"""
Varan live PowerPoint editing via COM automation (pywin32).

When a .pptx is open in Microsoft PowerPoint, PowerPoint holds an exclusive OS
handle on it, so any raw file write from Varan to that exact path is blocked.
The only way to edit an OPEN presentation live on screen is to drive PowerPoint
itself through its COM API: connect to the running PowerPoint.Application, locate
the open Presentation, and add slides / textboxes via PowerPoint's own API.
PowerPoint owns its file state, so there is no lock conflict and the user sees
new slides appear in the running show instantly.

The operations mirror PowerPointEditor.edit_presentation's vocabulary:
  add_slides: [ {layout/title/subtitle/bullets/...} ... ]  (best-effort via COM)

Raised PptNotAvailable when PowerPoint is not running or the presentation is not
open, so the caller can fall back to a file-based edit.
"""
from __future__ import annotations

import os
from pathlib import Path

# PPT layout enums: ppLayoutBlank=12, ppLayoutText=2, ppLayoutTitle=1
_PP_LAYOUT_BLANK = 12
_PP_LAYOUT_TEXT = 2
_PP_LAYOUT_TITLE = 1
# ppPlaceholderBody = 2, ppPlaceholderSubtitle = 11
_PP_PH_BODY = 2
_PP_PH_SUBTITLE = 11

# Same degenerate-filler guard as ppt_editor._assert_real_slide_content so the
# live COM path also refuses to write placeholder garbage ("- item", "item",
# "i | t | e | m") instead of silently "succeeding".
_FILLER_TOKENS = {
    "item", "items", "item1", "item 1", "insert", "placeholder", "placeholders",
    "your text here", "text here", "text", "content", "contents", "sample",
    "sample text", "lorem", "ipsum", "lorem ipsum", "todo", "tbd", "t.b.d.",
    "n/a", "na", "- item", "-item", "• item", "bullet", "bullets",
}


def _is_filler(value) -> bool:
    t = str(value or "").strip().lower()
    return t in _FILLER_TOKENS


def _assert_real_slide_content(spec: dict):
    bullets = spec.get("bullets") if spec.get("bullets") is not None else spec.get("set_bullets")
    if bullets is not None:
        real = [b for b in bullets if (b or "").strip()]
        if real and all(_is_filler(b) for b in real):
            raise FillerContentError(
                "Refusing placeholder-only content: every bullet is filler "
                f"({real!r}). Provide REAL slide content, or use edit_slides "
                "with 'set_bullets' / 'add_table' to fill an existing text box."
            )
    table = spec.get("table") if spec.get("table") is not None else spec.get("add_table")
    if table:
        headers = [h for h in (table.get("headers") or []) if (h or "").strip()]
        rows = [[v for v in (r or []) if (v or "").strip()] for r in (table.get("rows") or [])]
        cells = [v for row in rows for v in row]
        if headers and all(_is_filler(h) for h in headers) and (
            not cells or all(len(str(v).strip()) <= 1 for v in cells)
        ):
            raise FillerContentError(
                "Refusing placeholder-only table: "
                f"headers={headers!r} rows={[[str(v) for v in r] for r in rows]!r}. "
                "Provide REAL table data (headers + cell values), or use "
                "edit_slides with 'add_table' to fill a table on an existing slide."
            )


class PptNotAvailable(Exception):
    """Raised when Microsoft PowerPoint is not reachable via COM or doc not open."""


class FillerContentError(ValueError):
    """Raised when requested slide content is degenerate placeholder filler
    ('- item', 'item', 'i | t | e | m'). Unlike PptNotAvailable, this is a
    REFUSAL of bad content — it must propagate (fail loud), never be swallowed
    by the generic COM best-effort handlers, so a junk slide can never be
    reported as a successful edit."""


def _co_init():
    import pythoncom
    try:
        pythoncom.CoInitialize()
    except Exception:  # noqa: BLE001
        pass


def _cleanup(app):
    try:
        if app is not None:
            app = None
    except Exception:  # noqa: BLE001
        pass


def _connect():
    """Return the running PowerPoint.Application via COM, or raise PptNotAvailable."""
    try:
        import win32com.client
    except Exception as exc:  # noqa: BLE001
        raise PptNotAvailable("pywin32 is not installed.") from exc
    try:
        app = win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception as exc:  # noqa: BLE001
        raise PptNotAvailable(
            "Microsoft PowerPoint is not running (or COM is unavailable)."
        ) from exc
    try:
        app.Visible = True
    except Exception:  # noqa: BLE001
        pass
    return app


def _find_open_presentation(app, path: str | Path):
    want = os.path.normcase(os.path.realpath(str(path)))
    for i in range(app.Presentations.Count):
        prs = app.Presentations.Item(i + 1)
        try:
            full = prs.FullName
        except Exception:  # noqa: BLE001
            continue
        if os.path.normcase(os.path.realpath(full)) == want:
            return prs
    return None


def is_open_in_ppt(path: str | Path) -> bool:
    """Return True if the given file is currently open in Microsoft PowerPoint."""
    try:
        app = _connect()
    except PptNotAvailable:
        return False
    try:
        return _find_open_presentation(app, path) is not None
    finally:
        _cleanup(app)


def _add_textbox(slide, text: str, left=0.5, top=1.4, width=9.0, height=1.0):
    try:
        box = slide.Shapes.AddTextbox(1, left * 72, top * 72, width * 72, height * 72)
        tr = box.TextFrame.TextRange
        tr.Text = str(text)
        tr.Font.Size = 20
        return box
    except Exception:  # noqa: BLE001
        return None


def _find_body_placeholder(slide):
    """Return the slide's BODY placeholder (ppPlaceholderBody) via COM, or None."""
    try:
        placeholders = slide.Shapes.Placeholders
        for i in range(1, placeholders.Count + 1):
            try:
                ph = placeholders.Item(i)
                if int(ph.PlaceholderFormat.Type) == _PP_PH_BODY:
                    return ph
            except Exception:  # noqa: BLE001
                continue
    except Exception:  # noqa: BLE001
        pass
    return None


def _set_body_text(slide, lines: list[str]) -> bool:
    """Fill the body placeholder (or first non-title text box) with bullet lines."""
    try:
        body = _find_body_placeholder(slide)
        tr = None
        if body is not None:
            tr = body.TextFrame.TextRange
        if tr is None:
            title = None
            try:
                title = slide.Shapes.Title
            except Exception:  # noqa: BLE001
                pass
            for shape in slide.Shapes:
                if title is not None and shape == title:
                    continue
                try:
                    tr = shape.TextFrame.TextRange
                    break
                except Exception:  # noqa: BLE001
                    continue
        if tr is None:
            return False
        tr.Text = "\r".join(str(b) for b in lines)
        return True
    except Exception:  # noqa: BLE001
        return False


def _add_slide(prs, spec: dict) -> int:
    """Add one slide to the open presentation (best-effort COM)."""
    _assert_real_slide_content(spec)
    layout_name = (spec.get("layout") or "bullets").lower()
    if layout_name == "title":
        layout = _PP_LAYOUT_TITLE
    elif layout_name == "blank":
        layout = _PP_LAYOUT_BLANK
    else:
        layout = _PP_LAYOUT_TEXT

    idx = prs.Slides.Count + 1
    try:
        slide = prs.Slides.Add(idx, layout)
    except Exception:  # noqa: BLE001
        return 0

    title = spec.get("title")
    if title is not None:
        try:
            slide.Shapes.Title.TextFrame.TextRange.Text = str(title)
        except Exception:  # noqa: BLE001
            pass

    subtitle = spec.get("subtitle")
    if subtitle:
        placed = False
        try:
            placeholders = slide.Shapes.Placeholders
            for i in range(1, placeholders.Count + 1):
                ph = placeholders.Item(i)
                if int(ph.PlaceholderFormat.Type) == _PP_PH_SUBTITLE:
                    ph.TextFrame.TextRange.Text = str(subtitle)
                    placed = True
                    break
        except Exception:  # noqa: BLE001
            pass
        if not placed:
            _add_textbox(slide, str(subtitle), top=1.4, height=1.0)

    bullets = spec.get("bullets")
    if bullets:
        lines = [str(b) for b in bullets if (b or "").strip()]
        if lines and not _set_body_text(slide, lines):
            top = 1.6
            for b in lines:
                _add_textbox(slide, "- " + b, top=top, height=0.6)
                top += 0.6

    table = spec.get("table")
    if table:
        _add_live_table(slide, table)

    chart = spec.get("chart")
    if chart:
        _add_live_chart(slide, chart)

    return 1


def live_edit(path: str | Path, add_slides=None, rebuild_slides=None,
              edit_slides=None, remove_slides=None, save: bool = True) -> dict:
    """Append, rebuild, remove, or surgically edit slides in a presentation
    open in PowerPoint, live.

    append: add add_slides to the existing deck.
    rebuild: delete every existing slide and rebuild from rebuild_slides.
    edit_slides: list of {"slide": 1-based, "title": ..., "replace": {...},
                          "set_bullets": [...], "add_textbox": {...},
                          "add_table": {...}}.
    remove_slides: list of 1-based slide indices to DELETE from the deck.

    Returns {"ok": True, "path": ..., "mode": "live-ppt", "summary": {...}}.
    Raises PptNotAvailable if PowerPoint is not running or the presentation isn't
    open.
    """
    _co_init()
    app = _connect()
    try:
        prs = _find_open_presentation(app, path)
        if prs is None:
            raise PptNotAvailable(
                "The presentation is not open in the running PowerPoint session."
            )
        before = prs.Slides.Count
        n = 0
        if rebuild_slides is not None:
            # remove every existing slide (iterate backwards so indexes stay valid)
            for i in range(before, 0, -1):
                try:
                    prs.Slides(i).Delete()
                except Exception:  # noqa: BLE001
                    pass
            for slide in rebuild_slides:
                n += _add_slide(prs, slide)
            summary = {"mode": "rebuild", "removed": before, "added": n, "total": n}
        else:
            edited_idx = []
            for spec in (edit_slides or []):
                _apply_live_edit_slides(prs, spec)
                edited_idx.append(int(spec.get("slide", 1)))
            removed_idx = []
            for idx in sorted({int(i) for i in (remove_slides or [])}, reverse=True):
                try:
                    prs.Slides.Item(idx).Delete()
                    removed_idx.append(idx)
                except Exception:  # noqa: BLE001
                    pass
            for slide in (add_slides or []):
                n += _add_slide(prs, slide)
            mode = "append" if not edited_idx and not removed_idx else "edit_slides"
            summary = {"mode": mode, "removed": len(removed_idx),
                       "added": n, "edited": sorted(set(edited_idx)),
                       "removed_idx": sorted(removed_idx),
                       "total": before - len(removed_idx) + n}
        if save:
            try:
                prs.Save()
            except Exception:  # noqa: BLE001
                pass
    finally:
        _cleanup(app)
    return {"ok": True, "path": str(path), "mode": "live-ppt", "summary": summary}


def _apply_live_edit_slides(prs, spec: dict):
    """Best-effort surgical text/shape edit on ONE slide (1-based index)."""
    idx = int(spec.get("slide", 1))
    try:
        slide = prs.Slides.Item(idx)
    except Exception:  # noqa: BLE001
        return
    _assert_real_slide_content({
        "bullets": spec.get("set_bullets"),
        "table": spec.get("add_table"),
    })
    title = spec.get("title")
    if title is not None:
        try:
            slide.Shapes.Title.TextFrame.TextRange.Text = str(title)
        except Exception:  # noqa: BLE001
            pass
    replace_map = spec.get("replace") or {}
    if replace_map:
        try:
            for shape in slide.Shapes:
                if not hasattr(shape, "TextFrame"):
                    continue
                tr = shape.TextFrame.TextRange
                txt = tr.Text
                for old_text, new_sub in replace_map.items():
                    if old_text and old_text in txt:
                        txt = txt.replace(old_text, new_sub)
                tr.Text = txt
        except Exception:  # noqa: BLE001
            pass
    set_bullets = spec.get("set_bullets")
    if set_bullets is not None:
        lines = [str(b) for b in set_bullets if (b or "").strip()]
        if lines:
            _set_body_text(slide, lines)
    add_tb = spec.get("add_textbox")
    if add_tb:
        d = add_tb if isinstance(add_tb, dict) else {}
        _add_textbox(slide, str(d.get("text", "") if isinstance(add_tb, dict) else (add_tb or "")),
                     top=float(d.get("top", 1.5)), height=float(d.get("height", 1.0)))
    add_table = spec.get("add_table")
    if add_table is not None:
        _add_live_table(slide, add_table)
    add_chart = spec.get("add_chart")
    if add_chart is not None:
        _add_live_chart(slide, add_chart)


def _add_live_table(slide, spec: dict):
    """Add a table to an existing slide (best-effort COM)."""
    try:
        headers = spec.get("headers") or []
        rows = spec.get("rows") or []
        nrows = len(rows) + (1 if headers else 0)
        ncols = max([len(headers)] + [len(r) for r in rows]) if nrows else 0
        if nrows == 0 or ncols == 0:
            return
        top = float(spec.get("top", 2.0))
        left = float(spec.get("left", 0.6))
        shape = slide.Shapes.AddTable(
            nrows, ncols, left * 72, top * 72, 8.8 * 72, 3.0 * 72
        )
        tbl = shape.Table
        ri = 0
        if headers:
            for ci, h in enumerate(headers):
                tbl.Cell(1, ci + 1).Shape.TextFrame.TextRange.Text = str(h)
            ri = 1
        for r in rows:
            vals = list(r)
            for ci in range(ncols):
                v = vals[ci] if ci < len(vals) else ""
                tbl.Cell(ri + 1, ci + 1).Shape.TextFrame.TextRange.Text = str(v)
            ri += 1
    except Exception:  # noqa: BLE001
        pass


def _add_live_chart(slide, spec: dict):
    """Add a chart to an existing slide (best-effort COM). PowerPoint's chart
    object model needs a linked workbook for real data; this at least places a
    chart frame and its title so the slide is not left chart-free."""
    try:
        # xlColumnClustered=51, xlLineMarkers=65, xlPie=5
        ctype = (spec.get("type") or "bar").lower()
        xl = 51 if ctype == "line" else (5 if ctype == "pie" else 51)
        left = float(spec.get("left", 1.0))
        top = float(spec.get("top", 2.5))
        shape = slide.Shapes.AddChart(xl, left * 72, top * 72, 6 * 72, 4 * 72)
        if spec.get("title"):
            try:
                shape.Chart.ChartTitle.Text = str(spec["title"])
            except Exception:  # noqa: BLE001
                pass
    except Exception:  # noqa: BLE001
        pass


def append_text(path: str | Path, text: str, save: bool = True) -> dict:
    """Add a textbox with the given text to the last slide, live."""
    _co_init()
    app = _connect()
    try:
        prs = _find_open_presentation(app, path)
        if prs is None:
            raise PptNotAvailable(
                "The presentation is not open in the running PowerPoint session."
            )
        count = prs.Slides.Count
        if count == 0:
            prs.Slides.Add(1, _PP_LAYOUT_BLANK)
            count = 1
        slide = prs.Slides.Item(count)
        _add_textbox(slide, str(text), top=1.6, height=1.0)
        if save:
            try:
                prs.Save()
            except Exception:  # noqa: BLE001
                pass
    finally:
        _cleanup(app)
    return {"ok": True, "path": str(path), "mode": "live-ppt"}
