"""
Varan live Word editing via COM automation (pywin32).

When a .docx is open in Microsoft Word, Word holds an exclusive OS handle on it,
so any raw file write from Varan to that exact path is blocked. The only way to
edit an OPEN document live on screen is to drive Word itself through its COM
API: connect to the running Word.Application, locate the open Document, and run
Word's own commands (Find/Replace, Range insertion, append) on it. Word owns its
file state, so there is no lock conflict and the user sees changes appear in the
open window instantly — like typing in the document.

This makes Varan behave like a coding CLI (Cursor/VS Code) for Office files:
whatever document you have open, Varan edits THAT live document in place, never
leaving a duplicate.

The operations mirror OfficeWordEditor.edit_document's vocabulary so the agent's
edit_* calls just work, live:
  edits: [ {action: replace|insert_after|insert_before|delete, match: text,
            text: replacement/insertion} ... ]
  append: text to add at the end of the body

Raised WordNotAvailable when Word is not running or the doc is not open, so the
caller can fall back to a file-based edit.
"""
from __future__ import annotations

import os
from pathlib import Path


class WordNotAvailable(Exception):
    """Raised when Microsoft Word is not reachable via COM or doc is not open."""


class MatchNotFoundError(Exception):
    """Raised when a text match the user asked to edit/delete isn't in the doc.

    This is a USER-FACING problem (bad anchor), not an I/O or lock problem. It
    must never be confused with the file being locked, so the executor reports
    it clearly instead of pretending the file is locked.
    """


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
    """Return the running Word.Application via COM, or raise WordNotAvailable."""
    try:
        import win32com.client
    except Exception as exc:  # noqa: BLE001
        raise WordNotAvailable("pywin32 is not installed.") from exc
    try:
        app = win32com.client.GetActiveObject("Word.Application")
    except Exception as exc:  # noqa: BLE001
        raise WordNotAvailable(
            "Microsoft Word is not running (or COM is unavailable)."
        ) from exc
    try:
        app.Visible = True
    except Exception:  # noqa: BLE001
        pass
    return app


def _find_open_document(app, path: str | Path):
    want = os.path.normcase(os.path.realpath(str(path)))
    for i in range(app.Documents.Count):
        doc = app.Documents.Item(i + 1)
        try:
            full = doc.FullName
        except Exception:  # noqa: BLE001
            continue
        if os.path.normcase(os.path.realpath(full)) == want:
            return doc
    return None


def is_open_in_word(path: str | Path) -> bool:
    """Return True if the given file is currently open in Microsoft Word."""
    try:
        app = _connect()
    except WordNotAvailable:
        return False
    try:
        return _find_open_document(app, path) is not None
    finally:
        _cleanup(app)


def _do_append(doc, text: str) -> int:
    rng = doc.Content
    rng.Collapse(0)  # wdCollapseEnd
    if text:
        rng.InsertAfter("\r" + text)
    return 1


def _do_replace(doc, match: str, text: str, count: int = 1) -> int:
    """Replace `match` with `text`. count == -1 replaces EVERY occurrence in the
    whole document (replace-all, parity with the file editor's count=-1);
    otherwise only the FIRST occurrence is replaced and a MatchNotFoundError is
    raised if the text isn't present — never a silent no-op."""
    if count < 0:
        rng = doc.Range(0, 0)
        rng.End = doc.Content.End
        f = rng.Find
        # The fully-explicit positional form is required — relying on optional
        # arguments makes Word ignore the replacement on some builds.
        f.Execute(
            FindText=str(match),
            MatchCase=False,
            MatchWholeWord=False,
            MatchWildcards=False,
            MatchSoundsLike=False,
            MatchAllWordForms=False,
            Forward=True,
            Wrap=1,  # wdFindContinue
            Format=False,
            ReplaceWith=str(text),
            Replace=2,  # wdReplaceAll
        )
        return 1
    rng = _range_of_match(doc, match)
    if rng is None:
        raise MatchNotFoundError(f"match not found in open Word document: {match!r}")
    rng.Text = str(text)
    return 1


def _range_of_match(doc, match: str):
    rng = doc.Content
    find = rng.Find
    find.ClearFormatting()
    find.Text = match
    find.Forward = True
    find.Wrap = 0  # wdFindStop
    if find.Execute():
        return rng
    return None


def _range_of_match_last(doc, match: str):
    """Find the LAST occurrence of match in the document (returns a Range at it)
    or None. Used for the end-anchor of delete_range so a repeated heading/title
    anchors the section boundary to its last occurrence, which is what the user
    means by 'through the end of that section'."""
    rng = doc.Content
    find = rng.Find
    find.ClearFormatting()
    find.Text = match
    find.Forward = True
    find.Wrap = 0  # wdFindStop
    find.MatchCase = False
    last = None
    while find.Execute():
        found = rng.Duplicate
        last = (found.Start, found.End)
        rng.Collapse(0)  # wdCollapseEnd — search onward from after this match
        try:
            rng.End = doc.Content.End
        except Exception:  # noqa: BLE001
            pass
        find = rng.Find
        find.ClearFormatting()
        find.Text = match
        find.Forward = True
        find.Wrap = 0
        find.MatchCase = False
    if last is None:
        return None
    return doc.Range(last[0], last[1])


def _do_insert(doc, match: str, text: str, after: bool) -> int:
    rng = _range_of_match(doc, match)
    if rng is None:
        raise MatchNotFoundError(f"match not found in open Word document: {match!r}")
    insert = rng.Duplicate
    if after:
        insert.Collapse(0)  # wdCollapseEnd
    else:
        insert.Collapse(1)  # wdCollapseStart
    insert.InsertAfter("\r" + str(text))
    return 1


def _do_delete(doc, match: str) -> int:
    rng = _range_of_match(doc, match)
    if rng is None:
        raise MatchNotFoundError(f"match not found in open Word document: {match!r}")
    rng.Delete()
    return 1


def _do_delete_range(doc, start_match: str, end_match: str | None, end_level: int | None = None) -> int:
    """Delete reliably from the start anchor to the end anchor (inclusive of
    everything between them). This is the dependable way to remove a whole page
    / section without relying on Word's fragile page-pagination APIs.

    - start_match: first text of the block to remove.
    - end_match: last text of the block to remove. If None, the block runs to
      the very end of the document.
    - end_level: optional heading level (1-6). If given, the section ends just
      BEFORE the first later paragraph whose heading style is at this level —
      i.e. "delete this section up to the next Heading N". Takes precedence over
      end_match when both are supplied.
    """
    start_range = _range_of_match(doc, start_match)
    if start_range is None:
        raise MatchNotFoundError(
            f"to delete, I need the start of the section, but I couldn't find "
            f"{start_match!r} in the document."
        )
    start_pos = start_range.Start
    if end_level is not None:
        end_pos = _heading_level_end(doc, end_level, start_pos)
        if end_pos is None:
            raise MatchNotFoundError(
                f"I found the start {start_match!r} but no later Heading "
                f"{end_level} exists to mark the section end."
            )
    elif end_match and end_match.strip():
        # The end-anchor should target its LAST occurrence (a title/heading often
        # appears more than once; "through the end of that section" means the
        # last one), falling back to the first occurrence if there's only one.
        end_range = _range_of_match_last(doc, end_match)
        if end_range is None:
            end_range = _range_of_match(doc, end_match)
        if end_range is None:
            raise MatchNotFoundError(
                f"I found the start {start_match!r} but not the end {end_match!r} "
                f"of the section to delete."
            )
        end_pos = end_range.End
    else:
        end_pos = doc.Content.End
    if end_pos <= start_pos:
        raise MatchNotFoundError("the delete range is empty or inverted.")
    rng = doc.Range(start_pos, end_pos)
    rng.Delete()
    return 1


def _heading_level_end(doc, level: int, after_start: int) -> int | None:
    """Return the character position of the start of the first paragraph after
    `after_start` whose heading style is at the given level (1-6), or None if
    no such heading exists. This lets delete_range end "just before the next
    Heading N"."""
    level = int(level)
    if level < 1 or level > 6:
        raise MatchNotFoundError("heading level must be between 1 and 6.")
    paras = doc.Paragraphs
    count = paras.Count
    for i in range(1, count + 1):
        try:
            p = paras.Item(i)
            if p.Range.Start <= after_start:
                continue
        except Exception:  # noqa: BLE001
            continue
        try:
            st = p.Style.NameLocal
        except Exception:  # noqa: BLE001
            st = ""
        if _is_heading_level(st, level):
            return p.Range.Start
    return None


def _is_heading_level(style_name: str, level: int) -> bool:
    """Match a Word paragraph style name against a heading level. Style names are
    localized (e.g. 'Heading 1', 'Titre 1', 'Überschrift 1'), so we match the
    required trailing digit as the last token and accept any leading label that
    smells like a heading ('heading'/'title'/'titre'/'uberschrift'/etc.)."""
    parts = str(style_name).strip().split()
    if not parts:
        return False
    last = parts[-1]
    if not last.isdigit() or int(last) != level:
        return False
    first = " ".join(parts[:-1]).lower()
    heading_words = (
        "heading", "title", "titre", "titulo", "uberschrift", "überschrift",
        "headline", "rubrik", "enzcefal", "naslov", "titolo", "rubric",
    )
    return any(w in first for w in heading_words)


def _do_delete_page(doc, page: int) -> int:
    """Delete a whole numeric page via Word's pagination (GoTo).

    Works when the document is open in the user's INTERACTIVE Word session
    (where pagination is active) — which is exactly what the live editor
    attaches to. A hidden background session has no pagination and this falls
    back to deleting nothing (raising a clear MatchNotFoundError).
    """
    page = int(page)
    if page < 1:
        raise MatchNotFoundError("page number must be >= 1")
    try:
        doc.Repaginate()
        total_pages = int(doc.ComputeStatistics(2))  # wdStatisticPages
    except Exception:  # noqa: BLE001
        total_pages = 0
    if total_pages and page > total_pages:
        raise MatchNotFoundError(
            f"page {page} does not exist (the document has {total_pages} pages)."
        )

    start_of_page = doc.GoTo(1, 1, page).Start  # wdGoToPage / wdGoToAbsolute
    # Page N+1's start is the end boundary of page N (or document end for the
    # last page). Deleting [start_of_page .. start_of_next_page) removes page N.
    try:
        if page < (total_pages or page):
            start_of_next = doc.GoTo(1, 1, page + 1).Start
        else:
            start_of_next = doc.Content.End
    except Exception:  # noqa: BLE001
        start_of_next = doc.Content.End

    if start_of_page == 0:
        # Pagination unavailable (e.g. background session): be honest.
        raise MatchNotFoundError(
            "I couldn't locate page boundaries in this session. The document is "
            "likely being viewed in a way where Word hasn't laid out pages. "
            "Please keep it open in the main Word window and try again, or tell "
            "me the first and last line of the section to delete and I'll "
            "remove that block instead (action 'delete_range')."
        )
    if start_of_next <= start_of_page:
        raise MatchNotFoundError("page boundaries are empty or inverted.")
    rng = doc.Range(start_of_page, start_of_next)
    rng.Delete()
    return 1


def get_headings(path: str | Path, limit: int = 100) -> list[str]:
    """Return the non-empty paragraph texts of an open document (headings +
    body), useful for the model to pick a real anchor to delete/edit."""
    _co_init()
    app = _connect()
    try:
        doc = _find_open_document(app, path)
        if doc is None:
            raise WordNotAvailable("The document is not open in the running Word session.")
        paras = doc.Paragraphs
        out = []
        count = paras.Count
        for i in range(min(count, limit)):
            try:
                t = paras.Item(i + 1).Range.Text.strip()
            except Exception:  # noqa: BLE001
                t = ""
            if t:
                out.append(t)
        return out
    finally:
        _cleanup(app)


def _do_remove_empty_paragraphs(doc) -> int:
    """Remove every empty body paragraph from a live Word doc (the typical cause
    of 'blank pages'). Iterates backwards so indices stay valid. Returns how many
    paragraphs were removed."""
    removed = 0
    try:
        count = doc.Paragraphs.Count
    except Exception:  # noqa: BLE001
        return 0
    for i in range(count, 0, -1):
        p = doc.Paragraphs.Item(i)
        try:
            txt = p.Range.Text or ""
        except Exception:  # noqa: BLE001
            continue
        if not txt.strip():
            try:
                p.Range.Delete()
                removed += 1
            except Exception:  # noqa: BLE001
                pass
    return removed


def _do_insert_content(doc, blocks) -> int:
    """Append structured content BLOCKS to the open document as real styled
    Word paragraphs: true Heading styles, List Bullet / List Number, bold /
    italic runs, a Courier font, and centered scene-break dividers.

    Returns the number of blocks inserted.
    """
    n = 0
    for block in blocks or []:
        if not isinstance(block, dict):
            continue
        btype = block.get("type", "paragraph")
        text = str(block.get("text", "") or "")

        rng = doc.Content
        rng.Collapse(0)  # wdCollapseEnd
        rng.InsertAfter("\r" + text)  # expands rng to cover the new paragraph
        n += 1

        if btype in ("divider", "scene_break"):
            try:
                rng.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
            except Exception:  # noqa: BLE001
                pass
        else:
            style_name = "Normal"
            if btype == "title":
                style_name = "Title"
            elif btype == "heading":
                level = int(block.get("level", 1) or 1)
                level = max(1, min(level, 6))
                style_name = f"Heading {level}"
            elif btype == "bullet":
                style_name = "List Bullet"
            elif btype == "numbered":
                style_name = "List Number"
            try:
                rng.Style = doc.Styles(style_name)
            except Exception:  # noqa: BLE001
                pass  # template has no such style; keep Normal

        try:
            fname = block.get("font") or (block.get("courier") and "Courier New")
            if fname:
                rng.Font.Name = str(fname)
            if block.get("bold"):
                rng.Font.Bold = True
            if block.get("italic"):
                rng.Font.Italic = True
        except Exception:  # noqa: BLE001
            pass

        rng.Collapse(0)  # end of the paragraph just inserted, ready for next
    return n


def live_edit(path: str | Path, edits=None, append: str | None = None,
              page: int | None = None, save: bool = True,
              remove_empty_paragraphs: bool = False, content=None,
              table_edits=None) -> dict:
    """Apply a full edit job to a document open in Word, live.

    table_edits: [{"table": <1-based index>, "cell": {"r": row, "c": col} or
                   {"ref": "A1"}, "text": ...}] — set a cell's text via COM,
                   preserving the cell's formatting.

    Returns {"ok": True, "path": ..., "mode": "live-word"}.
    Raises WordNotAvailable if Word is not running or the doc isn't open.
    """
    _co_init()
    app = _connect()
    try:
        doc = _find_open_document(app, path)
        if doc is None:
            raise WordNotAvailable(
                "The document is not open in the running Word session."
            )
        n = 0
        if remove_empty_paragraphs:
            n += _do_remove_empty_paragraphs(doc)
        if page is not None:
            n += _do_delete_page(doc, page)
        for edit in (edits or []):
            action = (edit or {}).get("action", "replace")
            match = str((edit or {}).get("match", "") or "")
            text = str((edit or {}).get("text", "") or "")
            if action == "delete_page":
                n += _do_delete_page(doc, int(edit.get("page", 1) or 1))
            elif action == "delete_range":
                n += _do_delete_range(
                    doc, match,
                    str((edit or {}).get("end_match", "") or ""),
                    end_level=(edit or {}).get("end_level"),
                )
            elif action == "delete":
                n += _do_delete(doc, match)
            elif action in ("insert_after", "insert_before"):
                n += _do_insert(doc, match, text, after=(action == "insert_after"))
            else:  # replace
                count = int((edit or {}).get("count", 1) or 1)
                n += _do_replace(doc, match, text, count=count)
        for te in (table_edits or []):
            n += _do_table_edit(doc, te)
        if append is not None:
            n += _do_append(doc, str(append))
        if content:
            n += _do_insert_content(doc, content)
        if save:
            try:
                doc.Save()
            except Exception:  # noqa: BLE001
                pass
    finally:
        _cleanup(app)
    return {"ok": True, "path": str(path), "mode": "live-word", "removed": n}


def _do_table_edit(doc, te: dict) -> int:
    """Set the text of one cell in an existing table (best-effort COM)."""
    idx = int(te.get("table", 1) or 1)
    if idx < 1 or idx > doc.Tables.Count:
        raise MatchNotFoundError(
            f"table {idx} was not found (the open document has "
            f"{doc.Tables.Count} table(s))."
        )
    try:
        tbl = doc.Tables.Item(idx)
        cell = te.get("cell") or {}
        if isinstance(cell, dict):
            if "r" in cell and "c" in cell:
                r, c = int(cell["r"]) + 1, int(cell["c"]) + 1
            else:
                ref = str(cell.get("ref") or "A1")
                c = 0
                for ch in ref:
                    if ch.isalpha():
                        c = c * 26 + (ord(ch.upper()) - 64)
                    else:
                        break
                r = int("".join(ch for ch in ref if ch.isdigit()) or 1)
        else:
            ref = str(cell or "A1")
            c = 0
            for ch in ref:
                if ch.isalpha():
                    c = c * 26 + (ord(ch.upper()) - 64)
                else:
                    break
            r = int("".join(ch for ch in ref if ch.isdigit()) or 1)
        tbl.Cell(r, c).Range.Text = str(te.get("text", ""))
        return 1
    except Exception:  # noqa: BLE001
        return 0


def append_text(path: str | Path, text: str, save: bool = True) -> dict:
    """Append plain text to a document currently open in Word, live."""
    return live_edit(path, append=str(text), save=save)
