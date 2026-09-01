"""
Varan live Excel editing via COM automation (pywin32).

When an .xlsx is open in Microsoft Excel, Excel holds an exclusive OS handle on
it, so any raw file write from Varan to that exact path is blocked. The only way
to edit an OPEN workbook live on screen is to drive Excel itself through its COM
API: connect to the running Excel.Application, locate the open Workbook, and run
Excel's own commands (Range writes, formulas, adding sheets) on it. Excel owns
its file state, so there is no lock conflict and the user sees changes appear in
the open window instantly.

The operations mirror ExcelEditor.edit_workbook's vocabulary:
  writes:    [{"cell": "A1", "value": ...}]
  formulas:  [{"cell": "C2", "formula": "=SUM(A1:A10)"}]
  new_sheet: "name"
  add_chart: {"type": "bar"|"line"|"pie", ...}   (best-effort via AddChart2)

Raised ExcelNotAvailable when Excel is not running or the workbook is not open,
so the caller can fall back to a file-based edit.
"""
from __future__ import annotations

import os
from pathlib import Path


class ExcelNotAvailable(Exception):
    """Raised when Microsoft Excel is not reachable via COM or doc is not open."""


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
    """Return the running Excel.Application via COM, or raise ExcelNotAvailable."""
    try:
        import win32com.client
    except Exception as exc:  # noqa: BLE001
        raise ExcelNotAvailable("pywin32 is not installed.") from exc
    try:
        app = win32com.client.GetActiveObject("Excel.Application")
    except Exception as exc:  # noqa: BLE001
        raise ExcelNotAvailable(
            "Microsoft Excel is not running (or COM is unavailable)."
        ) from exc
    try:
        app.Visible = True
    except Exception:  # noqa: BLE001
        pass
    return app


def _find_open_workbook(app, path: str | Path):
    want = os.path.normcase(os.path.realpath(str(path)))
    for i in range(app.Workbooks.Count):
        wb = app.Workbooks.Item(i + 1)
        try:
            full = wb.FullName
        except Exception:  # noqa: BLE001
            continue
        if os.path.normcase(os.path.realpath(full)) == want:
            return wb
    return None


def is_open_in_excel(path: str | Path) -> bool:
    """Return True if the given file is currently open in Microsoft Excel."""
    try:
        app = _connect()
    except ExcelNotAvailable:
        return False
    try:
        return _find_open_workbook(app, path) is not None
    finally:
        _cleanup(app)


def _to_cell_value(value):
    # Excel COM needs plain Python types; map common wrappers.
    if isinstance(value, bool):
        return bool(value)
    if value is None:
        return None
    return value


def _range_of(ws, cell: str):
    # Returns the Range for a cell address on the given worksheet.
    return ws.Range(str(cell))


def _do_write(ws, cell: str, value) -> int:
    _range_of(ws, cell).Value = _to_cell_value(value)
    return 1


def _do_formula(ws, cell: str, formula: str) -> int:
    rng = _range_of(ws, cell)
    rng.Formula = str(formula)
    return 1


def _do_new_sheet(wb, name: str) -> int:
    for i in range(wb.Worksheets.Count):
        if wb.Worksheets.Item(i + 1).Name == name:
            return 0  # already exists
    ws = wb.Worksheets.Add()
    try:
        ws.Name = str(name)
    except Exception:  # noqa: BLE001
        pass
    return 1


def _do_add_chart(ws, spec: dict) -> int:
    try:
        ctype = (spec.get("type") or "bar").lower()
        if ctype == "pie":
            xl = 5  # xlPie
        elif ctype == "line":
            xl = 4  # xlLine
        else:
            xl = 51  # xlColumnClustered
        chart_objects = getattr(ws, "ChartObjects", None)
        if chart_objects is None:
            return 0
        co = chart_objects.Add()
        chart = co.Chart
        chart.ChartType = xl
        title = spec.get("title")
        if title:
            try:
                chart.HasTitle = True
                chart.ChartTitle.Text = str(title)
            except Exception:  # noqa: BLE001
                pass
        # Add data from a range if provided.
        data = spec.get("data")
        if data:
            try:
                chart.SetSourceData(ws.Range(data))
            except Exception:  # noqa: BLE001
                pass
        return 1
    except Exception:  # noqa: BLE001
        return 0


def _do_delete_sheet(wb, name: str) -> int:
    """Delete a worksheet by name (best-effort; never kills the last sheet)."""
    try:
        if wb.Worksheets.Count < 2:
            return 0
        ws_del = wb.Worksheets(str(name))
        disp = getattr(wb, "Application", None)
        old_disp = None
        if disp is not None:
            try:
                old_disp = disp.DisplayAlerts
                disp.DisplayAlerts = False
            except Exception:  # noqa: BLE001
                old_disp = None
        try:
            ws_del.Delete()
        finally:
            if disp is not None and old_disp is not None:
                try:
                    disp.DisplayAlerts = old_disp
                except Exception:  # noqa: BLE001
                    pass
        return 1
    except Exception:  # noqa: BLE001
        return 0


def _do_row_op(ws, op: dict) -> int:
    try:
        action = (str(op.get("action") or "insert")).lower()
        at = int(op.get("at", 1) or 1)
        rng = ws.Rows(at)
        if action == "delete":
            rng.Delete()
        else:
            rng.Insert()
        return 1
    except Exception:  # noqa: BLE001
        return 0


def _do_col_op(ws, op: dict) -> int:
    try:
        action = (str(op.get("action") or "insert")).lower()
        letter = str(op.get("at") or "A")
        rng = ws.Columns(letter)
        if action == "delete":
            rng.Delete()
        else:
            rng.Insert()
        return 1
    except Exception:  # noqa: BLE001
        return 0


def _do_clear(ws, rng_str: str) -> int:
    try:
        ws.Range(str(rng_str)).Clear()
        return 1
    except Exception:  # noqa: BLE001
        return 0


def _do_style(ws, spec: dict) -> int:
    try:
        rng = ws.Range(str(spec.get("cell")))
        if spec.get("bold") is not None:
            rng.Font.Bold = bool(spec["bold"])
        if spec.get("italic") is not None:
            rng.Font.Italic = bool(spec["italic"])
        if spec.get("size"):
            rng.Font.Size = float(spec["size"])
        if spec.get("font"):
            rng.Font.Name = str(spec["font"])
        fill = spec.get("fill")
        if fill:
            hex_color = str(fill).lstrip("#")
            if len(hex_color) == 6:
                r, g, b = hex_color[0:2], hex_color[2:4], hex_color[4:6]
                backend = int(b, 16) << 16 | int(g, 16) << 8 | int(r, 16)
                rng.Interior.Color = backend
        return 1
    except Exception:  # noqa: BLE001
        return 0


def live_edit(path: str | Path, sheet: str | None = None,
              writes=None, formulas=None, new_sheet: str | None = None,
              delete_sheet: str | None = None, rows=None, columns=None,
              clear=None, styles=None, add_chart=None,
              save: bool = True) -> dict:
    """Apply a full edit job to a workbook open in Excel, live.

    Mirrors ExcelEditor.edit_workbook's vocabulary, including the complex
    operations: rows/columns (insert/delete), clear (range strings), styles
    (bold/italic/size/font/fill), delete_sheet, new_sheet, writes, formulas,
    add_chart.

    Returns {"ok": True, "path": ..., "mode": "live-excel"}.
    Raises ExcelNotAvailable if Excel is not running or the workbook isn't open.
    """
    _co_init()
    app = _connect()
    try:
        wb = _find_open_workbook(app, path)
        if wb is None:
            raise ExcelNotAvailable(
                "The workbook is not open in the running Excel session."
            )
        ws = None
        if sheet is not None:
            try:
                ws = wb.Worksheets(str(sheet))
            except Exception:  # noqa: BLE001
                ws = None
        if ws is None:
            ws = wb.ActiveSheet

        n = 0
        for w in (writes or []):
            cell = (w or {}).get("cell")
            if cell:
                n += _do_write(ws, cell, (w or {}).get("value"))
        for f in (formulas or []):
            cell = (f or {}).get("cell")
            if cell:
                n += _do_formula(ws, cell, str((f or {}).get("formula") or ""))
        if new_sheet:
            n += _do_new_sheet(wb, new_sheet)
        if delete_sheet:
            n += _do_delete_sheet(wb, delete_sheet)
        for op in (rows or []):
            n += _do_row_op(ws, op)
        for op in (columns or []):
            n += _do_col_op(ws, op)
        for rng in (clear or []):
            n += _do_clear(ws, rng)
        for spec in (styles or []):
            n += _do_style(ws, spec)
        if add_chart:
            n += _do_add_chart(ws, add_chart)
        if save:
            try:
                wb.Save()
            except Exception:  # noqa: BLE001
                pass
    finally:
        _cleanup(app)
    return {"ok": True, "path": str(path), "mode": "live-excel"}


def append_text(path: str | Path, text: str, sheet: str | None = None,
                save: bool = True) -> dict:
    """Write plain text into the next empty cell of column A (live, on screen).

    Excel has no paragraph "append"; this lands the text in the first empty cell
    of column A so content visibly appears in the open workbook.
    """
    _co_init()
    app = _connect()
    try:
        wb = _find_open_workbook(app, path)
        if wb is None:
            raise ExcelNotAvailable(
                "The workbook is not open in the running Excel session."
            )
        ws = None
        if sheet is not None:
            try:
                ws = wb.Worksheets(str(sheet))
            except Exception:  # noqa: BLE001
                ws = None
        if ws is None:
            ws = wb.ActiveSheet
        # Find first empty cell in column A.
        row = 1
        while True:
            try:
                cell = ws.Range(f"A{row}")
                val = cell.Value
            except Exception:  # noqa: BLE001
                break
            if val in (None, ""):
                break
            row += 1
        _do_write(ws, f"A{row}", str(text))
        if save:
            try:
                wb.Save()
            except Exception:  # noqa: BLE001
                pass
    finally:
        _cleanup(app)
    return {"ok": True, "path": str(path), "mode": "live-excel"}
