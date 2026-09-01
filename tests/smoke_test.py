"""
Varan smoke test — verifies all three Office editors work
WITHOUT needing a live AI provider key.
Run: python tests/smoke_test.py
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, str(ROOT))

from office.registry import Registry  # noqa: E402


def _assert_live_safety():
    """Live editors must attach to the ALREADY-RUNNING Office app and never
    spawn/terminate Office processes (which would kill the user's real session).
    Guards against regressions like DispatchEx or .Quit()/taskkill creeping in."""
    import office.word_live as wl
    import office.excel_live as xl
    import office.ppt_live as pl
    for mod in (wl, xl, pl):
        f = getattr(mod, "__file__", None)
        assert f, f"{mod.__name__} has no __file__"
        src = open(Path(f), encoding="utf-8").read()
        assert "DispatchEx" not in src, f"{mod.__name__} uses DispatchEx (spawns Office)"
        assert "Dispatch(" not in src, f"{mod.__name__} uses Dispatch (spawns Office)"
        assert ".Quit()" not in src, f"{mod.__name__} calls .Quit() (kills Office)"
        assert "GetActiveObject" in src, f"{mod.__name__} does not attach to running Office"


def _assert_strict_mode():
    """--strict mode must surface unexpected live-edit errors verbatim instead of
    silently falling through (which previously hid the cause behind a lock error)."""
    from agent.tools import ToolExecutor
    out = ROOT / "outputs"

    def _run(strict: bool):
        te = ToolExecutor(out, strict=strict)
        te.registry.word.is_open_in_word = lambda path: True
        real = te.registry.word.live_edit

        def boom(*a, **k):
            raise RuntimeError("simulated unexpected live failure")

        te.registry.word.live_edit = boom
        try:
            return te.execute("edit_document", {
                "path": "strict_test.docx", "inplace": True,
                "edits": [{"action": "replace", "match": "a", "text": "b"}]})
        finally:
            te.registry.word.live_edit = real

    normal = _run(strict=False)
    assert "error" in normal and "unexpected error" in normal["error"], \
        "non-strict mode should report the unexpected error clearly, not fall through"
    strict = _run(strict=True)
    assert strict == {"error": "RuntimeError: simulated unexpected live failure"}, \
        "strict mode should surface the underlying exception verbatim"


def _assert_noop_nudge():
    """Guards against the 'it says done but did nothing' regression: if the model
    replies with plain text and NO tool call on an actionable request (create/edit/
    delete/summarize...), the loop must nudge it to actually perform the action
    instead of declaring done with no tool ever running."""
    from agent.loop import Agent
    from agent.providers import ChatMessage, ToolCall, BaseProvider
    from rich.console import Console
    from io import StringIO

    calls = {"n": 0}

    class _Fake(BaseProvider):
        name = "fake"
        def __init__(self):
            super().__init__({"provider": "fake"})
        def chat(self, messages, tools=None):
            calls["n"] += 1
            if calls["n"] == 1:  # text-only no-op (the bug)
                return ChatMessage(role="assistant", text="Done!", tool_calls=None)
            if calls["n"] == 2:  # nudged -> calls the tool
                return ChatMessage(role="assistant", text="", tool_calls=[
                    ToolCall(id="1", name="create_document",
                             arguments={"path": "_nudge.docx", "title": "T"},
                             raw_arguments='{"path": "_nudge.docx", "title": "T"}')])
            return ChatMessage(role="assistant", text="OK.", tool_calls=None)  # settle

    agent = Agent(_Fake(), {"provider": "fake"},
                  Console(file=StringIO(), force_terminal=False))
    events = []
    agent.run("create a report document", sink=lambda k, p: events.append((k, p)))

    assert calls["n"] == 3, f"expected 3 calls (no-op, tool, settle), got {calls['n']}"
    assert "tool" in [k for k, _ in events], "no tool event fired after the nudge"
    # relational check: the nudge heuristic flags real actions but not questions
    assert Agent._looks_actionable("create a report doc") is True
    assert Agent._looks_actionable("add a new section") is True
    assert Agent._looks_actionable("delete page two") is True
    assert Agent._looks_actionable("what is a docx file?") is False
    assert Agent._looks_actionable("how do i make a chart?") is False



def main():
    reg = Registry()
    out = ROOT / "outputs"
    out.mkdir(parents=True, exist_ok=True)
    ok = True

    print("-> Verifying live-editor safety (no Office process spawn/kill)...")
    _assert_live_safety()
    print("   OK - live editors attach to the running app only")
    print("-> Verifying strict mode surfaces unexpected live-edit errors...")
    _assert_strict_mode()
    print("   OK - strict mode surfaces the real error; normal mode reports clearly")
    print("-> Verifying no-op text-only replies get nudged into real tool use...")
    _assert_noop_nudge()
    print("   OK - 'says done but does nothing' is caught; actionable requests re-run the tool")
    # 1. Word document
    print("-> Testing Word editor...")
    doc_path = out / "smoke_test.docx"
    created = reg.word.create_document(
        doc_path,
        title="Varan Smoke Test",
        body=[
            {"type": "heading", "level": 1, "text": "Introduction"},
            {"type": "paragraph", "text": "This document was created by Varan.", "italic": True},
            {"type": "bullet", "text": "Bullet one"},
            {"type": "bullet", "text": "Bullet two"},
            {"type": "table", "headers": ["Name", "Age"], "rows": [["Ada", 36], ["Grace", 45]]},
        ],
    )
    assert os.path.exists(created), "docx not created"
    doc_read = reg.word.read_document(created)
    assert doc_read["paragraphs"], "docx has no paragraphs"
    summ = reg.word.summarize(created)
    print(f"   OK - {summ['title']}, {summ['paragraph_count']} paragraphs, {summ['table_count']} tables")

    # 2. Excel workbook with data
    print("-> Testing Excel editor...")
    xl_path = out / "smoke_test.xlsx"
    xcreated = reg.excel.create_workbook(
        xl_path,
        sheet="Sales",
        data=[
            {"cells": ["Month", "Revenue"]},
            {"cells": ["Jan", 100]},
            {"cells": ["Feb", 150]},
            {"cells": ["Mar", 130]},
        ],
    )
    assert os.path.exists(xcreated), "xlsx not created"
    # add a chart
    xedited = reg.excel.edit_workbook(
        xcreated,
        add_chart={"type": "bar", "title": "Monthly Revenue",
                   "categories": "A2:A4", "data": "B2:B4"},
    )
    xread = reg.excel.read_workbook(xedited)
    assert "Sales" in xread["sheets"], "sheet missing"
    print(f"   OK - sheets={list(xread['sheets'].keys())}")

    # 3. PowerPoint with slides, table, chart
    print("-> Testing PowerPoint editor...")
    ppt_path = out / "smoke_test.pptx"
    pcreated = reg.ppt.create_presentation(ppt_path, slides=[
        {"layout": "title", "title": "Varan Deck", "subtitle": "By Varan"},
        {"layout": "bullets", "title": "Agenda",
         "bullets": ["Intro", "Results", "Next Steps"]},
        {"layout": "title", "title": "Revenue",
         "chart": {"type": "bar", "title": "Revenue", "categories": ["Q1", "Q2", "Q3"],
                   "data": [100, 150, 130], "series_name": "Revenue"}},
    ])
    assert os.path.exists(pcreated), "pptx not created"
    pr = reg.ppt.read_presentation(pcreated)
    assert pr["slide_count"] == 3, "expected 3 slides"
    print(f"   OK - {pr['slide_count']} slides")

    # 4. Plain-text file
    print("-> Testing text editor...")
    txt_path = out / "smoke_test.md"
    txt_path.write_text("# Title\n\nAlpha\nBeta line\n\nGamma\n", encoding="utf-8")
    tsum = reg.text.summarize(txt_path)
    assert tsum["kind"] == "text" and tsum["word_count"] > 0
    t_edited = reg.text.edit(
        txt_path, edits=[{"action": "replace", "match": "Beta line", "text": "Beta EDITED"}],
        inplace=True,
    )
    assert "Beta EDITED" in Path(t_edited).read_text(encoding="utf-8")
    # delete_range removes the block from 'Alpha' through 'Beta EDITED'
    reg.text.edit(txt_path, edits=[{"action": "delete_range", "match": "Alpha", "end_match": "Beta EDITED"}], inplace=True)
    remaining = Path(t_edited).read_text(encoding="utf-8")
    assert "Gamma" in remaining and "Alpha" not in remaining
    print("   OK - read/replace/delete_range on .md")

    # 5. PDF read/summarize/edit (only if pypdf + reportlab are available)
    print("-> Testing PDF editor...")
    try:
        from reportlab.pdfgen import canvas as _canvas
        pdf_path = out / "smoke_test.pdf"
        c = _canvas.Canvas(str(pdf_path))
        c.setFont("Helvetica", 12)
        c.drawString(72, 780, "SECTION ALPHA")
        c.drawString(72, 760, "First body line here.")
        c.save()
        psum = reg.text.summarize(pdf_path)
        assert psum["kind"] == "pdf" and psum["page_count"] >= 1
        pdf_edited = reg.text.edit_pdf(
            pdf_path,
            edits=[{"action": "replace", "match": "First body line here.", "text": "EDITED LINE"}],
            inplace=True,
        )
        from pypdf import PdfReader as _PdfReader
        read_back = _PdfReader(pdf_edited).pages[0].extract_text()
        assert "EDITED LINE" in read_back and "First body line here." not in read_back
        print(f"   OK - read/summarize/edit {psum['page_count']} page PDF")
    except ImportError:
        print("   SKIP - reportlab/pypdf not installed for PDF test")

    if ok:
        print("\n[PASS] All Varan Office editors work correctly.")
        return 0
    return 1


if __name__ == "__main__":
    sys.exit(main())
