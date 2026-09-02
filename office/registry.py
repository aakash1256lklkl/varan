"""
Varan registry — routes a filename to the correct editor based on extension.
"""
from __future__ import annotations

from pathlib import Path

from .word_editor import WordEditor
from .excel_editor import ExcelEditor
from .ppt_editor import PowerPointEditor
from .text_editor import TextEditor
from . import anydoc_reader


class Registry:
    def __init__(self):
        self._word = WordEditor()
        self._excel = ExcelEditor()
        self._ppt = PowerPointEditor()
        self._text = TextEditor()

    @staticmethod
    def classify(path: str | Path) -> str:
        """Return 'word' | 'excel' | 'ppt' | 'text' | 'pdf' based on extension."""
        p = Path(path)
        suffix = p.suffix.lower()
        if suffix in (".docx", ".doc"):
            return "word"
        if suffix in (".xlsx", ".xls", ".csv"):
            return "excel"
        if suffix in (".pptx", ".ppt"):
            return "ppt"
        if suffix in (".pdf",):
            return "pdf"
        if suffix in (".txt", ".md", ".markdown", ".rst", ".log", ".tex"):
            return "text"
        return "unknown"

    def editor_for(self, path: str | Path):
        kind = self.classify(path)
        if kind == "word":
            return self._word
        if kind == "excel":
            return self._excel
        if kind == "ppt":
            return self._ppt
        if kind in ("text", "pdf"):
            return self._text
        raise ValueError(f"Unsupported file type: {path}")

    # Convenience delegates -------------------------------------------------
    @property
    def word(self) -> WordEditor:
        return self._word

    @property
    def excel(self) -> ExcelEditor:
        return self._excel

    @property
    def ppt(self) -> PowerPointEditor:
        return self._ppt

    @property
    def text(self) -> TextEditor:
        return self._text

    def read(self, path: str | Path):
        editor = self.editor_for(path)
        if isinstance(editor, WordEditor):
            result = editor.read_document(path)
        elif isinstance(editor, ExcelEditor):
            result = editor.read_workbook(path)
        elif isinstance(editor, TextEditor):
            result = editor.read(path)
        else:
            result = editor.read_presentation(path)
        # Optional anydoc enhancement: attach clean GitHub-Flavored Markdown
        # (headings/tables/lists preserved) when the package is installed and
        # the conversion succeeds. The structured result is unchanged, so
        # callers/tests keep working; the model gets LLM-ready context.
        md = anydoc_reader.to_markdown(path)
        if md is not None:
            result["markdown"] = md
        return result

    def summarize(self, path: str | Path, **kwargs):
        editor = self.editor_for(path)
        if isinstance(editor, WordEditor):
            result = editor.summarize(path, **kwargs)
        elif isinstance(editor, TextEditor):
            result = editor.summarize(path, **kwargs)
        else:
            result = editor.summarize(path)
        # Optional anydoc enhancement: attach a short Markdown preview so the
        # model can see the document's real structure (headings, tables, lists)
        # even for formats Varan's own readers only skim.
        md = anydoc_reader.to_markdown(path)
        if md is not None:
            result["markdown_preview"] = md[:1500]
        return result
