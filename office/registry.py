"""
Varan registry — routes a filename to the correct editor based on extension.
"""
from __future__ import annotations

from pathlib import Path

from .word_editor import WordEditor
from .excel_editor import ExcelEditor
from .ppt_editor import PowerPointEditor
from .text_editor import TextEditor


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
            return editor.read_document(path)
        if isinstance(editor, ExcelEditor):
            return editor.read_workbook(path)
        if isinstance(editor, TextEditor):
            return editor.read(path)
        return editor.read_presentation(path)

    def summarize(self, path: str | Path, **kwargs):
        editor = self.editor_for(path)
        if isinstance(editor, WordEditor):
            return editor.summarize(path, **kwargs)
        if isinstance(editor, TextEditor):
            return editor.summarize(path, **kwargs)
        return editor.summarize(path)
