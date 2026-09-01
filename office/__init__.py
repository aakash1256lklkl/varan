"""Varan office editors: Word, Excel, PowerPoint."""
from .word_editor import WordEditor
from .excel_editor import ExcelEditor
from .ppt_editor import PowerPointEditor
from .registry import Registry

__all__ = ["WordEditor", "ExcelEditor", "PowerPointEditor", "Registry"]
