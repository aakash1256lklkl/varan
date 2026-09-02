"""
Optional anydoc integration — convert Office/PDF files to clean Markdown.

Uses the `firecrawl-anydoc` package (Rust core, MIT license) when it is
installed. Every function here degrades gracefully to None when the package
is unavailable or a conversion fails, so Varan keeps working with its
built-in extractors (python-docx / openpyxl / python-pptx / pypdf).

anydoc turns Word, PowerPoint, Excel, OpenDocument, RTF, EPUB, CSV and PDF
into one consistent GitHub-Flavored Markdown representation (headings,
tables, lists, footnotes preserved) — ideal LLM-ready context for
read_file / summarize_file.
"""
from __future__ import annotations

from pathlib import Path

_ANYDOC = None
try:  # pragma: no cover - import guard
    import anydoc as _ANYDOC  # type: ignore
except Exception:  # noqa: BLE001 - any import failure -> fallback mode
    _ANYDOC = None

# Formats anydoc can convert (superset of Varan's own readers; includes
# legacy .doc/.ppt/.xls and OpenDocument/RTF/EPUB that Varan cannot read).
SUPPORTED_SUFFIXES = {
    ".doc", ".docx", ".docm",
    ".ppt", ".pps", ".pot", ".pptx", ".pptm", ".ppsx", ".ppsm",
    ".xls", ".xlsx", ".xlsm", ".xlsb",
    ".odt", ".ods", ".odp",
    ".rtf", ".epub", ".csv", ".pdf",
}


def available() -> bool:
    """True when the anydoc package is importable."""
    return _ANYDOC is not None


def supports(path: str | Path) -> bool:
    """True when anydoc can convert this file's extension."""
    return Path(path).suffix.lower() in SUPPORTED_SUFFIXES


def to_markdown(path: str | Path) -> str | None:
    """Return clean GitHub-Flavored Markdown for the file, or None when
    anydoc is unavailable or the conversion fails (e.g. scanned PDFs that
    need OCR — NeedsOcrError — or malformed files)."""
    if _ANYDOC is None or not supports(path):
        return None
    try:
        return _ANYDOC.to_markdown(str(path))
    except Exception:  # noqa: BLE001 - best-effort; caller falls back
        return None