"""
Varan text / PDF editor — create, read, summarize and edit plain-text-like
documents: .pdf, .txt, .md (and other plain text extensions).

Reading/summarizing PDFs uses pypdf. Editing a PDF uses pypdf's text-stream
rewriting (safe for simple, auto-generated PDFs); editing .txt/.md is a
straightforward line-based text rewrite.

NOTE: PDF is a fixed page-layout format, so unlike the Office COM live editor
there is no true "live" in-place edit here. Editing a PDF rewrites its text
streams and saves to an '_edited' copy (or in place for the selected target).
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Optional

TEXT_EXTENSIONS = {".txt", ".md", ".markdown", ".rst", ".log", ".csv"}


class UnsupportedPdfEditError(Exception):
    """Raised when a PDF cannot be edited (no extractable text / no text streams)."""


class TextEditor:
    extension = ".txt"

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    @staticmethod
    def _is_pdf(path: str | Path) -> bool:
        return Path(path).suffix.lower() == ".pdf"

    def _read_text(self, path: str | Path) -> str:
        src = Path(path)
        if not src.exists():
            raise FileNotFoundError(src)
        for enc in ("utf-8", "utf-8-sig", "latin-1"):
            try:
                return src.read_text(encoding=enc)
            except (UnicodeDecodeError, UnicodeError):
                continue
        return src.read_bytes().decode("utf-8", errors="replace")

    # ------------------------------------------------------------------
    # Reading
    # ------------------------------------------------------------------
    def read(self, path: str | Path) -> dict:
        if self._is_pdf(path):
            return self._read_pdf(path)
        text = self._read_text(path)
        lines = [l.rstrip() for l in text.splitlines()]
        return {
            "path": str(Path(path)),
            "kind": "text",
            "line_count": len(lines),
            "content": text,
        }

    def _read_pdf(self, path: str | Path) -> dict:
        from pypdf import PdfReader, PdfWriter  # noqa: F401

        src = Path(path)
        if not src.exists():
            raise FileNotFoundError(src)
        reader = PdfReader(str(src))
        pages = []
        for i, page in enumerate(reader.pages):
            try:
                txt = page.extract_text() or ""
            except Exception:
                txt = ""
            pages.append({"page": i + 1, "text": txt.strip()})
        return {
            "path": str(src),
            "kind": "pdf",
            "page_count": len(reader.pages),
            "pages": pages,
        }

    # ------------------------------------------------------------------
    # Summarizing
    # ------------------------------------------------------------------
    def summarize(self, path: str | Path, max_points: int = 10) -> dict:
        data = self.read(path)
        if data.get("kind") == "pdf":
            full = "\n\n".join(p["text"] for p in data["pages"])
            lines = [l.strip() for l in full.splitlines() if l.strip()]
            word_count = len(re.findall(r"\b\w+\b", full))
            return {
                "title": Path(path).stem,
                "kind": "pdf",
                "page_count": data.get("page_count", 0),
                "word_count": word_count,
                "text_preview": "\n".join(lines)[:2000],
            }
        text = data.get("content", "")
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        word_count = len(re.findall(r"\b\w+\b", text))
        return {
            "title": Path(path).stem,
            "kind": "text",
            "line_count": data.get("line_count", 0),
            "word_count": word_count,
            "text_preview": "\n".join(lines)[:2000],
        }

    # ------------------------------------------------------------------
    # Editing (text files)
    # ------------------------------------------------------------------
    def edit(self, path: str | Path, edits: list[dict] | None = None,
             append: str | None = None, inplace: bool = False) -> str:
        """Edit a .txt/.md/plain-text file. Returns the path written.

        edits: list of {"action": "replace"|"delete"|"insert_after"|
                               "insert_before"|"delete_range",
                        "match": text to find (exact substring),
                        "end_match": required for delete_range,
                        "text": replacement / inserted content}
        """
        src = Path(path)
        if not src.exists():
            raise FileNotFoundError(src)
        text = self._read_text(src)
        for edit in edits or []:
            text = self._apply_edit(text, edit)
        if append is not None:
            text = text.rstrip() + "\n" + str(append) + "\n"
        dst = src if inplace else self._edited_copy(src)
        dst.parent.mkdir(parents=True, exist_ok=True)
        dst.write_text(text, encoding="utf-8")
        return str(dst)

    def _apply_edit(self, text: str, edit: dict) -> str:
        action = edit.get("action", "replace")
        match = str(edit.get("match", ""))
        new = str(edit.get("text", ""))
        end_match = str(edit.get("end_match", ""))

        if action == "delete":
            if not match:
                return text
            idx = text.find(match)
            if idx < 0:
                raise LookupError(f"match not found: {match!r}")
            if match in "\r\n":
                # deleting a bare newline: remove the line's ending
                return text[:idx] + text[idx + len(match):]
            # delete through the end of the line containing the match
            end = text.find("\n", idx)
            if end < 0:
                end = len(text)
            return text[:idx] + text[end:]

        if action == "delete_range":
            if not match:
                return text
            start = text.find(match)
            if start < 0:
                raise LookupError(f"start match not found: {match!r}")
            if end_match and end_match.strip():
                last = _last_index_of(text, end_match)
                if last < 0:
                    raise LookupError(f"end match not found: {end_match!r}")
                end = last + len(end_match)
            else:
                end = len(text)
            if end <= start:
                raise LookupError("delete range empty or inverted")
            # consume the newline right after the end anchor so the block is
            # removed cleanly (the anchor line's ending does not linger).
            if end < len(text) and text[end] == "\n":
                end += 1
            return text[:start] + text[end:]

        if action in ("insert_after", "insert_before"):
            if not match:
                return text
            idx = text.find(match)
            if idx < 0:
                raise LookupError(f"match not found: {match!r}")
            pos = idx + len(match) if action == "insert_after" else idx
            return text[:pos] + new + text[pos:]

        # replace
        if not match:
            return text
        count = int(edit.get("count", 1) or 1)
        if count == -1:
            if match not in text:
                raise LookupError(f"match not found: {match!r}")
            return text.replace(match, new)
        idx = -1
        remaining_idx = 0
        for _ in range(count):
            idx = text.find(match, remaining_idx)
            if idx < 0:
                break
            text = text[:idx] + new + text[idx + len(match):]
            remaining_idx = idx + len(new)
        if idx < 0:
            raise LookupError(f"match not found: {match!r}")
        return text

    # ------------------------------------------------------------------
    # Editing (PDFs) — rewrite text streams via pypdf
    # ------------------------------------------------------------------
    def edit_pdf(self, path: str | Path, edits: list[dict] | None = None,
                 inplace: bool = False) -> str:
        """Edit text within a PDF by rewriting its content-stream text.

        This is best-effort: it works well on simple / auto-generated PDFs that
        store text in clear content streams. Complex, scanned, or layout-heavy
        PDFs may not edit cleanly.
        """
        from pypdf import PdfReader, PdfWriter
        from pypdf.generic import NameObject, ArrayObject, DecodedStreamObject

        src = Path(path)
        if not src.exists():
            raise FileNotFoundError(src)
        reader = PdfReader(str(src))
        # clone_from properly attaches each page to the writer.
        writer = PdfWriter(clone_from=reader)

        for page in writer.pages:
            data = _get_page_contents_data(page)
            if data is None:
                continue
            content = data.decode("latin-1")
            new_content = content
            for edit in edits or []:
                action = edit.get("action", "replace")
                match = str(edit.get("match", ""))
                new = str(edit.get("text", ""))
                if not match:
                    continue
                if action == "delete":
                    if match in new_content:
                        new_content = new_content.replace(match, "", 1)
                elif action in ("replace",):
                    if match in new_content:
                        new_content = new_content.replace(match, new, 1)
            if new_content != content:
                # Replace /Contents with an uncompressed DecodedStreamObject.
                # (Mutating the original encoded stream is unreliable, and raw
                # StreamObject/bytes is rejected — this is the robust path.)
                dec = DecodedStreamObject()
                dec.set_data(new_content.encode("latin-1"))
                page[NameObject("/Contents")] = writer._add_object(dec)

        dst = src if inplace else self._edited_copy(src, ".pdf")
        dst.parent.mkdir(parents=True, exist_ok=True)
        with open(dst, "wb") as fh:
            writer.write(fh)
        return str(dst)

    # ------------------------------------------------------------------
    # Live / open detection (limited for text)
    # ------------------------------------------------------------------
    def get_headings(self, path: str | Path, limit: int = 100) -> list[str]:
        try:
            data = self.read(path)
        except Exception:
            return []
        lines: list[str] = []
        if data.get("kind") == "pdf":
            full = "\n".join(p["text"] for p in data.get("pages", []))
        else:
            full = data.get("content", "")
        for line in full.splitlines():
            s = line.strip()
            if s and (s.startswith(("#", "==", "**")) or len(s) < 90):
                lines.append(s)
            if len(lines) >= limit:
                break
        return lines

    def _edited_copy(self, src: Path, ext: str | None = None) -> Path:
        stem = src.stem
        return src.with_name(f"{stem}_edited{(ext or self.extension)}")


def _last_index_of(text: str, sub: str) -> int:
    return text.rfind(sub)


def _get_page_contents_data(page) -> bytes | None:
    """Return the concatenated decoded bytes of a page's /Contents stream(s),
    or None if the page has no content or it can't be decoded."""
    from pypdf.generic import ArrayObject
    try:
        entry = page.get("/Contents", None)
        if entry is None:
            return None
        obj = entry.get_object()
        streams = obj if isinstance(obj, ArrayObject) else [obj]
        parts = []
        for s in streams:
            try:
                if hasattr(s, "get_data"):
                    parts.append(s.get_data())
            except Exception:
                continue
        if not parts:
            return None
        return b"\n".join(parts)
    except Exception:
        return None
