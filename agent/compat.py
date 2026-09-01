"""
Varan console compatibility helpers.

On Windows, the default console code page (e.g. cp1252) cannot encode many
Unicode characters (arrows, emoji, math symbols) that model output and Varan
messages may contain. This reconfigures stdout/stderr to UTF-8 with
error-tolerant replacement so the CLI never crashes on a Unicode character.
"""
from __future__ import annotations

import sys


def enable_utf8_stdio() -> None:
    """Best-effort reconfiguration of stdout/stderr to UTF-8 on Windows."""
    if sys.platform != "win32":
        return
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            try:
                reconfigure(encoding="utf-8", errors="replace")
            except Exception:  # noqa: BLE001
                pass
