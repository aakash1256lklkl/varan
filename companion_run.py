"""
Varan Companion launcher.

Launches the floating, always-on-top AI panel that sits beside Microsoft
Word / Excel / PowerPoint.

Run:  python companion_run.py
Toggle the panel with the global hotkey: Ctrl+Alt+Shift+V
"""
from __future__ import annotations

import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

from agent.compat import enable_utf8_stdio  # noqa: E402

enable_utf8_stdio()

from companion.window import main  # noqa: E402

if __name__ == "__main__":
    main()
