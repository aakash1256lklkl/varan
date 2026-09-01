"""
Varan shared state — lightweight persistence for things like the last
selected target file, so a selection survives companion/CLI restarts.
"""
from __future__ import annotations

import json
import os
from pathlib import Path

ROOT = Path(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_STATE_FILE = ROOT / "state.json"
_KEY_TARGET = "last_target_file"


class VaranState:
    def __init__(self, path: str | Path | None = None):
        self._path = Path(path) if path else _STATE_FILE
        self._data: dict = {}
        self._load()

    def _load(self):
        try:
            if self._path.exists():
                self._data = json.loads(self._path.read_text(encoding="utf-8"))
        except Exception:  # noqa: BLE001
            self._data = {}
        if not isinstance(self._data, dict):
            self._data = {}

    def save(self):
        try:
            self._path.write_text(
                json.dumps(self._data, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
        except Exception:  # noqa: BLE001
            pass

    def get_target_file(self) -> str | None:
        val = self._data.get(_KEY_TARGET)
        if val and Path(val).exists():
            return str(val)
        return None

    def set_target_file(self, path: str | None):
        if path and Path(path).exists():
            self._data[_KEY_TARGET] = str(Path(path).resolve())
        else:
            self._data.pop(_KEY_TARGET, None)
        self.save()
