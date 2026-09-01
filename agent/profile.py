"""
Varan user profile — persistent, human-authored "psychology + tasks" context.

This stores what the USER wants from Varan so the model can tailor every
response to them. Two parts:

  1. Psychology / preferences: who the user is, how they like to be spoken to
     (tone, verbosity, formality), and defaults (file/format habits, style).
  2. Task templates ("recipes"): named, reproducible tasks the user commonly
     requests, each with a description of the expected outcome so Varan
     understands what "done" looks like without re-explaining.

The profile is saved as JSON at varan/profile.json and injected into the model's
system prompt on every request (see agent/loop.py).
"""
from __future__ import annotations

import json
import os
from pathlib import Path

ROOT = Path(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_PROFILE_FILE = ROOT / "profile.json"

# Default/pre-filled task templates so the page isn't empty on first open.
_DEFAULT_TASK_TEMPLATES = [
    {
        "name": "Draft a screenplay / script",
        "goal": ("Write or add a screenplay/script section to the open document. "
                 "Use proper scene formatting (INT./EXT., Character:, dialogue), "
                 "a clear structure (title, logline, acts), and vivid direction."),
    },
    {
        "name": "Write a book / series bible",
        "goal": ("Produce organized book or series material: premise, characters, "
                 "plot outline, chapters. Keep the existing structure and append "
                 "or insert new sections rather than replacing."),
    },
    {
        "name": "Make a presentation / deck",
        "goal": ("Create an engaging slide deck: a strong title slide, clear "
                 "bullets (short, scannable), and logical flow. Mix layouts."),
    },
    {
        "name": "Build a spreadsheet / report",
        "goal": ("Build a tidy workbook: a header row, clean columns, useful "
                 "formulas (SUM/AVERAGE/percentages), and sensible sheet names."),
    },
]


def _default_profile() -> dict:
    return {
        "name": "",
        "role": "",
        # Psychology / preferences -----------------------------------------
        "tone": "professional",          # professional | friendly | casual | encouraging
        "verbosity": "balanced",         # brief | balanced | detailed
        "formality": "neutral",          # informal | neutral | formal
        "detail_level": "balanced",      # high-level | balanced | in-depth
        "structure_pref": "organized",   # organized | bullet-friendly | essay-like
        "follow_up": True,               # Varan should ask a clarifying question if unsure
        "edit_habit": "edit_in_place",   # edit_in_place | always_new_copy
        "extra_notes": "",
        # Task templates ---------------------------------------------------
        "task_templates": _DEFAULT_TASK_TEMPLATES,
    }


class UserProfile:
    def __init__(self, path: str | Path | None = None):
        self._path = Path(path) if path else _PROFILE_FILE
        self._load()

    def _load(self):
        if self._path.exists():
            try:
                data = json.loads(self._path.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    self._data = {**_default_profile(), **data}
                    self._data["task_templates"] = (
                        data.get("task_templates")
                        or _DEFAULT_TASK_TEMPLATES
                    )
                    return
            except Exception:  # noqa: BLE001
                pass
        self._data = _default_profile()

    # -- dict interface ----------------------------------------------------
    def get(self, key: str, default=None):
        return self._data.get(key, default)

    def __getitem__(self, key: str):
        return self._data[key]

    def __setitem__(self, key: str, value):
        self._data[key] = value

    def as_dict(self) -> dict:
        return dict(self._data)

    def save(self):
        try:
            self._path.write_text(
                json.dumps(self._data, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
        except Exception:  # noqa: BLE001
            pass

    # -- convenience -------------------------------------------------------
    def is_set(self) -> bool:
        """True if the user has given enough for the profile to be meaningful."""
        return bool(self._data.get("name")) or bool(self._data.get("extra_notes"))

    def to_prompt_block(self) -> str:
        """Render the profile into a concise system-prompt block the model reads
        so every request is tailored to this user."""
        d = self._data
        if not (d.get("name") or d.get("role") or d.get("tone")
                or d.get("extra_notes") or d.get("task_templates")):
            return ""
        lines: list[str] = ["[User profile — speak to and work for THIS user]"]
        if d.get("name"):
            lines.append(f"- User: {d['name']}")
        if d.get("role"):
            lines.append(f"- Role: {d['role']}")
        if d.get("tone"):
            lines.append(f"- Tone: {d['tone']}")
        if d.get("verbosity"):
            lines.append(f"- Verbosity: {d['verbosity']}")
        if d.get("formality"):
            lines.append(f"- Formality: {d['formality']}")
        if d.get("detail_level"):
            lines.append(f"- Detail level: {d['detail_level']}")
        if d.get("structure_pref"):
            lines.append(f"- Preferred structure: {d['structure_pref']}")
        if d.get("follow_up"):
            lines.append("- If a request is ambiguous, ask one clarifying question before acting.")
        if d.get("edit_habit"):
            lines.append(f"- Editing habit: {d['edit_habit']}")
        if d.get("extra_notes"):
            lines.append(f"- Notes: {d['extra_notes']}")

        templates = d.get("task_templates") or []
        named = [t for t in templates if (t or {}).get("name")]
        if named:
            lines.append("")
            lines.append("[Saved task recipes — follow these exactly when the user asks for them]")
            for t in named:
                lines.append(f"- \"{t['name']}\": {t.get('goal', '')}")
        return "\n".join(lines)


# Module-level accessors. We read a FRESH copy of the profile from disk on every
# call so a profile saved from the Companion (or edited in profile.json) is
# picked up immediately by the agent on the very next request — no stale cache.
def get_profile() -> UserProfile:
    return UserProfile()


def profile_prompt_block() -> str:
    try:
        return get_profile().to_prompt_block()
    except Exception:  # noqa: BLE001
        return ""
