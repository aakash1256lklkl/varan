"""
Varan Companion — a floating, always-on-top AI panel that sits beside
Microsoft Word / Excel / PowerPoint.

You toggle it with a global hotkey (default Ctrl+Alt+Shift+V), pick the
Office file you're working on (optional), and type natural-language prompts.
Varan reads, edits, summarizes or creates .docx / .xlsx / .pptx files and
writes results to the project outputs/ folder.

Runs the existing Varan agent engine in a background thread so the UI stays
responsive even during long model calls.
"""
from __future__ import annotations

import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext

ROOT = Path(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, str(ROOT))

from agent.config import load_config, CONFIG_PATH, PRESETS  # noqa: E402
from agent.providers import provider_factory  # noqa: E402
from agent.loop import Agent  # noqa: E402
from agent.profile import UserProfile  # noqa: E402

TOGGLE_HOTKEY = "<ctrl>+<alt>+<shift>+v"

PROVIDER_OPTIONS = sorted(PRESETS.keys()) + ["custom…"]

# Suggested models per provider (editable; you can type any model name).
CUSTOM_MODEL, *MODEL_SUGGESTIONS = [
    "custom…",
    "gpt-4o-mini", "gpt-4o", "gpt-3.5-turbo", "claude-3-5-sonnet-latest",
    "claude-3-haiku", "gemini-1.5-pro", "gemini-1.5-flash",
    "llama3.2", "llama3.1", "qwen2.5:3b", "qwen3:4b", "mistral-small-latest",
]


class CompanionApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Varan Companion")
        self.root.geometry("420x560")
        self.root.attributes("-topmost", True)

        cfg = load_config(CONFIG_PATH)
        self.agent = Agent(provider_factory(cfg), cfg, console=None)

        self._running = False  # a worker thread is active
        self._hotkey_listener = None
        self._hidden = False
        # A previously selected file is restored by the agent; mirror it here
        # so the status label and title reflect the active target.
        self._target_file = self.agent.target_file

        # User profile (psychology + tasks) shared with the agent.
        self.profile = UserProfile()

        self._build_ui()

        # Global hotkey toggle
        self._start_hotkey()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------
    # Model list helpers
    # ------------------------------------------------------------------
    def _available_models(self, provider: str) -> list[str]:
        """Return suggested model names for a provider.

        For local Ollama we query the running server for the real installed
        models; otherwise return the generic suggestion list.
        """
        if provider == "ollama":
            try:
                out = subprocess.run(
                    ["ollama", "list"],
                    capture_output=True, text=True, timeout=15,
                ).stdout or ""
                names = []
                for line in out.splitlines()[1:]:  # skip header
                    col = line.split()
                    if col:
                        names.append(col[0])
                if names:
                    return names
            except Exception:  # noqa: BLE001
                pass
        return list(MODEL_SUGGESTIONS)

    def _refresh_model_list(self, provider: str):
        self._model_options = self._available_models(provider)
        cur = self.agent.cfg.get("model", "")
        values = list(self._model_options)
        if cur and cur not in values:
            values.insert(0, cur)
        self.model_combo.configure(values=values)
        self.model_var.set(cur or "")

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        root = self.root

        top = tk.Frame(root, padx=8, pady=6)
        top.pack(fill="x")

        self.cfg_label = tk.Label(top, text=self._status_text(), anchor="w", fg="#444")
        self.cfg_label.pack(side="left", fill="x", expand=True)

        self.topmost_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            top, text="Always on top", variable=self.topmost_var,
            command=self._toggle_topmost,
        ).pack(side="right")

        # Output log
        self.log = scrolledtext.ScrolledText(
            root, wrap="word", height=20, state="disabled",
            font=("Consolas", 10), bg="#0f1720", fg="#d7e3f4",
            insertbackground="white",
        )
        self.log.pack(fill="both", expand=True, padx=8, pady=(0, 6))

        # Provider + Model dropdowns
        selrow = tk.Frame(root, padx=8)
        selrow.pack(fill="x", pady=(0, 4))

        tk.Label(selrow, text="Provider:").pack(side="left")
        self.provider_var = tk.StringVar()
        self.provider_combo = ttk.Combobox(
            selrow, textvariable=self.provider_var, state="readonly",
            values=PROVIDER_OPTIONS, width=13,
        )
        self.provider_combo.current(self._provider_index())
        self.provider_combo.bind("<<ComboboxSelected>>", self._on_provider_selected)
        self.provider_combo.pack(side="left", padx=(4, 12))

        tk.Label(selrow, text="Model:").pack(side="left")
        self.model_var = tk.StringVar()
        self.model_combo = ttk.Combobox(
            selrow, textvariable=self.model_var, width=20,
        )
        self.model_combo.pack(side="left", padx=(4, 6))
        self.model_combo.bind("<<ComboboxSelected>>", self._on_model_selected)
        self.model_combo.bind("<Return>", self._on_model_selected)

        # Quick actions row
        actions = tk.Frame(root, padx=8)
        actions.pack(fill="x")
        tk.Button(actions, text="Open Office file", command=self._pick_file).pack(side="left", padx=2)
        tk.Button(actions, text="Clear file", command=self._clear_target_file).pack(side="left", padx=2)
        tk.Button(actions, text="New provider…", command=self._new_provider_dialog).pack(side="left", padx=2)
        tk.Button(actions, text="Profile", command=self._open_profile).pack(side="left", padx=2)
        tk.Button(actions, text="Outputs", command=self._open_outputs).pack(side="left", padx=2)
        tk.Button(actions, text="Clear", command=self._clear_log).pack(side="right", padx=2)

        self._refresh_model_list(self.agent.cfg.get("provider", ""))

        # Prompt entry
        bottom = tk.Frame(root, padx=8, pady=8)
        bottom.pack(fill="x")
        self.entry_var = tk.StringVar()
        self.entry = tk.Entry(bottom, textvariable=self.entry_var, font=("Segoe UI", 11))
        self.entry.pack(side="left", fill="x", expand=True, padx=(0, 6), ipady=6)
        self.entry.bind("<Return>", lambda e: self._submit())

        self.run_btn = tk.Button(
            bottom, text="Run", command=self._submit, bg="#2f8c3e", fg="white",
            activebackground="#3aa349", font=("Segoe UI", 10, "bold"),
        )
        self.run_btn.pack(side="right")

        self.hint = tk.Label(
            root, text=f"Hotkey: {TOGGLE_HOTKEY} to show/hide this panel",
            fg="#888", anchor="w", padx=8,
        )
        self.hint.pack(fill="x", pady=(0, 4))

        self.entry.focus_set()

    def _status_text(self) -> str:
        c = self.agent.cfg
        target = self._target_file or "none"
        return f"Provider: {c.get('provider')}  Model: {c.get('model')}\nOpen file: {target}"

    def _toggle_topmost(self):
        self.root.attributes("-topmost", self.topmost_var.get())

    # ------------------------------------------------------------------
    # Logging helpers
    # ------------------------------------------------------------------
    def _append(self, text, tag=None):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n", tag or ())
        self.log.see("end")
        self.log.configure(state="disabled")

    def _info(self, text):
        self._append(text, "info")
        self.log.tag_config("info", foreground="#8fd3ff")

    def _ok(self, text):
        self._append(text, "ok")
        self.log.tag_config("ok", foreground="#7be37b")

    def _err(self, text):
        self._append(text, "err")
        self.log.tag_config("err", foreground="#ff8a8a")

    def _tool(self, text):
        self._append(text, "tool")
        self.log.tag_config("tool", foreground="#ffd479")

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    # ------------------------------------------------------------------
    # Prompt execution
    # ------------------------------------------------------------------
    def _submit(self):
        prompt = self.entry_var.get().strip()
        if not prompt or self._running:
            return
        self._clear_log()
        self._info("You: " + prompt)
        self._set_busy(True)
        thread = threading.Thread(target=self._worker, args=(prompt,), daemon=True)
        thread.start()
        self.entry_var.set("")

    def _set_busy(self, busy):
        self._running = busy
        self.run_btn.config(state="disabled" if busy else "normal", text="…" if busy else "Run")
        self.entry.config(state="disabled" if busy else "normal")

    def _worker(self, prompt):
        # agent.run is blocking; deliver events back to the GUI thread
        self.agent.run(prompt, sink=self._on_event)
        self.root.after(0, lambda: self._set_busy(False))

    def _on_event(self, kind, payload):
        # Called from a worker thread -> marshal to GUI thread
        self.root.after(0, lambda: self._handle_event(kind, payload))

    def _handle_event(self, kind, payload):
        if kind == "tool":
            name, args = payload
            path = args.get("path") if isinstance(args, dict) else None
            self._tool(f"→ {name}" + (f"  {path}" if path else ""))
        elif kind == "status":
            self._info(payload)
        elif kind == "text":
            self._ok("Varan: " + payload)
        elif kind == "result":
            path = payload.get("path")
            if path:
                self._ok("✓ Saved: " + path)
        elif kind == "error":
            self._err(payload)

    # ------------------------------------------------------------------
    # Target file / outputs
    # ------------------------------------------------------------------
    def _open_in_native_app(self, path: str):
        """Open a file in its default native application (Word/Excel/PowerPoint)
        so the user sees it live beside the Companion, ready for live edits."""
        path = str(path)
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # opens with its registered app (Word, etc.)
            else:
                subprocess.Popen(["xdg-open", path])
            return True
        except Exception as exc:  # noqa: BLE001
            self._info(f"Could not open file in its app: {exc}")
            return False

    def _pick_file(self):
        path = filedialog.askopenfilename(
            title="Select the Office file you're working on",
            filetypes=[
                ("Office files", "*.docx *.xlsx *.pptx *.doc *.xls *.ppt"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self._target_file = path
            self.agent.set_target_file(path)
            self.cfg_label.config(text=self._status_text())
            self._info(f"Target file set: {path}")
            # Also open the file in its native app so Varan can edit it LIVE
            # beside the Companion window.
            if self._open_in_native_app(path):
                self._info(f"Opened in its app for live editing: {os.path.basename(path)}")

    def _clear_target_file(self):
        self._target_file = None
        self.agent.clear_target_file()
        self.cfg_label.config(text=self._status_text())
        self._info("Target file cleared — new creates will make separate files.")

    def _open_outputs(self):
        outputs = ROOT / "outputs"
        outputs.mkdir(parents=True, exist_ok=True)
        subprocess.Popen(["explorer", str(outputs)])

    # ------------------------------------------------------------------
    # User profile (psychology + tasks)
    # ------------------------------------------------------------------
    _COMBO_OPTS = {
        "tone": ["professional", "friendly", "casual", "encouraging"],
        "verbosity": ["brief", "balanced", "detailed"],
        "formality": ["informal", "neutral", "formal"],
        "detail_level": ["high-level", "balanced", "in-depth"],
        "structure_pref": ["organized", "bullet-friendly", "essay-like"],
        "follow_up": None,  # bool checkbox
        "edit_habit": ["edit_in_place", "always_new_copy"],
    }

    def _open_profile(self):
        """Open a form where the user tells Varan how they like to work and the
        tasks they commonly ask for. Saved to profile.json and injected into the
        model's system prompt on every request."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Varan — User Profile & Tasks")
        dlg.geometry("640x720")
        dlg.transient(self.root)
        dlg.grab_set()

        canvas = tk.Canvas(dlg)
        vsb = ttk.Scrollbar(dlg, orient="vertical", command=canvas.yview)
        frm = ttk.Frame(canvas)
        frm.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=frm, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))

        pad = {"padx": 12, "pady": 4, "sticky": "w"}
        d = self.profile.as_dict()

        def add_fields(container, header, entries):
            ttk.Label(container, text=header, font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=6, pady=(10, 2))
            for key, label_text in entries:
                row = ttk.Frame(container)
                row.pack(fill="x", **pad)
                ttk.Label(row, text=label_text, width=18).pack(side="left")
                if self._COMBO_OPTS.get(key):
                    var = tk.StringVar(value=str(d.get(key, "")))
                    ttk.Combobox(row, textvariable=var, values=self._COMBO_OPTS[key],
                                 state="readonly", width=22).pack(side="left")
                else:
                    var = tk.StringVar(value=str(d.get(key, "") or ""))
                    ttk.Entry(row, textvariable=var, width=34).pack(side="left")
                self._profile_vars[key] = var

        self._profile_vars = {}
        add_fields(frm, "Who you are",
                   [("name", "Your name"), ("role", "Your role / job")])

        add_fields(frm, "How I should talk to you",
                   [("tone", "Tone"), ("verbosity", "Verbosity"),
                    ("formality", "Formality"), ("detail_level", "Detail level"),
                    ("structure_pref", "Structure")])

        # Edit habit + follow-up
        row = ttk.Frame(frm); row.pack(fill="x", **pad)
        ttk.Label(row, text="Editing habit", width=18).pack(side="left")
        eh = tk.StringVar(value=str(d.get("edit_habit", "edit_in_place")))
        ttk.Combobox(row, textvariable=eh, values=self._COMBO_OPTS["edit_habit"],
                     state="readonly", width=22).pack(side="left")
        self._profile_vars["edit_habit"] = eh

        row = ttk.Frame(frm); row.pack(fill="x", **pad)
        self._follow_up_var = tk.BooleanVar(value=bool(d.get("follow_up", True)))
        ttk.Checkbutton(row, text="Ask one clarifying question if a request is ambiguous",
                        variable=self._follow_up_var).pack(anchor="w")

        ttk.Label(frm, text="Extra notes Varan should always remember",
                  font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=6, pady=(10, 2))
        self._notes_text = tk.Text(frm, height=4, width=70, wrap="word")
        self._notes_text.insert("1.0", d.get("extra_notes", "") or "")
        self._notes_text.pack(fill="x", **pad)

        # Task templates --------------------------------------------------
        ttk.Label(frm, text="Tasks you give Varan (recipes it follows on request)",
                  font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=6, pady=(10, 2))
        ttk.Label(frm, text="You can edit the goals so Varan understands exactly what \"done\" "
                            "means. When you ask for a recipe by name, Varan follows it.",
                  wraplength=560, foreground="#555").pack(anchor="w", padx=6)
        self._task_frame = ttk.Frame(frm)
        self._task_frame.pack(fill="x", pady=4)
        self._render_tasks()

        ttk.Button(frm, text="+ Add a task", command=self._add_task_row).pack(anchor="w", padx=6, pady=4)

        # Footer buttons
        foot = ttk.Frame(frm)
        foot.pack(fill="x", pady=10)
        ttk.Button(foot, text="Save profile", command=lambda: self._save_profile(dlg)).pack(side="left", padx=6)
        ttk.Button(foot, text="Cancel", command=dlg.destroy).pack(side="left", padx=6)
        ttk.Button(foot, text="Reset to defaults", command=lambda: self._reset_profile(dlg)).pack(side="right", padx=6)

    def _render_tasks(self):
        for w in self._task_frame.winfo_children():
            w.destroy()
        self._task_entries = []
        for t in self.profile.get("task_templates") or []:
            self._add_task_row(t)

    def _add_task_row(self, task=None):
        task = task or {"name": "", "goal": ""}
        row = ttk.Frame(self._task_frame)
        row.pack(fill="x", pady=2)
        name_var = tk.StringVar(value=task.get("name", "") or "")
        goal_var = tk.StringVar(value=task.get("goal", "") or "")
        ttk.Entry(row, textvariable=name_var, width=16).pack(side="left", padx=(0, 4))
        ttk.Entry(row, textvariable=goal_var).pack(side="left", fill="x", expand=True, padx=(0, 4))
        ttk.Button(row, text="✕", width=2, command=lambda r=row: r.destroy()).pack(side="left")
        self._task_entries.append((row, name_var, goal_var))

    def _collect_task_templates(self):
        templates = []
        for _row, name_var, goal_var in getattr(self, "_task_entries", []):
            name = name_var.get().strip()
            goal = goal_var.get().strip()
            if name or goal:
                templates.append({"name": name or "Untitled", "goal": goal})
        return templates

    def _save_profile(self, dlg):
        p = self.profile
        for key, var in self._profile_vars.items():
            p[key] = var.get()
        p["follow_up"] = bool(self._follow_up_var.get())
        p["extra_notes"] = self._notes_text.get("1.0", "end").strip()
        p["task_templates"] = self._collect_task_templates()
        p.save()
        self._ok("Profile saved — Varan will tailor responses to it.")
        dlg.destroy()

    def _reset_profile(self, dlg):
        if not messagebox.askyesno("Profile", "Reset all profile fields to defaults?"):
            return
        from agent.profile import UserProfile
        blank = UserProfile()
        blank.save()
        # Reload in place and re-open the dialog.
        dlg.destroy()
        self.profile = UserProfile()
        self._open_profile()
        self._info("Profile reset. Re-open Profile to edit defaults.")

    # ------------------------------------------------------------------
    # Provider / model
    # ------------------------------------------------------------------
    def _provider_index(self) -> int:
        cur = self.agent.cfg.get("provider", "")
        opts = list(PROVIDER_OPTIONS)
        try:
            return opts.index(cur)
        except ValueError:
            # Custom provider not in preset list -> point at the "custom…" entry
            return opts.index("custom…")

    def _on_provider_selected(self, _event=None):
        label = self.provider_var.get()
        if label == "custom…":
            # Reset the dropdown to the actual provider so the user can re-open
            # the dialog or pick a preset.
            self.provider_combo.current(self._provider_index())
            self._new_provider_dialog()
            return
        try:
            self.agent.change_provider(label)
        except Exception as exc:  # noqa: BLE001
            self._err(f"Provider error: {exc}")
            self.provider_combo.current(self._provider_index())
            return
        self._refresh_model_list(label)
        self.cfg_label.config(text=self._status_text())
        self._ok(f"Provider → {label} (model: {self.agent.cfg.get('model')})")

    def _on_model_selected(self, _event=None):
        model = self.model_var.get().strip()
        if not model or model == "custom…":
            model = self._model_entry_dialog()
            if not model:
                self._refresh_model_list(self.agent.cfg.get("provider", ""))
                return
        try:
            self.agent.change_model(model)
        except Exception as exc:  # noqa: BLE001
            self._err(f"Model error: {exc}")
            return
        self.model_var.set(model)
        self.cfg_label.config(text=self._status_text())
        self._ok(f"Model → {model}")

    def _model_entry_dialog(self) -> str:
        dialog = tk.Toplevel(self.root)
        dialog.title("Set model")
        dialog.geometry("320x120")
        dialog.transient(self.root)
        dialog.attributes("-topmost", True)
        tk.Label(dialog, text="Enter any model name:").pack(pady=(10, 2))
        var = tk.StringVar(value=self.agent.cfg.get("model", ""))
        e = tk.Entry(dialog, textvariable=var)
        e.pack(padx=12, fill="x")
        e.focus_set()
        result = {"value": ""}

        def ok():
            result["value"] = var.get().strip()
            dialog.destroy()

        def cancel():
            dialog.destroy()

        tk.Button(dialog, text="OK", command=ok).pack(side="left", padx=20, pady=8)
        tk.Button(dialog, text="Cancel", command=cancel).pack(side="right", padx=20, pady=8)
        e.bind("<Return>", lambda ev: ok())
        dialog.bind("<Escape>", lambda ev: cancel())
        dialog.grab_set()
        self.root.wait_window(dialog)
        return result["value"]

    def _new_provider_dialog(self):
        """Form to add/edit a provider: name, base URL, model, API key."""
        dialog = tk.Toplevel(self.root)
        dialog.title("New AI Provider")
        dialog.geometry("460x300")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.attributes("-topmost", True)

        cfg = self.agent.cfg
        fields = {}

        def row(label, key, default="", show=None):
            f = tk.Frame(dialog, padx=12, pady=4)
            f.pack(fill="x")
            tk.Label(f, text=label, width=14, anchor="w").pack(side="left")
            var = tk.StringVar(value=default)
            e = tk.Entry(f, textvariable=var, show=show or "")
            e.pack(side="left", fill="x", expand=True)
            fields[key] = var
            return e

        row("Name", "provider", cfg.get("provider", ""))
        row("Base URL", "base_url", cfg.get("base_url", ""))
        row("Model", "model", cfg.get("model", ""))
        row("API key", "api_key", cfg.get("api_key", ""), show="*")
        hint = tk.Label(
            dialog,
            text="Base URL is the API endpoint, e.g. https://api.openai.com/v1\n"
                 "Use an OpenAI-compatible endpoint for most providers.",
            fg="#888", justify="left", padx=12,
        )
        hint.pack(fill="x", pady=(2, 6))

        def save():
            provider = fields["provider"].get().strip()
            if not provider:
                return
            base_url = fields["base_url"].get().strip()
            model = fields["model"].get().strip()
            api_key = fields["api_key"].get().strip()
            try:
                self.agent.change_provider_full(provider, model, base_url, api_key)
            except Exception as exc:  # noqa: BLE001
                self._err(f"Provider error: {exc}")
                return
            # Refresh the provider dropdown so the new provider is visible.
            self._refresh_provider_dropdown()
            self._refresh_model_list(provider)
            self.cfg_label.config(text=self._status_text())
            self._ok(f"Added provider → {provider} (model: {model or '(default)'})")
            dialog.destroy()

        def cancel():
            dialog.destroy()

        be = tk.Frame(dialog)
        be.pack(fill="x", pady=8)
        tk.Button(be, text="Save provider", command=save).pack(side="left", padx=20)
        tk.Button(be, text="Cancel", command=cancel).pack(side="right", padx=20)
        dialog.grab_set()

    def _refresh_provider_dropdown(self):
        self.provider_combo.configure(values=list(PROVIDER_OPTIONS))
        self.provider_combo.current(self._provider_index())

    # ------------------------------------------------------------------
    # Hotkey toggle (global)
    # ------------------------------------------------------------------
    def _start_hotkey(self):
        try:
            from pynput import keyboard
            from pynput.keyboard import GlobalHotKeys
        except Exception as exc:  # noqa: BLE001
            self._info(f"Global hotkey unavailable: {exc}")
            return

        # GlobalHotKeys accepts <ctrl>+<alt>+<shift>+v style combos
        def _do_toggle():
            self.root.after(0, self._toggle_visibility)
        hotkeys = GlobalHotKeys({TOGGLE_HOTKEY: _do_toggle})
        self._hk_listener = hotkeys
        hotkeys.start()

    def _toggle_visibility(self):
        if self._hidden:
            self.root.deiconify()
            self.root.lift()
            self.root.attributes("-topmost", self.topmost_var.get())
            self._hidden = False
        else:
            self.root.withdraw()
            self._hidden = True

    def _on_close(self):
        # Closing hides to tray-like behavior instead of killing
        if not self._running:
            if messagebox.askyesno("Varan Companion", "Quit Varan Companion?"):
                self.root.destroy()
            # else keep running
        else:
            self.root.withdraw()
            self._hidden = True

    # ------------------------------------------------------------------
    def mainloop(self):
        self.root.mainloop()


def main():
    app = CompanionApp()
    app.mainloop()


if __name__ == "__main__":
    main()
