"""
Varan agent loop — runs the prompt→tool-call→execute→reply cycle.
"""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Optional

from rich.console import Console

from .providers import provider_factory, BaseProvider, ProviderError
from .tools import TOOLS, ToolExecutor
from .state import VaranState
from .profile import profile_prompt_block

ROOT = Path(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
OUTPUTS_DIR = ROOT / "outputs"

SYSTEM_PROMPT = """You are Varan, an AI assistant that creates, edits, reads and summarizes
Microsoft Office files (Word .docx, Excel .xlsx, PowerPoint .pptx), PDF (.pdf)
and plain-text (.txt, .md) documents.

You have tools to perform these actions on real files. Follow these rules:
1. Parse the user's request and decide which tool(s) to call. You may need
   MULTIPLE tool calls for complex requests (e.g. create a doc, then a deck from it).
2. Always provide a 'path' argument. If it has no folder, it will go to the
   outputs/ folder automatically. Use descriptive filenames ending in the right
   extension (.docx / .xlsx / .pptx / .pdf / .txt / .md).
3. When creating slides, tables or charts, produce realistic, well-structured
   content. For presentations, use a mix of 'title', 'bullets', and 'blank' layouts.
4. After each tool call, the tool result is returned to you. Use it to craft a
   friendly final message telling the user what you created and the file path.
5. IMPORTANT — if a target file is open ([Target file: ...] in the request):
   treat any request that produces or adds content (e.g. "write a script",
   "add a section", "create content in the selected document") as an EDIT to
   THAT target file using edit_document / edit_workbook / edit_presentation /
   edit_text, NOT a create of a new file. Only create a brand-new file when the
   user explicitly asks for a separate/new document that is not the open target.
    When editing the target, keep its existing content and add/insert the new
    content rather than replacing everything. EXCEPTION: if the user explicitly
    asks to rewrite/retheme/replace the deck's slides or change every slide
    (a PowerPoint request), use 'rebuild_slides' per rule 11 instead of appending.
6. For PDFs, use edit_text to do text find/replace/delete (best-effort on simple
   PDFs). For plain text (.txt/.md) use edit_text. For a document that is open
   in its Office app (Word/Excel/PowerPoint), edits happen live in the open file.
7. DESTRUCTIVE CONFIRMATION — deleting (edit_document action 'delete' or
   'delete_range', or the 'page' parameter; likewise edit_text deletes) removes
   content permanently and cannot be undone. BEFORE such a delete:
   a. call get_paragraphs (or read_file) to see exactly what is on the page /
      in the section / around the anchor;
   b. tell the user precisely what will be removed ("this will delete the
      'ARC 2' section: <first line … last line>") and ASK them to confirm which
      section/page they mean, especially if the request is ambiguous (e.g.
      'delete this page' when one page holds several sections, or a heading that
      appears more than once);
   c. only call the delete tool AFTER the user confirms a specific anchor.
   Never delete content the user has not explicitly confirmed.
8. BLANK PAGES are typically caused by empty paragraphs (especially empty
   Heading 1 paragraphs used as page separators). Use 'remove_blank_pages: true'
   in edit_document to remove them — this is SAFE (only empty paragraphs are
   removed, no real content) and does NOT require confirm. If the user says
   'delete blank pages', 'remove empty pages', or 'clean up blank sections',
   use remove_blank_pages immediately — do NOT ask for confirmation on this
   operation. Call get_paragraphs first to show the empty-paragraph count, then
   call edit_document with remove_blank_pages: true and inplace: true.
9. STYLED ADDITIONS — when the user wants a styled section appended to an
   existing Word document (real headings, bullets, italics, a specific font like
   Courier, scene breaks), call edit_document with 'content' — an array of block
   dicts {'type': 'heading'|'title'|'paragraph'|'bullet'|'numbered'|'divider',
   'level', 'text', 'bold', 'italic', 'font', 'courier'}. Do NOT flatten blocks
   to plain text or write literal '##' / '-' markdown characters into the .docx:
   those show up verbatim in the file and produce unformatted garbage. Never
   claim styling was applied (Courier, small caps, headings, italics) unless the
   write actually used 'content' (or a live edit).
 10. BE CONCISE AND DIRECT — after a successful edit, report what changed in 1-3
    sentences (what was added/removed/style-applied and where). Do NOT ramble,
    do NOT offer the user a menu of options ('Option A vs Option B'); just do the
    most sensible thing and tell them. If a request genuinely cannot be done,
    say why in ONE sentence and do the closest safe alternative.
11. POWERPOINT: THREE MODES — pick the RIGHT one instead of always rebuilding:
      * SURGICAL ('edit_slides'): use when only SOME slide(s) change — rename one
        title, fix text on a single slide, replace a few bullets, or ADD a
        textbox/table/chart to a specific slide (keys 'add_textbox',
        'add_table', 'add_chart' inside the edit_slides item), or DELETE one or
        more specific slides (pass 'remove_slides': [1-based indices]).
        Each edit_slides item is {'slide': 1-based index, 'title': new title,
        'replace': {old: new}, 'set_bullets': [...], 'add_textbox': {...},
        'add_table': {...}, 'add_chart': {...}}. This is FAST and edits in
        place — use it for most "change slide 2's title" / "fix the typo on
        slide 4" / "add a table to slide 5" / "delete slide 3" requests.
      * APPEND ('add_slides'): use ONLY to stack genuinely NEW slides onto the
        deck. It does NOT touch existing slides.
      * REBUILD ('rebuild_slides'): use ONLY when the user asks to rewrite EVERY
        slide / change every slide's template / retheme the whole deck to match a
        topic ('make every slide about X', 'change every slide template',
        'rebuild the deck'). Rebuilds are expensive — do NOT use them to change
        just one or two slides.
After a PPT edit, report the tool's 'summary' (mode surgical/append/rebuild,
       how many slides affected) — never claim slides were rewritten or rethemed
       unless the tool result shows the matching mode.
       FILLING EMPTY SLIDES — when a deck already has section titles but the text
       boxes / body placeholders are empty, DO NOT rebuild the deck from scratch
       (and do not append new slides). Use SURGICAL 'edit_slides' with
       'set_bullets': [real bullets] and 'add_table': {headers, rows} to fill
       each numbered slide in place.
       NEVER WRITE PLACEHOLDERS — never emit placeholder filler as slide content:
       bullets like "- item" / "item", or tables whose cells read "item" /
       "i | t | e | m". The editor REFUSES such content, so write every bullet
       and cell as REAL information. If you cannot produce real content from the
       file and the user's request, ASK the user instead of inventing filler.
 12. REPORT TRUTHFULLY — read the tool result (especially the 'ok', 'error',
    'summary' fields). 'ok: true' means the tool ran and a file was written; it
    does not by itself mean a specific styling/rewrite happened. Check the
    returned summary/metadata before describing the outcome, and describe only
    what the tool result proves. If a tool returns an 'error', report that back
    and fix the input, do not narrate a fictional success.
Never invent fake confirmations, and never tell the user something was applied
when the tool result does not prove it.
Be concise and helpful."""


class Agent:
    def __init__(self, provider: BaseProvider, cfg: dict, console: Optional[Console] = None):
        self.provider = provider
        self.cfg = cfg
        self.console = console or Console()
        strict = bool(
            os.environ.get("VARAN_STRICT") or cfg.get("strict") or cfg.get("tool_strict")
        )
        self.executor = ToolExecutor(OUTPUTS_DIR, strict=strict)
        self.history: list[dict] = []
        self.max_tool_rounds = 8
        # Restore the last selected target file so a previous "Open file"
        # selection survives restarts of the companion / CLI.
        self._state = VaranState()
        self.target_file: str | None = self._state.get_target_file()

    # -- config switching ---------------------------------------------------
    def change_provider(self, name: str):
        cfg = dict(self.cfg)
        cfg["provider"] = name
        from .config import PRESETS
        preset = PRESETS.get(name)
        if preset:
            # For a known preset we always take the preset's default endpoint so
            # a stale base_url from a previous provider never leaks in (e.g. an
            # Ollama URL kept while the provider claims to be "groq").
            cfg["base_url"] = preset["base_url"]
            if not cfg.get("model") or cfg["model"] in ("", ):
                cfg["model"] = preset["model"]
        else:
            # Not a known preset (a manual/custom name): keep any manually
            # configured base URL, or default the endpoint to the name.
            if not cfg.get("base_url"):
                cfg["base_url"] = name
        if not cfg.get("model"):
            cfg["model"] = name
        new_provider = provider_factory(cfg)
        self.provider.close()
        self.provider = new_provider
        self.cfg = cfg
        from .config import save_prefs
        save_prefs(name, cfg.get("model", ""), cfg.get("base_url", ""))
        return cfg

    def change_provider_full(self, provider: str, model: str = "", base_url: str = "",
                             api_key: str = "") -> dict:
        """Switch to a provider with explicit settings (for custom providers).

        Builds a fresh config from the given values, reconstructs the provider,
        and persists the settings so they survive a restart.
        """
        cfg = dict(self.cfg)
        cfg["provider"] = provider
        if base_url:
            cfg["base_url"] = base_url
        if model:
            cfg["model"] = model
        if api_key:
            cfg["api_key"] = api_key
        if not cfg.get("base_url"):
            cfg["base_url"] = provider
        if not cfg.get("model"):
            cfg["model"] = provider
        new_provider = provider_factory(cfg)
        self.provider.close()
        self.provider = new_provider
        self.cfg = cfg
        from .config import save_prefs
        save_prefs(provider, cfg.get("model", ""), cfg.get("base_url", ""), api_key)
        return cfg

    def change_model(self, name: str):
        self.cfg = dict(self.cfg)
        self.cfg["model"] = name
        new_provider = provider_factory(self.cfg)
        self.provider.close()
        self.provider = new_provider
        return self.cfg

    def set_target_file(self, path: str):
        self.target_file = path
        self._state.set_target_file(path)

    def clear_target_file(self):
        self.target_file = None
        self._state.set_target_file(None)

    def list_outputs(self):
        files = self.executor._list_outputs({}).get("files", [])
        if not files:
            self.console.print("[dim](outputs folder is empty)[/dim]")
            return
        self.console.print("[bold]outputs/:[/bold]")
        for f in files:
            self.console.print(f"  {f}")

    # -- target-file aware routing -----------------------------------------
    @staticmethod
    def _kind_of(path: str) -> str:
        return Path(path or "").suffix.lower()

    def _remap_create_to_edit(self, name: str, args: dict):
        """When a target file is open, turn a "create" of the same file kind
        into an "edit" of that selected file, so content the user asked to put
        "in the selected document" actually lands there instead of a new file.

        Returns (edit_name, edit_args) to execute instead, or None to leave
        the tool call unchanged (e.g. a truly new/separate file or a kind
        that doesn't match the open target).
        """
        if not self.target_file:
            return None
        kind = self._kind_of(self.target_file)
        create_to_edit = {
            ".docx": ("create_document", "edit_document"),
            ".doc": ("create_document", "edit_document"),
            ".pptx": ("create_presentation", "edit_presentation"),
            ".ppt": ("create_presentation", "edit_presentation"),
            ".xlsx": ("create_workbook", "edit_workbook"),
            ".xls": ("create_workbook", "edit_workbook"),
            ".csv": ("create_workbook", "edit_workbook"),
        }
        mapping = create_to_edit.get(kind)
        if not mapping or name != mapping[0]:
            return None
        _, edit_name = mapping

        if edit_name == "edit_document":
            blocks: list[dict] = [b for b in (args.get("body") or []) if isinstance(b, dict)]
            if args.get("title"):
                blocks = [{"type": "title", "text": str(args["title"])}] + blocks
            if not blocks:
                return None
            # Forward the model's STRUCTURED blocks as 'content' so the editor
            # renders REAL Word heading/bullet/italic/font styling. Flattening to
            # plain text (literal '##'/'-' markdown inside the doc) is the historic
            # bug that made "add a styled section" silently appear unformatted.
            return (edit_name, {
                "path": self.target_file, "content": blocks, "inplace": True,
            })

        if edit_name == "edit_presentation":
            slides = args.get("slides") or []
            if not slides:
                return None
            # A `create_presentation` called while a target deck is open means
            # the user wants the TARGET's content replaced with this full deck
            # (like create -> the whole file is "this"). Routing to add_slides
            # would DUPLICATE the existing deck instead of retheming it — the
            # historic triplication bug. Route to rebuild_slides.
            return (edit_name, {"path": self.target_file, "rebuild_slides": slides, "inplace": True})

        if edit_name == "edit_workbook":
            rows = args.get("data") or []
            if not rows:
                return None
            # Write the data into a fresh sheet starting at A1.
            writes, row_i = [], 1
            for row in rows:
                cells = row.get("cells", []) if isinstance(row, dict) else row
                for col_i, value in enumerate(cells, start=1):
                    cell = f"{chr(64 + col_i)}{row_i}"
                    writes.append({"cell": cell, "value": value})
                row_i += 1
            return (edit_name, {
                "path": self.target_file,
                "new_sheet": (args.get("sheet") or "Data"),
                "writes": writes,
                "inplace": True,
            })

        return None

    def _force_inplace_on_target(self, name: str, args: dict):
        """Ensure any edit of the selected target file happens IN PLACE.

        The model (following the system prompt) often calls edit_* directly on
        the target without passing inplace=True, which would otherwise create a
        duplicate 'NAME_edited.*' copy. This forces inplace whenever the edited
        path resolves to the open target file, so the selected file is updated
        directly and no duplicate is produced.
        """
        if not self.target_file:
            return name, args
        if name not in ("edit_document", "edit_workbook", "edit_presentation", "edit_text"):
            return name, args
        path = args.get("path")
        if not path:
            return name, args
        target = Path(self.target_file).resolve()
        edited_path = Path(str(path)).resolve()
        if edited_path != target:
            return name, args
        if not args.get("inplace"):
            args = {**args, "inplace": True}
        return name, args

    @staticmethod
    def _blocks_to_text(title: str, body) -> str:
        """Render create_document title+blocks into a plain-text append payload."""
        parts: list[str] = []
        if title:
            parts.append(title.upper())
            parts.append("")
        for block in body or []:
            if not isinstance(block, dict):
                continue
            btype = block.get("type", "paragraph")
            text = str(block.get("text", "")).strip()
            if not text:
                continue
            if btype == "heading":
                parts.append(f"## {text}")
            elif btype == "bullet":
                parts.append(f"- {text}")
            elif btype == "numbered":
                parts.append(f"1. {text}")
            else:
                parts.append(text)
        return "\n\n".join(parts).strip()

    # -- main loop ----------------------------------------------------------
    def run(self, prompt: str, sink=None):
        """Run the agent loop.

        sink: optional callable sink(kind, payload) where kind is one of
              'tool'  (payload = (name, args_dict))
              'text'  (payload = assistant text)
              'result'(payload = summary dict: {done, path/filename, ...})
              'error' (payload = error string)
              'status'(payload = status string).
        If no sink is given, output goes to self.console (CLI mode).
        """
        # Inject target file context
        user_msg = prompt
        if self.target_file:
            user_msg = f"[Target file: {self.target_file}]\n" + prompt

        def emit(kind, payload):
            if sink:
                sink(kind, payload)

        messages = [{"role": "system", "content": SYSTEM_PROMPT}]
        profile_block = profile_prompt_block()
        if profile_block:
            messages[0] = {
                "role": "system",
                "content": SYSTEM_PROMPT + "\n\n" + profile_block,
            }
        messages += self.history
        messages.append({"role": "user", "content": user_msg})

        tools_ran = False       # did ANY tool actually execute this run() call?
        noop_nudges = 0         # how many times we've nudged a text-only reply

        for _round in range(self.max_tool_rounds):
            emit("status", f"Contacting {self.cfg.get('provider')}…")
            try:
                msg = self.provider.chat(messages, tools=TOOLS)
            except ProviderError as exc:
                self.console.print(f"[red]Provider error:[/red] {exc}")
                emit("error", f"Provider error: {exc}")
                return

            if not msg.tool_calls:
                # Text-only reply. If the user asked for something actionable in a
                # file (create/edit/delete/append/summarize/etc.) and NO tool ran,
                # the model is often just acknowledging instead of acting — that is
                # the "it says done but did nothing" symptom. Give it one bounded
                # nudge to actually call the tool before we accept the answer.
                if (
                    not tools_ran
                    and noop_nudges == 0
                    and self._looks_actionable(user_msg)
                ):
                    noop_nudges += 1
                    self.console.print(
                        "[yellow]  ↳ replied with no tool call and no tool ran — "
                        "nudging it to actually perform the action[/yellow]")
                    messages.append({
                        "role": "assistant",
                        "content": msg.text or "",
                    })
                    messages.append({
                        "role": "user",
                        "content": (
                            "You replied without calling any tool. If this request "
                            "requires doing something to a file (create / edit / "
                            "append / delete / summarize), call the correct tool now "
                            "and actually perform it. Do not just say you will do it. "
                            "Only give a plain-text answer if this was a pure question."
                        ),
                    })
                    continue

                # Final text reply
                text = msg.text or "(no response)"
                self.history.append({"role": "user", "content": user_msg})
                self.history.append({"role": "assistant", "content": text})
                if msg.text:
                    self.console.print(msg.text)
                else:
                    self.console.print("[dim]Done. (no text returned)[/dim]")
                emit("text", text)
                emit("result", {"done": True})
                return

            # Execute tool calls
            messages.append({
                "role": "assistant",
                "content": msg.text or None,
                "tool_calls": [
                    {
                        "id": tc.id,
                        "type": "function",
                        "function": {"name": tc.name, "arguments": tc.raw_arguments or "{}"},
                    }
                    for tc in msg.tool_calls
                ],
            })

            last_mutated: str | None = None
            last_tool: str | None = None
            for tc in msg.tool_calls:
                exec_name, exec_args = tc.name, tc.arguments
                tools_ran = True
                remapped = self._remap_create_to_edit(tc.name, tc.arguments)
                note = ""
                if remapped:
                    exec_name, exec_args = remapped
                    note = (f" (routed to '{exec_name}' on the selected file "
                            f"'{self.target_file}')")
                    self.console.print(
                        f"[yellow]  ↳ target file set — using {exec_name} on "
                        f"{self.target_file} instead of creating a new file[/yellow]")
                exec_name, exec_args = self._force_inplace_on_target(exec_name, exec_args)
                self.console.print(f"[cyan]-> {exec_name}[/cyan] {self._pretty_args(exec_args)}")
                emit("tool", (exec_name, exec_args))
                emit("status", f"Running {exec_name}{note}…")
                result = self.executor.execute(exec_name, exec_args)
                result_json = json.dumps(result, default=str)
                if note and isinstance(result, dict):
                    result_json = json.dumps({
                        **result,
                        "_note": f"Routed to {exec_name} because a target file is open.",
                    }, default=str)
                created = result.get("path") if isinstance(result, dict) else None
                if created:
                    emit("result", {"done": False, "path": created, "tool": exec_name})
                    last_mutated = created
                    last_tool = exec_name
                messages.append({
                    "role": "tool",
                    "tool_call_id": tc.id,
                    "name": tc.name,
                    "content": result_json,
                })

            # -- POST-EDIT VERIFICATION -------------------------------
            # After an EDIT tool runs, re-read the ACTUAL file on disk and hand
            # the real resulting structure back to the model. This closes the
            # "says done but the change isn't really there" gap: the model sees
            # ground truth and can correct a mismatch before reporting done,
            # instead of trusting an `ok: true` from a tool that may have silently
            # failed (missing match -> no-op) or hit the wrong target.
            # Plain creates are NOT verified: a fresh file's content is exactly
            # what the tool wrote (ok:true is trustworthy), so verifying it only
            # adds a needless round-trip.
            if last_mutated is not None and last_tool and last_tool.startswith("edit_"):
                verif = self._verify_read(last_mutated)
                if verif:
                    verify_note = (
                        "[VERIFICATION] Below is the ground-truth state of "
                        f"{Path(last_mutated).name} on disk immediately after the "
                        "edit above. Confirm the requested change is actually present. "
                        "If it is NOT (or is wrong), run the correct edit tool to "
                        "really apply it before you report done — never claim a change "
                        "that is not visible here.\n\n"
                        + verif
                    )
                    messages.append({
                        "role": "user",
                        "content": verify_note,
                    })

        # Max rounds hit
        self.console.print("[yellow]Reached max tool-call rounds. Please simplify the request.[/yellow]")
        emit("error", "Reached max tool-call rounds. Please simplify the request.")

    def _verify_read(self, path: str) -> str | None:
        """Re-read the ACTUAL file on disk after an edit and return a compact
        structural snapshot the model can use to confirm the change landed.

        Returns a short string (or None if the file can't be read). Uses the
        on-disk read path (the same tools the model would call to inspect a
        document), not the live COM session, so verification is stable and
        reflects what a fresh open will see."""
        try:
            low = str(path).lower()
            if low.endswith((".docx", ".doc")):
                res = self.executor.execute("get_paragraphs", {"path": path, "limit": 60})
                paras = res.get("paragraphs") if isinstance(res, dict) else None
                if isinstance(paras, list) and paras:
                    head = "\n".join(str(p) for p in paras[:40])
                    return f"paragraphs ({len(paras)} shown):\n{head}"
            elif low.endswith(
                (".pptx", ".ppt", ".xlsx", ".xls", ".csv", ".pdf", ".txt", ".md")
            ):
                res = self.executor.execute("summarize_file", {"path": path})
                summary = res.get("summary") if isinstance(res, dict) else None
                if summary:
                    s = str(summary)
                    return s[:2500] + (" …[truncated]" if len(s) > 2500 else "")
            else:
                res = self.executor.execute("read_file", {"path": path})
                data = res.get("data") if isinstance(res, dict) else None
                if data is not None:
                    s = str(data)
                    return s[:2500] + (" …[truncated]" if len(s) > 2500 else "")
        except Exception:  # noqa: BLE001
            return None
        return None

    @staticmethod
    def _looks_actionable(prompt: str) -> bool:
        """Heuristic: does the user's request clearly want us to DO something to a
        file (rather than answer a pure question)? Used to nudge a text-only reply
        that would otherwise "say done" without performing any action."""
        p = (prompt or "").strip().lower()
        if not p:
            return False
        # "how do I …", "what is …", "why …", "can you tell me …" etc. are advisory
        # questions, not file actions — don't nudge those.
        question_prefixes = (
            "how do", "how can", "how to", "how would", "how should",
            "what is", "what are", "what does", "what's", "whats",
            "when ", "where ", "why ", "which ", "who ",
            "can you tell", "could you tell", "please explain", "explain ",
            "tell me how", "tell me what", "do you know",
        )
        for q in question_prefixes:
            if p.startswith(q):
                return False
        action_markers = (
            # create/write verbs
            "create ", "make ", "write ", "generate ", "build ", "draft ",
            "produce ", "add ", "insert ", "append ", "put ",
            # edit/change verbs
            "edit ", "update ", "change ", "modify ", "replace ", "fix ",
            "correct ", "rename ", "delete ", "remove ", "clear ",
            # doc/summary verbs
            "summarize ", "summarise ", "format ", "convert ", "export ",
            "read ", "open ", "fill ",
            # target/document hints
            "in the document", "in the file", "to this file", "to this document",
            "the selected", "into the", "on the doc", "section ", "page ",
        )
        for m in action_markers:
            if m in p:
                return True
        # A trailing "save it" / "run it" / "do it" also signals an action.
        return p.endswith(("save it", "save", "run it", "do it"))

    @staticmethod
    def _pretty_args(args: dict) -> str:
        if not args:
            return "{}"
        path = args.get("path")
        if path:
            return f"path={path}"
        keys = list(args.keys())
        return ", ".join(f"{k}=..." for k in keys[:3])
