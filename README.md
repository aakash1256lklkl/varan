<div align="center">

# 🦎 Varan — AI Office Agent

**Provider-agnostic AI assistant for Word, Excel, PowerPoint, PDF & text.**
Create, edit, read and summarize Microsoft Office files — from natural language, right in your terminal or beside your open Office app.

![Python](https://img.shields.io/badge/Python-3.9%2B-3776AB?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-blue)
![Status](https://img.shields.io/badge/build-40%2F40%20passing-brightgreen)

---

```text
██████╗  █████╗ ██████╗  █████╗ ███╗   ██╗
██╔══██╗██╔══██╗██╔══██╗██╔══██╗████╗  ██║
██████╔╝███████║██████╔╝███████║██╔██╗ ██║
██╔══██╗██╔══██╗██╔═══╝ ██╔══██║██║╚██╗██║
██║  ██║██║  ██║██║     ██║  ██║██║ ╚████║
╚═╝  ╚═╝╚═╝  ╚═╝╚═╝     ╚═╝  ╚═╝╚═╝  ╚═══╝
         AI OFFICE AGENT
```

Talk to Varan in **plain English**:

```text
> create a 10-slide pitch deck about a fintech app
> make an Excel budget tracker with a monthly bar chart
> summarize proposal.docx and turn it into a 5-slide deck
> add a Q1 summary sheet to my workbook
```

Varan decides which tool to call, builds **real files**, and saves them to an `outputs/`
folder — your originals are never overwritten.

---

## ✨ Highlights

| | |
|---|---|
| 🧠 **Provider-agnostic** | One codebase, ten providers: OpenAI, Anthropic, Gemini, OpenRouter, Groq, Mistral, Ollama, LM Studio & any OpenAI-compatible server. |
| 📄 **Five formats** | Word `.docx`, Excel `.xlsx`, PowerPoint `.pptx`, PDF `.pdf`, plain text `.txt`/`.md` — create, edit, read, summarize. |
| 🔴 **Live editing** | If a file is *open* in Word, Excel or PowerPoint, Varan drives the app via COM automation — changes appear **live on screen**, no lock errors, no `_edited` duplicates. |
| 🪟 **Floating Companion** | An always-on-top panel that sits beside your Office app. Global hotkey, drag-droppable, provider/model switches. |
| ✂️ **Precise, formatting-preserving edits** | Powered by DocxEngine — replace/insert/delete **without** wiping bold, italic, fonts, styles, tracked changes or comments. |
| 🧼 **Editing safety** | Destructive deletes need confirmation. Post-edit `[VERIFICATION]` re-reads the on-disk file so a silent failure is never reported as success. |

---

## 🚀 Quick start

### 1. Install

```bash
cd varan
pip install -r requirements.txt
copy .env.example .env
```

### 2. Configure

Edit `.env` to pick your provider and model:

```ini
AI_PROVIDER=openrouter
AI_MODEL=openai/gpt-4o-mini
AI_API_KEY=sk-or-...
```

> **No API key?** Use [Ollama](https://ollama.com) — free and 100% local:
>
> ```bash
> ollama pull llama3.2
> ollama serve        # keep this running in another terminal
> ```
>
> Set `AI_PROVIDER=ollama` and `AI_MODEL=llama3.2` — no key needed.

### 3. Run

```bash
python main.py
```

You get a `varan>` prompt. Type natural-language requests, or use **slash commands**:

| Command | What it does |
|---|---|
| `/provider <name>` | Switch provider live |
| `/model <name>` | Switch model live |
| `/config` | Show current provider/model |
| `/open <file>` | Set a target file for edits |
| `/outputs` | List files in the `outputs/` folder |
| `/help` | Show help |
| `/exit` / `quit` | Leave Varan |

---

## 🪟 Varan Companion — work *inside* Office

Prefer your Office app over a terminal? The Companion is an **always-on-top floating panel**
that sits right beside Word, Excel or PowerPoint.

```bash
python companion_run.py        # or double-click VaranCompanion.bat
```

- **Global hotkey** `Ctrl+Alt+Shift+V` toggles the panel while you work.
- **Open Office file** picks the `.docx` / `.xlsx` / `.pptx` you have open and *launches it in
  its native app*, so the document sits next to the Companion — then your edits appear
  **live** in real time.
- Type a prompt (or press **Enter**) → Varan streams progress into the panel and reports
  where files were saved.
- **Provider…** / **Model…** buttons switch providers/models live.
- **Profile** button → tell Varan how you speak and the tasks you repeat.
- **Outputs** opens the generated-files folder.

> Because Office 2021 doesn't support the Microsoft-365-only "task pane" web add-in, the
> Companion floats alongside the app instead. Open your file, hit the hotkey, prompt away.

---

## 🧠 Personal profile & task templates

Varan works best when it knows *you*. The **Profile** button opens a form where you set:

- **How Varan speaks to you** — tone (professional / friendly / casual), verbosity,
  formality, detail level, preferred structure.
- **Who you are** — name, role, and notes it should always remember
  (e.g. *"I write screenplays and series bibles"*).
- **Task templates** (recipes) — named tasks with a **goal** that defines exactly what
  "done" means. Ask for a recipe by name and Varan follows it precisely.

The profile is saved to `profile.json` and **injected into the system prompt on every
request** — so every reply is tailored to you automatically.

---

## 🔌 Supported providers

One unified protocol (OpenAI-compatible Chat Completions) plus native adapters:

| Provider | `AI_PROVIDER` | Key? |
|---|---|---|
| OpenAI | `openai` | `AI_API_KEY` |
| OpenRouter (one key → many models) | `openrouter` | `AI_API_KEY` |
| Groq (fast free tier) | `groq` | `AI_API_KEY` |
| Mistral | `mistral` | `AI_API_KEY` |
| Anthropic Claude | `anthropic` | `AI_API_KEY` |
| Google Gemini | `gemini` | `AI_API_KEY` |
| **Ollama** (local, free) | `ollama` | *none* |
| **LM Studio** (local) | `lmstudio` | *none* |
| Any OpenAI-compatible server | `custom` | `AI_BASE_URL` + `AI_API_KEY` |

---

## 📦 What Varan can do

### Word — `create_document` / `edit_document`
- Title, headings (1–6), paragraphs, bold/italic, bullets, numbered lists, tables.
- **Formatting-preserving edits** via [DocxEngine](https://github.com/ruwadgroup/docxengine):
  replace / insert / delete without wiping styles, fonts, tracked changes or comments.
- **Tracked changes (redlines)** recorded under a named author.
- **Replace all** (`count: -1`), and **complex table edits** (`table_edits` sets one
  cell's text inside an existing table, e.g. `{"table": 1, "cell": {"ref": "B2"}, "text": "…"}`).
- **Live editing**: open the `.docx` in Word and Varan edits it via COM — changes stream
  into the open window, no file-lock error, no duplicate.

### Excel — `create_workbook` / `edit_workbook`
- Sheets, cell values, formulas (`=SUM(...)`), bar/line/pie charts, new sheets.
- **Complex edits**: `rows`/`columns` insert/delete, `clear` on cells or ranges,
  cell `styles` (bold / italic / size / font / fill), `delete_sheet`.
- **Live editing**: open workbook → COM drives the live worksheet.

### PowerPoint — `create_presentation` / `edit_presentation`
Three editing modes, chosen automatically:
- **`add_slides`** — stack new slides onto the existing deck.
- **`rebuild_slides`** — **replace** the entire deck ("retheme" / "rewrite every slide").
- **`edit_slides`** — **surgical**: change *only* specific slides by index — rename a
  title, fix text, swap bullets, or `add_textbox` / `add_table` / `add_chart` **onto** one
  slide. Fast and leaves every other slide untouched.
- **`remove_slides`** — surgically delete specific slide indices.
- **Placeholder garbage is refused**: content like `"- item"` or `"i | t | e | m"` is
  rejected with a clear error (fail-loud on file *and* live paths) — a deck can never
  "save ok" while holding placeholder text.
- **Create on an open target rethemes, never duplicates** — the historic
  "9 slides become 27" triplication bug is killed at the root.
- **Live editing**: open presentation → COM adds slides to the running show.

### PDF & text — `edit_text`
- Read, summarize and text-edit **PDFs** (pypdf) and **.txt / .md / .rst / .log** files.
- `replace` / `delete` / `insert_after` / `insert_before` / `delete_range`.
- PDF editing rewrites content streams — safe for simple / generated PDFs.

### Reading & summarizing
- `read_file` and `summarize_file` work across all five formats.

---

## 🛡️ Editing safety

- **Live editors attach, never relaunch**: COM `GetActiveObject` — Varan *never* spawns a
  hidden second Office instance or kills processes, so your real session and unsaved work
  are never touched.
- **Destructive actions need confirmation**: deletes are refused out of the blue — Varan
  lists what will be removed and only proceeds after you confirm (`confirm: true`).
- **In-place editing of your target**: when a file is your selected target it's edited in
  place — one continuously-updated document instead of piles of `_edited` copies.
- **`_edited` copies otherwise**: editing an existing *non-target* file writes a safe copy
  (e.g. `proposal_edited.docx`).
- **Anti "says-done-but-does-nothing"**: if no tool is called on an actionable request,
  Varan nudges itself to actually perform the action. After editing, the loop re-reads the
  on-disk file and verifies the change landed before reporting done.
- **Strict mode**: set `VARAN_STRICT=1` to surface unexpected live-edit errors verbatim
  instead of masking them.

> **Dependency note:** live Office editing needs the app installed (Word/Excel/PowerPoint)
> plus `pywin32`. If the app isn't running, Varan automatically falls back to file-based
> editing.

---

## 🗂️ Project layout

```
varan/
├── main.py               # CLI chat loop (rich)
├── companion_run.py      # launch the floating Companion panel
├── VaranCompanion.bat    # double-click launcher for the Companion
├── requirements.txt
├── .env.example          # provider + model + key config (template only)
├── agent/
│   ├── config.py         # .env → provider settings
│   ├── providers.py      # unified provider layer
│   ├── tools.py          # JSON-schema tool defs + executor
│   ├── loop.py           # prompt → tool-call → execute → reply loop
│   └── compat.py         # Windows UTF-8 console helpers
├── companion/
│   └── window.py         # tkinter floating panel UI
├── office/
│   ├── registry.py       # file extension → editor routing
│   ├── word_editor.py    # python-docx + DocxEngine
│   ├── excel_editor.py   # openpyxl
│   ├── ppt_editor.py     # python-pptx
│   ├── text_editor.py    # pypdf + plain-text editing
│   ├── word_live.py      # Word COM live editor
│   ├── excel_live.py     # Excel COM live editor
│   └── ppt_live.py       # PowerPoint COM live editor
├── tests/
│   ├── smoke_test.py      # happy path for every editor (no AI key needed)
│   └── editing_stress.py  # EVERY edit Varan must perform, per format
└── outputs/              # generated files live here (gitignored)
```

---

## 🧪 Testing (no AI key required)

```bash
python tests\smoke_test.py          # each editor creates + reads a valid file
python tests\editing_stress.py      # the full edit contract, via ToolExecutor
```

`editing_stress.py` is the checklist of edits Varan must be able to perform — **40 specs**:

- **Word (18)** — styled create, exact-phrase & replace-all (`count: -1`), insert before/
  after, append without markdown leaks, destructive-delete confirmation gate,
  `delete_range` (+ `end_level`), `remove_blank_pages`, tracked-changes (`w:del`/`w:ins`),
  in-place no-duplicate, table-cell edits…
- **Excel (4)** — create, cell writes + `=SUM()` formulas, new sheets + charts, complex
  row/column/style/sheet ops.
- **PowerPoint (8)** — create layouts, append, **rebuild** (retheme), surgical `edit_slides`,
  shape adds onto a slide, `remove_slides`, filler-content **refusal**, create-on-target
  rebuild remap.
- **PDF (2)** — read/summarize, best-effort replace.
- **Text (7)** — replace first/Nth/all, insert, append, delete, `delete_range`.
- **Cross-cutting (1)** — create-on-open-target remap + `[VERIFICATION]` on-disk re-read.

One guarantee the suite enforces: **a missing match is never a silent success.**

---

## ⚙️ How it works

1. The chat loop sends your request plus tool schemas to the provider.
2. The model returns a **tool call** (e.g. `create_document`).
3. Varan validates and **executes** it on a real file via the editor layer.
4. The tool result feeds back and Varan replies with the saved path.
5. Multi-step requests ("create a doc, *then* a deck from it") run across several tool
   rounds within one conversation.

---

## 📜 License

[MIT](./LICENSE) — do what you like; credit appreciated.

<div align="center">

**Built with 💚 for people who work in Office apps.**

</div>