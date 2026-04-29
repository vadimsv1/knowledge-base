# Knowledge Base Document Converter

Convert a folder of mixed documents (PDF, DOCX, XLSX, PPTX, images, Visio) into clean, structured Markdown — ready to feed an LLM as a personal or corporate knowledge base.

Built around a **two-tier hybrid pipeline**:

1. **Routing hybrid** — free local extraction (`pymupdf4llm`, `mammoth`, `openpyxl`) for text-only content; Claude Vision API only for pages that actually contain images, scans, or diagrams. Saves **80–95 %** on API cost vs. naive image-to-LLM pipelines.
2. **Vision hybrid** — when a page goes to Vision, **Google Cloud Vision OCR** extracts character-perfect text first, then **Claude Vision** is called with that OCR text injected into the prompt as ground truth. You get **exact character recognition** (Google's strength) **placed inside semantically structured Markdown** (Claude's strength). Most projects skip this layer and let Claude guess every character from pixels — this one doesn't.

On a typical text-heavy corpus you pay **~$0.02 per visual page** and **$0 for everything else**. Adding Google OCR adds ~9 % to per-page Vision cost in exchange for character-perfect numeric IDs, codes, and dense tables.

- Cross-platform (Windows / macOS / Linux / WSL)
- Crash-safe with MD5-tracked state — re-runs pick up where they stopped
- Per-file budget cap on API spend
- Purely additive Google OCR layer — toggled by a single env var, falls back gracefully on every error
- Optional one-click installer for fresh Windows machines (`deploy/`)

---

## Table of contents

- [Features](#features)
- [How it works](#how-it-works)
- [File-type support](#file-type-support)
- [Installation — pick your path](#installation--pick-your-path)
  - [Option 1 — Native install (any OS, recommended)](#option-1--native-install-any-os-recommended)
  - [Option 2 — Windows one-click installer (WSL + Ubuntu + Claude Code)](#option-2--windows-one-click-installer-wsl--ubuntu--claude-code)
  - [Which option should you pick?](#which-option-should-you-pick)
- [Usage](#usage)
- [Configuration](#configuration)
- [Project layout](#project-layout)
- [Output format](#output-format)
- [State management](#state-management)
- [Cost model](#cost-model)
- [Hybrid Vision: Google OCR + Claude Vision](#hybrid-vision-google-ocr--claude-vision)
- [Wiki compilation (Stage 3)](#wiki-compilation-stage-3)
- [Known limitations](#known-limitations)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

---

## Features

- **Hybrid Vision (Google OCR + Claude Vision)** — opt-in second-pass that grounds Claude's visual understanding with character-perfect OCR text from Google Cloud Vision. Eliminates the "Claude guesses long IDs from pixels" problem on technical/numeric content. See [Hybrid Vision](#hybrid-vision-google-ocr--claude-vision) below.
- **Per-page routing for PDFs** — text-only pages use free local extraction; pages with images, scans, or vector diagrams go through Vision API. Saves 80–95 % of API spend vs. sending every page to Vision.
- **Context bridging** — the last 200 chars of the previous page are passed to the next Vision call so tables and paragraphs don't break at page boundaries
- **DOCX inline images** — image descriptions are inserted at their **exact position** in the text via mammoth's `convert_image` callback, not appended at the end
- **XLSX preserved verbatim** — every sheet, every cell, as a Markdown table; no API calls
- **Resume-safe state** — MD5-tracked, crash-safe, deduplicates identical files, re-processes modified ones
- **Hard budget cap** — set a USD ceiling per run; the pipeline stops cleanly when reached
- **Language preservation** — Russian stays Russian, English stays English, etc. The prompts explicitly forbid translation
- **Smart prompts per content type** — document pages get a verbatim transcription prompt; standalone images get a knowledge-extraction prompt; DOCX images get a context-aware prompt with surrounding text

---

## How it works

Two tiers of hybrid logic.

### Tier 1 — Routing hybrid (always on)

Every page is analyzed locally first to decide which engine processes it.

```
rawdocs/  →  sort by type  →  per-page analysis  →  md_ready/
                                     │
                           ┌─────────┴─────────┐
                           │                   │
                     Text pages            Visual pages
                     (pymupdf4llm /        (sent to Vision,
                      mammoth /             see Tier 2 below)
                      openpyxl)
                     FREE, instant
                           │                   │
                           └─────────┬─────────┘
                                     │
                              Assembly with
                              context bridging
                                     │
                                Final .md
```

A PDF page is routed to Vision only if it has significant images (≥150 px), vector drawings (≥20 paths), or very little text (<80 chars). Everything else takes the free path. This typically saves **80–95 %** on API cost vs. sending every page to Vision.

### Tier 2 — Vision hybrid (opt-in via `GOOGLE_API_KEY`)

When a page does need Vision, the call itself is hybrid: Google OCR extracts exact text, Claude Vision interprets structure.

```
                       Image (rendered @ 300 DPI, ≤3000 px, base64 PNG)
                                          │
                                          ▼
                          ┌──── GOOGLE_API_KEY set? ────┐
                          │                             │
                         No                            Yes
                          │                             │
                          │                             ▼
                          │              ┌─────────────────────────────┐
                          │              │ Google Cloud Vision         │
                          │              │ DOCUMENT_TEXT_DETECTION     │
                          │              │ → exact characters          │
                          │              │   (numbers, IDs, codes,     │
                          │              │    dense tables)            │
                          │              └──────────────┬──────────────┘
                          │                             │
                          │                             ▼
                          │              ┌─────────────────────────────┐
                          │              │ Inject OCR text into the    │
                          │              │ Claude prompt as ground     │
                          │              │ truth ("use these verbatim")│
                          │              └──────────────┬──────────────┘
                          │                             │
                          └────────────┬────────────────┘
                                       ▼
                       ┌────────────────────────────────────┐
                       │ Claude Vision (Sonnet)             │
                       │ → structured Markdown              │
                       │   - tables, headings, layout       │
                       │   - diagrams + flowcharts          │
                       │   - language preserved             │
                       └────────────────────────────────────┘
                                       │
                                       ▼
                                Markdown for this page
```

**Why this matters:** Claude Vision is excellent at structure, layout, and meaning, but can drift on long alphanumeric strings (transaction IDs, hex codes, parameter lists, fine print). Google OCR returns those character-perfect but has no idea what's a heading vs. a caption. Combined, you get exact text placed correctly inside semantically structured Markdown.

Most public projects that use Claude Vision skip this layer — they send pixels to Claude and hope. Adding a deterministic OCR pass with prompt-level grounding fixes that, for ~9 % extra cost per Vision call.

The Google OCR layer is **purely additive**: if `GOOGLE_API_KEY` is unset, or if the OCR call fails for any reason, Claude is called alone with the original prompt. You never get worse output from enabling it.

---

## File-type support

| Type | Method | Cost | Notes |
|---|---|---|---|
| **PDF** | Hybrid per-page | Free for text pages, ~$0.02 / visual page | Pages with images >150 px, vector drawings, or no text → Vision API |
| **DOCX** | mammoth + Vision (for inline images) | Free if no images | Image descriptions inserted at exact position |
| **XLSX** | openpyxl | **Free** | All sheets → Markdown tables, pipe-escaped |
| **PPTX** | Vision API | ~$0.02 / slide | Slides are inherently visual; falls back to text-only via python-pptx if rendering fails |
| **Images** (JPG / PNG / GIF / WEBP / BMP / TIFF / SVG) | Vision API | ~$0.02 | Knowledge-extraction prompt — extracts meaning, not just description |
| **Visio** (.vsdx, .vsd) | Render via pymupdf → Vision | ~$0.02 / page | If rendering fails, exports an error and asks you to convert to PDF first |
| **TXT / CSV / JSON / XML / HTML / YAML / INI / CFG / LOG** | Direct read | Free | Encoding auto-detected (utf-8, utf-8-sig, cp1251, latin-1) |

Old binary formats (`.doc`, `.xls`, `.ppt`) are routed through the same converters as their modern counterparts but coverage depends on the underlying library — convert them to the modern format first if you hit issues.

---

## Installation — pick your path

You have two ways to get this running. **Most users want Option 1.** Option 2 is a convenience installer for fresh Windows machines that also want Claude Code CLI bundled in.

> **Note:** WSL is **not required** to run this project. The converter is pure Python and runs natively on Windows, macOS, Linux, or WSL — anywhere Python runs. WSL only appears in Option 2 because that bundle also installs the Claude Code CLI in a clean Linux environment for convenience.

---

### Option 1 — Native install (any OS, recommended)

Works on **Windows / macOS / Linux / WSL**, with no privileged operations and no extra moving parts. The only requirement is Python 3.10+ and a Claude API key.

**1. Clone the repo**

```bash
git clone https://github.com/vadimsv1/knowledge-base.git
cd knowledge-base
```

**2. Install Python dependencies**

```bash
pip install pymupdf pymupdf4llm mammoth openpyxl python-pptx Pillow anthropic
```

> On externally-managed Python (some Linux distros), add `--break-system-packages` or use a virtualenv:
> ```bash
> python -m venv .venv && source .venv/bin/activate   # macOS / Linux / WSL
> python -m venv .venv && .venv\Scripts\activate      # Windows PowerShell
> pip install pymupdf pymupdf4llm mammoth openpyxl python-pptx Pillow anthropic
> ```

**3. Set your API key**

```bash
export ANTHROPIC_API_KEY="sk-ant-..."          # macOS / Linux / WSL
$env:ANTHROPIC_API_KEY = "sk-ant-..."          # Windows PowerShell
set ANTHROPIC_API_KEY=sk-ant-...               # Windows cmd.exe
```

**4. Drop files into `rawdocs/` and run**

```bash
python scripts/convert.py
```

Output appears in `md_ready/`, one Markdown file per source document. The project root is auto-detected from the script location — clone the repo anywhere and it just works.

**That's it.** No WSL, no admin rights, no system tweaks.

---

### Option 2 — Windows one-click installer (WSL + Ubuntu + Claude Code)

For a Windows 10/11 machine with **no Python, no WSL, no Node** that you want to set up end-to-end in one go, including the Claude Code CLI for interactive use. The `deploy/` folder bootstraps the entire stack inside WSL2 + Ubuntu 24.04.

This option is useful if you:
- Want everything (Python deps + Claude Code CLI) installed in one shot
- Are setting up a brand-new Windows machine
- Prefer a Linux-style dev environment for working with the project

It is **not necessary** to use this option just to run the converter — Option 1 works on Windows natively.

**Step 1 — Enable WSL2 + install Ubuntu** *(Admin PowerShell)*

```powershell
cd C:\dev\knowledge-base\deploy
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\01-setup-wsl.ps1
```

Reboot when prompted.

**Step 2 — Install project + Python deps + Claude Code** *(Admin PowerShell, after reboot)*

```powershell
cd C:\dev\knowledge-base\deploy
.\02-setup-project.ps1
```

This script will:
1. Create `C:\dev\knowledge-base\` and all subdirectories
2. Copy `scripts/`, `CLAUDE.md`, `deploy/` into place
3. Run `03-setup-ubuntu.sh` inside WSL to install Python deps + Claude Code CLI
4. Prompt you for your `ANTHROPIC_API_KEY` and save it to `~/.bashrc`

**Step 3 — Daily use**

```powershell
.\04-launch.ps1          # Launch Claude Code in the project
.\quick-convert.ps1      # Run the converter headlessly
```

Full deployment details and troubleshooting in [`deploy/README.md`](deploy/README.md).

> The deploy scripts assume `C:\dev\knowledge-base\` as the install path. If you want it elsewhere, use Option 1 instead.

---

### Which option should you pick?

| You are… | Pick |
|---|---|
| On macOS or Linux | **Option 1** |
| On Windows and already have Python 3.10+ | **Option 1** |
| On Windows, want a quick run with no system changes | **Option 1** (one `pip install`, done) |
| On a fresh Windows machine with no Python / no Node, and want everything (including Claude Code CLI) bundled | **Option 2** |
| Inside WSL already and comfortable there | **Option 1** (run the bash commands inside WSL) |

If unsure, **start with Option 1**. You can always layer Option 2's Claude Code CLI on top later if you decide you want it.

---

## Usage

```bash
# Drop files into rawdocs/ first, then:

python scripts/convert.py                   # full pipeline
python scripts/convert.py --status          # show current state (no API calls)
python scripts/convert.py --reset           # wipe state + md_ready/, start fresh
python scripts/convert.py --retry-errors    # re-process only files that failed
python scripts/convert.py --max-budget 5.0  # hard cap on API spend (USD, default 10)
```

The `scripts/run.sh` wrapper is a thin shortcut: `./scripts/run.sh status`, `./scripts/run.sh reset`, `./scripts/run.sh retry`.

---

## Configuration

All configuration is via environment variables (or an optional `config/api_keys.json`).

| Variable | Purpose | Default |
|---|---|---|
| `ANTHROPIC_API_KEY` | **Required** for any Vision API call | — |
| `GOOGLE_API_KEY` | **Recommended** — enables [Hybrid Vision](#hybrid-vision-google-ocr--claude-vision) (Google OCR + Claude). Pure upside, ~9 % extra cost per Vision call. | — (off) |
| `KB_PROJECT_DIR` | Override project root | parent of `scripts/` directory |
| `KB_MODEL` | Claude model to use | `claude-sonnet-4-6` |
| `KB_MAX_TOKENS` | Max output tokens per Vision call | `8192` |

Alternatively, place a `config/api_keys.json` like:

```json
{
  "ANTHROPIC_API_KEY": "sk-ant-...",
  "GOOGLE_API_KEY": "AIza..."
}
```

Values from this file take precedence over environment variables. The file is **gitignored** by default.

---

## Project layout

```
knowledge-base/
├── rawdocs/          # DROP ZONE — put your source files here
├── sorted/           # auto-sorted copies, by file type           [gitignored]
├── md_ready/         # OUTPUT — final Markdown files              [gitignored]
├── wiki/             # compiled knowledge base (Stage 3+)         [gitignored]
├── scripts/
│   ├── convert.py    # main converter
│   └── run.sh        # shell wrapper
├── deploy/           # one-click Windows installer (see deploy/README.md)
├── config/           # state JSON + optional api_keys.json        [gitignored]
├── logs/             # per-run logs                               [gitignored]
├── CLAUDE.md         # full project spec for Claude Code
├── LICENSE           # MIT
└── README.md         # this file
```

User content (`rawdocs/`, `sorted/`, `md_ready/`, `wiki/`, `logs/`, `config/`) is gitignored — only code, scripts, and docs are tracked.

---

## Output format

Each `md_ready/<file>.md` starts with YAML frontmatter so downstream tools can filter by type, hash, or method:

```yaml
---
source_file: "annual-report-2024.pdf"
source_type: ".pdf"
category: "pdf"
source_size_bytes: 2451823
source_md5: "a1b2c3d4..."
pages: 42
pages_text: 38
pages_vision: 4
method: "hybrid"
api_calls: 4
est_cost_usd: 0.0784
converted_at: "2026-04-28T10:23:11+00:00"
---

# annual-report-2024

<!-- page 1 -->
...
```

`<!-- page N -->` markers separate logical pages so you can split or chunk later.

---

## State management

`config/convert_state.json` tracks every file by MD5 hash:

- **Resume-safe** — state is saved after each file. If the process crashes, the next run picks up where it stopped.
- **Duplicate detection** — identical files (same hash) are skipped, even under different filenames.
- **Modified-file detection** — if a file's content changes, it's re-processed automatically.
- **Cost tracking** — per-file and total API spend in USD.

To wipe state and reprocess everything: `python scripts/convert.py --reset`.

---

## Cost model

Vision API cost per page is roughly:

```
~1,600 input tokens × $3/MTok  +  ~800 output tokens × $15/MTok  ≈  $0.0168 / page
```

Real-world examples:
- A 50-page PDF where 45 pages are pure text and 5 have diagrams → **~$0.08**
- A 30-slide PPTX → **~$0.50**
- A 10-image folder → **~$0.17**
- A 20-sheet XLSX → **$0.00**

Use `--max-budget` to enforce a hard ceiling. The pipeline stops cleanly mid-run when the cap is hit; re-running picks up from the next file.

---

## Hybrid Vision: Google OCR + Claude Vision

This is the project's most distinctive design choice. **Strongly recommended** for any corpus that contains technical documents, receipts, financial records, code dumps, payment messages, or anything with critical alphanumeric strings.

### What it is

When `GOOGLE_API_KEY` is set, every Vision call goes through both engines instead of just Claude:

1. The image is sent to **Google Cloud Vision** (`DOCUMENT_TEXT_DETECTION`). Google returns character-perfect text — every digit, every code, every label — with no semantic interpretation.
2. That OCR text is **injected at the top of the Claude prompt** with explicit instructions: *"Use these exact text values for all numbers, codes, identifiers, and technical strings. Do NOT guess or approximate — use the OCR values verbatim."*
3. **Claude Vision** is then called with both the image **and** the OCR text as ground-truth context. Its job becomes: structure this content into Markdown using the OCR text as authoritative for characters, and your visual understanding for layout, tables, diagrams, and meaning.

### Why it works

| Engine | Strength | Weakness |
|---|---|---|
| **Claude Vision** | Layout, structure, tables, diagrams, semantic understanding | Can drift on long alphanumeric strings, dense codes, tiny print |
| **Google OCR** | Pixel-perfect character recognition in any layout | Zero understanding of structure, headings, tables, or meaning |
| **Combined (this pipeline)** | Exact characters placed inside semantically correct Markdown | None significant — see failure modes below |

You're paying Anthropic to *understand* the page and Google to *transcribe* it. Each does what it's best at; neither has to compensate for the other's weakness.

### Cost economics

| Engine | Per-image cost |
|---|---|
| Claude Vision (Sonnet) | ~$0.0168 |
| Google Vision DOCUMENT_TEXT_DETECTION | $0.0015 (free for first 1,000/month) |
| **Hybrid total** | **~$0.0183** (≈ +9 % over Claude alone) |

For a 50-page PDF with 5 visual pages, hybrid mode adds roughly **$0.0075** to the run. Almost always worth it on technical content; barely noticeable on prose-heavy material.

### Where the difference shows up

| Content type | Claude alone | + Google OCR |
|---|---|---|
| Long numeric IDs (`3.000..4B.280HSMPKHSMPKSig`) | May miss a digit, transpose chars | Character-perfect |
| Technical parameter lists (`termid,rspcode,SBPTMKPublic,...`) | May drop or rephrase | Exact, verbatim |
| Receipts, invoices, financial tables | Occasionally rounds or drops | Faithful to source |
| Small print, footnotes | Often misreads | Reads cleanly |
| Plain prose paragraphs | Already excellent | Same — no real difference |
| Diagrams and flowcharts | Already excellent | Same for structure; better for label text |

### Failure modes (defensive by design)

The Google OCR layer is **purely additive** — every failure path falls back to plain Claude Vision automatically:

- **`GOOGLE_API_KEY` unset** → OCR step is skipped, Claude is called alone. No error.
- **Google API rejects the request, returns an error, or times out** → logged as a warning, OCR returns empty string, Claude is called with the original prompt. No error surfaced to the user.
- **Google succeeds but returns no text** (blank page, photo of nothing) → no injection happens, Claude is called normally.
- **Claude Vision fails after retries** → the file is marked as `error` in state and can be retried later with `--retry-errors`. Same behavior with or without the OCR layer.

Bottom line: enabling `GOOGLE_API_KEY` cannot make output *worse* than Claude alone. It can only match or improve it.

### How to enable

```bash
# 1. Enable the Cloud Vision API in Google Cloud Console
#    https://console.cloud.google.com/apis/library/vision.googleapis.com
# 2. Create an API key, restrict it to "Cloud Vision API" only
# 3. Set the env var:

export GOOGLE_API_KEY="AIza..."        # macOS / Linux / WSL
$env:GOOGLE_API_KEY = "AIza..."        # Windows PowerShell
```

That's it. The pipeline auto-detects the variable and starts injecting OCR context on every Vision call. Look for `+ Google OCR context added` in the per-page logs to confirm it's firing.

### When to skip it

Skip the Google OCR layer if your corpus is:
- Mostly prose (articles, essays, narrative documents)
- Already-OCRed text PDFs (most modern PDFs — these never reach Vision anyway, see Tier 1 routing)
- Free-tier sensitive (you don't want to set up a Google Cloud account just to convert a few documents)

For everyone else: turn it on.

---

## Wiki compilation (Stage 3)

Once `md_ready/` is populated, you can ask Claude Code (or any LLM agent) to compile the contents into a structured wiki — Karpathy's "LLM Wiki" pattern — with:
- An index of all topics
- Cross-references and backlinks between related documents
- Per-topic summaries
- Deduplication of overlapping information

The wiki output goes into `wiki/` (gitignored by default). See [`CLAUDE.md`](CLAUDE.md) for the full pipeline spec.

---

## Known limitations

- **Visio (`.vsdx`)** — pymupdf rendering is partial. Best results: export to PDF from Visio first.
- **Old binary formats** (`.doc`, `.xls`, `.ppt`) — not natively supported. Convert with LibreOffice (`soffice --headless --convert-to docx file.doc`) or save as the modern format first.
- **XLSX charts** — openpyxl cannot extract chart images. Charts in Excel files are silently dropped.
- **Synchronous API calls** — pages are processed one at a time. Large presentations (~50+ slides) take several minutes.
- **DOCX page count** — DOCX has no native page concept; the value in frontmatter is estimated from paragraph count and is approximate.

---

## Troubleshooting

| Symptom | Likely cause / fix |
|---|---|
| `Missing: pymupdf, ...` | Run the `pip install` line from [Quick start](#quick-start). On externally-managed Python (Debian/Ubuntu), add `--break-system-packages` or use a venv. |
| `ANTHROPIC_API_KEY not set` | `export ANTHROPIC_API_KEY="sk-ant-..."` before running, or place it in `config/api_keys.json`. |
| `Need ~$X.XX, budget $Y.YY` | Raise the cap with `--max-budget` or split your `rawdocs/` into smaller batches. |
| File appears with `status: error` | `python scripts/convert.py --retry-errors` to retry only failed files. Inspect `logs/convert_*.log` for the traceback. |
| Vision output looks wrong / hallucinated | Open `md_ready/<file>.md`, find the page, and inspect the source. Consider enabling Google Vision OCR (see above) for higher accuracy on numeric/code-heavy pages. |
| Russian (or other non-Latin) text comes back translated | The prompts explicitly forbid translation; if it still happens, file an issue with the source page and the converted output. |
| WSL deployment errors | See [`deploy/README.md`](deploy/README.md) for WSL-specific troubleshooting. |

---

## Contributing

Issues and pull requests welcome. Please:
1. Open an issue first for non-trivial changes so we can align on approach.
2. Keep changes focused — one PR per concern.
3. All code, comments, and log messages in English. Prompts in English (with explicit instructions to preserve source language).

---

## License

[MIT](LICENSE) — free for personal and commercial use, no warranty.
