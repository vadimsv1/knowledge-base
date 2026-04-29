# Knowledge Base Document Converter

Convert a folder of mixed documents (PDF, DOCX, XLSX, PPTX, images, Visio) into clean, structured Markdown — ready to feed an LLM as a personal or corporate knowledge base.

Built around a **hybrid pipeline**: free local extraction (`pymupdf4llm`, `mammoth`, `openpyxl`) for text-only content, and the Claude Vision API only for pages that actually contain images, scans, or diagrams. On a typical text-heavy corpus you pay **~$0.02 per visual page** and **$0 for everything else**.

- Cross-platform (Windows / macOS / Linux / WSL)
- Crash-safe with MD5-tracked state — re-runs pick up where they stopped
- Per-file budget cap on API spend
- Optional one-click installer for fresh Windows machines (`deploy/`)

---

## Table of contents

- [Features](#features)
- [How it works](#how-it-works)
- [File-type support](#file-type-support)
- [Quick start](#quick-start)
- [One-click setup on a fresh Windows machine](#one-click-setup-on-a-fresh-windows-machine)
- [Usage](#usage)
- [Configuration](#configuration)
- [Project layout](#project-layout)
- [Output format](#output-format)
- [State management](#state-management)
- [Cost model](#cost-model)
- [Optional: Google Vision OCR for higher text accuracy](#optional-google-vision-ocr-for-higher-text-accuracy)
- [Wiki compilation (Stage 3)](#wiki-compilation-stage-3)
- [Known limitations](#known-limitations)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

---

## Features

- **Per-page routing for PDFs** — text-only pages use free local extraction; pages with images, scans, or vector diagrams go through Vision API
- **Context bridging** — the last 200 chars of the previous page are passed to the next Vision call so tables and paragraphs don't break at page boundaries
- **DOCX inline images** — image descriptions are inserted at their **exact position** in the text via mammoth's `convert_image` callback, not appended at the end
- **XLSX preserved verbatim** — every sheet, every cell, as a Markdown table; no API calls
- **Resume-safe state** — MD5-tracked, crash-safe, deduplicates identical files, re-processes modified ones
- **Hard budget cap** — set a USD ceiling per run; the pipeline stops cleanly when reached
- **Language preservation** — Russian stays Russian, English stays English, etc. The prompts explicitly forbid translation
- **Smart prompts per content type** — document pages get a verbatim transcription prompt; standalone images get a knowledge-extraction prompt; DOCX images get a context-aware prompt with surrounding text

---

## How it works

```
rawdocs/  →  sort by type  →  hybrid conversion  →  md_ready/
                                     │
                           ┌─────────┴─────────┐
                           │                   │
                     Text pages            Visual pages
                     (pymupdf4llm /        (Claude Vision API,
                      mammoth /             ~$0.02 / page)
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

For PDFs, every page is analyzed locally first: pages that are pure text take the free path; pages with significant images (≥150 px) or vector drawings (≥20 paths) take the Vision path. This routing typically saves **80–95 %** on API costs versus sending every page through Vision.

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

## Quick start

Works on any OS with Python 3.10+ and a Claude API key.

```bash
# 1. Clone
git clone https://github.com/<your-username>/knowledge-base.git
cd knowledge-base

# 2. Install dependencies
pip install pymupdf pymupdf4llm mammoth openpyxl python-pptx Pillow anthropic

# 3. Set your API key
export ANTHROPIC_API_KEY="sk-ant-..."          # macOS / Linux / WSL
# $env:ANTHROPIC_API_KEY = "sk-ant-..."        # Windows PowerShell

# 4. Drop files into rawdocs/, then run
python scripts/convert.py
```

Output appears in `md_ready/`, one Markdown file per source document.

The project root is auto-detected from the script location — clone it anywhere and it just works.

---

## One-click setup on a fresh Windows machine

For a Windows 10/11 box with **no Python, no WSL, no Node**, the `deploy/` folder bootstraps the entire stack.

### Step 1 — Enable WSL2 + Ubuntu (Admin PowerShell)

```powershell
cd C:\dev\knowledge-base\deploy
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\01-setup-wsl.ps1
```

**Reboot** when prompted.

### Step 2 — Install project + dependencies + Claude Code (Admin PowerShell)

```powershell
cd C:\dev\knowledge-base\deploy
.\02-setup-project.ps1
```

This will:
1. Create `C:\dev\knowledge-base\` and all subdirectories
2. Copy `scripts/`, `CLAUDE.md`, `deploy/` into place
3. Run `03-setup-ubuntu.sh` inside WSL to install Python deps + Claude Code CLI
4. Prompt you for your `ANTHROPIC_API_KEY` and save it to `~/.bashrc`

### Step 3 — Use it

```powershell
.\04-launch.ps1          # Launch Claude Code in the project
.\quick-convert.ps1      # Run the converter headlessly
```

Full deployment details and troubleshooting in [`deploy/README.md`](deploy/README.md).

> The deploy scripts assume `C:\dev\knowledge-base\` as the install path. If you want it elsewhere, install manually using the [Quick start](#quick-start) above instead.

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
| `ANTHROPIC_API_KEY` | Required for any Vision API call | — |
| `KB_PROJECT_DIR` | Override project root | parent of `scripts/` directory |
| `KB_MODEL` | Claude model to use | `claude-sonnet-4-6` |
| `KB_MAX_TOKENS` | Max output tokens per Vision call | `8192` |
| `GOOGLE_API_KEY` | Optional — enables Google Vision OCR for higher text accuracy on diagrams | — (off) |

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

## Optional: Google Vision OCR for higher text accuracy

For pages with critical numeric codes, identifiers, or tightly-packed text, you can enable a hybrid Google Vision OCR + Claude Vision flow. Google extracts exact text characters; Claude interprets structure and meaning.

Set `GOOGLE_API_KEY` and the pipeline will automatically prepend OCR text to each Vision prompt as a reference. No code changes needed.

This is **off by default** — Claude Vision alone is sufficient for most documents.

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
