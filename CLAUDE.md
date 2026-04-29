# Knowledge Base Document Converter

## Project Overview

This project converts company documents (PDF, DOCX, XLSX, PPTX, images, Visio) into clean Markdown files for building an AI-powered corporate knowledge base. The final Markdown files will be compiled into a structured wiki (Karpathy's LLM Wiki pattern) that can be fed to an AI agent as context.

## Architecture

### Pipeline Flow

```
rawdocs/ → sort by type → hybrid conversion → md_ready/
                              │
                    ┌─────────┴─────────┐
                    │                   │
              Text pages            Visual pages
              (pymupdf4llm)         (Claude Vision API)
              FREE, instant         $0.01-0.03/page
                    │                   │
                    └─────────┬─────────┘
                              │
                        Assembly with
                        context bridging
                              │
                         Final .md
```

### Conversion Strategy Per File Type

| File Type | Method | Details |
|-----------|--------|---------|
| **PDF** | Hybrid per-page | Each page analyzed: text-only → pymupdf4llm (free), has images/scans → render as PNG → Vision API |
| **DOCX** | Hybrid | mammoth extracts text with image placeholders → only images sent to Vision API with surrounding context |
| **XLSX** | Programmatic | openpyxl reads exact cell values → Markdown tables. No API needed |
| **PPTX** | Vision API | All slides rendered as PNG → Vision API (slides are inherently visual) |
| **Images** | Vision API | Knowledge-extraction prompt (not just description, but meaning) |
| **Visio** | Vision API | Render via pymupdf → Vision API. If render fails → error with instruction to export to PDF |
| **TXT/CSV/JSON/XML** | Direct read | Encoding detection (utf-8, cp1251, latin-1) |

### Key Design Decisions

1. **Per-page routing for PDF**: pymupdf analyzes each page for text content and significant images (>150px). Pages with only text go through pymupdf4llm (free). Pages with images, diagrams, or scans go through Vision API. This saves 90%+ on API costs for text-heavy documents.

2. **Context bridging**: When a Vision API page follows a text page (or vice versa), the last 200 chars of the previous page are passed in the prompt. This prevents table/paragraph breaks at page boundaries.

3. **Smart prompts**: Document pages get a transcription prompt (preserve original text exactly). Standalone images get a knowledge-extraction prompt (extract meaning, not just describe). DOCX images get context-aware prompt with surrounding text.

4. **Language preservation**: All prompts include strict instructions to output in the same language as the source document. No translation.

5. **mammoth image handler**: DOCX images are detected inline via mammoth's convert_image callback. Each image gets a unique placeholder inserted at its exact position in the Markdown. Context for Vision API is extracted from text surrounding the placeholder — not guessed.

6. **VisionAPIError**: If Vision API fails after all retries, an exception is raised (not a silent error string). The file is marked as ERROR in state, and --retry-errors will pick it up.

## Directory Structure

The project root is auto-detected from the script location (`scripts/convert.py` → `..`). It works from any path on any OS. Override with the `KB_PROJECT_DIR` environment variable to point elsewhere.

```
<project-root>/
├── rawdocs/                 # DROP ZONE: user puts all files here
├── sorted/                  # Auto-sorted copies by category
│   ├── pdf/
│   ├── docx/
│   ├── xlsx/
│   ├── pptx/
│   ├── image/
│   ├── visio/
│   ├── text/
│   └── unsupported/
├── md_ready/                # OUTPUT: final Markdown files
├── wiki/                    # FUTURE: compiled knowledge base
├── scripts/
│   ├── convert.py           # Main conversion script
│   └── run.sh               # Shell wrapper
├── config/
│   └── convert_state.json   # Tracks every file: status, pages, cost, hash
└── logs/
    └── convert_*.log        # Detailed logs per run
```

## Running the Converter

```bash
# Full pipeline
python3 scripts/convert.py

# Check status
python3 scripts/convert.py --status

# Reset everything and reprocess
python3 scripts/convert.py --reset
python3 scripts/convert.py

# Retry only failed files
python3 scripts/convert.py --retry-errors

# Set budget limit
python3 scripts/convert.py --max-budget 5.0
```

## State Management

`config/convert_state.json` tracks every file by MD5 hash. Features:
- **Resume-safe**: state saved after each file. Crash → re-run picks up where it left off.
- **Duplicate detection**: identical files (same hash) are skipped.
- **Modified file detection**: changed files (different hash) are re-processed.
- **Cost tracking**: per-file and total API spend in USD.

## Dependencies

```
pymupdf (fitz)    — PDF rendering, page analysis
pymupdf4llm       — PDF text extraction to Markdown
mammoth           — DOCX to Markdown with image handler
openpyxl          — XLSX cell-level reading
python-pptx       — PPTX text fallback
Pillow            — Image processing
anthropic         — Claude Vision API
```

Install: `pip3 install --break-system-packages pymupdf pymupdf4llm mammoth openpyxl python-pptx Pillow anthropic`

## Environment

| Variable | Purpose | Default |
|---|---|---|
| `ANTHROPIC_API_KEY` | Required for Vision API features | — |
| `KB_PROJECT_DIR` | Override project root path | parent of `scripts/` |
| `KB_MODEL` | Claude model to use | `claude-sonnet-4-6` |
| `KB_MAX_TOKENS` | Max output tokens per Vision call | `8192` |
| `GOOGLE_API_KEY` | Optional — enables Google Vision OCR for higher text accuracy | — (off) |

- Vision API: 300 DPI render, max 3000 px, 8192 max output tokens
- Rate limiting: 1.5s between calls, exponential backoff on 429
- Optional `config/api_keys.json` (JSON map) is loaded ahead of env vars

## Quality Verification Checklist

When testing output quality of md_ready/ files, verify:

1. **PDF files**: All pages have content (no empty `<!-- page N -->` sections). Text pages extracted cleanly. Vision pages have image descriptions inline.
2. **DOCX files**: Text preserved. If original had images, they should be described inline at their correct positions (not appended at the end).
3. **XLSX files**: All sheets present. Tables have correct columns/rows. Pipe characters escaped. No data loss.
4. **PPTX files**: All slides present with both text and visual descriptions.
5. **Images**: Knowledge extracted (meaning, instructions, relationships) — not just visual description.
6. **Language**: Russian documents stay in Russian. English stays in English. No translation.
7. **Tables**: Markdown tables render correctly. Multi-page tables should merge seamlessly via context bridging.
8. **Frontmatter**: Every .md file has YAML frontmatter with source_file, pages, method, api_calls, est_cost_usd.

## Known Limitations

1. **Visio (.vsdx)**: pymupdf support is limited. Best to export to PDF from Visio first.
2. **Old formats (.doc, .xls, .ppt)**: Not natively supported. Would need LibreOffice conversion.
3. **XLSX charts**: openpyxl cannot extract chart images. Charts in Excel are lost.
4. **Synchronous API calls**: Pages processed one at a time. Large presentations are slow (~15s per slide). Future optimization: async with semaphore.
5. **DOCX page count**: Estimated from paragraph count (no real page concept in DOCX).

## Future Stages

### Stage 3 — Wiki Compilation
Claude Code reads all md_ready/ files and compiles a structured wiki:
- Index file with all topics
- Cross-references and backlinks between related documents
- Summaries per topic
- Deduplication of information across documents
- Output in wiki/ directory

### Stage 4 — Validation
- Lint wiki for broken links, contradictions, duplicates
- Verify completeness against source documents

### Stage 5 — Packaging
- Package wiki as context for Claude Projects / Claude Code
- Optimize for context window size

### Stage 6 — Testing
- Ask real company questions and verify answer quality
- Iterate on wiki structure based on results

## Code Style

- All code, comments, variable names, and log messages in English
- Prompts for Vision API in English (with instructions to preserve source language)
- User-facing status output in English