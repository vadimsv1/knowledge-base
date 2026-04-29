#!/usr/bin/env python3
"""
Knowledge Base Document Converter (Final)

Hybrid pipeline:
  - PDF: per-page analysis -> text pages via pymupdf4llm, visual pages via Vision API
  - DOCX: mammoth text + Vision API for embedded images
  - XLSX: openpyxl (exact data, no visual needed)
  - PPTX: Vision API (slides are inherently visual)
  - Images: Vision API with knowledge-extraction prompt
  - Visio: render via pymupdf -> Vision API
  - Text files: direct read

Usage:
    python3 convert.py                  # process all files in rawdocs/
    python3 convert.py --status         # show progress
    python3 convert.py --retry-errors   # retry failed files
    python3 convert.py --reset          # wipe state and md_ready/, reprocess
    python3 convert.py --max-budget 5   # stop after $5 spent on API
"""

import os, sys, json, base64, hashlib, shutil, logging, argparse, traceback, time, re
from pathlib import Path
from datetime import datetime, timezone
from dataclasses import dataclass, field, asdict
from enum import Enum
from io import BytesIO
from html.parser import HTMLParser
from urllib.request import Request, urlopen
from urllib.error import HTTPError

MISSING = []
try: import fitz
except ImportError: MISSING.append("pymupdf")
try: import pymupdf4llm
except ImportError: MISSING.append("pymupdf4llm")
try: import mammoth
except ImportError: MISSING.append("mammoth")
try: import openpyxl
except ImportError: MISSING.append("openpyxl")
try: from pptx import Presentation as PptxPresentation
except ImportError: MISSING.append("python-pptx")
try: from PIL import Image
except ImportError: MISSING.append("Pillow")
try: import anthropic
except ImportError: MISSING.append("anthropic")

if MISSING:
    print(f"Missing: {', '.join(MISSING)}")
    print(f"Run: pip3 install --break-system-packages {' '.join(MISSING)}")
    sys.exit(1)

# Project root: env var override > script-relative (parent of scripts/ dir).
# This lets the project live anywhere on any OS without code changes.
PROJECT_DIR = Path(os.environ.get("KB_PROJECT_DIR") or Path(__file__).resolve().parent.parent)
RAWDOCS_DIR = PROJECT_DIR / "rawdocs"
SORTED_DIR = PROJECT_DIR / "sorted"
MD_READY_DIR = PROJECT_DIR / "md_ready"
STATE_FILE = PROJECT_DIR / "config" / "convert_state.json"
API_KEYS_FILE = PROJECT_DIR / "config" / "api_keys.json"

# Load API keys from config file, fall back to env vars
_api_keys = {}
if API_KEYS_FILE.exists():
    try: _api_keys = json.loads(API_KEYS_FILE.read_text(encoding="utf-8"))
    except Exception: pass
def _key(name): return _api_keys.get(name) or os.environ.get(name, "")

CATEGORY_MAP = {
    ".pdf": "pdf", ".docx": "docx", ".doc": "docx",
    ".xlsx": "xlsx", ".xls": "xlsx",
    ".pptx": "pptx", ".ppt": "pptx",
    ".png": "image", ".jpg": "image", ".jpeg": "image",
    ".gif": "image", ".webp": "image", ".bmp": "image",
    ".tiff": "image", ".tif": "image", ".svg": "image",
    ".vsdx": "visio", ".vsd": "visio",
    ".txt": "text", ".md": "text", ".csv": "text",
    ".json": "text", ".xml": "text", ".html": "text", ".htm": "text",
    ".log": "text", ".yaml": "text", ".yml": "text",
    ".ini": "text", ".cfg": "text", ".conf": "text",
}

MODEL = os.environ.get("KB_MODEL", "claude-sonnet-4-6")
MAX_TOKENS = int(os.environ.get("KB_MAX_TOKENS", "8192"))
GOOGLE_API_KEY = _key("GOOGLE_API_KEY")
GOOGLE_VISION_URL = "https://vision.googleapis.com/v1/images:annotate"
MAX_IMAGE_DIM = 3000
DPI = 300
RATE_LIMIT_DELAY = 1.5
MAX_RETRIES = 3
RETRY_BASE_DELAY = 30
MIN_SIG_IMG_PX = 150
MIN_TEXT_CHARS = 80
MIN_SIG_DRAWINGS = 20
COST_INPUT = 3.0 / 1_000_000
COST_OUTPUT = 15.0 / 1_000_000
EST_IMG_TOKENS = 1600
EST_OUT_TOKENS = 800

DOC_PROMPT = """You are an expert document analyst converting pages to Markdown for a corporate knowledge base.

Rules:
- Extract ALL visible text VERBATIM — every word, number, code, and symbol exactly as shown
- Do NOT summarize, combine, simplify, or paraphrase any content
- Reproduce tables as proper Markdown tables with correct columns and rows
- For diagrams, flowcharts, and sequence diagrams: extract EVERY step individually
  - Each numbered step must be a separate entry with: step number, sender → receiver, message name, and message code
  - Do NOT group or merge steps into phases or summaries
  - Copy all message codes EXACTLY (e.g. 3.000..4B.280HSMPKHSMPKSig)
  - Preserve all parameter lists exactly (e.g. termid,rspcode,SBPTMKPublic,SBPTMKSig,KCV)
  - The total step count in output MUST match the diagram
- For other images/charts: extract all labels, data points, and connections verbatim
- Place image/diagram content naturally where it appears on the page
- Maintain heading hierarchy (# ## ### etc.)
- Mark page boundaries with: <!-- page N -->
- CRITICAL: Output text in the EXACT same language as it appears on the page. Do NOT translate anything.
- Output ONLY clean Markdown. No code fences. No preamble like "Here is..."
- Do NOT add information that is not visible on the pages"""

DOC_PROMPT_CTX = """You are an expert document analyst converting pages to Markdown for a corporate knowledge base.

Context from the previous page (do NOT repeat this, continue seamlessly):
---
{context}
---

Rules:
- Continue the document naturally from where the previous page ended
- If a table or paragraph was split across pages, merge it seamlessly
- Extract ALL visible text VERBATIM — every word, number, code, and symbol exactly as shown
- Do NOT summarize, combine, simplify, or paraphrase any content
- Reproduce tables as proper Markdown tables with correct columns and rows
- For diagrams, flowcharts, and sequence diagrams: extract EVERY step individually
  - Each numbered step must be a separate entry with: step number, sender → receiver, message name, and message code
  - Do NOT group or merge steps into phases or summaries
  - Copy all message codes EXACTLY (e.g. 3.000..4B.280HSMPKHSMPKSig)
  - Preserve all parameter lists exactly
  - The total step count in output MUST match the diagram
- For other images/charts: extract all labels, data points, and connections verbatim
- Place image/diagram content naturally where it appears on the page
- Maintain heading hierarchy (# ## ### etc.)
- Mark page boundaries with: <!-- page N -->
- CRITICAL: Output text in the EXACT same language as it appears on the page. Do NOT translate anything.
- Output ONLY clean Markdown. No code fences. No preamble.
- Do NOT add information that is not visible on the pages"""

IMG_PROMPT = """You are an expert document analyst extracting content from an image for a corporate knowledge base.

Rules:
- Extract ALL visible text, data, labels, numbers, and codes VERBATIM — do NOT summarize or paraphrase
- For diagrams, flowcharts, and sequence diagrams:
  - Extract EVERY step individually as a separate entry
  - Each step must include: step number, sender → receiver, exact message name, and exact message code
  - Do NOT group, merge, or combine steps — list each one separately
  - Copy all message codes EXACTLY as shown (e.g. 3.000..4B.280HSMPKHSMPKSig)
  - Preserve all parameter lists exactly (e.g. termid,rspcode,SBPTMKPublic,SBPTMKSig,KCV)
  - Mark optional steps with asterisk as shown in the diagram (e.g. 1*, 17*)
  - The total step count in output MUST match the diagram exactly
- For tables, charts, infographics: reproduce all data points, labels, and values exactly
- For instructional content: present steps so a reader can follow them without seeing the image
- Explain relationships, hierarchies, flows, and connections between elements
- CRITICAL: Preserve the original language of any text in the image
- Output clean structured Markdown. No code fences. No preamble."""

DOCX_IMG_PROMPT = """You are an expert document analyst extracting content from an embedded image for a knowledge base.

Surrounding document text for context:
---
{context}
---

Rules:
- Extract ALL visible content from the image VERBATIM — do NOT summarize or paraphrase
- For diagrams, flowcharts, and sequence diagrams:
  - Extract EVERY step individually as a separate entry
  - Each step must include: step number, sender → receiver, exact message name, and exact message code
  - Do NOT group, merge, or combine steps — list each one separately
  - Copy all message codes and parameters EXACTLY as shown
  - The total step count must match the diagram exactly
- For other images: include all labels, data points, text visible in the image
- CRITICAL: Use the same language as the surrounding text
- Output clean Markdown. No code fences. No preamble."""

def setup_logging():
    log_dir = PROJECT_DIR / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"convert_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    fmt = logging.Formatter("[%(asctime)s] %(levelname)-7s %(message)s", datefmt="%H:%M:%S")
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(fmt); fh.setLevel(logging.DEBUG)
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt); ch.setLevel(logging.INFO)
    logger = logging.getLogger("convert")
    logger.setLevel(logging.DEBUG)
    logger.addHandler(fh); logger.addHandler(ch)
    return logger

log = setup_logging()

class Status(str, Enum):
    PENDING="pending"; CONVERTED="converted"; SKIPPED="skipped"; ERROR="error"

@dataclass
class FileRecord:
    source_path:str; filename:str; extension:str; category:str
    size_bytes:int; md5_hash:str; status:str=Status.PENDING.value
    output_file:str=""; error_message:str=""; pages_count:int=0
    pages_text:int=0; pages_vision:int=0; method:str=""
    api_calls:int=0; est_cost_usd:float=0.0; converted_at:str=""

@dataclass
class ConvertState:
    created_at:str=""; updated_at:str=""; total_cost_usd:float=0.0
    files:dict=field(default_factory=dict)
    def save(self):
        STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
        self.updated_at = datetime.now(timezone.utc).isoformat()
        data = {"created_at":self.created_at,"updated_at":self.updated_at,
                "total_cost_usd":self.total_cost_usd,
                "files":{k:asdict(v) for k,v in self.files.items()}}
        tmp = STATE_FILE.with_suffix(".tmp")
        tmp.write_text(json.dumps(data,indent=2,ensure_ascii=False),encoding="utf-8")
        tmp.replace(STATE_FILE)
    @classmethod
    def load(cls)->"ConvertState":
        if STATE_FILE.exists():
            try:
                data=json.loads(STATE_FILE.read_text(encoding="utf-8"))
                s=cls(created_at=data.get("created_at",""),updated_at=data.get("updated_at",""),
                      total_cost_usd=data.get("total_cost_usd",0.0))
                for k,v in data.get("files",{}).items(): s.files[k]=FileRecord(**v)
                return s
            except Exception as e: log.warning(f"State corrupted: {e}")
        return cls(created_at=datetime.now(timezone.utc).isoformat())

def file_md5(fp):
    h=hashlib.md5()
    with open(fp,"rb") as f:
        for c in iter(lambda:f.read(8192),b""): h.update(c)
    return h.hexdigest()

def safe_fn(name):
    s="".join(c if c.isalnum() or c in "-_. " else "_" for c in name)
    return s.strip().replace(" ","-").lower()

def unique_out(stem, rel_dir=Path(".")):
    out_dir=MD_READY_DIR/rel_dir; out_dir.mkdir(parents=True,exist_ok=True)
    name=safe_fn(stem)+".md"; p=out_dir/name; n=1
    while p.exists(): name=f"{safe_fn(stem)}_{n}.md"; p=out_dir/name; n+=1
    return p

def frontmatter(rec):
    return (f"---\nsource_file: \"{rec.filename}\"\nsource_type: \"{rec.extension}\"\n"
            f"category: \"{rec.category}\"\nsource_size_bytes: {rec.size_bytes}\n"
            f"source_md5: \"{rec.md5_hash}\"\npages: {rec.pages_count}\n"
            f"pages_text: {rec.pages_text}\npages_vision: {rec.pages_vision}\n"
            f"method: \"{rec.method}\"\napi_calls: {rec.api_calls}\n"
            f"est_cost_usd: {rec.est_cost_usd:.4f}\n"
            f"converted_at: \"{rec.converted_at}\"\n---\n\n")

def get_tail(text, chars=200):
    s=text.strip()
    return s if len(s)<=chars else "..."+s[-chars:]

def est_page_cost():
    return EST_IMG_TOKENS*COST_INPUT + EST_OUT_TOKENS*COST_OUTPUT

_client=None
def get_client():
    global _client
    if _client is None:
        key=_key("ANTHROPIC_API_KEY")
        if not key: log.error("ANTHROPIC_API_KEY not set"); sys.exit(1)
        _client=anthropic.Anthropic(api_key=key)
    return _client

def img_to_b64(img):
    img.thumbnail((MAX_IMAGE_DIM,MAX_IMAGE_DIM),Image.Resampling.LANCZOS)
    if img.mode not in("RGB","RGBA"): img=img.convert("RGB")
    buf=BytesIO(); img.save(buf,format="PNG")
    return base64.standard_b64encode(buf.getvalue()).decode("utf-8"),"image/png"

class VisionAPIError(Exception):
    """Raised when Vision API fails after all retries."""
    pass

def call_vision(images, prompt):
    content=[]
    for b64,mt in images:
        content.append({"type":"image","source":{"type":"base64","media_type":mt,"data":b64}})
    content.append({"type":"text","text":prompt})
    last_error=None
    for attempt in range(1,MAX_RETRIES+1):
        try:
            r=get_client().messages.create(model=MODEL,max_tokens=MAX_TOKENS,
                                           messages=[{"role":"user","content":content}])
            return r.content[0].text
        except anthropic.RateLimitError:
            w=RETRY_BASE_DELAY*attempt
            log.warning(f"    Rate limited, wait {w}s ({attempt}/{MAX_RETRIES})"); time.sleep(w)
            last_error="Rate limit exceeded"
        except anthropic.APIError as e:
            log.error(f"    API error ({attempt}/{MAX_RETRIES}): {e}")
            last_error=str(e)
            if attempt<MAX_RETRIES: time.sleep(5)
        except Exception as e:
            log.error(f"    Error: {e}")
            raise VisionAPIError(f"Vision API failed: {e}") from e
    raise VisionAPIError(f"Vision API failed after {MAX_RETRIES} retries: {last_error}")

def call_google_ocr(b64_data):
    """Send base64 image to Google Vision API, return OCR text or empty string."""
    if not GOOGLE_API_KEY: return ""
    body=json.dumps({"requests":[{
        "image":{"content":b64_data},
        "features":[{"type":"DOCUMENT_TEXT_DETECTION"}]
    }]}).encode("utf-8")
    req=Request(f"{GOOGLE_VISION_URL}?key={GOOGLE_API_KEY}",data=body,method="POST")
    req.add_header("Content-Type","application/json")
    try:
        with urlopen(req,timeout=30) as resp:
            result=json.loads(resp.read().decode("utf-8"))
        r=result.get("responses",[{}])[0]
        if "error" in r:
            log.warning(f"      Google OCR error: {r['error'].get('message','')[:100]}")
            return ""
        return r.get("fullTextAnnotation",{}).get("text","")
    except Exception as e:
        log.warning(f"      Google OCR failed: {e}")
        return ""

def inject_ocr_context(prompt, ocr_text):
    """Prepend exact OCR text to a prompt for accuracy."""
    if not ocr_text or not ocr_text.strip(): return prompt
    return (
        "IMPORTANT: Below is exact OCR text extracted from this image by a high-accuracy OCR engine.\n"
        "Use these exact text values for ALL numbers, codes, message names, identifiers, and technical strings.\n"
        "Do NOT guess or approximate — use the OCR values verbatim.\n\n"
        f"--- EXACT OCR TEXT ---\n{ocr_text.strip()}\n--- END OCR TEXT ---\n\n"
    ) + prompt

def vision_with_ocr(images, prompt):
    """Hybrid vision: Google OCR for text accuracy + Claude for understanding."""
    if images and GOOGLE_API_KEY:
        b64_data=images[0][0]
        ocr=call_google_ocr(b64_data)
        if ocr:
            prompt=inject_ocr_context(prompt,ocr)
            log.info("      + Google OCR context added")
    return call_vision(images,prompt)

class PgType(str,Enum):
    TEXT="text"; VISION="vision"

def analyze_page(doc, pn):
    page=doc[pn]; text=page.get_text("text").strip(); imgs=page.get_images(full=True)
    sig=[]
    for info in imgs:
        try:
            pix=fitz.Pixmap(doc,info[0])
            if pix.width>=MIN_SIG_IMG_PX or pix.height>=MIN_SIG_IMG_PX: sig.append(info)
        except Exception: sig.append(info)
    if sig: return PgType.VISION
    # Detect vector diagrams (Visio exports, flowcharts) — many drawing paths, no bitmaps
    try:
        if len(page.get_drawings()) >= MIN_SIG_DRAWINGS: return PgType.VISION
    except Exception: pass
    if len(text)<MIN_TEXT_CHARS: return PgType.VISION
    return PgType.TEXT

def render_page(doc, pn):
    page=doc[pn]; mat=fitz.Matrix(DPI/72,DPI/72); pix=page.get_pixmap(matrix=mat)
    return Image.frombytes("RGB",[pix.width,pix.height],pix.samples)

# -- HTML to Markdown (for mammoth output) ------------------------------------

class _HtmlToMd(HTMLParser):
    """Convert mammoth HTML to clean Markdown with proper tables."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self._parts = []
        self._stack = []
        # Table state
        self._in_table = False
        self._in_cell = False
        self._cell_buf = []
        self._row = []
        self._rows = []
        self._header_rows = set()
        self._in_thead = False
        # List state
        self._lists = []
        # Link state
        self._link_href = None

    def _out(self, text):
        if self._in_cell:
            self._cell_buf.append(text)
        else:
            self._parts.append(text)

    def handle_starttag(self, tag, attrs):
        ad = dict(attrs)
        self._stack.append(tag)
        if tag in ('h1','h2','h3','h4','h5','h6'):
            self._out('\n\n' + '#' * int(tag[1]) + ' ')
        elif tag == 'p':
            if not self._in_cell:
                self._out('\n\n')
        elif tag in ('strong', 'b'):
            self._out('**')
        elif tag in ('em', 'i'):
            self._out('*')
        elif tag == 'br':
            self._out('  \n' if self._in_cell else '\n')
        elif tag == 'table':
            self._in_table = True; self._rows = []; self._header_rows = set()
        elif tag == 'thead':
            self._in_thead = True
        elif tag == 'tbody':
            self._in_thead = False
        elif tag == 'tr':
            self._row = []
        elif tag in ('th', 'td'):
            self._in_cell = True; self._cell_buf = []
            if tag == 'th':
                self._header_rows.add(len(self._rows))
        elif tag == 'ul':
            self._lists.append(('ul', 0)); self._out('\n')
        elif tag == 'ol':
            self._lists.append(('ol', 0)); self._out('\n')
        elif tag == 'li':
            if self._lists:
                lt, cnt = self._lists[-1]
                indent = '  ' * (len(self._lists) - 1)
                if lt == 'ul':
                    self._out(f'\n{indent}- ')
                else:
                    cnt += 1; self._lists[-1] = (lt, cnt)
                    self._out(f'\n{indent}{cnt}. ')
        elif tag == 'img':
            self._out(f'![{ad.get("alt","")}]({ad.get("src","")})')
        elif tag == 'a':
            href = ad.get('href', '')
            if href:
                self._link_href = href; self._out('[')
            # ignore bookmark anchors (id-only, no href)

    def handle_endtag(self, tag):
        if self._stack and self._stack[-1] == tag:
            self._stack.pop()
        if tag in ('h1','h2','h3','h4','h5','h6'):
            self._out('\n')
        elif tag in ('strong', 'b'):
            self._out('**')
        elif tag in ('em', 'i'):
            self._out('*')
        elif tag in ('th', 'td'):
            cell = ''.join(self._cell_buf).strip().replace('|', '\\|')
            self._row.append(cell)
            self._in_cell = False; self._cell_buf = []
        elif tag == 'tr':
            if self._in_thead:
                self._header_rows.add(len(self._rows))
            self._rows.append(list(self._row))
        elif tag == 'thead':
            self._in_thead = False
        elif tag == 'table':
            self._flush_table(); self._in_table = False
        elif tag in ('ul', 'ol'):
            if self._lists: self._lists.pop()
            self._out('\n')
        elif tag == 'a':
            if self._link_href:
                self._out(f']({self._link_href})')
                self._link_href = None

    def handle_data(self, data):
        self._out(data)

    def _flush_table(self):
        if not self._rows: return
        mc = max((len(r) for r in self._rows), default=0)
        if mc == 0: return
        for r in self._rows:
            r.extend([''] * (mc - len(r)))
        self._parts.append('\n\n')
        if self._header_rows:
            for i, cells in enumerate(self._rows):
                if i in self._header_rows:
                    self._parts.append('| ' + ' | '.join(cells) + ' |\n')
            self._parts.append('| ' + ' | '.join(['---'] * mc) + ' |\n')
            for i, cells in enumerate(self._rows):
                if i not in self._header_rows:
                    self._parts.append('| ' + ' | '.join(cells) + ' |\n')
        else:
            for i, cells in enumerate(self._rows):
                self._parts.append('| ' + ' | '.join(cells) + ' |\n')
                if i == 0:
                    self._parts.append('| ' + ' | '.join(['---'] * mc) + ' |\n')
        self._parts.append('\n')

    def result(self):
        text = ''.join(self._parts)
        return re.sub(r'\n{3,}', '\n\n', text).strip()

def mammoth_html_to_md(html_str):
    p = _HtmlToMd(); p.feed(html_str); return p.result()

# -- CONVERTERS ----------------------------------------------------------------

def convert_pdf(fp, rec, budget):
    doc=fitz.open(str(fp))
    try:
        rec.pages_count=len(doc)
        ptypes=[analyze_page(doc,pn) for pn in range(len(doc))]
        nt=sum(1 for p in ptypes if p==PgType.TEXT)
        nv=sum(1 for p in ptypes if p==PgType.VISION)
        rec.pages_text=nt; rec.pages_vision=nv
        rec.method="hybrid" if nt>0 and nv>0 else ("vision_api" if nv>0 else "programmatic")
        log.info(f"    {rec.pages_count} pages: {nt} text + {nv} vision")
        cost=nv*est_page_cost()
        if cost>budget: raise RuntimeError(f"Need ~${cost:.2f}, budget ${budget:.2f}")
        # text pages
        text_md={}
        if nt>0:
            try:
                chunks=pymupdf4llm.to_markdown(str(fp),page_chunks=True,write_images=False,show_progress=False)
                for c in chunks:
                    pn=c.get("metadata",{}).get("page_number",1)-1
                    txt=c.get("text","")
                    if txt.strip():
                        text_md[pn]=txt
            except Exception as e:
                log.warning(f"    pymupdf4llm error: {e}")
            # fallback for any text page missing from text_md
            for pn in range(len(doc)):
                if ptypes[pn]==PgType.TEXT and (pn not in text_md or not text_md[pn].strip()):
                    fallback=doc[pn].get_text("text")
                    if fallback.strip():
                        text_md[pn]=fallback
        # vision pages with context bridge
        vision_md={}; prev_tail=""
        for pn in range(len(doc)):
            if ptypes[pn]==PgType.TEXT:
                md=text_md.get(pn,"")
                if md.strip(): prev_tail=get_tail(md)
                continue
            img=render_page(doc,pn); b64=img_to_b64(img)
            prompt=DOC_PROMPT_CTX.format(context=prev_tail) if prev_tail else DOC_PROMPT
            log.info(f"    Page {pn+1}/{rec.pages_count} -> Vision API")
            md=vision_with_ocr([b64],prompt)
            md=re.sub(r'^\s*<!--\s*page\s*\d+\s*-->\s*\n?','',md,flags=re.I)
            vision_md[pn]=md
            rec.api_calls+=1; rec.est_cost_usd+=est_page_cost()
            prev_tail=get_tail(md); time.sleep(RATE_LIMIT_DELAY)
    finally:
        doc.close()
    parts=[f"# {fp.stem}\n"]
    for pn in range(rec.pages_count):
        parts.append(f"\n<!-- page {pn+1} -->\n")
        parts.append(vision_md.get(pn, text_md.get(pn,"")))
    rec.status=Status.CONVERTED.value
    return "\n".join(parts)

def convert_docx(fp, rec, budget):
    # Step 1: Extract via mammoth HTML (proper tables) with image placeholders
    img_map = {}  # placeholder_id -> image bytes
    img_counter = [0]

    def handle_image(image):
        """Mammoth image handler: save blob, insert placeholder."""
        img_counter[0] += 1
        img_id = f"__IMG_{img_counter[0]}__"
        with image.open() as f:
            blob = f.read()
        try:
            im = Image.open(BytesIO(blob))
            if im.width >= MIN_SIG_IMG_PX or im.height >= MIN_SIG_IMG_PX:
                img_map[img_id] = blob
                return {"src": img_id}
        except Exception:
            pass
        return {}  # skip tiny/broken images

    with open(fp, "rb") as f:
        result = mammoth.convert_to_html(f, convert_image=mammoth.images.img_element(handle_image))
    md_text = mammoth_html_to_md(result.value)
    rec.pages_count = max(1, md_text.count("\n\n") // 20)
    rec.method = "programmatic"

    if not img_map:
        rec.status = Status.CONVERTED.value
        return f"# {fp.stem}\n\n{md_text}"

    # Step 2: For each placeholder, extract surrounding context and send to Vision API
    log.info(f"    {len(img_map)} image(s) with exact positions, sending to Vision API...")
    rec.method = "hybrid"; rec.pages_vision = len(img_map)
    cost = len(img_map) * est_page_cost()
    if cost > budget: raise RuntimeError(f"Need ~${cost:.2f}, budget ${budget:.2f}")

    for idx, (img_id, blob) in enumerate(img_map.items()):
        im = Image.open(BytesIO(blob)); b64 = img_to_b64(im)

        # Find placeholder position and extract surrounding text as context
        pos = md_text.find(img_id)
        if pos >= 0:
            ctx_start = max(0, pos - 500)
            ctx_end = min(len(md_text), pos + 500)
            context = md_text[ctx_start:ctx_end]
        else:
            context = md_text[:1000]

        prompt = DOCX_IMG_PROMPT.format(context=context)
        log.info(f"    Image {idx+1}/{len(img_map)} -> Vision API")
        desc = vision_with_ocr([b64], prompt)
        rec.api_calls += 1; rec.est_cost_usd += est_page_cost()
        time.sleep(RATE_LIMIT_DELAY)

        # Replace placeholder with AI description inline
        placeholder_md = f"![]({img_id})"
        replacement = f"\n\n{desc}\n\n"
        md_text = md_text.replace(placeholder_md, replacement)

    rec.status = Status.CONVERTED.value
    return f"# {fp.stem}\n\n{md_text}"

def convert_xlsx(fp, rec, budget):
    wb=openpyxl.load_workbook(str(fp),data_only=True,read_only=True)
    parts=[f"# {fp.stem}\n"]
    for sn in wb.sheetnames:
        ws=wb[sn]; rows=[]
        for row in ws.iter_rows(values_only=True):
            if any(c is not None for c in row):
                rows.append([str(c) if c is not None else "" for c in row])
        if not rows: continue
        mc=max(len(r) for r in rows)
        for r in rows: r.extend([""]*(mc-len(r)))
        parts.append(f"\n## Sheet: {sn}\n")
        parts.append("| "+" | ".join(rows[0])+" |")
        parts.append("| "+" | ".join(["---"]*mc)+" |")
        for r in rows[1:]:
            parts.append("| "+" | ".join(c.replace("|","\\|") for c in r)+" |")
    rec.pages_count=len(wb.sheetnames); wb.close()
    rec.method="programmatic"; rec.status=Status.CONVERTED.value
    return "\n".join(parts)

def convert_pptx(fp, rec, budget):
    try: doc=fitz.open(str(fp))
    except Exception:
        log.warning("    Cannot render PPTX, text fallback")
        return _pptx_text(fp,rec)
    rec.pages_count=len(doc); rec.pages_vision=len(doc); rec.method="vision_api"
    cost=rec.pages_count*est_page_cost()
    if cost>budget: doc.close(); raise RuntimeError(f"Need ~${cost:.2f}, budget ${budget:.2f}")
    log.info(f"    {rec.pages_count} slides -> Vision API")
    mds=[]; prev=""
    try:
        for pn in range(len(doc)):
            img=render_page(doc,pn); b64=img_to_b64(img)
            prompt=DOC_PROMPT_CTX.format(context=prev) if prev else DOC_PROMPT
            log.info(f"    Slide {pn+1}/{rec.pages_count} -> Vision API")
            md=vision_with_ocr([b64],prompt); mds.append(md)
            rec.api_calls+=1; rec.est_cost_usd+=est_page_cost()
            prev=get_tail(md); time.sleep(RATE_LIMIT_DELAY)
    finally:
        doc.close()
    rec.status=Status.CONVERTED.value
    return f"# {fp.stem}\n\n"+"\n\n".join(mds)

def _pptx_text(fp, rec):
    prs=PptxPresentation(str(fp)); rec.pages_count=len(prs.slides); rec.method="programmatic"
    parts=[f"# {fp.stem}\n"]
    for i,sl in enumerate(prs.slides):
        txts=[]
        for sh in sl.shapes:
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    t=p.text.strip()
                    if t: txts.append(t)
            if sh.has_table:
                for row in sh.table.rows:
                    txts.append("| "+" | ".join(c.text.strip() for c in row.cells)+" |")
        parts.append(f"\n## Slide {i+1}\n"); parts.append("\n".join(txts) if txts else "_[empty]_")
    rec.status=Status.CONVERTED.value; return "\n".join(parts)

def convert_image(fp, rec, budget):
    rec.pages_count=1; rec.pages_vision=1; rec.method="vision_api"
    cost=est_page_cost()
    if cost>budget: raise RuntimeError("Budget exceeded")
    img=Image.open(str(fp)); b64=img_to_b64(img)
    log.info("    -> Vision API (knowledge extraction)")
    md=vision_with_ocr([b64],IMG_PROMPT)
    rec.api_calls+=1; rec.est_cost_usd+=est_page_cost()
    rec.status=Status.CONVERTED.value; return f"# {fp.stem}\n\n{md}"

def convert_visio(fp, rec, budget):
    rec.method="vision_api"
    try: doc=fitz.open(str(fp))
    except Exception as e:
        rec.status=Status.ERROR.value
        rec.error_message=f"Cannot open. Export to PDF first. ({e})"
        return f"# {fp.stem}\n\n_Export to PDF and re-process._\n"
    rec.pages_count=len(doc); rec.pages_vision=len(doc)
    cost=rec.pages_count*est_page_cost()
    if cost>budget: doc.close(); raise RuntimeError("Budget exceeded")
    mds=[]
    try:
        for pn in range(len(doc)):
            img=render_page(doc,pn); b64=img_to_b64(img)
            log.info(f"    Page {pn+1}/{rec.pages_count} -> Vision API")
            md=vision_with_ocr([b64],IMG_PROMPT); mds.append(md)
            rec.api_calls+=1; rec.est_cost_usd+=est_page_cost(); time.sleep(RATE_LIMIT_DELAY)
    finally:
        doc.close()
    rec.status=Status.CONVERTED.value
    return f"# {fp.stem}\n\n"+"\n\n".join(mds)

def convert_text(fp, rec, budget):
    for enc in ("utf-8","utf-8-sig","cp1251","latin-1"):
        try: content=fp.read_text(encoding=enc); break
        except Exception: continue
    else: raise RuntimeError(f"Cannot decode {fp.name}")
    rec.pages_count=1; rec.method="programmatic"; rec.status=Status.CONVERTED.value
    ext=fp.suffix.lower()
    if ext==".md": return content
    if ext==".csv": return f"# {fp.stem}\n\n```csv\n{content}\n```\n"
    if ext==".json": return f"# {fp.stem}\n\n```json\n{content}\n```\n"
    if ext in(".xml",".html",".htm"): return f"# {fp.stem}\n\n```{ext.lstrip('.')}\n{content}\n```\n"
    return f"# {fp.stem}\n\n{content}\n"

CONVERTERS={"pdf":convert_pdf,"docx":convert_docx,"xlsx":convert_xlsx,
            "pptx":convert_pptx,"image":convert_image,"visio":convert_visio,"text":convert_text}

def ensure_dirs():
    for d in(RAWDOCS_DIR,SORTED_DIR,MD_READY_DIR,PROJECT_DIR/"config",PROJECT_DIR/"logs",PROJECT_DIR/"wiki"):
        d.mkdir(parents=True,exist_ok=True)
    for cat in set(CATEGORY_MAP.values()): (SORTED_DIR/cat).mkdir(exist_ok=True)
    (SORTED_DIR/"unsupported").mkdir(exist_ok=True)

def scan_rawdocs(state):
    new=[]
    for fp in sorted(RAWDOCS_DIR.rglob("*")):
        if not fp.is_file() or fp.name.startswith("."): continue
        key=str(fp.relative_to(RAWDOCS_DIR))
        if key in state.files:
            rec=state.files[key]
            if rec.md5_hash==file_md5(fp) and rec.status!=Status.ERROR.value: continue
        ext=fp.suffix.lower(); cat=CATEGORY_MAP.get(ext,"unsupported")
        state.files[key]=FileRecord(source_path=str(fp),filename=fp.name,extension=ext,
                                    category=cat,size_bytes=fp.stat().st_size,md5_hash=file_md5(fp))
        new.append(fp)
    return new

def sort_file(fp, rec, rel_dir=Path(".")):
    dd=SORTED_DIR/rec.category/rel_dir; dd.mkdir(parents=True,exist_ok=True)
    dest=dd/fp.name
    if dest.exists() and file_md5(dest)!=rec.md5_hash:
        n=1
        while dest.exists(): dest=dd/f"{fp.stem}_{n}{fp.suffix}"; n+=1
    if not dest.exists(): shutil.copy2(fp,dest)
    return dest

def convert_one(fp, rec, budget, rel_dir=Path(".")):
    if rec.category=="unsupported":
        rec.status=Status.SKIPPED.value; rec.error_message=f"Unsupported: {rec.extension}"
        log.warning(f"    SKIP ({rec.extension})"); return False
    conv=CONVERTERS.get(rec.category)
    if not conv: rec.status=Status.ERROR.value; rec.error_message="No converter"; return False
    try:
        md=conv(fp,rec,budget); out=unique_out(fp.stem, rel_dir)
        rec.converted_at=datetime.now(timezone.utc).isoformat()
        rec.output_file=str(out.relative_to(MD_READY_DIR))
        out.write_text(frontmatter(rec)+md,encoding="utf-8"); return True
    except Exception as e:
        rec.status=Status.ERROR.value; rec.error_message=f"{type(e).__name__}: {e}"
        log.error(f"    ERROR: {e}"); log.debug(traceback.format_exc()); return False

def print_status(state):
    c={s.value:0 for s in Status}
    for r in state.files.values(): c[r.status]=c.get(r.status,0)+1
    tt=sum(r.pages_text for r in state.files.values())
    tv=sum(r.pages_vision for r in state.files.values())
    ta=sum(r.api_calls for r in state.files.values())
    print(f"\n{'='*60}\n  Knowledge Base Converter\n{'='*60}")
    print(f"  Files: {len(state.files)} total, {c['converted']} done, {c['error']} err, {c['skipped']} skip, {c['pending']} pend")
    print(f"  Pages: {tt} text (free) + {tv} vision (paid)")
    print(f"  API calls: {ta}  |  Est. cost: ${state.total_cost_usd:.4f}\n{'='*60}")
    errs=[r for r in state.files.values() if r.status=="error"]
    if errs:
        print("  Errors:")
        for r in errs: print(f"    - {r.filename}: {r.error_message}")
    print()

def run_pipeline(max_budget):
    ensure_dirs()
    if not _key("ANTHROPIC_API_KEY"):
        log.warning("ANTHROPIC_API_KEY not set. Only text files will work.")
    state=ConvertState.load()
    log.info("Step 1/3: Scanning rawdocs/ ...")
    new=scan_rawdocs(state)
    if not new:
        log.info("No new files."); print_status(state); state.save(); return
    log.info(f"Found {len(new)} new file(s)")
    log.info("Step 2/3: Sorting ...")
    for fp in new:
        key=str(fp.relative_to(RAWDOCS_DIR)); rec=state.files[key]
        rel_dir=Path(key).parent
        sort_file(fp,rec,rel_dir); log.info(f"  {rec.category:>8s} <- {key}")
    log.info("Step 3/3: Converting ...")
    ok=err=0; t0=time.time()
    for i,fp in enumerate(new,1):
        key=str(fp.relative_to(RAWDOCS_DIR)); rec=state.files[key]
        rel_dir=Path(key).parent
        if rec.category=="unsupported": rec.status=Status.SKIPPED.value; continue
        brem=max_budget-state.total_cost_usd
        log.info(f"[{i}/{len(new)}] {key} (${brem:.2f} left)")
        sfp=SORTED_DIR/rec.category/rel_dir/fp.name
        if not sfp.exists(): sfp=fp
        if convert_one(sfp,rec,brem,rel_dir):
            state.total_cost_usd+=rec.est_cost_usd; ok+=1
            log.info(f"    -> {rec.output_file} ({rec.pages_text}t+{rec.pages_vision}v, ${rec.est_cost_usd:.4f})")
        else: err+=1
        state.save()
        if state.total_cost_usd>=max_budget:
            log.warning(f"Budget ${max_budget:.2f} reached."); break
    log.info(f"\nDone in {time.time()-t0:.1f}s: {ok} ok, {err} err, ${state.total_cost_usd:.4f}")
    print_status(state); state.save()

def main():
    p=argparse.ArgumentParser(description="Knowledge Base Document Converter")
    p.add_argument("--status",action="store_true")
    p.add_argument("--retry-errors",action="store_true")
    p.add_argument("--reset",action="store_true")
    p.add_argument("--max-budget",type=float,default=10.0,help="Max API $ (default: 10)")
    a=p.parse_args()
    if a.status: print_status(ConvertState.load()); return
    if a.reset:
        if STATE_FILE.exists(): STATE_FILE.unlink()
        if MD_READY_DIR.exists(): shutil.rmtree(MD_READY_DIR)
        MD_READY_DIR.mkdir(parents=True,exist_ok=True)
        log.info("Reset done."); return
    if a.retry_errors:
        s=ConvertState.load(); n=0
        for r in s.files.values():
            if r.status==Status.ERROR.value: r.status=Status.PENDING.value; r.error_message=""; n+=1
        s.save(); log.info(f"Reset {n} error(s)")
        if n: run_pipeline(a.max_budget)
        return
    run_pipeline(a.max_budget)

if __name__=="__main__": main()