#!/usr/bin/env python3
"""
ocr2docx - convert images containing text into an editable Word document.

Inputs: images (png, jpg, tif, bmp, webp) and PDFs (every page is converted).

Engines
  claude     Uses Claude's vision via the Anthropic API. Best for HANDWRITING,
             messy scans and mixed layouts. Needs ANTHROPIC_API_KEY.
  tesseract  Free, offline, best for PRINTED text. Low-confidence words can be
             highlighted in yellow so you know what to proofread.
   auto       (default) claude if ANTHROPIC_API_KEY is set, otherwise tesseract.

Reliability
   Claude calls retry transient failures with exponential backoff. In auto mode a page falls
   back to Tesseract if Claude stays unavailable. Results are cached under ~/.cache/wan-ocr
   (override with WAN_OCR_CACHE); --no-cache bypasses the cache and --offline forbids all API
   calls. Offline mode forces Tesseract and cannot run --analyze.

Examples
  python ocr2docx.py handwritten_notes.jpg
  python ocr2docx.py scanned_letter.pdf -o letter.docx
  python ocr2docx.py page1.png page2.png -o notes.docx --keep-image
  python ocr2docx.py printed.pdf --engine tesseract --highlight
   python ocr2docx.py notes1.pdf notes2.pdf --analyze --focus "decisions and deadlines"
   python ocr2docx.py scanned.pdf --offline --highlight
"""
import argparse
import base64
import hashlib
import io
import json
import os
import random
import sys
import time
from datetime import date
from pathlib import Path

import cv2
import numpy as np
from docx import Document
from docx.enum.text import WD_BREAK, WD_COLOR_INDEX
from docx.shared import Inches, Pt
from PIL import Image, ImageOps

SUPPORTED = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp", ".pdf"}
CACHE_DIR = Path(os.environ.get("WAN_OCR_CACHE", "~/.cache/wan-ocr")).expanduser()
MAX_ATTEMPTS = 5
RETRYABLE_STATUS = {408, 409, 429, 500, 502, 503, 504}


class ClaudeUnavailable(RuntimeError):
    """Raised when Claude cannot be reached or has refused the request."""


def is_transient(exc):
    if getattr(exc, "status_code", None) in RETRYABLE_STATUS:
        return True
    name = type(exc).__name__.lower()
    return any(part in name for part in ("ratelimit", "timeout", "connection", "overloaded"))


def retry_delay(exc, attempt):
    headers = getattr(getattr(exc, "response", None), "headers", {}) or {}
    header = headers.get("retry-after")
    if header:
        try:
            return min(60.0, max(0.0, float(header)))
        except ValueError:
            pass
    return min(30.0, 1.5 * 2 ** (attempt - 1)) + random.random()


def call_claude(fn, what="Claude request"):
    """Call the API with exponential backoff, then surface a clear CLI error."""
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            return fn()
        except Exception as exc:
            if not is_transient(exc) or attempt == MAX_ATTEMPTS:
                raise ClaudeUnavailable(f"{what}: {exc}") from exc
            delay = retry_delay(exc, attempt)
            reason = getattr(exc, "status_code", type(exc).__name__)
            print(f"  {what} hit {reason}; retrying in {delay:.1f}s "
                  f"({attempt}/{MAX_ATTEMPTS - 1})", file=sys.stderr)
            time.sleep(delay)


def cache_key(kind, parts):
    payload = json.dumps([kind, parts], sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(payload.encode()).hexdigest()


def cache_get(kind, parts, enabled=True):
    if not enabled:
        return None, None
    path = CACHE_DIR / f"{kind}-{cache_key(kind, parts)}.json"
    try:
        return json.loads(path.read_text(encoding="utf-8")), path
    except (OSError, ValueError):
        return None, path


def cache_put(path, value):
    if path is None:
        return
    try:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        tmp = path.with_name(f"{path.name}.{os.getpid()}.tmp")
        tmp.write_text(json.dumps(value), encoding="utf-8")
        tmp.replace(path)
    except OSError as exc:
        print(f"  Cache not written: {exc}", file=sys.stderr)


def image_digest(img):
    return hashlib.sha256(f"{img.mode}{img.size}".encode() + img.tobytes()).hexdigest()


# --------------------------------------------------------------------------
# Image loading and clean-up
# --------------------------------------------------------------------------
def load_image(path: Path) -> Image.Image:
    img = Image.open(path)
    img = ImageOps.exif_transpose(img)  # respect phone-camera rotation
    return img.convert("RGB")


def pdf_to_images(path: Path, dpi: int = 200) -> list:
    """Render every page of a PDF to an image (needs poppler + pdf2image)."""
    try:
        from pdf2image import convert_from_path
    except ImportError:
        sys.exit("PDF support needs:  pip install pdf2image  (and poppler: apt install poppler-utils)")
    return [p.convert("RGB") for p in convert_from_path(str(path), dpi=dpi)]


def deskew(binary: np.ndarray, max_angle: float = 8.0) -> np.ndarray:
    """Straighten slightly tilted text using a projection-profile search."""
    h, w = binary.shape
    scale = 800 / max(h, w) if max(h, w) > 800 else 1.0
    small = cv2.resize(255 - binary, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
    sh, sw = small.shape
    best_angle, best_score = 0.0, -1.0
    for angle in np.arange(-max_angle, max_angle + 0.01, 0.5):
        m = cv2.getRotationMatrix2D((sw / 2, sh / 2), angle, 1.0)
        rotated = cv2.warpAffine(small, m, (sw, sh), flags=cv2.INTER_NEAREST)
        score = float(np.var(rotated.sum(axis=1)))
        if score > best_score:
            best_angle, best_score = angle, score
    if abs(best_angle) < 0.3:
        return binary
    m = cv2.getRotationMatrix2D((w / 2, h / 2), best_angle, 1.0)
    return cv2.warpAffine(binary, m, (w, h), flags=cv2.INTER_CUBIC,
                          borderMode=cv2.BORDER_CONSTANT, borderValue=255)


def preprocess(img: Image.Image) -> np.ndarray:
    """Grayscale -> upscale small images -> denoise -> binarise -> deskew."""
    gray = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2GRAY)
    longest = max(gray.shape)
    if longest < 1800:  # Tesseract works best around 300 DPI
        factor = 1800 / longest
        gray = cv2.resize(gray, None, fx=factor, fy=factor, interpolation=cv2.INTER_CUBIC)
    gray = cv2.fastNlMeansDenoising(gray, None, h=10)
    binary = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 41, 15
    )
    return deskew(binary)


# --------------------------------------------------------------------------
# OCR engines. Each returns a list of paragraphs; a paragraph is a list of
# (word, confidence) tuples. Confidence is None when the engine gives none.
# --------------------------------------------------------------------------
def ocr_tesseract(img: Image.Image, lang: str, psm: int):
    import pytesseract

    processed = preprocess(img)
    data = pytesseract.image_to_data(
        processed, lang=lang, config=f"--oem 3 --psm {psm}",
        output_type=pytesseract.Output.DICT,
    )
    paragraphs, order = {}, []
    for i, word in enumerate(data["text"]):
        word = word.strip()
        conf = float(data["conf"][i])
        if not word or conf < 0:
            continue
        key = (data["block_num"][i], data["par_num"][i])
        if key not in paragraphs:
            paragraphs[key] = []
            order.append(key)
        paragraphs[key].append((word, conf))
    return [paragraphs[k] for k in order]


def ocr_claude(img: Image.Image, model: str):
    try:
        import anthropic
    except ImportError as exc:
        raise ClaudeUnavailable("The 'claude' engine needs:  pip install anthropic") from exc
    if not os.environ.get("ANTHROPIC_API_KEY"):
        raise ClaudeUnavailable("Set the ANTHROPIC_API_KEY environment variable to use the claude engine.")

    work = img.copy()
    work.thumbnail((2000, 2000))
    buf = io.BytesIO()
    work.save(buf, format="JPEG", quality=92)
    b64 = base64.standard_b64encode(buf.getvalue()).decode()

    client = anthropic.Anthropic()
    response = client.messages.create(
        model=model,
        max_tokens=4000,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image",
                 "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}},
                {"type": "text",
                 "text": ("Transcribe all text in this image exactly as written, preserving "
                          "paragraph breaks with a blank line. Do not correct spelling, add "
                          "commentary, or use markdown. If a word is illegible write [illegible]. "
                          "If there is no text, reply with exactly: [no text found]")},
            ],
        }],
    )
    text = "".join(b.text for b in response.content if b.type == "text").strip()
    return [[(w, None) for w in block.split()] for block in text.split("\n\n") if block.strip()]


# --------------------------------------------------------------------------
# Cross-page analysis (needs ANTHROPIC_API_KEY)
# --------------------------------------------------------------------------
AN_PROMPT = (
    "You are analysing text transcribed from several scanned or handwritten pages. Use ONLY what the "
    "text says: never invent facts, names, dates or numbers. Where pages disagree or something is "
    "unclear, mention it under open_questions. Use an empty list for anything the pages do not cover. "
    "Return ONLY JSON with exactly these keys: overview (3 to 5 sentence summary of everything "
    "together), key_points [{point, source}], themes [{theme, detail}] (patterns across pages), "
    "people_and_orgs [string], dates_and_deadlines [{date, what}], figures [{item, value}], "
    "action_items [{task, owner, source}], open_questions [string], per_document [{label, summary}]. "
    "'source' and 'label' are the page labels between the === markers."
)


def analyze_claude(pages, model: str, focus: str = ""):
    """pages: list of (label, text). Returns the analysis as a dict."""
    try:
        import anthropic
    except ImportError as exc:
        raise ClaudeUnavailable("Analysis needs:  pip install anthropic") from exc
    if not os.environ.get("ANTHROPIC_API_KEY"):
        raise ClaudeUnavailable("Set the ANTHROPIC_API_KEY environment variable to use --analyze.")
    docs = "\n\n".join(f"=== {label} ===\n{text}" for label, text in pages if text.strip())
    if not docs:
        sys.exit("No text was found, so there is nothing to analyse.")
    prompt = AN_PROMPT + (f" Pay particular attention to: {focus}." if focus else "") + "\n\n" + docs
    resp = anthropic.Anthropic().messages.create(
        model=model, max_tokens=4000, messages=[{"role": "user", "content": prompt}]
    )
    raw = "".join(b.text for b in resp.content if b.type == "text")
    try:
        return json.loads(raw[raw.index("{"): raw.rindex("}") + 1])
    except ValueError:
        sys.exit("Claude's analysis could not be read as JSON. Try again.")


def analysis_sections(a: dict):
    """Return [(heading, paragraph_or_None, [bullets])] skipping empty sections."""
    def lst(key):
        v = a.get(key)
        return v if isinstance(v, list) else []

    def src(o):
        return f" ({o['source']})" if isinstance(o, dict) and o.get("source") else ""

    def pair(o, k1, k2):
        if not isinstance(o, dict):
            return ""
        x, y = o.get(k1), o.get(k2)
        return f"{x}: {y}" if x and y else str(x or y or "")

    out = []
    if a.get("overview"):
        out.append(("Overview", str(a["overview"]), []))
    groups = [
        ("Key points", [f"{k['point']}{src(k)}" for k in lst("key_points") if isinstance(k, dict) and k.get("point")]),
        ("Themes across the pages", [pair(t, "theme", "detail") for t in lst("themes")]),
        ("People and organisations", [str(x) for x in lst("people_and_orgs")]),
        ("Dates and deadlines", [pair(d, "date", "what") for d in lst("dates_and_deadlines")]),
        ("Figures", [pair(f, "item", "value") for f in lst("figures")]),
        ("Action items", [f"{t['task']}" + (f" - {t['owner']}" if t.get("owner") else "") + src(t)
                          for t in lst("action_items") if isinstance(t, dict) and t.get("task")]),
        ("Open questions and gaps", [str(x) for x in lst("open_questions")]),
        ("Each page in brief", [pair(d, "label", "summary") for d in lst("per_document")]),
    ]
    for heading, items in groups:
        items = [i for i in items if i]
        if items:
            out.append((heading, None, items))
    return out


def write_analysis_docx(a: dict, output: Path, n_pages: int, label: str):
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    doc.add_heading("Document analysis", level=1)
    doc.add_paragraph(
        f"Based on {n_pages} page(s) of {label}, analysed on {date.today():%d %B %Y}. Written by Claude "
        "from the transcribed text, so check important details against the originals."
    )
    for heading, para, items in analysis_sections(a):
        doc.add_heading(heading, level=2)
        if para:
            doc.add_paragraph(para)
        for item in items:
            doc.add_paragraph(item, style="List Bullet")
    doc.save(output)


# --------------------------------------------------------------------------
# Word document builder
# --------------------------------------------------------------------------
def add_paragraphs(doc, paragraphs, highlight: bool, threshold: float):
    flagged = 0
    for words in paragraphs:
        p = doc.add_paragraph()
        for i, (word, conf) in enumerate(words):
            run = p.add_run(word + (" " if i < len(words) - 1 else ""))
            if highlight and conf is not None and conf < threshold:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                flagged += 1
    return flagged


def build_document(results, output: Path, keep_image: bool, highlight: bool, threshold: float):
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)

    total_flagged = 0
    multi = len(results) > 1
    for idx, (label, img, paragraphs) in enumerate(results):
        if idx > 0:
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        if multi:
            doc.add_heading(label, level=2)
        if keep_image:
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            buf.seek(0)
            doc.add_picture(buf, width=Inches(4.5))
        if not paragraphs:
            doc.add_paragraph("[No text detected]")
        total_flagged += add_paragraphs(doc, paragraphs, highlight, threshold)

    doc.save(output)
    return total_flagged


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Convert images with text into a Word document.")
    ap.add_argument("images", nargs="+", type=Path, help="one or more image or PDF files")
    ap.add_argument("-o", "--output", type=Path, help="output .docx (default: first image name)")
    ap.add_argument("--engine", choices=["auto", "claude", "tesseract"], default="auto")
    ap.add_argument("--dpi", type=int, default=200, help="resolution used to render PDF pages")
    ap.add_argument("--lang", default="eng", help="Tesseract language code(s), e.g. eng+swa")
    ap.add_argument("--psm", type=int, default=3,
                    help="Tesseract page mode: 3=auto (default), 6=single block, 4=single column")
    ap.add_argument("--model", default="claude-sonnet-5-5", help="model for the claude engine")
    ap.add_argument("--highlight", action="store_true",
                    help="highlight low-confidence words in yellow (tesseract only)")
    ap.add_argument("--threshold", type=float, default=60.0,
                    help="confidence below which words are highlighted (default 60)")
    ap.add_argument("--keep-image", action="store_true",
                    help="embed the original image above its text in the document")
    ap.add_argument("--analyze", action="store_true",
                    help="also write a summary and analysis across all pages (needs ANTHROPIC_API_KEY)")
    ap.add_argument("--focus", default="", help="what the analysis should pay attention to")
    ap.add_argument("--offline", action="store_true",
                    help="never call the API: force Tesseract and skip analysis")
    ap.add_argument("--no-cache", action="store_true", help="ignore and do not write cached results")
    args = ap.parse_args()

    if args.offline and args.analyze:
        sys.exit("--analyze needs Claude; remove it or run without --offline.")
    engine = args.engine
    if args.offline and engine in ("auto", "claude"):
        engine = "tesseract"
        print("Offline mode: using Tesseract.")
    elif engine == "auto":
        engine = "claude" if os.environ.get("ANTHROPIC_API_KEY") else "tesseract"
        if engine == "tesseract":
            print("Note: ANTHROPIC_API_KEY not set, using Tesseract (weak on handwriting).")

    output = args.output or args.images[0].with_suffix(".docx")
    pages = []  # (label, PIL image)
    for path in args.images:
        if not path.exists():
            sys.exit(f"File not found: {path}")
        if path.suffix.lower() not in SUPPORTED:
            sys.exit(f"Unsupported file type: {path.name} (use {', '.join(sorted(SUPPORTED))})")
        if path.suffix.lower() == ".pdf":
            imgs = pdf_to_images(path, args.dpi)
            pages += [(f"{path.name} - page {i}", im) for i, im in enumerate(imgs, 1)]
        else:
            pages.append((path.name, load_image(path)))

    results = []
    for label, img in pages:
        print(f"Reading {label} with {engine}...")
        if engine == "tesseract":
            kind = "tesseract-ocr-v1"
            key = {"image": image_digest(img), "lang": args.lang, "psm": args.psm}
        else:
            kind = "claude-ocr-v1"
            key = {"image": image_digest(img), "model": args.model}
        paragraphs, path = cache_get(kind, key, not args.no_cache)
        if paragraphs is not None:
            print("  cache hit")
        elif engine == "tesseract":
            paragraphs = ocr_tesseract(img, args.lang, args.psm)
            cache_put(path, paragraphs)
        else:
            try:
                paragraphs = call_claude(lambda: ocr_claude(img, args.model), f"OCR of {label}")
            except ClaudeUnavailable as exc:
                if engine != "claude":
                    print(f"  {exc}; falling back to Tesseract for {label}", file=sys.stderr)
                    paragraphs = ocr_tesseract(img, args.lang, args.psm)
                else:
                    raise
            else:
                cache_put(path, paragraphs)
        n_words = sum(len(p) for p in paragraphs)
        print(f"  {len(paragraphs)} paragraph(s), {n_words} word(s)")
        results.append((label, img, paragraphs))

    flagged = build_document(results, output, args.keep_image, args.highlight, args.threshold)
    print(f"Saved {output}" + (f"  ({flagged} low-confidence words highlighted)" if flagged else ""))

    if args.analyze:
        print("Analysing all pages with Claude...")
        texts = [(label, "\n\n".join(" ".join(w for w, _ in para) for para in paras))
                 for label, _, paras in results]
        key = {"model": args.model, "focus": args.focus, "pages": texts}
        analysis, path = cache_get("analysis-v1", key, not args.no_cache)
        if analysis is not None:
            print("  cache hit")
        else:
            try:
                analysis = call_claude(
                    lambda: analyze_claude(texts, args.model, args.focus), "Analysis")
            except ClaudeUnavailable as exc:
                print(f"Analysis skipped: {exc}", file=sys.stderr)
                return
            cache_put(path, analysis)
        out2 = output.with_name(output.stem + "-analysis.docx")
        write_analysis_docx(analysis, out2, len(results), args.images[0].name)
        print(f"Saved {out2}")


if __name__ == "__main__":
    main()
