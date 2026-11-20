# wan-ocr

## Booklist Factory

Install and start the local dashboard from this directory.

```sh
python3 -m pip install "python-docx>=1.2" "pandas>=2"
python3 booklists.py --self-test
python3 factory.py
```

Open http://127.0.0.1:8765 on the same computer. Start with 3–5 files. Choose a Booklists folder containing one folder per region, or select individual Word files and enter their region.

The factory reads the 50-column booklists, checks quantities and duplicate school names, then publishes a cumulative analytics bucket. A later upload with the same region and filename replaces that source in subsequent totals. Files with other names remain separate submissions. Check flagged entries and duplicate candidates before using the totals. Download all audit and summary files from the bucket.

Source versions and completed reports persist under `~/.wan-booklists`. Set `WAN_FACTORY_DATA` to choose a different archive (for example, a separate pilot archive), or `WAN_FACTORY_PORT` to choose another local port. Run one factory process per archive.

## OCR to Word

Convert handwritten (or printed) pages into an editable Word document.

Two ways to use it:

### 1. Command-line tool (`ocr2docx.py`)

Accepts images (PNG, JPG, TIFF, BMP, WebP) and PDFs. Every page of a PDF is converted.

```bash
pip install -r requirements.txt
# Also install system packages: tesseract-ocr and poppler-utils
#   Ubuntu/Debian: sudo apt install tesseract-ocr poppler-utils
#   macOS:         brew install tesseract poppler

export ANTHROPIC_API_KEY="your-key"   # for handwriting (Claude vision)

python ocr2docx.py notes.pdf -o notes.docx
python ocr2docx.py page1.jpg page2.jpg --keep-image
python ocr2docx.py printed.pdf --engine tesseract --highlight
```

| Engine | Best for | Needs |
|---|---|---|
| `claude` | Handwriting, messy scans | `ANTHROPIC_API_KEY` |
| `tesseract` | Printed text, offline use | Tesseract installed |
| `auto` (default) | Uses Claude if a key is set, otherwise Tesseract | |

#### Summary and analysis across pages

```bash
python ocr2docx.py notes1.pdf notes2.pdf --analyze --focus "decisions and deadlines"
```
Writes `notes1.docx` (the transcription) and `notes1-analysis.docx`: an overview, key points,
themes across the pages, people and organisations, dates and deadlines, figures, action items,
open questions and a short summary of each page. Claude uses only what the pages say, and every
point names the page it came from. Needs `ANTHROPIC_API_KEY`.

Options: `--lang eng+swa`, `--psm 6`, `--dpi 300`, `--highlight`, `--threshold 60`, `--keep-image`, `--model`, `--analyze`, `--focus`.

### 2. Web page (`handwriting-to-word.html`)

A single-file page: upload images or a PDF, review the text Claude reads, and save a `.docx`. A second button summarises and analyses all pages together and saves that as its own Word document.
It was built as a Claude artifact, so it runs inside claude.ai using the viewer's own Claude access.
To host it on your own site you would need to replace the `claude.use("sample")` call with a request
to your own server that calls the Anthropic API.

### Notes
- Handwriting OCR is never perfect. Words that can't be read are marked `[illegible]`.
- Tables and multi-column layouts come out as plain paragraphs.

### Analysis workspace

On the local dashboard, open **ANALYSIS**. Enter a project name (or select a previous project) and open it. Upload the regional folder or individual .docx files, then start the processing line. Each project accumulates its own submissions. Earlier runs created before project setup are available under **Booklists**.

The final **Analytics bucket** offers a Word summary with school counts, quantities by region and product, and review counts. **Download full ZIP** includes that summary and all ten audit and report files. Both exports cover the whole project, including when the dashboard table is filtered to one region.
