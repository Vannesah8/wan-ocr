"""Text conversion and keyword-based filing of original documents."""
import json
import re
import uuid
from datetime import datetime, timezone
from email import policy
from email.parser import BytesParser
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile, ZIP_DEFLATED

from docx import Document
from docx.table import Table


def form(handler):
    size = int(handler.headers.get("Content-Length", 0))
    if not 0 < size <= 128 * 1024**2:
        raise ValueError("Choose a batch under 128 MB.")
    mime = handler.headers.get("Content-Type", "")
    if not mime.startswith("multipart/form-data;"):
        raise ValueError("Choose files using the upload form.")
    message = BytesParser(policy=policy.default).parsebytes(
        f"Content-Type: {mime}\r\n\r\n".encode() + handler.rfile.read(size))
    fields, documents = {}, []
    for item in message.iter_parts():
        data = item.get_payload(decode=True)
        name = item.get_filename()
        if name:
            if Path(name).name != name or re.search(r'[\\:\x00-\x1f]', name) or name in (".", ".."):
                raise ValueError("Use plain document filenames.")
            documents.append((name, data))
        else:
            fields[item.get_param("name", header="content-disposition")] = data.decode("utf-8")
    if not documents:
        raise ValueError("Choose at least one document.")
    return fields, documents


def extract(name, data):
    ext = Path(name).suffix.lower()
    if ext == ".txt":
        return data.decode("utf-8-sig")
    if ext == ".docx":
        blocks = []
        for block in Document(BytesIO(data)).iter_inner_content():
            if isinstance(block, Table):
                blocks.extend("\t".join(cell.text for cell in row.cells) for row in block.rows)
            else:
                blocks.append(block.text)
        return "\n".join(blocks)
    if ext == ".pdf":
        import pymupdf
        with pymupdf.open(stream=data, filetype="pdf") as pdf:
            pages = []
            for page in pdf:
                text = page.get_text(sort=True)
                if not text.strip() and page.get_images():
                    raise ValueError("This PDF contains scanned pages. Run OCR before converting it.")
                pages.append(text)
            return "\n".join(pages)
    raise ValueError("Choose a Word, PDF or UTF-8 text document.")


def convert(name, data, target, dest):
    if target not in ("docx", "txt", "pdf"):
        raise ValueError("Choose a supported final format.")
    if Path(name).suffix.lower() == "." + target:
        raise ValueError("Choose a different final format.")
    content = extract(name, data)
    if not content.strip():
        raise ValueError("No readable text found. Run OCR on scanned pages first.")
    path = dest / ("converted." + target)
    if target == "txt":
        path.write_text(content, encoding="utf-8")
    elif target == "docx":
        doc = Document()
        for line in content.splitlines():
            doc.add_paragraph(line)
        doc.save(path)
    else:
        import pymupdf
        from html import escape
        story = pymupdf.Story(html='<pre style="white-space:pre-wrap">'+escape(content)+'</pre>')
        page = pymupdf.paper_rect("a4")
        with pymupdf.DocumentWriter(str(path)) as writer:
            more = True
            while more:
                device = writer.begin_page(page)
                more, _ = story.place(page + (36, 36, -36, -36))
                story.draw(device)
                writer.end_page()
    return path.name


def categories(spec):
    groups = {}
    for line in spec.splitlines():
        if not line.strip():
            continue
        label, _, terms = line.partition(":")
        label = label.strip()
        if (not label or len(label)>80 or re.search(r'[\\/:\x00-\x1f]', label)
                or label in (".","..") or label.casefold() in {"needs review", *(k.casefold() for k in groups)}):
            raise ValueError("Choose distinct category names without slashes. Needs review is reserved.")
        groups[label] = {term.strip().casefold() for term in [label, *terms.split(",")] if term.strip()}
    if not groups:
        raise ValueError("Create at least one category.")
    return groups


def compile_files(documents, spec, dest):
    groups = categories(spec)
    results = []
    used = set()
    with ZipFile(dest / "categorised-files.zip", "w", ZIP_DEFLATED) as archive:
        for name, data in documents:
            label, reason = "Needs review", "No matching terms"
            try:
                body = extract(name, data)
                haystack = (Path(name).stem.replace("_"," ") + "\n" + body).casefold()
                scores = {group: sum(bool(re.search(r"(?<!\w)"+re.escape(term)+r"(?!\w)",haystack))
                                     for term in terms) for group, terms in groups.items()}
                best = max(scores.values())
                winners = [g for g in scores if scores[g] == best]
                if best and len(winners)==1:
                    label, reason = winners[0], "Matched category terms"
                elif best:
                    reason = "Equal matches across categories"
            except Exception:
                reason = "Document could not be read"
            candidate = name
            suffix = 2
            while (label,candidate.casefold()) in used:
                candidate = f"{Path(name).stem} ({suffix}){Path(name).suffix}"
                suffix += 1
            used.add((label,candidate.casefold()))
            archive.writestr(f"{label}/{candidate}",data)
            results.append(dict(file=candidate,category=label,reason=reason))
    return results


def process_service(service, fields, documents, archive):
    root = archive / "services"
    root.mkdir(parents=True, exist_ok=True)
    job = uuid.uuid4().hex
    with TemporaryDirectory(dir=root) as tmp:
        dest = Path(tmp) / "result"
        dest.mkdir()
        entries = []
        if service == "conversion":
            if len(documents)!=1:
                raise ValueError("Choose one document for conversion.")
            name, data = documents[0]
            try:
                output = convert(name, data, fields.get("target"), dest)
            except ValueError:
                raise
            except Exception as exc:
                raise ValueError("The source document could not be converted.") from exc
            title = name
        else:
            entries = compile_files(documents,fields.get("categories",""),dest)
            output, title = "categorised-files.zip", "Compiler"
        result = dict(id=job,project=title,service=service,date=datetime.now(timezone.utc).isoformat(),
                      files=len(documents),output=output,entries=entries)
        (dest / "result.json").write_text(json.dumps(result),encoding="utf-8")
        dest.rename(root / job)
    return result
