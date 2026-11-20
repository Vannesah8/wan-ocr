"""Local desktop dashboard. Run python factory.py and open the printed URL."""
import csv
import json
import os
import re
import shutil
import threading
import uuid
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from tempfile import TemporaryDirectory
from urllib.parse import parse_qs, urlsplit

from booklists import process

BASE = Path(__file__).resolve().parent
DATA = Path(os.environ.get("WAN_FACTORY_DATA", Path.home() / ".wan-booklists")).resolve()
LOCK = threading.Lock()
STATE = {"phase": "idle"}
REPORTS = ("regional_totals.csv", "school_counts.csv", "flagged.csv", "duplicate_schools.csv",
           "product_mapping.csv", "raw_rows.csv", "booklists_wide.csv", "long.csv", "blanks.csv", "assumptions.txt")


def published():
    return sorted((p for p in DATA.glob("*/summary.json")), reverse=True)


def records(path):
    with path.open(encoding="utf-8-sig", newline="") as stream:
        yield from csv.DictReader(stream)


def run(batch):
    global STATE
    job = DATA / batch
    try:
        with TemporaryDirectory(dir=DATA) as tmp:
            root = Path(tmp)
            for src in [p.parent for p in reversed(published())] + [job]:
                for file in (src / "uploads").glob("*/*.docx"):
                    target = root / file.relative_to(src / "uploads")
                    target.parent.mkdir(exist_ok=True)
                    target.unlink(missing_ok=True)
                    os.link(file, target)
            files = sum(1 for _ in root.glob("*/*.docx"))

            def progress(done, total, current):
                global STATE
                STATE = dict(phase="processing", done=done, total=total, current=current)

            output = job / "reports"
            process(root, output, progress_callback=progress)
            STATE = {"phase": "collecting", "total": files}
            flags = records(output / "flagged.csv")
            first_flags = []
            count = 0
            for count, flag in enumerate(flags, 1):
                if count <= 25:
                    first_flags.append(flag)
            summary = dict(id=batch, files=files, date=datetime.now(timezone.utc).isoformat(),
                           totals=list(records(output / REPORTS[0])), schools=list(records(output / REPORTS[1])),
                           flags=count, preview=first_flags,
                           duplicates=len({r["school_key"] for r in records(output / REPORTS[3])}))
            staged = job / "pending.json"
            staged.write_text(json.dumps(summary), encoding="utf-8")
            staged.replace(job / "summary.json")
            STATE = {"phase": "complete", "total": files}
    except Exception as exc:
        STATE = {"phase": "error", "message": str(exc)}
        shutil.rmtree(job / "reports", ignore_errors=True)
    finally:
        LOCK.release()


def part(value):
    if not value.strip() or value in (".", "..") or re.search(r'[/\\:\x00-\x1f]', value):
        raise ValueError("Use a plain region or filename.")
    return value


class Handler(BaseHTTPRequestHandler):
    def send(self, body, mime, status=200, filename=None):
        self.send_response(status)
        self.send_header("Content-Type", mime)
        if filename:
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def reply(self, data, status=200):
        self.send(json.dumps(data).encode(), "application/json", status)

    def do_GET(self):
        url = urlsplit(self.path)
        if url.path == "/":
            self.send((BASE / "factory.html").read_bytes(), "text/html; charset=utf-8")
        elif url.path == "/state":
            items = published()
            latest = json.loads(items[0].read_text(encoding="utf-8")) if items else None
            self.reply(dict(state=STATE, bucket=latest, reports=REPORTS))
        elif url.path == "/report":
            query = parse_qs(url.query)
            name = query.get("name", [""])[0]
            batch = query.get("id", [""])[0]
            if name not in REPORTS or not re.fullmatch(r"[0-9]{20}-[a-f0-9]{8}", batch) or not (DATA / batch / "summary.json").is_file():
                return self.send_error(404)
            path = DATA / batch / "reports" / name
            self.send_response(200)
            self.send_header("Content-Disposition", f'attachment; filename="{name}"')
            self.send_header("Content-Type", "application/octet-stream")
            self.send_header("Content-Length", str(path.stat().st_size))
            self.end_headers()
            with path.open("rb") as stream:
                shutil.copyfileobj(stream, self.wfile)
        else:
            self.send_error(404)

    def do_POST(self):
        global STATE
        # Only same-origin dashboard requests may write to the local archive.
        if self.headers.get("Origin") != f"http://127.0.0.1:{self.server.server_port}" or self.headers.get("X-Factory") != "1":
            return self.send_error(403)
        url = urlsplit(self.path)
        args = parse_qs(url.query)
        try:
            if url.path == "/batch":
                batch = datetime.now(timezone.utc).strftime("%Y%m%d%H%M%S%f") + "-" + uuid.uuid4().hex[:8]
                (DATA / batch / "uploads").mkdir(parents=True)
                return self.reply({"id": batch})
            batch = args["id"][0]
            if not re.fullmatch(r"[0-9]{20}-[a-f0-9]{8}", batch):
                raise ValueError("Invalid batch.")
            job = DATA / batch
            if not LOCK.acquire(blocking=False):
                return self.reply({"error": "Another batch is processing."}, 409)
            worker_started = False
            try:
                if not (job / "uploads").is_dir() or (job / "summary.json").exists():
                    raise ValueError("Start a fresh batch.")
                if url.path == "/upload":
                    region = part(args["region"][0])
                    name = part(args["name"][0])
                    if not name.lower().endswith(".docx") or name.startswith("~$"):
                        raise ValueError("Choose a Word booklist.")
                    target = job / "uploads" / region / (name[:-5] + ".docx")
                    target.parent.mkdir(exist_ok=True)
                    size = int(self.headers.get("Content-Length", -1))
                    if size < 1 or size > 128 * 1024 * 1024:
                        raise ValueError("Each document must be between 1 byte and 128 MB.")
                    with target.open("xb") as stream:
                        while size:
                            chunk = self.rfile.read(min(size, 65536))
                            if not chunk:
                                raise ValueError("Incomplete upload.")
                            stream.write(chunk)
                            size -= len(chunk)
                    return self.reply({"ok": True})
                if url.path != "/run" or not any((job / "uploads").glob("*/*.docx")):
                    raise ValueError("Add at least one booklist.")
                # Timestamp publication order, independent of upload start order.
                final = datetime.now(timezone.utc).strftime("%Y%m%d%H%M%S%f") + "-" + uuid.uuid4().hex[:8]
                job.rename(DATA / final)
                STATE = {"phase": "collecting"}
                threading.Thread(target=run, args=(final,), daemon=True).start()
                worker_started = True
                self.reply({"ok": True})
            finally:
                if not worker_started:
                    LOCK.release()
        except (ValueError, KeyError, OSError) as exc:
            self.reply({"error": str(exc)}, 400)


if __name__ == "__main__":
    port = int(os.environ.get("WAN_FACTORY_PORT", "8765"))
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print(f"Booklist Factory: http://127.0.0.1:{port}", flush=True)
    server.serve_forever()
