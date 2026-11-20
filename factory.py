"""Local desktop dashboard. Run python factory.py and open the printed URL."""
from io import BytesIO
from docx import Document
import sessions
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
from analysis_exports import WORD, ZIP, export

BASE = Path(__file__).resolve().parent
DATA = Path(os.environ.get("WAN_FACTORY_DATA", Path.home() / ".wan-booklists")).resolve()
LOCK = threading.Lock()
STATE = {"phase": "idle"}
REPORTS = ("regional_totals.csv", "school_counts.csv", "flagged.csv", "duplicate_schools.csv",
           "product_mapping.csv", "raw_rows.csv", "booklists_wide.csv", "long.csv", "blanks.csv", "assumptions.txt")


def published(root):
    return sorted((p for p in root.glob("*/summary.json")), reverse=True)


def project_for(job):
    file = job / "project.txt"
    return file.read_text(encoding="utf-8") if file.exists() else "Booklists"


def records(path):
    with path.open(encoding="utf-8-sig", newline="") as stream:
        yield from csv.DictReader(stream)


def run(batch, archive):
    global STATE
    job = archive / batch
    project = project_for(job)
    try:
        with TemporaryDirectory(dir=archive) as tmp:
            root = Path(tmp)
            for src in [p.parent for p in reversed(published(archive)) if project_for(p.parent) == project] + [job]:
                for file in (src / "uploads").glob("*/*.docx"):
                    target = root / file.relative_to(src / "uploads")
                    target.parent.mkdir(exist_ok=True)
                    target.unlink(missing_ok=True)
                    os.link(file, target)
            sources = sorted(p.relative_to(root).as_posix() for p in root.glob("*/*.docx"))
            files = len(sources)

            def progress(done, total, current):
                global STATE
                STATE = dict(owner=archive.name, phase="processing", project=project, done=done, total=total, current=current, sources=sources, batch=batch)

            output = job / "reports"
            process(root, output, progress_callback=progress)
            STATE = {"owner": archive.name, "phase": "collecting", "project": project, "total": files}
            flags = records(output / "flagged.csv")
            first_flags = []
            flagged_sources = set()
            count = 0
            for count, flag in enumerate(flags, 1):
                flagged_sources.add(flag["source_file"])
                if count <= 25:
                    first_flags.append(flag)
            summary = dict(id=batch, project=project, files=files, date=datetime.now(timezone.utc).isoformat(),
                           totals=list(records(output / REPORTS[0])), schools=list(records(output / REPORTS[1])),
                           flags=count, preview=first_flags,
                           duplicates=len({r["school_key"] for r in records(output / REPORTS[3])}))
            summary["outcomes"] = {name: "Needs review" if name in flagged_sources else "Complete"
                                   for name in sources}
            format_path = job / "format.json"
            if format_path.exists():
                summary["format"] = json.loads(format_path.read_text(encoding="utf-8"))
            staged = job / "pending.json"
            staged.write_text(json.dumps(summary), encoding="utf-8")
            staged.replace(job / "summary.json")
            STATE = {"owner": archive.name, "phase": "complete", "total": files}
    except Exception as exc:
        STATE = {"owner": archive.name, "phase": "error", "project": project, "message": str(exc)}
        shutil.rmtree(job / "reports", ignore_errors=True)
    finally:
        LOCK.release()


def part(value):
    if not value.strip() or value in (".", "..") or re.search(r'[/\\:\x00-\x1f]', value):
        raise ValueError("Use a plain region or filename.")
    return value


class Handler(BaseHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        if hasattr(self, "session_cookie"):
            self.send_header("Set-Cookie", self.session_cookie)
        super().end_headers()

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
        if url.path in ("/", "/analysis"):
            self.send((BASE / "factory.html").read_bytes(), "text/html; charset=utf-8")
            return
        token = sessions.token_from(self.headers)
        user = sessions.identify(DATA, token)
        if url.path == "/session":
            return self.reply({"email": user[1] if user else None})
        if not user:
            return self.reply({"error": "Sign in to open your workspace."}, 401)
        archive = DATA / "users" / str(user[0])
        if url.path == "/files":
            files = []
            for path in published(archive):
                summary = json.loads(path.read_text(encoding="utf-8"))
                files.append({key: summary[key] for key in ("id", "project", "date", "files")})
            return self.reply({"files": files})
        if url.path == "/state":
            items = published(archive)
            project = parse_qs(url.query).get("project", ["Booklists"])[0]
            projects = {p: project_for(p.parent) for p in items}
            match = next((p for p in items if projects[p] == project), None)
            latest = json.loads(match.read_text(encoding="utf-8")) if match else None
            state = STATE
            if state.get("owner") != archive.name or state.get("project", project) != project:
                state = {"phase": "idle"}
            self.reply(dict(state=state, bucket=latest, reports=REPORTS,
                            projects=["Booklists", *sorted(set(projects.values()) - {"Booklists"})]))
        elif url.path == "/report":
            query = parse_qs(url.query)
            name = query.get("name", [""])[0]
            batch = query.get("id", [""])[0]
            if name not in (*REPORTS, WORD, ZIP) or not re.fullmatch(r"[0-9]{20}-[a-f0-9]{8}", batch) or not (archive / batch / "summary.json").is_file():
                return self.send_error(404)
            path = archive / batch / "reports" / name
            if name in (WORD, ZIP):
                path = export(archive / batch, name, REPORTS)
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
            token = sessions.token_from(self.headers)
            if url.path in ("/signup", "/login", "/logout"):
                age = 0
                if url.path == "/logout":
                    sessions.logout(DATA, token)
                    token = ""
                else:
                    size = int(self.headers.get("Content-Length", 0))
                    if not 0 < size <= 8192:
                        raise ValueError("Enter your email and password.")
                    credentials = json.loads(self.rfile.read(size))
                    if not isinstance(credentials, dict):
                        raise ValueError("Enter your email and password.")
                    token = sessions.authenticate(DATA, credentials.get("email"), credentials.get("password"),
                                                  url.path == "/signup")
                    age = sessions.LIFETIME
                self.session_cookie = f"{sessions.COOKIE}={token}; Path=/; Max-Age={age}; HttpOnly; SameSite=Strict"
                return self.reply({"ok": True})
            user = sessions.identify(DATA, token)
            if not user:
                return self.reply({"error": "Sign in to open your workspace."}, 401)
            archive = DATA / "users" / str(user[0])
            if url.path == "/batch":
                project = args.get("project", ["Booklists"])[0].strip()
                if not project or len(project) > 100 or any(ord(c) < 32 for c in project):
                    raise ValueError("Enter a project name of 1 to 100 characters.")
                kind = args.get("format", [""])[0]
                metadata = {}
                if kind:
                    size = int(self.headers.get("Content-Length", 0))
                    limit = 16 * 1024 * 1024 if kind == "template" else 16000
                    if not 0 < size <= limit:
                        raise ValueError("Choose a template up to 16 MB or enter a shorter description.")
                    body = self.rfile.read(size)
                    if len(body) != size:
                        raise ValueError("Incomplete format upload.")
                    if kind == "template":
                        filename = part(args["name"][0])
                        if not filename.lower().endswith(".docx"):
                            raise ValueError("Choose a .docx report template.")
                        try:
                            Document(BytesIO(body))
                        except Exception as exc:
                            raise ValueError("This template could not be opened as a Word document.") from exc
                        metadata = {"template": filename}
                    elif kind == "description":
                        description = body.decode("utf-8").strip()
                        if not description or any(ord(c) < 32 and c not in "\t\n\r" for c in description):
                            raise ValueError("Enter a readable format description.")
                        metadata = {"description": description}
                    else:
                        raise ValueError("Choose a template or description.")
                batch = datetime.now(timezone.utc).strftime("%Y%m%d%H%M%S%f") + "-" + uuid.uuid4().hex[:8]
                (archive / batch / "uploads").mkdir(parents=True)
                (archive / batch / "project.txt").write_text(project, encoding="utf-8")
                if metadata:
                    (archive / batch / "format.json").write_text(json.dumps(metadata), encoding="utf-8")
                    if kind == "template":
                        (archive / batch / "template.docx").write_bytes(body)
                return self.reply({"id": batch})
            batch = args["id"][0]
            if not re.fullmatch(r"[0-9]{20}-[a-f0-9]{8}", batch):
                raise ValueError("Invalid batch.")
            job = archive / batch
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
                job.rename(archive / final)
                STATE = {"owner": archive.name, "phase": "collecting", "project": project_for(archive / final)}
                threading.Thread(target=run, args=(final, archive), daemon=True).start()
                worker_started = True
                self.reply({"ok": True, "batch": final})
            finally:
                if not worker_started:
                    LOCK.release()
        except (ValueError, KeyError, OSError) as exc:
            self.reply({"error": str(exc)}, 400)


if __name__ == "__main__":
    port = int(os.environ.get("WAN_FACTORY_PORT", "8765"))
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print(f"wan-ocr: http://127.0.0.1:{port}", flush=True)
    server.serve_forever()
