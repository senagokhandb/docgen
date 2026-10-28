#!/usr/bin/env python3
"""DOCGEN chat UI (host side, stdlib only + docker).

Ask for a document in the browser, watch the agent work, preview the result
next to the chat, upload a file and ask for changes. Every message runs the
REAL system exactly like the CLI does — same image, same entrypoint, `docgen`
network with Gotenberg, .env — in a throw-away container:

    data/ui/<session>/files/          every upload and deliverable of the chat
    data/ui/<session>/turns/<n>/ws/   /workspace of message n
        uploads/   copies of files/ (read-only mount: the original stays intact)
        scratch/   the agent's working area
        outputs/   deliverables of message n -> copied back into files/
        previous/  mount points: /workspace/previous/<k>/ = scratch/ of message k (read-only)

Deliverables of earlier messages are uploads of the next one, so a follow-up
("başlığı kırmızı yap") edits the file the previous answer produced. The
scratch/ of earlier messages comes along too: the HTML a PDF was rendered
from, a deck's build script, the half-done work of a message that timed out.
Previews are Office->PDF conversions done by Gotenberg (cached per file).

    python3 ui/server.py                        # http://127.0.0.1:8765
    python3 ui/server.py --port 9000 --timeout 1800
"""
import argparse
import json
import mimetypes
import re
import shutil
import subprocess
import sys
import threading
import time
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse

import xlsx_view

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
# Shared with the MCP service, so a chat message and an MCP job get the same prompt (stdlib only).
from docgen_service.conversation import (  # noqa: E402
    DELIVERABLE, EVENTS, OFFICE, PREVIOUS, UPLOADABLE, build_prompt, clean_name, entries,
    events_from_stdout, read_events, summarize, unique_name,
)

STATIC = Path(__file__).resolve().parent / "static"
GOTENBERG = "http://gotenberg:3000"
SHEETS = {".xlsx", ".xlsm", ".xltx", ".csv"}  # shown as a grid (xlsx_view); everything else as PDF
MAX_UPLOAD = 200 * 1024 * 1024

for ext, typ in ((".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
                 (".pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
                 (".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                 (".md", "text/markdown")):
    mimetypes.add_type(typ, ext)


def sh(cmd, **kw):
    return subprocess.run(cmd, capture_output=True, text=True, **kw)


def model_name():
    try:
        for line in (ROOT / ".env").read_text().splitlines():
            if line.startswith("MODEL_NAME="):
                return line.split("=", 1)[1].strip()
    except OSError:
        pass
    return ""


# ---- sessions ---------------------------------------------------------------
class Store:
    def __init__(self, base, image, timeout):
        self.base, self.image, self.timeout = base, image, timeout
        self.lock = threading.RLock()
        self.sessions, self.ctl, self.live, self.previews, self.sheets = {}, {}, {}, {}, {}
        base.mkdir(parents=True, exist_ok=True)
        for f in sorted(base.glob("*/session.json")):
            try:
                sess = json.loads(f.read_text())
            except ValueError:
                print(f"bozuk oturum atlandı: {f}", file=sys.stderr)
                continue
            stale = [t for t in sess["turns"] if t["status"] == "running"]
            for t in stale:  # the server died while this message ran
                sh(["docker", "rm", "-f", self.container(sess["id"], t["n"])])
                t.update(status="interrupted", finished=time.time(),
                         error="Sunucu yeniden başlatıldı; bu istek yarıda kaldı.")
            self.sessions[sess["id"]] = sess
            if stale:
                self.save(sess)

    @staticmethod
    def container(sid, n):
        return f"docgen-ui-{sid}-{n}"

    def dir(self, sid):
        return self.base / sid

    def scratch(self, sid, n):
        return self.dir(sid) / "turns" / str(n) / "ws" / "scratch"

    def work(self, sess, n):
        """{k: entries} for every message before n that left files in its scratch/."""
        found = {t["n"]: entries(self.scratch(sess["id"], t["n"])) for t in sess["turns"] if t["n"] < n}
        return {k: v for k, v in found.items() if v}

    def save(self, sess):
        sess["updated"] = time.time()
        path = self.dir(sess["id"]) / "session.json"
        tmp = path.with_suffix(".tmp")
        tmp.write_text(json.dumps(sess, ensure_ascii=False, indent=1))
        tmp.replace(path)

    def create(self):
        sid = uuid.uuid4().hex[:12]
        sess = {"id": sid, "title": "", "created": time.time(), "updated": time.time(), "files": [], "turns": []}
        (self.dir(sid) / "files").mkdir(parents=True)
        with self.lock:
            self.sessions[sid] = sess
            self.save(sess)
        return sess

    def get(self, sid):
        return self.sessions.get(sid)

    def running(self, sess):
        return next((t for t in sess["turns"] if t["status"] == "running"), None)

    def delete(self, sid):
        with self.lock:
            sess = self.sessions.get(sid)
            if not sess or self.running(sess):
                return False
            del self.sessions[sid]
            self.sheets = {k: v for k, v in self.sheets.items() if k[0] != sid}
        shutil.rmtree(self.dir(sid), ignore_errors=True)
        return True

    def file(self, sess, name):
        return next((f for f in sess["files"] if f["name"] == name), None)

    def add_file(self, sess, name, src_or_stream, length=None, origin="upload", turn=None):
        with self.lock:
            name = unique_name({f["name"] for f in sess["files"]}, clean_name(name))
            rec = {"name": name, "origin": origin, "turn": turn, "size": 0, "added": time.time()}
            sess["files"].append(rec)  # reserve the name
        dest = self.dir(sess["id"]) / "files" / name
        try:
            if isinstance(src_or_stream, Path):
                shutil.copyfile(src_or_stream, dest)
            else:
                left = length
                with open(dest, "wb") as fh:
                    while left > 0:
                        chunk = src_or_stream.read(min(left, 1 << 20))
                        if not chunk:
                            raise IOError("upload interrupted")
                        fh.write(chunk)
                        left -= len(chunk)
        except Exception:
            with self.lock:
                sess["files"].remove(rec)
            dest.unlink(missing_ok=True)
            raise
        with self.lock:
            rec["size"] = dest.stat().st_size
            self.save(sess)
        return rec

    def remove_pending(self, sess, name):
        with self.lock:
            rec = self.file(sess, name)
            if not rec or rec["origin"] != "upload" or rec["turn"] is not None:
                return False
            sess["files"].remove(rec)
            self.sheets.pop((sess["id"], name), None)  # the name may be reused by the next upload
            self.save(sess)
        (self.dir(sess["id"]) / "files" / name).unlink(missing_ok=True)
        (self.dir(sess["id"]) / "previews" / (name + ".pdf")).unlink(missing_ok=True)
        return True

    # -- views --
    def steps(self, sid, turn):
        tdir = self.dir(sid) / "turns" / str(turn["n"])
        if turn["status"] != "running":
            try:
                return json.loads((tdir / "steps.json").read_text())
            except (OSError, ValueError):
                return []
        path = tdir / "ws" / EVENTS
        try:
            st = path.stat()
        except FileNotFoundError:
            return []
        key, sig = (sid, turn["n"]), (st.st_size, st.st_mtime_ns)
        cached = self.live.get(key)
        if not cached or cached[0] != sig:
            cached = self.live[key] = (sig, summarize(read_events(path))[0])
        return cached[1]

    def turn_view(self, sid, turn):
        return {**turn, "steps": self.steps(sid, turn)}

    def view(self, sess):
        with self.lock:
            sess = json.loads(json.dumps(sess))
        sess["turns"] = [self.turn_view(sess["id"], t) for t in sess["turns"]]
        return sess

    def summary(self, sess):
        return {"id": sess["id"], "title": sess["title"] or "Yeni sohbet", "updated": sess["updated"],
                "running": bool(self.running(sess))}

    # -- messages --
    def start_turn(self, sess, text, active, confirm):
        with self.lock:
            if self.running(sess):
                return None
            n = len(sess["turns"]) + 1
            attachments = []
            for f in sess["files"]:  # every upload not sent yet belongs to this message
                if f["origin"] == "upload" and f["turn"] is None:
                    f["turn"] = n
                    attachments.append(f["name"])
            turn = {"n": n, "text": text, "attachments": attachments,
                    "active": active if self.file(sess, active or "") else None, "confirm": bool(confirm),
                    "status": "running", "started": time.time(), "finished": None, "reply": "", "error": "",
                    "outputs": [], "stats": {}}
            sess["turns"].append(turn)
            if not sess["title"]:
                sess["title"] = text if len(text) <= 60 else text[:60].rstrip() + "…"
            self.save(sess)
            self.ctl[(sess["id"], n)] = {"cancelled": False, "proc": None, "done": False}
        threading.Thread(target=self.run_turn, args=(sess, turn), daemon=True).start()
        return turn

    def cancel(self, sess, n):
        ctl = self.ctl.get((sess["id"], n))
        if not ctl:
            return False
        ctl["cancelled"] = True

        def kill():  # the container may not exist yet: keep trying until the run is over
            while not ctl["done"]:
                sh(["docker", "kill", self.container(sess["id"], n)])
                time.sleep(1)
        threading.Thread(target=kill, daemon=True).start()
        return True

    def run_turn(self, sess, turn):
        sid, n = sess["id"], turn["n"]
        tdir = self.dir(sid) / "turns" / str(n)
        ws, ctl = tdir / "ws", self.ctl[(sid, n)]
        timed_out, rc = False, None
        try:
            for d in ("uploads", "scratch", "outputs"):
                (ws / d).mkdir(parents=True, exist_ok=True)
            with self.lock:
                files = [f["name"] for f in sess["files"] if f["turn"] is not None and f["turn"] <= n]
                work = self.work(sess, n)
                prompt = build_prompt(sess, turn, work)
            for name in files:
                shutil.copyfile(self.dir(sid) / "files" / name, ws / "uploads" / name)
            (tdir / "prompt.txt").write_text(prompt)
            sh(["chmod", "-R", "a+rwX", str(tdir)])  # the agent runs as uid 1000
            name = self.container(sid, n)
            mounts = ["-v", f"{ws}:/workspace", "-v", f"{ws / 'uploads'}:/workspace/uploads:ro"]
            for k in work:  # read-only: a follow-up must not change what an earlier message left behind
                mounts += ["-v", f"{self.scratch(sid, k)}:{PREVIOUS}/{k}:ro"]
            cmd = ["docker", "run", "--rm", "--name", name, "--network", "docgen", "--env-file", str(ROOT / ".env"),
                   "-e", f"GOTENBERG_URL={GOTENBERG}", "-e", "DOCGEN_OUTPUT_FORMAT=stream-json",
                   "-e", f"DOCGEN_OUT_FILE=/workspace/{EVENTS}", *mounts, self.image, prompt]
            with open(tdir / "stdout.txt", "wb") as out, open(tdir / "stderr.txt", "wb") as err:
                if not ctl["cancelled"]:
                    ctl["proc"] = subprocess.Popen(cmd, stdout=out, stderr=err)
                    try:
                        rc = ctl["proc"].wait(timeout=self.timeout)
                    except subprocess.TimeoutExpired:
                        timed_out = True
                        sh(["docker", "kill", name])
                        rc = ctl["proc"].wait()
            self.finish(sess, turn, tdir, rc, timed_out, ctl["cancelled"])
        except Exception as e:  # never leave a message "running"
            with self.lock:
                turn.update(status="error", finished=time.time(), error=f"Arayüz hatası: {type(e).__name__}: {e}")
                self.save(sess)
        finally:
            ctl["done"] = True
            self.ctl.pop((sid, n), None)
            self.live.pop((sid, n), None)

    def finish(self, sess, turn, tdir, rc, timed_out, cancelled):
        ws = tdir / "ws"
        stdout = (tdir / "stdout.txt").read_text(errors="replace")
        steps, result = summarize(read_events(ws / EVENTS) or events_from_stdout(stdout))
        (tdir / "steps.json").write_text(json.dumps(steps, ensure_ascii=False))
        outputs = []
        for p in sorted((ws / "outputs").iterdir()) if (ws / "outputs").is_dir() else []:
            if p.is_file() and p.suffix.lower() in DELIVERABLE and not p.name.startswith("."):
                outputs.append(self.add_file(sess, p.name, p, origin="output", turn=turn["n"])["name"])
        stderr = [x for x in (tdir / "stderr.txt").read_text(errors="replace").splitlines()
                  if x.strip() and "running headless with --yolo" not in x]
        error = ""
        if cancelled:
            status = "cancelled"
        elif timed_out:
            status, error = "error", f"Zaman aşımı: istek {self.timeout} sn içinde bitmedi."
        elif rc != 0 or not result or result["is_error"]:
            status = "error"
            error = (result or {}).get("error") or "\n".join(stderr[-15:]) or f"Ajan {rc} çıkış koduyla bitti."
        else:
            status = "done"
        with self.lock:
            turn.update(status=status, finished=time.time(), reply=(result or {}).get("reply", ""), error=error,
                        outputs=outputs, stats={"duration_ms": (result or {}).get("duration_ms"),
                                                "num_turns": (result or {}).get("num_turns"), "steps": len(steps)})
            self.save(sess)

    # -- preview --
    def preview(self, sess, rec):
        """-> (path, None) | (None, error). Office files are converted to PDF by Gotenberg once."""
        src = self.dir(sess["id"]) / "files" / rec["name"]
        if Path(rec["name"]).suffix.lower() not in OFFICE:
            return src, None
        out = self.dir(sess["id"]) / "previews" / (rec["name"] + ".pdf")
        with self.lock:
            lock = self.previews.setdefault(str(out), threading.Lock())
        with lock:
            if out.exists():
                return out, None
            ext = Path(rec["name"]).suffix.lower()
            r = subprocess.run(["docker", "run", "--rm", "-i", "--network", "docgen", "--entrypoint", "curl",
                                self.image, "-sS", "--fail-with-body", "--max-time", "150",
                                "-F", f"files=@-;filename=document{ext}",
                                f"{GOTENBERG}/forms/libreoffice/convert", "-o", "-"],
                               input=src.read_bytes(), capture_output=True, timeout=200)
            if r.returncode != 0 or not r.stdout.startswith(b"%PDF"):
                msg = (r.stdout[:300] if not r.stdout.startswith(b"%PDF") else b"") + r.stderr[-300:]
                return None, "Önizleme oluşturulamadı (Gotenberg): " + msg.decode(errors="replace").strip()
            out.parent.mkdir(exist_ok=True)
            out.write_bytes(r.stdout)
            return out, None

    def sheet(self, sess, rec):
        """-> (json bytes, None) | (None, error). Files never change, so the model is built once."""
        key = (sess["id"], rec["name"])
        if key not in self.sheets:
            try:
                model = xlsx_view.load(self.dir(sess["id"]) / "files" / rec["name"])
            except xlsx_view.ViewError as e:
                return None, str(e)
            except (OSError, ValueError, KeyError, IndexError, xlsx_view.ET.ParseError) as e:
                return None, f"Excel dosyası okunamadı: {type(e).__name__}: {e}"
            self.sheets[key] = json.dumps(model, ensure_ascii=False).encode()
        return self.sheets[key], None


# ---- health -----------------------------------------------------------------
class Health:
    def __init__(self, image):
        self.image, self.state, self.lock = image, {"checking": True}, threading.Lock()

    def check(self, start_gotenberg=False):
        with self.lock:
            s = {"checking": False, "env": (ROOT / ".env").exists(), "model": model_name(), "image_name": self.image}
            s["docker"] = sh(["docker", "version", "--format", "{{.Server.Version}}"]).returncode == 0
            s["image"] = s["docker"] and sh(["docker", "image", "inspect", self.image]).returncode == 0
            if s["image"] and start_gotenberg:
                sh(["docker", "compose", "up", "-d", "gotenberg"], cwd=ROOT)
            s["streaming"] = s["image"] and sh(["docker", "run", "--rm", "--entrypoint", "grep", self.image, "-q",
                                                "DOCGEN_OUT_FILE", "/opt/docgen/entrypoint.sh"]).returncode == 0
            s["gotenberg"] = False
            for _ in range(15 if start_gotenberg else 1):
                if s["image"] and sh(["docker", "run", "--rm", "--network", "docgen", "--entrypoint", "curl", self.image,
                                      "-sf", "--max-time", "5", f"{GOTENBERG}/health"]).returncode == 0:
                    s["gotenberg"] = True
                    break
                time.sleep(2)
            self.state = s
            return s


# ---- HTTP -------------------------------------------------------------------
class Handler(BaseHTTPRequestHandler):
    store: Store = None
    health: Health = None
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):
        if args and str(args[1]).startswith(("4", "5")):
            sys.stderr.write("%s %s\n" % (self.address_string(), fmt % args))

    def send(self, code, body=b"", ctype="application/json; charset=utf-8", headers=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (headers or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def fail(self, code, msg):
        self.send(code, {"error": msg})

    def body_json(self):
        n = int(self.headers.get("Content-Length") or 0)
        try:
            return json.loads(self.rfile.read(n) or b"{}")
        except ValueError:
            return {}

    def send_file(self, path, name, download):
        data = path.read_bytes()
        ctype = mimetypes.guess_type(name)[0] or "application/octet-stream"
        if ctype.startswith("text/") or ctype == "application/json":
            ctype += "; charset=utf-8"
        disp = "attachment" if download else "inline"
        self.send(200, data, ctype, {"Content-Disposition": f"{disp}; filename*=UTF-8''{quote(name)}"})

    def route(self):
        url = urlparse(self.path)
        parts = [unquote(p) for p in url.path.strip("/").split("/") if p]
        return parts, parse_qs(url.query)

    def session(self, sid):
        sess = self.store.get(sid) if re.fullmatch(r"[0-9a-f]{12}", sid or "") else None
        if not sess:
            self.fail(404, "Sohbet bulunamadı.")
        return sess

    def do_GET(self):
        parts, q = self.route()
        if not parts or parts == ["index.html"]:
            return self.send_file(STATIC / "index.html", "index.html", False)
        if parts[0] == "static" and len(parts) == 2 and parts[1] in {p.name for p in STATIC.iterdir()}:
            return self.send_file(STATIC / parts[1], parts[1], False)
        if parts[:2] == ["api", "health"]:
            if "refresh" in q:
                self.health.check()
            return self.send(200, {**self.health.state, "timeout": self.store.timeout,
                                   "uploadable": sorted(UPLOADABLE), "office": sorted(OFFICE), "sheets": sorted(SHEETS)})
        if parts[:2] == ["api", "sessions"] and len(parts) == 2:
            items = sorted(self.store.sessions.values(), key=lambda s: -s["updated"])
            return self.send(200, [self.store.summary(s) for s in items if s["turns"] or s["files"]])
        if parts[:2] != ["api", "sessions"] or len(parts) < 3:
            return self.fail(404, "Bulunamadı.")
        sess = self.session(parts[2])
        if not sess:
            return
        if len(parts) == 3:
            return self.send(200, self.store.view(sess))
        if parts[3] == "turns" and len(parts) == 5 and parts[4].isdigit():
            turn = next((t for t in sess["turns"] if t["n"] == int(parts[4])), None)
            return self.send(200, self.store.turn_view(sess["id"], turn)) if turn else self.fail(404, "Mesaj yok.")
        if parts[3] == "files" and len(parts) in (5, 6):
            rec = self.store.file(sess, parts[4])
            if not rec:
                return self.fail(404, "Dosya yok.")
            if len(parts) == 6 and parts[5] == "sheet":
                if Path(rec["name"]).suffix.lower() not in SHEETS:
                    return self.fail(415, "Bu dosya tablo olarak gösterilemez.")
                body, err = self.store.sheet(sess, rec)
                return self.fail(422, err) if err else self.send(200, body)
            if len(parts) == 6 and parts[5] == "preview":
                path, err = self.store.preview(sess, rec)
                if err:
                    return self.fail(502, err)
                return self.send_file(path, path.name, False)
            return self.send_file(self.store.dir(sess["id"]) / "files" / rec["name"], rec["name"], "download" in q)
        return self.fail(404, "Bulunamadı.")

    def do_POST(self):
        parts, _ = self.route()
        if parts == ["api", "sessions"]:
            return self.send(201, self.store.view(self.store.create()))
        if parts[:2] != ["api", "sessions"] or len(parts) < 4:
            return self.fail(404, "Bulunamadı.")
        sess = self.session(parts[2])
        if not sess:
            return
        if parts[3] == "turns" and len(parts) == 4:
            b = self.body_json()
            text = str(b.get("text") or "").strip()
            if not text:
                return self.fail(400, "Mesaj boş.")
            if not self.health.state.get("image"):
                return self.fail(503, f"{self.store.image} imajı yok: önce `docker compose build` çalıştırın.")
            if not (ROOT / ".env").exists():
                return self.fail(503, ".env yok: `cp .env.example .env` ile model ayarlarını yazın.")
            turn = self.store.start_turn(sess, text, b.get("active"), b.get("confirm"))
            if not turn:
                return self.fail(409, "Bu sohbette çalışan bir istek var.")
            return self.send(201, self.store.turn_view(sess["id"], turn))
        if parts[3] == "turns" and len(parts) == 6 and parts[5] == "cancel" and parts[4].isdigit():
            ok = self.store.cancel(sess, int(parts[4]))
            return self.send(200, {"ok": True}) if ok else self.fail(409, "Bu mesaj çalışmıyor.")
        return self.fail(404, "Bulunamadı.")

    def do_PUT(self):
        parts, q = self.route()
        if parts[:2] != ["api", "sessions"] or len(parts) != 4 or parts[3] != "files":
            return self.fail(404, "Bulunamadı.")
        sess = self.session(parts[2])
        if not sess:
            return
        name, length = (q.get("name") or [""])[0], int(self.headers.get("Content-Length") or 0)
        ext = Path(clean_name(name)).suffix.lower()
        if ext not in UPLOADABLE:
            self.close_connection = True
            return self.fail(415, f"Desteklenmeyen dosya türü: {ext or '(uzantısız)'}")
        if not 0 < length <= MAX_UPLOAD:
            self.close_connection = True
            return self.fail(413, "Dosya boş ya da çok büyük (en fazla 200 MB).")
        try:
            rec = self.store.add_file(sess, name, self.rfile, length)
        except OSError as e:
            self.close_connection = True
            return self.fail(400, f"Yükleme başarısız: {e}")
        return self.send(201, rec)

    def do_DELETE(self):
        parts, _ = self.route()
        if parts[:2] != ["api", "sessions"] or len(parts) < 3:
            return self.fail(404, "Bulunamadı.")
        sess = self.session(parts[2])
        if not sess:
            return
        if len(parts) == 3:
            return self.send(200, {"ok": True}) if self.store.delete(sess["id"]) else \
                self.fail(409, "Çalışan bir istek varken sohbet silinemez.")
        if parts[3] == "files" and len(parts) == 5:
            return self.send(200, {"ok": True}) if self.store.remove_pending(sess, parts[4]) else \
                self.fail(409, "Yalnızca henüz gönderilmemiş yüklemeler kaldırılabilir.")
        return self.fail(404, "Bulunamadı.")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=8765)
    ap.add_argument("--data", default=str(ROOT / "data" / "ui"), help="sessions directory")
    ap.add_argument("--image", default="docgen-agent:latest")
    ap.add_argument("--timeout", type=int, default=1500, help="seconds per message (default 1500)")
    a = ap.parse_args()

    Handler.store = Store(Path(a.data).resolve(), a.image, a.timeout)
    Handler.health = Health(a.image)
    threading.Thread(target=Handler.health.check, kwargs={"start_gotenberg": True}, daemon=True).start()
    server = ThreadingHTTPServer((a.host, a.port), Handler)
    server.daemon_threads = True
    print(f"DOCGEN arayüzü: http://{a.host}:{a.port}   (oturumlar: {a.data})", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
