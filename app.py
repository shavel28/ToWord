"""
ToWord — web PDF ke Word gratis, tanpa watermark.

Jalankan lokal:
    pip install -r requirements.txt
    python app.py            # http://localhost:8000
"""

from __future__ import annotations

import logging
import os
import re
import shutil
import tempfile
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor

from flask import Flask, abort, jsonify, render_template, request, send_file

from converter import ConversionError, PasswordRequired, convert_pdf_to_docx

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("toword.web")

MAX_MB = int(os.environ.get("TOWORD_MAX_MB", "100"))
KEEP_MINUTES = int(os.environ.get("TOWORD_KEEP_MINUTES", "30"))
WORKERS = int(os.environ.get("TOWORD_WORKERS", str(max(1, (os.cpu_count() or 2) // 2))))
DATA_DIR = os.environ.get("TOWORD_DATA_DIR") or os.path.join(tempfile.gettempdir(), "toword_jobs")
os.makedirs(DATA_DIR, exist_ok=True)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__, template_folder=BASE_DIR, static_folder=None)
app.config["MAX_CONTENT_LENGTH"] = MAX_MB * 1024 * 1024

executor = ThreadPoolExecutor(max_workers=WORKERS)
jobs: dict[str, dict] = {}
jobs_lock = threading.Lock()


def _safe_name(name: str) -> str:
    base = os.path.splitext(os.path.basename(name or "dokumen"))[0]
    base = re.sub(r"[^\w\-. ]+", "_", base, flags=re.UNICODE).strip(" ._") or "dokumen"
    return base[:120] + ".docx"


def _update(job_id, **kw):
    with jobs_lock:
        if job_id in jobs:
            jobs[job_id].update(kw)


def _run_job(job_id: str):
    job = jobs[job_id]
    _update(job_id, status="processing", progress=0.01, message="Memulai…")

    def progress(frac, msg):
        _update(job_id, progress=round(frac, 3), message=msg)

    try:
        info = convert_pdf_to_docx(job["pdf"], job["docx"], mode=job["mode"],
                                   password=job.get("password", ""), progress=progress)
        _update(job_id, status="done", progress=1.0, message="Selesai", info=info)
    except PasswordRequired as ex:
        _update(job_id, status="password", message=str(ex))
    except ConversionError as ex:
        _update(job_id, status="error", message=str(ex))
    except Exception:
        log.exception("Konversi gagal")
        _update(job_id, status="error",
                message="Terjadi kesalahan saat mengonversi. Coba mode 'Scan / Foto (OCR)'.")
    finally:
        job.pop("password", None)
        # File PDF asli langsung dihapus setelah diproses
        try:
            os.remove(job["pdf"])
        except OSError:
            pass


def _cleanup_loop():
    while True:
        time.sleep(60)
        limit = time.time() - KEEP_MINUTES * 60
        with jobs_lock:
            old = [k for k, v in jobs.items() if v["created"] < limit]
            for k in old:
                d = jobs.pop(k)["dir"]
                shutil.rmtree(d, ignore_errors=True)


threading.Thread(target=_cleanup_loop, daemon=True).start()


@app.get("/")
def index():
    return render_template("index.html", max_mb=MAX_MB, keep_minutes=KEEP_MINUTES)


@app.post("/api/convert")
def api_convert():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify(error="Pilih file PDF terlebih dahulu."), 400
    head = f.stream.read(1024)
    f.stream.seek(0)
    if b"%PDF" not in head:
        return jsonify(error="File ini bukan PDF."), 400

    mode = request.form.get("mode", "auto")
    if mode not in ("auto", "ocr", "layout"):
        mode = "auto"

    job_id = uuid.uuid4().hex
    d = os.path.join(DATA_DIR, job_id)
    os.makedirs(d)
    pdf_path = os.path.join(d, "input.pdf")
    f.save(pdf_path)

    job = {
        "id": job_id, "dir": d, "pdf": pdf_path, "docx": os.path.join(d, "hasil.docx"),
        "name": _safe_name(f.filename), "mode": mode,
        "password": request.form.get("password", ""),
        "status": "queued", "progress": 0.0, "message": "Dalam antrean…",
        "created": time.time(),
    }
    with jobs_lock:
        jobs[job_id] = job
    executor.submit(_run_job, job_id)
    return jsonify(id=job_id)


@app.get("/api/status/<job_id>")
def api_status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify(status="missing", message="Pekerjaan tidak ditemukan atau sudah kedaluwarsa."), 404
    return jsonify({k: job.get(k) for k in ("status", "progress", "message", "name", "info")})


@app.get("/api/download/<job_id>")
def api_download(job_id):
    job = jobs.get(job_id)
    if not job or job.get("status") != "done" or not os.path.exists(job["docx"]):
        abort(404)
    return send_file(job["docx"], as_attachment=True, download_name=job["name"],
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.delete("/api/job/<job_id>")
def api_delete(job_id):
    with jobs_lock:
        job = jobs.pop(job_id, None)
    if job:
        shutil.rmtree(job["dir"], ignore_errors=True)
    return jsonify(ok=True)


@app.errorhandler(413)
def too_large(_):
    return jsonify(error=f"Ukuran file maksimal {MAX_MB} MB."), 413


@app.get("/healthz")
def health():
    return "ok"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")), threaded=True)
