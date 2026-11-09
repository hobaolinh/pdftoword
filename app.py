"""Flask service with a loopback listener. Conversion results live in RAM."""

from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from functools import lru_cache
from io import BytesIO
from pathlib import PurePath
from threading import RLock, Thread
from time import monotonic, sleep
import argparse
import logging
import os
import re
import secrets
from urllib.parse import unquote

from flask import Flask, abort, jsonify, render_template, request, send_file
import pymupdf as pdf
from waitress import create_server

from layout import ConversionError, encoded, extract_page, ocr_command, raster
from word import export_word

MAX_FILE_MB = 100
LIMIT = MAX_FILE_MB * 1024 * 1024
MAX_PAGES = 900
MAX_SESSIONS = 10
DEFAULT_PORT = 80
RESULT_LIMIT = 512 * 1024 * 1024
TTL = 15 * 60
SERVER_OPTIONS = {
    "host": "127.0.0.1",
    "threads": 4,
    "inbuf_overflow": LIMIT + 1,
    "outbuf_overflow": RESULT_LIMIT + 1,
    "max_request_body_size": LIMIT + 1024 * 1024,
}
LANGUAGES = {"vie", "eng", "vie+eng"}
LOOPBACK = re.compile(r"^(localhost|127\.0\.0\.1|\[::1\])(?::\d{1,5})?$", re.I)

app = Flask(__name__)
pdf.no_recommend_layout()
app.config["MAX_CONTENT_LENGTH"] = LIMIT
pool = ThreadPoolExecutor(max_workers=1, thread_name_prefix="pdf")
lock = RLock()
jobs = {}


@dataclass
class Job:
    name: str
    language: str
    state: str = "queued"
    current: int = 0
    complete: int = 0
    total: int = 0
    error: str = ""
    expiry: float = field(default_factory=lambda: monotonic() + TTL)
    models: list = field(default_factory=list)
    originals: list = field(default_factory=list)
    media: list = field(default_factory=list)
    output: bytes = b""


def sweep():
    with lock:
        for key, job in list(jobs.items()):
            if job.expiry <= monotonic() and job.state not in {"queued", "working"}:
                jobs.pop(key)


def janitor():
    while True:
        sleep(15)
        sweep()


Thread(target=janitor, daemon=True, name="expiration").start()


def convert(key, job, content):
    with lock:
        job.state = "working"
    models, originals = [], []
    preview_bytes = 0

    @lru_cache(maxsize=1)
    def get_ocr():
        return ocr_command(job.language)

    try:
        try:
            document = pdf.open(stream=content, filetype="pdf")
        except (ValueError, RuntimeError) as exc:
            raise ConversionError("PDF bị hỏng hoặc có định dạng không hợp lệ.") from exc
        with document:
            if document.needs_pass:
                raise ConversionError("PDF có mật khẩu. Hãy mở khóa bằng công cụ cục bộ rồi thử lại.")
            total = len(document)
            if not 1 <= total <= MAX_PAGES:
                raise ConversionError(f"PDF phải có từ 1 đến {MAX_PAGES} trang. Hãy chia nhỏ tệp.")
            with lock:
                job.total = total
            preview_size = 820 if total > 100 else 1200
            for i, page in enumerate(document):
                if page.rect.is_empty or page.rect.is_infinite:
                    raise ConversionError(f"Trang {i+1} có kích thước không hợp lệ.")
                with lock:
                    job.current = i + 1
                image, _, _ = raster(page, page.rect, min(2, preview_size / max(page.rect.width, page.rect.height)))
                originals.append(encoded(image, "JPEG", quality=83, optimize=True))
                model = extract_page(page, get_ocr, job.language)
                # Character boxes are used during extraction, but are not needed by Word or the preview.
                for block in model["items"]:
                    lines = block.get("lines", [])
                    if block["kind"] == "table":
                        lines = [line for cell in block["cells"] for line in cell["lines"]]
                    for line in lines:
                        for span in line["spans"]:
                            span.pop("chars", None)
                models.append(model)
                preview_bytes += len(originals[-1]) + sum(map(len, model["media"].values()))
                if preview_bytes > RESULT_LIMIT:
                    raise ConversionError("Kết quả quá lớn. Hãy chia nhỏ tệp PDF.")
                with lock:
                    job.complete = i + 1
        result = export_word(models)
        if len(result) + preview_bytes > RESULT_LIMIT:
            raise ConversionError("Kết quả quá lớn. Hãy chia nhỏ tệp PDF.")
        media = [model.pop("media") for model in models]
        with lock:
            job.models, job.originals, job.media, job.output = models, originals, media, result
            job.state, job.expiry = "ready", monotonic() + TTL
    except Exception as exc:
        message = str(exc) if isinstance(exc, ConversionError) else (
            "Không thể đọc hoặc chuyển đổi PDF. Hãy thử lưu lại tệp PDF rồi tải lên.")
        if not isinstance(exc, ConversionError):
            logging.error("Conversion failed (%s)", type(exc).__name__)
        with lock:
            job.state, job.error, job.expiry = "failed", message, monotonic() + 120


def find_job(key):
    job = jobs.get(key)
    if not job:
        abort(404, "Phiên đã hết hạn hoặc tệp Word đã được tải xuống.")
    return job


@app.before_request
def local_only():
    forwarded_host = app.config.get("CODESPACES_HOST")
    local_host = LOOPBACK.fullmatch(request.host)
    if not local_host and request.host != forwarded_host:
        abort(403)
    origin = request.headers.get("Origin")
    if origin:
        allowed_origins = {f"http://{request.host}"} if local_host else set()
        if forwarded_host:
            allowed_origins.add(f"https://{forwarded_host}")
        if origin not in allowed_origins:
            abort(403)
    if request.method == "POST" and request.headers.get("X-Local-Converter") != "1":
        abort(403)
    sweep()


@app.after_request
def local_headers(response):
    response.headers["Cache-Control"] = "no-store"
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["Referrer-Policy"] = "no-referrer"
    response.headers["Cross-Origin-Resource-Policy"] = "same-origin"
    response.headers["Content-Security-Policy"] = (
        "default-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' blob:; "
        "object-src 'none'; base-uri 'none'; frame-ancestors 'none'; form-action 'self'")
    return response


@app.errorhandler(413)
def too_large(_):
    return jsonify(error=f"PDF vượt giới hạn {MAX_FILE_MB} MB. Hãy nén hoặc chia nhỏ tệp."), 413


@app.errorhandler(404)
def missing(error):
    return jsonify(error=error.description), 404


@app.get("/")
def index():
    return render_template("index.html", max_file_mb=MAX_FILE_MB, max_pages=MAX_PAGES)


@app.post("/api/jobs")
def submit():
    if request.mimetype != "application/pdf":
        return jsonify(error="Hãy chọn một tệp PDF."), 415
    lang = request.args.get("lang", "")
    if lang not in LANGUAGES:
        return jsonify(error="Hãy chọn ngôn ngữ OCR hợp lệ."), 400
    name = PurePath(unquote(request.headers.get("X-PDF-Name", "tai-lieu.pdf")).replace("\\", "/")).name
    name = re.sub(r"[\x00-\x1f\x7f]", "", name)[:220]
    if not name.lower().endswith(".pdf"):
        return jsonify(error="Tên tệp phải có đuôi .pdf."), 400
    with lock:
        if sum(j.state in {"working", "queued", "ready"} for j in jobs.values()) >= MAX_SESSIONS:
            return jsonify(error=f"Đã có {MAX_SESSIONS} phiên đang hoạt động. Hãy tải kết quả hoặc đợi phiên cũ hết hạn."), 429
        data = request.get_data(cache=False)
        if b"%PDF-" not in data[:1024]:
            return jsonify(error="Tệp trống, bị hỏng hoặc không có cấu trúc PDF."), 400
        key = secrets.token_urlsafe(24)
        job = Job(name, lang)
        jobs[key] = job
        pool.submit(convert, key, job, data)
    return jsonify(id=key), 202


@app.get("/api/jobs/<key>")
def status(key):
    with lock:
        job = find_job(key)
        queue_position = None
        if job.state == "queued":
            # Jobs are inserted under the same lock as submission to the FIFO worker.
            waiting = [job_id for job_id, item in jobs.items() if item.state == "queued"]
            queue_position = waiting.index(key) + 1
        return jsonify(state=job.state, current=job.current, complete=job.complete, total=job.total,
                       queue_position=queue_position,
                       expires_in_ms=max(0, round((job.expiry - monotonic()) * 1000))
                       if job.state == "ready" else None,
                       error=job.error, pages=[{"mode": p["mode"], "warnings": len(p["warnings"])}
                                               for p in job.models])


@app.get("/api/jobs/<key>/pages/<int:index>")
def preview(key, index):
    with lock:
        job = find_job(key)
        if index >= len(job.models):
            abort(404, "Chưa có bản xem trước của trang này.")
        return jsonify(job.models[index])


@app.get("/api/jobs/<key>/pages/<int:index>/original")
def original(key, index):
    with lock:
        job = find_job(key)
        if index >= len(job.originals):
            abort(404, "Không tìm thấy trang PDF.")
        data = job.originals[index]
    return send_file(BytesIO(data), mimetype="image/jpeg")


@app.get("/api/jobs/<key>/pages/<int:index>/assets/<int:asset>")
def image_asset(key, index, asset):
    with lock:
        job = find_job(key)
        if index >= len(job.media) or asset not in job.media[index]:
            abort(404, "Không tìm thấy ảnh.")
        data = job.media[index][asset]
    return send_file(BytesIO(data), mimetype="image/jpeg")


@app.post("/api/jobs/<key>/download")
def download(key):
    with lock:
        job = find_job(key)
        if job.state != "ready":
            return jsonify(error="Tệp Word chưa sẵn sàng."), 409
        response = send_file(BytesIO(job.output), as_attachment=True,
                             download_name=PurePath(job.name).stem + ".docx",
                             mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        jobs.pop(key)
        return response


@app.post("/api/jobs/<key>/cancel")
def cancel(key):
    with lock:
        job = jobs.get(key)
        if job is not None:
            if job.state != "ready":
                return jsonify(error="Chỉ có thể hủy phiên từ màn hình tải tệp Word."), 409
            jobs.pop(key)
    return jsonify(cancelled=True)


def local_server(port):
    try:
        server = create_server(app, port=port, **SERVER_OPTIONS)
    except OSError as exc:
        raise SystemExit(
            f"Không mở được cổng {port} trên máy này. Kiểm tra cổng hoặc chọn cổng khác bằng --port 8766. ({exc})"
        ) from exc
    address = "http://127.0.0.1" if int(server.effective_port) == 80 else f"http://127.0.0.1:{server.effective_port}"
    return server, address + "/"


def main():
    parser = argparse.ArgumentParser(description="PDF sang Word chạy trên máy này")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    parser.add_argument("--codespaces", action="store_true",
                        help="Cho phép địa chỉ chuyển tiếp riêng của GitHub Codespaces")
    args = parser.parse_args()
    if args.codespaces:
        if not 1 <= args.port <= 65535:
            parser.error("--codespaces cần một cổng cố định trong khoảng 1–65535.")
        name = os.environ.get("CODESPACE_NAME")
        domain = os.environ.get("GITHUB_CODESPACES_PORT_FORWARDING_DOMAIN")
        if os.environ.get("CODESPACES") != "true" or not name or not domain:
            parser.error("--codespaces cần chạy trong GitHub Codespaces.")
        app.config["CODESPACES_HOST"] = f"{name}-{args.port}.{domain}"
    server, url = local_server(args.port)
    print(f"PDF sang Word: {url}", flush=True)
    server.run()


if __name__ == "__main__":
    main()
