"""Flask frontend for the aged-care document redactor."""

from __future__ import annotations

import os
import uuid
from pathlib import Path

from anthropic import Anthropic
from dotenv import load_dotenv
from flask import Flask, abort, jsonify, render_template, request, send_from_directory
from werkzeug.utils import secure_filename

from process_docs import process_file

load_dotenv()

BASE_OUT = Path(os.environ.get("OUTPUT_DIR", "out")).resolve()
BASE_OUT.mkdir(parents=True, exist_ok=True)
MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "25"))
ALLOWED_EXTS = {".doc", ".docx"}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024

_client: Anthropic | None = None


def get_client() -> Anthropic:
    global _client
    if _client is None:
        _client = Anthropic()
    return _client


@app.route("/")
def index():
    return render_template("index.html", max_mb=MAX_UPLOAD_MB)


@app.route("/process", methods=["POST"])
def process():
    f = request.files.get("file")
    if f is None or not f.filename:
        return jsonify({"error": "no file uploaded"}), 400

    filename = secure_filename(f.filename)
    if not filename:
        return jsonify({"error": "invalid filename"}), 400
    ext = Path(filename).suffix.lower()
    if ext not in ALLOWED_EXTS:
        return jsonify({"error": f"unsupported file type: {ext}"}), 400

    job_id = uuid.uuid4().hex[:12]
    job_dir = BASE_OUT / job_id
    job_dir.mkdir(parents=True, exist_ok=True)
    upload_path = job_dir / filename
    f.save(str(upload_path))

    try:
        report = process_file(upload_path, job_dir, get_client())
    except Exception as e:
        app.logger.exception("processing failed for %s", filename)
        return jsonify({"error": f"processing failed: {e}"}), 500

    stem = Path(filename).stem
    redacted_rel = f"{stem}.redacted.docx"
    report_rel = f"{stem}.report.json"

    logos_web = []
    for lg in report["logos"]:
        rel = lg.get("file")
        logos_web.append(
            {
                **lg,
                "url": f"/files/{job_id}/{rel}" if rel else None,
            }
        )

    return jsonify(
        {
            "job_id": job_id,
            "source": filename,
            "entities": report["entities"],
            "logos": logos_web,
            "redacted_url": f"/files/{job_id}/{redacted_rel}?download=1",
            "report_url": f"/files/{job_id}/{report_rel}?download=1",
        }
    )


@app.route("/files/<job_id>/<path:filename>")
def serve_file(job_id: str, filename: str):
    # secure_filename strips slashes; we need to allow nested paths but block traversal.
    if ".." in filename.split("/") or job_id != secure_filename(job_id):
        abort(404)
    job_dir = BASE_OUT / job_id
    if not job_dir.is_dir():
        abort(404)
    as_attachment = request.args.get("download") == "1"
    try:
        return send_from_directory(job_dir, filename, as_attachment=as_attachment)
    except FileNotFoundError:
        abort(404)


@app.errorhandler(413)
def too_large(_):
    return jsonify({"error": f"file exceeds {MAX_UPLOAD_MB} MB limit"}), 413


if __name__ == "__main__":
    app.run(
        host=os.environ.get("HOST", "127.0.0.1"),
        port=int(os.environ.get("PORT", "5000")),
        debug=os.environ.get("FLASK_DEBUG") == "1",
    )
