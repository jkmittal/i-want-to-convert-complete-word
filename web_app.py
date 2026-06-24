import os
import re
import uuid
from pathlib import Path

from docx import Document
from flask import Flask, abort, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from word_topic_extractor import content_table_entries, extract_topics


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "work" / "web_uploads"
OUTPUT_DIR = BASE_DIR / "outputs" / "web_exports"
ALLOWED_EXTENSIONS = {".docx"}


app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "change-this-secret-key-before-production")
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_UPLOAD_MB", "25")) * 1024 * 1024

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def allowed_file(filename):
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def stored_upload_path(file_id):
    if not re.fullmatch(r"[a-f0-9]{32}", file_id or ""):
        abort(404)
    return UPLOAD_DIR / f"{file_id}.docx"


def safe_output_name(topics):
    first_topic = topics[0] if topics else "selected_topics"
    name = re.sub(r"[^A-Za-z0-9._-]+", "_", first_topic.strip()).strip("_")
    if len(topics) > 1:
        name = f"{name}_and_{len(topics) - 1}_more"
    return f"{name or 'selected_topics'}.docx"


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/upload")
def upload():
    uploaded_file = request.files.get("word_file")
    if not uploaded_file or not uploaded_file.filename:
        flash("Please choose a Word .docx file.")
        return redirect(url_for("index"))

    if not allowed_file(uploaded_file.filename):
        flash("Only .docx Word files are supported.")
        return redirect(url_for("index"))

    file_id = uuid.uuid4().hex
    original_name = secure_filename(uploaded_file.filename)
    upload_path = stored_upload_path(file_id)
    uploaded_file.save(upload_path)

    try:
        entries = content_table_entries(Document(upload_path))
    except Exception as error:
        upload_path.unlink(missing_ok=True)
        flash(f"Could not read this Word file: {error}")
        return redirect(url_for("index"))

    if not entries:
        upload_path.unlink(missing_ok=True)
        flash("No contents table or heading topics were found in this document.")
        return redirect(url_for("index"))

    return render_template(
        "select_topics.html",
        file_id=file_id,
        original_name=original_name,
        entries=entries,
    )


@app.post("/extract")
def extract():
    file_id = request.form.get("file_id", "")
    upload_path = stored_upload_path(file_id)
    if not upload_path.exists():
        flash("The uploaded file is no longer available. Please upload it again.")
        return redirect(url_for("index"))

    topics = request.form.getlist("topics")
    if not topics:
        flash("Please select at least one topic.")
        return redirect(url_for("index"))

    output_name = safe_output_name(topics)
    output_path = OUTPUT_DIR / f"{uuid.uuid4().hex}_{output_name}"

    try:
        extract_topics(upload_path, topics, output_path)
    except Exception as error:
        flash(f"Could not create the selected Word file: {error}")
        return redirect(url_for("index"))

    return send_file(
        output_path,
        as_attachment=True,
        download_name=output_name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.errorhandler(413)
def too_large(_error):
    flash("The uploaded file is too large. Try a smaller Word file or increase MAX_UPLOAD_MB.")
    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")), debug=True)
