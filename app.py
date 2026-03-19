import os
import uuid
import shutil
from flask import Flask, render_template, request, send_file, after_this_request
from werkzeug.utils import secure_filename
from converter import convert_docx_to_qti

app = Flask(__name__)

MAX_SIZE = 10 * 1024 * 1024
app.config['MAX_CONTENT_LENGTH'] = MAX_SIZE

BASE_UPLOAD = "jobs"
os.makedirs(BASE_UPLOAD, exist_ok=True)

def allowed_file(filename):
    return filename.lower().endswith(".docx")

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/convert", methods=["POST"])
def convert():
    if "file" not in request.files:
        return {"error": "No file part in request"}, 400

    file = request.files["file"]

    if file.filename == "":
        return {"error": "No file selected"}, 400

    if not allowed_file(file.filename):
        return {"error": "Only .docx files are allowed"}, 400

    job_id = str(uuid.uuid4())
    job_dir = os.path.join(BASE_UPLOAD, job_id)
    os.makedirs(job_dir)

    safe_name = secure_filename(file.filename)
    input_path = os.path.join(job_dir, safe_name)
    file.save(input_path)

    if os.path.getsize(input_path) > MAX_SIZE:
        shutil.rmtree(job_dir, ignore_errors=True)
        return {"error": "File exceeds 10MB limit"}, 400

    try:
        zip_path = convert_docx_to_qti(input_path, job_dir)
    except Exception as e:
        shutil.rmtree(job_dir, ignore_errors=True)
        return {"error": str(e)}, 500

    @after_this_request
    def cleanup(response):
        try:
            response.call_on_close(lambda: shutil.rmtree(job_dir, ignore_errors=True))
        except Exception:
            shutil.rmtree(job_dir, ignore_errors=True)
        return response

    return send_file(
        zip_path,
        as_attachment=True,
        download_name="converted_qti.zip",
        mimetype="application/zip"
    )