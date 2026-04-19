import os
import uuid
import shutil
from flask import Flask, render_template, request, send_file, after_this_request, jsonify
from werkzeug.utils import secure_filename
from converter import parse_docx_to_data, generate_qti_from_data

app = Flask(__name__)

MAX_SIZE = 10 * 1024 * 1024
app.config['MAX_CONTENT_LENGTH'] = MAX_SIZE

# We use the static folder to store temporary jobs so the frontend can access extracted images
STATIC_JOBS = os.path.join(app.static_folder, "jobs") if app.static_folder else os.path.join("static", "jobs")
os.makedirs(STATIC_JOBS, exist_ok=True)

def allowed_file(filename):
    return filename.lower().endswith(".docx")

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return {"error": "No file part in request"}, 400

    file = request.files["file"]

    if file.filename == "":
        return {"error": "No file selected"}, 400

    if not allowed_file(file.filename):
        return {"error": "Only .docx files are allowed"}, 400

    job_id = str(uuid.uuid4())
    job_dir = os.path.join(STATIC_JOBS, job_id)
    os.makedirs(job_dir, exist_ok=True)

    safe_name = secure_filename(file.filename)
    input_path = os.path.join(job_dir, safe_name)
    file.save(input_path)

    if os.path.getsize(input_path) > MAX_SIZE:
        shutil.rmtree(job_dir, ignore_errors=True)
        return {"error": "File exceeds 10MB limit"}, 400

    try:
        data = parse_docx_to_data(input_path, job_dir)
        data['job_id'] = job_id
        return jsonify(data)
    except Exception as e:
        shutil.rmtree(job_dir, ignore_errors=True)
        return {"error": str(e)}, 500

@app.route("/convert", methods=["POST"])
def convert():
    data = request.json
    if not data or "job_id" not in data:
        return {"error": "Invalid data"}, 400

    job_id = data['job_id']
    job_dir = os.path.join(STATIC_JOBS, job_id)
    
    if not os.path.exists(job_dir):
        return {"error": "Job directory not found. Please upload again."}, 404

    try:
        zip_path = generate_qti_from_data(data, job_dir)
    except Exception as e:
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

if __name__ == "__main__":
    app.run(debug=True)
