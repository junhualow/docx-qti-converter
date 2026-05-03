import os
import uuid
import shutil
from flask import Flask, render_template, request, send_file, after_this_request, jsonify
from werkzeug.utils import secure_filename
from converter import parse_docx_to_data, generate_qti_from_data
from PIL import Image

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

    # Handle optional PDF companion file
    pdf_path = None
    if "pdf_file" in request.files:
        pdf_file = request.files["pdf_file"]
        if pdf_file.filename and pdf_file.filename.lower().endswith(".pdf"):
            pdf_safe_name = secure_filename(pdf_file.filename)
            pdf_path = os.path.join(job_dir, pdf_safe_name)
            pdf_file.save(pdf_path)

    try:
        data = parse_docx_to_data(input_path, job_dir, pdf_path=pdf_path)
        data['job_id'] = job_id
        return jsonify(data)
    except Exception as e:
        shutil.rmtree(job_dir, ignore_errors=True)
        return {"error": str(e)}, 500

@app.route("/crop", methods=["POST"])
def crop_image():
    """Crop an extracted image based on user-defined coordinates."""
    data = request.json
    if not data:
        return {"error": "No data"}, 400

    job_id = data.get("job_id")
    filename = data.get("filename")
    # Crop coordinates as fractions (0.0 to 1.0) of the original image dimensions
    crop_top = data.get("top", 0)
    crop_left = data.get("left", 0)
    crop_right = data.get("right", 1)
    crop_bottom = data.get("bottom", 1)

    if not job_id or not filename:
        return {"error": "Missing job_id or filename"}, 400

    # Sanitize filename to prevent directory traversal
    filename = os.path.basename(filename)
    img_path = os.path.join(STATIC_JOBS, job_id, "assets", filename)

    if not os.path.exists(img_path):
        return {"error": "Image not found"}, 404

    try:
        img = Image.open(img_path)
        w, h = img.size

        # Convert fractional coordinates to pixel values
        left = int(crop_left * w)
        top = int(crop_top * h)
        right = int(crop_right * w)
        bottom = int(crop_bottom * h)

        # Ensure valid crop region
        left = max(0, min(left, w - 1))
        top = max(0, min(top, h - 1))
        right = max(left + 1, min(right, w))
        bottom = max(top + 1, min(bottom, h))

        cropped = img.crop((left, top, right, bottom))

        # Save as new file with _cropped suffix
        base, ext = os.path.splitext(filename)
        new_filename = f"{base}_cropped{ext}"
        new_path = os.path.join(STATIC_JOBS, job_id, "assets", new_filename)
        cropped.save(new_path)

        return jsonify({"new_filename": new_filename})
    except Exception as e:
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
