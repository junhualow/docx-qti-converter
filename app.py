import os
import uuid
from flask import Flask, render_template, request, send_file
from converter import convert_docx_to_qti

app = Flask(__name__)

JOBS_FOLDER = "jobs"
os.makedirs(JOBS_FOLDER, exist_ok=True)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert():

    file = request.files["file"]

    if not file.filename.endswith(".docx"):
        return "Please upload a DOCX file"

    # create a unique folder for this job
    job_id = str(uuid.uuid4())[:8]
    job_folder = os.path.join(JOBS_FOLDER, job_id)
    os.makedirs(job_folder)

    # save uploaded file
    path = os.path.join(job_folder, file.filename)
    file.save(path)

    # run converter
    zip_file = convert_docx_to_qti(path)

    return send_file(zip_file, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)