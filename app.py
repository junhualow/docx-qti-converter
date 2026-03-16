from flask import Flask, render_template, request, send_file
import os
from converter import convert_docx_to_qti

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/convert", methods=["POST"])
def convert():

    file = request.files["file"]

    if not file.filename.endswith(".docx"):
        return "Please upload a DOCX file"

    path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(path)

    zip_file = convert_docx_to_qti(path)

    return send_file(zip_file, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)