from flask import Flask, render_template, request, send_file
import os
import tempfile
import zipfile
import time
from werkzeug.utils import secure_filename
from docx import Document
import pdfplumber
from docx2pdf import convert as docx2pdf_convert

app = Flask(__name__)

def convert_docx_to_pdf(input_path):
    """
    Converts a DOCX file to PDF reliably.
    Returns the path to the PDF file.
    """
    output_dir = os.path.dirname(input_path)
    # Convert the DOCX; docx2pdf writes to the same directory
    docx2pdf_convert(input_path, output_dir)

    pdf_path = os.path.join(output_dir, os.path.basename(input_path).replace(".docx", ".pdf"))

    # Wait until PDF exists (max 5 seconds)
    timeout = 5
    while not os.path.exists(pdf_path) and timeout > 0:
        time.sleep(0.5)
        timeout -= 0.5

    if not os.path.exists(pdf_path):
        raise Exception(f"PDF conversion failed for {input_path}")

    return pdf_path

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/convert", methods=["POST"])
def convert_files():
    if "files[]" not in request.files:
        return "No files uploaded", 400

    files = request.files.getlist("files[]")
    converted_files = []

    for file in files:
        if file.filename == "":
            continue

        filename = secure_filename(file.filename)
        ext = filename.rsplit(".", 1)[1].lower()
        temp_input = os.path.join(tempfile.gettempdir(), filename)
        file.save(temp_input)

        if ext == "docx":
            # Word → PDF
            try:
                pdf_path = convert_docx_to_pdf(temp_input)
                converted_files.append(pdf_path)
            except Exception as e:
                print(f"Error converting {filename} to PDF:", e)

        elif ext == "pdf":
            # PDF → Word
            temp_docx = os.path.join(tempfile.gettempdir(), filename.replace(".pdf", ".docx"))
            doc = Document()
            try:
                with pdfplumber.open(temp_input) as pdf:
                    for page in pdf.pages:
                        text = page.extract_text()
                        if text:
                            doc.add_paragraph(text)
                doc.save(temp_docx)
                converted_files.append(temp_docx)
            except Exception as e:
                print(f"Error converting {filename} to DOCX:", e)

    # If only one file, send directly
    if len(converted_files) == 1:
        return send_file(converted_files[0], as_attachment=True)

    # Multiple files → ZIP
    zip_path = os.path.join(tempfile.gettempdir(), "converted_files.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for f in converted_files:
            if os.path.exists(f) and os.path.getsize(f) > 0:
                zipf.write(f, arcname=os.path.basename(f))

    return send_file(zip_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
