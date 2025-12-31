from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
import os

from pdf2docx import Converter
from docx2pdf import convert
from PIL import Image
from pdf2image import convert_from_path
from reportlab.pdfgen import canvas
from pptx import Presentation
import openpyxl
from docx import Document

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template("index.html")

@app.route('/convert', methods=['POST'])
def convert_file():
    file = request.files['file']
    conversion = request.form['conversion']

    if file.filename == "":
        return "No file selected"

    safe_name = secure_filename(file.filename)
    filename, ext = os.path.splitext(safe_name)

    input_path = os.path.abspath(os.path.join(UPLOAD_FOLDER, safe_name))
    file.save(input_path)

    # ---------- PDF to Word ----------
    if conversion == "pdf2word":
        output_path = os.path.abspath(
            os.path.join(OUTPUT_FOLDER, filename + ".docx")
        )
        cv = Converter(input_path)
        cv.convert(output_path)
        cv.close()

    # ---------- Word to PDF ----------
    elif conversion == "word2pdf":
        output_path = os.path.abspath(
            os.path.join(OUTPUT_FOLDER, filename + ".pdf")
        )
        convert(input_path, output_path)

    # ---------- Image to PDF ----------
    elif conversion == "image2pdf":
        output_path = os.path.abspath(
            os.path.join(OUTPUT_FOLDER, filename + ".pdf")
        )
        image = Image.open(input_path)
        image.convert("RGB").save(output_path)

    # ---------- Text to PDF ----------
    elif conversion == "text2pdf":
        output_path = os.path.abspath(
            os.path.join(OUTPUT_FOLDER, filename + ".pdf")
        )
        c = canvas.Canvas(output_path)
        with open(input_path, 'r', encoding="utf-8") as f:
            text = c.beginText(40, 800)
            for line in f:
                text.textLine(line)
            c.drawText(text)
        c.save()

    # ---------- PDF to Image ----------
    elif conversion == "pdf2image":
        output_path = os.path.abspath(
            os.path.join(OUTPUT_FOLDER, filename + ".jpg")
        )
        images = convert_from_path(input_path)
        images[0].save(output_path, "JPEG")

    # ---------- DOCX to TXT ----------
    elif conversion == "docx2txt":
        output_path = os.path.abspath(
            os.path.join(OUTPUT_FOLDER, filename + ".txt")
        )
        doc = Document(input_path)
        with open(output_path, "w", encoding="utf-8") as f:
            for p in doc.paragraphs:
                f.write(p.text + "\n")

    else:
        return "Invalid conversion type"

    # ---------- SAFETY CHECK ----------
    if not os.path.exists(output_path):
        return "Conversion failed. Output file not created."

    return send_file(output_path, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
