from flask import Flask, render_template, request, send_file, jsonify
from docx import Document
import os
from datetime import datetime

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        try:
            form_data = {
                "No": request.form["assignment_no"],
                "Here_Course_Code": request.form["course_code"],
                "Here_Course_Title": request.form["course_title"],
                "Here_TeacherName": request.form["teacher_name"],
                "designation": request.form["designation"],
                "Here_Teachers_Department_Name": request.form["teacher_dept"],
                "Here_StudentName": request.form["student_name"],
                "Here_StudentID": request.form["student_id"],
                "Here_Section": request.form["student_section"],
                "Here_DepartmentName": request.form["department_name"],
                "HereDate": request.form["submission_date"]
            }

            doc = Document("Assignment Cover Page.docx")
            for para in doc.paragraphs:
                for key, value in form_data.items():
                    if key in para.text:
                        for run in para.runs:
                            if key in run.text:
                                run.text = run.text.replace(key, value)

            os.makedirs("output", exist_ok=True)
            date_str = datetime.now().strftime("%Y-%m-%d")
            docx_path = os.path.join("output", f"CoverPage_{date_str}.docx")
            doc.save(docx_path)

            return jsonify({"success": True, "file": docx_path})
        except Exception as e:
            return jsonify({"success": False, "error": str(e)})

    return render_template("form.html")

@app.route("/download")
def download():
    from flask import request
    path = request.args.get("file")
    return send_file(path, as_attachment=True)

if __name__ == "__main__":
    print("✅ Server is running at http://127.0.0.1:5000")
    app.run(debug=True)

from flask import render_template

@app.route('/pdf-tools')
def pdf_tools():
    tools = [
        {"name": "Merge PDF", "description": "Combine multiple PDF files into one."},
        {"name": "Split PDF", "description": "Separate one PDF into multiple files."},
        {"name": "Compress PDF", "description": "Reduce PDF file size."},
        {"name": "PDF to DOCX", "description": "Convert PDF documents to Word format."},
        {"name": "PDF to PPTX", "description": "Convert PDF to PowerPoint slides."},
        {"name": "DOCX to PDF", "description": "Turn Word files into PDFs."},
        {"name": "Edit PDF", "description": "Add or modify text/images in PDF."},
        {"name": "PDF to JPG", "description": "Convert PDF pages to images."},
        {"name": "JPG to PDF", "description": "Convert image files to PDF."},
        {"name": "Sign & Watermark", "description": "Digitally sign and watermark your PDFs."},
        {"name": "Rotate PDF", "description": "Rotate pages within your PDF."},
        {"name": "Organize PDF", "description": "Reorder or delete pages in PDF."},
        {"name": "Add Page Numbers", "description": "Add page numbers to your PDF."},
    ]
    return render_template('pdf_tools.html', tools=tools)


import os
from werkzeug.utils import secure_filename
from PyPDF2 import PdfMerger

@app.route('/merge', methods=['GET', 'POST'])
def merge():
    merged_file = None
    if request.method == 'POST':
        files = request.files.getlist('pdfs')
        if files:
            merger = PdfMerger()
            save_path = os.path.join('static', 'merged')
            os.makedirs(save_path, exist_ok=True)
            output_filename = 'merged_result.pdf'
            output_path = os.path.join(save_path, output_filename)

            for file in files:
                if file and file.filename.endswith('.pdf'):
                    merger.append(file)

            merger.write(output_path)
            merger.close()
            merged_file = output_filename

    return render_template('merge.html', merged_file=merged_file)


from PyPDF2 import PdfReader, PdfWriter

@app.route('/split', methods=['GET', 'POST'])
def split():
    split_file = None
    if request.method == 'POST':
        file = request.files['pdf']
        start = int(request.form['start']) - 1  # 0-indexed
        end = int(request.form['end'])          # inclusive upper bound

        if file and file.filename.endswith('.pdf'):
            reader = PdfReader(file)
            writer = PdfWriter()

            # Ensure valid range
            num_pages = len(reader.pages)
            start = max(0, min(start, num_pages - 1))
            end = max(start + 1, min(end, num_pages))

            for page in range(start, end):
                writer.add_page(reader.pages[page])

            save_path = os.path.join('static', 'split')
            os.makedirs(save_path, exist_ok=True)
            output_filename = 'split_result.pdf'
            output_path = os.path.join(save_path, output_filename)

            with open(output_path, 'wb') as f:
                writer.write(f)

            split_file = output_filename

    return render_template('split.html', split_file=split_file)


import subprocess

@app.route('/compress', methods=['GET', 'POST'])
def compress():
    compressed_file = None
    if request.method == 'POST':
        file = request.files['pdf']
        if file and file.filename.endswith('.pdf'):
            input_path = os.path.join('static', 'compressed', 'original.pdf')
            output_filename = 'compressed_result.pdf'
            output_path = os.path.join('static', 'compressed', output_filename)

            os.makedirs('static/compressed', exist_ok=True)
            file.save(input_path)

            try:
                subprocess.run([
                    'gs', '-sDEVICE=pdfwrite', '-dCompatibilityLevel=1.4',
                    '-dPDFSETTINGS=/screen', '-dNOPAUSE', '-dQUIET', '-dBATCH',
                    f'-sOutputFile={output_path}', input_path
                ], check=True)
                compressed_file = output_filename
            except Exception as e:
                print("Compression error:", e)

    return render_template('compress.html', compressed_file=compressed_file)


from pdf2docx import Converter

@app.route('/pdf-to-docx', methods=['GET', 'POST'])
def pdf_to_docx():
    docx_file = None
    if request.method == 'POST':
        file = request.files['pdf']
        if file and file.filename.endswith('.pdf'):
            input_path = os.path.join('static', 'docx', 'input.pdf')
            output_filename = 'converted.docx'
            output_path = os.path.join('static', 'docx', output_filename)

            os.makedirs('static/docx', exist_ok=True)
            file.save(input_path)

            try:
                cv = Converter(input_path)
                cv.convert(output_path, start=0, end=None)
                cv.close()
                docx_file = output_filename
            except Exception as e:
                print("PDF to DOCX conversion error:", e)

    return render_template('pdf_to_docx.html', docx_file=docx_file)


from pdf2image import convert_from_path
from pptx import Presentation
from pptx.util import Inches

@app.route('/pdf-to-pptx', methods=['GET', 'POST'])
def pdf_to_pptx():
    pptx_file = None
    if request.method == 'POST':
        file = request.files['pdf']
        if file and file.filename.endswith('.pdf'):
            input_path = os.path.join('static', 'pptx', 'input.pdf')
            output_filename = 'converted.pptx'
            output_path = os.path.join('static', 'pptx', output_filename)

            os.makedirs('static/pptx', exist_ok=True)
            file.save(input_path)

            try:
                # Convert PDF pages to images
                images = convert_from_path(input_path)
                prs = Presentation()

                # Resize slide to match A4 (roughly)
                prs.slide_width = Inches(11.7)
                prs.slide_height = Inches(8.3)

                for image in images:
                    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
                    image_path = os.path.join('static', 'pptx', 'temp.jpg')
                    image.save(image_path, 'JPEG')
                    slide.shapes.add_picture(image_path, 0, 0, width=prs.slide_width, height=prs.slide_height)

                prs.save(output_path)
                pptx_file = output_filename
            except Exception as e:
                print("PDF to PPTX conversion error:", e)

    return render_template('pdf_to_pptx.html', pptx_file=pptx_file)


import platform
import subprocess

@app.route('/docx-to-pdf', methods=['GET', 'POST'])
def docx_to_pdf():
    pdf_file = None
    if request.method == 'POST':
        file = request.files['docx']
        if file and file.filename.endswith('.docx'):
            os.makedirs('static/docx_to_pdf', exist_ok=True)
            input_path = os.path.join('static', 'docx_to_pdf', 'input.docx')
            output_path = os.path.join('static', 'docx_to_pdf')
            file.save(input_path)

            try:
                if platform.system() == 'Windows':
                    from docx2pdf import convert
                    convert(input_path, output_path)
                else:
                    subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', output_path, input_path], check=True)

                pdf_file = 'input.pdf'
            except Exception as e:
                print("DOCX to PDF conversion error:", e)

    return render_template('docx_to_pdf.html', pdf_file=pdf_file)


@app.route('/edit-pdf', methods=['GET', 'POST'])
def edit_pdf():
    edited_file = None
    if request.method == 'POST':
        file = request.files['pdf']
        if file and file.filename.endswith('.pdf'):
            os.makedirs('static/edited', exist_ok=True)
            filename = 'editable.pdf'
            output_path = os.path.join('static', 'edited', filename)
            file.save(output_path)
            edited_file = filename

    return render_template('edit_pdf.html', edited_file=edited_file)


@app.route('/pdf-to-jpg', methods=['GET', 'POST'])
def pdf_to_jpg():
    images = []
    if request.method == 'POST':
        file = request.files['pdf']
        if file and file.filename.endswith('.pdf'):
            os.makedirs('static/pdf_to_jpg', exist_ok=True)
            input_path = os.path.join('static', 'pdf_to_jpg', 'input.pdf')
            file.save(input_path)

            try:
                from pdf2image import convert_from_path
                jpg_pages = convert_from_path(input_path)
                images = []
                for i, img in enumerate(jpg_pages):
                    image_filename = f'page_{i+1}.jpg'
                    image_path = os.path.join('static', 'pdf_to_jpg', image_filename)
                    img.save(image_path, 'JPEG')
                    images.append(image_filename)
            except Exception as e:
                print("PDF to JPG conversion error:", e)

    return render_template('pdf_to_jpg.html', images=images)


@app.route('/jpg-to-pdf', methods=['GET', 'POST'])
def jpg_to_pdf():
    pdf_file = None
    if request.method == 'POST':
        files = request.files.getlist('images')
        images = []

        os.makedirs('static/jpg_to_pdf', exist_ok=True)
        for file in files:
            if file and file.filename.lower().endswith('.jpg'):
                img = Image.open(file.stream).convert('RGB')
                images.append(img)

        if images:
            output_path = os.path.join('static', 'jpg_to_pdf', 'merged.pdf')
            images[0].save(output_path, save_all=True, append_images=images[1:])
            pdf_file = 'merged.pdf'

    return render_template('jpg_to_pdf.html', pdf_file=pdf_file)


@app.route('/sign-watermark-pdf', methods=['GET', 'POST'])
def sign_watermark_pdf():
    result_file = None
    if request.method == 'POST':
        file = request.files['pdf']
        watermark_text = request.form['watermark']
        if file and file.filename.endswith('.pdf') and watermark_text:
            os.makedirs('static/signed', exist_ok=True)
            input_path = os.path.join('static', 'signed', 'original.pdf')
            watermark_path = os.path.join('static', 'signed', 'watermark.pdf')
            output_path = os.path.join('static', 'signed', 'signed.pdf')
            file.save(input_path)

            # Create watermark PDF
            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import letter
            c = canvas.Canvas(watermark_path, pagesize=letter)
            c.setFont("Helvetica", 36)
            c.setFillAlpha(0.3)
            c.drawCentredString(300, 400, watermark_text)
            c.save()

            # Merge watermark with original
            from PyPDF2 import PdfReader, PdfWriter
            original = PdfReader(input_path)
            watermark = PdfReader(watermark_path)
            writer = PdfWriter()

            for page in original.pages:
                page.merge_page(watermark.pages[0])
                writer.add_page(page)

            with open(output_path, "wb") as f:
                writer.write(f)

            result_file = "signed.pdf"

    return render_template('sign_watermark_pdf.html', result_file=result_file)


@app.route('/rotate-pdf', methods=['GET', 'POST'])
def rotate_pdf():
    rotated_file = None
    if request.method == 'POST':
        file = request.files['pdf']
        rotation = int(request.form['rotation'])
        if file and file.filename.endswith('.pdf'):
            os.makedirs('static/rotated', exist_ok=True)
            input_path = os.path.join('static', 'rotated', 'input.pdf')
            output_path = os.path.join('static', 'rotated', 'rotated.pdf')
            file.save(input_path)

            from PyPDF2 import PdfReader, PdfWriter
            reader = PdfReader(input_path)
            writer = PdfWriter()

            for page in reader.pages:
                page.rotate(rotation)
                writer.add_page(page)

            with open(output_path, "wb") as f:
                writer.write(f)

            rotated_file = "rotated.pdf"

    return render_template('rotate_pdf.html', rotated_file=rotated_file)


@app.route('/organize-pdf', methods=['GET', 'POST'])
def organize_pdf():
    organized_file = None
    if request.method == 'POST':
        file = request.files['pdf']
        page_order = request.form['page_order']
        if file and file.filename.endswith('.pdf') and page_order:
            os.makedirs('static/organized', exist_ok=True)
            input_path = os.path.join('static', 'organized', 'input.pdf')
            output_path = os.path.join('static', 'organized', 'organized.pdf')
            file.save(input_path)

            order_list = [int(i.strip()) - 1 for i in page_order.split(',') if i.strip().isdigit()]

            from PyPDF2 import PdfReader, PdfWriter
            reader = PdfReader(input_path)
            writer = PdfWriter()

            for idx in order_list:
                if 0 <= idx < len(reader.pages):
                    writer.add_page(reader.pages[idx])

            with open(output_path, "wb") as f:
                writer.write(f)

            organized_file = "organized.pdf"

    return render_template('organize_pdf.html', organized_file=organized_file)


@app.route('/add-page-numbers', methods=['GET', 'POST'])
def add_page_numbers():
    numbered_file = None
    if request.method == 'POST':
        file = request.files['pdf']
        if file and file.filename.endswith('.pdf'):
            os.makedirs('static/numbered', exist_ok=True)
            input_path = os.path.join('static', 'numbered', 'input.pdf')
            output_path = os.path.join('static', 'numbered', 'numbered.pdf')
            file.save(input_path)

            from PyPDF2 import PdfReader, PdfWriter
            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import letter
            from PyPDF2.pdf import PageObject
            import io

            reader = PdfReader(input_path)
            writer = PdfWriter()

            for i, page in enumerate(reader.pages):
                packet = io.BytesIO()
                can = canvas.Canvas(packet, pagesize=letter)
                can.setFont("Helvetica", 10)
                can.drawString(500, 20, f"{i + 1}")
                can.save()

                packet.seek(0)
                new_pdf = PdfReader(packet)
                page.merge_page(new_pdf.pages[0])
                writer.add_page(page)

            with open(output_path, "wb") as f:
                writer.write(f)

            numbered_file = "numbered.pdf"

    return render_template('add_page_numbers.html', numbered_file=numbered_file)

@app.route('/home')
def home():
    return render_template('home.html')
