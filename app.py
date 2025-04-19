from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
import os
import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
import zipfile

app = Flask(__name__)

# Set up upload folder
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'docx'}

# Function to check allowed file extensions
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    # Check if files are part of the request
    if 'excel_file' not in request.files or 'word_template' not in request.files:
        return 'No file part', 400

    excel_file = request.files['excel_file']
    word_template = request.files['word_template']

    # Validate the files
    if excel_file.filename == '' or not allowed_file(excel_file.filename):
        return 'Invalid Excel file', 400
    if word_template.filename == '' or not allowed_file(word_template.filename):
        return 'Invalid Word template file', 400

    # Save files securely
    excel_filename = secure_filename(excel_file.filename)
    word_filename = secure_filename(word_template.filename)

    excel_path = os.path.join(UPLOAD_FOLDER, excel_filename)
    word_path = os.path.join(UPLOAD_FOLDER, word_filename)

    excel_file.save(excel_path)
    word_template.save(word_path)

    # Process the files and generate PDFs
    df = pd.read_excel(excel_path)
    doc = DocxTemplate(word_path)

    output_files = []
    for index, row in df.iterrows():
        context = {
            "employeeName": row["employeeName"],
            "Credits": row["Credits"]
        }
        doc.render(context)

        output_word_path = os.path.join(OUTPUT_FOLDER, f"{row['employeeName']}.docx")
        doc.save(output_word_path)
        convert(output_word_path)  # Convert to PDF

        output_pdf_path = os.path.join(OUTPUT_FOLDER, f"{row['employeeName']}.pdf")
        output_files.append(output_pdf_path)

    # Create a ZIP file of the generated PDFs
    zip_filename = 'generated_pdfs.zip'
    zip_path = os.path.join(OUTPUT_FOLDER, zip_filename)
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for file in output_files:
            zipf.write(file, os.path.basename(file))

    # Return the ZIP file as a download
    return send_file(zip_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
