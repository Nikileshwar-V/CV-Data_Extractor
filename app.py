from flask import Flask, render_template, request, redirect, url_for, send_from_directory
import os
import re
import PyPDF2
import docx
import pandas as pd
from werkzeug.utils import secure_filename

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf', 'doc', 'docx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def extract_text_from_pdf(file_path):
    text = ''
    with open(file_path, 'rb') as f:
        pdf_reader = PyPDF2.PdfReader(f)
        for page in pdf_reader.pages:
            text += page.extract_text()
    return text

def extract_text_from_docx(file_path):
    doc = docx.Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

def extract_contact_numbers(text):
    phone_numbers = re.findall(r'\b(?:0|\+?91)?[.-]?[0-9]{10}\b', text)
    return phone_numbers

def extract_email_addresses(text):
    email_addresses = re.findall(r'\b[\w.-]+?@\w+?\.\w+?\b', text)
    return email_addresses

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_files():
    if request.method == 'POST':
        uploaded_files = request.files.getlist('file')
        all_texts = []
        all_contact_numbers = []
        all_email_addresses = []

        for file in uploaded_files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(file_path)
                if filename.endswith('.pdf'):
                    text = extract_text_from_pdf(file_path)
                elif filename.endswith('.doc') or filename.endswith('.docx'):
                    text = extract_text_from_docx(file_path)
                else:
                    continue

                contact_numbers = extract_contact_numbers(text)
                email_addresses = extract_email_addresses(text)

                all_texts.append(text)
                all_contact_numbers.append('\n'.join(contact_numbers))
                all_email_addresses.append('\n'.join(email_addresses))

        data = {
            'Text': all_texts,
            'Contact Numbers': all_contact_numbers,
            'Email Addresses': all_email_addresses
        }
        df = pd.DataFrame(data)
        excel_file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'extracted_info.xlsx')
        df.to_excel(excel_file_path, index=False)
        return redirect(url_for('uploaded_file', filename='extracted_info.xlsx'))
    return render_template('upload.html')

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
