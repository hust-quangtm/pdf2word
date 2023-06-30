import os
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from pdf2docx import Converter

app = Flask(__name__)

# Create the 'tmp' directory if it doesn't exist
if not os.path.exists('tmp'):
    os.makedirs('tmp')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    # Get the uploaded PDF file from the form
    pdf_file = request.files['pdf_file']
    pdf_filename = secure_filename(pdf_file.filename)

    # Save the PDF file to a temporary location
    pdf_path = os.path.join('tmp', pdf_filename)
    pdf_file.save(pdf_path)

    # Convert PDF to Word
    docx_filename = pdf_filename.rsplit('.', 1)[0] + '.docx'
    docx_path = os.path.join('tmp', docx_filename)
    convert_pdf_to_docx(pdf_path, docx_path)

    # Send the converted Word document as a download
    return send_file(docx_path, as_attachment=True, download_name=docx_filename)

def convert_pdf_to_docx(pdf_path, docx_path):
    # Create a PDF to Word converter object
    cv = Converter(pdf_path)

    # Convert PDF to Word
    cv.convert(docx_path, start=0, end=None)

    # Close the converter
    cv.close()

if __name__ == '__main__':
    app.run(debug=True)
