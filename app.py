from flask import Flask, request, render_template, jsonify, Response, url_for
import os
from reportlab.pdfgen import canvas
from docx2pdf import convert

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert_docx_to_pdf():
    if 'docx_file' not in request.files:
        error_message = 'No file provided.'
        return render_template('index.html', error=error_message)

    docx_file = request.files['docx_file']

    if docx_file.filename == '':
        error_message = 'No file selected.'
        return render_template('index.html', error=error_message)

    if not docx_file.filename.endswith('.docx'):
        error_message = 'The selected file has not .docx extension!'
        return render_template('index.html', error=error_message)
    
    docx_filepath = os.path.join(app.config['UPLOAD_FOLDER'], docx_file.filename)
    pdf_filepath = os.path.splitext(docx_filepath)[0] + '.pdf'

    docx_file.save(docx_filepath)  # Save the uploaded DOCX file

    # Convert DOCX to PDF using python-docx2pdf
    try:
        convert(docx_filepath, pdf_filepath)
    except Exception as e:
        error_message = str(e)
        return render_template('index.html', error=error_message)

    # Serve the converted PDF file as a response
    with open(pdf_filepath, 'rb') as pdf_file:
        pdf_data = pdf_file.read()

    #success_message = 'Conversion done. Download started.'
    response = Response(pdf_data, content_type='application/pdf')
    response.headers['Content-Disposition'] = f'attachment; filename="{os.path.basename(pdf_filepath)}"'
    return response



if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(debug=True, threaded=False)
