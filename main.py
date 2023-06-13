from flask import Flask, render_template, request, redirect, send_file
import os
import pdf2docx
from werkzeug.utils import secure_filename
#from word import  allowed_file, word, upload_file, download



app = Flask(__name__)

# Set the upload folder path
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Route to handle the upload page
@app.route('/')
def index():
    return render_template('index.html')

# Route to handle file upload and conversion
@app.route('/upload', methods=['POST'])
def upload_file():
    # Get the uploaded file
    file = request.files['file']

    # Save the file to the uploads folder
    filename = secure_filename(file.filename)
    file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

    # Convert the file to DOCX format
    pdf_file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    docx_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{os.path.splitext(filename)[0]}.docx")
    pdf2docx.parse(pdf_file_path, docx_file_path)

    # Download the converted DOCX file
    return send_file(docx_file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)


#word to pdf converter



