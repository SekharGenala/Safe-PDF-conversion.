import os
from flask import Flask, render_template, request, send_from_directory
from PIL import Image
from docx import Document

app = Flask(__name__)

# Configure upload folder
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


@app.route('/')
def index():
    return render_template('jpgw.html')


@app.route('/uploaded', methods=['POST'])
def uploaded():
    # Check if file was submitted
    if 'file' not in request.files:
        return 'No file uploaded'

    file = request.files['file']

    # Check if the file is an image
    if file.filename == '':
        return 'No file selected'

    if file and allowed_file(file.filename):
        # Save the uploaded file to the upload folder
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)

        # Convert the image to Word document
        image = Image.open(file_path)
        document = Document()
        document.add_picture(file_path)
        word_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{os.path.splitext(file.filename)[0]}.docx')
        document.save(word_file_path)

        return render_template('download_jpg_word.html', filename=os.path.basename(word_file_path))
    else:
        return 'Invalid file format'


@app.route('/download/<filename>')
def download(filename):
    # Return the converted Word file for download
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)


def allowed_file(filename):
    # Check if the file has an allowed extension
    allowed_extensions = {'jpg', 'jpeg', 'png'}
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions


if __name__ == '__main__':
    app.run(debug=True)
