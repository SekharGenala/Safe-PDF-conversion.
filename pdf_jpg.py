from flask import Flask, render_template, request, send_from_directory
from pdf2image import convert_from_path
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'


@app.route('/')
def pdf():
    return render_template('pdf.html')


@app.route('/convert', methods=['POST'])
def convert():
    # Get the uploaded PDF file
    file = request.files['pdf']

    # Save the PDF file
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(pdf_path)

    # Convert PDF to JPG
    images = convert_from_path(pdf_path)

    # Save each image as JPG
    jpg_filenames = []
    for i, image in enumerate(images):
        jpg_filename = f"{i}.jpg"
        jpg_path = os.path.join(app.config['UPLOAD_FOLDER'], jpg_filename)
        image.save(jpg_path)
        jpg_filenames.append(jpg_filename)

    return render_template('download_pdf_jpg.html', filenames=jpg_filenames)


@app.route('/download/<filename>')
def download(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)


if __name__ == '__main__':
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])
    app.run(debug=True)
