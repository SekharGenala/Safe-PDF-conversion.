from flask import Flask, render_template, request, send_from_directory
from PIL import Image
from fpdf import FPDF

app = Flask(__name__)

# Set the upload folder path
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


@app.route('/')
def index():
    return render_template('jpg.html')


@app.route('/convert', methods=['POST'])
def convert():
    # Get the uploaded image file
    image_file = request.files['image']

    # Save the image to the upload folder
    image_path = f"{app.config['UPLOAD_FOLDER']}/{image_file.filename}"
    image_file.save(image_path)

    # Convert the image to PDF
    pdf_path = f"{app.config['UPLOAD_FOLDER']}/{image_file.filename.split('.')[0]}.pdf"
    image = Image.open(image_path)
    pdf = FPDF()
    pdf.add_page()
    pdf.image(image_path, 0, 0, pdf.w, pdf.h)
    pdf.output(pdf_path, "F")

    # Provide the download link for the converted PDF file
    return render_template('download_jpg.html', pdf_filename=f"{image_file.filename.split('.')[0]}.pdf")


@app.route('/download/<filename>')
def download(filename):
    # Get the path of the PDF file
    pdf_path = f"{app.config['UPLOAD_FOLDER']}/{filename}"

    # Download the PDF file
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
