from flask import Flask, render_template, request, send_file
import os
import PyPDF2
import pandas as pd

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def convert_pdf_to_excel(file_path, output_path):
    with open(file_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        num_pages = len(pdf_reader.pages)
        data = []
        for page_num in range(num_pages):
            page = pdf_reader.pages[page_num]
            page_data = page.extract_text()
            data.append(page_data)

        df = pd.DataFrame(data)
        df.to_excel(output_path, index=False)

@app.route('/')
def index():
    return render_template('pdfe.html')

@app.route('/convert', methods=['POST'])
def convert():
    file = request.files['pdf_file']
    filename = file.filename
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)

    output_file = os.path.splitext(filename)[0] + '.xlsx'
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_file)

    convert_pdf_to_excel(file_path, output_path)

    return render_template('download_pdf_excel.html', filename=output_file)


@app.route('/uploads/<filename>', methods=['GET'])
def download(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
