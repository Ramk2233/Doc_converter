from flask import Flask, render_template, request, send_from_directory, url_for, redirect
from werkzeug.utils import secure_filename
from docx2pdf import convert
import requests
import os
import pythoncom


app = Flask(__name__)

# change locations accordingly for upload and download locations
UPLOAD_FOLDER = '/Users/hp/Desktop/python-doc_converter/doc_converter/Uploads'
DOWNLOAD_FOLDER = '/Users/hp/Desktop/python-doc_converter/doc_converter/Downloads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER


@app.route('/')
def upload_file():
    return render_template('fileconverter.html')


@app.route('/converter/<filename>', methods=['GET', 'POST'])
def converted_file(filename):

    return send_from_directory(app.config['DOWNLOAD_FOLDER'], filename, as_attachment=True)


@app.route('/converter', methods=['GET', 'POST'])
def upload_file_1():

    pythoncom.CoInitialize()

    if request.method == 'POST':
        f = request.files['file']

        f.save(os.path.join(app.config['UPLOAD_FOLDER'], f.filename))

        uploaded_filename = f.filename
        new_file_name = uploaded_filename.rsplit('.', 1)[0]

        convert(f"{UPLOAD_FOLDER}/{uploaded_filename}",
                f"{DOWNLOAD_FOLDER}/{new_file_name}.pdf")

        Converted_file_name = new_file_name + ".pdf"

    return redirect(url_for('converted_file', filename=Converted_file_name))


if __name__ == '__main__':
    app.run(debug=True, port=5000)
