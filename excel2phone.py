import os
import pandas as pd
import re
import zipfile
import time
from datetime import datetime
from flask import Flask, request, send_file, render_template, make_response
from werkzeug.utils import secure_filename
import openpyxl
from io import BytesIO

app = Flask(__name__)

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    files = request.files.getlist("file")
    if not files:
        return '没有选中的文件', 400
    else:
        current_time = datetime.now().strftime('%Y%m%d%H%M')  # current date and time in 'YYYYMMDDHHMM' format
        filenames = []
        for file in files:
            if file:
                filename = current_time + "_" + secure_filename(file.filename)  # prepend the timestamp to the filename
                filepath = os.path.join('/tmp', filename)
                file.save(filepath)
                if os.path.isdir(filepath):  # If the uploaded file is a directory
                    process_directory(filepath, filenames)
                else:
                    processed_filepath = process_file(filepath, file.filename)
                    if processed_filepath:
                        filenames.append(os.path.basename(processed_filepath))

        num_files = len(filenames)

        if not filenames:
            return '没有找到符合要求的文件', 400

        return render_template('index.html', filenames=filenames, num_files=num_files)

def process_directory(directory, filenames):
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.lower().endswith(('.xls', '.xlsx')):
                filepath = os.path.join(root, file)
                processed_filepath = process_file(filepath, file)
                if processed_filepath:
                    filenames.append(os.path.basename(processed_filepath))

def process_file(filepath, original_filename):
    try:
        if original_filename.lower().endswith('.xls'):
            # Convert .xls to .xlsx
            filepath_xlsx = os.path.splitext(filepath)[0] + '.xlsx'
            xls = pd.ExcelFile(filepath, engine='xlrd')
            with pd.ExcelWriter(filepath_xlsx, engine='openpyxl') as writer:
                for sheet_name in xls.sheet_names:
                    df_xls = pd.read_excel(xls, sheet_name=sheet_name)
                    df_xls.to_excel(writer, sheet_name=sheet_name, index=False)
            filepath = filepath_xlsx
            original_filename = os.path.splitext(original_filename)[0] + '.xlsx'

        xlsx = pd.ExcelFile(filepath, engine='openpyxl')
        identifiable_numbers = set()
        for sheet_name in xlsx.sheet_names:
            df = pd.read_excel(xlsx, sheet_name=sheet_name)
            process_worksheet(df, identifiable_numbers)

        base_filename = os.path.splitext(original_filename)[0]
        output_filename = os.path.join('/tmp', base_filename + ".txt")
        with open(output_filename, 'w') as f:
            for phone_number in identifiable_numbers:
                f.write(phone_number + '\n')
        return output_filename
    except Exception as e:
        print("Error processing file:", str(e))
        return None

def process_worksheet(df, identifiable_numbers):
    phone_pattern = r"(?:13[0-9]|14[0-1,4-9]|15[0-3,5-9]|16[2,5-7]|17[1-9]|18[0-9]|19[0,1,3,5-9])\d{8}"
    email_pattern = r"^([a-zA-Z0-9_.+-]+@(126.com|163.com|wo.cn|189.com|139.com))$"
    for column in df:
        for cell in df[column]:
            if pd.notna(cell):
                cell_strs = re.split('\n|;|\s', str(cell))
                cell_strs = [i for s in cell_strs for i in s.split('/')]
                for cell_str in cell_strs:
                    if re.match(email_pattern, cell_str):
                        email_split = cell_str.split("@")
                        email_numbers = re.findall(phone_pattern, email_split[0])
                        identifiable_numbers.update(email_numbers)
                    else:
                        numbers = re.findall(phone_pattern, cell_str)
                        identifiable_numbers.update(numbers)

@app.route('/download', methods=['GET'])
def download_file():
    current_time = datetime.now().strftime('%Y%m%d%H%M')
    zip_filename = "/tmp/" + current_time + "_processed.zip"
    filenames = request.args.getlist("filename")
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for filename in filenames:
            filepath = os.path.join('/tmp', filename)
            zipf.write(filepath, os.path.basename(filepath))

    data = BytesIO()
    with open(zip_filename, 'rb') as f:
        data.write(f.read())
    data.seek(0)

    os.remove(zip_filename)

    response = make_response(data.getvalue())
    response.headers.set('Content-Type', 'application/zip')
    response.headers.set('Content-Disposition', 'attachment', filename=current_time + "_processed.zip")

    return response

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=80)
