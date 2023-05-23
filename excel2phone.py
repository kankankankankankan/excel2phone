from flask import Flask, request, send_file, render_template
from werkzeug.utils import secure_filename
import os
import pandas as pd
import re
import zipfile
import time
from datetime import datetime

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
        filenames = []
        for file in files:
            if file:
                current_time = datetime.now().strftime('%Y%m%d%H%M')  # current date and time in 'YYYYMMDDHHMM' format
                filename = current_time + "_" + secure_filename(file.filename)  # prepend the timestamp to the filename
                filepath = os.path.join('/tmp', filename)
                file.save(filepath)
                processed_filepath = process_file(filepath, file.filename)
                filenames.append(processed_filepath)
        
        # Create a zip file
        zip_filename = "/tmp/" + current_time + "_processed.zip"
        with zipfile.ZipFile(zip_filename, 'w') as zipf:
            for filename in filenames:
                # Add file to the zip file
                # Second argument is the name of the file in the zip file
                zipf.write(filename, os.path.basename(filename))

        return send_file(zip_filename, as_attachment=True)


def process_file(filepath, original_filename):
    df = pd.read_excel(filepath)
    identifiable_numbers = set()  # use set to avoid duplicates
    phone_pattern = r"^((13[0-9])|(14([0-1]|[4-9]))|(15([0-3]|[5-9]))|(16(2|[5-7]))|(17[1-9])|(18[0-9])|(19[0|1|3])|(19[5-9]))\d{8}$"
    email_pattern = r"^([a-zA-Z0-9_.+-]+@(126.com|163.com|wo.cn|189.com|139.com))$"
    for column in df:
        for cell in df[column]:
            if pd.notna(cell):
                cell_str = str(cell)
                if re.match(email_pattern, cell_str):
                    email_split = cell_str.split("@")
                    email_number = re.search(phone_pattern, email_split[0])
                    if email_number:
                        identifiable_numbers.add(email_number.group())
                else:
                    number = re.search(phone_pattern, cell_str)
                    if number:
                        identifiable_numbers.add(number.group())
    base_filename = os.path.splitext(original_filename)[0]  # get the original file's base name (no extension)
    output_filename = filepath.replace(secure_filename(original_filename), base_filename) + ".txt"
    with open(output_filename, 'w') as f:
        for phone_number in identifiable_numbers:
            f.write(phone_number + '\n')
    return output_filename  # return the name of created file



if __name__ == "__main__":
    app.run(host='0.0.0.0', port=80)
