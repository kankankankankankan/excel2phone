from flask import Flask, request, send_file, render_template
from werkzeug.utils import secure_filename
import os
import pandas as pd
import re

app = Flask(__name__)

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return '没有文件被上传', 400
    file = request.files['file']
    if file.filename == '':
        return '没有选中的文件', 400
    if file:
        filename = secure_filename(file.filename)
        filepath = os.path.join('/tmp', filename)
        file.save(filepath)
        process_file(filepath)
        return send_file('phone.txt', as_attachment=True)

def process_file(filepath):
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
    with open('phone.txt', 'w') as f:
        for phone_number in identifiable_numbers:
            f.write(phone_number + '\n')

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=80)
