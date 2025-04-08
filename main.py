from flask import Flask, request, send_file, redirect, url_for
import os
import shutil
import pandas as pd
from docx import Document
import zipfile
from datetime import datetime

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
ACTS_FOLDER = 'generated_acts'
CONTRACTS_FOLDER = 'generated_contracts'
TEMPLATES_FOLDER = 'templates'
ACT_TEMPLATE = os.path.join(TEMPLATES_FOLDER, 'templ_akt.docx')
CONTRACT_TEMPLATE = os.path.join(TEMPLATES_FOLDER, 'templ_dogovor.docx')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(ACTS_FOLDER, exist_ok=True)
os.makedirs(CONTRACTS_FOLDER, exist_ok=True)
os.makedirs(TEMPLATES_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {'xls', 'xlsx', 'docx'}

success_message = ""


def allowed_file(filename, types):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in types


@app.route('/')
def index():
    global success_message
    message_html = f'<p style="color: green;">{success_message}</p>' if success_message else ''
    success_message = ""  # сброс после показа

    return f'''
        <html>
            <head>
                <title>Генератор документов</title>
                <style>
                    body {{
                        font-family: Arial, sans-serif;
                        background-color: #f0f0f5;
                        margin: 0;
                        padding: 0;
                        display: flex;
                        align-items: center;
                        justify-content: center;
                        height: 100vh;
                    }}
                    .container {{
                        background: white;
                        padding: 30px;
                        border-radius: 10px;
                        box-shadow: 0 0 15px rgba(0,0,0,0.1);
                        width: 400px;
                        text-align: center;
                    }}
                    h1 {{
                        font-size: 20px;
                        margin-bottom: 20px;
                    }}
                    input[type="file"] {{
                        margin: 10px 0;
                    }}
                    input[type="submit"] {{
                        padding: 8px 16px;
                        border: none;
                        background-color: #4CAF50;
                        color: white;
                        border-radius: 5px;
                        cursor: pointer;
                    }}
                    input[type="submit"]:hover {{
                        background-color: #45a049;
                    }}
                    p.note {{
                        font-size: 12px;
                        color: #888;
                        margin-top: 20px;
                    }}
                </style>
            </head>
            <body>
                <div class="container">
                    <h1>Генерация документов</h1>
                    {message_html}
                    <form method="POST" enctype="multipart/form-data" action="/upload">
                        <input type="file" name="file" required><br>
                        <input type="submit" value="Загрузить Excel">
                    </form>
                    <hr>
                    <form method="POST" enctype="multipart/form-data" action="/upload_template">
                        <label>Обновить шаблон акта:</label><br>
                        <input type="file" name="template_akt">
                        <input type="submit" name="submit_type" value="Загрузить акт"><br><br>
                        <label>Обновить шаблон договора:</label><br>
                        <input type="file" name="template_dogovor">
                        <input type="submit" name="submit_type" value="Загрузить договор">
                    </form>
                    <p class="note">Готовый архив ищите в "Загрузках".</p>
                </div>
            </body>
        </html>
    '''


@app.route('/upload_template', methods=['POST'])
def upload_template():
    global success_message
    if 'submit_type' in request.form:
        if request.form['submit_type'] == 'Загрузить акт' and 'template_akt' in request.files:
            file = request.files['template_akt']
            if allowed_file(file.filename, {'docx'}):
                file.save(ACT_TEMPLATE)
                success_message = 'Шаблон акта обновлён.'
        elif request.form['submit_type'] == 'Загрузить договор' and 'template_dogovor' in request.files:
            file = request.files['template_dogovor']
            if allowed_file(file.filename, {'docx'}):
                file.save(CONTRACT_TEMPLATE)
                success_message = 'Шаблон договора обновлён.'
    return redirect(url_for('index'))


@app.route('/upload', methods=['POST'])
def upload_file():
    global success_message
    if 'file' not in request.files:
        success_message = 'Нет файла в запросе.'
        return redirect(url_for('index'))

    file = request.files['file']
    if file.filename == '':
        success_message = 'Файл не выбран.'
        return redirect(url_for('index'))

    if file and allowed_file(file.filename, {'xls', 'xlsx'}):
        filename = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filename)
        return process_file(filename)
    else:
        success_message = 'Неверный формат файла.'
        return redirect(url_for('index'))


def replace_placeholders(doc: Document, data_row):
    for para in doc.paragraphs:
        for column in data_row.index:
            placeholder = f"{{{column}}}"
            if placeholder in para.text:
                para.text = para.text.replace(placeholder, str(data_row[column]))

    for table in doc.tables:
        for row_cells in table.rows:
            for cell in row_cells.cells:
                for column in data_row.index:
                    placeholder = f"{{{column}}}"
                    if placeholder in cell.text:
                        cell.text = cell.text.replace(placeholder, str(data_row[column]))


def process_file(file_path):
    global success_message

    data = pd.read_excel(file_path, dtype=str)
    if "date_pass" in data.columns:
        data["date_pass"] = pd.to_datetime(data["date_pass"], errors='coerce').dt.strftime("%d.%m.%Y")

    for index, row in data.iterrows():
        act_filename = os.path.join(ACTS_FOLDER, f"{row['name']}_act.docx")
        shutil.copy(ACT_TEMPLATE, act_filename)
        new_act = Document(act_filename)
        replace_placeholders(new_act, row)
        new_act.save(act_filename)

        contract_filename = os.path.join(CONTRACTS_FOLDER, f"{row['name']}_contract.docx")
        shutil.copy(CONTRACT_TEMPLATE, contract_filename)
        new_contract = Document(contract_filename)
        replace_placeholders(new_contract, row)
        new_contract.save(contract_filename)

    zip_filename = 'generated_documents.zip'
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for root, dirs, files in os.walk(ACTS_FOLDER):
            for file in files:
                zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), ACTS_FOLDER))
        for root, dirs, files in os.walk(CONTRACTS_FOLDER):
            for file in files:
                zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), CONTRACTS_FOLDER))

    success_message = 'Документы сгенерированы. Архив загружается...'
    return send_file(zip_filename, as_attachment=True)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.getenv('PORT', 5000)))
