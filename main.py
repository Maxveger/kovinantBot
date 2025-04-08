from flask import Flask, request, send_file, redirect
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
TEMPLATE_FOLDER = 'templates'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(ACTS_FOLDER, exist_ok=True)
os.makedirs(CONTRACTS_FOLDER, exist_ok=True)
os.makedirs(TEMPLATE_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {'xls', 'xlsx'}
TEMPLATES = {
    'templ_akt.docx': 'Шаблон акта',
    'templ_dogovor.docx': 'Шаблон договора'
}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    act_time = get_template_update_time('templ_akt.docx')
    dog_time = get_template_update_time('templ_dogovor.docx')
    return f'''
        <html>
            <head><title>Генератор документов</title></head>
            <body style="font-family: sans-serif">
                <h1>Генерация документов</h1>
                <form method="POST" enctype="multipart/form-data" action="/upload">
                    <p><b>Excel-файл</b>:</p>
                    <input type="file" name="file" />
                    <input type="submit" value="Сгенерировать" />
                </form>
                <hr>
                <h2>Обновить шаблоны</h2>
                <form method="POST" enctype="multipart/form-data" action="/upload_template">
                    <p><b>Загрузить новый шаблон акта (templ_akt.docx)</b><br>
                    Последнее обновление: {act_time}</p>
                    <input type="file" name="template" />
                    <input type="submit" value="Обновить акт" />
                </form>
                <form method="POST" enctype="multipart/form-data" action="/upload_template">
                    <p><b>Загрузить новый шаблон договора (templ_dogovor.docx)</b><br>
                    Последнее обновление: {dog_time}</p>
                    <input type="file" name="template" />
                    <input type="submit" value="Обновить договор" />
                </form>
            </body>
        </html>
    '''

@app.route('/upload_template', methods=['POST'])
def upload_template():
    file = request.files.get('template')
    if not file or not file.filename:
        return 'Файл не выбран', 400

    filename = file.filename
    if filename not in TEMPLATES:
        return f'Имя файла должно быть одним из: {", ".join(TEMPLATES.keys())}', 400

    file.save(os.path.join(TEMPLATE_FOLDER, filename))
    return redirect('/')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'Нет файла в запросе', 400
    file = request.files['file']
    if file.filename == '':
        return 'Файл не выбран', 400
    if file and allowed_file(file.filename):
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)
        return process_file(filepath)
    return 'Неверный формат файла', 400

def process_file(file_path):
    data = pd.read_excel(file_path, dtype=str)
    if "date_pass" in data.columns:
        data["date_pass"] = pd.to_datetime(data["date_pass"], errors='coerce').dt.strftime("%d.%m.%Y")

    for index, row in data.iterrows():
        act_filename = os.path.join(ACTS_FOLDER, f"{row['name']}_act.docx")
        shutil.copy(os.path.join(TEMPLATE_FOLDER, "templ_akt.docx"), act_filename)
        fill_doc(act_filename, row)

        contract_filename = os.path.join(CONTRACTS_FOLDER, f"{row['name']}_contract.docx")
        shutil.copy(os.path.join(TEMPLATE_FOLDER, "templ_dogovor.docx"), contract_filename)
        fill_doc(contract_filename, row)

    zip_filename = 'generated_documents.zip'
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for root, _, files in os.walk(ACTS_FOLDER):
            for file in files:
                zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), ACTS_FOLDER))
        for root, _, files in os.walk(CONTRACTS_FOLDER):
            for file in files:
                zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), CONTRACTS_FOLDER))

    return send_file(zip_filename, as_attachment=True)

def fill_doc(doc_path, row):
    doc = Document(doc_path)
    for para in doc.paragraphs:
        for column, value in row.items():
            placeholder = f"{{{column}}}"
            if placeholder in para.text:
                para.text = para.text.replace(placeholder, str(value))
    for table in doc.tables:
        for row_cells in table.rows:
            for cell in row_cells.cells:
                for column, value in row.items():
                    placeholder = f"{{{column}}}"
                    if placeholder in cell.text:
                        cell.text = cell.text.replace(placeholder, str(value))
    doc.save(doc_path)

def get_template_update_time(filename):
    path = os.path.join(TEMPLATE_FOLDER, filename)
    if os.path.exists(path):
        timestamp = os.path.getmtime(path)
        return datetime.fromtimestamp(timestamp).strftime('%d.%m.%Y %H:%M:%S')
    return 'Не загружено'

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.getenv('PORT', 5000)))
