from flask import Flask, request, send_from_directory
import os
import shutil
import pandas as pd
from docx import Document
from docxtpl import DocxTemplate

app = Flask(__name__)

# Папка для загрузки файлов и сгенерированных документов
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'generated_documents'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Разрешённые расширения файлов
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

# Функция для проверки разрешённого расширения файла
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Главная страница с улучшенным оформлением
@app.route('/')
def index():
    return '''
        <html>
            <head>
                <title>Генератор документов</title>
                <style>
                    body {
                        font-family: Arial, sans-serif;
                        background-color: #f4f4f9;
                        padding: 20px;
                    }
                    h1 {
                        color: #333;
                    }
                    form {
                        background-color: #fff;
                        padding: 20px;
                        border-radius: 8px;
                        box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
                    }
                    input[type="file"] {
                        margin-top: 10px;
                    }
                </style>
            </head>
            <body>
                <h1>Загрузите Excel файл для генерации документов</h1>
                <form method="POST" enctype="multipart/form-data" action="/upload">
                    <input type="file" name="file" />
                    <input type="submit" value="Загрузить файл" />
                </form>
            </body>
        </html>
    '''

# Страница для загрузки файла
@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'Нет файла в запросе', 400

    file = request.files['file']

    if file.filename == '':
        return 'Файл не выбран', 400

    if file and allowed_file(file.filename):
        filename = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filename)
        return process_file(filename)
    else:
        return 'Неверный формат файла', 400

# Обработка Excel файла и генерация документов
def process_file(file_path):
    data = pd.read_excel(file_path, dtype=str)
    if "date_pass" in data.columns:
        data["date_pass"] = pd.to_datetime(data["date_pass"], errors='coerce').dt.strftime("%d.%m.%Y")

    for index, row in data.iterrows():
        new_filename = os.path.join(OUTPUT_FOLDER, f"{row['name']}.docx")
        shutil.copy("templates/templ_akt.docx", new_filename)
        new_doc = Document(new_filename)

        # Заменяем текст в абзацах
        for para in new_doc.paragraphs:
            for column in data.columns:
                placeholder = f"{{{column}}}"
                if placeholder in para.text:
                    para.text = para.text.replace(placeholder, str(row[column]))

        # Заменяем текст в таблицах
        for table in new_doc.tables:
            for row_cells in table.rows:
                for cell in row_cells.cells:
                    for column in data.columns:
                        placeholder = f"{{{column}}}"
                        if placeholder in cell.text:
                            cell.text = cell.text.replace(placeholder, str(row[column]))

        new_doc.save(new_filename)

    return send_from_directory(OUTPUT_FOLDER, f"{row['name']}.docx", as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.getenv('PORT', 5000)))
