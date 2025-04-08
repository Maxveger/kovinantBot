import os
from flask import Flask, request, render_template, send_from_directory
from docx import Document
import pandas as pd
from datetime import datetime

app = Flask(__name__)

# Папка для шаблонов
UPLOAD_FOLDER = 'templates/'
ALLOWED_EXTENSIONS = {'docx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Шаблоны по умолчанию
DEFAULT_TEMPLATES = {
    'akt': 'templates/templ_akt.docx',
    'dogovor': 'templates/templ_dogovor.docx'
}

# Функция для проверки формата файла
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Функция для загрузки и замены шаблона
def upload_template(template_type, file):
    if not allowed_file(file.filename):
        return "Invalid file format. Please upload a .docx file", 400

    if template_type not in DEFAULT_TEMPLATES:
        return "Invalid template type", 400
    
    # Определяем путь для шаблона
    template_path = os.path.join(app.config['UPLOAD_FOLDER'], f"templ_{template_type}.docx")
    
    # Проверяем, существует ли файл, если да — заменяем
    if os.path.exists(template_path):
        os.remove(template_path)
    
    # Сохраняем новый файл
    file.save(template_path)
    return f"{template_type.capitalize()} template updated successfully", 200

@app.route('/')
def home():
    # Отображаем текущие даты последних обновлений шаблонов
    akt_last_updated = get_last_updated_date('templ_akt.docx')
    dogovor_last_updated = get_last_updated_date('templ_dogovor.docx')
    return render_template('index.html', 
                           akt_last_updated=akt_last_updated, 
                           dogovor_last_updated=dogovor_last_updated)

@app.route('/upload_template_akt', methods=['POST'])
def upload_template_akt():
    if 'file' not in request.files:
        return "No file part", 400
    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400
    return upload_template('akt', file)

@app.route('/upload_template_dogovor', methods=['POST'])
def upload_template_dogovor():
    if 'file' not in request.files:
        return "No file part", 400
    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400
    return upload_template('dogovor', file)

@app.route('/generate_documents', methods=['POST'])
def generate_documents():
    # Загружаем данные из Excel
    data = pd.read_excel("to_akts.xlsx", dtype=str)
    if "date_pass" in data.columns:
        data["date_pass"] = pd.to_datetime(data["date_pass"], errors='coerce').dt.strftime("%d.%m.%Y")

    # Проверяем наличие шаблонов и выбираем, какой использовать
    akt_template = DEFAULT_TEMPLATES['akt']
    dogovor_template = DEFAULT_TEMPLATES['dogovor']

    # Если шаблон был обновлен, используем новый
    if os.path.exists('templates/templ_akt.docx'):
        akt_template = 'templates/templ_akt.docx'
    if os.path.exists('templates/templ_dogovor.docx'):
        dogovor_template = 'templates/templ_dogovor.docx'

    # Создаём папку для готовых документов
    output_folder = "generated_documents"
    os.makedirs(output_folder, exist_ok=True)

    # Генерируем документы для актов
    for index, row in data.iterrows():
        new_filename = os.path.join(output_folder, f"{row['name']}_akt.docx")
        document = Document(akt_template)
        replace_placeholders_in_document(document, row, data.columns)
        document.save(new_filename)

    # Генерируем документы для договоров
    for index, row in data.iterrows():
        new_filename = os.path.join(output_folder, f"{row['name']}_dogovor.docx")
        document = Document(dogovor_template)
        replace_placeholders_in_document(document, row, data.columns)
        document.save(new_filename)

    return f"Все документы сохранены в папке: '{output_folder}'"

# Функция для замены плейсхолдеров в документе
def replace_placeholders_in_document(document, row, columns):
    for para in document.paragraphs:
        for column in columns:
            placeholder = f"{{{column}}}"
            if placeholder in para.text:
                para.text = para.text.replace(placeholder, str(row[column]))
    
    for table in document.tables:
        for row_cells in table.rows:
            for cell in row_cells.cells:
                for column in columns:
                    placeholder = f"{{{column}}}"
                    if placeholder in cell.text:
                        cell.text = cell.text.replace(placeholder, str(row[column]))

# Функция для получения даты последнего обновления шаблона
def get_last_updated_date(template_name):
    template_path = os.path.join(app.config['UPLOAD_FOLDER'], template_name)
    if os.path.exists(template_path):
        timestamp = os.path.getmtime(template_path)
        return datetime.fromtimestamp(timestamp).strftime("%d.%m.%Y %H:%M:%S")
    return "Not updated yet"

if __name__ == '__main__':
    # Используем переменную окружения для порта
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
