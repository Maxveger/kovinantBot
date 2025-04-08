import os
import shutil
from flask import Flask, render_template, request, redirect, url_for
from werkzeug.utils import secure_filename
from docx import Document
import pandas as pd

app = Flask(__name__)

# Конфигурация для загрузки файлов
app.config['UPLOAD_FOLDER'] = 'templates'
app.config['ALLOWED_EXTENSIONS'] = {'docx'}

# Функция для проверки разрешённых расширений
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

# Загрузка шаблона
@app.route('/upload', methods=['GET', 'POST'])
def upload_template():
    if request.method == 'POST':
        # Проверяем, что файл был выбран
        if 'template_file' not in request.files:
            return 'Нет файла'

        file = request.files['template_file']
        
        # Если файл не был выбран
        if file.filename == '':
            return 'Не выбрано ни одного файла'
        
        # Проверяем, что файл имеет допустимое расширение
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            # Сохраняем файл с оригинальным именем
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            return redirect(url_for('index'))

    return render_template('upload.html')

# Главная страница
@app.route('/')
def index():
    return render_template('index.html')

# Генерация документов
@app.route('/generate', methods=['POST'])
def generate_documents():
    # Загружаем данные из Excel
    data = pd.read_excel("to_akts.xlsx", dtype=str)  # Читаем всё как строки
    if "date_pass" in data.columns:
        data["date_pass"] = pd.to_datetime(data["date_pass"], errors='coerce').dt.strftime("%d.%m.%Y")  # Форматируем дату

    # Создаём папку для готовых документов
    output_folder = "generated_documents"
    os.makedirs(output_folder, exist_ok=True)

    # Загружаем шаблоны
    try:
        template_akt = Document(os.path.join('templates', 'templ_akt.docx'))
        template_dogovor = Document(os.path.join('templates', 'templ_dogovor.docx'))
    except Exception as e:
        return f"Ошибка загрузки шаблонов: {str(e)}"

    # Проходим по каждой строке в Excel и генерируем документы
    for index, row in data.iterrows():
        # Генерация для акта
        akt_filename = os.path.join(output_folder, f"{row['name']}_akt.docx")
        new_doc_akt = Document(template_akt)
        replace_placeholders(new_doc_akt, row)
        new_doc_akt.save(akt_filename)

        # Генерация для договора
        dogovor_filename = os.path.join(output_folder, f"{row['name']}_dogovor.docx")
        new_doc_dogovor = Document(template_dogovor)
        replace_placeholders(new_doc_dogovor, row)
        new_doc_dogovor.save(dogovor_filename)

    return 'Документы успешно сгенерированы!'

# Замена заполнителей в документах
def replace_placeholders(doc, row):
    for para in doc.paragraphs:
        for column in row.index:
            placeholder = f"{{{{ {column} }}}}"
            if placeholder in para.text:
                para.text = para.text.replace(placeholder, str(row[column]))

    for table in doc.tables:
        for row_cells in table.rows:
            for cell in row_cells.cells:
                for column in row.index:
                    placeholder = f"{{{{ {column} }}}}"
                    if placeholder in cell.text:
                        cell.text = cell.text.replace(placeholder, str(row[column]))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
