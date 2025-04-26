from flask import Flask, request, send_file, redirect, url_for, render_template_string
import os
import shutil
import pandas as pd
from docx import Document
import zipfile
from datetime import datetime
import threading
import time

app = Flask(__name__)

TEMPLATES_FOLDER = 'templates'
ACT_TEMPLATE = os.path.join(TEMPLATES_FOLDER, 'templ_akt.docx')
CONTRACT_TEMPLATE = os.path.join(TEMPLATES_FOLDER, 'templ_dogovor.docx')
GENERATED_FOLDER = 'generated_documents'
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <title>Генерация документов</title>
    <style>
        body {
            font-family: sans-serif;
            background-color: #f4f4f9;
            margin: 0;
            padding: 2em;
        }
        .container {
            max-width: 600px;
            margin: auto;
            background: white;
            padding: 2em;
            border-radius: 10px;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
        }
        h2 {
            text-align: center;
            margin-bottom: 1em;
        }
        form {
            margin-bottom: 1em;
        }
        input[type=file], button {
            padding: 0.5em;
            margin-top: 0.5em;
        }
        hr {
            margin: 2em 0;
        }
        ul {
            padding-left: 1.2em;
        }
    </style>
</head>
<body>
    <div class="container">
        <h2>Генерация документов</h2>

        <form method="POST" enctype="multipart/form-data" action="/upload">
            <input type="file" name="file">
            <br>
            <button type="submit">Загрузить Excel</button>
        </form>

        <hr>

        <p><strong>Обновить шаблон акта:</strong></p>
        <form method="POST" enctype="multipart/form-data" action="/upload_template/akt">
            <input type="file" name="file">
            <br>
            <button type="submit">Загрузить акт</button>
        </form>

        <p><strong>Обновить шаблон договора:</strong></p>
        <form method="POST" enctype="multipart/form-data" action="/upload_template/dogovor">
            <input type="file" name="file">
            <br>
            <button type="submit">Загрузить договор</button>
        </form>

        <hr>

        <h3>Как пользоваться:</h3>
        <ul>
            <li>Подготовьте Excel-файл в формате .xls или .xlsx.</li>
            <li>В первой строке — заголовки колонок (например: Имя, Дата, Адрес).</li>
            <li>Каждая строка — отдельная запись для генерации документов.</li>
            <li>Нажмите «Загрузить Excel» — начнётся обработка.</li>
            <li>Через несколько секунд начнётся скачивание ZIP-архива с готовыми файлами.</li>
        </ul>
        <p><strong>Важно:</strong> Убедитесь, что Excel-файл заполнен корректно.</p>
    </div>
</body>
</html>
"""

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def generate_documents(df):
    shutil.rmtree(GENERATED_FOLDER, ignore_errors=True)
    os.makedirs(GENERATED_FOLDER, exist_ok=True)
    for i, row in df.iterrows():
        act = Document(ACT_TEMPLATE)
        contract = Document(CONTRACT_TEMPLATE)

        # Здесь ваша логика подстановки значений в шаблоны
        act.paragraphs[0].text = str(row[0])
        contract.paragraphs[0].text = str(row[0])

        act.save(f"{GENERATED_FOLDER}/akt_{i+1}.docx")
        contract.save(f"{GENERATED_FOLDER}/contract_{i+1}.docx")

    zip_path = os.path.join(GENERATED_FOLDER, 'archive.zip')
    with zipfile.ZipFile(zip_path, 'w') as zf:
        for file in os.listdir(GENERATED_FOLDER):
            if file.endswith('.docx'):
                zf.write(os.path.join(GENERATED_FOLDER, file), file)
    return zip_path

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files.get('file')
    if file and allowed_file(file.filename):
        df = pd.read_excel(file)
        archive_path = generate_documents(df)
        return send_file(archive_path, as_attachment=True)
    return redirect(url_for('index'))

@app.route('/upload_template/<template_type>', methods=['POST'])
def upload_template(template_type):
    file = request.files.get('file')
    if file and file.filename.endswith('.docx'):
        if template_type == 'akt':
            file.save(ACT_TEMPLATE)
        elif template_type == 'dogovor':
            file.save(CONTRACT_TEMPLATE)
    return redirect(url_for('index'))

def cleanup_old_temp():
    while True:
        now = time.time()
        for folder in ['generated_documents', 'temp']:  # если что-то ещё будет
            if os.path.exists(folder):
                for f in os.listdir(folder):
                    path = os.path.join(folder, f)
                    if os.path.isfile(path) and now - os.path.getmtime(path) > 600:
                        os.remove(path)
        time.sleep(300)

if __name__ == '__main__':
    os.makedirs(TEMPLATES_FOLDER, exist_ok=True)
    os.makedirs(GENERATED_FOLDER, exist_ok=True)
    threading.Thread(target=cleanup_old_temp, daemon=True).start()
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
