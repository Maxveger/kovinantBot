from flask import Flask, request, send_file, redirect, url_for, Response
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

ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

def allowed_file(filename, types):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in types

def delayed_delete(path, delay_seconds=300):
    """Удаляет файл через delay секунд после отправки пользователю"""
    def delete_file():
        time.sleep(delay_seconds)
        if os.path.exists(path):
            try:
                os.remove(path)
            except:
                pass
    threading.Thread(target=delete_file).start()

def clean_tmp_folder():
    """Удаляет старые файлы Excel, docx, zip из /tmp при старте"""
    patterns = ('.xlsx', '.docx', '.zip')
    for filename in os.listdir('/tmp'):
        if filename.endswith(patterns):
            filepath = os.path.join('/tmp', filename)
            try:
                os.remove(filepath)
            except Exception as e:
                print(f"Не удалось удалить {filepath}: {e}")

@app.route('/')
def index():
    return '''
    <!doctype html>
    <html lang="ru">
    <head>
        <meta charset="UTF-8">
        <title>Генерация документов</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; }
            h1 { color: #333; }
            h2 { color: #555; }
            ul { line-height: 1.6; }
            p { margin-top: 10px; }
            .instruction { background: #f9f9f9; padding: 20px; border-radius: 8px; }
        </style>
    </head>
    <body>
        <h1>Загрузка Excel-файла для генерации документов</h1>
        <form action="/upload" method="post" enctype="multipart/form-data">
            <input type="file" name="file" accept=".xls,.xlsx" required>
            <input type="submit" value="Отправить">
        </form>

        <div class="instruction">
            <h2>Как пользоваться:</h2>
            <ul>
                <li>Подготовьте Excel-файл в формате .xls или .xlsx.</li>
                <li>В первой строке должны быть заголовки колонок (например: Имя, Дата, Адрес).</li>
                <li>Каждая строка ниже — это отдельная запись для генерации документов.</li>
                <li>Выберите файл и нажмите "Отправить".</li>
                <li>Через несколько секунд начнётся скачивание архива ZIP с готовыми файлами.</li>
            </ul>
            <p><b>Важно:</b> Убедитесь, что файл заполнен корректно, чтобы избежать ошибок при генерации.</p>
        </div>
    </body>
    </html>
    '''

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'Нет файла в запросе', 400

    file = request.files['file']

    if file.filename == '':
        return 'Файл не выбран', 400

    if file and allowed_file(file.filename, ALLOWED_EXTENSIONS):
        temp_excel_path = f"/tmp/{file.filename}"
        file.save(temp_excel_path)

        try:
            df = pd.read_excel(temp_excel_path)

            # Проверка содержимого Excel
            if df.empty or df.shape[1] == 0:
                os.remove(temp_excel_path)
                return 'Ошибка: файл пустой или без столбцов.', 400

            output_folder = "/tmp/generated_docs"
            os.makedirs(output_folder, exist_ok=True)

            for index, row in df.iterrows():
                act = Document(ACT_TEMPLATE)
                contract = Document(CONTRACT_TEMPLATE)

                # Замена полей
                for p in act.paragraphs:
                    for key, value in row.items():
                        if f"{{{{{key}}}}}" in p.text:
                            p.text = p.text.replace(f"{{{{{key}}}}}", str(value))

                for p in contract.paragraphs:
                    for key, value in row.items():
                        if f"{{{{{key}}}}}" in p.text:
                            p.text = p.text.replace(f"{{{{{key}}}}}", str(value))

                act_path = os.path.join(output_folder, f"akt_{index + 1}.docx")
                contract_path = os.path.join(output_folder, f"contract_{index + 1}.docx")
                act.save(act_path)
                contract.save(contract_path)

            zip_filename = f"/tmp/generated_docs_{datetime.now().strftime('%Y%m%d%H%M%S')}.zip"
            with zipfile.ZipFile(zip_filename, 'w') as zipf:
                for root, dirs, files in os.walk(output_folder):
                    for file in files:
                        filepath = os.path.join(root, file)
                        zipf.write(filepath, arcname=file)

            shutil.rmtree(output_folder)
            os.remove(temp_excel_path)

            response = send_file(zip_filename, as_attachment=True)

            @response.call_on_close
            def cleanup():
                try:
                    os.remove(zip_filename)
                except:
                    pass

            delayed_delete(zip_filename, delay_seconds=300)

            return response

        except Exception as e:
            if os.path.exists(temp_excel_path):
                os.remove(temp_excel_path)
            return f'Ошибка обработки: {str(e)}', 500

    else:
        return 'Неподдерживаемый тип файла', 400

# Чистим мусор в /tmp при запуске сервера
clean_tmp_folder()

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
