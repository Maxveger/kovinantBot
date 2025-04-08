from flask import Flask, request, send_file, render_template_string
import os
import zipfile
import pandas as pd  # если используешь для обработки
from io import BytesIO

app = Flask(__name__)

# HTML-форма
UPLOAD_FORM = """
<!doctype html>
<title>Загрузить Excel</title>
<h1>Загрузите Excel-файл</h1>
<form method=post enctype=multipart/form-data>
  <input type=file name=file>
  <input type=submit value=Загрузить>
</form>
"""

@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        uploaded_file = request.files['file']
        if uploaded_file.filename.endswith('.xlsx') or uploaded_file.filename.endswith('.xls'):
            # Обработка файла
            input_excel = uploaded_file.read()
            df = pd.read_excel(BytesIO(input_excel))

            # Создание документа(ов) — пока просто сохраняем Excel как CSV
            output = BytesIO()
            with zipfile.ZipFile(output, 'w') as zipf:
                csv_data = df.to_csv(index=False).encode('utf-8')
                zipf.writestr('converted.csv', csv_data)

            output.seek(0)
            return send_file(output, as_attachment=True, download_name='result.zip', mimetype='application/zip')
        else:
            return 'Поддерживаются только Excel-файлы (.xls, .xlsx)'
    return render_template_string(UPLOAD_FORM)

@app.route('/')
def home():
    return '<h2>Сервис работает. Перейдите на <a href="/upload">/upload</a> чтобы загрузить файл.</h2>'
