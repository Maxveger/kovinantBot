import os
from flask import Flask, request, send_file, render_template_string
from docxtpl import DocxTemplate
import pandas as pd
import zipfile
from io import BytesIO

app = Flask(__name__)

UPLOAD_FORM = """
<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <title>Загрузка Excel-файла</title>
  <style>
    body {
      font-family: sans-serif;
      background-color: #f8f8f8;
      display: flex;
      justify-content: center;
      align-items: center;
      height: 100vh;
    }
    .container {
      background: white;
      padding: 2em;
      border-radius: 10px;
      box-shadow: 0 0 10px rgba(0,0,0,0.1);
      text-align: center;
    }
    input[type=file], input[type=submit] {
      margin: 1em 0;
    }
  </style>
</head>
<body>
  <div class="container">
    <h1>Генерация Word-документов</h1>
    <p>Загрузите Excel-файл с данными:</p>
    <form method=post enctype=multipart/form-data>
      <input type=file name=file>
      <br>
      <input type=submit value="Загрузить и получить ZIP">
    </form>
  </div>
</body>
</html>
"""

@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        uploaded_file = request.files['file']
        if uploaded_file and uploaded_file.filename.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(uploaded_file)
            output = BytesIO()

            with zipfile.ZipFile(output, 'w') as zipf:
                for i, row in df.iterrows():
                    context = row.to_dict()

                    for template_file, suffix in [('templ_akt.docx', 'akt'), ('templ_dogovor.docx', 'dogovor')]:
                        doc = DocxTemplate(os.path.join('templates', template_file))
                        doc.render(context)
                        result_stream = BytesIO()
                        doc.save(result_stream)
                        result_stream.seek(0)

                        filename = f"{context.get('name', 'document')}_{suffix}.docx"
                        zipf.writestr(filename, result_stream.read())

            output.seek(0)
            return send_file(output, as_attachment=True, download_name='documents.zip', mimetype='application/zip')
        else:
            return 'Поддерживаются только Excel-файлы (.xls, .xlsx)'
    return render_template_string(UPLOAD_FORM)

@app.route('/')
def home():
    return '<h2>Сервис работает. Перейдите на <a href="/upload">/upload</a>, чтобы загрузить Excel-файл.</h2>'

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
