from flask import Flask, request, send_file
import pandas as pd
from docx import Document
import zipfile
import os
import io

app = Flask(__name__)

@app.route('/')
def index():
    return 'Сервис работает. Загрузите файл на /upload'

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return 'Файл не найден', 400

    file = request.files['file']
    df = pd.read_excel(file)

    # Обработка и создание Word-документов
    doc1 = Document('templates/шаблон1.docx')
    doc2 = Document('templates/шаблон2.docx')

    # Пример подстановки — сделай как у тебя в локальной версии
    for p in doc1.paragraphs:
        p.text = p.text.replace('{{name}}', str(df.iloc[0]['Имя']))

    for p in doc2.paragraphs:
        p.text = p.text.replace('{{name}}', str(df.iloc[0]['Имя']))

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, 'w') as zipf:
        doc1_io = io.BytesIO()
        doc1.save(doc1_io)
        zipf.writestr('output1.docx', doc1_io.getvalue())

        doc2_io = io.BytesIO()
        doc2.save(doc2_io)
        zipf.writestr('output2.docx', doc2_io.getvalue())

    buffer.seek(0)
    return send_file(buffer, mimetype='application/zip', as_attachment=True, download_name='documents.zip')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
