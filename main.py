from flask import Flask, request, send_from_directory, render_template, redirect
import os
import shutil
import zipfile
from docxtpl import DocxTemplate
import pandas as pd
from datetime import datetime

app = Flask(__name__)

TEMPLATES_DIR = 'templates'
UPLOAD_FOLDER = 'uploads'
os.makedirs(TEMPLATES_DIR, exist_ok=True)
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/upload_template', methods=['POST'])
def upload_template():
    if 'contract_template' in request.files:
        contract_template = request.files['contract_template']
        contract_template.save(os.path.join(TEMPLATES_DIR, 'templ_dogovor.docx'))
    
    if 'act_template' in request.files:
        act_template = request.files['act_template']
        act_template.save(os.path.join(TEMPLATES_DIR, 'templ_akt.docx'))

    return redirect('/')

@app.route('/upload_data', methods=['POST'])
def upload_data():
    if 'datafile' not in request.files:
        return 'No file part', 400
    
    datafile = request.files['datafile']
    datafile.save(os.path.join(UPLOAD_FOLDER, 'datafile.xlsx'))

    # Читаем Excel файл
    data = pd.read_excel(os.path.join(UPLOAD_FOLDER, 'datafile.xlsx'), dtype=str)
    if "date_pass" in data.columns:
        data["date_pass"] = pd.to_datetime(data["date_pass"], errors='coerce').dt.strftime("%d.%m.%Y")
    
    # Создаем папку для документов
    output_folder = os.path.join(UPLOAD_FOLDER, "generated_documents")
    os.makedirs(output_folder, exist_ok=True)

    # Проверяем существующие шаблоны
    contract_template = os.path.join(TEMPLATES_DIR, 'templ_dogovor.docx')
    act_template = os.path.join(TEMPLATES_DIR, 'templ_akt.docx')

    if not os.path.exists(contract_template) or not os.path.exists(act_template):
        return 'Шаблоны документов не найдены, загрузите их снова', 400

    # Генерация документов
    for index, row in data.iterrows():
        # Договор
        contract_filename = os.path.join(output_folder, f"{row['name']}_dogovor.docx")
        doc = DocxTemplate(contract_template)
        doc.render(row.to_dict())
        doc.save(contract_filename)
        
        # Акт
        act_filename = os.path.join(output_folder, f"{row['name']}_akt.docx")
        doc = DocxTemplate(act_template)
        doc.render(row.to_dict())
        doc.save(act_filename)

    # Архивируем результаты
    archive_path = os.path.join(UPLOAD_FOLDER, 'generated_documents.zip')
    with zipfile.ZipFile(archive_path, 'w') as zipf:
        for root, dirs, files in os.walk(output_folder):
            for file in files:
                zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), output_folder))

    return send_from_directory(UPLOAD_FOLDER, 'generated_documents.zip', as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
