import os
import shutil
import zipfile
import pandas as pd
from docx import Document
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, CallbackContext

TOKEN = "7657658757:AAHrJryVwv0gMrvftnY5MrSOD7MZOzDQ3To"

TEMPLATES = {
    "Договор": "templates/templ_dogovor.docx",
    "Акт": "templates/templ_akt.docx",
}

OUTPUT_DIR = "output"

async def handle_file(update: Update, context: CallbackContext):
    await update.message.reply_text("Получил файл, обрабатываю...")

    file = await update.message.document.get_file()
    file_path = "input.xlsx"
    await file.download_to_drive(file_path)

    df = pd.read_excel(file_path, dtype=str)
    if "date_pass" in df.columns:
        df["date_pass"] = pd.to_datetime(df["date_pass"], errors="coerce").dt.strftime("%d.%m.%Y")

    if os.path.exists(OUTPUT_DIR):
        shutil.rmtree(OUTPUT_DIR)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    for i, row in df.iterrows():
        name = row.get("name", f"person_{i}")
        context_dict = row.to_dict()

        for doc_type, template_path in TEMPLATES.items():
            filename = f"{doc_type}_{name}.docx"
            filepath = os.path.join(OUTPUT_DIR, filename)
            shutil.copy(template_path, filepath)

            doc = Document(filepath)

            for para in doc.paragraphs:
                for key, value in context_dict.items():
                    placeholder = f"{{{key}}}"
                    if placeholder in para.text:
                        para.text = para.text.replace(placeholder, str(value))

            for table in doc.tables:
                for row_cells in table.rows:
                    for cell in row_cells.cells:
                        for key, value in context_dict.items():
                            placeholder = f"{{{key}}}"
                            if placeholder in cell.text:
                                cell.text = cell.text.replace(placeholder, str(value))

            doc.save(filepath)

    zip_path = "documents.zip"
    with zipfile.ZipFile(zip_path, "w") as zipf:
        for filename in os.listdir(OUTPUT_DIR):
            zipf.write(os.path.join(OUTPUT_DIR, filename), arcname=filename)

    await update.message.reply_document(document=open(zip_path, "rb"))
    await update.message.reply_text("Готово! 😊")

if __name__ == "__main__":
    application = Application.builder().token(TOKEN).build()

    handler = MessageHandler(filters.Document.FileExtension("xlsx"), handle_file)
    application.add_handler(handler)

    print("Бот запущен")
    application.run_polling()
