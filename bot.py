import os
import shutil
import zipfile
import xml.etree.ElementTree as ET
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, MessageHandler, CommandHandler, CallbackQueryHandler, filters, ContextTypes

TOKEN = os.environ["TOKEN"]

# Файлы, которые разрешено оставить
ALLOWED_XML = {
    '[Content_Types].xml',
    '_rels/.rels',
    'word/document.xml',
    'word/styles.xml',
    'word/_rels/document.xml.rels'
}

def purge_docx(input_path: str, output_path: str):
    with zipfile.ZipFile(input_path, 'r') as zin:
        zin.extractall('temp_raw')

    for root, dirs, files in os.walk('temp_raw'):
        for fname in files:
            rel = os.path.relpath(os.path.join(root, fname), 'temp_raw')
            if rel not in ALLOWED_XML:
                os.remove(os.path.join(root, fname))

    doc_xml = 'temp_raw/word/document.xml'
    if os.path.exists(doc_xml):
        tree = ET.parse(doc_xml)
        rt = tree.getroot()
        tags = [
            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author',
            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}trackRevisions',
            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeStart',
            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeEnd',
            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentReference'
        ]
        for tag in tags:
            for el in rt.findall('.//' + tag):
                parent = rt.find('.//' + tag + '/..')
                if parent is not None and el in parent:
                    parent.remove(el)
        tree.write(doc_xml)

    with zipfile.ZipFile(output_path, 'w') as zout:
        for folder, _, files in os.walk('temp_raw'):
            for fname in files:
                full = os.path.join(folder, fname)
                arc = os.path.relpath(full, 'temp_raw')
                zout.write(full, arc)

    shutil.rmtree('temp_raw')

# Главное меню
async def show_main_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [InlineKeyboardButton("🧹 Удалить метаданные", callback_data='clean')],
        [InlineKeyboardButton("⚙️ Изменить параметры метаданных", callback_data='edit')]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)

    if update.message:
        await update.message.reply_text("👋 Добро пожаловать! Что хотите сделать?", reply_markup=reply_markup)
    elif update.callback_query:
        await update.callback_query.message.reply_text("Что хотите сделать дальше?", reply_markup=reply_markup)

# Обработка кнопок
async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == 'clean':
        await query.message.reply_text("📎 Отправьте .docx файл для очистки.")
    elif query.data == 'edit':
        await query.message.reply_text("🛠️ Функция редактирования метаданных пока в разработке.")

# Обработка .docx файлов
async def handle_docx(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document

    if not doc.file_name.lower().endswith('.docx'):
        await update.message.reply_text("❌ Пожалуйста, отправьте именно .docx файл.")
        return

    in_path = f"input_{doc.file_unique_id}.docx"
    out_path = f"cleaned_{doc.file_unique_id}.docx"

    tg_file = await doc.get_file()
    await tg_file.download_to_drive(in_path)

    purge_docx(in_path, out_path)

    await update.message.reply_document(document=open(out_path, 'rb'))

    # Меню после обработки
    keyboard = [
        [InlineKeyboardButton("🧹 Удалить ещё один файл", callback_data='clean')],
        [InlineKeyboardButton("⚙️ Изменить параметры метаданных", callback_data='edit')]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text("✅ Готово! Что делаем дальше?", reply_markup=reply_markup)

    os.remove(in_path)
    os.remove(out_path)

def main():
    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", show_main_menu))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_docx))

    print("🚀 Бот запущен.")
    app.run_polling()

if __name__ == "__main__":
    main()
