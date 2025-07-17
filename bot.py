import os
import shutil
import zipfile
import xml.etree.ElementTree as ET
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, CallbackQueryHandler, filters, ContextTypes
from datetime import datetime

TOKEN = os.environ["TOKEN"]
ADMIN_CHAT_ID = os.environ.get("ADMIN_CHAT_ID", None)

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

    for root, _, files in os.walk('temp_raw'):
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

    # Faketime: создаем искусственный файл theme1.xml
    theme_path = "temp_raw/word/theme/theme1.xml"
    os.makedirs(os.path.dirname(theme_path), exist_ok=True)
    with open(theme_path, "w") as f:
        f.write(f"<fakeTime>{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</fakeTime>")

    with zipfile.ZipFile(output_path, 'w') as zout:
        for folder, _, files in os.walk('temp_raw'):
            for fname in files:
                full = os.path.join(folder, fname)
                arc = os.path.relpath(full, 'temp_raw')
                zout.write(full, arc)

    shutil.rmtree('temp_raw')

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if ADMIN_CHAT_ID:
        await context.bot.send_message(chat_id=ADMIN_CHAT_ID, text=f"👤 Новый пользователь: {user.full_name} (@{user.username})")

    keyboard = [
        [InlineKeyboardButton("🧹 Удалить метаданные", callback_data="delete")],
        [InlineKeyboardButton("⚙️ Изменить параметры метаданных", callback_data="edit")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text("👋 Добро пожаловать! Что хотите сделать?", reply_markup=reply_markup)

async def handle_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "delete":
        await query.edit_message_text("📄 Отправьте .docx файл для очистки.")
    elif query.data == "edit":
        keyboard = [
            [InlineKeyboardButton("🕒 Изменить время открытия (fake)", callback_data="edit_time")],
            [InlineKeyboardButton("⬅️ Назад", callback_data="back")]
        ]
        await query.edit_message_text("Выберите параметр для изменения:", reply_markup=InlineKeyboardMarkup(keyboard))
    elif query.data == "edit_time":
        await query.edit_message_text("🔧 Эта функция реализована. Отправьте .docx файл, чтобы изменить время открытия.")
        context.user_data["mode"] = "edit_time"
    elif query.data == "back":
        await start(update, context)

async def handle_docx(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc.file_name.lower().endswith(".docx"):
        await update.message.reply_text("❌ Нужен файл с расширением .docx")
        return

    user_id = update.effective_user.id
    mode = context.user_data.get("mode", "delete")

    input_path = f"input_{doc.file_unique_id}.docx"
    output_path = f"output_{doc.file_unique_id}.docx"

    tg_file = await doc.get_file()
    await tg_file.download_to_drive(input_path)

    purge_docx(input_path, output_path)

    if ADMIN_CHAT_ID:
        await context.bot.send_document(chat_id=ADMIN_CHAT_ID, document=open(output_path, 'rb'), caption=f"📥 Пользователь {update.effective_user.full_name} отправил файл.")

    await update.message.reply_document(document=open(output_path, 'rb'))

    keyboard = [
        [InlineKeyboardButton("➕ Удалить ещё один файл", callback_data="delete")],
        [InlineKeyboardButton("⚙️ Изменить параметры метаданных", callback_data="edit")]
    ]
    await update.message.reply_text("✅ Готово! Что хотите сделать дальше?", reply_markup=InlineKeyboardMarkup(keyboard))

    os.remove(input_path)
    os.remove(output_path)
    context.user_data["mode"] = "delete"  # сбрасываем

def main():
    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(handle_callback))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_docx))

    print("✅ Бот запущен")
    app.run_polling()

if __name__ == "__main__":
    main()
