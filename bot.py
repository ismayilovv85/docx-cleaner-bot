import os
import shutil
import zipfile
import datetime
import xml.etree.ElementTree as ET
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, InputFile
from telegram.ext import (
    ApplicationBuilder, MessageHandler, CommandHandler,
    CallbackQueryHandler, ConversationHandler, filters, ContextTypes
)

TOKEN = os.environ["TOKEN"]

# Conversation states
(
    MENU,
    WAIT_DOC_FOR_CLEAN,
    WAIT_DOC_FOR_EDIT,
    ASK_AUTHOR,
    ASK_COMPANY,
    ASK_CREATED,
    ASK_MODIFIED,
    ASK_LASTPRINT
) = range(8)

user_data_store = {}

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [[
        InlineKeyboardButton("\U0001F9F1 Удалить метаданные", callback_data="clean"),
        InlineKeyboardButton("\u2699 Изменить параметры метаданных", callback_data="edit")
    ]]
    await update.message.reply_text(
        "\U0001F44B Добро пожаловать! Что хотите сделать?",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )
    return MENU

async def menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "clean":
        await query.message.reply_text("\U0001F4C4 Отправьте .docx файл для очистки.")
        return WAIT_DOC_FOR_CLEAN
    else:
        await query.message.reply_text("\u270F\ufe0f Эта функция реализована. Отправьте .docx файл, чтобы изменить метаданные.")
        return WAIT_DOC_FOR_EDIT

# =================== Очистка =======================
def purge_docx(input_path: str, output_path: str):
    with zipfile.ZipFile(input_path, 'r') as zin:
        zin.extractall('temp_raw')

    ALLOWED_XML = {
        '[Content_Types].xml',
        '_rels/.rels',
        'word/document.xml',
        'word/styles.xml',
        'word/_rels/document.xml.rels'
    }
    for root, _, files in os.walk('temp_raw'):
        for fname in files:
            rel = os.path.relpath(os.path.join(root, fname), 'temp_raw')
            if rel not in ALLOWED_XML:
                os.remove(os.path.join(root, fname))

    doc_xml = 'temp_raw/word/document.xml'
    if os.path.exists(doc_xml):
        tree = ET.parse(doc_xml)
        rt = tree.getroot()
        for tag in ['w:author', 'w:trackRevisions', 'w:commentRangeStart', 'w:commentRangeEnd', 'w:commentReference']:
            for el in rt.findall(f'.//{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}{tag[2:]}'):
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

async def handle_clean_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc.file_name.lower().endswith('.docx'):
        await update.message.reply_text("Нужен именно .docx файл.")
        return WAIT_DOC_FOR_CLEAN

    in_path = f"input_{doc.file_unique_id}.docx"
    out_path = f"cleaned_{doc.file_unique_id}.docx"
    await doc.get_file().download_to_drive(in_path)
    purge_docx(in_path, out_path)
    await update.message.reply_document(document=InputFile(out_path))
    os.remove(in_path)
    os.remove(out_path)

    return await restart_menu(update)

# =================== Изменение метаданных =======================
def update_metadata(docx_path, output_path, author, company, created, modified, last_print):
    with zipfile.ZipFile(docx_path, 'r') as zin:
        zin.extractall('temp_raw')

    core_path = 'temp_raw/docProps/core.xml'
    if os.path.exists(core_path):
        tree = ET.parse(core_path)
        root = tree.getroot()
        ns = {'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
              'dc': 'http://purl.org/dc/elements/1.1/',
              'dcterms': 'http://purl.org/dc/terms/',
              'xsi': 'http://www.w3.org/2001/XMLSchema-instance'}

        def set_elem(tag, val):
            elem = root.find(f'dc:{tag}', ns)
            if elem is None:
                elem = ET.SubElement(root, f'{{{ns['dc']}}}{tag}')
            elem.text = val

        set_elem('creator', author)
        set_elem('title', 'Document cleaned by Docx Cleaner')

        for tag, val in [('created', created), ('modified', modified), ('lastPrinted', last_print)]:
            e = root.find(f'dcterms:{tag}', ns)
            if e is None:
                e = ET.SubElement(root, f'{{{ns['dcterms']}}}{tag}')
            e.text = val
            e.set(f'{{{ns['xsi']}}}type', 'dcterms:W3CDTF')

        tree.write(core_path)

    app_path = 'temp_raw/docProps/app.xml'
    if os.path.exists(app_path):
        tree = ET.parse(app_path)
        root = tree.getroot()
        company_elem = root.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Company')
        if company_elem is None:
            company_elem = ET.SubElement(root, '{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Company')
        company_elem.text = company
        tree.write(app_path)

    with zipfile.ZipFile(output_path, 'w') as zout:
        for folder, _, files in os.walk('temp_raw'):
            for fname in files:
                full = os.path.join(folder, fname)
                arc = os.path.relpath(full, 'temp_raw')
                zout.write(full, arc)

    shutil.rmtree('temp_raw')

async def handle_edit_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    doc = update.message.document
    if not doc.file_name.endswith('.docx'):
        await update.message.reply_text("Пожалуйста, отправьте .docx файл.")
        return WAIT_DOC_FOR_EDIT

    path = f"meta_{doc.file_unique_id}.docx"
    await doc.get_file().download_to_drive(path)
    user_data_store[user_id] = {"file_path": path}

    await update.message.reply_text("Введите имя автора:")
    return ASK_AUTHOR

async def ask_author(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_data_store[update.message.from_user.id]["author"] = update.message.text
    await update.message.reply_text("Введите название организации:")
    return ASK_COMPANY

async def ask_company(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_data_store[update.message.from_user.id]["company"] = update.message.text
    await update.message.reply_text("Введите дату создания (в формате 2025-07-17T10:10:00Z):")
    return ASK_CREATED

async def ask_created(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_data_store[update.message.from_user.id]["created"] = update.message.text
    await update.message.reply_text("Введите дату последнего изменения (в том же формате):")
    return ASK_MODIFIED

async def ask_modified(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_data_store[update.message.from_user.id]["modified"] = update.message.text
    await update.message.reply_text("Введите дату последнего открытия (в том же формате):")
    return ASK_LASTPRINT

async def ask_lastprint(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    user_data_store[user_id]["lastPrinted"] = update.message.text

    in_path = user_data_store[user_id]["file_path"]
    out_path = f"output_{user_id}.docx"

    update_metadata(
        in_path, out_path,
        author=user_data_store[user_id]['author'],
        company=user_data_store[user_id]['company'],
        created=user_data_store[user_id]['created'],
        modified=user_data_store[user_id]['modified'],
        last_print=user_data_store[user_id]['lastPrinted']
    )

    await update.message.reply_document(document=InputFile(out_path))
    os.remove(in_path)
    os.remove(out_path)

    return await restart_menu(update)

async def restart_menu(update: Update):
    keyboard = [[
        InlineKeyboardButton("➕ Удалить ещё один файл", callback_data="clean"),
        InlineKeyboardButton("⚙️ Изменить параметры метаданных", callback_data="edit")
    ]]
    await update.message.reply_text("✅ Готово! Что хотите сделать дальше?", reply_markup=InlineKeyboardMarkup(keyboard))
    return MENU


def main():
    app = ApplicationBuilder().token(TOKEN).build()

    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            MENU: [CallbackQueryHandler(menu_handler)],
            WAIT_DOC_FOR_CLEAN: [MessageHandler(filters.Document.ALL, handle_clean_file)],
            WAIT_DOC_FOR_EDIT: [MessageHandler(filters.Document.ALL, handle_edit_file)],
            ASK_AUTHOR: [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_author)],
            ASK_COMPANY: [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_company)],
            ASK_CREATED: [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_created)],
            ASK_MODIFIED: [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_modified)],
            ASK_LASTPRINT: [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_lastprint)],
        },
        fallbacks=[]
    )

    app.add_handler(conv)
    print("Bot is running...")
    app.run_polling()

if __name__ == "__main__":
    main()
