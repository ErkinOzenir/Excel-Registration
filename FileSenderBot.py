import os
from io import BytesIO
import pandas as pd
from telegram.ext import Updater, CommandHandler, MessageHandler, Filters
from telegram import ChatAction

api_token = None
stored_excel_file = None
stored_excel_filename = None

def start(update, context):
    update.message.reply_text(
        "Welcome to the Excel File Manager. Send an Excel file or ask for one using /get_excel."
    )

def get_excel(update, context):
    global stored_excel_file, stored_excel_filename

    if stored_excel_file and stored_excel_filename:
        update.message.reply_document(document=stored_excel_file, filename=stored_excel_filename)
    else:
        update.message.reply_text("No Excel file available. Please send an Excel file first.")

def handle_excel_file(update, context):
    global stored_excel_file, stored_excel_filename

    if update.message.document:
        file = context.bot.getFile(update.message.document.file_id)
        file_bytes = file.download_as_bytearray()

        try:
            excel_data = pd.read_excel(BytesIO(file_bytes))
            stored_excel_file = BytesIO(file_bytes)
            stored_excel_filename = update.message.document.file_name
            update.message.reply_text("New Excel file received and stored.")
        except pd.errors.ParserError:
            update.message.reply_text("The file you sent is not a valid Excel file.")
    else:
        update.message.reply_text("Please send an Excel file.")

def main():
    updater = Updater(api_token, use_context=True)

    dp = updater.dispatcher

    dp.add_handler(CommandHandler("start", start))
    dp.add_handler(CommandHandler("get_excel", get_excel))

    dp.add_handler(MessageHandler(Filters.document, handle_excel_file))

    updater.start_polling()
    updater.idle()

if __name__ == '__main__':
    main()
