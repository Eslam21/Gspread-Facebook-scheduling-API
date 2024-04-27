from typing import Final
from telegram import Update, InputMediaPhoto
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

ALLOWED_USER_ID: Final[int] = 1334659701  # Replace with your user ID

def connect_to_sheet():
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
            "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
    client = gspread.authorize(creds)
    sheet = client.open("NU Confessions (Responses)").sheet1  # Open the spreadsheet

    conf = sheet.col_values(2)  # all confessions
    date_posted = sheet.col_values(6)  # confessions posted only, which have a date
    num_conf = "Number of new confessions: " + str(len(conf) - len(date_posted))
    confs = conf[len(date_posted):]

    print("Processing confessions")

    return num_conf, confs

with open('telegram_keys.txt', 'r') as f:
    user_name = f.readline().strip()
    TOKEN = f.readline().strip()

async def get_confessions_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.from_user.id != ALLOWED_USER_ID:
        await update.message.reply_text("You are not allowed to view this")
    else:
        await update.message.reply_text("Connecting to Sheet")
        num_conf, confs = connect_to_sheet()
        await update.message.reply_text(f"{num_conf}")

        for i, conf in enumerate(confs):
            await update.message.reply_text(f"Confession Number {i + 1}:\n{conf}")
            time.sleep(1.5)

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("عاوز انيييييك")

async def handle_response(text: str, update: Update) -> str:
    
    if 'hello' in text.lower():
        return 'عامل ايه يا باشا'
    
    if 'how are you' in text.lower():
        return 'كله تمام'

    if 'anas' in text.lower():
        with open('anas photo.jpg', 'rb') as photo:
            await update.message.reply_photo(photo=photo)
        return 'هذه الصورة لك'

    if 'who is there' in text.lower():
        return 'your MOTHER'

    return 'مش فاهم منك حاجة يا خرا انت'

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message_type: str = update.message.chat.type
    text: str = update.message.text

    print(f'User ({update.message.chat_id}) in {message_type}: "{text}"')

    if message_type == 'group':
        if user_name in text:
            new_text: str = text.replace(user_name, '').strip()
            response: str = await handle_response(new_text, update)
        else:
            return
    else:
        response: str = await handle_response(text, update)

    print("Bot: ", response)
    await update.message.reply_text(response)

async def error(update: Update, context: ContextTypes.DEFAULT_TYPE):
    print(f"Update {update} caused error {context.error}")

if __name__ == '__main__':
    print("Starting Bot...")

    app = Application.builder().token(TOKEN).build()

    app.add_handler(CommandHandler('confessions', get_confessions_command))
    app.add_handler(CommandHandler('help', help_command))

    app.add_handler(MessageHandler(filters.TEXT, handle_message))

    app.add_error_handler(error)

    app.run_polling(poll_interval=5)
