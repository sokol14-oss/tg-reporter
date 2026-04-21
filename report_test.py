import os
import smtplib
import pandas as pd
from telethon import TelegramClient
from datetime import datetime, timezone, timedelta
from email.message import EmailMessage

# Данные для входа
API_ID = int(os.getenv('API_ID').replace('"', '').strip())
API_HASH = os.getenv('API_HASH').replace('"', '').strip()
client = TelegramClient('session_data', API_ID, API_HASH)

def send_email(file_path):
    msg = EmailMessage()
    msg['Subject'] = f"📊 Отчет по задачам за {datetime.now().date()}"
    msg['From'] = "nastya.sokol2013@yandex.ru"
    msg['To'] = "nastya.sokol2013@yandex.ru"
    msg.set_content("Привет! Во вложении отчет по задачам за сегодня.")

    with open(file_path, 'rb') as f:
        msg.add_attachment(
            f.read(),
            maintype='application',
            subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            filename=file_path
        )

    with smtplib.SMTP_SSL('smtp.yandex.ru', 465) as smtp:
        smtp.login("nastya.sokol2013@yandex.ru", "bzpvhkswafdwacia") 
        smtp.send_message(msg)
    print("📧 Письмо успешно отправлено через Яндекс!")

async def main():
    chat_id = -1003435555222 # Твой ID чата
    topics = {3: "НН", 5: "Самара(Дзержинского)", 2: "СПБ", 6: "Самара(Смышляевка)", 4: "Тула"}
    keywords = ["ITHELP", "SBLEQUIP"]
    
    today = datetime.now(timezone.utc).date()
    file_name = "report_main.xlsx"
    all_tasks = []

    for tid, tname in topics.items():
        print(f"Сканирую {tname}...")
        async for msg in client.iter_messages(chat_id, reply_to=tid, limit=100):
            if msg.date.date() == today and msg.text:
                if any(word in msg.text.upper() for word in keywords):
                    all_tasks.append({
                        'Ветка': tname,
                        'Время': (msg.date + timedelta(hours=3)).strftime('%H:%M'),
                        'Автор': (await msg.get_sender()).first_name if await msg.get_sender() else "N/A",
                        'Задача': msg.text
                    })

    if all_tasks:
        df = pd.DataFrame(all_tasks)
        df = df.sort_values(by=['Время'])
        if os.path.exists(file_name):
            with pd.ExcelWriter(file_name, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                start_row = writer.book.active.max_row
                df.to_excel(writer, index=False, header=False, startrow=start_row)
                print(f"Добавили новые строки в {file_name}")
        else:
            df.to_excel(file_name, index=False)
            print(f"Создали новый файл {file_name}")
        # Вызываем отправку ПОСЛЕ закрытия writer, чтобы файл успел сохраниться
        send_email(file_name)
        print(f"дата выполнения скрипта{today}")
    else:
        print(f"Задач с ключевыми словами {keywords} не найдено.")
        print(f"Дата выполнения скрипта{today}")

with client:
    client.loop.run_until_complete(main())
