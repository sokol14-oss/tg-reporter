import os
import smtplib
import pandas as pd
from telethon import TelegramClient
from datetime import datetime, timezone, timedelta
from email.message import EmailMessage

# Данные для входа
EMAIL_PASS = os.getenv('EMAIL_PASSWORD').replace('"', '').strip()
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
        smtp.login("nastya.sokol2013@yandex.ru", EMAIL_PASS) 
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
                        'Дата': msg.date.strftime('%d.%m.%Y'),
                        'Время': (msg.date + timedelta(hours=3)).strftime('%H:%M'),
                        'Автор': (await msg.get_sender()).first_name if await msg.get_sender() else "N/A",
                        'Задача': msg.text
                    })

    if all_tasks:
        df = pd.DataFrame(all_tasks)
        df = df.sort_values(by=['Дата','Время'])
        if os.path.exists(file_name):
                df_old=pd.read_excel(file_name)
                df_combined = pd.concat([df_old, df], ignore_index=True)
                df_final = df_combined.drop_duplicates(subset=['Ветка','Дата', 'Время', 'Автор', 'Задача'], keep='first')
                df_final.to_excel(file_name, index=False)
                print(f"Добавили новые строки в {file_name}")
        else:
            df.to_excel(file_name, index=False)
            df_final=df
            print(f"Создали новый файл {file_name}")
        # Вызываем отправку ПОСЛЕ закрытия writer, чтобы файл успел сохраниться
        send_email(file_name)
        print(f"дата выполнения скрипта{today}")
    else:
        print(f"Задач с ключевыми словами {keywords} не найдено.")
        print(f"Дата выполнения скрипта{today}")

with client:
    client.loop.run_until_complete(main())
