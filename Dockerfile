FROM python:3.9-slim

WORKDIR /app

# Копируем список библиотек и устанавливаем их
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Копируем весь наш код в контейнер
COPY . .

# Команда для запуска (пока оставим так)
CMD ["python", "report.py"]
