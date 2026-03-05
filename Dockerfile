FROM python:3.11-slim

# Устанавливаем рабочую директорию
WORKDIR /app

# Копируем список зависимостей и устанавливаем их
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Устанавливаем часовой пояс (опционально, чтобы время в CSV было локальным)
ENV TZ=Europe/Moscow