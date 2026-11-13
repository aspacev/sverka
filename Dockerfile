FROM python:3.13-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    SVERKA_DATA=/data \
    HOST=0.0.0.0 \
    PORT=5000 \
    TZ=Asia/Baku

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py db.py service.py compare_excel.py ./
COPY templates ./templates
COPY static ./static

# База, загруженные файлы и Excel-отчёты лежат здесь; папку подключают снаружи
VOLUME /data
EXPOSE 5000

CMD ["python", "app.py"]
