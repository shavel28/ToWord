FROM python:3.12-slim

# Tesseract OCR + bahasa Indonesia & Inggris
RUN apt-get update && apt-get install -y --no-install-recommends \
        tesseract-ocr tesseract-ocr-ind tesseract-ocr-eng libgl1 libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

ENV PORT=8000 OMP_THREAD_LIMIT=2
EXPOSE 8000
# 1 proses (antrean pekerjaan disimpan di memori), banyak thread
CMD gunicorn -w 1 --threads 8 --timeout 600 -b 0.0.0.0:${PORT} app:app
