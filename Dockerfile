FROM python:3.11-slim

# Cài Tesseract OCR + gói ngôn ngữ tiếng Việt
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    tesseract-ocr-vie \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Render tự inject biến PORT, code của bạn đã đọc os.environ.get("PORT", 10000)
EXPOSE 10000

CMD ["python", "sepay_bot_with_ocr.py"]
