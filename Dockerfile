FROM python:3.11-slim

# Ngăn chặn prompt tương tác làm đứng tiến trình apt
ENV DEBIAN_FRONTEND=noninteractive
ENV PYTHONUNBUFFERED=1

# Cài đặt bản LibreOffice tối giản cho headless conversion
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy requirements và cài đặt thư viện
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Copy mã nguồn
COPY . .

# Khởi chạy Uvicorn theo biến cổng $PORT động của Render
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-10000}"]