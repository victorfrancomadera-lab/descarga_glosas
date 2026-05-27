FROM python:3.11-slim

# Dependencias del sistema para Playwright, PyMuPDF, Tesseract y pdfminer
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    tesseract-ocr-spa \
    libglib2.0-0 \
    libnss3 \
    libnspr4 \
    libatk1.0-0 \
    libatk-bridge2.0-0 \
    libcups2 \
    libdrm2 \
    libdbus-1-3 \
    libxcb1 \
    libxkbcommon0 \
    libx11-6 \
    libxcomposite1 \
    libxdamage1 \
    libxext6 \
    libxfixes3 \
    libxrandr2 \
    libgbm1 \
    libpango-1.0-0 \
    libcairo2 \
    libasound2 \
    libatspi2.0-0 \
    wget \
    curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copiar requirements primero (cache de Docker)
COPY requirements.txt .

RUN pip install --no-cache-dir -r requirements.txt

# Instalar Playwright y su browser
RUN playwright install chromium
RUN playwright install-deps chromium

# Copiar el resto del proyecto
COPY . .

# Crear carpetas necesarias
RUN mkdir -p uploads output

EXPOSE 5000

CMD ["python", "app.py"]
