# ── Dockerfile — Imagen base ─────────────────────────────────
FROM python:3.11-slim

LABEL maintainer="tu-email@ejemplo.com"
LABEL description="Consulta Miembro de Mesa ONPE — Lab 04"

# Instalar Chrome y dependencias para Selenium
RUN apt-get update && apt-get install -y --no-install-recommends \
    wget curl gnupg unzip ca-certificates \
    chromium chromium-driver \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Dependencias Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Código de la aplicación y datos
COPY app.py .
COPY dnis.xlsx .

# Variable de entorno para Selenium (usar Chromium del sistema)
ENV CHROME_BIN=/usr/bin/chromium
ENV CHROMEDRIVER_PATH=/usr/bin/chromedriver
ENV EXCEL_PATH=/app/dnis.xlsx

EXPOSE 5000

CMD ["python", "app.py"]
