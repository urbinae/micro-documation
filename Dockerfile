FROM python:3.12-slim

# Instalar LibreOffice headless y paquetes de fuentes estándar para evitar desfases de texto
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc \
    fonts-dejavu \
    fonts-liberation \
    fontconfig \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY 2_generar_recibo_service.py app.py
COPY plantilla_maestra_fixed.xlsx .

EXPOSE 8080
CMD ["gunicorn", "-w", "2", "-b", "0.0.0.0:8080", "--timeout", "90", "app:app"]