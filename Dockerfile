FROM python:3.12-slim

# LibreOffice headless para la conversión xlsx -> pdf
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc \
    fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY 2_generar_recibo_service.py app.py
COPY plantilla_maestra_fixed.xlsx .

EXPOSE 8080
CMD ["gunicorn", "-w", "2", "-b", "0.0.0.0:8080", "--timeout", "60", "app:app"]
