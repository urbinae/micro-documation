FROM debian:bookworm-slim

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONUNBUFFERED=1 \
    HOME=/tmp \
    PATH="/opt/venv/bin:$PATH"

RUN apt-get update && apt-get install -y --no-install-recommends \
    python3 \
    python3-venv \
    libreoffice-calc \
    libreoffice-core \
    python3-uno \
    fonts-dejavu \
    fonts-liberation \
    fonts-noto-core \
    && rm -rf /var/lib/apt/lists/*

# Venv que HEREDA los paquetes del sistema (ahí vive "uno", instalado por
# python3-uno para el Python del sistema). Al usar --system-site-packages
# evitamos el desajuste de versión/ABI que causaba "No module named 'uno'".
RUN python3 -m venv /opt/venv --system-site-packages

WORKDIR /app
COPY requirements.txt ./
RUN /opt/venv/bin/pip install --no-cache-dir -r requirements.txt
COPY app ./app

EXPOSE 10000
CMD ["sh", "-c", "uvicorn app.main:app --host 0.0.0.0 --port ${PORT:-10000}"]