FROM debian:bookworm-slim

ENV LANG=es_AR.UTF-8 \
    LC_ALL=es_AR.UTF-8 \
    DEBIAN_FRONTEND=noninteractive \
    PYTHONUNBUFFERED=1 \
    HOME=/tmp \
    PATH="/opt/venv/bin:$PATH"

RUN apt-get update && apt-get install -y --no-install-recommends \
    python3 \
    python3-venv \
    libreoffice-calc \
    libreoffice-core \
    libreoffice-l10n-es \
    python3-uno \
    fonts-dejavu \
    fonts-liberation \
    fonts-noto-core \
    locales \
    && rm -rf /var/lib/apt/lists/* \
    && sed -i -e 's/# es_AR.UTF-8 UTF-8/es_AR.UTF-8 UTF-8/' /etc/locale.gen \
    && dpkg-reconfigure --frontend=noninteractive locales \
    && update-locale LANG=es_AR.UTF-8

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