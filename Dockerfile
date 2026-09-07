<<<<<<< HEAD
FROM python:3.12-bookworm

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONUNBUFFERED=1 \
    HOME=/tmp

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc \
    libreoffice-core \
    python3-uno \
    fonts-dejavu \
    fonts-liberation \
    fonts-noto-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt
COPY app ./app

EXPOSE 10000
CMD ["sh", "-c", "uvicorn app.main:app --host 0.0.0.0 --port ${PORT:-10000}"]
=======
FROM ubuntu:22.04

ENV DEBIAN_FRONTEND=noninteractive

RUN apt-get update && apt-get install -y \
    libreoffice-calc \
    python3 \
    python3-pip \
    python3-uno \
    fonts-dejavu \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip3 install --no-cache-dir -r requirements.txt

COPY app.py start.sh ./
RUN chmod +x start.sh

EXPOSE 8080

CMD ["./start.sh"]
>>>>>>> fix-paginas-2
