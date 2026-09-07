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