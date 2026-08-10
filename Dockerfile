FROM python:3.10-slim

# تثبيت LibreOffice
RUN apt-get update && apt-get install -y libreoffice-core libreoffice-writer && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt

COPY . .

CMD ["gunicorn", "--bind", "0.0.0.0:8080", "app:app"]