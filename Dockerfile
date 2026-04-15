FROM python:3.11-slim

# Install only LibreOffice Calc (not the full suite — much faster & smaller)
RUN apt-get update && apt-get install -y \
    libreoffice-calc \
    --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 10000
# Increase timeout to 120s for PDF conversion; preload app for faster cold starts
CMD ["gunicorn", "--bind", "0.0.0.0:10000", "--timeout", "120", "--preload", "app:app"]
