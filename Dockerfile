FROM python:3.11-slim

# Dependencias base del sistema
RUN apt-get update && apt-get install -y \
    wget \
    gnupg \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Instalar Chromium + todas sus dependencias del sistema
RUN playwright install chromium && playwright install-deps chromium

COPY . .

EXPOSE 10000

ENV FLASK_ENV=production
ENV HEADLESS=true

# gunicorn con 1 worker (Playwright no es thread-safe) y timeout largo para scraping
CMD ["gunicorn", "-w", "1", "--timeout", "600", "-b", "0.0.0.0:10000", "app:app"]
