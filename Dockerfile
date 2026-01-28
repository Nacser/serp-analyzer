FROM python:3.11-slim

# Instalar dependencias del sistema para Playwright
RUN apt-get update && apt-get install -y \
    wget \
    gnupg \
    && rm -rf /var/lib/apt/lists/*

# Crear directorio de trabajo
WORKDIR /app

# Copiar requirements primero (para cache de Docker)
COPY requirements.txt .

# Instalar dependencias Python
RUN pip install --no-cache-dir -r requirements.txt

# Instalar Playwright y navegador Chromium
RUN playwright install chromium
RUN playwright install-deps chromium

# Copiar el resto del código
COPY . .

# Puerto que usa Render
EXPOSE 10000

# Variable de entorno para producción
ENV FLASK_ENV=production

# Comando para iniciar la app
CMD ["python", "app.py"]
