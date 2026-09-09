FROM python:3.11-slim

# Instalar LibreOffice Writer y fuentes tipográficas básicas de forma optimizada
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copiar e instalar dependencias de Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar todo el proyecto al contenedor
COPY . .

# Comando de arranque para servidor web en producción con Gunicorn
CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:${PORT:-10000} --timeout 120 app:app"]