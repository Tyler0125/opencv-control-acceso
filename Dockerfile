FROM python:3.13-slim

# Dependencias del sistema requeridas por OpenCV
RUN apt-get update && apt-get install -y \
    libgl1 \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Instalar dependencias Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar código fuente
COPY . .

# Cloud Run inyecta la variable PORT automáticamente
ENV PORT=8080
EXPOSE 8080

# Gunicorn para producción
CMD exec gunicorn --bind :$PORT --workers 1 --threads 8 --timeout 60 server:app
