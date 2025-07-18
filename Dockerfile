# Imagen base oficial de Python
FROM python:3.11-slim

# Evita errores con caracteres (español, acentos, etc.)
ENV PYTHONIOENCODING=UTF-8

# Crea el directorio de trabajo dentro del contenedor
WORKDIR /app

# Copia todos los archivos del proyecto al contenedor
COPY . /app

# Instalar dependencias del sistema, incluyendo curl
RUN apt-get update && apt-get install -y curl && apt-get clean
 
# Dar permisos de ejecución al script ping.sh
RUN chmod +x ping.sh

# Instala las dependencias
RUN pip install --no-cache-dir -r requirements.txt

# Expone el puerto de Streamlit
EXPOSE 8501

# Comando que se ejecuta cuando arranca el contenedor
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0", "--server.enableCORS=false", "--server.enableXsrfProtection=false"]