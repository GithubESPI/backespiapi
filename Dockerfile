FROM python:3.11

WORKDIR /code

# Installer LibreOffice pour la conversion DOCX vers PDF
RUN apt-get update && apt-get install -y libreoffice

# Copier les dépendances et installer
COPY ./requirements.txt /code/requirements.txt
RUN pip install --no-cache-dir --upgrade -r /code/requirements.txt

# Créer les répertoires nécessaires
RUN mkdir -p /code/documents /code/downloads /code/excel /code/template /code/json

# Copier les fichiers nécessaires dans les répertoires adéquats
COPY ./app /code/app
COPY ./excel /code/excel  
COPY ./template /code/template  
COPY ./json /code/json

# Copie du fichier .env (si nécessaire)
COPY ./.env /code/.env

# Définir les permissions si nécessaire
RUN chmod -R 755 /code

# Lancer l'application
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8080"]
