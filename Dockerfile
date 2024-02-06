# Utilisez l'image de base Python spécifiée dans votre fichier DevContainer.json
FROM mcr.microsoft.com/devcontainers/python:0-3.11

# Mettez à jour les paquets et installez LibreOffice
RUN apt-get update && apt-get install -y libreoffice

# Définissez le répertoire de travail dans le conteneur
WORKDIR /app

# Copiez les fichiers de dépendance dans le conteneur
COPY requirements.txt .

# Installez les dépendances
RUN pip3 install --no-cache-dir -r requirements.txt

# Créez les dossiers nécessaires pour les fichiers pptx et pdf
RUN mkdir -p files/pptx files/pdf

# Copiez le code source dans le conteneur
COPY . .

# Exposez le port sur lequel votre application s'exécute (par exemple, 9000 pour votre application)
EXPOSE 9000

# Définissez la commande pour exécuter votre application
CMD ["flask", "run", "--host=0.0.0.0"]
