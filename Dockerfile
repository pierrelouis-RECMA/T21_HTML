# Utilisation d'une image Python stable
FROM python:3.12-slim

WORKDIR /app

# Installation des dépendances
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copie uniquement les fichiers nécessaires au rapport HTML
COPY app.py .
COPY generate_html.py .
COPY fill_template.py .

# Exposition du port pour Render
EXPOSE 10000

# Commande de démarrage avec Gunicorn
CMD ["gunicorn", "--bind", "0.0.0.0:10000", "app:app"]
