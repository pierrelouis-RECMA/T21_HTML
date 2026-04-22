# Utilise une image Python légère
FROM python:3.12-slim

# Définit le dossier de travail
WORKDIR /app

# Copie d'abord les dépendances pour optimiser le cache
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# COPIE UNIQUEMENT LES FICHIERS NÉCESSAIRES
# On ne liste plus generate_pptx_v3.py ici
COPY app.py .
COPY generate_html.py .
COPY fill_template.py .

# Si tu as d'autres fichiers indispensables, ajoute-les, 
# mais retire toute référence à generate_pptx_v3.py

# Commande de démarrage
CMD ["gunicorn", "--bind", "0.0.0.0:10000", "app:app"]
