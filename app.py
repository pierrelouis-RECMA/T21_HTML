import io, os, threading, traceback, uuid
from datetime import datetime, timedelta
import pandas as pd
from flask import Flask, abort, jsonify, render_template_string, request, send_file

# --- IMPORTS DE TES MODULES PERSONNALISÉS ---
# Assure-toi que ces fonctions existent dans tes fichiers .py
from generate_html import build_report_html 
# J'ajoute l'import supposé pour le PPTX
# from generate_pptx import build_agency_pptx 

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024

# Configuration
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "T21_HK_Agencies_Glass_v13.pptx")
_cache = {}
_lock = threading.Lock()

# --- FONCTIONS UTILITAIRES ---

def store_file(data, filename):
    token = str(uuid.uuid4())
    expiry = datetime.now() + timedelta(minutes=10)
    with _lock:
        _cache[token] = {"data": data, "filename": filename, "expiry": expiry}
    return token

def purge_cache():
    now = datetime.now()
    with _lock:
        dead = [k for k, v in _cache.items() if v["expiry"] < now]
        for k in dead:
            del _cache[k]

# --- ROUTES ---

@app.route("/")
def index():
    return render_template_string(HTML)

# Nouvelle route pour générer la version HTML (Rapport Web)
@app.route("/report", methods=["POST"])
def report():
    try:
        file = request.files.get("file")
        if not file:
            return "Fichier manquant", 400
        df = pd.read_excel(file)
        html_content = build_report_html(df)
        return html_content, 200, {"Content-Type": "text/html; charset=utf-8"}
    except Exception as e:
        return f"Erreur lors de la génération HTML : {str(e)}", 500

# Route pour générer le PPTX (utilisée par le bouton "Générer" du HTML)
@app.route("/generate", methods=["POST"])
def generate():
    try:
        file = request.files.get("file")
        if not file:
            return jsonify({"status": "error", "error": "Fichier Excel manquant"}), 400

        df = pd.read_excel(file)

        required = ["Agency", "NewBiz", "Advertiser", "Integrated Spends"]
        missing = [c for c in required if c not in df.columns]
        if missing:
            return jsonify({"status": "error", "error": f"Colonnes manquantes : {', '.join(missing)}"}), 400

        # /!\ ATTENTION : Assure-toi que build_agency_pptx est importé /!\
        # pptx_bytes = build_agency_pptx(df, TEMPLATE_PATH) 
        
        # Simulation pour le test si build_agency_pptx n'est pas chargé :
        return jsonify({"status": "error", "error": "Fonction build_agency_pptx non définie"}), 500

        ts = datetime.now().strftime("%Y%m%d_%H%M")
        filename = f"NBB_Report_{ts}.pptx"
        token = store_file(pptx_bytes, filename)
        purge_cache()

        return jsonify({
            "status": "success",
            "download_url": f"{request.host_url.rstrip('/')}/download/{token}",
            "filename": filename
        })
    except Exception as e:
        app.logger.error(traceback.format_exc())
        return jsonify({"status": "error", "error": str(e)}), 500

@app.route("/download/<token>")
def download(token):
    purge_cache()
    with _lock:
        entry = _cache.get(token)
    if not entry:
        abort(404, "Lien expiré. Régénérez le rapport.")
    return send_file(
        io.BytesIO(entry["data"]),
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        as_attachment=True,
        download_name=entry["filename"],
    )

@app.route("/health")
def health():
    return jsonify({"status": "ok", "template_present": os.path.exists(TEMPLATE_PATH)})

# --- VARIABLE HTML (Contenu de ton script) ---
HTML = r"""...ton code HTML ici..."""

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)