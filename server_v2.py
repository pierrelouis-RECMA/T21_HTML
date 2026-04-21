from flask import Flask, request, render_template_string, send_file
import pandas as pd
import io, os
from generate_pptx_v2 import build_agency_pptx

app = Flask(__name__)

HTML_INTERFACE = """
<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>RECMA NBB Generator</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; background: #2D5C54; color: white; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
        .container { background: rgba(255, 255, 255, 0.1); backdrop-filter: blur(10px); padding: 40px; border-radius: 20px; border: 1px solid rgba(255,255,255,0.2); text-align: center; width: 400px; }
        h1 { margin-bottom: 10px; font-weight: 300; }
        p { font-size: 0.9em; opacity: 0.8; margin-bottom: 30px; }
        input[type="file"] { margin: 20px 0; display: block; width: 100%; color: white; }
        button { background: #CC2229; color: white; border: none; padding: 12px 30px; border-radius: 5px; cursor: pointer; font-size: 16px; transition: 0.3s; width: 100%; }
        button:hover { background: #a81c22; transform: scale(1.02); }
    </style>
</head>
<body>
    <div class="container">
        <h1>NBB Generator</h1>
        <p>Génération Slides 3, 4, 6 et Détails (7+)</p>
        <form action="/upload" method="post" enctype="multipart/form-data">
            <input type="file" name="file" accept=".xlsx" required>
            <button type="submit">GÉNÉRER LE POWERPOINT</button>
        </form>
    </div>
</body>
</html>
"""

@app.route('/')
def home():
    return render_template_string(HTML_INTERFACE)

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files.get('file')
    if not file: return "Fichier manquant", 400
    try:
        df = pd.read_excel(file)
        template = "T21_HK_Agencies_Glass_v12.pptx"
        if not os.path.exists(template):
            return f"Erreur: Template {template} manquant sur GitHub.", 500

        pptx_bytes = build_agency_pptx(df, template)
        return send_file(
            io.BytesIO(pptx_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name="NBB_Final_Report.pptx"
        )
    except Exception as e:
        return f"Erreur : {str(e)}", 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)