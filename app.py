from flask import Flask, jsonify, request, send_file, render_template
from flask_cors import CORS
from dotenv import load_dotenv
from config import load_config, save_config
from sheets_manager import SheetsManager
from export_manager import export_to_excel, export_all_to_excel
from datetime import datetime
from pathlib import Path
from io import BytesIO
import os
import logging
import tempfile

# ── Chargement des variables d'environnement (.env) ──────────────────
load_dotenv(Path(__file__).parent / ".env")

# ── Logging structuré ─────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger(__name__)

# ── App Flask + CORS ──────────────────────────────────────────────────
app = Flask(__name__)
CORS(app)   # autorise l'accès depuis n'importe quelle origine (Android, PC...)

sheets_mgr  = None
all_data    = []
config_data = load_config()

# Surcharger avec .env si présent
if os.getenv("SPREADSHEET_ID"):
    config_data["spreadsheet_id"] = os.getenv("SPREADSHEET_ID")


def _safe_float(val):
    try:
        return float(str(val).replace(",", ".").replace(" ", ""))
    except (ValueError, TypeError):
        return 0.0


# ------------------------------------------------------------------ Pages ---

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/manifest.json")
def manifest():
    return send_file("static/manifest.json", mimetype="application/manifest+json")


@app.route("/sw.js")
def service_worker():
    resp = send_file("static/sw.js", mimetype="application/javascript")
    resp.headers["Service-Worker-Allowed"] = "/"
    resp.headers["Cache-Control"] = "no-cache"
    return resp


# ----------------------------------------------------------------- Config ---

@app.route("/api/config", methods=["GET"])
def get_config():
    return jsonify({
        "spreadsheet_id": config_data.get("spreadsheet_id", ""),
        "sheet_name": config_data.get("sheet_name", ""),
        "exports_folder": config_data.get("exports_folder", "exports"),
    })


@app.route("/api/config", methods=["POST"])
def post_config():
    global config_data, sheets_mgr
    data = request.json or {}
    config_data.update({k: v for k, v in data.items() if k in ("spreadsheet_id", "sheet_name", "exports_folder")})
    save_config(config_data)
    sheets_mgr = None
    return jsonify({"ok": True})


# ------------------------------------------------------------------- Sync ---

@app.route("/api/sync", methods=["POST"])
def sync():
    global sheets_mgr, all_data
    if not config_data.get("spreadsheet_id"):
        return jsonify({"error": "ID Google Sheet non configuré"}), 400
    try:
        logger.info("Sync demandé")
        if not sheets_mgr:
            sheets_mgr = SheetsManager(
                config_data["spreadsheet_id"],
                config_data.get("sheet_name", "")
            )
        sheets_mgr.connect()
        sheets = sheets_mgr.list_sheets()
        all_data = sheets_mgr.get_all_data()
        return jsonify({
            "ok": True,
            "data": all_data,
            "sheets": sheets,
            "current_sheet": sheets_mgr.sheet_name,
        })
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/switch_sheet", methods=["POST"])
def switch_sheet():
    global all_data
    if not sheets_mgr:
        return jsonify({"error": "Non connecté"}), 400
    sheet_name = (request.json or {}).get("sheet_name")
    if not sheet_name:
        return jsonify({"error": "Nom d'onglet manquant"}), 400
    try:
        sheets_mgr.switch_sheet(sheet_name)
        all_data = sheets_mgr.get_all_data()
        return jsonify({"ok": True, "data": all_data, "current_sheet": sheet_name})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ------------------------------------------------------------------- Data ---

@app.route("/api/data", methods=["GET"])
def get_data():
    return jsonify({"data": all_data})


# ------------------------------------------------------------------- CRUD ---

@app.route("/api/update_qty", methods=["POST"])
def update_qty():
    global all_data
    body = request.json or {}
    idx = body.get("index")
    new_qty = body.get("qty", "0")
    if idx is None:
        return jsonify({"error": "Index manquant"}), 400
    try:
        qty = float(str(new_qty).replace(",", "."))
        row = list((all_data[idx] + [""] * 9)[:9])
        prix = _safe_float(row[5])
        valorise = round(qty * prix, 2)
        row[7] = str(qty)
        row[8] = str(valorise)
        all_data[idx] = row
        if sheets_mgr:
            sheets_mgr.update_qte_m31(idx, str(qty), valorise)
        return jsonify({"ok": True, "row": row})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/update_item", methods=["POST"])
def update_item():
    global all_data
    body = request.json or {}
    idx = body.get("index")
    values = body.get("values", [])
    if idx is None:
        return jsonify({"error": "Index manquant"}), 400
    try:
        all_data[idx] = values
        if sheets_mgr:
            sheets_mgr.batch_update_row(idx, values)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/add_item", methods=["POST"])
def add_item():
    global all_data
    values = (request.json or {}).get("values", [])
    try:
        all_data.append(values)
        if sheets_mgr:
            sheets_mgr.add_row(values)
        return jsonify({"ok": True, "index": len(all_data) - 1})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/hide_item", methods=["POST"])
def hide_item():
    global all_data
    idx = (request.json or {}).get("index")
    if idx is None:
        return jsonify({"error": "Index manquant"}), 400
    try:
        row = list((all_data[idx] + [""] * 9)[:9])
        if not row[0].startswith("[MASQUÉ]"):
            row[0] = f"[MASQUÉ] {row[0]}".strip()
        all_data[idx] = row
        if sheets_mgr:
            sheets_mgr.batch_update_row(idx, row)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/bulk_assign_category", methods=["POST"])
def bulk_assign_category():
    global all_data
    body     = request.json or {}
    indices  = body.get("indices", [])
    category = body.get("category", "").strip()
    if not category:
        return jsonify({"error": "Catégorie vide"}), 400
    try:
        updates = []
        for idx in indices:
            row = list((all_data[idx] + [""] * 9)[:9])
            row[0] = category
            all_data[idx] = row
            updates.append((idx, row))
        if sheets_mgr:
            for idx, row in updates:
                sheets_mgr.batch_update_row(idx, row)
        return jsonify({"ok": True, "updated": len(updates)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/restore_item", methods=["POST"])
def restore_item():
    global all_data
    idx = (request.json or {}).get("index")
    if idx is None:
        return jsonify({"error": "Index manquant"}), 400
    try:
        row = list((all_data[idx] + [""] * 9)[:9])
        row[0] = row[0].replace("[MASQUÉ] ", "").replace("[MASQUÉ]", "").strip()
        all_data[idx] = row
        if sheets_mgr:
            sheets_mgr.batch_update_row(idx, row)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ----------------------------------------------------------------- Export ---

def _excel_response(path, download_name):
    """Envoie le fichier Excel puis le supprime (compatible cloud sans disque persistant)."""
    with open(path, "rb") as f:
        data = f.read()
    try:
        os.remove(path)
    except Exception:
        pass
    return send_file(
        BytesIO(data),
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/api/export_excel", methods=["GET"])
def export_excel():
    if not all_data:
        return jsonify({"error": "Aucune donnée — synchronisez d'abord"}), 400
    try:
        label      = datetime.now().strftime("%Y_%m")
        sheet_name = sheets_mgr.sheet_name if sheets_mgr else "Inventaire"
        folder     = tempfile.mkdtemp()
        path       = export_to_excel(all_data, folder, label, sheet_name)
        name       = f"Inventaire_{sheet_name}_{label}.xlsx"
        logger.info(f"Export Excel onglet '{sheet_name}'")
        return _excel_response(path, name)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/export_all_excel", methods=["GET"])
def export_all_excel():
    if not sheets_mgr:
        return jsonify({"error": "Non connecté — synchronisez d'abord"}), 400
    SKIP = {"RELEVES COMPTEURS", "Règles d'usage"}
    try:
        label       = datetime.now().strftime("%Y_%m")
        sheets_data = []
        current     = sheets_mgr.sheet_name
        for name in sheets_mgr.list_sheets():
            if name in SKIP:
                continue
            sheets_mgr.switch_sheet(name)
            sheets_data.append((name, sheets_mgr.get_all_data()))
        sheets_mgr.switch_sheet(current)
        folder = tempfile.mkdtemp()
        path   = export_all_to_excel(sheets_data, folder, label)
        logger.info(f"Export Excel complet — {len(sheets_data)} onglets")
        return _excel_response(path, f"Inventaire_COMPLET_{label}.xlsx")
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/close_month", methods=["POST"])
def close_month():
    global all_data
    if not all_data:
        return jsonify({"error": "Aucune donnée — synchronisez d'abord"}), 400
    try:
        label = datetime.now().strftime("%Y_%m")
        path = export_to_excel(all_data, config_data.get("exports_folder", "exports"), label)
        if sheets_mgr:
            sheets_mgr.close_month(all_data)
            all_data = sheets_mgr.get_all_data()
        return jsonify({
            "ok": True,
            "filename": os.path.basename(path),
            "data": all_data,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


import atexit

def _shutdown():
    global sheets_mgr
    if sheets_mgr and sheets_mgr.gc:
        logger.info("Fermeture connexion Google Sheets")
        sheets_mgr = None

atexit.register(_shutdown)

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    logger.info(f"Démarrage InventaireApp sur port {port}")
    app.run(host="0.0.0.0", port=port, debug=False)
