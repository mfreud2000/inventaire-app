import json
from pathlib import Path

APP_DIR = Path(__file__).parent
CONFIG_FILE = APP_DIR / "config.json"
EXPORTS_DIR = APP_DIR / "exports"
CREDENTIALS_FILE = APP_DIR / "credentials.json"
TOKEN_FILE = APP_DIR / "token.json"

COLUMNS = ["CATEGORIE", "REFERENCE", "FOURNISSEUR", "PRODUIT",
           "CONDITIONNEMENT", "PRIX HT", "QTE M-1", "QTE M31", "INVENTAIRE VALORISE"]

COL_WIDTHS = [130, 110, 130, 210, 130, 85, 80, 80, 135]

DEFAULT_CONFIG = {
    "spreadsheet_id": "",
    "sheet_name": "Feuil1",
    "exports_folder": str(EXPORTS_DIR),
}


def load_config():
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return {**DEFAULT_CONFIG, **json.load(f)}
    return DEFAULT_CONFIG.copy()


def save_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2, ensure_ascii=False)
