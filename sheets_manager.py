import gspread
import json
import os
from config import CREDENTIALS_FILE, TOKEN_FILE


def _build_client():
    """
    Connexion Google Sheets.
    Priorité :
      1. Variable d'environnement GOOGLE_SERVICE_ACCOUNT (JSON string) → hébergement cloud
      2. Fichier service_account.json local                             → local avec service account
      3. credentials.json OAuth2                                        → local legacy
    """
    sa_env = os.getenv("GOOGLE_SERVICE_ACCOUNT")
    if sa_env:
        info = json.loads(sa_env)
        return gspread.service_account_from_dict(info)

    sa_file = CREDENTIALS_FILE.parent / "service_account.json"
    if sa_file.exists():
        return gspread.service_account(filename=str(sa_file))

    if CREDENTIALS_FILE.exists():
        return gspread.oauth(
            credentials_filename=str(CREDENTIALS_FILE),
            authorized_user_filename=str(TOKEN_FILE),
        )

    raise FileNotFoundError(
        "Aucune credentials Google trouvées.\n"
        "• Cloud : définissez la variable GOOGLE_SERVICE_ACCOUNT\n"
        "• Local  : placez service_account.json ou credentials.json dans le dossier"
    )


class SheetsManager:
    def __init__(self, spreadsheet_id, sheet_name=""):
        self.spreadsheet_id = spreadsheet_id
        self.sheet_name     = sheet_name
        self.gc             = None
        self.spreadsheet    = None
        self.sheet          = None

    def connect(self):
        self.gc          = _build_client()
        self.spreadsheet = self.gc.open_by_key(self.spreadsheet_id)
        if self.sheet_name:
            self.sheet = self.spreadsheet.worksheet(self.sheet_name)
        else:
            self.sheet      = self.spreadsheet.get_worksheet(0)
            self.sheet_name = self.sheet.title

    def list_sheets(self):
        return [ws.title for ws in self.spreadsheet.worksheets()]

    def switch_sheet(self, sheet_name):
        self.sheet_name = sheet_name
        self.sheet      = self.spreadsheet.worksheet(sheet_name)

    def get_all_data(self):
        all_values = self.sheet.get_all_values()
        return all_values[1:] if len(all_values) > 1 else []

    def update_qte_m31(self, row_index, new_qty, new_valorise):
        sheet_row = row_index + 2
        self.sheet.update_cell(sheet_row, 8, new_qty)
        self.sheet.update_cell(sheet_row, 9, new_valorise)

    def batch_update_row(self, row_index, values):
        sheet_row = row_index + 2
        n         = min(len(values), 9)
        col_end   = chr(ord('A') + n - 1)
        self.sheet.update(f'A{sheet_row}:{col_end}{sheet_row}', [values[:n]])

    def close_month(self, data_rows):
        updates = []
        for i, row in enumerate(data_rows):
            sheet_row = i + 2
            padded    = (row + [""] * 9)[:9]
            qte_m31   = padded[7] if padded[7] else "0"
            updates.append({
                'range':  f'G{sheet_row}:I{sheet_row}',
                'values': [[qte_m31, "0", "0"]]
            })
        if updates:
            self.sheet.batch_update(updates)

    def add_row(self, values):
        self.sheet.append_row(values[:9])

    def delete_row(self, row_index):
        self.sheet.delete_rows(row_index + 2)
