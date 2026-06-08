# -*- coding: utf-8 -*-
"""
Monitor de sincronização — roda em background no PC.
Verifica a cada 2 minutos se o app solicitou uma sincronização.
"""
import time, sys, os, subprocess
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gspread
from google.oauth2.service_account import Credentials

CREDS_PATH = r"G:\Meu Drive\Finanças\Automacao\credenciais_google.json"
SHEET_ID   = "1LR63NFna6y1z88aI1HxRKq8kyNGaPN_3-dLzX3kn5Ys"
SCRIPT     = r"G:\Meu Drive\Finanças\Automacao\sincronizar_sheets.py"
INTERVALO  = 120  # segundos

def get_ws():
    creds = Credentials.from_service_account_file(CREDS_PATH, scopes=[
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ])
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SHEET_ID)
    try:
        return sh.worksheet("Config")
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet("Config", 10, 2)
        ws.update([["Key", "Value"], ["sync_pending", "FALSE"], ["last_sync", "-"]])
        return ws

def checar_e_sincronizar():
    ws   = get_ws()
    data = {r["Key"]: r["Value"] for r in ws.get_all_records() if r.get("Key")}
    if data.get("sync_pending") != "TRUE":
        return

    print(f"[{datetime.now():%H:%M:%S}] Sincronização solicitada — executando...")
    cell = ws.find("sync_pending")
    ws.update_cell(cell.row, 2, "RUNNING")

    result = subprocess.run(
        [sys.executable, SCRIPT],
        capture_output=True, text=True, timeout=180
    )

    agora = datetime.now().strftime("%d/%m/%Y %H:%M")
    ws.update_cell(cell.row, 2, "FALSE")

    cell_sync = ws.find("last_sync")
    ws.update_cell(cell_sync.row, 2, agora)

    if result.returncode == 0:
        print(f"[{datetime.now():%H:%M:%S}] Sync concluído com sucesso.")
    else:
        print(f"[{datetime.now():%H:%M:%S}] Erro no sync: {result.stderr[:200]}")

if __name__ == "__main__":
    print(f"Monitor de sync iniciado — verificando a cada {INTERVALO}s")
    while True:
        try:
            checar_e_sincronizar()
        except Exception as e:
            print(f"[{datetime.now():%H:%M:%S}] Erro: {e}")
        time.sleep(INTERVALO)
