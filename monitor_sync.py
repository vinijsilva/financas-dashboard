# -*- coding: utf-8 -*-
"""
Monitor — roda a cada 2 minutos via Agendador do Windows.
Verifica se o app solicitou sincronizacao ou importacao e executa.
"""
import sys, os, subprocess
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gspread
from google.oauth2.service_account import Credentials

CREDS_PATH   = r"G:\Meu Drive\Finanças\Automacao\credenciais_google.json"
SHEET_ID     = "1LR63NFna6y1z88aI1HxRKq8kyNGaPN_3-dLzX3kn5Ys"
SCRIPT_SYNC  = r"G:\Meu Drive\Finanças\Automacao\sincronizar_sheets.py"
SCRIPT_IMP   = r"G:\Meu Drive\Finanças\Automacao\teste_completo.py"

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
        ws.update([["Key", "Value"],
                   ["sync_pending",   "FALSE"], ["last_sync",   "-"],
                   ["import_pending", "FALSE"], ["last_import", "-"]])
        return ws

def _executar(ws, flag_key, resultado_key, script, timeout=300):
    cell_flag = ws.find(flag_key)
    ws.update_cell(cell_flag.row, 2, "RUNNING")

    result = subprocess.run(
        [sys.executable, script],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=timeout
    )

    agora = datetime.now().strftime("%d/%m/%Y %H:%M")

    # Extrai resumo do output (ultima linha nao vazia)
    linhas = [l.strip() for l in result.stdout.splitlines() if l.strip()]
    resumo = linhas[-1] if linhas else ("OK" if result.returncode == 0 else "Erro")
    valor_resultado = f"{agora} — {resumo}"

    ws.update_cell(cell_flag.row, 2, "FALSE")
    cell_res = ws.find(resultado_key)
    ws.update_cell(cell_res.row, 2, valor_resultado)

    return result.returncode == 0

def checar_e_importar(ws, data):
    if data.get("import_pending") != "TRUE":
        return
    print(f"[{datetime.now():%H:%M:%S}] Importacao solicitada — executando...")
    ok = _executar(ws, "import_pending", "last_import", SCRIPT_IMP, timeout=300)
    print(f"[{datetime.now():%H:%M:%S}] Importacao {'concluida' if ok else 'com erro'}.")

def checar_e_sincronizar(ws, data):
    if data.get("sync_pending") != "TRUE":
        return
    print(f"[{datetime.now():%H:%M:%S}] Sincronizacao solicitada — executando...")
    ok = _executar(ws, "sync_pending", "last_sync", SCRIPT_SYNC, timeout=180)
    print(f"[{datetime.now():%H:%M:%S}] Sync {'concluido' if ok else 'com erro'}.")

if __name__ == "__main__":
    # Executa uma verificacao e encerra — o agendador do Windows chama a cada 2 min
    try:
        ws   = get_ws()
        data = {r["Key"]: r["Value"] for r in ws.get_all_records() if r.get("Key")}

        # Garante que as chaves de import existem na aba Config
        if "import_pending" not in data:
            ws.append_row(["import_pending", "FALSE"])
            ws.append_row(["last_import", "-"])
            data["import_pending"] = "FALSE"

        checar_e_importar(ws, data)
        checar_e_sincronizar(ws, data)
    except Exception as e:
        print(f"[{datetime.now():%H:%M:%S}] Erro: {e}")
