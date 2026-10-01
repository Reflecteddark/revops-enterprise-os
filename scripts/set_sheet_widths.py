"""Set column widths for 🎙️ ИИ_Аудит in Google Sheet"""
import gspread
import sys

sys.stdout.reconfigure(encoding="utf-8")

client = gspread.service_account(filename="service_account.json")
sh = client.open_by_key("1QnjrrbpqhYssofchee7G06szWrFvBCcOqGIVokjqdVc")
ws = sh.worksheet("🎙️ ИИ_Аудит")
sheet_id = ws.id

col_widths = {
    0: 90,   # A
    1: 220,  # B
    2: 120,  # C
    3: 130,  # D
    4: 85,   # E
    5: 140,  # F
    6: 75,   # G
    7: 75,   # H
    8: 75,   # I
    9: 75,   # J
    10: 75,  # K
    11: 75,  # L
    12: 75,  # M
    13: 75,  # N
    14: 75,  # O
    15: 75,  # P
    16: 75,  # Q
    17: 75,  # R
    18: 75,  # S
    19: 90,  # T
    20: 120, # U
    21: 140  # V
}

requests = []
for col_idx, width in col_widths.items():
    requests.append({
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "COLUMNS",
                "startIndex": col_idx,
                "endIndex": col_idx + 1
            },
            "properties": {
                "pixelSize": width
            },
            "fields": "pixelSize"
        }
    })

sh.batch_update({"requests": requests})
print("Successfully set column widths in Google Sheet!")
