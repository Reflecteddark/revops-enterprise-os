"""Set column widths for 🎙️ ИИ_Аудит in Google Sheet"""
import gspread
import sys

sys.stdout.reconfigure(encoding="utf-8")

client = gspread.service_account(filename="service_account.json")
sh = client.open_by_key("1QnjrrbpqhYssofchee7G06szWrFvBCcOqGIVokjqdVc")
ws = sh.worksheet("🎙️ ИИ_Аудит")
sheet_id = ws.id

col_widths = {
    0: 110,  # A (ID / Cards / Nav)
    1: 260,  # B (Менеджер / Клиент / Топик)
    2: 130,  # C (Статус / Дата)
    3: 130,  # D (Сделка / Этап)
    4: 110,  # E (Длительность / Роли)
    5: 150,  # F (Критерий / Риск)
    6: 75,   # G (1. Установление контакта)
    7: 75,   # H (2. Квалификация BANT)
    8: 75,   # I (3. Выявление болей)
    9: 75,   # J (4. Презентация ценности)
    10: 75,  # K (5. Отработка возражений)
    11: 75,  # L (6. Фиксация Next Step)
    12: 75,  # M (7. Инициатива в диалоге)
    13: 75,  # N (8. Чистота речи)
    14: 75,  # O (9. Скорость / темп)
    15: 75,  # P (10. Экспертность / тон)
    16: 75,  # Q (11. Внесение в CRM)
    17: 75,  # R (12. Соблюдение регламента)
    18: 85,  # S (Итог баллов)
    19: 105, # T (% Соблюдения)
    20: 135, # U (Грейд звонка)
    21: 175  # V (Ключевая рекомендация)
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
