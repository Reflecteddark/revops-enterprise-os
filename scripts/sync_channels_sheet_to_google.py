"""
Скрипт синхронизации листа '📢 Каналы_Продаж' с Google Таблицей:
https://docs.google.com/spreadsheets/d/1AbaQGHhzar7Qsb9U9TByN57rAmOn36zw4HOD2XVrmTc
"""

import sys
from pathlib import Path
import openpyxl
import gspread

sys.stdout.reconfigure(encoding="utf-8")

SPREADSHEET_ID = "1AbaQGHhzar7Qsb9U9TByN57rAmOn36zw4HOD2XVrmTc"
LOCAL_EXCEL = Path(r"C:\Users\strel\Desktop\RevOps_Founder_Workspace_Planner.xlsx")

client = gspread.service_account(filename="service_account.json")
sh = client.open_by_key(SPREADSHEET_ID)
wb = openpyxl.load_workbook(LOCAL_EXCEL, data_only=True)

s_name = "📢 Каналы_Продаж"
ws_local = wb[s_name]

# Читаем данные из Excel
values = []
for r in range(1, ws_local.max_row + 1):
    row_vals = []
    for c in range(1, ws_local.max_column + 1):
        val = ws_local.cell(r, c).value
        row_vals.append("" if val is None else str(val))
    values.append(row_vals)

# Проверяем или создаем лист в Google Sheets
try:
    ws_cloud = sh.worksheet(s_name)
except gspread.WorksheetNotFound:
    ws_cloud = sh.add_worksheet(title=s_name, rows=len(values) + 15, cols=len(values[0]) + 2)
    print(f"[+] Создан новый облачный лист: {s_name}")

sheet_id = ws_cloud.id

# Очищаем и заливаем данные
ws_cloud.clear()
range_str = f"A1:{openpyxl.utils.get_column_letter(len(values[0]))}{len(values)}"
ws_cloud.update(range_name=range_str, values=values)
print(f"[✓] Данные записаны на лист '{s_name}' (строк: {len(values)})")

# Наводим красоту через batch_update
def color_rgb(r, g, b):
    return {"red": r / 255.0, "green": g / 255.0, "blue": b / 255.0}

c_navy_dark = color_rgb(15, 23, 42)
c_navy_hdr = color_rgb(30, 41, 59)
c_blue_primary = color_rgb(30, 58, 138)
c_blue_light = color_rgb(239, 246, 255)
c_card_bg = color_rgb(248, 250, 252)
c_emerald = color_rgb(5, 150, 105)
c_emerald_light = color_rgb(236, 253, 245)
c_amber = color_rgb(217, 119, 6)
c_amber_light = color_rgb(255, 251, 235)
c_purple_light = color_rgb(245, 243, 255)
c_purple = color_rgb(124, 58, 237)
c_border = color_rgb(203, 213, 225)
c_white = color_rgb(255, 255, 255)
c_text_dark = color_rgb(30, 41, 59)
c_text_muted = color_rgb(148, 163, 184)

requests = []

# Сброс объединений
requests.append({
    "unmergeCells": {
        "range": {
            "sheetId": sheet_id,
            "startRowIndex": 0,
            "endRowIndex": 55,
            "startColumnIndex": 0,
            "endColumnIndex": 12
        }
    }
})

def req_merge(r_s, r_e, c_s, c_e):
    return {
        "mergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": r_s,
                "endRowIndex": r_e,
                "startColumnIndex": c_s,
                "endColumnIndex": c_e
            },
            "mergeType": "MERGE_ALL"
        }
    }

def req_format(r_s, r_e, c_s, c_e, bg=None, fg=None, bold=False, size=9, halign="LEFT", valign="MIDDLE", wrap=True):
    cell_fmt = {
        "textFormat": {
            "fontFamily": "Segoe UI",
            "fontSize": int(round(size)),
            "bold": bold
        },
        "horizontalAlignment": halign,
        "verticalAlignment": valign,
        "wrapStrategy": "WRAP" if wrap else "OVERFLOW_CELL"
    }
    if bg:
        cell_fmt["backgroundColor"] = bg
    if fg:
        cell_fmt["textFormat"]["foregroundColor"] = fg

    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": r_s,
                "endRowIndex": r_e,
                "startColumnIndex": c_s,
                "endColumnIndex": c_e
            },
            "cell": {"userEnteredFormat": cell_fmt},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)"
        }
    }

# 1. Заголовки (R1-R2)
requests.append(req_merge(0, 1, 0, 10))
requests.append(req_format(0, 1, 0, 10, bg=c_navy_dark, fg=c_white, bold=True, size=13, halign="LEFT"))
requests.append(req_merge(1, 2, 0, 10))
requests.append(req_format(1, 2, 0, 10, bg=c_navy_dark, fg=c_text_muted, size=9, halign="LEFT"))

# 2. KPI Карточки (R4-R5)
requests.append(req_merge(3, 4, 0, 2))
requests.append(req_merge(3, 4, 2, 4))
requests.append(req_merge(3, 4, 4, 6))
requests.append(req_merge(3, 4, 6, 8))
requests.append(req_merge(3, 4, 8, 10))

requests.append(req_merge(4, 5, 0, 2))
requests.append(req_merge(4, 5, 2, 4))
requests.append(req_merge(4, 5, 4, 6))
requests.append(req_merge(4, 5, 6, 8))
requests.append(req_merge(4, 5, 8, 10))

requests.append(req_format(3, 4, 0, 2, bg=c_blue_light, fg=c_blue_primary, bold=True, size=9, halign="CENTER"))
requests.append(req_format(3, 4, 2, 4, bg=c_purple_light, fg=c_purple, bold=True, size=9, halign="CENTER"))
requests.append(req_format(3, 4, 4, 6, bg=c_blue_light, fg=c_blue_primary, bold=True, size=9, halign="CENTER"))
requests.append(req_format(3, 4, 6, 8, bg=c_amber_light, fg=c_amber, bold=True, size=9, halign="CENTER"))
requests.append(req_format(3, 4, 8, 10, bg=c_emerald_light, fg=c_emerald, bold=True, size=9, halign="CENTER"))

requests.append(req_format(4, 5, 0, 2, bg=c_white, fg=c_blue_primary, bold=True, size=11, halign="CENTER"))
requests.append(req_format(4, 5, 2, 4, bg=c_white, fg=c_purple, bold=True, size=11, halign="CENTER"))
requests.append(req_format(4, 5, 4, 6, bg=c_white, fg=c_blue_primary, bold=True, size=11, halign="CENTER"))
requests.append(req_format(4, 5, 6, 8, bg=c_white, fg=c_amber, bold=True, size=11, halign="CENTER"))
requests.append(req_format(4, 5, 8, 10, bg=c_white, fg=c_emerald, bold=True, size=11, halign="CENTER"))

# 3. Раздел 1 (R7-R28)
requests.append(req_merge(6, 7, 0, 10))
requests.append(req_format(6, 7, 0, 10, bg=c_card_bg, fg=c_text_dark, bold=True, size=11, halign="LEFT"))
# Шапка таблицы каналов (R8)
requests.append(req_format(7, 8, 0, 10, bg=c_blue_primary, fg=c_white, bold=True, size=9, halign="CENTER"))

# Строки каналов (R9-R28)
for r_i in range(8, 28):
    bg_c = c_card_bg if r_i % 2 == 1 else c_white
    requests.append(req_format(r_i, r_i + 1, 0, 10, bg=bg_c, fg=c_text_dark, size=9, halign="LEFT"))
    requests.append(req_format(r_i, r_i + 1, 0, 1, bg=c_blue_light, fg=c_blue_primary, bold=True, size=9, halign="CENTER"))
    requests.append(req_format(r_i, r_i + 1, 1, 2, bg=bg_c, fg=c_text_dark, bold=True, size=9, halign="LEFT"))
    requests.append(req_format(r_i, r_i + 1, 5, 6, bg=bg_c, fg=c_text_dark, bold=True, size=9, halign="CENTER"))
    requests.append(req_format(r_i, r_i + 1, 6, 7, bg=c_amber_light, fg=c_amber, bold=True, size=9, halign="LEFT"))
    requests.append(req_format(r_i, r_i + 1, 8, 9, bg=c_emerald_light, fg=c_emerald, bold=True, size=9, halign="CENTER"))

# 4. Раздел 2: Креативы (R30-R36)
requests.append(req_merge(29, 30, 0, 10))
requests.append(req_format(29, 30, 0, 10, bg=c_purple_light, fg=c_purple, bold=True, size=11, halign="LEFT"))

# Пост 1
requests.append(req_merge(30, 31, 0, 10))
requests.append(req_format(30, 31, 0, 10, bg=c_blue_light, fg=c_blue_primary, bold=True, size=9.5, halign="LEFT"))
requests.append(req_merge(31, 32, 0, 10))
requests.append(req_format(31, 32, 0, 10, bg=c_white, fg=c_navy_dark, size=8.5, halign="LEFT"))

# Пост 2
requests.append(req_merge(32, 33, 0, 10))
requests.append(req_format(32, 33, 0, 10, bg=c_blue_light, fg=c_blue_primary, bold=True, size=9.5, halign="LEFT"))
requests.append(req_merge(33, 34, 0, 10))
requests.append(req_format(33, 34, 0, 10, bg=c_white, fg=c_navy_dark, size=8.5, halign="LEFT"))

# Пост 3
requests.append(req_merge(34, 35, 0, 10))
requests.append(req_format(34, 35, 0, 10, bg=c_blue_light, fg=c_blue_primary, bold=True, size=9.5, halign="LEFT"))
requests.append(req_merge(35, 36, 0, 10))
requests.append(req_format(35, 36, 0, 10, bg=c_white, fg=c_navy_dark, size=8.5, halign="LEFT"))

# 5. Раздел 3: Инструкция (R37-R42)
requests.append(req_merge(37, 38, 0, 10))
requests.append(req_format(37, 38, 0, 10, bg=c_emerald_light, fg=c_emerald, bold=True, size=11, halign="LEFT"))

for r_r in range(38, 42):
    requests.append(req_merge(r_r, r_r + 1, 0, 3))
    requests.append(req_merge(r_r, r_r + 1, 3, 10))
    bg_r = c_card_bg if r_r % 2 == 1 else c_white
    requests.append(req_format(r_r, r_r + 1, 0, 3, bg=bg_r, fg=c_text_dark, bold=True, size=9, halign="LEFT"))
    requests.append(req_format(r_r, r_r + 1, 3, 10, bg=bg_r, fg=c_navy_dark, size=8.5, halign="LEFT"))

# 6. Границы (Borders)
requests.append({
    "updateBorders": {
        "range": {
            "sheetId": sheet_id,
            "startRowIndex": 0,
            "endRowIndex": len(values),
            "startColumnIndex": 0,
            "endColumnIndex": 10
        },
        "top": {"style": "SOLID", "color": c_border},
        "bottom": {"style": "SOLID", "color": c_border},
        "left": {"style": "SOLID", "color": c_border},
        "right": {"style": "SOLID", "color": c_border},
        "innerHorizontal": {"style": "SOLID", "color": c_border},
        "innerVertical": {"style": "SOLID", "color": c_border}
    }
})

# 7. Ширины колонок
col_pixel_widths = [75, 130, 220, 250, 95, 125, 175, 280, 140, 140]
for col_idx, width in enumerate(col_pixel_widths):
    requests.append({
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "COLUMNS",
                "startIndex": col_idx,
                "endIndex": col_idx + 1
            },
            "properties": {"pixelSize": width},
            "fields": "pixelSize"
        }
    })

# 8. Высоты строк
requests.append({
    "updateDimensionProperties": {
        "range": {
            "sheetId": sheet_id,
            "dimension": "ROWS",
            "startIndex": 0,
            "endIndex": 1
        },
        "properties": {"pixelSize": 42},
        "fields": "pixelSize"
    }
})
requests.append({
    "updateDimensionProperties": {
        "range": {
            "sheetId": sheet_id,
            "dimension": "ROWS",
            "startIndex": 7,
            "endIndex": 28
        },
        "properties": {"pixelSize": 36},
        "fields": "pixelSize"
    }
})

# Отправляем форматирование
sh.batch_update({"requests": requests})
print(f"🚀 Облачный лист '{s_name}' успешно оформлен и синхронизирован!")
print(f"URL: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit?gid={sheet_id}#gid={sheet_id}")
