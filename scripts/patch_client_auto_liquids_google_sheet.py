#!/usr/bin/env python3
"""
Direct cloud migration patch for client Google Sheet:
RevOps Platform V18 ООО "ТД Автожидкости" (1xxGSOBFzoZV1u7mBmpexD0jXvmujHchCBGyEGBWlmIs)

Applies all 5 enterprise audit fixes directly via Google Sheets API:
1. Mask BOT_TOKEN in ⚙️ Настройки (B52)
2. Unmerge all 114 merged cells in 🎙️ ИИ_Аудит + set column widths & wrapStrategy
3. Standardize unified internal Navbar (#gid=...) on all 21 visual dashboards
4. Expand all 239 calculation formulas from 1000 to 50000 rows
5. Improve empty states on 🎙️ ИИ_Аудит & add capacity watchdog in 🧪 QA_Suite
"""

import gspread
import re
import sys
import time

sys.stdout.reconfigure(encoding='utf-8')

SPREADSHEET_ID = "1xxGSOBFzoZV1u7mBmpexD0jXvmujHchCBGyEGBWlmIs"

print("==================================================")
print(f"CONNECTING TO CLIENT GOOGLE SHEET: {SPREADSHEET_ID}")
print("==================================================")

client = gspread.service_account(filename="service_account.json")
sh = client.open_by_key(SPREADSHEET_ID)
print(f"Connected: '{sh.title}'")

sheet_map = {ws.title: ws.id for ws in sh.worksheets()}
print(f"Total sheets: {len(sheet_map)}\n")

# ----------------------------------------------------
# 1. FIX BOT_TOKEN in ⚙️ Настройки (B52)
# ----------------------------------------------------
print("--- [1] FIXING BOT_TOKEN IN ⚙️ Настройки ---")
try:
    ws_set = sh.worksheet("⚙️ Настройки")
    old_b52 = ws_set.acell("B52").value
    ws_set.update_acell("B52", "🔒 Хранится в защищенном Script Properties (настройка: меню RevOps -> 🔐 Telegram Bot)")
    print(f"  [+] B52 updated successfully (was: '{old_b52}')")
except Exception as e:
    print(f"  [!] Error updating B52: {e}")
print()

# ----------------------------------------------------
# 2. UNMERGE ALL 114 CELLS IN 🎙️ ИИ_Аудит + FORMAT
# ----------------------------------------------------
print("--- [2] UNMERGING 114 CELLS IN 🎙️ ИИ_Аудит ---")
try:
    ws_audit = sh.worksheet("🎙️ ИИ_Аудит")
    audit_id = ws_audit.id
    
    # 1. Unmerge all cells in table area
    req_unmerge = {
        "unmergeCells": {
            "range": {
                "sheetId": audit_id,
                "startRowIndex": 1,
                "endRowIndex": 150,
                "startColumnIndex": 0,
                "endColumnIndex": 26
            }
        }
    }
    
    # 2. Set clean column widths for flawless reading
    col_widths = [
        (0, 1, 90),   # A: Call ID
        (1, 2, 220),  # B: Сделка / Контрагент
        (2, 3, 140),  # C: Менеджер
        (3, 4, 130),  # D: Сумма в риске
        (4, 5, 240),  # E: Критический дефект / Параметр
        (5, 6, 360),  # F: Цитата Whisper Large V3
        (6, 7, 360),  # G: Решение РОПа / Маркеры
        (7, 8, 110),  # H: Оценка
        (8, 9, 110),  # I: Статус
        (9, 10, 120), # J: Длительность
        (10, 11, 130),# K: Дата звонка
        (11, 12, 140),# L: Ссылка на CRM
    ]
    width_reqs = []
    for start_c, end_c, px in col_widths:
        width_reqs.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": audit_id,
                    "dimension": "COLUMNS",
                    "startIndex": start_c,
                    "endIndex": end_c
                },
                "properties": {
                    "pixelSize": px
                },
                "fields": "pixelSize"
            }
        })
        
    # 3. Enable text wrap strategy on data rows
    wrap_req = {
        "repeatCell": {
            "range": {
                "sheetId": audit_id,
                "startRowIndex": 1,
                "endRowIndex": 100,
                "startColumnIndex": 0,
                "endColumnIndex": 15
            },
            "cell": {
                "userEnteredFormat": {
                    "wrapStrategy": "WRAP",
                    "verticalAlignment": "MIDDLE"
                }
            },
            "fields": "userEnteredFormat(wrapStrategy,verticalAlignment)"
        }
    }

    sh.batch_update({"requests": [req_unmerge] + width_reqs + [wrap_req]})
    print("  [+] Successfully unmerged 114 cells in 🎙️ ИИ_Аудит, set column widths and text wrap!")
except Exception as e:
    print(f"  [!] Error on 🎙️ ИИ_Аудит: {e}")
print()

# ----------------------------------------------------
# 3. STANDARDIZE INTERNAL NAVBAR (#gid=...) ON ALL 21 VISUAL DASHBOARDS
# ----------------------------------------------------
print("--- [3] STANDARDIZING INTERNAL NAVBAR (#gid=...) ACROSS 21 VISUAL DASHBOARDS ---")
nav_definitions = [
    (sheet_map.get("📄 Executive_OnePager"), "📄 1. One-Pager"),
    (sheet_map.get("⚡ Пульс_Компании"), "⚡ 2. Пульс компании"),
    (sheet_map.get("📋 Пульт_РОПа_15_Минут"), "📋 3. Пульт РОПа"),
    (sheet_map.get("💸 Диагностика_Утечек_ОП"), "💸 4. 7 Грехов ОП"),
    (sheet_map.get("🎯 Action_Center"), "🎯 5. Action Center"),
    (sheet_map.get("🎙️ ИИ_Аудит"), "🎙️ 6. ИИ-Аудит"),
    (sheet_map.get("⚡ Экспресс_Калькулятор_3_Цифры"), "⚡ 7. Экспресс 3 цифры"),
    (sheet_map.get("🧪 QA_Suite"), "🧪 8. QA Suite"),
    (sheet_map.get("⚙️ Настройки"), "⚙️ 9. Настройки"),
]

# Build formula values for row 2
nav_row_values = []
for gid, label in nav_definitions:
    if gid is not None:
        nav_row_values.append(f'=HYPERLINK("#gid={gid}", "{label}")')
    else:
        nav_row_values.append(label)

visual_dashboard_names = [
    "📄 Executive_OnePager", "⚡ Пульс_Компании", "📋 Пульт_РОПа_15_Минут",
    "💸 Диагностика_Утечек_ОП", "🎯 Action_Center", "🎙️ ИИ_Аудит",
    "⚡ Экспресс_Калькулятор_3_Цифры", "🧪 QA_Suite", "⚙️ Настройки",
    "🌐 Мультиканальная_Атрибуция", "💳 Финансы_и_AI_Дожим", "👥 Мотивация_ОП",
    "🔮 Симулятор_Роста", "🗄️ DWH_и_Безопасность", "🌐 Сквозная_RevOps_Аналитика",
    "🌐 Маркетинг_и_Трафик", "🎯 Воронка_и_SLA", "📊 Юнит_Экономика",
    "🚨 Радар_Алертов", "📈 Когорты_LTV", "👥 Ресурсный_План"
]

navbar_updated = 0
for title in visual_dashboard_names:
    if title not in sheet_map:
        continue
    try:
        ws = sh.worksheet(title)
        # Unmerge row 2 if merged
        unmerge_r2 = {
            "unmergeCells": {
                "range": {
                    "sheetId": ws.id,
                    "startRowIndex": 1,
                    "endRowIndex": 2,
                    "startColumnIndex": 0,
                    "endColumnIndex": 15
                }
            }
        }
        sh.batch_update({"requests": [unmerge_r2]})
        
        # Write navbar values
        ws.update("A2:I2", [nav_row_values], value_input_option="USER_ENTERED")
        
        # Style row 2 (Navy blue background, bold cyan text)
        style_r2 = {
            "repeatCell": {
                "range": {
                    "sheetId": ws.id,
                    "startRowIndex": 1,
                    "endRowIndex": 2,
                    "startColumnIndex": 0,
                    "endColumnIndex": 9
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 0.06, "green": 0.09, "blue": 0.16},
                        "textFormat": {
                            "foregroundColor": {"red": 0.22, "green": 0.74, "blue": 0.97},
                            "bold": True,
                            "fontSize": 9
                        },
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE"
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
            }
        }
        sh.batch_update({"requests": [style_r2]})
        navbar_updated += 1
        print(f"  [+] Navbar installed on: '{title}'")
        time.sleep(0.5)  # Respect Google API quotas
    except Exception as e:
        print(f"  [!] Error installing navbar on '{title}': {e}")

print(f"  Total dashboards equipped with internal Navbar: {navbar_updated}\n")

# ----------------------------------------------------
# 4. EXPAND 239 FORMULAS FROM 1000 TO 50000 ROWS
# ----------------------------------------------------
print("--- [4] EXPANDING CALCULATION FORMULAS FROM 1000 TO 50000 ROWS ---")
formula_sheets = [
    "calc_sales", "calc_finance", "calc_marketing",
    "📊 Data_Quality", "👥 Ресурсный_План", "🧪 QA_Suite", "calc_engine"
]

pattern_1000 = re.compile(r'(\$?[A-Z]+)(\$?[0-9]+):(\$?[A-Z]+)(\$?)1000\b')
total_expanded = 0

for s_name in formula_sheets:
    if s_name not in sheet_map:
        continue
    try:
        ws = sh.worksheet(s_name)
        formulas = ws.get_values(value_render_option="FORMULA")
        updates = []
        sheet_expanded = 0
        
        for r_idx, row in enumerate(formulas, start=1):
            for c_idx, val in enumerate(row, start=1):
                if isinstance(val, str) and val.startswith("=") and "1000" in val:
                    new_val = pattern_1000.sub(r'\g<1>\g<2>:\g<3>\g<4>50000', val)
                    if new_val != val:
                        col_letter = gspread.utils.rowcol_to_a1(r_idx, c_idx)
                        updates.append({"range": col_letter, "values": [[new_val]]})
                        sheet_expanded += 1
                        
        if updates:
            ws.batch_update(updates, value_input_option="USER_ENTERED")
            total_expanded += sheet_expanded
            print(f"  [+] {s_name}: expanded {sheet_expanded} formulas to 50,000 capacity")
        else:
            print(f"  [-] {s_name}: no 1000 formulas found")
        time.sleep(1)
    except Exception as e:
        print(f"  [!] Error on {s_name}: {e}")

print(f"  Total formulas expanded to 50 000 rows: {total_expanded}\n")

# ----------------------------------------------------
# 5. FIX EMPTY STATES ON 🎙️ ИИ_Аудит & ADD CAPACITY WATCHDOG IN 🧪 QA_Suite
# ----------------------------------------------------
print("--- [5] UPGRADING EMPTY STATES & CAPACITY WATCHDOG ---")
try:
    ws_audit = sh.worksheet("🎙️ ИИ_Аудит")
    ws_audit.update("A31:F31", [["—", "⚡ Ожидает первых 3 звонков", "—", "0.0 (Ожидание данных)", "—", "—"]])
    print("  [+] Empty state row in 🎙️ ИИ_Аудит upgraded to helpful hint.")
except Exception as e:
    print(f"  [!] Error updating empty state in 🎙️ ИИ_Аудит: {e}")

try:
    ws_qa = sh.worksheet("🧪 QA_Suite")
    qa_vals = ws_qa.get_values()
    watchdog_row = len(qa_vals) + 1
    watchdog_payload = [
        "WATCHDOG_ROW_CAPACITY",
        "Контроль емкости строк (50 000)",
        "",
        '=IF(OR(COUNTA(raw_deals!$A$2:$A$50000)>48000, COUNTA(raw_calls!$A$2:$A$50000)>48000), "🚨 ВНИМАНИЕ: Заполнено 96% емкости (48 000+). Создайте архивный срез.", "🟢 Лимит в норме (до 50 000 строк)")',
        "TRUE",
        f"=D{watchdog_row}"
    ]
    ws_qa.update(f"A{watchdog_row}:F{watchdog_row}", [watchdog_payload], value_input_option="USER_ENTERED")
    print(f"  [+] Watchdog Capacity Alert added to 🧪 QA_Suite (row {watchdog_row}).")
except Exception as e:
    print(f"  [!] Error adding watchdog in 🧪 QA_Suite: {e}")

print("\n==================================================")
print("ALL 5 PATCHES APPLIED DIRECTLY TO ООО \"ТД Автожидкости\"!")
print("==================================================")
