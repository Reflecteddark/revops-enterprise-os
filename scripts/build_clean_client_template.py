import sys
import time
import gspread

sys.stdout.reconfigure(encoding="utf-8")

SRC_ID = "1QnjrrbpqhYssofchee7G06szWrFvBCcOqGIVokjqdVc"
DST_ID = "1jBBotOfFh-XEJGScJyQi10jiFrna66OpJ91yPDrXy2A"

client = gspread.service_account(filename="service_account.json")
sh_src = client.open_by_key(SRC_ID)
sh_dst = client.open_by_key(DST_ID)

print(f"Source: '{sh_src.title}' ({len(sh_src.worksheets())} sheets)")
print(f"Target: '{sh_dst.title}'")

# Rename Target spreadsheet
sh_dst.update_title("RevOps Platform V17.6 (Clean Client Starter Template)")
print("Updated target spreadsheet title.")

# Order of sheets to copy
src_worksheets = sh_src.worksheets()
total = len(src_worksheets)

print(f"Starting copy of {total} sheets...")

copied_map = [] # (original_title, copied_id)

for idx, ws in enumerate(src_worksheets, 1):
    orig_title = ws.title
    print(f"[{idx}/{total}] Copying '{orig_title}'...", end=" ", flush=True)
    try:
        res = ws.copy_to(DST_ID)
        copied_sheet_id = res["sheetId"]
        copied_title = res["title"]
        copied_map.append((orig_title, copied_sheet_id, copied_title))
        print(f"OK (id: {copied_sheet_id})")
        time.sleep(0.5) # avoid rate limit
    except Exception as e:
        print(f"ERROR: {e}")

print(f"\nAll {len(copied_map)} sheets copied. Now renaming copied sheets in destination...")

# Rename copied sheets (strip ' (копия)')
requests = []
for orig_title, sheet_id, _ in copied_map:
    requests.append({
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "title": orig_title
            },
            "fields": "title"
        }
    })

if requests:
    sh_dst.batch_update({"requests": requests})
    print("Renamed all copied sheets to their original titles!")

# Delete old initial sheet 'УТП' if it exists
try:
    ws_old = sh_dst.worksheet("УТП")
    sh_dst.del_worksheet(ws_old)
    print("Deleted old 'УТП' placeholder sheet.")
except Exception as e:
    print("Old sheet notice:", e)

# Reorder presentation tabs to front (indices 0 to 5)
presentation_tabs = [
    "⚡ Экспресс_Калькулятор_3_Цифры",
    "📋 Пульт_РОПа_15_Минут",
    "🎙️ ИИ_Аудит",
    "💸 Диагностика_Утечек_ОП",
    "📄 Executive_OnePager",
    "⚙️ Настройки"
]

reorder_reqs = []
for new_idx, title in enumerate(presentation_tabs):
    try:
        ws = sh_dst.worksheet(title)
        reorder_reqs.append({
            "updateSheetProperties": {
                "properties": {
                    "sheetId": ws.id,
                    "index": new_idx
                },
                "fields": "index"
            }
        })
    except Exception as e:
        print(f"Notice reordering {title}: {e}")

if reorder_reqs:
    sh_dst.batch_update({"requests": reorder_reqs})
    print("Reordered presentation tabs to front (indices 0-5).")

print(f"\n🚀 Clean Client Starter Template successfully created in Google Sheets!")
print(f"URL: https://docs.google.com/spreadsheets/d/{DST_ID}/edit")
