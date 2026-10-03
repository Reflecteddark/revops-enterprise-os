#!/usr/bin/env python3
"""
Finish remaining calculation sheets on client Google Sheet:
RevOps Platform V18 ООО "ТД Автожидкости" (1xxGSOBFzoZV1u7mBmpexD0jXvmujHchCBGyEGBWlmIs)
With rate-limiting protection to respect Google Sheets API 60 req/min quota.
"""

import gspread
import re
import sys
import time

sys.stdout.reconfigure(encoding='utf-8')

SPREADSHEET_ID = "1xxGSOBFzoZV1u7mBmpexD0jXvmujHchCBGyEGBWlmIs"

client = gspread.service_account(filename="service_account.json")
sh = client.open_by_key(SPREADSHEET_ID)
sheet_map = {ws.title: ws.id for ws in sh.worksheets()}

pattern_1000 = re.compile(r'(\$?[A-Z]+)(\$?[0-9]+):(\$?[A-Z]+)(\$?)1000\b')

sheets_to_finish = [
    "calc_finance",
    "calc_marketing",
    "📊 Data_Quality",
    "👥 Ресурсный_План",
    "calc_engine"
]

print("Finishing remaining calculation formulas with rate limiting...")
total_formulas = 0

for s_name in sheets_to_finish:
    if s_name not in sheet_map:
        continue
    try:
        ws = sh.worksheet(s_name)
        formulas = ws.get_values(value_render_option="FORMULA")
        updates = []
        sheet_count = 0
        
        for r_idx, row in enumerate(formulas, start=1):
            for c_idx, val in enumerate(row, start=1):
                if isinstance(val, str) and val.startswith("=") and "1000" in val:
                    new_val = pattern_1000.sub(r'\g<1>\g<2>:\g<3>\g<4>50000', val)
                    if new_val != val:
                        col_letter = gspread.utils.rowcol_to_a1(r_idx, c_idx)
                        updates.append({"range": col_letter, "values": [[new_val]]})
                        sheet_count += 1
                        
        if updates:
            # Batch update all formulas on this sheet in a single API call!
            ws.batch_update(updates, value_input_option="USER_ENTERED")
            total_formulas += sheet_count
            print(f"  [+] {s_name}: successfully expanded {sheet_count} formulas to 50,000 capacity!")
        else:
            print(f"  [-] {s_name}: already up to date.")
            
        time.sleep(2)  # Respect rate limit
    except Exception as e:
        print(f"  [!] Error on {s_name}: {e}")
        time.sleep(5)

# Empty states update
print("\nUpdating empty states on 🎙️ ИИ_Аудит...")
try:
    time.sleep(2)
    ws_audit = sh.worksheet("🎙️ ИИ_Аудит")
    ws_audit.update([["—", "⚡ Ожидает первых 3 звонков", "—", "0.0 (Ожидание данных)", "—", "—"]], range_name="A31:F31")
    print("  [+] Empty state row on 🎙️ ИИ_Аудит updated successfully!")
except Exception as e:
    print(f"  [!] Error on empty states: {e}")

# Watchdog update
print("\nAdding Watchdog in 🧪 QA_Suite...")
try:
    time.sleep(2)
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
    ws_qa.update([watchdog_payload], range_name=f"A{watchdog_row}:F{watchdog_row}", value_input_option="USER_ENTERED")
    print(f"  [+] Watchdog Capacity Alert added to 🧪 QA_Suite row {watchdog_row}!")
except Exception as e:
    print(f"  [!] Error on watchdog: {e}")

print(f"\nSuccessfully finished! Total formulas expanded: {total_formulas}")
