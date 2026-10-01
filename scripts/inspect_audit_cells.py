import gspread
import sys

sys.stdout.reconfigure(encoding="utf-8")

client = gspread.service_account(filename="service_account.json")
sh = client.open_by_key("1QnjrrbpqhYssofchee7G06szWrFvBCcOqGIVokjqdVc")
ws = sh.worksheet("🎙️ ИИ_Аудит")

print("Rows:", ws.row_count, "Cols:", ws.col_count)
cells = ws.get_all_values()
for i, r in enumerate(cells[:25]):
    if any(cell.strip() for cell in r):
        print(f"Row {i+1}: {r[:10]}")
