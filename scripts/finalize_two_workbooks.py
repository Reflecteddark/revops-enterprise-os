import sys
import requests
import gspread

sys.stdout.reconfigure(encoding="utf-8")

MASTER_ID = "1QnjrrbpqhYssofchee7G06szWrFvBCcOqGIVokjqdVc"
CLEAN_ID = "1jBBotOfFh-XEJGScJyQi10jiFrna66OpJ91yPDrXy2A"

client = gspread.service_account(filename="service_account.json")

# 1. Update Title of Master Presentation Sheet
sh_master = client.open_by_key(MASTER_ID)
sh_master.update_title("RevOps Platform V17.6 (Showcase Master — Презентационная)")
print(f"Master Sheet Title: '{sh_master.title}'")

# 2. Update Title of Clean Client Sheet
sh_clean = client.open_by_key(CLEAN_ID)
sh_clean.update_title("RevOps Platform V17.6 (Clean Client Starter — Чистая для клиентов)")
print(f"Clean Sheet Title: '{sh_clean.title}'")

# 3. Export Clean Sheet as local XLSX
url = f"https://docs.google.com/spreadsheets/d/{CLEAN_ID}/export?format=xlsx"
res = requests.get(url)
if res.status_code == 200:
    out_xlsx = "docs/RevOps_Platform_V17.6_Clean_Client_Starter.xlsx"
    with open(out_xlsx, "wb") as f:
        f.write(res.content)
    print(f"Exported clean client workbook to: {out_xlsx} ({len(res.content)} bytes)")
    # Also copy to presentation folder
    out_pres = "presentation/RevOps_Platform_V17.6_Clean_Client_Starter.xlsx"
    with open(out_pres, "wb") as f:
        f.write(res.content)
    print(f"Copied to: {out_pres}")
else:
    print(f"Notice exporting clean XLSX: status {res.status_code}")

print("\n🚀 Оба варианта таблиц успешно подготовлены и зафиксированы!")
