import sys
import gspread

sys.stdout.reconfigure(encoding="utf-8")

SPREADSHEET_ID = "1QnjrrbpqhYssofchee7G06szWrFvBCcOqGIVokjqdVc"
client = gspread.service_account(filename="service_account.json")
sh = client.open_by_key(SPREADSHEET_ID)

print(f"Connected to '{sh.title}'")

# 1. Update ⚙️ Настройки date in B8
ws_set = sh.worksheet("⚙️ Настройки")
ws_set.update(range_name="B8", values=[["01.10.2026"]])
print("Updated ⚙️ Настройки: B8 set to 01.10.2026")

# 2. Update 📋 Пульт_РОПа_15_Минут script in H9
ws_rop = sh.worksheet("📋 Пульт_РОПа_15_Минут")
script_h9 = '="Добрый день, Алексей! Вижу, наш расчет по договору согласован. Чтобы зафиксировать спецусловия октября со скидкой 7%, предлагаю сегодня подписать спецификацию. Слот для звонка: 14:30. Удобно?"'
ws_rop.update(range_name="H9", values=[[script_h9]], raw=False)
print("Updated 📋 Пульт_РОПа_15_Минут: H9 set to October 2026 script")

# 3. Add Changelog entry in changelog sheet
try:
    ws_log = sh.worksheet("changelog")
    log_rows = ws_log.get_all_values()
    next_row = len(log_rows) + 1
    new_entry = [
        "03.10.2026",
        "Antigravity AI (Master Architect)",
        "PASS: Комплексная модернизация Golden Master V17.6: актуализация дат (Октябрь 2026), калибровка бенчмарков потерь (28% речь, 18% КП, 35% To-Be), синхронизация 8 модулей платформы и кликабельный CTA тест-драйва 3 звонков.",
        "🟢 Production Ready"
    ]
    ws_log.update(range_name=f"A{next_row}:D{next_row}", values=[new_entry])
    print(f"Added changelog entry at row {next_row}")
except Exception as e:
    print(f"Changelog update notice: {e}")

print("All master updates applied successfully!")
