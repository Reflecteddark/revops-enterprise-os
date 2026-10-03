import sys
import gspread

sys.stdout.reconfigure(encoding="utf-8")
client = gspread.service_account(filename="service_account.json")
sh = client.open_by_key("1QnjrrbpqhYssofchee7G06szWrFvBCcOqGIVokjqdVc")
ws_calc = sh.worksheet("⚡ Экспресс_Калькулятор_3_Цифры")

# Update Benchmark descriptions in Calculator
ws_calc.update(range_name="C15", values=[["28% выручки теряется из-за ошибок менеджеров на звонках (бенчмарк 2026)"]])
ws_calc.update(range_name="C17", values=[["12% — потенциал реактивации списанных отказников CRM"]])
ws_calc.update(range_name="C18", values=[["10% выставленных счетов зависают на согласовании (DSO)"]])

# Update CTA Link in R25
link_formula = '=HYPERLINK("https://ai-rop.ru/?utm_source=sheets_calculator", "🚀 ЗАПУСТИТЬ БЕСПЛАТНЫЙ ТЕСТ-ДРАЙВ НА 3 ЗВОНКАХ (ПЕРЕЙТИ НА AI-ROP.RU)")'
ws_calc.update(range_name="A25", values=[[link_formula]], raw=False)

print("⚡ Экспресс_Калькулятор_3_Цифры updated successfully!")
