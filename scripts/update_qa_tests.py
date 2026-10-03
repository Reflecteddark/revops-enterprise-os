import sys
import gspread

sys.stdout.reconfigure(encoding="utf-8")

client = gspread.service_account(filename="service_account.json")
sh = client.open_by_key("1QnjrrbpqhYssofchee7G06szWrFvBCcOqGIVokjqdVc")
ws_qa = sh.worksheet("🧪 QA_Suite")

# Update T-29
row_35 = [
    "T-29",
    "ИИ-Аудит речи",
    "Хронометраж звонка в ИИ_Аудит!E17 зафиксирован в формате времени",
    True,
    True,
    '=ISNUMBER(SEARCH(":", \'🎙️ ИИ_Аудит\'!E17))',
    '=IF(F35=TRUE, "🟢 PASS", "🔴 FAIL")',
    "В колонке длительности звонка нарушен формат времени."
]
ws_qa.update(range_name="A35:H35", values=[row_35], raw=False)

# Update T-30
row_36 = [
    "T-30",
    "ИИ-Аудит речи",
    "Выручка под угрозой в ИИ_Аудит!J4 сходится с итогом выборки V27",
    "='🎙️ ИИ_Аудит'!V27",
    "='🎙️ ИИ_Аудит'!J4",
    "=D36",
    '=IF(E36=F36, "🟢 PASS", "🔴 FAIL")',
    "Сумма финансового риска брака звонков расходится с контрольной суммой выборки."
]
ws_qa.update(range_name="A36:H36", values=[row_36], raw=False)

print("🧪 QA_Suite tests T-29 and T-30 updated!")
