import sys
import gspread

sys.stdout.reconfigure(encoding="utf-8")

DST_ID = "1jBBotOfFh-XEJGScJyQi10jiFrna66OpJ91yPDrXy2A"
gc = gspread.service_account(filename="service_account.json")
sh = gc.open_by_key(DST_ID)
ws = sh.worksheet("🎙️ ИИ_Аудит")

# 1. Clear rows 18-26 completely (A18:V26)
empty_rows_matrix = [["" for _ in range(22)] for _ in range(9)]

# 2. Set clean row 17
row_17 = ["—", "«Ожидание первой аудиозаписи звонка из телефонии»", "—", 0, "—", "—"] + ["—"] * 13 + [0, "🟢 ОЖИДАНИЕ", 0]

# 3. Row 27 summary
row_27 = ["ИТОГО ПОТЕРЬ И РИСКОВ ПО ЗВОНКАМ:", "", "", "=SUM(D17:D26)", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "=IFERROR(AVERAGE(T17:T26), 0)", '=COUNTIF(U17:U26, "*СЛИВ*") & " сливов"', "=SUM(V17:V26)"]

# 4. Clear rows 32-35 completely (A32:N35)
empty_rows_defects = [["" for _ in range(14)] for _ in range(4)]

# 5. Set clean row 31
row_31 = ["—", "«Дефекты речи не обнаружены (ожидание звонков)»", "—", 0, "—", "", "—", "", "", "", "", "", "", "«При обнаружении дефекта речи нейросеть выделит цитату и сформирует инструкцию РОПу»"]

batch_data = [
    {"range": "'🎙️ ИИ_Аудит'!A17:V17", "values": [row_17]},
    {"range": "'🎙️ ИИ_Аудит'!A18:V26", "values": empty_rows_matrix},
    {"range": "'🎙️ ИИ_Аудит'!A27:V27", "values": [row_27]},
    {"range": "'🎙️ ИИ_Аудит'!A31:N31", "values": [row_31]},
    {"range": "'🎙️ ИИ_Аудит'!A32:N35", "values": empty_rows_defects}
]

sh.values_batch_update({
    "valueInputOption": "USER_ENTERED",
    "data": batch_data
})

print("Successfully cleaned 🎙️ ИИ_Аудит rows 17-35!")
