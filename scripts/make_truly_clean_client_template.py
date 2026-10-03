import sys
import gspread

sys.stdout.reconfigure(encoding="utf-8")

DST_ID = "1jBBotOfFh-XEJGScJyQi10jiFrna66OpJ91yPDrXy2A"
client = gspread.service_account(filename="service_account.json")
sh_dst = client.open_by_key(DST_ID)

print(f"Connecting to Clean Client Template: '{sh_dst.title}' ({DST_ID})")

# 1. Clean ⚡ Экспресс_Калькулятор_3_Цифры
ws_calc = sh_dst.worksheet("⚡ Экспресс_Калькулятор_3_Цифры")
print("Cleaning ⚡ Экспресс_Калькулятор_3_Цифры...")

# Clear input cells B5, B6, B7, B8
ws_calc.update(range_name="B5:B8", values=[[""], [""], [""], [""]])

# Set clean safe formulas
calc_formulas = [
    # Row 5
    ("D5", '=IF(OR(B5="", B6=""), 0, B5 * B6)'),
    ("H5", '=D7'),
    # Row 6
    ("D6", '=IF(OR(D5=0, D7=0), 0, D7/D5)'),
    ("H6", '=A11'),
    # Row 7
    ("D7", '=IFERROR(calc_engine!B19, 0)'),
    ("H7", '=C11'),
    # Row 8
    ("D8", '=IF(OR(B5="", B7=""), "—", ROUND(B5 / MAX(B7,1), 0) & " лид/мес")'),
    # Row 11
    ("A11", '=SUM(D15:D18)'),
    ("B11", '=IF(D7=0, 0, D7 + A11*0.15)'),
    ("C11", '=IF(D7=0, 0, D7 + A11*0.3)'),
    ("D11", '=IF(D7=0, 0, D7 + A11*0.5)'),
    # Rows 15-18
    ("D15", '=IF(D7=0, IF(AND(ISNUMBER(B5), ISNUMBER(B6), B5*B6>0), ROUND(B5*B6*0.045*\'⚙️ Настройки\'!B43, 0), 0), ROUND(D7 * \'⚙️ Настройки\'!B43, 0))'),
    ("D16", '=IF(D7=0, IF(AND(ISNUMBER(B5), ISNUMBER(B6), B5*B6>0), ROUND(B5*B6*0.045*\'⚙️ Настройки\'!B44, 0), 0), ROUND(D7 * \'⚙️ Настройки\'!B44, 0))'),
    ("D17", '=IF(D7=0, IF(AND(ISNUMBER(B5), ISNUMBER(B6), B5*B6>0), ROUND(B5*B6*0.045*\'⚙️ Настройки\'!B45, 0), 0), ROUND(D7 * \'⚙️ Настройки\'!B45, 0))'),
    ("D18", '=IF(D7=0, IF(AND(ISNUMBER(B5), ISNUMBER(B6), B5*B6>0), ROUND(B5*B6*0.045*\'⚙️ Настройки\'!B46, 0), 0), ROUND(D7 * \'⚙️ Настройки\'!B46, 0))'),
    # Row 19
    ("D19", '=SUM(D15:D18)'),
    ("F19", '=IF(D19=0, "0 ₽ в первый месяц", ROUND(D19 * \'⚙️ Настройки\'!B48, 0) & " ₽ в первый месяц")')
]

for cell_ref, form in calc_formulas:
    ws_calc.update(range_name=cell_ref, values=[[form]], raw=False)

print("⚡ Экспресс_Калькулятор_3_Цифры is now 100% clean and reactive!")

# 2. Clean 📋 Пульт_РОПа_15_Минут
ws_rop = sh_dst.worksheet("📋 Пульт_РОПа_15_Минут")
print("Cleaning 📋 Пульт_РОПа_15_Минут...")

rop_updates = [
    # Metrics
    ("A5", '=IFERROR(calc_engine!B41, 0)'),
    ("B5", '=IFERROR(COUNTIFS(raw_calls!$T$2:$T1000, "<9"), 0)'),
    ("C5", '=IFERROR(COUNTIF(E9:E13, "*КРИТИЧЕСКИЙ*"), 0)'),
    ("D5", '=IFERROR(SUM(D9:D13), 0)'),
    ("E5", '=IF(A5=0, "🟢 ОЖИДАНИЕ ЗВОНКОВ И СДЕЛОК ИЗ CRM", IF(C5>0, "🚨 ТРЕБУЕТСЯ 15 МИНУТ КОНТРОЛЯ РОПа", "🟢 ВСЕ В РАБОЧЕМ ГРАФИКЕ"))'),
    # Clean rows 9-13
    ("A9", 1), ("B9", "«Ожидание синхронизации сделок с CRM»"), ("C9", "Сделки в риске не зафиксированы"), ("D9", 0), ("E9", "—"), ("F9", "—"), ("G9", "—"), ("H9", "«Персональный скрипт сформируется автоматически при возникновении риска по сделке»"),
    ("A10", 2), ("B10", "—"), ("C10", "—"), ("D10", 0), ("E10", "—"), ("F10", "—"), ("G10", "—"), ("H10", "—"),
    ("A11", 3), ("B11", "—"), ("C11", "—"), ("D11", 0), ("E11", "—"), ("F11", "—"), ("G11", "—"), ("H11", "—"),
    ("A12", 4), ("B12", "—"), ("C12", "—"), ("D12", 0), ("E12", "—"), ("F12", "—"), ("G12", "—"), ("H12", "—"),
    ("A13", 5), ("B13", "—"), ("C13", "—"), ("D13", 0), ("E13", "—"), ("F13", "—"), ("G13", "—"), ("H13", "—"),
    ("D14", '=SUM(D9:D13)')
]

for cell_ref, val in rop_updates:
    is_f = isinstance(val, str) and val.startswith("=")
    ws_rop.update(range_name=cell_ref, values=[[val]], raw=not is_f)

print("📋 Пульт_РОПа_15_Минут is now 100% clean!")

# 3. Clean 🎙️ ИИ_Аудит
ws_audit = sh_dst.worksheet("🎙️ ИИ_Аудит")
print("Cleaning 🎙️ ИИ_Аудит...")

audit_updates = [
    # Metrics
    ("A4", "0 звонков / сут"),
    ("C4", "0.0 из 13"),
    ("E4", "0.0% (ожидание звонков)"),
    ("G4", "0 диалогов"),
    ("J4", 0),
    ("M4", "🟢 ОЖИДАНИЕ ПЕРВЫХ ЗВОНКОВ ПО API"),
    # Leaderboard rows 8-12
    ("B8", "Менеджер 1"), ("C8", 0), ("D8", 0.0), ("E8", "0.0%"), ("F8", "—"),
    ("B9", "Менеджер 2"), ("C9", 0), ("D9", 0.0), ("E9", "0.0%"), ("F9", "—"),
    ("B10", "Менеджер 3"), ("C10", 0), ("D10", 0.0), ("E10", "0.0%"), ("F10", "—"),
    ("B11", "Менеджер 4"), ("C11", 0), ("D11", 0.0), ("E11", "0.0%"), ("F11", "—"),
    ("B12", "Менеджер 5"), ("C12", 0), ("D12", 0.0), ("E12", "0.0%"), ("F12", "—"),
    ("C13", 0), ("D13", "0.0 / 13"), ("E13", "0.0%"), ("F13", "—"),
    # Calls table R17-R26
    ("A17", "—"), ("B17", "«Ожидание первой аудиозаписи звонка из телефонии»"), ("C17", "—"), ("D17", 0), ("E17", "—"), ("F17", "—"),
    ("A18", ""), ("B18", ""), ("C18", ""), ("D18", ""), ("E18", ""), ("F18", ""),
    ("A19", ""), ("B19", ""), ("C19", ""), ("D19", ""), ("E19", ""), ("F19", ""),
    ("A20", ""), ("B20", ""), ("C20", ""), ("D20", ""), ("E20", ""), ("F20", ""),
    ("A21", ""), ("B21", ""), ("C21", ""), ("D21", ""), ("E21", ""), ("F21", ""),
    ("A22", ""), ("B22", ""), ("C22", ""), ("D22", ""), ("E22", ""), ("F22", ""),
    ("A23", ""), ("B23", ""), ("C23", ""), ("D23", ""), ("E23", ""), ("F23", ""),
    ("A24", ""), ("B24", ""), ("C24", ""), ("D24", ""), ("E24", ""), ("F24", ""),
    ("A25", ""), ("B25", ""), ("C25", ""), ("D25", ""), ("E25", ""), ("F25", ""),
    ("A26", ""), ("B26", ""), ("C26", ""), ("D26", ""), ("E26", ""), ("F26", ""),
    ("D27", 0), ("V27", 0),
    # Speech defects R31-R35
    ("A31", "—"), ("B31", "«Дефекты речи не обнаружены (ожидание звонков)»"), ("C31", "—"), ("D31", 0), ("E31", "—"), ("G31", "—"),
    ("A32", ""), ("B32", ""), ("C32", ""), ("D32", ""), ("E32", ""), ("G32", ""),
    ("A33", ""), ("B33", ""), ("C33", ""), ("D33", ""), ("E33", ""), ("G33", ""),
    ("A34", ""), ("B34", ""), ("C34", ""), ("D34", ""), ("E34", ""), ("G34", ""),
    ("A35", ""), ("B35", ""), ("C35", ""), ("D35", ""), ("E35", ""), ("G35", "")
]

for cell_ref, val in audit_updates:
    is_f = isinstance(val, str) and val.startswith("=")
    ws_audit.update(range_name=cell_ref, values=[[val]], raw=not is_f)

print("🎙️ ИИ_Аудит is now 100% clean!")

# 4. Clean 💸 Диагностика_Утечек_ОП
ws_leak = sh_dst.worksheet("💸 Диагностика_Утечек_ОП")
print("Cleaning 💸 Диагностика_Утечек_ОП...")

leak_updates = [
    # Top metrics
    ("A5", '=IFERROR(calc_engine!B19, 0)'),
    ("B5", '=IFERROR(calc_engine!B18, 0)'),
    ("C5", '=IFERROR(E16, 0)'),
    ("D5", '=A5 + E16'),
    # 7 Sins table rows 9-15
    ("D9", '=IFERROR(calc_engine!B17, 0)'), ("E9", '=IFERROR(calc_engine!B17 - calc_engine!B18, 0)'),
    ("D10", '=IFERROR(calc_engine!C5, 0)'), ("E10", '=IFERROR(ROUND(D10 * 0.45, 0), 0)'),
    ("D11", '=IFERROR(SUMIFS(raw_deals!$C$2:$C1000, raw_deals!$H$2:$H1000, 101, raw_deals!$D$2:$D1000, "<=5"), 0)'), ("E11", '=IFERROR(ROUND(D11 * 0.3, 0), 0)'),
    ("D12", '=IFERROR(calc_engine!B17, 0)'), ("E12", '=IFERROR(ROUND(D12 * 0.25, 0), 0)'),
    ("D13", '=IFERROR(SUM(calc_engine!D52:D55), 0)'), ("E13", '=IFERROR(calc_engine!D56, 0)'),
    ("D14", '=IFERROR(calc_engine!B146, 0)'), ("E14", '=D14'),
    ("D15", '=IFERROR(SUMIF(raw_marketing!$B$2:$B1000, "Google Ads", raw_marketing!$E$2:$E1000), 0)'), ("E15", '=D15'),
    ("D16", '=SUM(D9:D15)'), ("E16", '=SUM(E9:E15)'),
    ("B21", '=ROUND(E16 * 0.3, 0)'),
    ("B22", '=B21 - B20'),
    ("B23", '=IF(B20=0, "0%", ROUND((B21 - B20) / B20 * 100, 0) & "%")'),
    ("B24", '=IF(B21=0, "0 дн.", ROUND(B20 / (B21 / 30), 0) & " дн.")')
]

for cell_ref, val in leak_updates:
    is_f = isinstance(val, str) and val.startswith("=")
    ws_leak.update(range_name=cell_ref, values=[[val]], raw=not is_f)

print("💸 Диагностика_Утечек_ОП is now 100% clean!")

# 5. Clean 📄 Executive_OnePager
ws_one = sh_dst.worksheet("📄 Executive_OnePager")
print("Cleaning 📄 Executive_OnePager...")

one_updates = [
    # Top KPI row 5
    ("A5", '=IFERROR(calc_engine!B19, 0)'),
    ("B5", '=IFERROR(calc_engine!B20, 0)'),
    ("C5", '=IFERROR(calc_engine!B18, 0)'),
    ("D5", '=IFERROR(calc_engine!B29, 0)'),
    ("E5", '=IFERROR(calc_engine!B17, 0)'),
    ("F5", 0),
    ("G5", 0.0),
    # Funnel stages R9-R14 deals count & amount
    ("B9", 0), ("C9", 0),
    ("B10", 0), ("C10", 0),
    ("B11", 0), ("C11", 0),
    ("B12", 0), ("C12", 0),
    ("B13", 0), ("C13", 0),
    ("B14", 0), ("C14", 0),
    # Top 5 risks table (F9:I13)
    ("F9", "«Ожидание сделок из CRM»"), ("G9", "—"), ("H9", 0), ("I9", "🟢 В НОРМЕ"),
    ("F10", "—"), ("G10", "—"), ("H10", 0), ("I10", "—"),
    ("F11", "—"), ("G11", "—"), ("H11", 0), ("I11", "—"),
    ("F12", "—"), ("G12", "—"), ("H12", 0), ("I12", "—"),
    ("F13", "—"), ("G13", "—"), ("H13", 0), ("I13", "—")
]

for cell_ref, val in one_updates:
    is_f = isinstance(val, str) and val.startswith("=")
    ws_one.update(range_name=cell_ref, values=[[val]], raw=not is_f)

print("📄 Executive_OnePager is now 100% clean!")

# 6. Clean ⚙️ Настройки staff placeholders
ws_set = sh_dst.worksheet("⚙️ Настройки")
print("Cleaning ⚙️ Настройки staff and company info...")

set_updates = [
    ("B3", "ООО «Название Компании Клиента»"),
    ("E3", "CLIENT-TENANT-001"),
    ("F3", "(Пространство нового клиента)"),
    ("E4", "RevOps OS V17.6 Client Edition"),
    ("B21", "Менеджер 1 (РОП)"),
    ("B22", "Менеджер 2 (КАМ)"),
    ("B23", "Менеджер 3 (SDR)"),
    ("B24", "Менеджер 4 (SDR-2)")
]

for cell_ref, val in set_updates:
    ws_set.update(range_name=cell_ref, values=[[val]])

print("⚙️ Настройки is now 100% clean!")

print("\n🚀 All showcase sheets of Clean Client Template are now 100% PRISTINE CLEAN!")
