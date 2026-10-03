import sys
import time
import gspread

sys.stdout.reconfigure(encoding="utf-8")

DST_ID = "1jBBotOfFh-XEJGScJyQi10jiFrna66OpJ91yPDrXy2A"
client = gspread.service_account(filename="service_account.json")
sh = client.open_by_key(DST_ID)

print(f"Connected to Clean Client Template: '{sh.title}' ({DST_ID})")

# Step 1: Batch clear all data rows (Row 2 to 1000) from raw and log tables
ranges_to_clear = [
    "'raw_deals'!A2:AZ1000",
    "'raw_calls'!A2:AZ1000",
    "'raw_touchpoints'!A2:AZ1000",
    "'raw_invoices'!A2:AZ1000",
    "'raw_payments'!A2:AZ1000",
    "'raw_marketing'!A2:AZ1000",
    "'raw_alerts'!A2:AZ1000",
    "'raw_audit_log'!A2:AZ1000",
    "'📅 Deal_Event_Log'!A2:AZ1000",
    "'📊 Stage_History'!A2:AZ1000",
    "'🎯 Action_Center'!A2:AZ1000"
]

print("1. Clearing all demo data rows across 11 raw/log sheets in 1 batch call...")
for attempt in range(1, 4):
    try:
        sh.values_batch_clear(body={"ranges": ranges_to_clear})
        print("  -> Successfully cleared all raw and log tables!")
        break
    except Exception as e:
        print(f"  Attempt {attempt} failed: {e}. Retrying after 20s...")
        time.sleep(20)

# Step 2: Prepare batch update for showcase sheets, registry, tasks, settings
batch_data = []

def add_batch(sheet_title, cell_range, values):
    batch_data.append({
        "range": f"'{sheet_title}'!{cell_range}",
        "values": values
    })

# 2.1 🆔 ID_Registry reset
add_batch("🆔 ID_Registry", "C2:E9", [
    [101, "—", 0],
    [101, "—", 0],
    [101, "—", 0],
    [101, "—", 0],
    [101, "—", 0],
    [101, "—", 0],
    [101, "—", 0],
    [101, "—", 0]
])

# 2.2 📋 Задачи_и_Спринты: Client Onboarding Tasks
add_batch("📋 Задачи_и_Спринты", "A3:F13", [
    ["T-01", "Подключение API amoCRM / Битрикс24", "Интеграция", "P0", "RevOps Lead", "День 1"],
    ["T-02", "Подключение телефонии к Whisper AI речевой аналитике", "Телефония", "P0", "Интегратор", "День 2"],
    ["T-03", "Установка целевого плана продаж и лимитов менеджеров в ⚙️ Настройки", "Конфигурация", "P0", "РОП", "День 2"],
    ["T-04", "Подключение Telegram-бота супервайзера @revops_supervisor_bot", "Алерты", "P1", "РОП", "День 3"],
    ["T-05", "Проведение первых 10 звонков и проверка скоринга в 🎙️ ИИ_Аудит", "ОКК", "P1", "РОП / ОКК", "День 4"],
    ["T-06", "Сверка счетов и дебиторской задолженности в 💳 Финансы", "Финансы", "P1", "Бухгалтерия", "День 5"],
    ["", "", "", "", "", ""],
    ["", "", "", "", "", ""],
    ["", "", "", "", "", ""],
    ["", "", "", "", "", ""],
    ["ИТОГО ЗАДАЧ", "=COUNTA(A3:A8)", "Выполнено:", 0, "В работе:", "=COUNTA(A3:A8)"]
])

# 2.3 Connect manager names in calc_engine and 👥 Ресурсный_План
add_batch("calc_engine", "B172:B174", [
    ["='⚙️ Настройки'!B21"],
    ["='⚙️ Настройки'!B22"],
    ["='⚙️ Настройки'!B23"]
])
add_batch("👥 Ресурсный_План", "B8:B10", [
    ["='⚙️ Настройки'!B21"],
    ["='⚙️ Настройки'!B22"],
    ["='⚙️ Настройки'!B23"]
])

# 2.4 Protect division in 👥 Мотивация_ОП row 4 col F
add_batch("👥 Мотивация_ОП", "F4", [["=IFERROR(A4 / calc_engine!$B$19, 0)"]])

# 2.5 Ensure ⚡ Экспресс_Калькулятор_3_Цифры has completely clean empty inputs and reactive formulas
add_batch("⚡ Экспресс_Калькулятор_3_Цифры", "B5:B8", [[""], [""], [""], [""]])
add_batch("⚡ Экспресс_Калькулятор_3_Цифры", "D5", [['=IF(OR(B5="", B6=""), 0, B5 * B6)']])
add_batch("⚡ Экспресс_Калькулятор_3_Цифры", "H5", [['=D7']])
add_batch("⚡ Экспресс_Калькулятор_3_Цифры", "D6", [['=IF(OR(D5=0, D7=0), 0, D7/D5)']])
add_batch("⚡ Экспресс_Калькулятор_3_Цифры", "H6", [['=A11']])
add_batch("⚡ Экспресс_Калькулятор_3_Цифры", "D7", [['=IFERROR(calc_engine!B19, 0)']])
add_batch("⚡ Экспресс_Калькулятор_3_Цифры", "H7", [['=C11']])
add_batch("⚡ Экспресс_Калькулятор_3_Цифры", "D8", [['=IF(OR(B5="", B7=""), "—", ROUND(B5 / MAX(B7,1), 0) & " лид/мес")']])
add_batch("⚡ Экспресс_Калькулятор_3_Цифры", "A11", [['=SUM(D15:D18)']])
add_batch("⚡ Экспресс_Калькулятор_3_Цифры", "B11", [['=IF(D7=0, 0, D7 + A11*0.15)']])
add_batch("⚡ Экспресс_Калькулятор_3_Цифры", "C11", [['=IF(D7=0, 0, D7 + A11*0.3)']])
add_batch("⚡ Экспресс_Калькулятор_3_Цифры", "D11", [['=IF(D7=0, 0, D7 + A11*0.5)']])
add_batch("⚡ Экспресс_Калькулятор_3_Цифры", "D15", [['=IF(D7=0, IF(AND(ISNUMBER(B5), ISNUMBER(B6), B5*B6>0), ROUND(B5*B6*0.045*\'⚙️ Настройки\'!B43, 0), 0), ROUND(D7 * \'⚙️ Настройки\'!B43, 0))']])
add_batch("⚡ Экспресс_Калькулятор_3_Цифры", "D16", [['=IF(D7=0, IF(AND(ISNUMBER(B5), ISNUMBER(B6), B5*B6>0), ROUND(B5*B6*0.045*\'⚙️ Настройки\'!B44, 0), 0), ROUND(D7 * \'⚙️ Настройки\'!B44, 0))']])
add_batch("⚡ Экспресс_Калькулятор_3_Цифры", "D17", [['=IF(D7=0, IF(AND(ISNUMBER(B5), ISNUMBER(B6), B5*B6>0), ROUND(B5*B6*0.045*\'⚙️ Настройки\'!B45, 0), 0), ROUND(D7 * \'⚙️ Настройки\'!B45, 0))']])
add_batch("⚡ Экспресс_Калькулятор_3_Цифры", "D18", [['=IF(D7=0, IF(AND(ISNUMBER(B5), ISNUMBER(B6), B5*B6>0), ROUND(B5*B6*0.045*\'⚙️ Настройки\'!B46, 0), 0), ROUND(D7 * \'⚙️ Настройки\'!B46, 0))']])
add_batch("⚡ Экспресс_Калькулятор_3_Цифры", "D19", [['=SUM(D15:D18)']])
add_batch("⚡ Экспресс_Калькулятор_3_Цифры", "F19", [['=IF(D19=0, "0 ₽ в первый месяц", ROUND(D19 * \'⚙️ Настройки\'!B48, 0) & " ₽ в первый месяц")']])

# 2.6 📋 Пульт_РОПа_15_Минут
add_batch("📋 Пульт_РОПа_15_Минут", "A5:E5", [[
    '=IFERROR(calc_engine!B41, 0)',
    '=IFERROR(COUNTIFS(raw_calls!$T$2:$T1000, "<9"), 0)',
    '=IFERROR(COUNTIF(E9:E13, "*КРИТИЧЕСКИЙ*"), 0)',
    '=IFERROR(SUM(D9:D13), 0)',
    '=IF(A5=0, "🟢 ОЖИДАНИЕ ЗВОНКОВ И СДЕЛОК ИЗ CRM", IF(C5>0, "🚨 ТРЕБУЕТСЯ 15 МИНУТ КОНТРОЛЯ РОПа", "🟢 ВСЕ В РАБОЧЕМ ГРАФИКЕ"))'
]])

add_batch("📋 Пульт_РОПа_15_Минут", "A9:H13", [
    [1, "«Ожидание синхронизации сделок с CRM»", "Сделки в риске не зафиксированы", 0, "—", "—", "—", "«Персональный скрипт сформируется автоматически при возникновении риска по сделке»"],
    [2, "—", "—", 0, "—", "—", "—", "—"],
    [3, "—", "—", 0, "—", "—", "—", "—"],
    [4, "—", "—", 0, "—", "—", "—", "—"],
    [5, "—", "—", 0, "—", "—", "—", "—"]
])
add_batch("📋 Пульт_РОПа_15_Минут", "D14", [['=SUM(D9:D13)']])

# 2.7 🎙️ ИИ_Аудит
add_batch("🎙️ ИИ_Аудит", "A4:N4", [[
    "0 звонков / сут", "", "0.0 из 13", "", "0.0% (ожидание звонков)", "", "0 диалогов", "", "", 0, "", "", "🟢 ОЖИДАНИЕ ПЕРВЫХ ЗВОНКОВ ПО API", ""
]])
add_batch("🎙️ ИИ_Аудит", "B8:F12", [
    ["Менеджер 1", 0, 0.0, "0.0%", "—"],
    ["Менеджер 2", 0, 0.0, "0.0%", "—"],
    ["Менеджер 3", 0, 0.0, "0.0%", "—"],
    ["Менеджер 4", 0, 0.0, "0.0%", "—"],
    ["Менеджер 5", 0, 0.0, "0.0%", "—"]
])
add_batch("🎙️ ИИ_Аудит", "C13:F13", [[0, "0.0 / 13", "0.0%", "—"]])

calls_blank = [["—", "«Ожидание первой аудиозаписи звонка из телефонии»", "—", 0, "—", "—"]]
for _ in range(9):
    calls_blank.append(["", "", "", "", "", ""])
add_batch("🎙️ ИИ_Аудит", "A17:F26", calls_blank)
add_batch("🎙️ ИИ_Аудит", "D27", [[0]])
add_batch("🎙️ ИИ_Аудит", "V27", [[0]])

defects_blank = [["—", "«Дефекты речи не обнаружены (ожидание звонков)»", "—", 0, "—", "", "—"]]
for _ in range(4):
    defects_blank.append(["", "", "", "", "", "", ""])
add_batch("🎙️ ИИ_Аудит", "A31:G35", defects_blank)

# 2.8 💸 Диагностика_Утечек_ОП
add_batch("💸 Диагностика_Утечек_ОП", "A5:D5", [[
    '=IFERROR(calc_engine!B19, 0)',
    '=IFERROR(calc_engine!B18, 0)',
    '=IFERROR(E16, 0)',
    '=A5 + E16'
]])
add_batch("💸 Диагностика_Утечек_ОП", "D9:E15", [
    ['=IFERROR(calc_engine!B17, 0)', '=IFERROR(calc_engine!B17 - calc_engine!B18, 0)'],
    ['=IFERROR(calc_engine!C5, 0)', '=IFERROR(ROUND(D10 * 0.45, 0), 0)'],
    ['=IFERROR(SUMIFS(raw_deals!$C$2:$C1000, raw_deals!$H$2:$H1000, 101, raw_deals!$D$2:$D1000, "<=5"), 0)', '=IFERROR(ROUND(D11 * 0.3, 0), 0)'],
    ['=IFERROR(calc_engine!B17, 0)', '=IFERROR(ROUND(D12 * 0.25, 0), 0)'],
    ['=IFERROR(SUM(calc_engine!D52:D55), 0)', '=IFERROR(calc_engine!D56, 0)'],
    ['=IFERROR(calc_engine!B146, 0)', '=D14'],
    ['=IFERROR(SUMIF(raw_marketing!$B$2:$B1000, "Google Ads", raw_marketing!$E$2:$E1000), 0)', '=D15']
])
add_batch("💸 Диагностика_Утечек_ОП", "D16:E16", [['=SUM(D9:D15)', '=SUM(E9:E15)']])
add_batch("💸 Диагностика_Утечек_ОП", "B21:B24", [
    ['=ROUND(E16 * 0.3, 0)'],
    ['=B21 - B20'],
    ['=IF(B20=0, "0%", ROUND((B21 - B20) / B20 * 100, 0) & "%")'],
    ['=IF(B21=0, "0 дн.", ROUND(B20 / (B21 / 30), 0) & " дн.")']
])

# 2.9 📄 Executive_OnePager
add_batch("📄 Executive_OnePager", "A5:G5", [[
    '=IFERROR(calc_engine!B19, 0)',
    '=IFERROR(calc_engine!B20, 0)',
    '=IFERROR(calc_engine!B18, 0)',
    '=IFERROR(calc_engine!B29, 0)',
    '=IFERROR(calc_engine!B17, 0)',
    0,
    0.0
]])
add_batch("📄 Executive_OnePager", "B9:C14", [
    [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0]
])
add_batch("📄 Executive_OnePager", "F9:I13", [
    ["«Ожидание сделок из CRM»", "—", 0, "🟢 В НОРМЕ"],
    ["—", "—", 0, "—"],
    ["—", "—", 0, "—"],
    ["—", "—", 0, "—"],
    ["—", "—", 0, "—"]
])

# 2.10 ⚙️ Настройки
add_batch("⚙️ Настройки", "B3", [["ООО «Название Компании Клиента»"]])
add_batch("⚙️ Настройки", "E3", [["CLIENT-TENANT-001"]])
add_batch("⚙️ Настройки", "F3", [["(Пространство нового клиента)"]])
add_batch("⚙️ Настройки", "E4", [["RevOps OS V17.6 Client Edition"]])
add_batch("⚙️ Настройки", "B21:B24", [
    ["Менеджер 1 (РОП)"],
    ["Менеджер 2 (КАМ)"],
    ["Менеджер 3 (SDR)"],
    ["Менеджер 4 (SDR-2)"]
])

print(f"2. Applying batch updates for {len(batch_data)} ranges in 1 batch call...")
for attempt in range(1, 4):
    try:
        sh.values_batch_update({
            "valueInputOption": "USER_ENTERED",
            "data": batch_data
        })
        print("  -> Successfully applied all formulas, placeholders, and links!")
        break
    except Exception as e:
        print(f"  Attempt {attempt} failed: {e}. Retrying after 20s...")
        time.sleep(20)

print("\n🚀 COMPLETE RESET FINISHED! Checking calculated values...")
