import sys
import gspread

sys.stdout.reconfigure(encoding="utf-8")

DST_ID = "1jBBotOfFh-XEJGScJyQi10jiFrna66OpJ91yPDrXy2A"
client = gspread.service_account(filename="service_account.json")
sh_dst = client.open_by_key(DST_ID)

print(f"Connected to '{sh_dst.title}' ({DST_ID})")

# 1. Clean raw sheets (keep Row 1 headers, clear rows 2+)
raw_sheets = [
    "raw_deals",
    "raw_calls",
    "raw_touchpoints",
    "raw_invoices",
    "raw_payments",
    "raw_marketing",
    "raw_alerts",
    "raw_audit_log"
]

for s_name in raw_sheets:
    try:
        ws = sh_dst.worksheet(s_name)
        headers = ws.row_values(1)
        row_count = ws.row_count
        col_count = len(headers)
        
        # Batch clear from row 2 onwards
        if row_count > 1:
            ws.batch_clear([f"A2:Z{row_count}"])
            print(f"Cleaned '{s_name}' (headers preserved: {len(headers)} columns)")
    except Exception as e:
        print(f"Notice on '{s_name}': {e}")

# 2. Update ⚙️ Настройки with client placeholders
try:
    ws_set = sh_dst.worksheet("⚙️ Настройки")
    ws_set.update(range_name="B3", values=[["ООО «Название Компании Клиента»"]])
    ws_set.update(range_name="E3", values=[["CLIENT-TENANT-001"]])
    ws_set.update(range_name="F3", values=[["(Рабочее пространство клиента)"]])
    ws_set.update(range_name="B4", values=[["RUB (₽)"]])
    ws_set.update(range_name="E4", values=[["Пространство Клиента (Production)"]])
    ws_set.update(range_name="B8", values=[["01.10.2026"]])
    ws_set.update(range_name="B9", values=[["3,000,000 ₽"]])
    print("Updated ⚙️ Настройки with clean client placeholders.")
except Exception as e:
    print(f"Notice on ⚙️ Настройки: {e}")

# 3. Create or Update 🚀 Быстрый_Старт_Клиента sheet
guide_sheet_title = "🚀 Быстрый_Старт_Клиента"
try:
    try:
        ws_guide = sh_dst.worksheet(guide_sheet_title)
    except:
        ws_guide = sh_dst.add_worksheet(title=guide_sheet_title, rows=30, cols=10, index=0)

    guide_content = [
        ["🚀 REVOPS ENTERPRISE OS V17.6 — ПАМЯТКА БЫСТРОГО СТАРТА НОВОГО КЛИЕНТА", "", "", "", "", "", ""],
        ["Инструкция по развертыванию рабочего контура, подключению CRM и запуску 100% ИИ-контроля звонков за 15 минут", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["ШАГ 1: АКТИВАЦИЯ И СОХРАНЕНИЕ КОПИИ", "", "", "", "ШАГ 2: БАЗОВЫЕ НАСТРОЙКИ КОМПАНИИ", "", ""],
        ["Действие:", "В меню Google Sheets выберите: «Файл» ➔ «Создать копию».", "", "", "Действие:", "Откройте лист «⚙️ Настройки» (вкладка 6).", ""],
        ["Зачем:", "Чтобы создать собственное изолированное рабочее пространство вашей компании.", "", "", "Зачем:", "Укажите название компании, месяц и целевой план выручки (B9).", ""],
        ["Срок:", "1 минута", "", "", "Срок:", "2 минуты", ""],
        ["", "", "", "", "", "", ""],
        ["ШАГ 3: ПОДКЛЮЧЕНИЕ CRM (amoCRM / БИТРИКС24)", "", "", "", "ШАГ 4: ЗАПУСК TELEGRAM-СУПЕРВАЙЗЕРА", "", ""],
        ["Действие:", "Вставьте API-ключ вашей CRM в модуль интеграции (MOD-03).", "", "", "Действие:", "Подключите Telegram-бота @revops_supervisor_bot.", ""],
        ["Результат:", "Сделки и записи звонков начнут автоматически поступать в raw_deals и raw_calls.", "", "", "Результат:", "РОП и директор начнут получать алерты за 30с со сценариями спасения сделок.", ""],
        ["Срок:", "5-10 минут", "", "", "Срок:", "2 минуты", ""],
        ["", "", "", "", "", "", ""],
        ["💎 ВАШИ КЛЮЧЕВЫЕ РАБОЧИЕ ДАШБОРДЫ (НАВИГАЦИЯ ПО ВКЛАДКАМ):", "", "", "", "", "", ""],
        ["№", "Название вкладки", "Для кого предназначена", "Какую бизнес-задачу решает", "", "", ""],
        ["1", "⚡ Экспресс_Калькулятор_3_Цифры", "Собственник, CEO, РОП", "Моделирование выручки, оцифровка 4 точек сливов (брак звонков 28%, зависание КП 18%)", "", "", ""],
        ["2", "📋 Пульт_РОПа_15_Минут", "РОП, Руководитель группы", "Ежедневный утренний пульт: ТОП-5 горящих сделок в риске + готовые скрипты перехвата", "", "", ""],
        ["3", "🎙️ ИИ_Аудит", "РОП, Контроль качества (ОКК)", "100% аудит звонков нейросетью Whisper + DeepSeek по 13 жестким стандартам речи", "", "", ""],
        ["4", "💸 Диагностика_Утечек_ОП", "CFO, Финансовый директор", "Аудит 7 смертных грехов продаж + расчет чистой окупаемости спринта (ROI 367%)", "", "", ""],
        ["5", "📄 Executive_OnePager", "Собственник, Совет директоров", "Сводная витрина одного экрана: план/факт, активный пайплайн, риски выручки", "", "", ""],
        ["6", "⚙️ Настройки", "Администратор системы", "Модульный конструктор тарифов (8 фич), бенчмарки потерь, параметры контура 152-ФЗ", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["🔒 БЕЗОПАСНОСТЬ И КОНФИДЕНЦИАЛЬНОСТЬ (152-ФЗ РФ):", "", "", "", "", "", ""],
        ["• Все данные обрабатываются в закрытом защищенном контуре в дата-центрах РФ.", "", "", "", "", "", ""],
        ["• Перед анализом нейросетью биометрия и персональные данные клиентов деперсонализируются.", "", "", "", "", "", ""],
        ["• Доступ к сырым таблицам защищен на уровне ролевой модели (RBAC).", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["📞 НУЖНА ПОМОЩЬ С НАСТРОЙКОЙ ИЛИ ПОДКЛЮЧЕНИЕМ API?", "", "", "", "", "", ""],
        ["Напишите в службу заботы RevOps OS: Telegram: @revops_support | Сайт: https://ai-rop.ru", "", "", "", "", "", ""]
    ]

    ws_guide.update(range_name=f"A1:G{len(guide_content)}", values=guide_content)
    print("Created 🚀 Быстрый_Старт_Клиента guide sheet!")
except Exception as e:
    print(f"Notice on guide sheet: {e}")

# Reorder presentation tabs with Guide at front (index 0)
clean_presentation_order = [
    "🚀 Быстрый_Старт_Клиента",
    "⚡ Экспресс_Калькулятор_3_Цифры",
    "📋 Пульт_РОПа_15_Минут",
    "🎙️ ИИ_Аудит",
    "💸 Диагностика_Утечек_ОП",
    "📄 Executive_OnePager",
    "⚙️ Настройки"
]

reorder_reqs = []
for new_idx, title in enumerate(clean_presentation_order):
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
    print("Ordered clean template presentation tabs 0 to 6.")

print("All clean client template configurations successfully completed!")
