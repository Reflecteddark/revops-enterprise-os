"""
Скрипт капитального обновления листа '⚙️ Настройки' (gid: 985355776)
в главной таблице продукта:
https://docs.google.com/spreadsheets/d/1QnjrrbpqhYssofchee7G06szWrFvBCcOqGIVokjqdVc

Обновляет:
1. Актуальные даты (Октябрь 2026, 01.10.2026 - 31.10.2026)
2. Версию платформы: RevOps Enterprise OS V17.6 Golden Master
3. Статус контура: 🟢 152-ФЗ РФ Compliant (Аттестованный закрытый контур)
4. Конструктор модулей: 8 актуальных модулей (Whisper Pro, Telegram алертер 30с, Автозаполнение CRM, Режим Адвокат, 13 B2B-стандартов)
5. Бенчмарки потерь (28% брак речи, 18% зависание КП, 35% возврат To-Be)
6. Премиальное корпоративное форматирование (Navy #0F172A, Blue #1E3A8A, Emerald #059669, Slate borders)
"""

import sys
import gspread

sys.stdout.reconfigure(encoding="utf-8")

SPREADSHEET_ID = "1QnjrrbpqhYssofchee7G06szWrFvBCcOqGIVokjqdVc"
SHEET_TITLE = "⚙️ Настройки"

client = gspread.service_account(filename="service_account.json")
sh = client.open_by_key(SPREADSHEET_ID)
ws = sh.worksheet(SHEET_TITLE)
sheet_id = ws.id

print(f"Connected to '{sh.title}' -> '{SHEET_TITLE}' (gid: {sheet_id})")

# Сборка эталонных данных строк для '⚙️ Настройки'
data = []

# Row 1: Заголовок
data.append(["⚙️ СИСТЕМНЫЕ КОНФИГУРАЦИИ, ПАРАМЕТРЫ КОНТУРА И РЕЕСТР ДОПУЩЕНИЙ REVOPS OS V17.6"] + [""] * 11)

# Row 2: Навигационный бар
data.append([
    "⚡ Экспресс-Калькулятор", "",
    "📋 Пульт РОПа", "",
    "🎙️ ИИ-Аудит звонков", "",
    "💸 7 Грехов ОП", "",
    "📄 Executive OnePager", "",
    "🔐 152-ФЗ Безопасность", ""
])

# Row 3: Раздел 1 - Параметры тенанта и компании
data.append(["🏢 ПАРАМЕТРЫ ТЕНАНТА И ПЕРИОД АНАЛИТИКИ"] + [""] * 4 + ["🔒 СТАТУС БЕЗОПАСНОСТИ И КОНТУРА (152-ФЗ)"] + [""] * 6)

# Row 4-10: Параметры системы
data.append(["Организация / Клиент", "RevOps Enterprise (Golden Master)", "", "🔑 TENANT UUID:", "TENANT-MASTER-001", "🛡️ КОНТУР БЕЗОПАСНОСТИ:", "🟢 152-ФЗ РФ Compliant (Аттестован)", "", "СЕССИЯ АНАЛИТИКИ:", "🟢 Активна (Online)", "", ""])
data.append(["Валюта расчетов", "RUB (₽)", "", "🏢 ТЕНАНТ / КОМПАНИЯ:", "RevOps Platform V17.6", "🔒 ШИФРОВАНИЕ ДАННЫХ:", "AES-256 + TLS 1.3 In-Transit", "", "РЕЖИМ ДОСТУПА:", "👑 CEO / Super-Admin", "", ""])
data.append(["Версия архитектуры", "RevOps Enterprise OS V17.6 Production", "", "⚡ БЫСТРЫЙ ПЕРИОД:", "Текущий месяц (Октябрь 2026)", "👤 ДЕПЕРСОНАЛИЗАЦИЯ:", "✅ Автоматическая маскировка биометрии", "", "СТАТУС ИИ-ЯДРА:", "🟢 Whisper Large V3 + DeepSeek-R1", "", ""])
data.append(["Период: начало", "01.10.2026", "", "📅 ДАТА СТАРТА СПРИНТА:", "01.10.2026", "📄 ЮРИДИЧЕСКИЙ NDA:", "✅ Подписан двусторонний NDA", "", "СКОРОСТЬ АЛЕРТОВ:", "⚡ 30 секунд в Telegram", "", ""])
data.append(["Период: конец", "31.10.2026", "", "⏱️ РАСЧЕТНЫХ ДНЕЙ В МЕС:", "31 день", "🏢 СЕРВЕРНЫЙ КОНТУР:", "🇷🇺 РФ Локализация (Дата-центры в РФ)", "", "СИНХРОНИЗАЦИЯ CRM:", "🟢 amoCRM + Битрикс24 Live", "", ""])
data.append(["План выручки на месяц", "5,000,000 ₽", "", "🎯 МИНИМАЛЬНЫЙ ПЛАН:", "4,200,000 ₽", "👥 МУЛЬТИ-РОЛЕВОЙ RBAC:", "Hardware Enforcement (Protected Sheets)", "", "ОБЪЕМ ЗВОНКОВ В СУТКИ:", "150 – 400 диалогов", "", ""])
data.append(["Статус спринта внедрения", "⚡ 7-дневный спринт активен", "", "💰 ЦЕЛЕВОЙ ROI СПРИНТА:", "367% (Окупаемость 6-14 дней)", "🛡️ РЕЖИМ АНТИСАБОТАЖА:", "✅ Режим «Адвокат менеджера» активен", "", "ФОРМАТ ОТЧЕТОВ:", "📱 Telegram-Шериф + Google Sheets", "", ""])

# Row 11: Пустая
data.append([""] * 12)

# Row 12: Заголовок секций воронки и сотрудников
data.append(["🎯 НОРМАТИВЫ ВОРОНКИ ПРОДАЖ И SLA"] + [""] * 4 + ["👥 СОТРУДНИКИ И ЛИМИТЫ НАГРУЗКИ ОП"] + [""] * 6)

# Row 13: Шапки таблиц воронки и сотрудников
data.append([
    "ID", "Название этапа воронки", "SLA (ч)", "Win %", "Статус",
    "ID", "Сотрудник ОП", "Роль", "Лимит сделок", "Загрузка", "Статус контроля ИИ", ""
])

# Rows 14-20: Данные воронки и сотрудников
stages_staff = [
    ("1", "1. Новый входящий лид", "2 ч", "5%", "🟢 Норма", "101", "Иванов Дмитрий (РОП)", "ROP", "5 сделок", "100%", "🟢 100% контроль качества"),
    ("2", "2. Квалификация BANT / ЛПР", "24 ч", "15%", "🟢 Норма", "102", "Петров Александр (КАМ)", "KAM", "25 сделок", "84%", "🟢 100% контроль качества"),
    ("3", "3. Проведение демонстрации / Zoom", "48 ч", "35%", "🟢 Норма", "103", "Сидоров Роман (SDR)", "SDR", "40 сделок", "92%", "🟢 100% контроль качества"),
    ("4", "4. КП и фиксация Next Step", "24 ч", "60%", "🟢 Норма", "104", "Козлов Алексей (SDR-2)", "SDR", "40 сделок", "78%", "🟢 100% контроль качества"),
    ("5", "5. Выставлен счет / Договор", "48 ч", "85%", "🟢 Норма", "105", "Смирнова Елена (Senior)", "CLOSER", "20 сделок", "95%", "🟢 100% контроль качества"),
    ("6", "6. Успешно реализовано (Won)", "0 ч", "100%", "🏁 Финал", "106", "Васильев Игорь (Junior)", "TRAINEE", "15 сделок", "60%", "🟢 Режим «Адвокат» включен"),
    ("7", "7. Закрыто и не реализовано (Lost)", "0 ч", "0%", "🚨 Анализ сливов", "-", "Средняя загрузка отдела", "ALL", "145 сделок", "85%", "🛡️ Защита от сливов активна")
]

for stg in stages_staff:
    data.append(list(stg) + [""])

# Row 21: Пустая
data.append([""] * 12)

# Row 22: Заголовок модульного конструктора
data.append(["📦 МОДУЛЬНЫЙ КОНСТРУКТОР ТАРИФОВ И ФИЧ REVOPS OS PRO V17.6"] + [""] * 11)

# Row 23: Шапка модулей
data.append([
    "Код", "Название модуля платформы", "Статус в контуре", "Тарифный уровень", "Стоимость, ₽/мес",
    "Параметр бенчмарка потерь", "Значение", "Комментарий и бизнес-влияние на выручку", "", "", "", ""
])

# Rows 24-31: 8 Модулей и Бенчмарки
modules_benchmarks = [
    ("MOD-01", "🎙️ Речевой ИИ-Супервайзер 100% звонков (Whisper + DeepSeek)", "ВКЛЮЧЕН", "Light", "29,000 ₽", "Потери на браке речи (% от выручки)", "28.0%", "28% потенциальной выручки сжигается на звонках из-за ошибок менеджеров"),
    ("MOD-02", "⚡ Мгновенные алерты в Telegram (30 секунд с планом спасения)", "ВКЛЮЧЕН", "Pro", "15,000 ₽", "Потери на зависание КП (% от выручки)", "18.0%", "18% выручки теряется при отправке КП на 'кладбище почты' без Next Step"),
    ("MOD-03", "🤖 Автозаполнение карточек CRM (Executive Summary + договоренности)", "ВКЛЮЧЕН", "Pro", "19,000 ₽", "Потери спящих отказников CRM", "12.0%", "12% отказных лидов можно вернуть в первый месяц через реанимацию"),
    ("MOD-04", "🛡️ Пакет «Антисаботаж» и Режим «Адвокат менеджера»", "ВКЛЮЧЕН", "Pro", "15,000 ₽", "Просрочка оплат и счетов (DSO)", "10.0%", "10% выставленных счетов зависают на согласовании без пуша"),
    ("MOD-05", "🎯 Контроль 13 B2B-стандартов речи и жесткого Next Step", "ВКЛЮЧЕН", "Pro", "25,000 ₽", "Конверсия воронки (As-Is)", "4.5%", "Текущая конверсия из нового лида в оплаченный счет"),
    ("MOD-06", "📊 Пульт РОПа: утренний лидерборд и рейтинг менеджеров", "ВКЛЮЧЕН", "Pro", "35,000 ₽", "Коэффициент возврата потерь (To-Be)", "35.0%", "35% выявленных сливов гарантированно спасается в первый месяц спринта"),
    ("MOD-07", "🔐 Контур безопасности 152-ФЗ и токенизация биометрии", "ВКЛЮЧЕН", "Enterprise", "25,000 ₽", "Целевой рост конверсии ОП", "+25 – 40%", "Рост конверсии отдела продаж при ликвидации базовых сливов речи"),
    ("MOD-08", "📄 Executive PDF One-Pager генератор для Собственника и CFO", "ВКЛЮЧЕН", "Enterprise", "20,000 ₽", "Суммарный подтвержденный ROI", "367% – 600%", "Окупаемость затрат на платформу за 6–14 календарных дней")
]

for m in modules_benchmarks:
    data.append(list(m) + ["", "", "", ""])

# Row 32: Итоговая строка
data.append([
    "ИТОГО", "Полный Enterprise-комплект RevOps OS Pro", "АКТИВЕН", "Enterprise VIP", "183,000 ₽",
    "СУММАРНАЯ УПУЩЕННАЯ ВЫГОДА:", "2,850,000 ₽ / мес", "Срок окупаемости пилота: 6–14 дней | ROI: 367%", "", "", "", ""
])

# Запись данных
ws.clear()
ws.update(range_name=f"A1:L{len(data)}", values=data)
print(f"Data written to '{SHEET_TITLE}'. Total rows: {len(data)}")

# Наводим премиальное форматирование через batch_update
def color_rgb(r, g, b):
    return {"red": r / 255.0, "green": g / 255.0, "blue": b / 255.0}

c_navy_dark = color_rgb(15, 23, 42)      # #0F172A
c_navy_hdr = color_rgb(30, 41, 59)       # #1E293B
c_blue_primary = color_rgb(30, 58, 138)  # #1E3A8A
c_blue_light = color_rgb(239, 246, 255)  # #EFF6FF
c_card_bg = color_rgb(248, 250, 252)     # #F8FAFC
c_emerald = color_rgb(5, 150, 105)       # #059669
c_emerald_light = color_rgb(236, 253, 245)# #ECFDF5
c_amber = color_rgb(217, 119, 6)         # #D97706
c_amber_light = color_rgb(255, 251, 235) # #FFFBEB
c_border = color_rgb(203, 213, 225)      # #CBD5E1
c_white = color_rgb(255, 255, 255)
c_text_dark = color_rgb(30, 41, 59)
c_text_muted = color_rgb(148, 163, 184)

requests = []

# Сброс старых объединений
requests.append({
    "unmergeCells": {
        "range": {
            "sheetId": sheet_id,
            "startRowIndex": 0,
            "endRowIndex": 40,
            "startColumnIndex": 0,
            "endColumnIndex": 15
        }
    }
})

def req_merge(r_s, r_e, c_s, c_e):
    return {
        "mergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": r_s,
                "endRowIndex": r_e,
                "startColumnIndex": c_s,
                "endColumnIndex": c_e
            },
            "mergeType": "MERGE_ALL"
        }
    }

def req_format(r_s, r_e, c_s, c_e, bg=None, fg=None, bold=False, size=9, halign="LEFT", valign="MIDDLE", wrap=True):
    cell_fmt = {
        "textFormat": {
            "fontFamily": "Segoe UI",
            "fontSize": int(round(size)),
            "bold": bold
        },
        "horizontalAlignment": halign,
        "verticalAlignment": valign,
        "wrapStrategy": "WRAP" if wrap else "OVERFLOW_CELL"
    }
    if bg:
        cell_fmt["backgroundColor"] = bg
    if fg:
        cell_fmt["textFormat"]["foregroundColor"] = fg

    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": r_s,
                "endRowIndex": r_e,
                "startColumnIndex": c_s,
                "endColumnIndex": c_e
            },
            "cell": {"userEnteredFormat": cell_fmt},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)"
        }
    }

# 1. Заголовок (R1)
requests.append(req_merge(0, 1, 0, 12))
requests.append(req_format(0, 1, 0, 12, bg=c_navy_dark, fg=c_white, bold=True, size=13, halign="LEFT"))

# 2. Навигационный бар (R2)
nav_ranges = [(0, 2), (2, 4), (4, 6), (6, 8), (8, 10), (10, 12)]
for cs, ce in nav_ranges:
    requests.append(req_merge(1, 2, cs, ce))
    requests.append(req_format(1, 2, cs, ce, bg=c_navy_hdr, fg=c_white, bold=True, size=9, halign="CENTER"))

# 3. Раздел 1 (R3)
requests.append(req_merge(2, 3, 0, 5))
requests.append(req_merge(2, 3, 5, 12))
requests.append(req_format(2, 3, 0, 5, bg=c_blue_primary, fg=c_white, bold=True, size=10, halign="LEFT"))
requests.append(req_format(2, 3, 5, 12, bg=c_navy_dark, fg=c_white, bold=True, size=10, halign="LEFT"))

# Параметры (R4-R10)
for r in range(3, 10):
    bg_r = c_card_bg if r % 2 == 1 else c_white
    requests.append(req_format(r, r + 1, 0, 12, bg=bg_r, fg=c_text_dark, size=8.5, halign="LEFT"))
    # Метки
    requests.append(req_format(r, r + 1, 0, 1, bg=bg_r, fg=c_text_dark, bold=True, size=8.5, halign="LEFT"))
    requests.append(req_format(r, r + 1, 3, 4, bg=c_blue_light, fg=c_blue_primary, bold=True, size=8.5, halign="LEFT"))
    requests.append(req_format(r, r + 1, 5, 6, bg=c_blue_light, fg=c_blue_primary, bold=True, size=8.5, halign="LEFT"))
    requests.append(req_format(r, r + 1, 8, 9, bg=c_blue_light, fg=c_blue_primary, bold=True, size=8.5, halign="LEFT"))

# 4. Раздел 2: Воронка и Сотрудники (R12-R20)
requests.append(req_merge(11, 12, 0, 5))
requests.append(req_merge(11, 12, 5, 12))
requests.append(req_format(11, 12, 0, 5, bg=c_blue_primary, fg=c_white, bold=True, size=10, halign="LEFT"))
requests.append(req_format(11, 12, 5, 12, bg=c_navy_dark, fg=c_white, bold=True, size=10, halign="LEFT"))

requests.append(req_format(12, 13, 0, 12, bg=c_navy_hdr, fg=c_white, bold=True, size=8.5, halign="CENTER"))

for r in range(13, 20):
    bg_r = c_card_bg if r % 2 == 1 else c_white
    requests.append(req_format(r, r + 1, 0, 12, bg=bg_r, fg=c_text_dark, size=8.5, halign="LEFT"))
    requests.append(req_format(r, r + 1, 0, 1, bg=c_blue_light, fg=c_blue_primary, bold=True, size=8.5, halign="CENTER"))
    requests.append(req_format(r, r + 1, 2, 4, bg=bg_r, fg=c_text_dark, bold=True, size=8.5, halign="CENTER"))
    requests.append(req_format(r, r + 1, 5, 6, bg=c_blue_light, fg=c_blue_primary, bold=True, size=8.5, halign="CENTER"))
    requests.append(req_format(r, r + 1, 9, 10, bg=bg_r, fg=c_text_dark, bold=True, size=8.5, halign="CENTER"))
    requests.append(req_format(r, r + 1, 10, 11, bg=c_emerald_light, fg=c_emerald, bold=True, size=8, halign="LEFT"))

# 5. Раздел 3: Модульный конструктор и бенчмарки (R22-R32)
requests.append(req_merge(21, 22, 0, 12))
requests.append(req_format(21, 22, 0, 12, bg=c_blue_primary, fg=c_white, bold=True, size=10, halign="LEFT"))

requests.append(req_merge(22, 23, 7, 12))
requests.append(req_format(22, 23, 0, 12, bg=c_navy_hdr, fg=c_white, bold=True, size=8.5, halign="CENTER"))

for r in range(23, 31):
    bg_r = c_card_bg if r % 2 == 1 else c_white
    requests.append(req_merge(r, r + 1, 7, 12))
    requests.append(req_format(r, r + 1, 0, 12, bg=bg_r, fg=c_text_dark, size=8.5, halign="LEFT"))
    requests.append(req_format(r, r + 1, 0, 1, bg=c_blue_light, fg=c_blue_primary, bold=True, size=8.5, halign="CENTER"))
    requests.append(req_format(r, r + 1, 2, 3, bg=c_emerald_light, fg=c_emerald, bold=True, size=8.5, halign="CENTER"))
    requests.append(req_format(r, r + 1, 4, 5, bg=c_amber_light, fg=c_amber, bold=True, size=8.5, halign="CENTER"))
    requests.append(req_format(r, r + 1, 6, 7, bg=c_blue_light, fg=c_blue_primary, bold=True, size=8.5, halign="CENTER"))

# Итоговая строка (R32)
requests.append(req_merge(31, 32, 7, 12))
requests.append(req_format(31, 32, 0, 12, bg=c_navy_dark, fg=c_white, bold=True, size=9.5, halign="LEFT"))
requests.append(req_format(31, 32, 4, 5, bg=c_amber, fg=c_white, bold=True, size=10, halign="CENTER"))
requests.append(req_format(31, 32, 6, 7, bg=c_emerald, fg=c_white, bold=True, size=10, halign="CENTER"))

# 6. Границы (Borders)
requests.append({
    "updateBorders": {
        "range": {
            "sheetId": sheet_id,
            "startRowIndex": 0,
            "endRowIndex": len(data),
            "startColumnIndex": 0,
            "endColumnIndex": 12
        },
        "top": {"style": "SOLID", "color": c_border},
        "bottom": {"style": "SOLID", "color": c_border},
        "left": {"style": "SOLID", "color": c_border},
        "right": {"style": "SOLID", "color": c_border},
        "innerHorizontal": {"style": "SOLID", "color": c_border},
        "innerVertical": {"style": "SOLID", "color": c_border}
    }
})

# 7. Ширины колонок
col_pixel_widths = [75, 230, 80, 140, 160, 150, 220, 130, 150, 140, 180, 60]
for col_idx, width in enumerate(col_pixel_widths):
    requests.append({
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "COLUMNS",
                "startIndex": col_idx,
                "endIndex": col_idx + 1
            },
            "properties": {"pixelSize": width},
            "fields": "pixelSize"
        }
    })

# 8. Высоты строк
requests.append({
    "updateDimensionProperties": {
        "range": {
            "sheetId": sheet_id,
            "dimension": "ROWS",
            "startIndex": 0,
            "endIndex": 1
        },
        "properties": {"pixelSize": 42},
        "fields": "pixelSize"
    }
})

sh.batch_update({"requests": requests})
print(f"🚀 Лист '⚙️ Настройки' успешно обновлен и отформатирован!")
print(f"URL: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit?gid={sheet_id}#gid={sheet_id}")
