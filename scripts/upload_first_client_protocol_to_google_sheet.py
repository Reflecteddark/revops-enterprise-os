"""
Скрипт загрузки регламента '🚀 Протокол_Первого_Клиента' в Google Таблицу:
https://docs.google.com/spreadsheets/d/1AbaQGHhzar7Qsb9U9TByN57rAmOn36zw4HOD2XVrmTc

Создает новый лист с идеальным корпоративным стилем McKinsey / RevOps OS:
- Топовые карточки KPI
- 12 пошаговых боевых шагов от лида до подписки
- Готовые шаблоны сообщений MSG-01..05 для копирования
- Отработка 3 главных возражений
- Чек-лист готовности к созвону
"""

import sys
import gspread

sys.stdout.reconfigure(encoding="utf-8")

SPREADSHEET_ID = "1AbaQGHhzar7Qsb9U9TByN57rAmOn36zw4HOD2XVrmTc"
SHEET_TITLE = "🚀 Протокол_Первого_Клиента"

client = gspread.service_account(filename="service_account.json")
sh = client.open_by_key(SPREADSHEET_ID)

# Проверяем или создаем лист
try:
    ws = sh.worksheet(SHEET_TITLE)
    print(f"Sheet '{SHEET_TITLE}' already exists. Rebuilding...")
except gspread.WorksheetNotFound:
    ws = sh.add_worksheet(title=SHEET_TITLE, rows=60, cols=10)
    print(f"Sheet '{SHEET_TITLE}' created.")

sheet_id = ws.id

# 1. Формируем массив данных строк
data = []

# Row 1: Заголовок
data.append(["🚀 RevOps Enterprise OS | Пошаговый Регламент Ведения Первого Клиента (End-to-End SOP)"] + [""] * 7)

# Row 2: Подзаголовок
data.append(["Полная инструкция действий основателя: от входящей заявки до аудита 3 звонков за 24ч, закрытия пилота 45к-150к ₽ и абонентского сопровождения"] + [""] * 7)

# Row 3: Пустая
data.append([""] * 8)

# Row 4-5: KPI Карточки
data.append([
    "⚡ СКОРОСТЬ ПЕРВОГО ОТВЕТА", "",
    "⏱️ ФОРМАТ ПЕРВОГО ДЕМО", "",
    "🎯 ГЛАВНЫЙ ВХОДНОЙ ОФФЕР", "",
    "💰 ЦЕЛЕВОЙ ЧЕК ПИЛОТА", ""
])
data.append([
    "< 10 минут (в WhatsApp/TG)", "",
    "Ровно 20 минут в Zoom", "",
    "Тест 3 звонков за 24ч (0 ₽)", "",
    "45 000 – 150 000 ₽", ""
])

# Row 6: Пустая
data.append([""] * 8)

# Row 7: Секция 1
data.append(["📋 ПОШАГОВЫЙ БОЕВОЙ АЛГОРИТМ: 12 ШАГОВ ВЕДЕНИЯ КЛИЕНТА"] + [""] * 7)

# Row 8: Шапка таблицы шагов
data.append([
    "Шаг", "Фаза воронки", "Срок / Дедлайн", "Что конкретно делает фаундер",
    "Инструмент / Где лежит", "Что сказать / Написать (скрипт)", "Критерий успеха (DoD)", "Опасная ошибка (Чего нельзя делать)"
])

# Rows 9-20: 12 Шагов
steps = [
    (
        "Шаг 1", "Входящий лид", "0 – 10 минут",
        "Лид упал с сайта ai-rop.ru на почту info@ai-rop.ru или в Telegram @dm1918. Моментально пишем клиенту. Проводим экспресс-квалификацию по 3 вопросам.",
        "Почта / Telegram @dm1918",
        "«[Имя], добрый день! Получил вашу заявку на ai-rop.ru. Подскажите, сколько у вас менеджеров в отделе, какой средний чек и какую CRM используете?»",
        "Получен ответ: подтверждено от 3 менеджеров и средний чек от 50 000 ₽.",
        "Ждать больше 15 минут; отправлять громоздкие КП и коммерческие предложения в первом сообщении."
    ),
    (
        "Шаг 2", "Назначение демо", "В течение 1 часа",
        "Предложить 2 конкретных слота времени на выбор для 20-минутного Zoom-разбора. Скинуть ссылку на встречу и внести в календарь.",
        "Google Календарь / Zoom / Яндекс.Телемост",
        "«Отлично! Предлагаю созвониться на 20 минут в Zoom: покажу на пальцах, как ИИ слушает 100% звонков и находит потери. Завтра удобно в 11:30 или в 15:00?»",
        "Встреча подтверждена клиентом и стоит в календаре.",
        "Спрашивать: «Когда вам удобно?» — это размазывает договоренность."
    ),
    (
        "Шаг 3", "20-мин Презентация", "День 2 (строго 20 мин)",
        "Провести созвон строго по готовому Word-сценарию. Открыть сайт ai-rop.ru, Excel-калькулятор (3 числа) и Google Таблицу (13 стандартов B2B).",
        "Сценарий_20_минутной_Презентации_RevOps.docx",
        "Вести разговор по файлу: 3 мин боль -> 4 мин Калькулятор потерь -> 6 мин Живой аудит -> 3 мин 152-ФЗ -> 4 мин Оффер на тест 3 звонков.",
        "Клиент в шоке от суммы потенциальных потерь и согласен на тест 3 звонков.",
        "Пытаться продать годовой софт за 300к на первой встрече. Цель первой встречи — ТОЛЬКО получить 3 звонка!"
    ),
    (
        "Шаг 4", "Забор 3 звонков", "Сразу после созвона",
        "Отправить персональную ссылку на облачную папку для загрузки аудио (Яндекс.Диск/Google Drive) или предложить прислать файлы прямо в Telegram.",
        "Яндекс.Диск / Telegram @dm1918",
        "«[Имя], спасибо за созвон! Вот ссылка на закрытую папку: [Ссылка]. Загрузите 3 любых звонка за вчера (1 успешный и 2 слитых). Ровно через 24ч верну аудит».",
        "3 аудиофайла (.mp3 или .wav) сохранены на вашем компьютере.",
        "Давать сложные технические инструкции по API телефонии. Брать аудиофайлы простым архивом!"
    ),
    (
        "Шаг 5", "Защита 152-ФЗ", "По запросу клиента",
        "Если клиент или его СБ сомневаются в конфиденциальности — мгновенно отправить подписанный NDA и памятку по 152-ФЗ.",
        "Папка «152 ФЗ» на Рабочем столе",
        "«Мы работаем строго по 152-ФЗ РФ. Перед началом подписываем двусторонний NDA. Данные деперсонализируются в защищенном контуре. Направляю проект NDA».",
        "NDA подписан с обеих сторон (скан или ЭДО).",
        "Спорить со службой безопасности; говорить, что данные отправляются на зарубежные сервера."
    ),
    (
        "Шаг 6", "Fulfillment (Запуск ИИ)", "В течение 4–6 часов",
        "Запустить Start_Sales_Calls_AI.bat (Faster-Whisper на 8000, DeepSeek-R1 на 11434, n8n на 5678). Прогнать 3 аудиозаписи через транскрибацию и скоринг.",
        "Start_Sales_Calls_AI.bat, transcriber_service",
        "Автоматический процесс: получение расшифровки с тайм-кодами и скоринга по 13 критериям B2B.",
        "Получены точные транскрипты и баллы соответствия по всем 13 стандартам.",
        "Оставлять сырой нечитаемый текст без разбивки на реплики менеджера и клиента."
    ),
    (
        "Шаг 7", "Сборка отчета", "В течение 8 часов",
        "Сделать копию мастер-таблицы Google Sheets. Заполнить вкладку «ИИ_Аудит»: вставить цитаты, выставить зеленые/красные чипы, рассчитать недополученную маржу.",
        "Google Sheets Master (gid 852624872)",
        "Форматирование в корпоративном стиле: чипы ошибок, выделение желтым зон роста, формулы ROI.",
        "Готовая брендированная ссылка на Google Таблицу с правами «Просмотр» для клиента.",
        "Использовать неаккуратные шрифты или оставлять ошибки '###' в ячейках."
    ),
    (
        "Шаг 8", "Human-in-the-Loop", "За 2 часа до сдачи",
        "Фаундер лично переслушивает 2 ключевых спорных фрагмента. Формулирует 2 убийственных инсайта (например: 'Менеджер дал скидку 100к без торга').",
        "Google Таблица + блокнот инсайтов",
        "«Инсайт 1: Менеджер Александр не выявил ЛПР и отпустил клиента без даты повторного звонка. Инсайт 2: Слив скидки 15% снизил маржинальность на 240 000 ₽».",
        "Готовы 2 железобетонных факта потерь в деньгах, с которыми собственник не сможет поспорить.",
        "Полагаться на 100% только на авто-генерацию ИИ без личной проверки спикером."
    ),
    (
        "Шаг 9", "Презентация аудита", "Ровно через 24 часа",
        "Созвон на 15 минут. Демонстрация экрана: «Вот ваши 3 звонка, вот где менеджеры слили 400 000 ₽». Показ готовой Google Таблицы.",
        "Google Таблица клиента + Zoom",
        "«[Имя], посмотрите строку 12: клиент спросил про рассрочку, менеджер ответил 'нет' и положил трубку. Здесь компания потеряла 350к. И это только 1 звонок из 3!»",
        "Клиент признает: аудит объективен, проблема реальна и стоит больших денег.",
        "Отправлять ссылку молча в мессенджер без личной демонстрации на созвоне."
    ),
    (
        "Шаг 10", "Закрытие на пилот", "Финальные 5 мин созвона",
        "Презентовать 7-дневный спринт: 100% звонков отдела за 45 000 – 75 000 ₽ (или спринт на месяц за 150 000 ₽) с гарантией окупаемости.",
        "Спецификация пилота / Памятка 7 дней",
        "«Мы отслушали 3 звонка и нашли слив на 400к. За 7 дней мы подключим 100% звонков отдела, дадим дашборд РОПу и спасем от 1 млн ₽. Стоимость пилота — 55 000 ₽. Стартуем?»",
        "Устное «ДА» от собственника и запрос счета/договора.",
        "Делать скидки при первом сомнении. Держать цену твердо, опираясь на ROI калькулятора."
    ),
    (
        "Шаг 11", "Оплата и договор", "В течение 24 часов",
        "Выставить счет от ИП/Самозанятого с QR-кодом. Направить типовой договор оказания услуг. Проконтролировать приход денег.",
        "Банк-клиент (Точка / Тинькофф / Сбер)",
        "«[Имя], направил договор и счет с QR-кодом на почту и в WhatsApp. Как только бухгалтерия проведет оплату, мы за 2 часа запускаем коннектор звонков».",
        "Деньги поступили на расчетный счет. Подписан скан договора.",
        "Начинать полноценную работу без предоплаты."
    ),
    (
        "Шаг 12", "Сопровождение и LTV", "Дни 1 – 7 пилота",
        "Ежедневная выгрузка звонков, генерация утренней сводки в Telegram для РОПа в 09:00. В пятницу — итоговый отчет и продажа абонентской подписки (60к-120к/мес).",
        "Telegram-бот @RevOps_Super_Audit_Bot",
        "«За 7 дней мы разобрали 420 звонков, подняли показатель жесткого следующего шага с 32% до 78%, спасли 3 сделки на 1.2 млн ₽. Переходим на помесячный контракт?»",
        "Подписан долгосрочный контракт на регулярную аналитику (MRR от 60 000 ₽).",
        "Оставлять клиента без ежедневного контакта; не показывать динамику улучшений."
    )
]

for s in steps:
    data.append(list(s))

# Row 21: Пустая
data.append([""] * 8)

# Row 22: Секция 2 - Шаблоны сообщений
data.append(["💬 ШАБЛОНЫ БЫСТРЫХ СООБЩЕНИЙ ДЛЯ КОПИРОВАНИЯ (FAST COPY-PASTE)"] + [""] * 7)

# Row 23: Шапка шаблонов
data.append(["Код", "Событие", "Канал", "Готовый текст сообщения для клиента", "", "", "", ""])

# Rows 24-28: Шаблоны
templates = [
    (
        "MSG-01", "Первый ответ на заявку с сайта", "Telegram / WhatsApp",
        "«[Имя], приветствую! Получил вашу заявку на сервисе ai-rop.ru по аудиту звонков отдела продаж. Подскажите, пожалуйста: 1) Сколько менеджеров сейчас в линии? 2) Какой средний чек сделки? И какую CRM используете (amoCRM или Битрикс24)? Хочу подготовить релевантные примеры под вашу нишу перед созвоном»."
    ),
    (
        "MSG-02", "Напоминание за 1 час до Zoom-демо", "Telegram / WhatsApp",
        "«[Имя], добрый день! Напоминаю, что через 1 час (в [Время]) у нас запланирован 20-минутный созвон в Zoom. Ссылка на встречу: [Ссылка]. Покажу на живых цифрах, как работает ИИ-аудит и сколько выручки теряется в звонках. До встречи!»"
    ),
    (
        "MSG-03", "Инструкция по передаче 3 звонков", "Telegram / WhatsApp",
        "«[Имя], спасибо за встречу! Как договорились, жду от вас 3 любых аудиозаписи за вчера (один успешный звонок и два с сорвавшимися сделками). Загрузить можно сюда: [Ссылка на Яндекс.Диск] либо скинуть аудиофайлами прямо в этот чат. Завтра в [Время] верну готовую аналитику по 13 критериям»."
    ),
    (
        "MSG-04", "Уведомление о готовности аудита", "Telegram / WhatsApp",
        "«[Имя], добрый день! Аудит ваших 3 звонков полностью готов. Нашли 2 критические системные ошибки менеджеров, из-за которых сорвались сделки на суммарно ~[Сумма] ₽. Собрал всё в персональную Google Таблицу. Предлагаю созвониться на 15 минут — выведу экран и покажу конкретные цитаты и точки слива. Удобно в [Время]?»"
    ),
    (
        "MSG-05", "Отправка счета на 7-дневный пилот", "Email + Telegram",
        "«[Имя], направил вам договор и счет на 7-дневный пилот RevOps (сумма 55 000 ₽). В стоимость входит анализ 100% звонков вашего отдела, ежедневная утренняя сводка РОПу в Telegram и оцифровка точек потерь. Оплатить можно по реквизитам или по QR-коду в счете. Стартуем сразу после подтверждения!»"
    )
]

for code, ev, ch, txt in templates:
    data.append([code, ev, ch, txt, "", "", "", ""])

# Row 29: Пустая
data.append([""] * 8)

# Row 30: Секция 3 - Отработка возражений
data.append(["🛡️ ШПАРГАЛКА: ОТРАБОТКА 3 ГЛАВНЫХ ВОЗРАЖЕНИЙ НА СОЗВОНЕ"] + [""] * 7)

# Row 31: Шапка возражений
data.append(["Возражение клиента", "Истинная причина сомнения", "", "Убойный ответ спикера (слово в слово)", "", "", "", ""])

# Rows 32-34: Возражения
objs = [
    (
        "«У нас РОП сам каждый день слушает звонки»",
        "Оправдывает текущие затраты на ФОТ руководителя.",
        "«Сколько звонков успевает прослушать РОП? 3–5 из 100 — это меньше 5% выборки! При этом РОП тратит 25 часов в неделю на рутину вместо дожима сделок. Наш сервис слушает 100% звонков за секунды и приносит РОПу список из 4 проблемных сделок. Мы не заменяем РОПа, мы даем ему суперсилу»."
    ),
    (
        "«У нас специфический бизнес, ИИ не разберется»",
        "Боится, что алгоритм ошибется в терминах.",
        "«Именно поэтому мы не используем шаблонные решения. Перед аудитом мы заносим в промпт ваш словарь терминов, маржинальные позиции и стоп-слова. На тесте 3 звонков вы сами лично оцените точность терминологии»."
    ),
    (
        "«Пришлите КП на почту, мы подумаем»",
        "Вежливый способ слить встречу.",
        "«[Имя], КП я обязательно отправлю. Но без реальных записей вашей компании это останется абстрактным файлом. Скиньте прямо сейчас 3 звонка — и через 24 часа вы получите КП уже с оцифрованными потерями вашей команды. Это бесплатно»."
    )
]

for ob, reason, resp in objs:
    data.append([ob, reason, "", resp, "", "", "", ""])

# Row 35: Пустая
data.append([""] * 8)

# Row 36: Секция 4 - Чек-лист готовности
data.append(["✅ ЧЕК-ЛИСТ ГОТОВНОСТИ ЗА 5 МИНУТ ДО СОЗВОНА"] + [""] * 7)

# Rows 37-41: Чек-лист
checklist = [
    ("☑ 1", "Сайт https://ai-rop.ru открыт во вкладке браузера (подтверждает надежность и статус компании)"),
    ("☑ 2", "Открыта Google Таблица RevOps Master на вкладке '🎙️ ИИ_Аудит' со светофором из 13 критериев"),
    ("☑ 3", "Открыт файл 'Презентационная_Таблица_RevOps.xlsx' на листе 'Калькулятор_ROI' для мгновенного расчета потерь"),
    ("☑ 4", "Создана и скопирована персональная ссылка на защищенную папку Яндекс.Диска для приема 3 аудиозаписей"),
    ("☑ 5", "Рядом открыт блокнот с шаблоном NDA (152-ФЗ) и карточкой банковских реквизитов для выставления счета")
]

for ch_id, ch_desc in checklist:
    data.append([ch_id, ch_desc, "", "", "", "", "", ""])

# Очищаем лист и загружаем данные
ws.clear()
ws.update(range_name="A1:H41", values=data)
print("Data written successfully. Total rows:", len(data))

# 2. Форматирование через Google Sheets batchUpdate API
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
c_rose_light = color_rgb(255, 241, 242)  # #FFF1F2
c_rose = color_rgb(225, 29, 72)          # #E11D48
c_border = color_rgb(203, 213, 225)      # #CBD5E1

c_white = color_rgb(255, 255, 255)
c_text_dark = color_rgb(30, 41, 59)
c_text_muted = color_rgb(148, 163, 184)

requests = []

# Очистить все предыдущие объединения ячеек
requests.append({
    "unmergeCells": {
        "range": {
            "sheetId": sheet_id,
            "startRowIndex": 0,
            "endRowIndex": 50,
            "startColumnIndex": 0,
            "endColumnIndex": 10
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

def req_format(r_s, r_e, c_s, c_e, bg=None, fg=None, bold=False, size=10, halign="LEFT", valign="MIDDLE", wrap=True):
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

# 1. Заголовки (R1-R2)
requests.append(req_merge(0, 1, 0, 8))
requests.append(req_format(0, 1, 0, 8, bg=c_navy_dark, fg=c_white, bold=True, size=13, halign="LEFT", valign="MIDDLE"))

requests.append(req_merge(1, 2, 0, 8))
requests.append(req_format(1, 2, 0, 8, bg=c_navy_dark, fg=c_text_muted, bold=False, size=9, halign="LEFT", valign="MIDDLE"))

# 2. KPI Карточки (R4-R5)
requests.append(req_merge(3, 4, 0, 2))
requests.append(req_merge(3, 4, 2, 4))
requests.append(req_merge(3, 4, 4, 6))
requests.append(req_merge(3, 4, 6, 8))

requests.append(req_merge(4, 5, 0, 2))
requests.append(req_merge(4, 5, 2, 4))
requests.append(req_merge(4, 5, 4, 6))
requests.append(req_merge(4, 5, 6, 8))

requests.append(req_format(3, 4, 0, 4, bg=c_blue_light, fg=c_blue_primary, bold=True, size=9, halign="CENTER"))
requests.append(req_format(3, 4, 4, 6, bg=c_emerald_light, fg=c_emerald, bold=True, size=9, halign="CENTER"))
requests.append(req_format(3, 4, 6, 8, bg=c_amber_light, fg=c_amber, bold=True, size=9, halign="CENTER"))

requests.append(req_format(4, 5, 0, 4, bg=c_white, fg=c_blue_primary, bold=True, size=11, halign="CENTER"))
requests.append(req_format(4, 5, 4, 6, bg=c_white, fg=c_emerald, bold=True, size=11, halign="CENTER"))
requests.append(req_format(4, 5, 6, 8, bg=c_white, fg=c_amber, bold=True, size=11, halign="CENTER"))

# 3. Секция 1: 12 Шагов (R7-R20)
requests.append(req_merge(6, 7, 0, 8))
requests.append(req_format(6, 7, 0, 8, bg=c_card_bg, fg=c_text_dark, bold=True, size=11, halign="LEFT"))

# Шапка таблицы (R8)
requests.append(req_format(7, 8, 0, 8, bg=c_blue_primary, fg=c_white, bold=True, size=9, halign="CENTER"))

# Строки 12 шагов (R9-R20)
for r_idx in range(8, 20):
    bg_r = c_card_bg if r_idx % 2 == 0 else c_white
    requests.append(req_format(r_idx, r_idx + 1, 0, 8, bg=bg_r, fg=c_text_dark, size=8.5, halign="LEFT"))
    # Шаг бейдж
    requests.append(req_format(r_idx, r_idx + 1, 0, 1, bg=c_blue_light, fg=c_blue_primary, bold=True, size=9, halign="CENTER"))
    # Фаза
    requests.append(req_format(r_idx, r_idx + 1, 1, 2, bg=bg_r, fg=c_text_dark, bold=True, size=9, halign="LEFT"))
    # Срок
    requests.append(req_format(r_idx, r_idx + 1, 2, 3, bg=bg_r, fg=c_text_dark, bold=True, size=8.5, halign="CENTER"))
    # Ошибка (Колонка H)
    requests.append(req_format(r_idx, r_idx + 1, 7, 8, bg=c_rose_light, fg=c_rose, bold=True, size=8, halign="LEFT"))

# 4. Секция 2: Шаблоны сообщений (R22-R28)
requests.append(req_merge(21, 22, 0, 8))
requests.append(req_format(21, 22, 0, 8, bg=c_card_bg, fg=c_text_dark, bold=True, size=11, halign="LEFT"))

requests.append(req_merge(22, 23, 3, 8))
requests.append(req_format(22, 23, 0, 8, bg=c_navy_hdr, fg=c_white, bold=True, size=9, halign="CENTER"))

for r_t in range(23, 28):
    requests.append(req_merge(r_t, r_t + 1, 3, 8))
    bg_t = c_card_bg if r_t % 2 == 1 else c_white
    requests.append(req_format(r_t, r_t + 1, 0, 1, bg=c_blue_light, fg=c_blue_primary, bold=True, size=9, halign="CENTER"))
    requests.append(req_format(r_t, r_t + 1, 1, 2, bg=bg_t, fg=c_text_dark, bold=True, size=8.5, halign="LEFT"))
    requests.append(req_format(r_t, r_t + 1, 2, 3, bg=bg_t, fg=c_text_dark, size=8.5, halign="CENTER"))
    requests.append(req_format(r_t, r_t + 1, 3, 8, bg=bg_t, fg=c_navy_dark, size=8.5, halign="LEFT"))

# 5. Секция 3: Отработка возражений (R30-R34)
requests.append(req_merge(29, 30, 0, 8))
requests.append(req_format(29, 30, 0, 8, bg=c_card_bg, fg=c_text_dark, bold=True, size=11, halign="LEFT"))

requests.append(req_merge(30, 31, 1, 3))
requests.append(req_merge(30, 31, 3, 8))
requests.append(req_format(30, 31, 0, 8, bg=c_navy_hdr, fg=c_white, bold=True, size=9, halign="CENTER"))

for r_o in range(31, 34):
    requests.append(req_merge(r_o, r_o + 1, 1, 3))
    requests.append(req_merge(r_o, r_o + 1, 3, 8))
    bg_o = c_card_bg if r_o % 2 == 1 else c_white
    requests.append(req_format(r_o, r_o + 1, 0, 1, bg=bg_o, fg=c_rose, bold=True, size=8.5, halign="LEFT"))
    requests.append(req_format(r_o, r_o + 1, 1, 3, bg=bg_o, fg=c_text_dark, size=8.5, halign="LEFT"))
    requests.append(req_format(r_o, r_o + 1, 3, 8, bg=bg_o, fg=c_navy_dark, size=8.5, halign="LEFT"))

# 6. Секция 4: Чек-лист (R36-R41)
requests.append(req_merge(35, 36, 0, 8))
requests.append(req_format(35, 36, 0, 8, bg=c_card_bg, fg=c_text_dark, bold=True, size=11, halign="LEFT"))

for r_c in range(36, 41):
    requests.append(req_merge(r_c, r_c + 1, 1, 8))
    bg_c = c_card_bg if r_c % 2 == 1 else c_white
    requests.append(req_format(r_c, r_c + 1, 0, 1, bg=c_emerald_light, fg=c_emerald, bold=True, size=10, halign="CENTER"))
    requests.append(req_format(r_c, r_c + 1, 1, 8, bg=bg_c, fg=c_text_dark, size=9, halign="LEFT"))

# 7. Сетки границ (Borders)
requests.append({
    "updateBorders": {
        "range": {
            "sheetId": sheet_id,
            "startRowIndex": 0,
            "endRowIndex": 41,
            "startColumnIndex": 0,
            "endColumnIndex": 8
        },
        "top": {"style": "SOLID", "color": c_border},
        "bottom": {"style": "SOLID", "color": c_border},
        "left": {"style": "SOLID", "color": c_border},
        "right": {"style": "SOLID", "color": c_border},
        "innerHorizontal": {"style": "SOLID", "color": c_border},
        "innerVertical": {"style": "SOLID", "color": c_border}
    }
})

# 8. Ширины колонок
col_pixel_widths = [80, 140, 110, 280, 200, 320, 220, 220]
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

# 9. Высоты строк
requests.append({
    "updateDimensionProperties": {
        "range": {
            "sheetId": sheet_id,
            "dimension": "ROWS",
            "startIndex": 0,
            "endIndex": 1
        },
        "properties": {"pixelSize": 44},
        "fields": "pixelSize"
    }
})
requests.append({
    "updateDimensionProperties": {
        "range": {
            "sheetId": sheet_id,
            "dimension": "ROWS",
            "startIndex": 1,
            "endIndex": 2
        },
        "properties": {"pixelSize": 26},
        "fields": "pixelSize"
    }
})
# Высота шагов
requests.append({
    "updateDimensionProperties": {
        "range": {
            "sheetId": sheet_id,
            "dimension": "ROWS",
            "startIndex": 8,
            "endIndex": 20
        },
        "properties": {"pixelSize": 68},
        "fields": "pixelSize"
    }
})
# Высота шаблонов сообщений
requests.append({
    "updateDimensionProperties": {
        "range": {
            "sheetId": sheet_id,
            "dimension": "ROWS",
            "startIndex": 23,
            "endIndex": 28
        },
        "properties": {"pixelSize": 56},
        "fields": "pixelSize"
    }
})

sh.batch_update({"requests": requests})
print(f"Batch update completed successfully! GID: {sheet_id}")
print(f"URL: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit?gid={sheet_id}#gid={sheet_id}")
