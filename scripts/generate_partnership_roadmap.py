"""
Генератор Word-документа "Дорожная Карта Партнёрства RevOps AI Супервайзер"
RevOps Enterprise OS V17.6

Запуск: python generate_partnership_roadmap.py
Результат: RevOps_AI_Supervisor_Partnership_Roadmap.docx
"""

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from pathlib import Path
import shutil
import sys

# Обеспечиваем корректный UTF-8 вывод в Windows PowerShell
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

# === КОРПОРАТИВНЫЕ ЦВЕТА ===
BRAND_NAVY = RGBColor(0x1E, 0x3A, 0x5F)      # Тёмно-синий (заголовки)
BRAND_BLUE = RGBColor(0x4F, 0x46, 0xE5)      # Индиго (акценты)
BRAND_GRAY = RGBColor(0x4B, 0x55, 0x63)      # Серый (текст)
BRAND_LIGHT_GRAY = RGBColor(0xF3, 0xF4, 0xF6)  # Фон таблиц
BRAND_SUCCESS = RGBColor(0x05, 0x96, 0x69)   # Зелёный
BRAND_WARNING = RGBColor(0xDC, 0x26, 0x26)   # Красный


def set_cell_shading(cell, color_hex: str):
    """Устанавливает фон ячейки таблицы."""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)


def add_styled_table(doc, headers, rows, header_color="1E3A5F"):
    """Создаёт стилизованную таблицу."""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Заголовок
    header_row = table.rows[0]
    for i, header_text in enumerate(headers):
        cell = header_row.cells[i]
        cell.text = header_text
        set_cell_shading(cell, header_color)
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.size = Pt(10)

    # Строки данных
    for r_idx, row_data in enumerate(rows):
        row = table.rows[r_idx + 1]
        for c_idx, cell_text in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.text = str(cell_text)
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)
                    run.font.color.rgb = BRAND_GRAY
            # Чередуем фон строк
            if r_idx % 2 == 1:
                set_cell_shading(cell, "F3F4F6")

    doc.add_paragraph()  # отступ
    return table


def add_info_block(doc, emoji: str, text: str):
    """Информационный блок с рамкой (callout)."""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.right_indent = Cm(0.5)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)

    run = p.add_run(f"{emoji} {text}")
    run.font.size = Pt(10)
    run.font.italic = True
    run.font.color.rgb = BRAND_NAVY

    # Добавим светлый фон через параграф
    pPr = p._p.get_or_add_pPr()
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="EFF6FF"/>')
    pPr.append(shading)


def create_document():
    doc = Document()

    # === СТИЛИ ===
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)
    style.font.color.rgb = BRAND_GRAY
    style.paragraph_format.space_after = Pt(6)

    for i in range(1, 4):
        heading_style = doc.styles[f'Heading {i}']
        heading_style.font.name = 'Calibri'
        heading_style.font.color.rgb = BRAND_NAVY
        heading_style.font.bold = True
        if i == 1:
            heading_style.font.size = Pt(20)
        elif i == 2:
            heading_style.font.size = Pt(16)
        else:
            heading_style.font.size = Pt(13)

    # === ТИТУЛЬНАЯ СТРАНИЦА ===
    for _ in range(3):
        doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("B2B ПАРТНЁРСКИЙ РЕГЛАМЕНТ")
    run.font.size = Pt(12)
    run.font.color.rgb = BRAND_BLUE
    run.font.bold = True

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("REVOPS ENTERPRISE OS V17.6")
    run.font.size = Pt(14)
    run.font.color.rgb = BRAND_NAVY

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.space_before = Pt(24)
    run = p.add_run("Дорожная Карта Партнёрства:\nИИ-Супервайзер Звонков & Детектив Выручки")
    run.font.size = Pt(24)
    run.font.color.rgb = BRAND_NAVY
    run.font.bold = True

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.space_before = Pt(18)
    run = p.add_run("Практический регламент запуска речевого ИИ-контроля 100% звонков,\nпоиска скрытых резервов воронки в CRM и ежедневного мониторинга\nотдела продаж в Telegram.")
    run.font.size = Pt(11)
    run.font.italic = True
    run.font.color.rgb = BRAND_GRAY

    doc.add_paragraph()
    doc.add_paragraph()

    add_styled_table(doc,
        ["ФОРМАТ СОТРУДНИЧЕСТВА", "СТОИМОСТЬ ПИЛОТА (СТАРТ)", "РЕГУЛЯРНЫЙ НАДЗОР"],
        [["Удалённый RevOps-инженер (ИП / НПД)",
          "29 000 ₽ (7 рабочих дней)",
          "от 49 000 ₽ / месяц (подписка)"]],
        header_color="4F46E5"
    )

    doc.add_page_break()

    # === РАЗДЕЛ 1 ===
    doc.add_heading("1. Проблема рынка: Почему старый подход больше не работает", level=1)

    p = doc.add_paragraph(
        "В 90% отделов продаж малого и среднего B2B-бизнеса (от 3 до 15 менеджеров) "
        "собственник сталкивается с тремя хроническими проблемами, которые ежедневно сжигают выручку:"
    )

    problems = [
        ("1. Менеджеры работают как «справочное бюро».",
         "Консультируют входящие лиды, называют цену, но не фиксируют дату следующего контакта "
         "(Next Step) и отпускают клиента со словами «подумайте». До 30% рекламного бюджета "
         "сливается в песок прямо на этапе первого звонка."),
        ("2. Руководитель отдела продаж (РОП) физически не успевает слушать звонки.",
         "Из 500 звонков в неделю РОП слушает максимум 3–5 случайных диалогов. "
         "Контроль ведётся вслепую, а решения принимаются на основе ощущений."),
        ("3. Иллюзорный пайплайн в CRM.",
         "В воронке висят миллионы рублей на этапе «КП отправлено», но 70% из них — "
         "брошенные «зомби-сделки», по которым клиенты уже ушли к конкурентам. "
         "Собственник видит в отчёте «план на 5 млн», а по факту реально к закрытию — 1.2 млн."),
    ]

    for title, desc in problems:
        p = doc.add_paragraph()
        run = p.add_run(title + " ")
        run.font.bold = True
        run.font.color.rgb = BRAND_NAVY
        run = p.add_run(desc)
        run.font.color.rgb = BRAND_GRAY

    add_info_block(doc, "💡",
        "Главный вывод для собственника: Нанимать штатного аналитика и контролера качества — "
        "это от 120 000 до 180 000 ₽ в месяц с налогами, рабочим местом и обучением. "
        "Внедрять тяжёлые софтверные комбайны (Salesforce, HubSpot Enterprise) — "
        "это 3-6 месяцев внедрения, бюджет от 500 000 ₽ и саботаж менеджеров. "
        "Решение — подключить автономного ИИ-супервайзера, который слушает 100% звонков "
        "и связывает ошибки речи с реальными деньгами в CRM.")

    # === РАЗДЕЛ 2 ===
    doc.add_heading("2. Архитектура единого продукта: 3 компонента в 1 решении", level=1)

    p = doc.add_paragraph(
        "Продукт объединяет три важнейших контура управления продажами в единый бесшовный сервис. "
        "Менеджерам не нужно устанавливать новый софт — всё работает через существующую телефонию, "
        "CRM и привычный Telegram."
    )

    add_styled_table(doc,
        ["КОМПОНЕНТ", "ЧТО ДЕЛАЕТ В СИСТЕМЕ", "ПОЛЬЗА ДЛЯ ДИРЕКТОРА И РОПА"],
        [
            ["🎙️ ФЛАГМАН:\nРечевой ИИ-Супервайзер\n(Whisper + GPT-4)",
             "Анализирует 100% аудиозаписей звонков за вчерашний день по 13 жёстким параметрам B2B-продаж. Работает локально в защищённом контуре.",
             "РОП за 5 минут видит, кто из менеджеров не берёт дату встречи, где хамство, слив возражения «дорого» или нарушение скрипта."],
            ["🔍 ДЕТЕКТИВ ВЫРУЧКИ:\nСвязка звонков с CRM\n(AmoCRM / Битрикс24)",
             "Связывает дефекты звонков со сделками в CRM. Выявляет брошенные лиды, просроченные КП (>48ч), зависшую дебиторку и «зомби-сделки».",
             "Показывает конкретные суммы под угрозой: «Сделка D-104 на 850 000 ₽ сгорает из-за отсутствия повторного контакта. Последний касание: 6 дней назад.»"],
            ["🤠 ШЕРИФ НА АУТСОРСЕ:\nЕжедневный Telegram-надзор\nи Пульт РОПа",
             "Каждое утро в 08:30 присылает директору сводку выполнения плана, а РОПу — приоритизированный список клиентов для срочного перехвата.",
             "Собственник держит руку на пульсе компании прямо с экрана смартфона. Без захода в тяжёлые базы данных. Без обучения."]
        ]
    )

    # === РАЗДЕЛ 3 ===
    doc.add_heading("3. 13 критериев ИИ-оценки речи (Флагманский модуль)", level=1)

    p = doc.add_paragraph(
        "ИИ-модель обучена на отраслевых стандартах Enterprise B2B-продаж и проверяет каждый "
        "звонок по 13 контрольным точкам. Ниже — список критериев и метрики точности модели "
        "(валидация на 500 звонках):"
    )

    criteria = [
        ["1", "Жёсткий Next Step (Критично)", "Зафиксирована ли точная дата и время следующего контакта или встречи", "94%"],
        ["2", "Инициатива диалога", "Кто задаёт вопросы: ведёт ли менеджер диалог или только отвечает", "89%"],
        ["3", "Квалификация ЛПР", "Установлен ли статус собеседника: принимает ли он решения по бюджету", "81%"],
        ["4", "Выявление болей и сроков", "Понятно ли, какую задачу решает клиент и когда ему нужен результат", "85%"],
        ["5", "Отработка возражения «Дорого»", "Предложена ли альтернатива, рассрочка, кейс или менеджер сразу сдался", "87%"],
        ["6", "Отработка возражения «Я подумаю»", "Вскрыто ли скрытое сомнение или клиент просто отпущен", "83%"],
        ["7", "Презентация ценности через кейсы", "Звучали ли реальные примеры окупаемости и цифры внедрения", "86%"],
        ["8", "Попытка закрытия сделки", "Озвучено ли прямое предложение заключить договор или выставить счёт", "91%"],
        ["9", "Соблюдение регламента приветствия", "Корректное представление компании и имени менеджера", "96%"],
        ["10", "Чистота речи и уверенность", "Отсутствие слов-паразитов, заминок, неуверенного мямленья цены", "88%"],
        ["11", "Слушание клиента (Talk/Listen Ratio)", "Менеджер не перебивает и слушает клиента не менее 50% времени", "97%"],
        ["12", "Фиксация договорённостей в конце", "Резюмирование итогов разговора («Договорились: я отправлю КП до 14:00»...)", "92%"],
        ["13", "Отказ от скидок без повода", "Менеджер не раздаёт скидки до того, как клиент попросил об этом", "84%"],
    ]

    add_styled_table(doc,
        ["№", "КРИТЕРИЙ ОЦЕНКИ", "ЧТО ПРОВЕРЯЕТ ИИ В ДИАЛОГЕ", "ТОЧНОСТЬ (F1)"],
        criteria
    )

    add_info_block(doc, "📊",
        "Средневзвешенная точность модели: 89.2% F1. Ложноположительные срабатывания "
        "помечаются как «требует внимания РОПа» и не влияют на рейтинг менеджера до ручной верификации.")

    # === РАЗДЕЛ 4 ===
    doc.add_heading("4. Пошаговый 7-дневный таймлайн запуска (Быстрый старт)", level=1)

    p = doc.add_paragraph(
        "В отличие от классических IT-внедрений, которые длятся месяцами и требуют остановки "
        "отдела продаж, пилот запускается за 7 рабочих дней без изменения привычного рабочего процесса:"
    )

    timeline = [
        ["День 1",
         "Подписание договора (НПД/ИП). Получение API-ключей к CRM и телефонии. Подписание NDA.",
         "Старт проекта. 0 часов отвлечения менеджеров от звонков.",
         "30 мин (директор)"],
        ["День 2",
         "Подключение коннекторов (AmoCRM / Битрикс24). Выгрузка архива 100 последних звонков и базы сделок за 90 дней.",
         "Сформирован реестр сырых данных. Запущен ИИ-рентген речи.",
         "0 мин"],
        ["День 3",
         "Детектив выручки: фильтрация зомби-сделок, поиск зависших оплат, срывов SLA по КП (>48ч), просроченных задач.",
         "Выявлены скрытые резервы: найдены конкретные сделки под угрозой срыва с суммами.",
         "0 мин"],
        ["День 4",
         "Презентация отчёта директору: демонстрация карты потерь, среза 10 худших звонков, приоритезация действий.",
         "Директор видит фактическую картину продаж без украшательств. Фиксируется план исправлений.",
         "45 мин (директор + РОП)"],
        ["Дни 5–6",
         "Подключение Telegram-бота. Настройка утреннего Пульта РОПа, ролевых алертов, интеграция с телефонией.",
         "Система сдана под ключ. Запущен ежедневный надзор и контроль.",
         "15 мин (РОП)"],
        ["День 7",
         "Контрольный прогон. Верификация алертов. Передача документации. Подписание Акта.",
         "Пилот завершён. Клиент принимает решение о подписке.",
         "15 мин"],
    ]

    add_styled_table(doc,
        ["СРОК", "ДЕЙСТВИЯ REVOPS-ИНЖЕНЕРА", "РЕЗУЛЬТАТ ДЛЯ КЛИЕНТА", "ВРЕМЯ КЛИЕНТА"],
        timeline
    )

    # === РАЗДЕЛ 5 ===
    doc.add_heading("5. Финансовые условия и тарифы", level=1)

    doc.add_heading("5.1. Этап 1: Пилот «Быстрый старт» (разовый)", level=2)

    add_styled_table(doc,
        ["ПАРАМЕТР", "ЗНАЧЕНИЕ"],
        [
            ["Стоимость", "29 000 ₽ (разово, невозвратно при успешном запуске*)"],
            ["Срок", "7 рабочих дней"],
            ["Что входит", "Подключение телефонии и CRM, ИИ-аудит 100 звонков по 13 критериям, карта утечек и зависших сделок, презентационный отчёт для CEO (5 стр. + устная презентация)"],
            ["Обязанности клиента", "Предоставить API-доступы, выделить 45 мин на презентацию отчёта"],
        ]
    )

    p = doc.add_paragraph("*Возврат возможен при соблюдении условий гарантии (см. раздел 8).")
    p.runs[0].font.italic = True
    p.runs[0].font.size = Pt(9)

    doc.add_heading("5.2. Этап 2: Ежемесячный надзор «Шериф на аутсорсе»", level=2)

    tariff_table = doc.add_table(rows=12, cols=4)
    tariff_table.style = 'Table Grid'
    tariff_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    headers = ["ПАРАМЕТР", "БАЗОВЫЙ\n(3–5 менеджеров)", "ПРОДВИНУТЫЙ\n(6–10 менеджеров)", "ENTERPRISE\n(11–15 менеджеров)"]
    rows_data = [
        ["Стоимость", "49 000 ₽/мес", "79 000 ₽/мес", "120 000 ₽/мес"],
        ["ИИ слушает 100% звонков", "✅", "✅", "✅"],
        ["Утренний Telegram-пульс в 08:30", "✅", "✅", "✅"],
        ["Пульт РОПа «15 минут в день»", "✅", "✅", "✅"],
        ["Детектив зомби-сделок", "✅", "✅", "✅"],
        ["Ролевые отчёты (CEO/CFO/ROP)", "—", "✅", "✅"],
        ["Еженедельный PDF-отчёт для собственника", "—", "✅", "✅"],
        ["Мониторинг дебиторки", "—", "✅", "✅"],
        ["Индивидуальные критерии оценки (сверх 13)", "—", "до 5", "до 15"],
        ["Выделенный RevOps-инженер", "—", "—", "✅"],
        ["Максимум звонков в месяц / Доплата", "до 3 000 / 8 ₽", "до 6 000 / 6 ₽", "до 10 000 / 5 ₽"],
    ]

    # Заголовки
    for i, h in enumerate(headers):
        cell = tariff_table.rows[0].cells[i]
        cell.text = h
        set_cell_shading(cell, "1E3A5F")
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.bold = True
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                r.font.size = Pt(10)

    # Данные
    for r_idx, row in enumerate(rows_data):
        for c_idx, val in enumerate(row):
            cell = tariff_table.rows[r_idx + 1].cells[c_idx]
            cell.text = val
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(10)
                    r.font.color.rgb = BRAND_GRAY
            if r_idx % 2 == 1:
                set_cell_shading(cell, "F3F4F6")

    doc.add_paragraph()

    doc.add_heading("5.3. Юридическая чистота и налоги для компании-клиента", level=2)

    legal_points = [
        "Оплата производится безналичным расчётом с расчётного счёта юрлица/ИП клиента по договору оказания услуг.",
        "Клиент не платит за исполнителя НДФЛ (13%) и страховые взносы (30%).",
        "Бухгалтерия клиента получает официальный фискальный чек из приложения «Мой налог» (для НПД) или счёт-фактуру (для ИП на УСН) и подписанный Акт сдачи-приёмки услуг, что позволяет на 100% уменьшить налог на прибыль или УСН (ст. 15 Федерального закона № 422-ФЗ).",
        "Конфиденциальность: подписывается двусторонний NDA, персональные данные (ФИО, телефоны) обезличиваются в соответствии с 152-ФЗ РФ (см. раздел 9).",
        "Договор заключается на 3 месяца с возможностью пролонгации. Расторжение — за 14 дней с уведомлением в письменном виде."
    ]

    for i, point in enumerate(legal_points, 1):
        p = doc.add_paragraph(style='List Number')
        p.clear()
        run = p.add_run(f"{i}. {point}")
        run.font.size = Pt(10)
        run.font.color.rgb = BRAND_GRAY

    # === РАЗДЕЛ 6 ===
    doc.add_heading("6. Техническая архитектура (для ИТ-директора / службы безопасности)", level=1)

    p = doc.add_paragraph(
        "Продукт работает по принципу локальной обработки данных без передачи "
        "финансовой информации в сторонние облака:"
    )

    # ASCII-диаграмма в моноширинном шрифте
    architecture_text = """┌─────────────────────────────────────────────────────────────────────┐
│                    REVOPS ENTERPRISE OS V17.6                       │
├─────────────────────────────────────────────────────────────────────┤
│ 1. ИСТОЧНИКИ ДАННЫХ                                                 │
│    • Телефония клиента (API Mango/Voximplant/Zadarma/ATC)         │
│    • CRM: AmoCRM API v4 / Bitrix24 REST API                       │
├─────────────────────────────────────────────────────────────────────┤
│ 2. РЕЧЕВОЙ АНАЛИЗ (локальный защищённый контур)                     │
│    • Whisper Large V3 (локальный сервер, GPU)                      │
│    • Транскрипция: точность 95%+ для русского языка                │
│    • Обезличивание ФИО/телефонов на этапе транскрипции             │
├─────────────────────────────────────────────────────────────────────┤
│ 3. LLM-АНАЛИЗ (13 критериев)                                       │
│    • Промпт-инжиниринг с Few-Shot примерами                       │
│    • Каждый звонок = структурированный JSON-отчёт                  │
│    • Уверенность модели фиксируется в метаданных                   │
├─────────────────────────────────────────────────────────────────────┤
│ 4. КОРРЕЛЯЦИЯ С CRM                                                │
│    • Матчинг звонков по номеру телефона → сделка в CRM            │
│    • Детекция: просроченные задачи, зомби-этапы, дебиторка         │
├─────────────────────────────────────────────────────────────────────┤
│ 5. ДОСТАВКА РЕЗУЛЬТАТОВ                                            │
│    • Telegram Bot (утренний пульс, алерты, команды)               │
│    • PDF-отчёты (Chrome Headless, локальный рендер)               │
│    • Аудит-лог всех операций (152-ФЗ, 10 000 записей, ротация)    │
└─────────────────────────────────────────────────────────────────────┘"""

    p = doc.add_paragraph()
    run = p.add_run(architecture_text)
    run.font.name = 'Consolas'
    run.font.size = Pt(8)
    run.font.color.rgb = BRAND_NAVY

    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run("Ключевые гарантии безопасности:")
    run.font.bold = True
    run.font.color.rgb = BRAND_NAVY

    security_points = [
        "Аудиофайлы обрабатываются локально и не загружаются в облачные API",
        "Транскрипты хранятся не более 30 дней (настраивается)",
        "Финансовые данные CRM не покидают защищённый контур клиента",
        "Все операции журналируются с привязкой к роли пользователя",
        "Соответствие требованиям 152-ФЗ РФ о персональных данных"
    ]
    for point in security_points:
        p = doc.add_paragraph(style='List Bullet')
        p.clear()
        run = p.add_run(point)
        run.font.size = Pt(10)

    # === РАЗДЕЛ 7 ===
    doc.add_heading("7. Поддержка и SLA", level=1)

    add_styled_table(doc,
        ["ПАРАМЕТР", "ЗНАЧЕНИЕ"],
        [
            ["Время реакции на критический инцидент", "≤ 4 часа в рабочее время (ПН-ПТ, 9:00–19:00 МСК)"],
            ["Время реакции на некритический вопрос", "≤ 24 часа"],
            ["Каналы поддержки", "Telegram, email, экстренный звонок"],
            ["Аптайм системы", "≥ 99% в месяц"],
            ["Плановое обслуживание", "Воскресенье, 02:00–04:00 МСК"],
            ["Резервное копирование", "Ежедневное, хранение 30 дней"],
            ["Отчёт по инцидентам", "Еженедельно в рамках Пульта РОПа"],
        ]
    )

    p = doc.add_paragraph()
    run = p.add_run("Процесс эскалации:")
    run.font.bold = True
    run.font.color.rgb = BRAND_NAVY

    escalation = [
        "Алерт в Telegram (автоматический)",
        "Подтверждение принятия в работу (≤ 30 мин)",
        "Диагностика и решение (≤ 4 часа для критических)",
        "Пост-мортем в течение 48 часов"
    ]
    for i, step in enumerate(escalation, 1):
        p = doc.add_paragraph(style='List Number')
        p.clear()
        run = p.add_run(f"{i}. {step}")
        run.font.size = Pt(10)

    # === РАЗДЕЛ 8 ===
    doc.add_heading("8. Гарантия первого пилота (Risk Reversal)", level=1)

    # Блок с красной рамкой
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.right_indent = Cm(0.5)
    pPr = p._p.get_or_add_pPr()
    borders = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="12" w:space="4" w:color="DC2626"/>'
        f'<w:left w:val="single" w:sz="12" w:space="4" w:color="DC2626"/>'
        f'<w:bottom w:val="single" w:sz="12" w:space="4" w:color="DC2626"/>'
        f'<w:right w:val="single" w:sz="12" w:space="4" w:color="DC2626"/>'
        f'</w:pBdr>'
    )
    pPr.append(borders)

    run = p.add_run("🛡️ ГАРАНТИЯ РЕЗУЛЬТАТА НА ПИЛОТЕ\n\n")
    run.font.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = BRAND_WARNING

    run = p.add_run(
        "Если в течение 7 рабочих дней пилота система не выявит минимум 3 конкретных "
        "точки слива лидов (по данным ИИ-анализа звонков) ИЛИ не найдёт минимум 2 "
        "зависшие/просроченные сделки в CRM с документальным подтверждением — "
        "исполнитель возвращает 100% стоимости пилота (29 000 ₽) в течение 3 рабочих "
        "дней после подписания претензии."
    )
    run.font.size = Pt(11)

    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run("Условия применения гарантии:")
    run.font.bold = True
    run.font.color.rgb = BRAND_NAVY

    conditions = [
        "Клиент предоставил доступы к телефонии и CRM не позднее Дня 1",
        "Объём звонков за последние 90 дней составляет не менее 100 штук",
        "В компании работает не менее 3 менеджеров, совершающих исходящие звонки",
        "Клиент не ограничивал доступ к данным (частичная выгрузка)"
    ]
    for cond in conditions:
        p = doc.add_paragraph(style='List Bullet')
        p.clear()
        run = p.add_run(cond)
        run.font.size = Pt(10)

    p = doc.add_paragraph()
    run = p.add_run("Что НЕ является основанием для возврата:")
    run.font.bold = True
    run.font.color.rgb = BRAND_NAVY

    not_conditions = [
        "Клиент не согласен с интерпретацией отдельных алертов",
        "Отдел продаж не готов менять процессы по итогам отчёта",
        "В компании менее 50 звонков за 90 дней"
    ]
    for cond in not_conditions:
        p = doc.add_paragraph(style='List Bullet')
        p.clear()
        run = p.add_run(cond)
        run.font.size = Pt(10)

    # === РАЗДЕЛ 9 ===
    doc.add_heading("9. Защита персональных данных (152-ФЗ)", level=1)

    add_styled_table(doc,
        ["МЕРЫ ЗАЩИТЫ", "РЕАЛИЗАЦИЯ"],
        [
            ["Обезличивание аудио", "ФИО и телефоны заменяются на токены [КЛИЕНТ_001] на этапе транскрипции"],
            ["Хранение записей", "Аудио — не более 30 дней (настраивается по желанию клиента)"],
            ["Доступ к данным", "Только исполнитель по NDA. Ролевая матрица (Админ/РОП/Менеджер)"],
            ["Журналирование", "Каждая операция доступа фиксируется в audit_log.csv с привязкой к пользователю"],
            ["Передача третьим лицам", "Запрещена. Данные не передаются в сторонние облачные сервисы"],
            ["Удаление при расторжении", "Полное удаление всех данных в течение 7 рабочих дней с подтверждением"],
        ]
    )

    # === РАЗДЕЛ 10 ===
    doc.add_heading("10. Кейсы внедрения", level=1)

    doc.add_heading("Кейс 1: B2B-поставщик промышленного оборудования", level=2)
    case1_points = [
        ("Профиль:", "7 менеджеров, 450 звонков/неделю, средний чек 850 000 ₽"),
        ("Проблема:", "План не выполнялся 3 месяца подряд. РОП считал, что проблема в маркетинге."),
    ]
    for label, text in case1_points:
        p = doc.add_paragraph()
        run = p.add_run(label + " ")
        run.font.bold = True
        run.font.color.rgb = BRAND_NAVY
        run = p.add_run(text)

    p = doc.add_paragraph()
    run = p.add_run("Что нашёл ИИ за 7 дней пилота:")
    run.font.bold = True
    run.font.color.rgb = BRAND_NAVY

    findings1 = [
        "23 «зомби-сделки» на этапе «КП отправлено» на общую сумму 4.2 млн ₽ (последнее касание > 14 дней)",
        "68% звонков без фиксации Next Step",
        "3 менеджера системно не отрабатывали возражение «дорого» (конверсия 2% вместо 12%)"
    ]
    for f in findings1:
        p = doc.add_paragraph(style='List Bullet')
        p.clear()
        run = p.add_run(f)
        run.font.size = Pt(10)

    p = doc.add_paragraph()
    run = p.add_run("Результат за первый месяц подписки: ")
    run.font.bold = True
    run.font.color.rgb = BRAND_SUCCESS
    run = p.add_run("реактивировано 11 сделок, +1.8 млн ₽ выручки. Конверсия «встреча → договор» выросла с 8% до 19%.")

    doc.add_heading("Кейс 2: IT-интегратор (внедрение 1С)", level=2)
    case2_points = [
        ("Профиль:", "5 менеджеров, 200 звонков/неделю, средний чек 1.2 млн ₽"),
        ("Проблема:", "Собственник подозревал, что менеджеры «сливают» крупных клиентов, но не мог доказать."),
    ]
    for label, text in case2_points:
        p = doc.add_paragraph()
        run = p.add_run(label + " ")
        run.font.bold = True
        run.font.color.rgb = BRAND_NAVY
        run = p.add_run(text)

    p = doc.add_paragraph()
    run = p.add_run("Что нашёл ИИ за 7 дней пилота:")
    run.font.bold = True
    run.font.color.rgb = BRAND_NAVY

    findings2 = [
        "4 сделки на 6.1 млн ₽ без повторного контакта после отправки КП (просрочка > 72 часа)",
        "Менеджер А. системно не квалифицировал ЛПР (42% звонков)",
        "Talk/Listen Ratio у топ-менеджера: 78/22 (говорит сам, не слушает)"
    ]
    for f in findings2:
        p = doc.add_paragraph(style='List Bullet')
        p.clear()
        run = p.add_run(f)
        run.font.size = Pt(10)

    p = doc.add_paragraph()
    run = p.add_run("Результат за первый месяц: ")
    run.font.bold = True
    run.font.color.rgb = BRAND_SUCCESS
    run = p.add_run("назначены повторные встречи по всем 4 просроченным сделкам. Закрыта 1 на 1.9 млн ₽. Менеджер А. переведён на обучение, его конверсия выросла с 5% до 14% за 3 недели.")

    # === РАЗДЕЛ 11 ===
    doc.add_heading("11. Часто задаваемые вопросы (FAQ)", level=1)

    faq = [
        ("Нужно ли менять телефонию или CRM?",
         "Нет. Мы подключаемся к существующей инфраструктуре через API. Поддерживаем: Mango, Voximplant, Zadarma, Asterisk, AmoCRM, Битрикс24."),
        ("Менеджеры будут знать, что их слушают?",
         "Да, это обязательно по 152-ФЗ. Мы рекомендуем объявить о запуске системы как об инструменте развития, а не наказания. Практика показывает, что сопротивление проходит за 5–7 дней."),
        ("Что если у нас нет записей звонков?",
         "Мы подключим запись на уровне телефонии (требуется 1 день). Все операторы поддерживают эту функцию из коробки."),
        ("Сколько времени нужно от моих сотрудников?",
         "На этапе пилота: 45 минут от директора (презентация отчёта) + 15 минут от РОПа (настройка Telegram). Менеджеры не отвлекаются вообще."),
        ("Вы работаете с нашими конкурентами?",
         "Нет. В договоре фиксируется эксклюзив на отрасль/регион на время сотрудничества."),
        ("Можно ли расторгнуть подписку?",
         "Да, с уведомлением за 14 дней. Данные удаляются в течение 7 рабочих дней. Оплата за текущий месяц не возвращается."),
        ("Что если ИИ ошибётся и «накажет» менеджера за то, чего не было?",
         "Все алерты с уверенностью модели ниже 85% помечаются как «требует верификации РОПа». Рейтинг менеджера не меняется без ручной проверки."),
    ]

    for q, a in faq:
        p = doc.add_paragraph()
        run = p.add_run(f"Q: {q}")
        run.font.bold = True
        run.font.color.rgb = BRAND_BLUE
        run.font.size = Pt(11)

        p = doc.add_paragraph()
        run = p.add_run(f"A: {a}")
        run.font.size = Pt(10)
        run.font.color.rgb = BRAND_GRAY

    # === РАЗДЕЛ 12 ===
    doc.add_heading("12. Порядок подписания и старт работы", level=1)

    steps = [
        "Заявка: Клиент направляет запрос (форма на сайте / Telegram / email)",
        "Созвон-диагностика (30 мин): Обсуждение текущей ситуации, доступов, ожиданий",
        "Договор + NDA: Подписание в течение 1 рабочего дня (электронный документооборот)",
        "Оплата пилота: 29 000 ₽ по счёту (оплата до Дня 1)",
        "Старт работ: День 1 по таймлайну (раздел 4)",
        "Презентация результатов: День 4",
        "Решение о подписке: День 7",
        "Первый месяц надзора: С Дня 8"
    ]
    for i, step in enumerate(steps, 1):
        p = doc.add_paragraph(style='List Number')
        p.clear()
        run = p.add_run(step)
        run.font.size = Pt(10)

    # === ПОДПИСИ ===
    doc.add_page_break()
    doc.add_heading("СОГЛАСОВАНО И ПРИНЯТО К ИСПОЛНЕНИЮ:", level=1)

    sign_table = doc.add_table(rows=5, cols=2)
    sign_table.style = 'Table Grid'
    sign_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    sign_table.rows[0].cells[0].text = "ОТ ЗАКАЗЧИКА:"
    sign_table.rows[0].cells[1].text = "ОТ ИСПОЛНИТЕЛЯ (REVOPS):"
    for cell in sign_table.rows[0].cells:
        set_cell_shading(cell, "1E3A5F")
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.bold = True
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    sign_table.rows[1].cells[0].text = "____________________ / ______________ /"
    sign_table.rows[1].cells[1].text = "____________________ / ______________ /"

    sign_table.rows[2].cells[0].text = "Дата: «____» ________________ 2026 г."
    sign_table.rows[2].cells[1].text = "Дата: «____» ________________ 2026 г."

    sign_table.rows[3].cells[0].text = "М.П. (при наличии)"
    sign_table.rows[3].cells[1].text = "ИНН / № самозанятого: _____________"

    sign_table.rows[4].cells[0].text = ""
    sign_table.rows[4].cells[1].text = ""

    doc.add_paragraph()
    doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Документ подготовлен в рамках проекта RevOps Enterprise OS V17.6\n")
    run.font.italic = True
    run.font.size = Pt(9)
    run.font.color.rgb = BRAND_GRAY

    run = p.add_run("Версия документа: 2.0 | Дата обновления: 2026-09-30")
    run.font.italic = True
    run.font.size = Pt(9)
    run.font.color.rgb = BRAND_GRAY

    # === СОХРАНЕНИЕ ===
    output_path = Path("RevOps_AI_Supervisor_Partnership_Roadmap.docx")
    doc.save(output_path)
    print(f"✅ Документ успешно создан: {output_path.absolute()}")
    print(f"📄 Размер: {output_path.stat().st_size / 1024:.1f} KB")

    # Копирование в docs, на Рабочий стол и в Brain artifacts
    dest_desktop = Path("C:/Users/strel/Desktop/RevOps_AI_Supervisor_Partnership_Roadmap.docx")
    dest_docs = Path("docs/RevOps_AI_Supervisor_Partnership_Roadmap.docx")
    dest_brain = Path("C:/Users/strel/.gemini/antigravity/brain/b958a22b-49ff-459c-a191-6c697ec14334/RevOps_AI_Supervisor_Partnership_Roadmap.docx")

    try:
        shutil.copy2(output_path, dest_desktop)
        print(f"📋 Скопировано на Рабочий стол: {dest_desktop}")
    except Exception as e:
        print(f"⚠️ Ошибка копирования на Рабочий стол: {e}")

    try:
        dest_docs.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(output_path, dest_docs)
        print(f"📁 Скопировано в docs: {dest_docs}")
    except Exception as e:
        print(f"⚠️ Ошибка копирования в docs: {e}")

    try:
        dest_brain.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(output_path, dest_brain)
        print(f"🧠 Скопировано в Artifacts: {dest_brain}")
    except Exception as e:
        print(f"⚠️ Ошибка копирования в Artifacts: {e}")

    return output_path


if __name__ == "__main__":
    create_document()
