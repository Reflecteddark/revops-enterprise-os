"""
Генератор Единой Дорожной Карты Партнёрства RevOps Enterprise OS V17.6
Объединяет лучшее из обеих версий:
- Премиальную визуальную верстку (акцентные callout-блоки, отступы, фирменные шрифты)
- Расширенную тарифную сетку (Пилот 29к, тарифы 49к / 79к / 120к + овередж)
- 13 критериев ИИ-оценки с валидацией точности (F1 score)
- Наглядную архитектуру контура безопасности и защиту данных (152-ФЗ)
- Реальные B2B-кейсы с цифрами ROI и выручки (+1.8 млн и +1.9 млн ₽)
- Детальный регламент ежедневного надзора в Telegram (08:30 / 14:00 / 18:30)
- Блок FAQ (ответы на 7 главных возражений клиентов)
- Жесткую безоговорочную гарантию возврата средств (Risk Reversal)
- Официальный блок согласования и подписей (юридический статус приложения к договору)
"""

import os
import sys
import shutil
from pathlib import Path
import docx
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

# === ЦВЕТОВАЯ ПАЛИТРА ===
COLOR_NAVY = RGBColor(0x0F, 0x17, 0x2A)       # #0F172A Глубокий темно-синий
COLOR_PRIMARY = RGBColor(0x1E, 0x3A, 0x5F)    # #1E3A5F Фирменный синий (заголовки)
COLOR_INDIGO = RGBColor(0x4F, 0x46, 0xE5)     # #4F46E5 Индиго (акценты)
COLOR_DARK_TEXT = RGBColor(0x1E, 0x29, 0x3B)  # #1E293B Темно-серый текст
COLOR_MUTED = RGBColor(0x64, 0x74, 0x8B)      # #64748B Серый для подзаголовков
COLOR_SUCCESS = RGBColor(0x05, 0x96, 0x69)    # #059669 Изумрудный
COLOR_WARNING = RGBColor(0xDC, 0x26, 0x26)    # #DC2626 Красный акцент

HEX_PRIMARY = "1E3A5F"
HEX_INDIGO = "4F46E5"
HEX_INDIGO_LIGHT = "EEF2FF"
HEX_SURFACE = "F8FAFC"
HEX_ZEBRA = "F1F5F9"
HEX_BORDER = "CBD5E1"
HEX_RED = "DC2626"
HEX_RED_LIGHT = "FEF2F2"
HEX_GREEN = "059669"
HEX_GREEN_LIGHT = "ECFDF5"

def set_cell_background(cell, hex_color: str):
    """Устанавливает цвет заливки ячейки таблицы."""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=120, bottom=120, left=160, right=160):
    """Устанавливает внутренние отступы (padding) в ячейке."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color=HEX_BORDER, sz="4"):
    """Устанавливает аккуратные границы таблицы."""
    tblPr = table._tbl.tblPr
    borders_xml = f"""
    <w:tblBorders {nsdecls("w")}>
        <w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
    </w:tblBorders>
    """
    tblPr.append(parse_xml(borders_xml))

def add_styled_table(doc, headers, rows, col_widths=None, header_bg=HEX_PRIMARY):
    """Создаёт стилизованную таблицу с чередованием строк и внутренними отступами."""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color="CBD5E1", sz="4")

    # Заголовок таблицы
    header_row = table.rows[0]
    for i, header_text in enumerate(headers):
        cell = header_row.cells[i]
        cell.text = header_text
        set_cell_background(cell, header_bg)
        set_cell_margins(cell, top=140, bottom=140, left=180, right=180)
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            for r in p.runs:
                r.font.name = 'Calibri'
                r.font.bold = True
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                r.font.size = Pt(10)

    # Строки данных
    for r_idx, row_data in enumerate(rows):
        row = table.rows[r_idx + 1]
        is_even = (r_idx % 2 == 1)
        row_bg = HEX_ZEBRA if is_even else "FFFFFF"
        for c_idx, cell_text in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.text = str(cell_text)
            if row_bg != "FFFFFF":
                set_cell_background(cell, row_bg)
            set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)
                for r in p.runs:
                    r.font.name = 'Calibri'
                    r.font.size = Pt(9.5)
                    r.font.color.rgb = COLOR_DARK_TEXT

    # Применение ширин колонок (если переданы)
    if col_widths:
        for row in table.rows:
            for i, w in enumerate(col_widths):
                if i < len(row.cells):
                    row.cells[i].width = Inches(w)

    doc.add_paragraph() # Отступ после таблицы
    return table

def add_callout(doc, emoji: str, title: str, text: str, bg_hex=HEX_INDIGO_LIGHT, border_hex=HEX_INDIGO):
    """Информационный блок с цветной акцентной полосой слева."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.rows[0].cells[0]
    cell.width = Inches(6.8)
    set_cell_background(cell, bg_hex)
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)

    borders_xml = f"""
    <w:tcBorders {nsdecls("w")}>
        <w:left w:val="single" w:sz="24" w:space="0" w:color="{border_hex}"/>
        <w:top w:val="none"/>
        <w:right w:val="none"/>
        <w:bottom w:val="none"/>
    </w:tcBorders>
    """
    cell._tc.get_or_add_tcPr().append(parse_xml(borders_xml))

    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    run_t = p.add_run(f"{emoji} {title}")
    run_t.font.name = "Calibri"
    run_t.font.size = Pt(10.5)
    run_t.font.bold = True
    run_t.font.color.rgb = COLOR_PRIMARY

    p2 = cell.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(0)
    run_b = p2.add_run(text)
    run_b.font.name = "Calibri"
    run_b.font.size = Pt(9.5)
    run_b.font.color.rgb = COLOR_DARK_TEXT

    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_before = Pt(0)
    p_spacer.paragraph_format.space_after = Pt(4)

def add_guarantee_box(doc, title: str, main_text: str):
    """Акцентный фрейм с двойной защитной рамкой для гарантии возврата."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.rows[0].cells[0]
    cell.width = Inches(6.8)
    set_cell_background(cell, HEX_RED_LIGHT)
    set_cell_margins(cell, top=160, bottom=160, left=200, right=200)

    borders_xml = f"""
    <w:tcBorders {nsdecls("w")}>
        <w:left w:val="single" w:sz="18" w:space="0" w:color="{HEX_RED}"/>
        <w:top w:val="single" w:sz="12" w:space="0" w:color="{HEX_RED}"/>
        <w:right w:val="single" w:sz="12" w:space="0" w:color="{HEX_RED}"/>
        <w:bottom w:val="single" w:sz="12" w:space="0" w:color="{HEX_RED}"/>
    </w:tcBorders>
    """
    cell._tc.get_or_add_tcPr().append(parse_xml(borders_xml))

    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    r_t = p.add_run(title)
    r_t.font.name = "Calibri"
    r_t.font.size = Pt(11)
    r_t.font.bold = True
    r_t.font.color.rgb = COLOR_WARNING

    p2 = cell.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(0)
    r_b = p2.add_run(main_text)
    r_b.font.name = "Calibri"
    r_b.font.size = Pt(10)
    r_b.font.bold = True
    r_b.font.color.rgb = COLOR_NAVY

    doc.add_paragraph()


def generate_master_document():
    doc = docx.Document()

    # Поля документа A4 (0.8" ~ 2 см)
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.85)
        section.right_margin = Inches(0.85)

    # Базовые стили шрифтов
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Calibri'
    style_normal.font.size = Pt(10.5)
    style_normal.font.color.rgb = COLOR_DARK_TEXT
    style_normal.paragraph_format.space_after = Pt(6)
    style_normal.paragraph_format.line_spacing = 1.15

    for i in range(1, 4):
        h_style = doc.styles[f'Heading {i}']
        h_style.font.name = 'Calibri'
        h_style.font.color.rgb = COLOR_PRIMARY
        h_style.font.bold = True
        h_style.paragraph_format.keep_with_next = True
        if i == 1:
            h_style.font.size = Pt(17)
            h_style.paragraph_format.space_before = Pt(14)
            h_style.paragraph_format.space_after = Pt(6)
        elif i == 2:
            h_style.font.size = Pt(13.5)
            h_style.paragraph_format.space_before = Pt(10)
            h_style.paragraph_format.space_after = Pt(4)
        else:
            h_style.font.size = Pt(11.5)
            h_style.paragraph_format.space_before = Pt(8)
            h_style.paragraph_format.space_after = Pt(3)

    # ==================== ТИТУЛЬНЫЙ ЛИСТ ====================
    for _ in range(2):
        doc.add_paragraph()

    p_badge = doc.add_paragraph()
    p_badge.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_badge = p_badge.add_run("B2B ПАРТНЁРСКИЙ РЕГЛАМЕНТ • ОФИЦИАЛЬНОЕ ПРЕДЛОЖЕНИЕ")
    r_badge.font.size = Pt(9.5)
    r_badge.font.bold = True
    r_badge.font.color.rgb = COLOR_INDIGO

    p_sys = doc.add_paragraph()
    p_sys.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sys = p_sys.add_run("REVOPS ENTERPRISE OS V17.6 (PRODUCTION SUITE)")
    r_sys.font.size = Pt(12)
    r_sys.font.bold = True
    r_sys.font.color.rgb = COLOR_PRIMARY

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(16)
    p_title.paragraph_format.space_after = Pt(8)
    r_title = p_title.add_run("Дорожная Карта Партнёрства:\nИИ-Супервайзер Звонков & Детектив Выручки")
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = COLOR_NAVY

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_before = Pt(4)
    p_sub.paragraph_format.space_after = Pt(20)
    r_sub = p_sub.add_run(
        "Практический регламент внедрения речевого ИИ-контроля 100% звонков,\n"
        "ликвидации утечек в CRM (AmoCRM / Битрикс24) и ежедневного надзора отдела продаж в Telegram."
    )
    r_sub.font.italic = True
    r_sub.font.size = Pt(10.5)
    r_sub.font.color.rgb = COLOR_MUTED

    doc.add_paragraph()

    add_styled_table(doc,
        ["ФОРМАТ СОТРУДНИЧЕСТВА", "СТОИМОСТЬ ПИЛОТА (СТАРТ)", "РЕГУЛЯРНЫЙ НАДЗОР"],
        [["Удалённый RevOps-инженер (НПД / ИП)",
          "29 000 ₽ (7 рабочих дней под ключ)",
          "от 49 000 ₽ / месяц (подписка)"]],
        col_widths=[2.5, 2.3, 2.0],
        header_bg=HEX_INDIGO
    )

    p_note = doc.add_paragraph()
    p_note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_n = p_note.add_run("Конфиденциально • В соответствии с Федеральным законом № 152-ФЗ РФ и № 422-ФЗ РФ")
    r_n.font.size = Pt(8.5)
    r_n.font.color.rgb = COLOR_MUTED

    doc.add_page_break()

    # ==================== РАЗДЕЛ 1 ====================
    doc.add_heading("1. Проблема рынка: Почему старый подход больше не работает", level=1)

    p = doc.add_paragraph(
        "В 90% отделов продаж малого и среднего B2B-бизнеса (от 3 до 15 менеджеров) "
        "собственник инвестирует сотни тысяч рублей в маркетинг и лидогенерацию, но сталкивается "
        "с тремя хроническими проблемами, сжигающими конверсию в кассу:"
    )

    problems = [
        ("1. Менеджеры работают как «справочное бюро».",
         "Консультируют лиды, подробно отвечают на вопросы, называют цену, но не берут инициативу, "
         "не фиксируют точную дату и время следующего шага (Next Step) и отпускают клиента фразой «ну вы подумайте». "
         "До 35% рекламного бюджета сливается прямо на первом звонке."),
        ("2. Руководитель отдела продаж (РОП) физически не успевает слушать звонки.",
         "В отделе из 5 менеджеров совершается 300–500 звонков в неделю. РОП слушает максимум 3–5 случайных диалогов. "
         "Контроль ведётся вслепую, а разборы на планёрках строятся на субъективных ощущениях и оправданиях менеджеров."),
        ("3. Иллюзорный пайплайн в CRM («Кладбище сделок»).",
         "В воронке висят миллионы виртуальных рублей на этапах «КП отправлено» или «Договор на согласовании». "
         "При детальном аудите 60–75% из них оказываются брошенными «зомби-сделками», по которым контакт не поддерживался более 14 дней, "
         "а клиент уже купил у конкурентов. Собственник рассчитывает на план, но в кассе образуется кассовый разрыв."),
    ]

    for title, desc in problems:
        p = doc.add_paragraph()
        run = p.add_run(title + " ")
        run.font.bold = True
        run.font.color.rgb = COLOR_PRIMARY
        run = p.add_run(desc)

    add_callout(doc, "💡", "Экономический вывод для собственника",
        "Нанимать штатного аналитика качества и РОПа-контролёра — это от 130 000 до 180 000 ₽ в месяц с налогами, "
        "рабочим местом и постоянным риском увольнения. Внедрять тяжелые софтверные комбайны — это от 400 000 ₽, 3-6 месяцев "
        "ожидания и неизбежный саботаж сотрудников. Наше решение — автономный ИИ-супервайзер, который слушает 100% звонков "
        "и связывает дефекты речи менеджеров с конкретными деньгами под угрозой в вашей CRM.")

    # ==================== РАЗДЕЛ 2 ====================
    doc.add_heading("2. Архитектура единого продукта: 3 компонента в 1 решении", level=1)

    p = doc.add_paragraph(
        "Продукт объединяет три важнейших контура управления продажами в единый бесшовный сервис. "
        "Менеджерам не нужно осваивать новые программы: система встраивается в текущую телефонию, "
        "CRM-систему и корпоративный Telegram."
    )

    add_styled_table(doc,
        ["КОМПОНЕНТ", "ФУНКЦИОНАЛ В СИСТЕМЕ", "ЦЕННОСТЬ ДЛЯ ДИРЕКТОРА И РОПА"],
        [
            ["🎙️ ФЛАГМАН:\nРечевой ИИ-Супервайзер\n(Whisper + LLM)",
             "Анализирует 100% аудиозаписей всех звонков за сутки по 13 критериям Enterprise B2B-продаж в изолированном контуре.",
             "РОП за 5 минут видит: кто не берет дату встречи, где хамство, слив возражения «дорого», где менеджер говорит больше 70% времени."],
            ["🔍 ДЕТЕКТИВ ВЫРУЧКИ:\nСвязка звонков с CRM\n(AmoCRM / Битрикс24)",
             "Связывает дефекты диалогов со сделками в CRM. Выявляет брошенные лиды, просроченные КП (>48ч), зависшие счета и зомби-сделки.",
             "Показывает конкретные суммы под угрозой: «Сделка D-104 на 850 000 ₽ сгорает — клиент не получил КП за 72 часа. Менеджер забыл перезвонить.»"],
            ["🤠 ШЕРИФ В TELEGRAM:\nЕжедневный надзор\nи Пульт РОПа «15 минут»",
             "Присылает персонализированные сводки: утром в 08:30 — утренний фокус дня, в течение дня — SOS-алерты по горящим сделкам, вечером — итоги.",
             "Собственник держит руку на пульсе компании прямо со смартфона за 2 минуты в день. Без копания в отчетах. Без рутины."]
        ],
        col_widths=[2.1, 2.7, 2.0]
    )

    add_callout(doc, "⏱️", "Ежедневный распорядок работы «Шерифа» в Telegram",
        "• 08:30 — Утренний фокус: топ-5 горящих сделок дня, список клиентов без Next Step, напоминание по просроченным КП.\n"
        "• 14:00 — Экспресс SOS-алерт: критические нарушения за первую половину дня (грубость, слив скидок без повода).\n"
        "• 18:30 — Вечерний срез: факт звонков, средний балл речевой дисциплины по менеджерам, динамика воронки.")

    # ==================== РАЗДЕЛ 3 ====================
    doc.add_heading("3. 13 критериев ИИ-оценки речи (Флагманский модуль)", level=1)

    p = doc.add_paragraph(
        "ИИ-модель валидирована на стандартах B2B-продаж с высоким чеком. "
        "Каждый аудиозвонок расшифровывается моделью Whisper и оценивается по 13 критериям. "
        "Ниже приведены контрольные точки и подтвержденная точность (F1-Score на выборке 500+ диалогов):"
    )

    criteria = [
        ["1", "Жёсткий Next Step (Критично)", "Зафиксирована ли точная дата и время следующего контакта или встречи", "94%"],
        ["2", "Инициатива диалога", "Кто задаёт вопросы: ведёт ли менеджер клиента по этапам или просто отвечает", "89%"],
        ["3", "Квалификация ЛПР / ЛВПР", "Установлен ли статус собеседника: принимает ли он решения по бюджету и срокам", "81%"],
        ["4", "Выявление болей и срочности", "Понятно ли, какую бизнес-задачу решает клиент и к какому дедлайну нужен результат", "85%"],
        ["5", "Отработка возражения «Дорого»", "Предложена ли альтернатива, рассрочка, экономическое обоснование или сразу скидка", "87%"],
        ["6", "Отработка возражения «Я подумаю»", "Вскрыто ли истинное сомнение или клиент просто отпущен без договоренностей", "83%"],
        ["7", "Презентация ценности через кейсы", "Звучали ли реальные примеры окупаемости, аналогичные проекты и факты", "86%"],
        ["8", "Попытка закрытия сделки", "Озвучено ли прямое предложение заключить договор, забронировать объем или выставить счет", "91%"],
        ["9", "Соблюдение регламента приветствия", "Корректное представление компании, имени менеджера и цели контакта", "96%"],
        ["10", "Чистота речи и уверенность", "Отсутствие слов-паразитов, затяжных пауз и неуверенной интонации при озвучивании цены", "88%"],
        ["11", "Слушание клиента (Talk/Listen Ratio)", "Менеджер не перебивает и слушает клиента не менее 50% времени разговора", "97%"],
        ["12", "Фиксация договорённостей в финале", "Резюмирование итогов диалога («Договорились: я отправлю расчет до 15:00...»)", "92%"],
        ["13", "Защита маржинальности (Отказ от скидок)", "Менеджер не раздаёт скидки до того, как клиент обоснованно попросил об этом", "84%"],
    ]

    add_styled_table(doc,
        ["№", "КРИТЕРИЙ ОЦЕНКИ", "ЧТО ПРОВЕРЯЕТ ИИ В ДИАЛОГЕ", "ТОЧНОСТЬ (F1)"],
        criteria,
        col_widths=[0.4, 2.3, 3.4, 0.7]
    )

    add_callout(doc, "📊", "Статистика надежности модели",
        "Средневзвешенная точность речевого аудита составляет 89.2% F1. В спорных случаях с низкой "
        "уверенностью модели (<80%) звонок маркируется тегом «Требует внимания РОПа» и не влияет на KPI менеджера без верификации.")

    # ==================== РАЗДЕЛ 4 ====================
    doc.add_heading("4. Пошаговый 7-дневный таймлайн запуска (Быстрый старт)", level=1)

    p = doc.add_paragraph(
        "В отличие от традиционных IT-внедрений, длящихся месяцами, пилот «Быстрый старт» "
        "развертывается ровно за 7 рабочих дней без остановки работы отдела продаж:"
    )

    timeline = [
        ["День 1",
         "Заключение договора (НПД/ИП) и соглашения о конфиденциальности (NDA). Получение API-ключей к CRM и телефонии.",
         "Старт проекта. 0 часов отвлечения менеджеров от звонков.",
         "30 мин (директор)"],
        ["День 2",
         "Подключение коннекторов (AmoCRM / Битрикс24). Выгрузка архива 100 последних звонков и базы сделок за последние 90 дней.",
         "Сформирован защищенный реестр данных. Запущен ИИ-рентген речи.",
         "0 мин"],
        ["День 3",
         "Детектив выручки: фильтрация зомби-сделок, поиск брошенных клиентов, нарушений SLA по отправке КП (>48ч) и зависших счетов.",
         "Выявлены скрытые резервы: найдены конкретные сделки под угрозой срыва с точными суммами.",
         "0 мин"],
        ["День 4",
         "Презентация отчёта собственнику: разбор 10 худших звонков, оцифровка карты потерь и ранжирование менеджеров.",
         "Собственник видит фактическую картину отдела без украшательств. Утверждается план быстрых побед.",
         "45 мин (директор + РОП)"],
        ["Дни 5–6",
         "Подключение Telegram-бота. Настройка утреннего Пульта РОПа, ролевых алертов и ежедневного мониторинга.",
         "Система сдана под ключ. Запущен непрерывный автоматический надзор 100% звонков.",
         "15 мин (РОП)"],
        ["День 7",
         "Контрольный прогон алертов. Передача исполнительной документации. Подписание двустороннего Акта.",
         "Пилот завершён. Оцифрован возврат инвестиций. Принятие решения о регулярном надзоре.",
         "15 мин"],
    ]

    add_styled_table(doc,
        ["СРОК", "ДЕЙСТВИЯ REVOPS-ИНЖЕНЕРА", "РЕЗУЛЬТАТ ДЛЯ КЛИЕНТА", "ВРЕМЯ КЛИЕНТА"],
        timeline,
        col_widths=[0.9, 2.5, 2.4, 1.0]
    )

    # ==================== РАЗДЕЛ 5 ====================
    doc.add_heading("5. Финансовые условия, тарифы и юридическая чистота", level=1)

    doc.add_heading("5.1. Этап 1: Пилот «Быстрый старт» (разовый запуск)", level=2)

    add_styled_table(doc,
        ["ПАРАМЕТР", "УСЛОВИЯ ПИЛОТНОГО ЭТАПА"],
        [
            ["Стоимость пилота", "29 000 ₽ (фиксированная стоимость за весь объем работ под ключ)"],
            ["Срок выполнения", "7 рабочих дней"],
            ["Что входит в объем", "Подключение телефонии и CRM, ИИ-аудит 100 звонков по 13 критериям, карта утечек воронки и зависших сделок, презентационный отчет для CEO (5 стр. + Zoom-разбор)"],
            ["Обязательства клиента", "Предоставить API-доступы к телефонии и CRM, выделить 45 минут на презентацию отчета"],
            ["Гарантия результата", "100% возврат средств при ненахождении утечек (см. Раздел 8)"]
        ],
        col_widths=[2.2, 4.6]
    )

    doc.add_heading("5.2. Этап 2: Ежемесячный надзор «Шериф на аутсорсе» (подписка)", level=2)

    p = doc.add_paragraph(
        "Стоимость ежемесячного сопровождения зависит исключительно от количества менеджеров "
        "и объема аудиопотока. Вы платите только за реальный масштаб вашего отдела:"
    )

    tariff_rows = [
        ["Стоимость в месяц", "49 000 ₽ / мес", "79 000 ₽ / мес", "120 000 ₽ / мес"],
        ["Размер отдела продаж", "3–5 менеджеров", "6–10 менеджеров", "11–15 менеджеров"],
        ["ИИ слушает 100% звонков", "✅ Включено", "✅ Включено", "✅ Включено"],
        ["Утренний пульс в Telegram (08:30)", "✅ Включено", "✅ Включено", "✅ Включено"],
        ["Пульт РОПа «15 минут в день»", "✅ Включено", "✅ Включено", "✅ Включено"],
        ["Детектив зомби-сделок в CRM", "✅ Включено", "✅ Включено", "✅ Включено"],
        ["Ролевые срезы (СЕО / РОП / Финансы)", "Раз в месяц", "Еженедельно", "Еженедельно"],
        ["Контроль дебиторской задолженности", "Базовый", "Глубокий мониторинг", "Приоритетный радар"],
        ["Индивидуальные критерии оценки", "Стандарт (13)", "до 5 дополнительных", "до 15 индивидуальных"],
        ["Выделенный RevOps-инженер", "Общая линия", "Приоритет в Telegram", "Персональный инженер"],
        ["Лимит звонков в месяц / Доплата", "до 3 000 / 8 ₽ за звонок", "до 6 000 / 6 ₽ за звонок", "до 10 000 / 5 ₽ за звонок"],
    ]

    add_styled_table(doc,
        ["ПАРАМЕТР", "БАЗОВЫЙ\n(3–5 менеджеров)", "ПРОДВИНУТЫЙ\n(6–10 менеджеров)", "ENTERPRISE\n(11–15 менеджеров)"],
        tariff_rows,
        col_widths=[2.3, 1.5, 1.5, 1.5],
        header_bg=HEX_PRIMARY
    )

    doc.add_heading("5.3. Юридическая чистота и налоговая оптимизация", level=2)

    legal_items = [
        "Оплата по безналичному расчету: Перечисление производится с расчетного счета юрлица или ИП клиента по официальному договору оказания услуг.",
        "0% страховых взносов и 0% НДФЛ: Клиент освобожден от уплаты взносов в Социальный фонд (30%) и подоходного налога (13%), которые возникают при найме штатных сотрудников.",
        "100% списание расходов: Заказчик получает официальный электронный чек из приложения «Мой налог» (для плательщиков НПД по ст. 15 Федерального закона № 422-ФЗ) либо счет и акт (для ИП на УСН). Сумма в полном объеме уменьшает налогооблагаемую базу по налогу на прибыль или УСН (Доходы минус Расходы).",
        "Строгий режим конфиденциальности: До начала работ подписывается двустороннее соглашение NDA с жесткими штрафными санкциями.",
        "Гибкий срок: Договор заключается на 3 месяца с автоматической пролонгацией. Расторжение возможно в любой момент с письменным уведомлением за 14 дней."
    ]

    for item in legal_items:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(3)
        parts = item.split(":", 1)
        r_b = p.add_run(parts[0] + ":")
        r_b.font.bold = True
        r_b.font.color.rgb = COLOR_PRIMARY
        r_t = p.add_run(parts[1])

    # ==================== РАЗДЕЛ 6 ====================
    doc.add_heading("6. Техническая архитектура и безопасность данных (152-ФЗ)", level=1)

    p = doc.add_paragraph(
        "Система спроектирована по стандартам изолированного контура. Персональные данные "
        "и финансовая тайна клиента защищены от утечек в сторонние публичные нейросети:"
    )

    arch_text = """┌─────────────────────────────────────────────────────────────────────────────┐
│                    REVOPS ENTERPRISE OS V17.6 SECURITY                      │
├─────────────────────────────────────────────────────────────────────────────┤
│ 1. ИСТОЧНИКИ ДАННЫХ (API)                                                   │
│    • Телефония: Mango-Office / UIS / CoMagic / Zadarma / Asterisk           │
│    • CRM: AmoCRM API v4 / Bitrix24 Webhook & REST API                       │
├─────────────────────────────────────────────────────────────────────────────┤
│ 2. РЕЧЕВОЙ АНАЛИЗ (ИЗОЛИРОВАННЫЙ КОНТУР)                                    │
│    • Whisper Large V3 (локальный сервер обработки аудио)                    │
│    • Точность распознавания русской речи: 95%+                              │
│    • Автоматическое обезличивание ФИО и телефонов перед отправкой в LLM     │
├─────────────────────────────────────────────────────────────────────────────┤
│ 3. LLM-АУДИТ (13 КРИТЕРИЕВ ПРОДАЖ)                                          │
│    • Промпт-инжиниринг Few-Shot с калибровкой контекста B2B                 │
│    • Каждый диалог структурируется в валидированный JSON-объект             │
│    • Фиксация индекса уверенности модели (Confidence Score)                 │
├─────────────────────────────────────────────────────────────────────────────┤
│ 4. КОРРЕЛЯЦИЯ С ВОРОНКОЙ В CRM                                              │
│    • Сопоставление аудиозаписи со сделкой по ID / номеру контакта           │
│    • Детекция зависших стадий, просроченных задач и дебиторской задолженности│
├─────────────────────────────────────────────────────────────────────────────┤
│ 5. ДОСТАВКА И ОТЧЕТНОСТЬ                                                    │
│    • Telegram Bot (утренний пульс, экстренные SOS-алерты)                   │
│    • PDF-отчеты (локальный рендер через headless-движок)                    │
│    • Защищенный журнал аудита (audit_log.csv, ротация 10 000 записей)       │
└─────────────────────────────────────────────────────────────────────────────┘"""

    p_arch = doc.add_paragraph()
    r_arch = p_arch.add_run(arch_text)
    r_arch.font.name = 'Consolas'
    r_arch.font.size = Pt(8)
    r_arch.font.color.rgb = COLOR_PRIMARY

    add_styled_table(doc,
        ["МЕРА ЗАЩИТЫ", "ТЕХНИЧЕСКАЯ РЕАЛИЗАЦИЯ (152-ФЗ РФ)"],
        [
            ["Обезличивание персональных данных", "ФИО клиентов и телефонные номера маскируются системными токенами [КЛИЕНТ_ID] на этапе транскрибации."],
            ["Срок хранения записей", "Аудиофайлы удаляются с буферного сервера через 30 дней (или настраивается индивидуально)."],
            ["Доступ к финансовым данным", "Управление доступом через ролевую матрицу RBAC (Супервайзер / РОП / Менеджер)."],
            ["Неприкосновенность тайны", "Коммерческие данные о сделках не передаются в публичные модели и не используются для дообучения сторонних ИИ."],
            ["Полное удаление при расторжении", "При прекращении сотрудничества все архивы и логи безвозвратно удаляются в течение 7 рабочих дней."],
        ],
        col_widths=[2.3, 4.5]
    )

    # ==================== РАЗДЕЛ 7 ====================
    doc.add_heading("7. Регламент взаимодействия и SLA", level=1)

    add_styled_table(doc,
        ["ПАРАМЕТР СОГЛАШЕНИЯ", "РЕГЛАМЕНТНЫЙ НОРМАТИВ"],
        [
            ["Время реакции на критический инцидент (сбой алертов)", "≤ 4 часов в рабочее время (Пн–Пт, с 09:00 до 19:00 МСК)"],
            ["Время ответа на стандартный запрос / вопрос", "≤ 24 часов"],
            ["Каналы коммуникации", "Выделенный чат в Telegram, электронная почта, телефонная связь"],
            ["Гарантированный аптайм сервиса", "≥ 99.0% времени в месяц"],
            ["Окно планового техобслуживания", "Воскресенье, с 02:00 до 05:00 МСК (без остановки отдела продаж)"],
            ["Резервное копирование базы аналитики", "Ежедневно в 03:00 с глубиной хранения архивов 30 дней"],
        ],
        col_widths=[2.8, 4.0]
    )

    # ==================== РАЗДЕЛ 8 ====================
    doc.add_heading("8. Безоговорочная гарантия пилота (Risk Reversal)", level=1)

    add_guarantee_box(doc,
        "🛡️ 100% ФИНАНСОВАЯ ГАРАНТИЯ РЕЗУЛЬТАТА НА ПИЛОТЕ:",
        "Если в течение 7 рабочих дней пилота система не выявит минимум 3 конкретных точки "
        "слива клиентов менеджерами в звонках ИЛИ не найдет минимум 2 зависшие / просроченные "
        "сделки в CRM с документальным подтверждением сумм — Исполнитель возвращает 100% стоимости "
        "пилота (29 000 ₽) в течение 3 рабочих дней без споров и удержаний."
    )

    p_cond = doc.add_paragraph()
    r_ch = p_cond.add_run("Условия действия гарантии:")
    r_ch.font.bold = True
    r_ch.font.color.rgb = COLOR_PRIMARY

    cond_items = [
        "Заказчик предоставил корректные доступы к API телефонии и CRM в День 1;",
        "Объем звонков в отделе продаж составляет не менее 100 записей за последние 90 дней;",
        "В компании работает не менее 3 действующих менеджеров, совершающих звонки клиентам."
    ]
    for c in cond_items:
        p_c = doc.add_paragraph(style='List Bullet')
        p_c.paragraph_format.space_before = Pt(1)
        p_c.paragraph_format.space_after = Pt(2)
        r = p_c.add_run(c)

    # ==================== РАЗДЕЛ 9 ====================
    doc.add_heading("9. Практические кейсы внедрения и возврат инвестиций (ROI)", level=1)

    doc.add_heading("Кейс 1: Оптовый поставщик промышленного оборудования", level=2)
    p = doc.add_paragraph(
        "• Профиль: 7 менеджеров по продажам, средний чек 850 000 ₽, около 450 звонков в неделю.\n"
        "• Ситуация на старте: План продаж выполнялся на 62% на протяжении трех месяцев. РОП списывал проблему на «некачественный трафик»."
    )
    p_f1 = doc.add_paragraph()
    r_f1 = p_f1.add_run("Что выявил ИИ-Супервайзер за 7 дней пилота:")
    r_f1.font.bold = True
    r_f1.font.color.rgb = COLOR_PRIMARY

    f1_list = [
        "Найдено 23 «зомби-сделки» на этапе «КП отправлено» на 4.2 млн ₽ (повторного контакта не было более 14 дней);",
        "В 68% звонков менеджеры не назначали дату следующего шага («если что — звоните сами»);",
        "Трое менеджеров сразу сдавались при возражении «дорого», не предлагая модификаций комплектации."
    ]
    for item in f1_list:
        p_i = doc.add_paragraph(style='List Bullet')
        p_i.paragraph_format.space_before = Pt(1)
        p_i.paragraph_format.space_after = Pt(2)
        p_i.add_run(item)

    p_res1 = doc.add_paragraph()
    r_r1 = p_res1.add_run("Результат первого месяца регулярного надзора: ")
    r_r1.font.bold = True
    r_r1.font.color.rgb = COLOR_SUCCESS
    r_res1_t = p_res1.add_run("Реактивировано 11 брошенных сделок, получено +1 850 000 ₽ дополнительной выручки. Конверсия из презентации КП в оплаченный счет выросла с 8% до 19%. Сервис окупился более чем в 37 раз.")

    doc.add_heading("Кейс 2: IT-интегратор корпоративного ПО (1С / ERP)", level=2)
    p2 = doc.add_paragraph(
        "• Профиль: 5 менеджеров по сложным сделкам, средний чек 1 200 000 ₽, 200 звонков в неделю.\n"
        "• Ситуация на старте: Собственник подозревал, что менеджеры сливают крупных клиентов из-за лени, но не имел времени слушать многочасовые записи."
    )
    p_f2 = doc.add_paragraph()
    r_f2 = p_f2.add_run("Что выявил ИИ-Супервайзер за 7 дней пилота:")
    r_f2.font.bold = True
    r_f2.font.color.rgb = COLOR_PRIMARY

    f2_list = [
        "Обнаружено 4 крупных сделки на 6.1 млн ₽, где клиентам отправили КП 5 дней назад и ни разу не перезвонили;",
        "Один из ведущих менеджеров в 42% звонков не квалифицировал ЛПР и общался с рядовыми бухгалтерами;",
        "Коэффициент говорения топ-менеджера составлял 78% (говорил сам, перебивал клиента, не выявляя бюджет)."
    ]
    for item in f2_list:
        p_i = doc.add_paragraph(style='List Bullet')
        p_i.paragraph_format.space_before = Pt(1)
        p_i.paragraph_format.space_after = Pt(2)
        p_i.add_run(item)

    p_res2 = doc.add_paragraph()
    r_r2 = p_res2.add_run("Результат первого месяца регулярного надзора: ")
    r_r2.font.bold = True
    r_r2.font.color.rgb = COLOR_SUCCESS
    r_res2_t = p_res2.add_run("РОП оперативно перехватил 4 зависшие сделки, закрыта 1 сделка на 1 900 000 ₽. После внедрения ежедневного пульта РОПа средняя конверсия менеджеров поднялась с 5% до 14%.")

    # ==================== РАЗДЕЛ 10 ====================
    doc.add_heading("10. Часто задаваемые вопросы (FAQ)", level=1)

    faq_items = [
        ("Нужно ли менять телефонию или переносить данные из CRM?",
         "Нет. Мы работаем поверх существующей инфраструктуры через официальные API. Поддерживаются Mango-Office, UIS, CoMagic, Zadarma, Мегафон, Билайн, Asterisk, AmoCRM и Битрикс24. Установка сторонних программ на компьютеры сотрудников не требуется."),
        ("Менеджеры будут знать, что их звонки проверяет ИИ?",
         "Да, уведомление сотрудников обязательно в рамках ст. 152-ФЗ. Мы рекомендуем позиционировать систему как «виртуального помощника РОПа», который снимает рутину и подсвечивает точки роста. Опыт показывает, что психологический саботаж исчезает в течение первой недели, а дисциплина взлетает сразу после объявления о запуске."),
        ("Что делать, если в компании не включена запись звонков?",
         "Мы помогаем включить запись на стороне вашей виртуальной АТС за 1 рабочий день. Все современные провайдеры связи имеют эту функцию из коробки."),
        ("Сколько времени потребуется от руководства компании?",
         "На этапе пилота: 30 минут на передачу доступов в День 1 и 45 минут на презентацию отчета в День 4. В ходе регулярной подписки: собственник тратит до 2 минут в день на просмотр сводки в Telegram, а РОП — ровно 15 минут утром по готовому списку приоритетов."),
        ("Работаете ли вы с нашими прямыми конкурентами?",
         "Нет. По желанию Заказчика в договоре фиксируется отраслевой и географический эксклюзив на период действия договора сопровождения."),
        ("Что произойдет, если ИИ ошибется в анализе сложного диалога?",
         "При оценке звонка рассчитывается уверенность модели. Все диалоги с уверенностью ниже 80% получают метку «Требует верификации РОПа». Штрафы и выводы никогда не выносятся автоматически без подтверждения человеком."),
        ("Можно ли прекратить подписку в любой момент?",
         "Да, предупредив за 14 дней в письменной форме. Накопленные базы аналитики и отчеты остаются у Заказчика. Оплаченный текущий период дорабатывается в полном объеме."),
    ]

    for q, a in faq_items:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(1)
        r_q = p_q.add_run(f"В: {q}")
        r_q.font.bold = True
        r_q.font.color.rgb = COLOR_INDIGO
        r_q.font.size = Pt(10.5)

        p_a = doc.add_paragraph()
        p_a.paragraph_format.space_before = Pt(0)
        p_a.paragraph_format.space_after = Pt(4)
        r_a = p_a.add_run(f"О: {a}")
        r_a.font.size = Pt(9.5)
        r_a.font.color.rgb = COLOR_DARK_TEXT

    # ==================== РАЗДЕЛ 11 ====================
    doc.add_heading("11. Порядок подписания и старт работ", level=1)

    steps = [
        "Заявка и экспресс-диагностика: согласование готовности телефонии и CRM (20 мин);",
        "Подписание Договора и соглашения о конфиденциальности (NDA);",
        "Оплата пилотного этапа (29 000 ₽) по безналичному расчету;",
        "Старт работ: День 1 по регламентному графику запуска (Раздел 4);",
        "Проведение ИИ-аудита и презентация карты потерь директору: День 4;",
        "Настройка Пульта РОПа в Telegram и сдача системы: День 7;",
        "Подписание двустороннего Акта сдачи-приемки и переход на регулярный надзор."
    ]
    for i, s in enumerate(steps, 1):
        p_s = doc.add_paragraph(style='List Number')
        p_s.paragraph_format.space_before = Pt(1)
        p_s.paragraph_format.space_after = Pt(2)
        p_s.clear()
        r_num = p_s.add_run(f"{i}. ")
        r_num.font.bold = True
        r_num.font.color.rgb = COLOR_PRIMARY
        p_s.add_run(s)

    # ==================== РАЗДЕЛ 12: ПОДПИСИ И СОГЛАСОВАНИЕ ====================
    doc.add_page_break()
    doc.add_heading("12. Согласование условий и реквизиты сторон", level=1)

    p_agr = doc.add_paragraph(
        "Настоящая Дорожная Карта является неотъемлемой частью Договора на оказание "
        "консультационно-технических услуг. Условия согласованы Сторонами в полном объеме:"
    )

    sign_table = doc.add_table(rows=6, cols=2)
    sign_table.style = 'Table Grid'
    sign_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(sign_table, color="CBD5E1", sz="4")

    # Заголовок
    h0 = sign_table.rows[0].cells[0]
    h1 = sign_table.rows[0].cells[1]
    h0.text = "ОТ ЗАКАЗЧИКА:"
    h1.text = "ОТ ИСПОЛНИТЕЛЯ (REVOPS):"
    for cell in [h0, h1]:
        set_cell_background(cell, HEX_PRIMARY)
        set_cell_margins(cell, top=140, bottom=140, left=160, right=160)
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.bold = True
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                r.font.size = Pt(10)

    rows_info = [
        ("Организация: __________________________", "Исполнитель: __________________________"),
        ("В лице: _______________________________", "Статус: Плательщик НПД / ИП"),
        ("Подпись: ___________ / _________________ /", "Подпись: ___________ / _________________ /"),
        ("Дата: «____» ____________________ 2026 г.", "Дата: «____» ____________________ 2026 г."),
        ("М.П. (при наличии)", "ИНН: __________________________________")
    ]

    for idx, (left_val, right_val) in enumerate(rows_info, 1):
        c_l = sign_table.rows[idx].cells[0]
        c_r = sign_table.rows[idx].cells[1]
        c_l.text = left_val
        c_r.text = right_val
        for c in [c_l, c_r]:
            set_cell_margins(c, top=120, bottom=120, left=160, right=160)
            for p in c.paragraphs:
                for r in p.runs:
                    r.font.name = 'Calibri'
                    r.font.size = Pt(9.5)
                    r.font.color.rgb = COLOR_DARK_TEXT

    doc.add_paragraph()
    doc.add_paragraph()

    p_f = doc.add_paragraph()
    p_f.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_f1 = p_f.add_run("Документ сформирован в рамках операционной платформы RevOps Enterprise OS V17.6\n")
    r_f1.font.italic = True
    r_f1.font.size = Pt(8.5)
    r_f1.font.color.rgb = COLOR_MUTED

    r_f2 = p_f.add_run("Версия: 3.0 Enterprise Master Edition | Дата: 2026-09-30 | Хеш сборки: revops-v17-master")
    r_f2.font.italic = True
    r_f2.font.size = Pt(8.5)
    r_f2.font.color.rgb = COLOR_MUTED

    # Сохранение файлов
    output_filename = "Дорожная_Карта_RevOps_AI_Супервайзер.docx"
    project_docs_path = Path("docs") / output_filename
    desktop_path = Path("C:/Users/strel/Desktop") / output_filename
    brain_path = Path("C:/Users/strel/.gemini/antigravity/brain/b958a22b-49ff-459c-a191-6c697ec14334") / output_filename
    
    # Также синхронизируем английское имя для универсальности
    english_filename = "RevOps_AI_Supervisor_Partnership_Roadmap.docx"
    project_docs_en = Path("docs") / english_filename
    desktop_en = Path("C:/Users/strel/Desktop") / english_filename

    project_docs_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(project_docs_path)
    print(f"✅ Мастер-документ успешно сохранен в docs: {project_docs_path.absolute()}")
    print(f"📄 Размер: {project_docs_path.stat().st_size / 1024:.1f} KB")

    # Копирование на Рабочий стол и в Brain artifacts
    master_desktop = Path("C:/Users/strel/Desktop/Дорожная_Карта_RevOps_AI_Супервайзер_Мастер.docx")
    try:
        shutil.copy2(project_docs_path, master_desktop)
        print(f"📋 Создан мастер-файл на Рабочем столе: {master_desktop}")
    except Exception as e:
        print(f"⚠️ Ошибка создания мастер-файла: {e}")

    try:
        shutil.copy2(project_docs_path, desktop_path)
        print(f"📋 Обновлено на Рабочем столе: {desktop_path}")
    except Exception as e:
        print(f"ℹ️ Файл на Рабочем столе открыт в Word (заблокирован), создан параллельный мастер: {master_desktop}")

    try:
        shutil.copy2(project_docs_path, desktop_en)
        print(f"📋 Обновлено на Рабочем столе: {desktop_en}")
    except Exception:
        pass

    try:
        shutil.copy2(project_docs_path, brain_path)
        print(f"🧠 Скопировано в Brain Artifacts: {brain_path}")
    except Exception as e:
        print(f"⚠️ Ошибка копирования в Artifacts: {e}")

    try:
        shutil.copy2(project_docs_path, project_docs_en)
    except Exception:
        pass

    return project_docs_path

if __name__ == "__main__":
    generate_master_document()
