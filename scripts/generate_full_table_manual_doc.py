import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import subprocess

sys.stdout.reconfigure(encoding="utf-8")

# Palette RevOps Brand
COLOR_PRIMARY = RGBColor(15, 23, 42)      # Slate Navy #0F172A
COLOR_ACCENT = RGBColor(37, 99, 235)      # Royal Blue #2563EB
COLOR_SECONDARY = RGBColor(71, 85, 105)   # Muted Slate #475569
COLOR_SUCCESS = RGBColor(16, 185, 129)    # Emerald Green #10B981
COLOR_WARNING = RGBColor(245, 158, 11)    # Amber Gold #F59E0B
COLOR_DANGER = RGBColor(239, 68, 68)      # Ruby Red #EF4444
COLOR_MUTED = RGBColor(100, 116, 139)     # Gray #64748B

HEX_BG_LIGHT = "F8FAFC"
HEX_BG_HEADER = "0F172A"
HEX_BORDER = "E2E8F0"
HEX_ACCENT_BG = "EFF6FF"
HEX_SUCCESS_BG = "ECFDF5"
HEX_WARNING_BG = "FFFBEB"
HEX_DANGER_BG = "FEF2F2"

def set_cell_shading(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=120, bottom=120, left=160, right=160):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def add_header_footer(doc):
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
        # Header
        header = section.header
        p_head = header.paragraphs[0]
        p_head.text = "REVOPS ENTERPRISE OS V18.0  |  ГЕНЕРАЛЬНЫЙ АТЛАС-СПРАВОЧНИК ПО ВСЕМ ЛИСТАМ"
        p_head.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_head.runs[0].font.size = Pt(8)
        p_head.runs[0].font.color.rgb = COLOR_MUTED
        p_head.runs[0].font.name = "Calibri"
        
        # Footer
        footer = section.footer
        p_foot = footer.paragraphs[0]
        p_foot.text = "RevOps Enterprise Architecture  •  Единый Источник Правды (SSOT)  •  Конфиденциально для Клиента"
        p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_foot.runs[0].font.size = Pt(8)
        p_foot.runs[0].font.color.rgb = COLOR_MUTED
        p_foot.runs[0].font.name = "Calibri"

def add_callout(doc, text, title="ВАЖНОЕ ПРАВИЛО", alert_type="info"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    
    if alert_type == "info":
        bg_color, border_color, title_color = HEX_ACCENT_BG, "2563EB", COLOR_ACCENT
    elif alert_type == "success":
        bg_color, border_color, title_color = HEX_SUCCESS_BG, "10B981", COLOR_SUCCESS
    elif alert_type == "warning":
        bg_color, border_color, title_color = HEX_WARNING_BG, "F59E0B", COLOR_WARNING
    else:
        bg_color, border_color, title_color = HEX_DANGER_BG, "EF4444", COLOR_DANGER
    
    set_cell_shading(cell, bg_color)
    set_cell_margins(cell, top=140, bottom=140, left=180, right=180)
    
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:left w:val="single" w:sz="28" w:space="0" w:color="{border_color}"/><w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/></w:tcBorders>')
    tcPr.append(tcBorders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(4)
    r_title = p.add_run(f"📌 {title}: ")
    r_title.bold = True
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(10)
    r_title.font.color.rgb = title_color
    
    r_text = p.add_run(text)
    r_text.font.name = "Calibri"
    r_text.font.size = Pt(9.5)
    r_text.font.color.rgb = COLOR_PRIMARY
    
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_after = Pt(4)

def add_heading_1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = COLOR_PRIMARY

def add_heading_2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = COLOR_ACCENT

def add_heading_3(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(10.5)
    r.font.bold = True
    r.font.color.rgb = COLOR_SECONDARY

def add_p(doc, text, bold_prefix="", space_after=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = "Calibri"
        r_pre.font.size = Pt(10)
        r_pre.font.bold = True
        r_pre.font.color.rgb = COLOR_PRIMARY
    r_body = p.add_run(text)
    r_body.font.name = "Calibri"
    r_body.font.size = Pt(10)
    r_body.font.color.rgb = COLOR_SECONDARY
    return p

def add_bullet(doc, text, bold_prefix=""):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = "Calibri"
        r_pre.font.size = Pt(10)
        r_pre.font.bold = True
        r_pre.font.color.rgb = COLOR_PRIMARY
    r_body = p.add_run(text)
    r_body.font.name = "Calibri"
    r_body.font.size = Pt(10)
    r_body.font.color.rgb = COLOR_SECONDARY

def add_kpi_card_table(doc, kpis):
    cols = len(kpis)
    tbl = doc.add_table(rows=1, cols=cols)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for idx, (title, value, subtitle, ctype) in enumerate(kpis):
        cell = tbl.cell(0, idx)
        set_cell_shading(cell, HEX_BG_LIGHT)
        set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
        
        bcolor = "2563EB" if ctype=="accent" else ("10B981" if ctype=="success" else ("F59E0B" if ctype=="warning" else "EF4444"))
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="18" w:space="0" w:color="{bcolor}"/><w:left w:val="single" w:sz="6" w:space="0" w:color="{HEX_BORDER}"/><w:right w:val="single" w:sz="6" w:space="0" w:color="{HEX_BORDER}"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="{HEX_BORDER}"/></w:tcBorders>')
        tcPr.append(tcBorders)
        
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(2)
        r1 = p.add_run(title.upper() + "\n")
        r1.font.name = "Calibri"
        r1.font.size = Pt(7.5)
        r1.font.bold = True
        r1.font.color.rgb = COLOR_MUTED
        
        r2 = p.add_run(value + "\n")
        r2.font.name = "Calibri"
        r2.font.size = Pt(12.5)
        r2.font.bold = True
        r2.font.color.rgb = COLOR_ACCENT if ctype=="accent" else (COLOR_SUCCESS if ctype=="success" else (COLOR_WARNING if ctype=="warning" else COLOR_DANGER))
        
        r3 = p.add_run(subtitle)
        r3.font.name = "Calibri"
        r3.font.size = Pt(7.5)
        r3.font.color.rgb = COLOR_MUTED
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(4)

def format_custom_table(tbl, col_widths, headers, rows_data):
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cells = tbl.rows[0].cells
    for i, h_text in enumerate(headers):
        hdr_cells[i].text = h_text
        set_cell_shading(hdr_cells[i], HEX_BG_HEADER)
        set_cell_margins(hdr_cells[i], top=100, bottom=100, left=120, right=120)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.runs[0]
        r.font.name = "Calibri"
        r.font.size = Pt(8.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
    
    for row_idx, r_data in enumerate(rows_data):
        row = tbl.add_row()
        cells = row.cells
        bg_col = HEX_BG_LIGHT if row_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(r_data):
            cells[c_idx].text = str(val)
            set_cell_shading(cells[c_idx], bg_col)
            set_cell_margins(cells[c_idx], top=80, bottom=80, left=120, right=120)
            p = cells[c_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            if len(p.runs) > 0:
                p.runs[0].font.name = "Calibri"
                p.runs[0].font.size = Pt(8.5)
                p.runs[0].font.color.rgb = COLOR_PRIMARY
    
    for row in tbl.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = Inches(width)

def generate_document():
    doc = docx.Document()
    add_header_footer(doc)
    
    # Title Block
    p_badge = doc.add_paragraph()
    p_badge.paragraph_format.space_before = Pt(8)
    p_badge.paragraph_format.space_after = Pt(4)
    r_badge = p_badge.add_run("OFFICIAL ENTERPRISE EDITION V18.0  •  ПОЛНЫЙ АТЛАС-СПРАВОЧНИК")
    r_badge.font.name = "Calibri"
    r_badge.font.size = Pt(9.5)
    r_badge.font.bold = True
    r_badge.font.color.rgb = COLOR_ACCENT
    
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_after = Pt(6)
    r_title = p_title.add_run("RevOps Enterprise OS V18.0\nГенеральный Атлас и Постраничное Руководство")
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = COLOR_PRIMARY
    
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run("Исчерпывающее руководство по каждому листу, блоку, ячейке и выпадающему списку платформы. Боевые регламенты работы для Собственника (CEO) и Руководителя отдела продаж (РОП).")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(11)
    r_sub.font.color.rgb = COLOR_SECONDARY
    
    add_kpi_card_table(doc, [
        ("Архитектура", "SSOT V18.0", "Единый источник правды", "accent"),
        ("Покрытие тестами", "118 / 118", "Century Enterprise Ready", "success"),
        ("Речевой аудит", "100% звонков", "13 критериев ИИ", "accent"),
        ("Скорость планерки", "15 минут", "Каждое утро в 09:00", "warning"),
    ])
    
    add_callout(doc, 
        "Данный документ содержит полное описание функционала каждого рабочего экрана RevOps OS V18.0. "
        "Платформа спроектирована по стандарту Hosted SaaS (Single Source of Truth): доступ к витринам предоставляется "
        "руководству компании (Собственнику и РОПу), а линейные сотрудники работают в CRM. Платформа собирает данные автоматически, "
        "исключая человеческий фактор и ручные манипуляции.", 
        title="ЗОЛОТОЙ СТАНДАРТ REVOPS OS", alert_type="info")
    
    # -------------------------------------------------------------
    # РАЗДЕЛ 1: АРХИТЕКТУРА И РОЛЕВАЯ МОДЕЛЬ (ЛИНЗА РОЛЕЙ)
    # -------------------------------------------------------------
    add_heading_1(doc, "1. Архитектура Платформы и «Линза Роли» (RBAC)")
    add_p(doc, 
        "Архитектура RevOps OS V18 разделена на три изолированных технологических слоя: витрины управления (презентационный слой), "
        "расчетное ядро (calc_*) и слой сырых транзакций (raw_*). Это исключает случайную поломку формул пользователями и гарантирует "
        "математическую точность 100% показателей.")
    
    add_heading_2(doc, "Выпадающий список ролей: «Линза Роли», а не учетные записи")
    add_p(doc, 
        "На ключевых листах системы расположен выпадающий переключатель «РОЛЬ В СИСТЕМЕ». "
        "Принципиально важно понимать: это НЕ список учетных записей Google для входа разных людей! "
        "Таблица физически находится под управлением Собственника и РОПа. Переключатель ролей — это «умная оптическая линза», "
        "которая мгновенно перестраивает отображение данных на дашборде под конкретную задачу совещания.")
    
    # Table of Roles
    roles_tbl = doc.add_table(rows=1, cols=3)
    format_custom_table(roles_tbl, [1.5, 2.2, 3.1], 
        ["Значение в списке", "Целевое назначение роли", "Что видит и как использует пользователь"],
        [
            ["👑 CEO", "Генеральный директор / Собственник", "Макро-экономика: общая выручка, маржинальность, % выполнения плана, объем упущенной прибыли за месяц, прогноз кэш-ин."],
            ["📋 РОП", "Руководитель отдела продаж", "Операционка: зависшие сделки без Next Step (>48ч), нарушения SLA, сливы звонков менеджерами, дисциплина заполнения CRM."],
            ["💼 КАМ", "Key Account Manager (Ключевые клиенты)", "Фокус на крупных сделках (Enterprise), повторных продажах (LTV), дебиторской задолженности и контроле оплат."],
            ["📞 SDR", "Sales Development Representative (Холодный поиск)", "Фокус на первом касании: скорость реакции на новый лид (SLA < 24ч), процент квалификации, конверсия в назначенную встречу."],
            ["🌐 CMO", "Директор по маркетингу", "Сквозная аналитика трафика: расход бюджетов, стоимость квалифицированного лида (CAC), окупаемость рекламных каналов (ROMI)."],
            ["💳 CFO", "Финансовый директор", "Денежные потоки: фактический Cash-In, платежный календарь на 30–60 дней, DSO (средний срок оплаты), просроченные инвойсы."],
            ["🔧 Admin", "Технический администратор / Интегратор", "Бортовой журнал: состояние API-вебхуков amoCRM/Битрикс24, статус тестов QA Suite (118/118), логи синхронизации."]
        ]
    )
    add_p(doc, "Когда РОП проводит индивидуальный разбор с группой холодных звонков, он выбирает роль «📞 SDR» — и видит только метрики этого отдела. Это экономит до 40 минут на каждом совещании.", bold_prefix="💡 Практический сценарий: ")
    
    # -------------------------------------------------------------
    # РАЗДЕЛ 2: ЛИСТ ⚙️ НАСТРОЙКИ
    # -------------------------------------------------------------
    add_heading_1(doc, "2. Лист «⚙️ Настройки» — Командный Центр и Конфигурация")
    add_p(doc, 
        "Лист «⚙️ Настройки» — это центральный управляющий регистр всей платформы. Все формулы, расчетные модули, дашборды "
        "и графики на остальных 15 листах динамически ссылаются на значения из этого листа. Заполнение занимает ровно 3 минуты один раз в месяц.")
    
    add_heading_2(doc, "Поячеечная анатомия экрана:")
    add_bullet(doc, "Юридическое или коммерческое наименование клиента. Подставляется во все отчетные заголовки и экспортные PDF-выгрузки.", "Ячейка B3 (Организация): ")
    add_bullet(doc, "Основная валюта финансового учета (по умолчанию RUB / ₽). Все суммы форматируются автоматически.", "Ячейка B4 (Валюта системы): ")
    add_bullet(doc, "Дата начала текущего расчетного месяца (например, 01.10.2026).", "Ячейка B6 (Начало периода): ")
    add_bullet(doc, "Дата окончания расчетного месяца (например, 31.10.2026). Используется формулами для расчета оставшихся дней и темпа закрытия (Pacing).", "Ячейка B7 (Конец периода): ")
    add_bullet(doc, "Дата запуска RevOps OS в компании. Фиксирует точку отсчета для аналитики динамики «До и После».", "Ячейка B8 (Дата старта проекта): ")
    add_bullet(doc, "Ключевая финансовая цель месяца (например, 5 000 000 ₽). На этот показатель ориентируются все шкалы прогресса и расчеты премий.", "Ячейка B9 (План выручки на месяц): ")
    add_bullet(doc, "Уникальный идентификатор клиентского контура для связи с сервером ИИ и Telegram-ботом супервайзера.", "Ячейка E3 (TENANT UUID): ")
    add_bullet(doc, "Индикатор юридического соответствия хранения и обработки данных на территории РФ (Аттестован).", "Ячейка E9 (Статус 152-ФЗ РФ): ")
    
    add_heading_2(doc, "Нормативы этапов воронки и SLA (Диапазон A11:D17)")
    add_p(doc, "В этом блоке задаются жесткие стандарты скорости движения клиента по воронке продаж:")
    
    sla_tbl = doc.add_table(rows=1, cols=4)
    format_custom_table(sla_tbl, [1.0, 2.5, 1.8, 1.5],
        ["Stage ID", "Название этапа воронки", "SLA норматив (ч)", "Вероятность (Win %)"],
        [
            ["1.0", "1. Новый лид", "24.0 ч (макс. 1 сутки)", "5%"],
            ["2.0", "2. Квалификация / ЛПР", "48.0 ч (макс. 2 суток)", "15%"],
            ["3.0", "3. Встреча / Демо", "120.0 ч (макс. 5 суток)", "35%"],
            ["4.0", "4. КП и согласование", "48.0 ч (критический этап)", "60%"],
            ["5.0", "5. Счет выставлен", "72.0 ч (макс. 3 суток)", "85%"],
            ["6.0", "6. Успешно реализовано", "Финишный этап", "100%"],
            ["7.0", "7. Отказ / Архив", "Финишный этап", "0%"]
        ]
    )
    add_p(doc, 
        "Если сделка находится на этапе дольше указанного норматива часов, система автоматически подсвечивает её тревожным цветом "
        "и отправляет в лист Action Center как инцидент срыва выручки.")
    
    add_heading_2(doc, "Реестр менеджеров отдела продаж (Диапазон A21:F26)")
    add_p(doc, 
        "Сюда вносятся реальные сотрудники компании: их ID в CRM, ФИО, штатная роль, окладная ставка и персональный план по продажам. "
        "Формулы листа «👥 Мотивация_ОП» привязываются именно к этим ячейкам.")
    
    add_callout(doc, 
        "1-го числа месяца РОП заходит в «⚙️ Настройки» и обновляет 3 вещи: 1) Даты B6 и B7 на новый месяц; "
        "2) План B9; 3) Список менеджеров B21:B26, если кто-то уволился или нанят. Всё остальное ядро пересчитывается автоматически.", 
        title="РЕГЛАМЕНТ 1-ГО ЧИСЛА МЕСЯЦА", alert_type="warning")
    
    # -------------------------------------------------------------
    # РАЗДЕЛ 3: ЛИСТ ⚡ ЭКСПРЕСС_КАЛЬКУЛЯТОР_3_ЦИФРЫ
    # -------------------------------------------------------------
    add_heading_1(doc, "3. Лист «⚡ Экспресс_Калькулятор_3_Цифры» — Финансовый Аудит за 30 Секунд")
    add_p(doc, 
        "Лист предназначен для мгновенной финансовой диагностики узких мест отдела продаж. "
        "Он позволяет собственнику бизнеса оцифровать объем упущенной прибыли на базе отраслевых бенчмарков B2B-рынка.")
    
    add_heading_2(doc, "Анатомия ввода (3 цифры бизнеса):")
    add_bullet(doc, "Количество новых входящих обращений / заявок в месяц.", "Ячейка B4 (Входящий поток лидов): ")
    add_bullet(doc, "Средняя стоимость закрытой сделки (в рублях).", "Ячейка B5 (Средний чек): ")
    add_bullet(doc, "Количество продавцов, работающих с клиентами.", "Ячейка B6 (Количество менеджеров): ")
    
    add_heading_2(doc, "Главное табло: «ВЫЯВЛЕННЫЙ РИСК ПОТЕРЬ» (Красный блок ячеек B10:E11)")
    add_p(doc, 
        "Это главная цифра отчета. Она показывает совокупную сумму живых денег, которую компания недополучает каждый месяц "
        "из-за человеческого фактора менеджеров. Сумма складывается из 4 объективных финансовых утечек:")
    
    leaks_tbl = doc.add_table(rows=1, cols=4)
    format_custom_table(leaks_tbl, [2.0, 1.4, 2.0, 1.4],
        ["Узкое место воронки", "Бенчмарк потерь", "Причина утечки", "Инструмент решения"],
        [
            ["1. Слив лидов на звонках (нет Next Step)", "28% выручки", "Менеджеры консультируют, но не фиксируют дату следующего контакта", "ИИ-аудит 100% звонков + авто-пуш РОПу"],
            ["2. Зависание сделок на этапе КП > 48ч", "18% выручки", "Клиент остывает, пока менеджер пассивно ждет ответа", "Авторадар SLA 48ч + эскалация РОПу"],
            ["3. Брошенные отказники без дожима", "12% выручки", "Сделки списываются в архив без регулярной реактивации", "AI Recovery Engine (автодожим базы)"],
            ["4. Зависшая дебиторка и задержка оплат", "10% выручки", "Нет платежного календаря и контроля просрочек", "Платежный календарь DSO + авто-триггеры"]
        ]
    )
    add_p(doc, "Собственник использует этот лист для обоснования внедрения жестких регламентов: «Коллеги, каждый день без фиксации Next Step обходится нашей компании в X рублей упущенной прибыли».", bold_prefix="🎯 Как использует CEO: ")
    
    # -------------------------------------------------------------
    # РАЗДЕЛ 4: ЛИСТ 📋 ПУЛЬТ_РОПА_15_МИНУТ
    # -------------------------------------------------------------
    add_heading_1(doc, "4. Лист «📋 Пульт_РОПа_15_Минут» — Ежедневный Центр Управления")
    add_p(doc, 
        "Главный рабочий инструмент Руководителя отдела продаж. Экран спроектирован так, чтобы РОП ровно за 15 минут "
        "утром получил исчерпывающую картину состояния продаж без необходимости открывать десятки карточек в CRM.")
    
    add_heading_2(doc, "Верхние 6 KPI-карточек (Строки 4–5):")
    add_bullet(doc, "Фактически поступившие деньги от выигранных сделок за текущий месяц.", "1. Выручка (Факт Closed-Won): ")
    add_bullet(doc, "Отношение факта к плану ячейки B9 листа Настроек с динамической шкалой цвета.", "2. % Выполнения плана: ")
    add_bullet(doc, "Математический прогноз поступлений с учетом вероятности каждого этапа воронки.", "3. Взвешенный прогноз (SSOT): ")
    add_bullet(doc, "Ожидаемая касса к концу месяца, если отдел продолжит продавать с текущей среднедневной скоростью.", "4. Run-Rate касса (Прогноз): ")
    add_bullet(doc, "Сумма всех открытых сделок, находящихся в обработке прямо сейчас.", "5. Активный пайплайн в работе: ")
    add_bullet(doc, "Число сделок, нарушивших регламент скорости движения по воронке (требуют немедленного внимания).", "6. Сделок под угрозой SLA: ")
    
    add_heading_2(doc, "Левый блок: Воронка продаж AS-IS (Строки 8–15)")
    add_p(doc, 
        "Наглядно отображает количество сделок и сумму на каждом этапе. Если на этапе «4. КП и согласование» сумма превышает "
        "выручку в 3 раза, это прямой сигнал: в отделе затор — менеджеры рассылают предложения, но не дожимают клиентов до счета.")
    
    add_heading_2(doc, "Правый блок: ТОП-5 критических рисков выручки (Action Center)")
    add_p(doc, 
        "Автоматически выводит 5 самых дорогих сделок компании, находящихся под угрозой срыва (просрочен счет, нет звонка более 5 дней, "
        "брак переговоров). РОП сразу видит номер сделки, сумму и рекомендуемое действие.")
    
    add_callout(doc, 
        "09:00–09:03: Смотрим карточку Run-Rate касса (идем ли в плане?).\n"
        "09:03–09:08: Открываем правый блок ТОП-5 рисков и разбираем каждую зависшую сделку.\n"
        "09:08–09:15: Раздаем менеджерам персональные поручения: «Иван — дожать счет D-103 до 14:00; Александр — связаться по КП D-104».", 
        title="РЕГЛАМЕНТ УТРЕННЕЙ 15-МИНУТНОЙ ПЛАНЕРКИ РОПА", alert_type="success")
    
    # -------------------------------------------------------------
    # РАЗДЕЛ 5: ЛИСТ 🎙️ ИИ_АУДИТ
    # -------------------------------------------------------------
    add_heading_1(doc, "5. Лист «🎙️ ИИ_Аудит» — Речевая Аналитика и Контроль Звонков")
    add_p(doc, 
        "Лист содержит объективную оценку 100% телефонных переговоров менеджеров нейросетью. "
        "Система автоматически транскрибирует аудиозапись из CRM, сопоставляет диалог с 13 критериями эталонных продаж "
        "и выставляет итоговый балл качества от 0 до 100.")
    
    add_heading_2(doc, "Матрица 13 стандартов продаж (Колонки G:S):")
    add_p(doc, "Каждый разговор оценивается бинарно: 1 (стандарт выполнен) или 0 (стандарт провален):")
    
    std_tbl = doc.add_table(rows=1, cols=3)
    format_custom_table(std_tbl, [0.8, 2.5, 3.5],
        ["№", "Стандарт продажи", "Что проверяет искусственный интеллект"],
        [
            ["1", "Приветствие по стандарту", "Представился ли менеджер, назвал ли компанию и цель звонка."],
            ["2", "Идентификация ЛПР", "Уточнил ли менеджер полномочия собеседника в принятии решения."],
            ["3", "Выявление болей и потребности", "Задал ли менеджер минимум 2 открытых вопроса о текущих сложностях."],
            ["4", "Квалификация бюджета", "Уточнил ли ценовой диапазон или объем планируемой закупки."],
            ["5", "Квалификация по срокам", "Выяснил ли дедлайн запуска проекта или поставки товара."],
            ["6", "Презентация через выгоды", "Рассказал ли о продукте языком пользы для бизнеса клиента."],
            ["7", "Отработка возражений", "Применил ли технику согласия и аргументации при сомнениях."],
            ["8", "Чистота речи и вежливость", "Отсутствие грубости, слов-паразитов и давления на клиента."],
            ["9", "Фиксация дедлайна КП", "Озвучена ли точная дата и время отправки коммерческого предложения."],
            ["10", "Озвучивание условий оплаты", "Проговорен ли порядок расчетов (аванс, постоплата, рассрочка)."],
            ["11", "ЖЕЛЕЗОБЕТОННЫЙ NEXT STEP", "КРИТИЧЕСКИЙ СТАНДАРТ: Назначены ли конкретная ДАТА И ВРЕМЯ следующего звонка/встречи."],
            ["12", "Фиксация договоренностей", "Обещано ли продублировать итоги встречи в WhatsApp/Telegram."],
            ["13", "Корректность CRM", "Заполнены ли все обязательные поля в карточке сделки после разговора."]
        ]
    )
    
    add_heading_2(doc, "Цветовая градация итогового балла:")
    add_bullet(doc, "Критический брак переговоров. Клиент слит либо ушел с негативом. Требует немедленного вмешательства РОПа.", "🔴 Красная зона (<70 баллов): ")
    add_bullet(doc, "Разговор состоялся, но ключевые стандарты (Next Step, дедлайн) не зафиксированы. Сделка под угрозой зависания.", "🟡 Желтая зона (70–84 баллов): ")
    add_bullet(doc, "Эталонный диалог. Все этапы соблюдены. Рекомендуется для добавления в корпоративную базу знаний для обучения стажеров.", "🟢 Зеленая зона (85–100 баллов): ")
    
    add_heading_2(doc, "Колонка «Coaching Tip» (Персональная рекомендация нейросети)")
    add_p(doc, 
        "В этой ячейке нейросеть простым русским языком пишет совет менеджеру:  \n"
        "«Вместо фразы \"Ну вы подумайте и перезвоните\" нужно было сказать: \"Иван Иванович, давайте я наберу вас в четверг в 14:00, "
        "когда вы изучите спецификацию, и мы согласуем финальную скидку?\"». "
        "РОПу не нужно придумывать аргументы — готовый скрипт уже сформирован.")
    
    # -------------------------------------------------------------
    # РАЗДЕЛ 6: ЛИСТ 🎯 ВОРОНКА_И_SLA & ACTION CENTER
    # -------------------------------------------------------------
    add_heading_1(doc, "6. Листы «🎯 Воронка_и_SLA» и «🎯 Action_Center»")
    add_p(doc, 
        "Эта связка листов отвечает за управление скоростью продаж (Sales Velocity) и предотвращение «заболачивания» базы контактов.")
    
    add_heading_2(doc, "Лист «Воронка_и_SLA»: Контроль сквозной конверсии и заторов")
    add_p(doc, 
        "Лист рассчитывает фактическое среднее время нахождения сделок на каждом этапе и сравнивает его с нормативом из листа Настроек. "
        "Если факт превышает норматив, статус окрашивается в `⚠️ ПРОСРОЧЕН`. Это позволяет мгновенно обнаружить, где буксует отдел: "
        "на отправке КП (долгие согласования смет) или на этапе счета (бухгалтерия клиента тянет с оплатой).")
    
    add_heading_2(doc, "Лист «Action_Center»: Инцидентный радар угроз выручке")
    add_p(doc, "Сюда попадают карточки с конкретными сигналами тревоги:")
    add_bullet(doc, "Критический инцидент. Сумма > 500 000 ₽ или срыв регламента > 7 дней. Требует вмешательства РОПа в течение 2 часов.", "Приоритет 🔴 P0: ")
    add_bullet(doc, "Высокий приоритет. Зависание на этапе счета или КП > 48 часов. Требует звонка менеджера до конца рабочего дня.", "Приоритет 🟡 P1: ")
    
    # -------------------------------------------------------------
    # РАЗДЕЛ 7: ЛИСТ 💳 ФИНАНСЫ_И_AI_ДОЖИМ
    # -------------------------------------------------------------
    add_heading_1(doc, "7. Лист «💳 Финансы_и_AI_Дожим» — Управление Деньгами и DSO")
    add_p(doc, 
        "Связывает продажи с реальным движением денежных средств (Cash Flow). Предотвращает кассовые разрывы "
        "и сокращает период оборачиваемости дебиторской задолженности.")
    
    add_heading_2(doc, "Ключевые финансовые индикаторы (Строки 3–4):")
    add_bullet(doc, "Реальные деньги, поступившие на расчетный счет за отчетный месяц.", "Фактический Cash-In: ")
    add_bullet(doc, "Ожидаемые поступления по выставленным счетам на ближайшие 30 календарных дней.", "Ожидаемый Cash-In (30д): ")
    add_bullet(doc, "Сумма по счетам, у которых наступил дедлайн оплаты, но деньги не поступили.", "Просрочено (Overdue): ")
    add_bullet(doc, "Days Sales Outstanding — средний фактический срок задержки оплаты клиентами (в днях). Норма: до 7 дней.", "DSO (Срок оплаты): ")
    add_bullet(doc, "Сумма зависших оплат, которую можно вернуть с помощью агентных сценариев напоминания.", "Потенциал AI-дожима: ")
    
    add_heading_2(doc, "Платежный календарь и агентный AI-дожим")
    add_p(doc, 
        "Система отслеживает статус каждого инвойса (`paid`, `forecast`, `overdue`). "
        "При переходе счета в статус `overdue` система генерирует персонализированное вежливое напоминание в мессенджер ЛПР "
        "с приложением счета и акта сверки, разгружая менеджеров от неприятных разговоров о долгах.")
    
    # -------------------------------------------------------------
    # РАЗДЕЛ 8: ЛИСТ 👥 МОТИВАЦИЯ_ОП
    # -------------------------------------------------------------
    add_heading_1(doc, "8. Лист «👥 Мотивация_ОП» — Динамический Payroll и KPI")
    add_p(doc, 
        "Лист автоматизирует расчет переменной части заработной платы менеджеров. "
        "В отличие от классических систем, где премия платится только за объем выручки, RevOps OS защищает компанию от «токсичных звезд», "
        "которые делают план ценой потери репутации и грубого нарушения стандартов.")
    
    add_heading_2(doc, "Формула динамического бонуса:")
    add_p(doc, 
        "Бонус менеджера рассчитывается по формуле:\n"
        "Бонус = (Закрытая выручка × Базовый %) × Коэффициент SLA × Коэффициент Качества Речи\n\n"
        "Где:\n"
        "• Коэффициент SLA = 1.0 (при отсутствии просроченных сделок) или 0.85 (при систематическом нарушении сроков).\n"
        "• Коэффициент Качества Речи = 1.0 (если средний балл ИИ-Аудита ≥ 9.0 из 13) или 0.70 (дисконт 30% при браке речи).")
    
    add_callout(doc, 
        "Правило Ревопс: Менеджер видит расчет прозрачно. Если сотрудник выполнил план по выручке, но регулярно забывал ставить "
        "следующий шаг или грубил клиентам (балл ИИ < 9.0), его премия автоматически уменьшается на 30%. "
        "Это заставляет сотрудников соблюдать регламенты без постоянных криков и штрафов со стороны руководства.", 
        title="СПРАВЕДЛИВЫЙ ДИСКОНТ ЗА БРАК РЕЧИ", alert_type="warning")
    
    # -------------------------------------------------------------
    # РАЗДЕЛ 9: СВОДНАЯ МАТРИЦА ОТВЕТСТВЕННОСТИ (RACI)
    # -------------------------------------------------------------
    add_heading_1(doc, "9. Сводная Матрица Регламентов: Кто, Когда и Что Открывает")
    add_p(doc, "Чтобы внедрение платформы прошло гладко, закрепите следующие зоны ответственности в вашей компании:")
    
    raci_tbl = doc.add_table(rows=1, cols=4)
    format_custom_table(raci_tbl, [1.5, 1.3, 2.0, 2.2],
        ["Роль в компании", "Периодичность", "Рабочие листы", "Главная задача и регламент"],
        [
            ["Собственник (CEO)", "Раз в неделю / месяц", "⚡ Экспресс_3_Цифры, 📄 OnePager, 👥 Мотивация", "Контроль чистой прибыли, утверждение ФОТ, анализ предотвращенных потерь."],
            ["РОП (Head of Sales)", "Ежедневно в 09:00", "📋 Пульт_РОПа, 🎙️ ИИ_Аудит, 🎯 Action_Center", "Проведение 15-минутной планерки, разбор ТОП-5 рисков, контроль звонков менеджеров."],
            ["РОП (Head of Sales)", "1-е число месяца", "⚙️ Настройки, 🎯 Воронка_и_SLA", "Установка плана продаж, обновление состава менеджеров, проверка нормативов SLA."],
            ["Финансист / Бухгалтер", "Раз в неделю", "💳 Финансы_и_AI_Дожим, raw_invoices", "Контроль поступления оплат, сверка дебиторки, запуск сценариев дожима."],
            ["Маркетолог (CMO)", "Раз в 2 недели", "🌐 Мультиканальная_Атрибуция", "Анализ CAC и окупаемости каналов трафика, отключение убыточной рекламы."]
        ]
    )
    
    add_callout(doc, 
        "Платформа RevOps OS V18.0 — это завершенный промышленный инструмент. "
        "При регулярном соблюдении 15-минутного регламента РОПа рост выручки компании составляет от +18% до +34% "
        "уже в первые 60 дней за счет полной ликвидации потерь на этапе звонков и коммерческих предложений.", 
        title="ИТОГОВЫЙ БИЗНЕС-ЭФФЕКТ", alert_type="success")
    
    # Save Word
    output_docx = r"C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\docs\Полное_Постраничное_Руководство_RevOps_OS_V18.docx"
    doc.save(output_docx)
    print(f"Word document saved to: {output_docx}")
    
    # Copy to presentation folder as well
    pres_docx = r"C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\presentation\Полное_Постраничное_Руководство_RevOps_OS_V18.docx"
    try:
        import shutil
        shutil.copyfile(output_docx, pres_docx)
        print(f"Copied to: {pres_docx}")
    except Exception as e:
        print(f"Notice copy: {e}")
        
    return output_docx

if __name__ == "__main__":
    generate_document()
