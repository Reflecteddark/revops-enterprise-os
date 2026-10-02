"""
Генератор Word-документа "Партнёрский Оффер для Интеграторов amoCRM и Битрикс24"
RevOps Enterprise OS V17.6

Цель: Компактный, высококонверсионный One-Pager / Two-Pager для руководителей
агентств внедрения CRM и частных интеграторов с предложением 25-30% ревшара (MRR).
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

# === КОРПОРАТИВНЫЕ ЦВЕТА ===
COLOR_NAVY = RGBColor(0x0F, 0x17, 0x2A)       # #0F172A
COLOR_PRIMARY = RGBColor(0x1E, 0x3A, 0x5F)    # #1E3A5F
COLOR_INDIGO = RGBColor(0x4F, 0x46, 0xE5)     # #4F46E5
COLOR_DARK = RGBColor(0x1E, 0x29, 0x3B)       # #1E293B
COLOR_MUTED = RGBColor(0x64, 0x74, 0x8B)      # #64748B
COLOR_SUCCESS = RGBColor(0x05, 0x96, 0x69)    # #059669

HEX_PRIMARY = "1E3A5F"
HEX_INDIGO = "4F46E5"
HEX_INDIGO_LIGHT = "EEF2FF"
HEX_ZEBRA = "F8FAFC"
HEX_GREEN = "059669"
HEX_GREEN_LIGHT = "ECFDF5"
HEX_BORDER = "CBD5E1"

def set_cell_background(cell, hex_color: str):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color=HEX_BORDER, sz="4"):
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
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color="CBD5E1", sz="4")

    # Header
    header_row = table.rows[0]
    for i, header_text in enumerate(headers):
        cell = header_row.cells[i]
        cell.text = header_text
        set_cell_background(cell, header_bg)
        set_cell_margins(cell, top=120, bottom=120, left=140, right=140)
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            for r in p.runs:
                r.font.name = 'Calibri'
                r.font.bold = True
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                r.font.size = Pt(9.5)

    # Data rows
    for r_idx, row_data in enumerate(rows):
        row = table.rows[r_idx + 1]
        is_even = (r_idx % 2 == 1)
        row_bg = HEX_ZEBRA if is_even else "FFFFFF"
        for c_idx, cell_text in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.text = str(cell_text)
            if row_bg != "FFFFFF":
                set_cell_background(cell, row_bg)
            set_cell_margins(cell, top=90, bottom=90, left=120, right=120)
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)
                for r in p.runs:
                    r.font.name = 'Calibri'
                    r.font.size = Pt(9)
                    r.font.color.rgb = COLOR_DARK

    if col_widths:
        for row in table.rows:
            for i, w in enumerate(col_widths):
                if i < len(row.cells):
                    row.cells[i].width = Inches(w)

    doc.add_paragraph()
    return table

def add_callout(doc, emoji: str, title: str, text: str, bg_hex=HEX_INDIGO_LIGHT, border_hex=HEX_INDIGO):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.rows[0].cells[0]
    cell.width = Inches(6.9)
    set_cell_background(cell, bg_hex)
    set_cell_margins(cell, top=120, bottom=120, left=160, right=160)

    borders_xml = f"""
    <w:tcBorders {nsdecls("w")}>
        <w:left w:val="single" w:sz="20" w:space="0" w:color="{border_hex}"/>
        <w:top w:val="none"/>
        <w:right w:val="none"/>
        <w:bottom w:val="none"/>
    </w:tcBorders>
    """
    cell._tc.get_or_add_tcPr().append(parse_xml(borders_xml))

    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    run_t = p.add_run(f"{emoji} {title}")
    run_t.font.name = "Calibri"
    run_t.font.size = Pt(10)
    run_t.font.bold = True
    run_t.font.color.rgb = COLOR_PRIMARY

    p2 = cell.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(0)
    run_b = p2.add_run(text)
    run_b.font.name = "Calibri"
    run_b.font.size = Pt(9)
    run_b.font.color.rgb = COLOR_DARK

    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_before = Pt(0)
    p_spacer.paragraph_format.space_after = Pt(2)

def generate_integrator_offer():
    doc = docx.Document()

    # Поля страницы 0.7" (компактный формат)
    for section in doc.sections:
        section.top_margin = Inches(0.65)
        section.bottom_margin = Inches(0.65)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Calibri'
    style_normal.font.size = Pt(10)
    style_normal.font.color.rgb = COLOR_DARK
    style_normal.paragraph_format.space_after = Pt(4)
    style_normal.paragraph_format.line_spacing = 1.12

    # Заголовочный блок
    p_badge = doc.add_paragraph()
    p_badge.paragraph_format.space_before = Pt(0)
    p_badge.paragraph_format.space_after = Pt(2)
    r_b = p_badge.add_run("B2B ПАРТНЁРСКАЯ ПРОГРАММА ДЛЯ ИНТЕГРАТОРОВ AMOCRM И БИТРИКС24")
    r_b.font.size = Pt(9)
    r_b.font.bold = True
    r_b.font.color.rgb = COLOR_INDIGO

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(4)
    r_t = p_title.add_run("Превратите внедрение CRM в пожизненный MRR-доход:\nдо 30% Rev-Share с каждого клиента")
    r_t.font.size = Pt(17)
    r_t.font.bold = True
    r_t.font.color.rgb = COLOR_NAVY

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(10)
    r_sub = p_sub.add_run(
        "Вы внедрили клиенту CRM. Мы подключаем автономного ИИ-Супервайзера, который контролирует 100% звонков, "
        "спасает брошенные сделки в воронке и защищает продление лицензий. Вы получаете до 36 000 ₽ с каждого клиента ежемесячно."
    )
    r_sub.font.size = Pt(9.5)
    r_sub.font.color.rgb = COLOR_MUTED

    # Раздел 1: Главная боль интегратора
    h1 = doc.add_heading("1. Скрытая проблема рынка интеграций CRM", level=2)
    h1.paragraph_format.space_before = Pt(6)
    h1.paragraph_format.space_after = Pt(3)

    p1 = doc.add_paragraph(
        "Каждый интегратор amoCRM и Битрикс24 сталкивается с одной и той же экономической ловушкой: "
        "после завершения базового внедрения (80 000 – 250 000 ₽) клиент уходит, а постоянный денежный поток (LTV) падает до нуля. "
        "Хуже того: через 2-3 месяца менеджеры клиента начинают саботировать CRM, бросать сделки без Next Step и не заполнять поля. "
        "Клиент разочаровывается в CRM («она нам не помогла») и не продлевает платные лицензии."
    )

    add_callout(doc, "🎯", "Суть нашего предложения интегратору",
        "Мы берем на себя пост-внедренческий надзор и аналитику. Мы подключаем готовую систему «RevOps AI-Супервайзер»: "
        "она через официальные API слушает 100% звонков, привязывает ошибки к сделкам в вашей воронке и выдает РОПу готовый Action-лист в Telegram. "
        "Вам не нужно нанимать аналитиков, писать код или вести техподдержку. Вы просто рекомендуете решение и получаете ежемесячный пассивный доход.",
        bg_hex=HEX_INDIGO_LIGHT, border_hex=HEX_INDIGO
    )

    # Раздел 2: Экономика партнерства (Калькулятор заработка)
    h2 = doc.add_heading("2. Партнерская сетка вознаграждения (Rev-Share)", level=2)
    h2.paragraph_format.space_before = Pt(6)
    h2.paragraph_format.space_after = Pt(3)

    revenue_table_headers = [
        "ТАРИФ КЛИЕНТА",
        "ЧИСЛО МЕНЕДЖЕРОВ",
        "ОПЛАТА КЛИЕНТА В МЕСЯЦ",
        "ВАШ РЕВШАР (25–30%)",
        "ВАШ ДОХОД ЗА ГОД С 1 КЛИЕНТА"
    ]
    revenue_rows = [
        ["Старт: Пилот (7 дней)", "Любой отдел", "29 000 ₽ (разово)", "20% = 5 800 ₽", "5 800 ₽ (сразу в День 1)"],
        ["Базовый", "3–5 менеджеров", "49 000 ₽ / месяц", "25% = 12 250 ₽ / мес", "147 000 ₽ / год"],
        ["Продвинутый", "6–10 менеджеров", "79 000 ₽ / месяц", "25% = 19 750 ₽ / мес", "237 000 ₽ / год"],
        ["Enterprise", "11–15 менеджеров", "120 000 ₽ / месяц", "30% = 36 000 ₽ / мес", "432 000 ₽ / год"],
    ]

    add_styled_table(doc, revenue_table_headers, revenue_rows, col_widths=[1.5, 1.4, 1.4, 1.4, 1.2], header_bg=HEX_PRIMARY)

    add_callout(doc, "💰", "Пример пассивного дохода интегратора с базы клиентов",
        "Если вы подключите всего 5 клиентов (например, 3 Базовых + 2 Продвинутых): "
        "ваш регулярный пассивный доход составит 76 250 ₽ в месяц (915 000 ₽ в год).\n"
        "При 10 активных клиентах на сопровождении — свыше 170 000 ₽ в месяц чистой прибыли без операционных расходов!",
        bg_hex=HEX_GREEN_LIGHT, border_hex=HEX_GREEN
    )

    # Раздел 3: Почему клиенты скажут вам спасибо
    h3 = doc.add_heading("3. 4 прямые выгоды для вашего агентства", level=2)
    h3.paragraph_format.space_before = Pt(6)
    h3.paragraph_format.space_after = Pt(3)

    benefits = [
        ("1. Защита продления лицензий CRM на 100%:",
         "Когда ИИ ежедневно возвращает собственнику от 400 000 до 1.8 млн ₽ за счет спасения зависших сделок, у клиента никогда не возникнет мысли отключить amoCRM или Битрикс24."),
        ("2. Нулевые затраты на разработку и инфраструктуру:",
         "Мы предоставляем полностью протестированный стек: коннекторы к amoCRM / Битрикс24, распознавание речи в защищенном контуре РФ (152-ФЗ), аудит-логи и готовый Telegram-бот."),
        ("3. Повышение ценности вашего основного чека:",
         "Вы можете упаковать сервис в свое пакетное предложение: «Внедрение CRM + 1 месяц ИИ-контроля продаж под ключ», отстраиваясь от конкурентов-демпингеров."),
        ("4. Прозрачные выплаты и юридическая чистота:",
         "Заключаем официальный партнерский договор (ИП / ООО / НПД). Выплаты партнерской комиссии осуществляются ежемесячно в течение 3 рабочих дней после поступления оплаты от клиента.")
    ]

    for title, desc in benefits:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        r_b = p.add_run(title + " ")
        r_b.font.bold = True
        r_b.font.color.rgb = COLOR_PRIMARY
        r_t = p.add_run(desc)

    # Раздел 4: Как мы работаем (3 простых шага)
    h4 = doc.add_heading("4. Порядок запуска партнерства: от контакта до первой выплаты", level=2)
    h4.paragraph_format.space_before = Pt(6)
    h4.paragraph_format.space_after = Pt(3)

    steps = [
        ["Шаг 1. Экспресс-знакомство (15 минут)", "Показываем вам работу системы на реальном демо-стенде. Согласовываем индивидуальные условия ревшара и подписываем NDA."],
        ["Шаг 2. Пилотный клиент из вашей базы", "Вы передаете контакт 1 лояльного клиента, которому хотите поднять продажи (или делаем совместный тест-драйв 3-5 звонков)."],
        ["Шаг 3. Мы внедряем — вы получаете комиссию", "Мы проводим пилот за 7 дней под ключ, презентуем собственнику спасенные сделки, закрываем на подписку и перечисляем ваш Rev-Share."]
    ]

    add_styled_table(doc, ["ЭТАП", "ЧТО ПРОИСХОДИТ"], steps, col_widths=[2.3, 4.6], header_bg=HEX_INDIGO)

    # Финальный призыв к действию (CTA) и контакты
    p_cta = doc.add_paragraph()
    p_cta.paragraph_format.space_before = Pt(8)
    p_cta.paragraph_format.space_after = Pt(4)
    p_cta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_cta = p_cta.add_run("ГОТОВЫ ОБСУДИТЬ ПАРТНЕРСТВО И ПОСМОТРЕТЬ ДЕМО?")
    r_cta.font.size = Pt(11)
    r_cta.font.bold = True
    r_cta.font.color.rgb = COLOR_NAVY

    p_contact = doc.add_paragraph()
    p_contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_contact.paragraph_format.space_before = Pt(0)
    p_contact.paragraph_format.space_after = Pt(0)
    r_c = p_contact.add_run(
        "Напишите основателю напрямую в Telegram для назначения 15-минутного созвона или запроса партнерского договора:\n"
        "Telegram: @dm1918 | Email: info@ai-rop.ru | Сайт платформы: https://ai-rop.ru\n"
        "RevOps Enterprise OS • ИИ-Супервайзер звонков & Партнерская программа 20% RevShare"
    )
    r_c.font.size = Pt(9.5)
    r_c.font.bold = True
    r_c.font.color.rgb = COLOR_NAVY

    # Сохранение файлов
    output_filename = "Оффер_Интеграторам_CRM_20_процентов.docx"
    partner_dir = Path(r"C:\Users\strel\Desktop\RevOps Platform\Каналы продаж\Партнёрка")
    partner_dir.mkdir(parents=True, exist_ok=True)
    desktop_dest = partner_dir / output_filename
    docs_dest = Path("docs") / output_filename

    docs_dest.parent.mkdir(parents=True, exist_ok=True)
    doc.save(docs_dest)
    doc.save(desktop_dest)
    print(f"✅ Партнерский оффер сохранен в docs: {docs_dest.absolute()}")
    print(f"📋 Скопировано в папку Партнёрка: {desktop_dest}")

    return docs_dest

if __name__ == "__main__":
    generate_integrator_offer()
