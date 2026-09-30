import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import shutil

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

DOCS_DIR = r"C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\docs"
os.makedirs(DOCS_DIR, exist_ok=True)
OUT_DOCX = os.path.join(DOCS_DIR, "Дорожная_Карта_RevOps_AI_Супервайзер.docx")
DESKTOP_DOCX = os.path.expanduser(r"~\Desktop\Дорожная_Карта_RevOps_AI_Супервайзер.docx")
BRAIN_DOCX = r"C:\Users\strel\.gemini\antigravity\brain\b958a22b-49ff-459c-a191-6c697ec14334\Дорожная_Карта_RevOps_AI_Супервайзер.docx"

doc = docx.Document()

# Page setup: A4, 2 cm margins
for section in doc.sections:
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)

# Color Palette Constants
COLOR_PRIMARY = RGBColor(15, 23, 42)      # #0F172A Midnight Navy
COLOR_ACCENT = RGBColor(79, 70, 229)     # #4F46E5 Indigo
COLOR_SUCCESS = RGBColor(21, 128, 61)    # #15803D Emerald
COLOR_MUTED = RGBColor(100, 116, 139)    # #64748B Slate Gray
COLOR_DARK = RGBColor(30, 41, 59)        # #1E293B

HEX_PRIMARY = "0F172A"
HEX_ACCENT = "4F46E5"
HEX_ACCENT_LIGHT = "EEF2FF"
HEX_SUCCESS_LIGHT = "ECFDF5"
HEX_WARNING_LIGHT = "FFFBEB"
HEX_SURFACE = "F8FAFC"
HEX_BORDER = "CBD5E1"

# Helper XML styling
def set_cell_background(cell, hex_color):
    shading_xml = f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>'
    cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))

def set_cell_margins(cell, top=120, bottom=120, left=160, right=160):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color="E2E8F0", sz="4"):
    tblPr = table._tbl.tblPr
    borders_xml = f"""
    <w:tblBorders {nsdecls("w")}>
        <w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:left w:val="none"/>
        <w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:right w:val="none"/>
        <w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:insideV w:val="none"/>
    </w:tblBorders>
    """
    tblPr.append(parse_xml(borders_xml))

def add_callout(doc, title, text, bg_hex=HEX_ACCENT_LIGHT, border_hex=HEX_ACCENT):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.rows[0].cells[0]
    cell.width = Inches(6.8)
    set_cell_background(cell, bg_hex)
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)
    
    # Left border only
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
    p.paragraph_format.space_after = Pt(4)
    run_t = p.add_run(title)
    run_t.font.name = "Arial"
    run_t.font.size = Pt(11)
    run_t.font.bold = True
    run_t.font.color.rgb = COLOR_PRIMARY
    
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(0)
    run_b = p2.add_run(text)
    run_b.font.name = "Arial"
    run_b.font.size = Pt(9.5)
    run_b.font.color.rgb = COLOR_DARK
    
    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_before = Pt(0)
    p_spacer.paragraph_format.space_after = Pt(6)

# ==================== TITULAR / HEADER ====================
p_badge = doc.add_paragraph()
p_badge.paragraph_format.space_before = Pt(0)
p_badge.paragraph_format.space_after = Pt(4)
r_badge = p_badge.add_run("B2B ПАРТНЕРСКИЙ РЕГЛАМЕНТ • REVOPS REVENUE & VOICE GUARDIAN")
r_badge.font.name = "Arial"
r_badge.font.size = Pt(8.5)
r_badge.font.bold = True
r_badge.font.color.rgb = COLOR_ACCENT

p_title = doc.add_paragraph()
p_title.paragraph_format.space_before = Pt(2)
p_title.paragraph_format.space_after = Pt(6)
r_title = p_title.add_run("Дорожная Карта Партнерства:\nИИ-Супервайзер Звонков & Детектив Выручки")
r_title.font.name = "Georgia"
r_title.font.size = Pt(22)
r_title.font.bold = True
r_title.font.color.rgb = COLOR_PRIMARY

p_sub = doc.add_paragraph()
p_sub.paragraph_format.space_before = Pt(0)
p_sub.paragraph_format.space_after = Pt(16)
r_sub = p_sub.add_run("Практический регламент запуска речевого ИИ-контроля 100% звонков, поиска скрытых резервов воронки в CRM и ежедневного мониторинга отдела продаж в Telegram.")
r_sub.font.name = "Arial"
r_sub.font.size = Pt(10.5)
r_sub.font.color.rgb = COLOR_MUTED

# Meta Grid (Table)
meta_tbl = doc.add_table(rows=2, cols=3)
meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_tbl.autofit = False
set_table_borders(meta_tbl, color="E2E8F0", sz="4")

headers_meta = ["ФОРМАТ СОТРУДНИЧЕСТВА", "СТОИМОСТЬ ПИЛОТА (СТАРТ)", "РЕГУЛЯРНЫЙ НАДЗОР"]
vals_meta = ["Удаленный RevOps-шериф (НПД)", "29 000 ₽ (3–5 дней запуска)", "29 000 ₽ / месяц (подписка)"]

for c_idx in range(3):
    c_h = meta_tbl.rows[0].cells[c_idx]
    c_h.width = Inches(2.26)
    set_cell_background(c_h, "F8FAFC")
    set_cell_margins(c_h, top=80, bottom=60, left=100, right=100)
    p = c_h.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(headers_meta[c_idx])
    r.font.name = "Arial"
    r.font.size = Pt(8)
    r.font.bold = True
    r.font.color.rgb = COLOR_MUTED

    c_v = meta_tbl.rows[1].cells[c_idx]
    c_v.width = Inches(2.26)
    set_cell_background(c_v, "FFFFFF")
    set_cell_margins(c_v, top=80, bottom=100, left=100, right=100)
    p = c_v.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(vals_meta[c_idx])
    r.font.name = "Arial"
    r.font.size = Pt(9.5)
    r.font.bold = True
    r.font.color.rgb = COLOR_ACCENT if c_idx == 1 else COLOR_PRIMARY

p_div = doc.add_paragraph()
p_div.paragraph_format.space_before = Pt(12)
p_div.paragraph_format.space_after = Pt(12)

# ==================== SECTION 1 ====================
h1 = doc.add_heading(level=1)
h1.paragraph_format.space_before = Pt(14)
h1.paragraph_format.space_after = Pt(6)
r_h1 = h1.add_run("1. Проблема рынка: Почему старый подход больше не работает")
r_h1.font.name = "Georgia"
r_h1.font.size = Pt(14)
r_h1.font.bold = True
r_h1.font.color.rgb = COLOR_PRIMARY

doc.add_paragraph(
    "В 90% отделов продаж малого и среднего B2B-бизнеса (от 3 до 10 менеджеров) собственник сталкивается с тремя хроническими проблемами:\n"
    "1. Менеджеры работают как «справочное бюро»: консультируют входящие лиды, называют цену, но не фиксируют дату следующего контакта (Next Step) и отпускают клиента со словами «подумайте». До 30% рекламного бюджета сливается в песок прямо на этапе первого звонка.\n"
    "2. Руководитель отдела продаж (РОП) физически не успевает слушать звонки: из 500 звонков в неделю РОП слушает максимум 3–5 случайных диалогов. Контроль ведется вслепую.\n"
    "3. Иллюзорный пайплайн в CRM: в воронке висят миллионы рублей на этапе «КП отправлено», но 70% из них — брошенные «зомби-сделки», по которым клиенты уже ушли к конкурентам."
)

add_callout(
    doc,
    "💡 Главный вывод собственника",
    "Нанимать штатного аналитика и контролера качества — это от 100 000 до 150 000 ₽ в месяц с налогами и рабочим местом. Внедрять тяжелые софтверные комбайны — долго и встречает саботаж менеджеров. Решение — подключить автономного ИИ-супервайзера, который слушает 100% звонков и связывает ошибки речи с реальными деньгами в CRM.",
    bg_hex=HEX_ACCENT_LIGHT,
    border_hex=HEX_ACCENT
)

# ==================== SECTION 2 ====================
h1 = doc.add_heading(level=1)
h1.paragraph_format.space_before = Pt(14)
h1.paragraph_format.space_after = Pt(6)
r_h1 = h1.add_run("2. Архитектура единого продукта: 3 компонента в 1 решении")
r_h1.font.name = "Georgia"
r_h1.font.size = Pt(14)
r_h1.font.bold = True
r_h1.font.color.rgb = COLOR_PRIMARY

doc.add_paragraph(
    "Продукт объединяет три важнейших контура управления продажами в единый бесшовный сервис без необходимости ставить сложный софт сотрудникам:"
)

# Table 3 components
tbl_comp = doc.add_table(rows=4, cols=3)
tbl_comp.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl_comp.autofit = False
set_table_borders(tbl_comp, color="CBD5E1", sz="4")

col_widths = [Inches(1.8), Inches(2.5), Inches(2.5)]
row_data = [
    ["КОМПОНЕНТ", "ЧТО ДЕЛАЕТ В СИСТЕМЕ", "ПОЛЬЗА ДЛЯ ДИРЕКТОРА И РОПА"],
    [
        "ФЛАГМАН:\nРечевой ИИ-Супервайзер\n(Whisper + LLM)",
        "Анализирует 100% аудиозаписей звонков за вчерашний день по 13 жестким параметрам B2B-продаж.",
        "РОП за 5 минут видит, кто из менеджеров не берет дату встречи, где хамство или слив возражения «дорого»."
    ],
    [
        "ДЕТЕКТИВ ВЫРУЧКИ:\nСвязка звонков с CRM\n(AmoCRM / Битрикс24)",
        "Связывает дефекты звонков со сделками в CRM, выявляет брошенные лиды, просроченные КП (>48ч) и зависшую дебиторку.",
        "Показывает конкретные суммы под угрозой: «Сделка D-104 на 850 000 ₽ сгорает из-за отсутствия повторного контакта»."
    ],
    [
        "ШЕРИФ НА АУТСОРСЕ:\nЕжедневный Telegram-надзор\nи Пульт РОПа",
        "Каждое утро в 08:30 присылает директору сводку выполнения плана, а РОПу — список клиентов для срочного перехвата.",
        "Собственник держит руку на пульсе компании прямо с экрана смартфона без захода в тяжелые базы данных."
    ]
]

for r_idx, row in enumerate(row_data):
    for c_idx, val in enumerate(row):
        cell = tbl_comp.rows[r_idx].cells[c_idx]
        cell.width = col_widths[c_idx]
        if r_idx == 0:
            set_cell_background(cell, "F1F5F9")
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.name = "Arial"
            r.font.size = Pt(8.5)
            r.font.bold = True
            r.font.color.rgb = COLOR_PRIMARY
        else:
            set_cell_background(cell, "FFFFFF" if r_idx % 2 == 1 else "F8FAFC")
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.name = "Arial"
            r.font.size = Pt(9)
            r.font.color.rgb = COLOR_DARK
            if c_idx == 0:
                r.font.bold = True

p_div = doc.add_paragraph()
p_div.paragraph_format.space_before = Pt(8)
p_div.paragraph_format.space_after = Pt(8)

# ==================== SECTION 3 ====================
h1 = doc.add_heading(level=1)
h1.paragraph_format.space_before = Pt(14)
h1.paragraph_format.space_after = Pt(6)
r_h1 = h1.add_run("3. 13 критериев ИИ-оценки речи (Флагманский модуль)")
r_h1.font.name = "Georgia"
r_h1.font.size = Pt(14)
r_h1.font.bold = True
r_h1.font.color.rgb = COLOR_PRIMARY

doc.add_paragraph(
    "ИИ-модель обучена на отраслевых стандартах Enterprise B2B-продаж и проверяет каждый звонок по 13 контрольным точкам:"
)

tbl_speech = doc.add_table(rows=14, cols=3)
tbl_speech.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl_speech.autofit = False
set_table_borders(tbl_speech, color="E2E8F0", sz="4")

col_w_sp = [Inches(0.6), Inches(2.7), Inches(3.5)]
speech_params = [
    ["№", "КРИТЕРИЙ ОЦЕНКИ", "ЧТО ПРОВЕРЯЕТ ИИ В ДИАЛОГЕ"],
    ["1", "Жесткий Next Step (Критично)", "Зафиксирована ли точная дата и время следующего контакта или встречи"],
    ["2", "Инициатива диалога", "Кто задает вопросы: ведет ли менеджер диалог или только отвечает"],
    ["3", "Квалификация ЛПР", "Установлен ли статус собеседника: принимает ли он решения по бюджету"],
    ["4", "Выявление болей и сроков", "Понятно ли, какую задачу решает клиент и когда ему нужен результат"],
    ["5", "Отработка возражения «Дорого»", "Предложена ли альтернатива, рассрочка, кейс или менеджер сразу сдался"],
    ["6", "Отработка возражения «Я подумаю»", "Вскрыто ли скрытое сомнение или клиент просто отпущен"],
    ["7", "Презентация ценности через кейсы", "Звучали ли реальные примеры окупаемости и цифры внедрения"],
    ["8", "Попытка закрытия сделки", "Озвучено ли прямое предложение заключить договор или выставить счет"],
    ["9", "Соблюдение регламента приветствия", "Корректное представление компании и имени менеджера"],
    ["10", "Чистота речи и уверенность", "Отсутствие слов-паразитов, заминок, неуверенного мямленья цены"],
    ["11", "Слушание клиента (Talk/Listen Ratio)", "Менеджер не перебивает и слушает клиента не менее 50% времени"],
    ["12", "Фиксация договоренностей в конце", "Резюмирование итогов разговора («Договорились: я отправлю КП до 14:00»...)"],
    ["13", "Отказ от скидок без повода", "Менеджер не раздает скидки до того, как клиент попросил об этом"]
]

for r_idx, row in enumerate(speech_params):
    for c_idx, val in enumerate(row):
        cell = tbl_speech.rows[r_idx].cells[c_idx]
        cell.width = col_w_sp[c_idx]
        if r_idx == 0:
            set_cell_background(cell, "F1F5F9")
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.name = "Arial"
            r.font.size = Pt(8)
            r.font.bold = True
            r.font.color.rgb = COLOR_PRIMARY
        else:
            set_cell_background(cell, "FFFFFF" if r_idx % 2 == 1 else "F8FAFC")
            set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.name = "Arial"
            r.font.size = Pt(8.5)
            r.font.color.rgb = COLOR_DARK
            if c_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if r_idx == 1 and c_idx == 1:
                r.font.bold = True
                r.font.color.rgb = COLOR_ACCENT

# ==================== SECTION 4 ====================
h1 = doc.add_heading(level=1)
h1.paragraph_format.space_before = Pt(14)
h1.paragraph_format.space_after = Pt(6)
r_h1 = h1.add_run("4. Пошаговый 7-дневный таймлайн запуска (Быстрый старт)")
r_h1.font.name = "Georgia"
r_h1.font.size = Pt(14)
r_h1.font.bold = True
r_h1.font.color.rgb = COLOR_PRIMARY

doc.add_paragraph(
    "В отличие от классических IT-внедрений, которые длятся месяцами, пилот запускается за 7 рабочих дней без остановки работы отдела продаж:"
)

tbl_time = doc.add_table(rows=6, cols=3)
tbl_time.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl_time.autofit = False
set_table_borders(tbl_time, color="E2E8F0", sz="4")

col_w_t = [Inches(1.2), Inches(2.6), Inches(3.0)]
time_steps = [
    ["СРОК", "ДЕЙСТВИЯ REVOPS-ИНЖЕНЕРА", "РЕЗУЛЬТАТ ДЛЯ КЛИЕНТА"],
    ["День 1", "Подписание договора с самозанятым. Получение доступов к CRM и телефонии.", "Старт проекта. 0 часов отвлечения менеджеров от звонков."],
    ["День 2", "Подключение коннекторов. Выгрузка архива 100 последних звонков и базы сделок.", "Сформирован реестр сырых данных. Запущен ИИ-рентген речи."],
    ["День 3", "Детектив выручки: фильтрация зомби-сделок, поиск зависших оплат и срывов SLA.", "Выявлены скрытые резервы: найдены конкретные сделки под угрозой срыва."],
    ["День 4", "Презентация отчета директору: демонстрация карты потерь и среза звонков.", "Директор видит фактическую картину продаж без украшательств."],
    ["Дни 5–7", "Подключение Telegram-бота. Настройка утреннего Пульта РОПа (15 мин в день).", "Система сдана под ключ. Запущен ежедневный надзор и контроль."]
]

for r_idx, row in enumerate(time_steps):
    for c_idx, val in enumerate(row):
        cell = tbl_time.rows[r_idx].cells[c_idx]
        cell.width = col_w_t[c_idx]
        if r_idx == 0:
            set_cell_background(cell, "F1F5F9")
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.name = "Arial"
            r.font.size = Pt(8)
            r.font.bold = True
            r.font.color.rgb = COLOR_PRIMARY
        else:
            set_cell_background(cell, "FFFFFF" if r_idx % 2 == 1 else "F8FAFC")
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.name = "Arial"
            r.font.size = Pt(8.5)
            r.font.color.rgb = COLOR_DARK
            if c_idx == 0:
                r.font.bold = True

# ==================== SECTION 5 ====================
h1 = doc.add_heading(level=1)
h1.paragraph_format.space_before = Pt(14)
h1.paragraph_format.space_after = Pt(6)
r_h1 = h1.add_run("5. Финансовые условия и работа с самозанятым (НПД)")
r_h1.font.name = "Georgia"
r_h1.font.size = Pt(14)
r_h1.font.bold = True
r_h1.font.color.rgb = COLOR_PRIMARY

doc.add_paragraph(
    "Взаимодействие выстроено предельно открыто, без скрытых платежей и сложных корпоративных барьеров:"
)

# Table pricing
tbl_price = doc.add_table(rows=3, cols=3)
tbl_price.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl_price.autofit = False
set_table_borders(tbl_price, color="CBD5E1", sz="4")

col_w_pr = [Inches(2.0), Inches(1.8), Inches(3.0)]
price_rows = [
    ["ЭТАП", "СТОИМОСТЬ", "ЧТО ВХОДИТ И УСЛОВИЯ"],
    [
        "Этап 1: Быстрый старт\n(Диагностика и Детектив)",
        "29 000 ₽\n(разово)",
        "• Подключение телефонии и CRM\n• ИИ-аудит звонков по 13 критериям\n• Карта утечек и зависших сделок\n• Презентационный отчет для CEO (5 стр.)"
    ],
    [
        "Этап 2: Ежемесячный надзор\n(Шериф на аутсорсе)",
        "29 000 ₽\n(ежемесячно)",
        "• ИИ слушает 100% звонков каждый день\n• Утренний Telegram-пульс в 08:30\n• Пульт РОПа «15 минут в день»\n• Реактивация списанных сделок и дебиторки"
    ]
]

for r_idx, row in enumerate(price_rows):
    for c_idx, val in enumerate(row):
        cell = tbl_price.rows[r_idx].cells[c_idx]
        cell.width = col_w_pr[c_idx]
        if r_idx == 0:
            set_cell_background(cell, "F1F5F9")
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.name = "Arial"
            r.font.size = Pt(8.5)
            r.font.bold = True
            r.font.color.rgb = COLOR_PRIMARY
        else:
            set_cell_background(cell, "FFFFFF" if r_idx == 1 else "F8FAFC")
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.name = "Arial"
            r.font.size = Pt(9)
            r.font.color.rgb = COLOR_DARK
            if c_idx == 1:
                r.font.bold = True
                r.font.color.rgb = COLOR_ACCENT

add_callout(
    doc,
    "⚖️ Юридическая чистота и налоги для компании-клиента",
    "1. Оплата производится безналичным расчетом с расчетного счета юрлица/ИП клиента по договору оказания услуг с самозанятым.\n"
    "2. Клиент не платит за исполнителя НДФЛ (13%) и страховые взносы (30%) — налог 6% уплачивается исполнителем самостоятельно.\n"
    "3. Бухгалтерия клиента получает официальный фискальный чек из приложения «Мой налог» и подписанный Акт сдачи-приемки услуг, что позволяет на 100% уменьшить налог на прибыль или УСН (ст. 15 Федерального закона № 422-ФЗ).\n"
    "4. Конфиденциальность: подписывается двухсторонний NDA, персональные данные (ФИО, телефоны) обезличиваются (152-ФЗ РФ).",
    bg_hex=HEX_SUCCESS_LIGHT,
    border_hex=HEX_PRIMARY
)

# ==================== SECTION 6 ====================
h1 = doc.add_heading(level=1)
h1.paragraph_format.space_before = Pt(14)
h1.paragraph_format.space_after = Pt(6)
r_h1 = h1.add_run("6. Железобетонная гарантия первого пилота (Risk Reversal)")
r_h1.font.name = "Georgia"
r_h1.font.size = Pt(14)
r_h1.font.bold = True
r_h1.font.color.rgb = COLOR_PRIMARY

add_callout(
    doc,
    "🛡️ 100% Гарантия возврата средств на пилоте",
    "Если в течение первых 7 дней работы ИИ-супервайзер и модуль Детектива выручки не выявят конкретных точек слива лидов менеджерами и не найдут зависших сделок в CRM на сумму, минимум в 5 раз превышающую стоимость пилота (от 150 000 ₽) — исполнитель возвращает 100% оплаты за старт без лишних споров и согласований.",
    bg_hex="FFFBEB",
    border_hex="D97706"
)

# Signature block
p_sig = doc.add_paragraph()
p_sig.paragraph_format.space_before = Pt(20)
p_sig.paragraph_format.space_after = Pt(4)
r_sig = p_sig.add_run("СОГЛАСОВАНО И ПРИНЯТО К ИСПОЛНЕНИЮ:")
r_sig.font.name = "Arial"
r_sig.font.size = Pt(9.5)
r_sig.font.bold = True
r_sig.font.color.rgb = COLOR_PRIMARY

tbl_sig = doc.add_table(rows=1, cols=2)
tbl_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl_sig.autofit = False
set_table_borders(tbl_sig, color="CBD5E1", sz="4")

col_w_s = [Inches(3.4), Inches(3.4)]
for c_idx, title in enumerate(["ОТ ЗАКАЗЧИКА:", "ОТ ИСПОЛНИТЕЛЯ (REVOPS):"]):
    cell = tbl_sig.rows[0].cells[c_idx]
    cell.width = col_w_s[c_idx]
    set_cell_background(cell, "FFFFFF")
    set_cell_margins(cell, top=80, bottom=120, left=100, right=100)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(title + "\n")
    r.font.name = "Arial"
    r.font.size = Pt(8.5)
    r.font.bold = True
    r.font.color.rgb = COLOR_MUTED
    
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_after = Pt(0)
    r2 = p2.add_run("____________________ / ______________ /\nДата: «____» ________________ 2026 г.")
    r2.font.name = "Arial"
    r2.font.size = Pt(8.5)
    r2.font.color.rgb = COLOR_DARK

doc.save(OUT_DOCX)
print(f"Generated DOCX: {OUT_DOCX} ({os.path.getsize(OUT_DOCX):,} bytes)")

# Copy to Desktop
try:
    shutil.copyfile(OUT_DOCX, DESKTOP_DOCX)
    print(f"Copied to Desktop: {DESKTOP_DOCX}")
except Exception as e:
    print(f"Warning desktop copy: {e}")

# Copy to Brain
try:
    shutil.copyfile(OUT_DOCX, BRAIN_DOCX)
    print(f"Copied to Brain Artifacts: {BRAIN_DOCX}")
except Exception as e:
    print(f"Warning brain copy: {e}")
