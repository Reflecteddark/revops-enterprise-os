"""
Генератор Word-документа "152-ФЗ Комплект Юридической Безопасности: Речевая ИИ-Аналитика в Отделе Продаж"
RevOps Enterprise OS V17.6

Включает полный комплект документов под ключ:
1. Памятка для генерального директора и службы безопасности: 4 шага легального запуска;
2. Согласие сотрудника на обработку биометрических персональных данных (голоса);
3. Регламент и формулировки голосовых дисклеймеров для телефонии (АТС);
4. Положение о порядке использования ИИ-аналитики и защиты данных клиентов;
5. Типовой Акт об уничтожении записей и транскриптов (ротация 30 дней по нормам РКН).
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
COLOR_NAVY = RGBColor(0x0F, 0x17, 0x2A)       # #0F172A Глубокий синий
COLOR_PRIMARY = RGBColor(0x1E, 0x3A, 0x5F)    # #1E3A5F Заголовки
COLOR_INDIGO = RGBColor(0x4F, 0x46, 0xE5)     # #4F46E5 Акцент
COLOR_DARK = RGBColor(0x1E, 0x29, 0x3B)       # #1E293B Основной текст
COLOR_MUTED = RGBColor(0x64, 0x74, 0x8B)      # #64748B Пояснения
COLOR_SUCCESS = RGBColor(0x05, 0x96, 0x69)    # #059669
COLOR_WARNING = RGBColor(0xDC, 0x26, 0x26)    # #DC2626

HEX_PRIMARY = "1E3A5F"
HEX_INDIGO = "4F46E5"
HEX_INDIGO_LIGHT = "EEF2FF"
HEX_ZEBRA = "F8FAFC"
HEX_GREEN = "059669"
HEX_GREEN_LIGHT = "ECFDF5"
HEX_RED = "DC2626"
HEX_RED_LIGHT = "FEF2F2"
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

    # Rows
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
    cell.width = Inches(6.8)
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
    p_spacer.paragraph_format.space_after = Pt(3)


def generate_152fz_compliance_kit():
    doc = docx.Document()

    # Поля страницы 0.75" (~1.9 см)
    for section in doc.sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Calibri'
    style_normal.font.size = Pt(10)
    style_normal.font.color.rgb = COLOR_DARK
    style_normal.paragraph_format.space_after = Pt(4)
    style_normal.paragraph_format.line_spacing = 1.15

    for i in range(1, 4):
        h_style = doc.styles[f'Heading {i}']
        h_style.font.name = 'Calibri'
        h_style.font.color.rgb = COLOR_PRIMARY
        h_style.font.bold = True
        if i == 1:
            h_style.font.size = Pt(16)
            h_style.paragraph_format.space_before = Pt(12)
            h_style.paragraph_format.space_after = Pt(4)
        elif i == 2:
            h_style.font.size = Pt(13)
            h_style.paragraph_format.space_before = Pt(8)
            h_style.paragraph_format.space_after = Pt(3)
        else:
            h_style.font.size = Pt(11)
            h_style.paragraph_format.space_before = Pt(6)
            h_style.paragraph_format.space_after = Pt(2)

    # ==================== ТИТУЛ ====================
    for _ in range(1):
        doc.add_paragraph()

    p_badge = doc.add_paragraph()
    p_badge.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_badge = p_badge.add_run("ЮРИДИЧЕСКИЙ КОМПЛЕКТ БЕЗОПАСНОСТИ • СООТВЕТСТВИЕ 152-ФЗ РФ")
    r_badge.font.size = Pt(9.5)
    r_badge.font.bold = True
    r_badge.font.color.rgb = COLOR_INDIGO

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(2)
    p_title.paragraph_format.space_after = Pt(6)
    r_title = p_title.add_run("Правовой Регламент Внедрения\nРечевой ИИ-Аналитики в Отделе Продаж")
    r_title.font.size = Pt(20)
    r_title.font.bold = True
    r_title.font.color.rgb = COLOR_NAVY

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run(
        "Полный пакет юридических шаблонов, согласий на обработку биометрии (голоса),\n"
        "голосовых дисклеймеров для АТС и регламентов уничтожения данных по нормам Роскомнадзора."
    )
    r_sub.font.size = Pt(10)
    r_sub.font.italic = True
    r_sub.font.color.rgb = COLOR_MUTED

    add_callout(doc, "⚖️", "Правовой статус обработки голоса в 2026 году",
        "В соответствии со ст. 11 Федерального закона № 152-ФЗ «О персональных данных», аудиозапись голоса "
        "физического лица относится к категории биометрических персональных данных. Использование ИИ для транскрибации "
        "и оценки звонков требует строгого соблюдения двух условий: 1) письменного согласия сотрудников компании; "
        "2) надлежащего уведомления внешних клиентов до начала фиксации звука. Данный комплект полностью закрывает риски штрафов Роскомнадзора.",
        bg_hex=HEX_INDIGO_LIGHT, border_hex=HEX_INDIGO
    )

    doc.add_page_break()

    # ==================== РАЗДЕЛ 1: ПАМЯТКА РУКОВОДИТЕЛЮ ====================
    doc.add_heading("1. Памятка для Руководителя: 4 шага легализации ИИ-контроля", level=1)

    steps_table = [
        ["Шаг 1. Согласие сотрудников (ст. 9, 11 152-ФЗ)", "Каждый менеджер подписывает отдельное письменное Согласие на обработку биометрии (голоса) при приеме на работу или в виде Дополнения к трудовому договору."],
        ["Шаг 2. Автоинформатор АТС (ст. 6 152-ФЗ)", "В телефонию загружается голосовой дисклеймер. Продолжение разговора клиентом после предупреждения является конклюдентным согласием."],
        ["Шаг 3. Внутренний регламент компании", "Приказом генерального директора утверждается Положение об оценке качества звонков с помощью ИИ. Запрещается использовать аудио вне служебных целей."],
        ["Шаг 4. Ротация и уничтожение (ст. 21 152-ФЗ)", "Аудиозаписи хранятся в защищенном контуре не более 30 дней, после чего автоматически уничтожаются с формированием электронного лога/Акта."]
    ]
    add_styled_table(doc, ["ЭТАП ЛЕГАЛИЗАЦИИ", "ОБЯЗАТЕЛЬНЫЕ ДЕЙСТВИЯ ОРГАНИЗАЦИИ"], steps_table, col_widths=[2.5, 4.4], header_bg=HEX_PRIMARY)

    # ==================== РАЗДЕЛ 2: ШАБЛОН СОГЛАСИЯ СОТРУДНИКА ====================
    doc.add_heading("2. Шаблон документа: Согласие работника на обработку биометрии (голоса)", level=1)
    doc.add_paragraph("Ниже приведен юридически выверенный текст согласия для подписания менеджерами отдела продаж:")

    p_doc1_box = doc.add_paragraph()
    p_doc1_box.paragraph_format.left_indent = Cm(0.5)
    p_doc1_box.paragraph_format.right_indent = Cm(0.5)
    pPr = p_doc1_box._p.get_or_add_pPr()
    borders = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="8" w:space="4" w:color="{HEX_PRIMARY}"/>'
        f'<w:left w:val="single" w:sz="16" w:space="4" w:color="{HEX_PRIMARY}"/>'
        f'<w:bottom w:val="single" w:sz="8" w:space="4" w:color="{HEX_PRIMARY}"/>'
        f'<w:right w:val="single" w:sz="8" w:space="4" w:color="{HEX_PRIMARY}"/>'
        f'</w:pBdr>'
    )
    pPr.append(borders)

    doc1_content = """СОГЛАСИЕ РАБОТНИКА
на обработку персональных данных и биометрических персональных данных (голоса)

г. ___________________                                            «____» ________________ 202__ г.

Я, ____________________________________________________________________________________,
паспорт: серия _______ № ______________, выдан _________________________________________
______________________________________________________, код подразделения ____________,
зарегистрированный(ая) по адресу: _____________________________________________________,
являясь работником __________________________________________________ (далее — «Работодатель»),

в соответствии со ст. 9 и ст. 11 Федерального закона от 27.07.2006 № 152-ФЗ «О персональных данных», 
свободно, своей волей и в своем интересе даю конкретное, предметное, информированное, 
сознательное и однозначное согласие Работодателю на обработку моих персональных данных:

1. Перечень обрабатываемых данных:
• Фамилия, имя, отчество, должность;
• Биометрические персональные данные: запись голоса и акустические характеристики речи, 
  зафиксированные средствами корпоративной телефонной связи при исполнении трудовых обязанностей;
• Текстовые расшифровки (транскрипты) телефонных переговоров с контрагентами и клиентами Работодателя.

2. Цели обработки:
• Контроль соблюдения стандартов клиентского сервиса, сценариев и регламентов B2B-продаж;
• Оценка качества ведения переговоров с использованием автоматизированных систем речевой аналитики;
• Обучение, повышение квалификации и объективный расчет показателей премирования (KPI);
• Урегулирование спорных ситуаций с клиентами и защита законных интересов Работодателя.

3. Способы обработки:
Сбор, запись, систематизация, накопление, хранение, распознавание с использованием технологий 
искусственного интеллекта (STT/LLM в изолированном контуре), обезличивание и удаление. 
Передача третьим лицам запрещена, за исключением уполномоченных технических операторов платформы 
RevOps на основании договора с условием строгой конфиденциальности.

4. Срок действия и уничтожение:
Согласие действует в течение всего срока действия трудового договора. Аудиозаписи переговоров 
хранятся не более 30 (тридцати) календарных дней с момента записи, после чего подлежат уничтожению.

Подпись работника: __________________ / ___________________________________ /"""

    r_d1 = p_doc1_box.add_run(doc1_content)
    r_d1.font.name = "Consolas"
    r_d1.font.size = Pt(8.5)
    r_d1.font.color.rgb = COLOR_DARK

    doc.add_page_break()

    # ==================== РАЗДЕЛ 3: РЕГЛАМЕНТ АТС И ДИСКЛЕЙМЕРЫ ====================
    doc.add_heading("3. Шаблоны голосовых уведомлений для АТС (Работа с клиентами)", level=1)

    p_disc = doc.add_paragraph(
        "Для внешних клиентов (ЛПР, контрагентов) согласие на запись звонка признается конклюдентным, "
        "если компания явно предупредила абонента до начала фиксации разговора. "
        "Ниже приведены утвержденные отраслевые формулировки для виртуальных АТС:"
    )

    add_styled_table(doc,
        ["КАНАЛ СВЯЗИ", "ГОЛОСОВОЙ МОДУЛЬ (ДИСКЛЕЙМЕР)", "ПРАВОВОЙ КОММЕНТАРИЙ"],
        [
            ["Входящие вызовы (IVR / автоответчик)",
             "«Здравствуйте! Вы позвонили в компанию [Название]. В целях контроля качества обслуживания и обучения персонала все разговоры записываются. Соединяю со специалистом...»",
             "Включается до ответа менеджера. Факт того, что клиент остался на линии — 100% подтверждает согласие."],
            ["Исходящие звонки менеджеров (скрипт)",
             "«Добрый день, [Имя Клиента]! Меня зовут [Имя], компания [Название]. Наш разговор записывается для контроля договоренностей. Подскажите, вам удобно сейчас обсудить...»",
             "Обязательно озвучивать на первых 10 секундах звонка. Вшивается в 1-й пункт скрипта менеджера."],
            ["Оферта на сайте / политика конфиденциальности",
             "«Компания осуществляет запись телефонных соединений и их автоматизированную аналитику исключительно в целях контроля сервиса и подтверждения условий сделок. Срок хранения — 30 дней.»",
             "Добавляется отдельным пунктом в Политику обработки персональных данных на веб-сайте."]
        ],
        col_widths=[1.8, 3.3, 1.8],
        header_bg=HEX_INDIGO
    )

    # ==================== РАЗДЕЛ 4: РЕГЛАМЕНТ ХРАНЕНИЯ И ЗАЩИТЫ ДАННЫХ ====================
    doc.add_heading("4. Положение о защите данных при использовании ИИ-аналитики", level=1)

    sec_rules = [
        ("1. Изоляция финансового контура:",
         "Система RevOps Enterprise OS разделяет голосовые данные и коммерческие параметры сделок. В сторонние нейросетевые модели никогда не передаются номера расчетных счетов, паспортные данные или банковские реквизиты."),
        ("2. Автоматическое обезличивание (Tokenization):",
         "При транскрибации речи модулем Whisper номера телефонов и фамилии заменяются на системные идентификаторы [КЛИЕНТ_ID]. ИИ анализирует контекст диалога, а не персональные профили."),
        ("3. Хранение исключительно в РФ:",
         "Все серверные мощности и базы данных расположены в дата-центрах на территории Российской Федерации (ст. 18 152-ФЗ). Запрещен экспорт сырого аудиопотока за пределы юрисдикции РФ."),
        ("4. Журнал аудита операций (audit_log.csv):",
         "Каждое обращение к записям звонков и отчетам фиксируется в неизменяемом логе с указанием роли (Администратор, РОП, Аналитик), даты и IP-адреса с автоматической ротацией до 10 000 записей.")
    ]

    for title, desc in sec_rules:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        r_b = p.add_run(title + " ")
        r_b.font.bold = True
        r_b.font.color.rgb = COLOR_PRIMARY
        r_t = p.add_run(desc)

    # ==================== РАЗДЕЛ 5: ТИПОВОЙ АКТ ОБ УНИЧТОЖЕНИИ ДАННЫХ ====================
    doc.add_heading("5. Типовой Акт об уничтожении персональных данных (Ротация 30 дней)", level=1)
    doc.add_paragraph("Документ подтверждает выполнение требований Роскомнадзора об удалении данных после достижения целей обработки:")

    p_doc2_box = doc.add_paragraph()
    p_doc2_box.paragraph_format.left_indent = Cm(0.5)
    p_doc2_box.paragraph_format.right_indent = Cm(0.5)
    pPr2 = p_doc2_box._p.get_or_add_pPr()
    borders2 = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="8" w:space="4" w:color="{HEX_PRIMARY}"/>'
        f'<w:left w:val="single" w:sz="16" w:space="4" w:color="{HEX_PRIMARY}"/>'
        f'<w:bottom w:val="single" w:sz="8" w:space="4" w:color="{HEX_PRIMARY}"/>'
        f'<w:right w:val="single" w:sz="8" w:space="4" w:color="{HEX_PRIMARY}"/>'
        f'</w:pBdr>'
    )
    pPr2.append(borders2)

    doc2_content = """УТВЕРЖДАЮ:
Генеральный директор __________________________
____________________ / ________________________ /
«____» ________________ 202__ г.

АКТ
об уничтожении персональных данных (аудиозаписей звонков и транскриптов)

г. ___________________                                            «____» ________________ 202__ г.

Комиссия в составе:
Председатель: ________________________________________________ (должность, ФИО)
Члены комиссии: ______________________________________________ (должность, ФИО)
                ______________________________________________ (должность, ФИО)

составила настоящий Акт о том, что в соответствии с Положением о защите персональных данных 
и регламентом 30-дневного срока хранения информации произведено плановое уничтожение 
массива данных телефонных переговоров за период с «____» __________ по «____» __________ 202__ г.

1. Перечень уничтоженных данных:
• Аудиофайлы входящих и исходящих вызовов отдела продаж в количестве: ________ шт.
• Временные текстовые транскрипты и рабочие аудио-логи платформы.

2. Способ уничтожения:
Программное безвозвратное удаление с серверов хранения данных методом перезаписи секторов 
без возможности восстановления.

3. Результат:
Персональные данные абонентов и биометрические образцы голоса за указанный период уничтожены в полном объеме.

Подписи членов комиссии:
____________________ / ________________________ /
____________________ / ________________________ /"""

    r_d2 = p_doc2_box.add_run(doc2_content)
    r_d2.font.name = "Consolas"
    r_d2.font.size = Pt(8.5)
    r_d2.font.color.rgb = COLOR_DARK

    doc.add_paragraph()
    p_end = doc.add_paragraph()
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_e = p_end.add_run("Комплект документов подготовлен в соответствии со ст. 6, 9, 11, 18, 21 Федерального закона № 152-ФЗ РФ\n")
    r_e.font.size = Pt(8.5)
    r_e.font.italic = True
    r_e.font.color.rgb = COLOR_MUTED
    r_e2 = p_end.add_run("RevOps Enterprise OS V17.6 Compliance Suite • Версия: 2026-10-01")
    r_e2.font.size = Pt(8.5)
    r_e2.font.italic = True
    r_e2.font.color.rgb = COLOR_MUTED

    # Сохранение файлов
    output_filename = "152-ФЗ_Комплект_Безопасности_RevOps_AI.docx"
    docs_path = Path("docs") / output_filename
    desktop_path = Path("C:/Users/strel/Desktop") / output_filename
    brain_path = Path("C:/Users/strel/.gemini/antigravity/brain/b958a22b-49ff-459c-a191-6c697ec14334") / output_filename

    docs_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(docs_path)
    print(f"✅ 152-ФЗ Compliance Kit успешно сохранен в docs: {docs_path.absolute()}")
    print(f"📄 Размер: {docs_path.stat().st_size / 1024:.1f} KB")

    try:
        shutil.copy2(docs_path, desktop_path)
        print(f"📋 Скопировано на Рабочий стол: {desktop_path}")
    except Exception as e:
        print(f"⚠️ Ошибка копирования на Рабочий стол: {e}")

    try:
        shutil.copy2(docs_path, brain_path)
        print(f"🧠 Скопировано в Brain Artifacts: {brain_path}")
    except Exception as e:
        print(f"⚠️ Ошибка копирования в Artifacts: {e}")

    return docs_path

if __name__ == "__main__":
    generate_152fz_compliance_kit()
