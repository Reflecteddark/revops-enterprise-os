"""
Генератор Word-документа: "Холодный Outbound-Пак для Telegram: 3-Шаговая Воронка Привлечения ЛПР"
Создает боевое руководство по неспамному выходу на собственников и коммерческих директоров.
"""

import sys
from pathlib import Path
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

COLOR_NAVY = RGBColor(0x0F, 0x17, 0x2A)
COLOR_PRIMARY = RGBColor(0x1E, 0x3A, 0x8A)
COLOR_BLUE = RGBColor(0x25, 0x63, 0xEB)
COLOR_DARK = RGBColor(0x1E, 0x29, 0x3B)
COLOR_MUTED = RGBColor(0x64, 0x74, 0x8B)
COLOR_EMERALD = RGBColor(0x05, 0x96, 0x69)
COLOR_AMBER = RGBColor(0xD9, 0x77, 0x06)
COLOR_ROSE = RGBColor(0xE1, 0x1D, 0x48)

HEX_NAVY = "0F172A"
HEX_PRIMARY = "1E3A8A"
HEX_BLUE = "2563EB"
HEX_BLUE_LIGHT = "EFF6FF"
HEX_DARK = "1E293B"
HEX_LIGHT_BG = "F8FAFC"
HEX_BORDER = "CBD5E1"
HEX_EMERALD = "059669"
HEX_EMERALD_LIGHT = "ECFDF5"
HEX_AMBER = "D97706"
HEX_AMBER_LIGHT = "FFFBEB"
HEX_ROSE = "E11D48"
HEX_ROSE_LIGHT = "FFF1F2"


def set_cell_background(cell, hex_color: str):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shd)


def set_cell_margins(cell, top=140, bottom=140, left=180, right=180):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def set_table_borders(table, hex_color="CBD5E1"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="single" w:sz="6" w:space="0" w:color="{hex_color}"/>'
        f'  <w:left w:val="none"/>'
        f'  <w:bottom w:val="single" w:sz="8" w:space="0" w:color="{hex_color}"/>'
        f'  <w:right w:val="none"/>'
        f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{hex_color}"/>'
        f'  <w:insideV w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)


def add_callout(doc, text, title="ПРАВИЛО КАСАНИЯ", bg_color=HEX_BLUE_LIGHT, border_color=HEX_BLUE):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tbl.columns[0].width = Inches(6.8)

    cell = tbl.rows[0].cells[0]
    set_cell_background(cell, bg_color)
    set_cell_margins(cell, top=140, bottom=140, left=200, right=180)

    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'  <w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/>'
        f'  <w:top w:val="none"/>'
        f'  <w:bottom w:val="none"/>'
        f'  <w:right w:val="none"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)

    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    run_t = p.add_run(f"⚡ {title.upper()}\n")
    run_t.font.name = "Arial"
    run_t.font.size = Pt(9.5)
    run_t.font.bold = True
    run_t.font.color.rgb = COLOR_PRIMARY

    run_b = p.add_run(text)
    run_b.font.name = "Arial"
    run_b.font.size = Pt(10)
    run_b.font.color.rgb = COLOR_DARK

    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def build_outbound_pack_docx(output_path: str):
    doc = docx.Document()

    for section in doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

        footer = section.footer
        f_p = footer.paragraphs[0]
        f_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        f_run = f_p.add_run("RevOps Platform | Холодный Outbound-Пак для Telegram (B2B Lead Gen)")
        f_run.font.name = "Arial"
        f_run.font.size = Pt(8)
        f_run.font.color.rgb = COLOR_MUTED

    # Шапка
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(2)

    r_badge = p_title.add_run("ЛИДОГЕНЕРАЦИЯ БЕЗ БЮДЖЕТА • ПЕРВЫЕ 20 КЛИЕНТОВ НА ТЕСТ-ДРАЙВ\n")
    r_badge.font.name = "Arial"
    r_badge.font.size = Pt(9.5)
    r_badge.font.bold = True
    r_badge.font.color.rgb = COLOR_BLUE

    r_t = p_title.add_run("Холодный Outbound-Пак для Telegram\n")
    r_t.font.name = "Arial"
    r_t.font.size = Pt(20)
    r_t.font.bold = True
    r_t.font.color.rgb = COLOR_NAVY

    r_sub = p_title.add_run("3-шаговая цепочка сообщений для ЛПР, механика поиска контактов и триггеры через открытые вакансии HH.ru")
    r_sub.font.name = "Arial"
    r_sub.font.size = Pt(11)
    r_sub.font.color.rgb = COLOR_MUTED

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    add_callout(
        doc,
        "Почему этот метод дает конверсию в ответ >25%:\n"
        "Мы НЕ пишем спам 'Предлагаем услуги речевой аналитики'. Мы пишем в контексте актуальной проблемы компании: "
        "они прямо сейчас тратят деньги на поиск РОПа или менеджеров на HH.ru. Наш оффер — готовое решение их проблемы прямо сейчас, "
        "пока вакансия еще открыта.",
        title="ЗОЛОТОЙ ПРИНЦИП КОНТЕКСТНОГО АУТРИЧА"
    )

    # 1. Поиск ЛПР за 2 минуты
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(10)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("1. Как найти контакт ЛПР за 2 минуты (Без платных баз)")
    r.font.name = "Arial"
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = COLOR_PRIMARY

    steps_find = [
        ("Шаг 1: Берем название компании из таблицы", "В файле 'Горячие_Лиды_HH_RevOps.xlsx' копируем название компании или ИНН (указан на странице работодателя на hh.ru)."),
        ("Шаг 2: Пробиваем в Rusprofile / Чекко", "Вбиваем ИНН на Rusprofile.ru. Смотрим ФИО Генерального директора или Учредителя. Это наш ЛПР №1."),
        ("Шаг 3: Поиск в Telegram", "В поисковой строке Telegram вбиваем: 'Имя Фамилия' или номер телефона из открытых реестров компании. В 70% случаев аккаунт директора находится за 30 секунд."),
        ("Шаг 4: Альтернатива — контактное лицо в вакансии", "В 30% вакансий на HH прямо указан контакт рекрутера или HRD (Telegram/WhatsApp). Пишем им: просим связать с коммерческим директором.")
    ]

    for s_title, s_desc in steps_find:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r_b = p.add_run(f"• {s_title}: ")
        r_b.font.name = "Arial"
        r_b.font.bold = True
        r_b.font.color.rgb = COLOR_BLUE
        r_d = p.add_run(s_desc)
        r_d.font.name = "Arial"
        r_d.font.color.rgb = COLOR_DARK

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 2. 3-шаговая цепочка сообщений
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(10)
    h2.paragraph_format.space_after = Pt(4)
    r = h2.add_run("2. Боевая 3-шаговая цепочка сообщений в Telegram")
    r.font.name = "Arial"
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = COLOR_PRIMARY

    # Касание 1
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    p_k1 = doc.add_paragraph()
    r = p_k1.add_run("КАСАНИЕ 1 (ДЕНЬ 1): МЯГКИЙ ВХОД ЧЕРЕЗ КОНТЕКСТ ВАКАНСИИ")
    r.font.name = "Arial"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = COLOR_BLUE

    add_callout(
        doc,
        "«[Имя], добрый день!\n\n"
        "Обратил внимание, что вы сейчас в [Название компании] открыли вакансию [Название вакансии, например: Руководителя отдела продаж].\n\n"
        "Пока идет найм сильного кандидата (а это обычно от 3 до 6 недель), в отделе продаж часто проседает контроль: менеджеры не фиксируют дедлайны встреч, отпускают клиентов 'подумать' и сливают рекламный бюджет.\n\n"
        "Мы развернули сервис ai-rop.ru: нейросеть берет на себя 100% прослушки звонков менеджеров, оценивает диалоги по 13 критериям B2B и каждое утро присылает вам готовую выжимку с зонами сливов сделок.\n\n"
        "Готовы бесплатно разобрать 3 любых реальных звонка ваших менеджеров за вчера, чтобы показать, где прямо сейчас теряется выручка. Интересно взглянуть на пример такого отчета?»",
        title="ШАБЛОН СООБЩЕНИЯ 1 (КОПИРОВАТЬ В TELEGRAM)",
        bg_color=HEX_BLUE_LIGHT,
        border_color=HEX_BLUE
    )

    # Касание 2
    p_k2 = doc.add_paragraph()
    p_k2.paragraph_format.space_before = Pt(8)
    p_k2.paragraph_format.space_after = Pt(2)
    r = p_k2.add_run("КАСАНИЕ 2 (ЧЕРЕЗ 48 ЧАСОВ, ЕСЛИ МОЛЧИТ): ИНСАЙТ И ЦИФРЫ ПОТЕРЬ")
    r.font.name = "Arial"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = COLOR_AMBER

    add_callout(
        doc,
        "«[Имя], приветствую снова!\n\n"
        "Понимаю вашу занятость. Просто для наглядности: недавно делали аудит похожей B2B-компании. "
        "Выяснили, что менеджеры в 42% звонков не задавали вопрос: 'Кто кроме вас участвует в согласовании договора?', "
        "тратили по 40 минут на общение со специалистами без полномочий, а настоящие ЛПР уходили к конкурентам. Сумма сорванных сделок за месяц — 1.8 млн ₽.\n\n"
        "Наш аудит 3 звонков делается за 24 часа без доработок CRM и серверов. Вы увидите реальную картину в диалогах вашей команды.\n\n"
        "Скинуть короткое видео (1 минута) или ссылку на Google Таблицу с примером разбора?»",
        title="ШАБЛОН СООБЩЕНИЯ 2 (ЧЕРЕЗ 2 ДНЯ)",
        bg_color=HEX_AMBER_LIGHT,
        border_color=HEX_AMBER
    )

    # Касание 3
    p_k3 = doc.add_paragraph()
    p_k3.paragraph_format.space_before = Pt(8)
    p_k3.paragraph_format.space_after = Pt(2)
    r = p_k3.add_run("КАСАНИЕ 3 (ЧЕРЕЗ 4 ДНЯ): ОФФЕР С МЯГКИМ ДЕДЛАЙНОМ")
    r.font.name = "Arial"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = COLOR_EMERALD

    add_callout(
        doc,
        "«[Имя], добрый день!\n\n"
        "На этой неделе берем ровно 3 компании на бесплатный пилотный аудит (разбор 3 звонков по 13 критериям + расчет упущенной выгоды). "
        "Остался 1 слот до пятницы.\n\n"
        "Если актуально посмотреть, где менеджеры сейчас недожимают клиентов — пришлите 3 аудиозаписи прямо в этот чат или скажите, куда направить ссылку на защищенную папку.\n\n"
        "Если сейчас не в фокусе — дайте знать, не буду отвлекать!»",
        title="ШАБЛОН СООБЩЕНИЯ 3 (МЯГКИЙ ДЕДЛАЙН)",
        bg_color=HEX_EMERALD_LIGHT,
        border_color=HEX_EMERALD
    )

    # Касание 4 - Break-up
    p_k4 = doc.add_paragraph()
    p_k4.paragraph_format.space_before = Pt(8)
    p_k4.paragraph_format.space_after = Pt(2)
    r = p_k4.add_run("КАСАНИЕ 4 (ЧЕРЕЗ 7 ДНЕЙ): ВЕЖЛИВЫЙ BREAK-UP (ЗАКРЫТИЕ ДИАЛОГА)")
    r.font.name = "Arial"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = COLOR_ROSE

    add_callout(
        doc,
        "«[Имя], добрый день! Вижу, что вопрос контроля звонков сейчас, скорее всего, не в приоритете. "
        "Больше писать не буду, чтобы не засорять переписку.\n\n"
        "Если в будущем вернетесь к задаче увеличения конверсии ОП и снятия рутины с РОПа — сохраните контакт или загляните на наш сайт https://ai-rop.ru.\n\n"
        "Успешного закрытия вакансии и отличных продаж!»",
        title="ШАБЛОН СООБЩЕНИЯ 4 (ВЕЖЛИВЫЙ ОТХОД)",
        bg_color=HEX_ROSE_LIGHT,
        border_color=HEX_ROSE
    )

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 3. Воронка цифр
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(10)
    h3.paragraph_format.space_after = Pt(4)
    r = h3.add_run("3. Экономика и воронка холодного аутрича")
    r.font.name = "Arial"
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = COLOR_PRIMARY

    funnel_table = doc.add_table(rows=5, cols=4)
    funnel_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    funnel_table.autofit = False

    col_w = [Inches(1.5), Inches(1.8), Inches(1.8), Inches(1.7)]
    headers = ["Этап воронки", "Конверсия", "Количество (на 50 контактов)", "Результат"]

    for j, h_text in enumerate(headers):
        cell = funnel_table.rows[0].cells[j]
        cell.width = col_w[j]
        set_cell_background(cell, HEX_PRIMARY)
        set_cell_margins(cell, 100, 100, 120, 120)
        p = cell.paragraphs[0]
        r = p.add_run(h_text)
        r.font.name = "Arial"
        r.font.size = Pt(8.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    funnel_data = [
        ("Отправка первого сообщения", "100%", "50 ЛПР", "Компании из базы HH"),
        ("Ответ / вступление в диалог", "25 – 35%", "12 – 17 ответов", "«Давайте посмотрим / пришлите пример»"),
        ("Согласие на тест 3 звонков", "50% от ответивших", "6 – 8 компаний", "Прислали 3 файла аудиозаписей"),
        ("Закрытие на платный пилот", "40 – 50% от аудитов", "3 – 4 сделки", "Чеки 45 000 – 150 000 ₽ (Выручка: 180к–450к ₽)")
    ]

    for row_idx, data in enumerate(funnel_data, start=1):
        row = funnel_table.rows[row_idx]
        bg = HEX_LIGHT_BG if row_idx % 2 == 1 else "FFFFFF"
        for col_idx, text in enumerate(data):
            cell = row.cells[col_idx]
            cell.width = col_w[col_idx]
            set_cell_background(cell, bg)
            set_cell_margins(cell, 90, 90, 110, 110)
            p = cell.paragraphs[0]
            r = p.add_run(text)
            r.font.name = "Arial"
            r.font.size = Pt(8.5)
            r.font.color.rgb = COLOR_DARK
            if col_idx == 3:
                r.font.bold = True
                r.font.color.rgb = COLOR_EMERALD

    set_table_borders(funnel_table)
    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    print(f"Outbound pack docx saved at: {output_path}")


if __name__ == "__main__":
    dest1 = r"C:\Users\strel\Desktop\RevOps Platform\Каналы продаж\Холодный_Outbound_Пак_Telegram.docx"
    dest2 = r"C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\docs\Холодный_Outbound_Пак_Telegram.docx"
    build_outbound_pack_docx(dest1)
    build_outbound_pack_docx(dest2)
