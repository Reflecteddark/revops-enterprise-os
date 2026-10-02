import os
import sys
from pathlib import Path
import docx
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

# === ЦВЕТОВАЯ ПАЛИТРА ENTERPRISE ===
COLOR_NAVY = RGBColor(0x0F, 0x17, 0x2A)        # #0F172A
COLOR_PRIMARY = RGBColor(0x1E, 0x3A, 0x8A)     # #1E3A8A Blue 900
COLOR_BLUE = RGBColor(0x25, 0x63, 0xEB)        # #2563EB Blue 600
COLOR_DARK = RGBColor(0x1E, 0x29, 0x3B)        # #1E293B Slate 800
COLOR_MUTED = RGBColor(0x64, 0x74, 0x8B)       # #64748B Slate 500
COLOR_EMERALD = RGBColor(0x05, 0x96, 0x69)     # #059669 Emerald 600
COLOR_AMBER = RGBColor(0xD9, 0x77, 0x06)       # #D97706 Amber 600
COLOR_ROSE = RGBColor(0xE1, 0x1D, 0x48)        # #E11D48 Rose 600

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


def add_callout(doc, text, title="КЛЮЧЕВОЙ АКЦЕНТ ДЛЯ СОБСТВЕННИКА", bg_color=HEX_BLUE_LIGHT, border_color=HEX_BLUE):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tbl.columns[0].width = Inches(6.8)
    
    cell = tbl.rows[0].cells[0]
    set_cell_background(cell, bg_color)
    set_cell_margins(cell, top=160, bottom=160, left=240, right=200)
    
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
    run_b.font.italic = False
    run_b.font.color.rgb = COLOR_DARK
    
    doc.add_paragraph().paragraph_format.space_after = Pt(6)


def build_presentation_docx(output_path: str):
    doc = docx.Document()
    
    # Page setup - A4
    for section in doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
        
        # Header / Footer
        footer = section.footer
        f_p = footer.paragraphs[0]
        f_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        f_run = f_p.add_run("RevOps Platform | Сценарий 20-минутной презентации (B2B High-Ticket)")
        f_run.font.name = "Arial"
        f_run.font.size = Pt(8)
        f_run.font.color.rgb = COLOR_MUTED

    # ==========================
    # ТИТУЛЬНАЯ ШАПКА
    # ==========================
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(4)
    title_p.paragraph_format.space_after = Pt(2)
    run_badge = title_p.add_run("БОЕВОЙ РЕГЛАМЕНТ ПЕРЕГОВОРОВ • ЗАКРЫТИЕ СДЕЛОК 150K–450K ₽\n")
    run_badge.font.name = "Arial"
    run_badge.font.size = Pt(9.5)
    run_badge.font.bold = True
    run_badge.font.color.rgb = COLOR_BLUE
    
    run_title = title_p.add_run("Сценарий 20-минутной Презентации «ИИ-РОП RevOps»\n")
    run_title.font.name = "Arial"
    run_title.font.size = Pt(20)
    run_title.font.bold = True
    run_title.font.color.rgb = COLOR_NAVY
    
    run_sub = title_p.add_run("Пошаговый путеводитель для спикера: экраны, дословные скрипты, вопросы-крючки и механика дожима")
    run_sub.font.name = "Arial"
    run_sub.font.size = Pt(11)
    run_sub.font.color.rgb = COLOR_MUTED
    
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    
    # Meta box
    meta_tbl = doc.add_table(rows=1, cols=4)
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_tbl.autofit = False
    widths = [Inches(1.7), Inches(1.7), Inches(1.7), Inches(1.7)]
    
    labels = [
        ("ХРОНОМЕТРАЖ", "Ровно 20 минут", COLOR_BLUE),
        ("ЦЕЛЕВАЯ АУДИТОРИЯ", "Собственник / Комдир", COLOR_DARK),
        ("ГЛАВНАЯ ЦЕЛЬ", "Согласие на тест 3 звонков", COLOR_EMERALD),
        ("ЧЕК СДЕЛКИ", "45 000 – 250 000 ₽", COLOR_AMBER)
    ]
    
    for i, (title, val, col) in enumerate(labels):
        cell = meta_tbl.rows[0].cells[i]
        cell.width = widths[i]
        set_cell_background(cell, HEX_LIGHT_BG)
        set_cell_margins(cell, 80, 80, 100, 100)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        
        r1 = p.add_run(f"{title}\n")
        r1.font.name = "Arial"
        r1.font.size = Pt(7.5)
        r1.font.bold = True
        r1.font.color.rgb = COLOR_MUTED
        
        r2 = p.add_run(val)
        r2.font.name = "Arial"
        r2.font.size = Pt(9.5)
        r2.font.bold = True
        r2.font.color.rgb = col
        
    set_table_borders(meta_tbl)
    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Вводный акцент
    add_callout(
        doc,
        "Главный закон 20-минутной встречи: 80% времени говорить на языке ДЕНЕГ клиента, а не технических нейросетей. "
        "Собственнику плевать, какая у вас модель (DeepSeek или Whisper). Ему важно одно: «Сколько миллионов я прямо сейчас недополучаю "
        "из-за косяков менеджеров, и как вы вернете мне эти деньги уже на следующей неделе».",
        title="ЗОЛОТОЕ ПРАВИЛО СПИКЕРА"
    )

    # ==========================
    # ТАЙМИНГ-КАРТА ВСТРЕЧИ
    # ==========================
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(12)
    h1.paragraph_format.space_after = Pt(6)
    r_h1 = h1.add_run("1. Тайминг-карта 20-минутной встречи")
    r_h1.font.name = "Arial"
    r_h1.font.size = Pt(14)
    r_h1.font.bold = True
    r_h1.font.color.rgb = COLOR_PRIMARY

    time_table = doc.add_table(rows=6, cols=4)
    time_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    time_table.autofit = False
    
    col_w = [Inches(1.2), Inches(1.8), Inches(2.2), Inches(1.6)]
    headers = ["Минуты", "Этап встречи", "Что показывать на экране", "Целевой результат"]
    
    hdr_row = time_table.rows[0]
    for j, h_text in enumerate(headers):
        cell = hdr_row.cells[j]
        cell.width = col_w[j]
        set_cell_background(cell, HEX_PRIMARY)
        set_cell_margins(cell, 100, 100, 120, 120)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(h_text)
        r.font.name = "Arial"
        r.font.size = Pt(8.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        
    schedule_data = [
        ("00:00 – 03:00\n(3 мин)", "1. Крючок и проблематизация", "Сайт ai-rop.ru (главный экран) или камера", "Клиент признает: «Да, звонки почти не слушаем»"),
        ("03:00 – 07:00\n(4 мин)", "2. Диагностика и Калькулятор", "Excel: Лист «Калькулятор_ROI» (3 числа)", "Шок от цифры потерь (от 1.5 до 4 млн ₽)"),
        ("07:00 – 13:00\n(6 мин)", "3. Демонстрация ИИ-Аудита", "Google Таблица: Лист «🎙️ ИИ_Аудит»", "«Вау-эффект»: видит 13 критериев и ошибки"),
        ("13:00 – 16:00\n(3 мин)", "4. Безопасность 152-ФЗ и процесс", "Слайд / Схема: Контур безопасности", "Снят страх утечки баз и геморроя с IT"),
        ("16:00 – 20:00\n(4 мин)", "5. Оффер и закрытие на тест", "Слайд Дорожная карта / Памятка 7 дней", "Клиент согласен скинуть 3 звонка прямо сейчас")
    ]
    
    for row_idx, data in enumerate(schedule_data, start=1):
        row = time_table.rows[row_idx]
        bg = HEX_LIGHT_BG if row_idx % 2 == 1 else "FFFFFF"
        for col_idx, text in enumerate(data):
            cell = row.cells[col_idx]
            cell.width = col_w[col_idx]
            set_cell_background(cell, bg)
            set_cell_margins(cell, 90, 90, 110, 110)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(text)
            r.font.name = "Arial"
            r.font.size = Pt(8.5)
            r.font.color.rgb = COLOR_DARK
            if col_idx == 0:
                r.font.bold = True
                r.font.color.rgb = COLOR_BLUE

    set_table_borders(time_table)
    doc.add_paragraph().paragraph_format.space_after = Pt(14)

    # ==========================
    # ПОШАГОВЫЙ СЦЕНАРИЙ (ДОСЛОВНЫЙ СКРИПТ)
    # ==========================
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(10)
    h2.paragraph_format.space_after = Pt(6)
    r_h2 = h2.add_run("2. Пошаговый сценарий 20 минут: Слово в слово")
    r_h2.font.name = "Arial"
    r_h2.font.size = Pt(14)
    r_h2.font.bold = True
    r_h2.font.color.rgb = COLOR_PRIMARY

    # --- ЭТАП 1 ---
    p_step1 = doc.add_paragraph()
    p_step1.paragraph_format.space_before = Pt(8)
    p_step1.paragraph_format.space_after = Pt(2)
    r = p_step1.add_run("ЭТАП 1 [00:00 – 03:00]: ВХОД, КОНТЕКСТ И ПРОБЛЕМАТИЗАЦИЯ")
    r.font.name = "Arial"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = COLOR_BLUE

    doc.add_paragraph(
        "• Что на экране: Камера спикера или титульный экран сайта https://ai-rop.ru.\n"
        "• Ваша цель: За 180 секунд перевести разговор из вежливого смолл-тока в признание ключевой боли отдела продаж."
    )

    add_callout(
        doc,
        "«[Имя клиента], приветствую! Давайте сразу договоримся по регламенту встречи. "
        "У нас ровно 20 минут. Моя задача — не продавать вам абстрактный искусственный интеллект, а показать одну конкретную вещь: "
        "где прямо сейчас в диалогах ваших менеджеров теряется от 15% до 30% чистой выручки, и как за 7 дней остановить этот слив без раздувания штата.\n\n"
        "Перед тем как открыть цифры, задам всего один вопрос: скажите, сколько звонков в день сейчас совершают ваши менеджеры, "
        "и какую долю из них физически успевает отслушать РОП?»",
        title="ЧТО СКАЗАТЬ СЛОВО В СЛОВО (СКРИПТ ВХОДА)",
        bg_color=HEX_BLUE_LIGHT,
        border_color=HEX_BLUE
    )

    doc.add_paragraph(
        "💡 Реакция клиента:\n"
        "В 95% случаев клиент скажет: «Ну, звонков 100–200, а РОП успевает дай бог 3–5 штук, если время есть». Либо: «Мы вообще не слушаем, некогда».\n\n"
        "👉 Ваш моментальный хук:\n"
        "«Вот именно. То есть 97% разговоров ваших менеджеров с клиентами — это полная 'черная дыра'. Вы платите за рекламу, лид приходит, "
        "менеджер что-то говорит, сделка срывается, а в CRM стоит комментарий 'думает' или 'дорого'. Давайте я покажу, сколько это стоит компании в рублях»."
    )

    # --- ЭТАП 2 ---
    p_step2 = doc.add_paragraph()
    p_step2.paragraph_format.space_before = Pt(12)
    p_step2.paragraph_format.space_after = Pt(2)
    r = p_step2.add_run("ЭТАП 2 [03:00 – 07:00]: «ПРАВИЛО 3 ЧИСЕЛ» И КАЛЬКУЛЯТОР ПОТЕРЬ")
    r.font.name = "Arial"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = COLOR_BLUE

    doc.add_paragraph(
        "• Что на экране: Открыть файл «Презентационная_Таблица_RevOps.xlsx», вкладка «Калькулятор_ROI».\n"
        "• Ваша цель: Вбить цифры клиента и показать размер ежемесячной катастрофы."
    )

    add_callout(
        doc,
        "«Посмотрите на экран. Это управленческий экспресс-калькулятор. Давайте вобьем ваши параметры:\n"
        "1. Сколько у вас менеджеров в линии? (Вбиваем, например: 5 чел)\n"
        "2. Какой средний чек закрытой сделки? (Вбиваем: 250 000 ₽)\n"
        "3. Какая текущая конверсия из лида в оплату? (Вбиваем: 8%)\n\n"
        "Смотрите, формула автоматически считает: даже если менеджеры всего в 15% звонков забывают назначить жесткий следующий шаг "
        "или сливают возражение 'я подумаю' без попытки квалификации, компания ежемесячно недополучает ровно 2 850 000 ₽.\n\n"
        "Это не виртуальные деньги — это зарплаты, которые вы уже заплатили менеджерам, и рекламный бюджет Яндекс.Директа, который улетел в трубу»",
        title="ЧТО СКАЗАТЬ СЛОВО В СЛОВО (ЭКОНОМИКА ПОТЕРЬ)",
        bg_color=HEX_AMBER_LIGHT,
        border_color=HEX_AMBER
    )

    # --- ЭТАП 3 ---
    p_step3 = doc.add_paragraph()
    p_step3.paragraph_format.space_before = Pt(12)
    p_step3.paragraph_format.space_after = Pt(2)
    r = p_step3.add_run("ЭТАП 3 [07:00 – 13:00]: СЕРДЦЕ ПРЕЗЕНТАЦИИ — ЖИВОЙ ИИ-АУДИТ")
    r.font.name = "Arial"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = COLOR_BLUE

    doc.add_paragraph(
        "• Что на экране: Открыть вкладку браузера с Google Таблицей: Лист «🎙️ ИИ_Аудит».\n"
        "• Ваша цель: Создать «вау-эффект» от глубины разбора (13 критериев) и показать, как ИИ ловит менеджеров на фатальных ошибках."
    )

    add_callout(
        doc,
        "«А теперь самое главное. Как мы находим эти потери. Переключаюсь на живой дашборд ИИ-Аудита.\n\n"
        "Мы анализируем не просто интонацию или паузы (как бесполезные отчеты у сотовых операторов), а бизнес-смысл по 13 жестким стандартам B2B:\n"
        "1. Выявление ЛПР: спросил ли менеджер, кто принимает финансовое решение, или 40 минут рассказывал секретарю?\n"
        "2. Жесткий следующий шаг: назначена ли конкретная дата и время звонка ('вторник в 14:00'), или менеджер сказал губительное 'ну вы подумайте и наберите'?\n"
        "3. Отработка цены: назвал ли менеджер ценность до того, как выпалил прайс?\n"
        "4. Защита маржинальности: не дал ли менеджер скидку сходу без встречных уступок?\n\n"
        "Посмотрите на строку 18: вот реальный звонок. Зеленый чип — стандарт выполнен. Красный чип — критическая ошибка. "
        "Нейросеть вытаскивает цитату менеджера: 'Ну если что, наш прайс на почте'. И тут же калькулятор фиксирует: упущенная вероятность сделки — 35%. "
        "РОПу больше не надо сидеть в наушниках по 3 часа. Он открывает таблицу в 09:00 и видит готовый светофор: кого похвалить, а кого срочно отправить на доработку»",
        title="ЧТО СКАЗАТЬ СЛОВО В СЛОВО (ДЕМОНСТРАЦИЯ 13 КРИТЕРИЕВ)",
        bg_color=HEX_BLUE_LIGHT,
        border_color=HEX_BLUE
    )

    # --- ЭТАП 4 ---
    p_step4 = doc.add_paragraph()
    p_step4.paragraph_format.space_before = Pt(12)
    p_step4.paragraph_format.space_after = Pt(2)
    r = p_step4.add_run("ЭТАП 4 [13:00 – 16:00]: БЕЗОПАСНОСТЬ 152-ФЗ И БЫСТРЫЙ СТАРТ")
    r.font.name = "Arial"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = COLOR_BLUE

    doc.add_paragraph(
        "• Что на экране: Лист «Контур_152ФЗ» в Excel или показать документ 152-ФЗ на бланке.\n"
        "• Ваша цель: Упредить страх утечки баз и снять возражение «наши айтишники будут настраивать это полгода»."
    )

    add_callout(
        doc,
        "«Знаю, какой вопрос у вас сейчас возник: 'Насколько это безопасно и сколько месяцев займет интеграция?'.\n\n"
        "Отвечаю сразу на оба:\n"
        "1. По безопасности (152-ФЗ): Мы работаем по защищенному контуру РФ. Перед стартом подписываем юридический NDA с материальной ответственностью. "
        "Все имена клиентов и номера телефонов деперсонализируются перед отправкой в речевые модели. Ваши базы остаются внутри компании.\n\n"
        "2. По интеграции: Вам НЕ НУЖНО переписывать CRM, нанимать программистов или покупать сервера. "
        "Первый аудит запускается за 10 минут: вы просто даете выгрузку 10-20 звонков из amoCRM или Битрикс24, либо подключаем легкий коннектор. "
        "Никакой головной боли для вашего IT-отдела»",
        title="ЧТО СКАЗАТЬ СЛОВО В СЛОВО (СНЯТИЕ СТРАХОВ)",
        bg_color=HEX_EMERALD_LIGHT,
        border_color=HEX_EMERALD
    )

    # --- ЭТАП 5 ---
    p_step5 = doc.add_paragraph()
    p_step5.paragraph_format.space_before = Pt(12)
    p_step5.paragraph_format.space_after = Pt(2)
    r = p_step5.add_run("ЭТАП 5 [16:00 – 20:00]: БЕЗОТКАЗНЫЙ ОФФЕР И ЗАКРЫТИЕ")
    r.font.name = "Arial"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = COLOR_BLUE

    doc.add_paragraph(
        "• Что на экране: Переключиться обратно на камеру или показать слайд «Дорожная карта пилота за 7 дней».\n"
        "• Ваша цель: Получить 3 звонка прямо сегодня до конца рабочего дня."
    )

    add_callout(
        doc,
        "«[Имя клиента], я не предлагаю вам сейчас подписывать годовой контракт на полмиллиона рублей. "
        "В бизнесе слова проверяются делом.\n\n"
        "Предлагаю сделать так:\n"
        "Вы прямо сейчас выбираете 3 любых звонка за вчерашний день — например, один успешный и два сорвавшихся. "
        "Скидываете их мне в защищенную папку или Telegram.\n\n"
        "Ровно через 24 часа я возвращаю вам готовый аналитический отчет в точно такой же Google Таблице по вашим 13 критериям. "
        "Вы своими глазами увидите, где менеджеры допустили просадку, и сколько денег это стоило компании.\n\n"
        "Для вас это абсолютно бесплатно и ни к чему не обязывает. Если увидите реальную пользу — обсудим 7-дневный пилот. "
        "Если нет — у вас останется глубокий независимый аудит звонков вашей команды.\n\n"
        "Куда удобнее скинуть ссылку на папку для записей — в WhatsApp или Telegram?»",
        title="ЧТО СКАЗАТЬ СЛОВО В СЛОВО (ЗАКРЫВАЮЩИЙ СКРИПТ)",
        bg_color=HEX_BLUE_LIGHT,
        border_color=HEX_BLUE
    )

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # ==========================
    # ШПАБАРГАЛКА ПО ВОЗРАЖЕНИЯМ
    # ==========================
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(10)
    h3.paragraph_format.space_after = Pt(6)
    r_h3 = h3.add_run("3. Шпаргалка: Отработка возражений во время демо")
    r_h3.font.name = "Arial"
    r_h3.font.size = Pt(14)
    r_h3.font.bold = True
    r_h3.font.color.rgb = COLOR_PRIMARY

    obj_table = doc.add_table(rows=5, cols=3)
    obj_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    obj_table.autofit = False
    
    col_w_obj = [Inches(1.8), Inches(2.2), Inches(2.8)]
    obj_headers = ["Возражение клиента", "Что он на самом деле имеет в виду", "Убойный ответ спикера"]
    
    hdr_row = obj_table.rows[0]
    for j, h_text in enumerate(obj_headers):
        cell = hdr_row.cells[j]
        cell.width = col_w_obj[j]
        set_cell_background(cell, HEX_PRIMARY)
        set_cell_margins(cell, 100, 100, 120, 120)
        p = cell.paragraphs[0]
        r = p.add_run(h_text)
        r.font.name = "Arial"
        r.font.size = Pt(8.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    obj_data = [
        ("«У нас специфический бизнес, ИИ не поймет наши термины»", "Боится, что робот наставит двоек нормальным продавцам.", "«Именно поэтому мы не используем шаблонные решения. Перед аудитом мы заносим в промпт ваш словарь терминов, маржинальные позиции и стоп-слова. На тесте 3 звонков вы сами проверите точность терминологии»."),
        ("«У нас РОП сам слушает звонки, ему за это платят»", "Оправдывает текущие затраты на ФОТ руководителя.", "«Конечно! И средняя зарплата сильного РОПа — 150 000–250 000 ₽. Неужели вы хотите, чтобы специалист такой стоимости тратил 25 часов в неделю на прослушивание 'гудков' и монотонных бесед, вместо того чтобы лично дожимать миллионные сделки?»"),
        ("«Пришлите на почту коммерческое, мы изучим»", "Вежливый способ слить встречу и ничего не делать.", "«[Имя], на почту я отправлю и калькулятор, и презентацию. Но без реальных звонков вашей компании это останется просто набором слайдов. Давайте прикрепим к КП отчет по 3 вашим звонкам — это займет у вас 2 минуты на пересылку, зато покажет факты»."),
        ("«Сколько стоит полноценное внедрение?»", "Проверяет, по карману ли продукт.", "«Пилотный спринт на 7 дней со 100% контролем отдела стоит от 45 000 до 75 000 ₽. С первой же спасенной сделки среднего чека в [200 000 ₽] он окупается в 3–4 раза. Но сначала давайте сделаем бесплатный тест 3 звонков, чтобы вы лично убедились в цифрах»."),
    ]

    for row_idx, data in enumerate(obj_data, start=1):
        row = obj_table.rows[row_idx]
        bg = HEX_LIGHT_BG if row_idx % 2 == 1 else "FFFFFF"
        for col_idx, text in enumerate(data):
            cell = row.cells[col_idx]
            cell.width = col_w_obj[col_idx]
            set_cell_background(cell, bg)
            set_cell_margins(cell, 90, 90, 110, 110)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(text)
            r.font.name = "Arial"
            r.font.size = Pt(8.5)
            r.font.color.rgb = COLOR_DARK
            if col_idx == 0:
                r.font.bold = True
                r.font.color.rgb = COLOR_ROSE

    set_table_borders(obj_table)
    doc.add_paragraph().paragraph_format.space_after = Pt(14)

    # ==========================
    # ИТОГОВЫЙ ЧЕК-ЛИСТ ПЕРЕД ЗВОНКОМ
    # ==========================
    h4 = doc.add_paragraph()
    h4.paragraph_format.space_before = Pt(8)
    h4.paragraph_format.space_after = Pt(4)
    r_h4 = h4.add_run("4. Финальный чек-лист готовности к презентации")
    r_h4.font.name = "Arial"
    r_h4.font.size = Pt(12)
    r_h4.font.bold = True
    r_h4.font.color.rgb = COLOR_PRIMARY

    checklist_items = [
        "Открыт сайт https://ai-rop.ru (демонстрирует надежность и технологичность).",
        "Открыта Google Таблица RevOps Master на вкладке «🎙️ ИИ_Аудит» со светофором критериев.",
        "Открыт Excel «Презентационная_Таблица_RevOps.xlsx» на листе «Калькулятор_ROI».",
        "Создана и скопирована персональная ссылка на папку (Яндекс.Диск / Google Диск) для приема 3 аудиозаписей.",
        "Рядом лежит открытый блокнот с реквизитами и шаблоном соглашения NDA (152-ФЗ)."
    ]

    for item in checklist_items:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r_box = p.add_run("  ☑  ")
        r_box.font.name = "Arial"
        r_box.font.bold = True
        r_box.font.color.rgb = COLOR_EMERALD
        r_text = p.add_run(item)
        r_text.font.name = "Arial"
        r_text.font.size = Pt(9.5)
        r_text.font.color.rgb = COLOR_DARK

    # Ensure parent dir exists
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    print(f"Presentation script successfully created at: {output_path}")

if __name__ == "__main__":
    dest1 = r"C:\Users\strel\Desktop\RevOps Platform\Презентация\Сценарий_20_минутной_Презентации_RevOps.docx"
    dest2 = r"C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\docs\Сценарий_20_минутной_Презентации_RevOps.docx"
    build_presentation_docx(dest1)
    build_presentation_docx(dest2)
