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

# Palette
COLOR_PRIMARY = RGBColor(15, 23, 42)      # Slate Navy #0F172A
COLOR_ACCENT = RGBColor(37, 99, 235)      # Royal Blue #2563EB
COLOR_SECONDARY = RGBColor(71, 85, 105)   # Muted Slate #475569
COLOR_SUCCESS = RGBColor(16, 185, 129)    # Emerald Green #10B981
COLOR_WARNING = RGBColor(245, 158, 11)    # Amber #F59E0B
COLOR_MUTED = RGBColor(100, 116, 139)     # Gray #64748B

HEX_BG_LIGHT = "F8FAFC"
HEX_BG_HEADER = "0F172A"
HEX_BORDER = "E2E8F0"
HEX_ACCENT_BG = "EFF6FF"
HEX_SUCCESS_BG = "ECFDF5"

def set_cell_shading(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=120, bottom=120, left=160, right=160):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def add_header(doc):
    section = doc.sections[0]
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)
    
    header = section.header
    p = header.paragraphs[0]
    p.text = "REVOPS ENTERPRISE OS V18.0  |  РУКОВОДСТВО ПОЛЬЗОВАТЕЛЯ"
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.runs[0].font.size = Pt(8.5)
    p.runs[0].font.color.rgb = COLOR_MUTED
    p.runs[0].font.name = "Calibri"

def add_callout(doc, text, title="ВАЖНОЕ ПРАВИЛО", alert_type="info"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    
    bg_color = HEX_ACCENT_BG if alert_type == "info" else HEX_SUCCESS_BG
    border_color = "2563EB" if alert_type == "info" else "10B981"
    
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
    r_title.font.color.rgb = COLOR_ACCENT if alert_type == "info" else COLOR_SUCCESS
    
    r_text = p.add_run(text)
    r_text.font.name = "Calibri"
    r_text.font.size = Pt(10)
    r_text.font.color.rgb = COLOR_PRIMARY
    
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_after = Pt(6)

def build_client_manual_doc():
    doc = docx.Document()
    add_header(doc)
    
    # Title Block
    p_badge = doc.add_paragraph()
    p_badge.paragraph_format.space_before = Pt(10)
    p_badge.paragraph_format.space_after = Pt(4)
    r_badge = p_badge.add_run("OFFICIAL CLIENT MANUAL  •  ENTERPRISE EDITION V18.0")
    r_badge.font.name = "Calibri"
    r_badge.font.size = Pt(9.5)
    r_badge.font.bold = True
    r_badge.font.color.rgb = COLOR_ACCENT
    
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_after = Pt(8)
    r_title = p_title.add_run("RevOps Enterprise OS V18.0\nРуководство пользователя")
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(26)
    r_title.font.bold = True
    r_title.font.color.rgb = COLOR_PRIMARY
    
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(16)
    r_sub = p_sub.add_run("Регламент утреннего 15-минутного контроля сделок РОПом, сквозная аналитика выручки и 100% ИИ-аудит звонков (Whisper + DeepSeek)")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(12)
    r_sub.font.color.rgb = COLOR_SECONDARY
    
    add_callout(
        doc,
        "Данный документ содержит регламент работы руководства компании и отдела продаж с операционной системой RevOps OS. Система ликвидирует 4 ключевые утечки кассы и оцифровывает качество каждого разговора менеджера с клиентом.",
        title="НАЗНАЧЕНИЕ ПЛАТФОРМЫ",
        alert_type="info"
    )

    # Section 1: Concept
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("1. Ключевая бизнес-задача RevOps OS")
    r_h1.font.name = "Calibri"
    r_h1.font.color.rgb = COLOR_PRIMARY
    r_h1.font.size = Pt(16)
    
    p_desc = doc.add_paragraph()
    p_desc.paragraph_format.space_after = Pt(8)
    p_desc.add_run(
        "В 90% B2B-компаний отдел продаж теряет до 40% потенциальной выручки не из-за плохого продукта или нехватки лидов, "
        "а из-за 4 системных слепых зон в операционной работе. RevOps OS закрывает эти дыры за счет ежедневного сквозного мониторинга:"
    )
    
    # Table of 4 Leaks
    tbl_leaks = doc.add_table(rows=5, cols=4)
    tbl_leaks.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["№", "Узкое место воронки продаж", "Отраслевой ущерб", "Как RevOps OS устраняет утечку"]
    
    for c_idx, h_text in enumerate(headers):
        cell = tbl_leaks.cell(0, c_idx)
        set_cell_shading(cell, HEX_BG_HEADER)
        set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
        p = cell.paragraphs[0]
        r = p.add_run(h_text)
        r.font.name = "Calibri"
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    leaks_data = [
        ("1", "Слив лидов на звонках (нет Next Step)", "28% кассы", "ИИ-аудит 100% звонков через Whisper: детекция фиксации точной даты и времени следующего шага."),
        ("2", "Зависание сделок на этапе КП > 48 часов", "18% кассы", "Авторадар просрочки SLA: эскалация РОПу на 3-й день с готовым сценарием спасения."),
        ("3", "Брошенные отказники без дожима (L1–L4)", "12% кассы", "AI Recovery Engine: системная реактивация списанных сделок через WhatsApp и Telegram."),
        ("4", "Зависшая дебиторка и задержка оплат", "10% кассы", "Платежный календарь DSO: контроль кассовых разрывов и скоринг риска задержки траншей.")
    ]
    
    for r_idx, row_vals in enumerate(leaks_data, 1):
        for c_idx, val in enumerate(row_vals):
            cell = tbl_leaks.cell(r_idx, c_idx)
            set_cell_shading(cell, HEX_BG_LIGHT if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, top=90, bottom=90, left=120, right=120)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.name = "Calibri"
            r.font.size = Pt(9.5)
            r.font.color.rgb = COLOR_PRIMARY
            if c_idx == 2:
                r.bold = True
                r.font.color.rgb = COLOR_WARNING
                
    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 2: Dashboards Navigation
    h2 = doc.add_heading(level=1)
    r_h2 = h2.add_run("2. Архитектура рабочих экранов платформы")
    r_h2.font.name = "Calibri"
    r_h2.font.color.rgb = COLOR_PRIMARY
    r_h2.font.size = Pt(16)
    
    dashboards = [
        ("⚡ 1. Экспресс_Калькулятор_3_Цифры", "Для кого: Собственник, CEO, РОП", "Диагностика потенциала кассы за 30 секунд. Вы вводите 3 цифры (лиды в месяц, средний чек, количество менеджеров) — система мгновенно рассчитывает емкость базы, упущенную выгоду по 4 узким местам и формирует 3 финансовых сценария роста (+15%, +30%, +50%)."),
        ("📋 2. Пульт_РОПа_15_Минут", "Для кого: Коммерческий директор, РОП", "Главный оперативный экран. Каждое утро в 09:00 РОП открывает этот пульт: система формирует ТОП-5 горящих сделок в риске с расчетом суммы под угрозой срыва и автоматически генерирует персональный скрипт перехвата клиента."),
        ("🎙️ 3. ИИ_Аудит", "Для кого: РОП, Руководитель группы, ОКК", "Витрина речевой аналитики 100% диалогов. Нейросеть Whisper Large V3 переводит звонки в текст, а алгоритм DeepSeek V3 оценивает каждый звонок по 13 жестким стандартам B2B-продаж. Выводит точные текстовые цитаты ошибок менеджеров и советы РОПу."),
        ("💸 4. Диагностика_Утечек_ОП", "Для кого: Финансовый директор, CFO, Собственник", "Глубокий аудит 7 смертных грехов отдела продаж (сливы звонков, скидки без торга, потеря темпа в CRM). Рассчитывает чистую финансовую окупаемость спринта внедрения (ROI 300%+ за 21 день)."),
        ("📄 5. Executive_OnePager", "Для кого: Собственник бизнеса, Совет директоров", "Управленческий дайджест на один экран. Сводка план/факт выручки, выполнение месячной цели, объем активного пайплайна, прогноз Run-Rate и интегральный индекс здоровья воронки.")
    ]
    
    for title, role, text in dashboards:
        p_t = doc.add_paragraph()
        p_t.paragraph_format.space_before = Pt(8)
        p_t.paragraph_format.space_after = Pt(2)
        r_t = p_t.add_run(title)
        r_t.bold = True
        r_t.font.name = "Calibri"
        r_t.font.size = Pt(11.5)
        r_t.font.color.rgb = COLOR_ACCENT
        
        p_r = doc.add_paragraph()
        p_r.paragraph_format.space_after = Pt(3)
        r_r = p_r.add_run(role)
        r_r.italic = True
        r_r.font.name = "Calibri"
        r_r.font.size = Pt(9.5)
        r_r.font.color.rgb = COLOR_MUTED
        
        p_x = doc.add_paragraph()
        p_x.paragraph_format.space_after = Pt(8)
        r_x = p_x.add_run(text)
        r_x.font.name = "Calibri"
        r_x.font.size = Pt(10)
        r_x.font.color.rgb = COLOR_PRIMARY

    # Section 3: 15-min Ritual for ROP
    h3 = doc.add_heading(level=1)
    r_h3 = h3.add_run("3. Боевой регламент РОПа: «15 минут утреннего контроля»")
    r_h3.font.name = "Calibri"
    r_h3.font.color.rgb = COLOR_PRIMARY
    r_h3.font.size = Pt(16)
    
    p_rit = doc.add_paragraph()
    p_rit.paragraph_format.space_after = Pt(8)
    p_rit.add_run("Чтобы отдел продаж выполнял план, руководителю больше не нужно вручную переслушивать сотни часов записей или часами копаться в CRM. Достаточно следовать простому 3-шаговому регламенту:")

    steps = [
        ("Шаг 1 (09:00 – 09:05): Экспресс-аудит Пульта РОПа", "Откройте вкладку «📋 Пульт_РОПа_15_Минут». Посмотрите верхние индикаторы: есть ли сделки с критическим риском срыва (колонка C) и какова общая сумма в риске (колонка D)."),
        ("Шаг 2 (09:05 – 09:12): Распределение задач по ТОП-5 рискам", "Изучите таблицу ТОП-5 сделок (строки 9–13). Обратите внимание на колонку H («Персональный сценарий спасения сделки»). Скопируйте готовый скрипт и отправьте его ответственному менеджеру или совершите звонок клиенту лично."),
        ("Шаг 3 (09:12 – 09:15): Контроль дефектов в «🎙️ ИИ_Аудит»", "Перейдите на лист «🎙️ ИИ_Аудит» (строка 30). Просмотрите дословные цитаты грубых ошибок менеджеров за вчерашний день (отпустил без Next Step, дал необоснованную скидку, не отработал возражение). Озвучьте ошибки на утренней 10-минутной планерке.")
    ]
    
    for s_title, s_text in steps:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.space_before = Pt(6)
        p_st.paragraph_format.space_after = Pt(2)
        r_st = p_st.add_run(s_title)
        r_st.bold = True
        r_st.font.name = "Calibri"
        r_st.font.size = Pt(11)
        r_st.font.color.rgb = COLOR_PRIMARY
        
        p_sx = doc.add_paragraph()
        p_sx.paragraph_format.space_after = Pt(6)
        r_sx = p_sx.add_run(s_text)
        r_sx.font.name = "Calibri"
        r_sx.font.size = Pt(10)
        r_sx.font.color.rgb = COLOR_SECONDARY

    add_callout(
        doc,
        "В течение рабочего дня РОПу не нужно держать таблицу открытой: при возникновении критического брака на звонке (<9 баллов) бот @revops_supervisor_bot пришлет алерт в Telegram в течение 30 секунд с аудиофрагментом и цитатой ошибки.",
        title="РЕЖИМ TELEGRAM-АЛЕРТОВ",
        alert_type="success"
    )

    # Section 4: 13 Speech Standards
    h4 = doc.add_heading(level=1)
    r_h4 = h4.add_run("4. 13 Стандартов качества диалогов (Чек-лист ИИ)")
    r_h4.font.name = "Calibri"
    r_h4.font.color.rgb = COLOR_PRIMARY
    r_h4.font.size = Pt(16)
    
    p_std = doc.add_paragraph()
    p_std.paragraph_format.space_after = Pt(8)
    p_std.add_run("Нейросеть проверяет каждый разговор менеджера по объективным маркерам (0 или 1 балл):")
    
    tbl_std = doc.add_table(rows=14, cols=3)
    tbl_std.alignment = WD_TABLE_ALIGNMENT.CENTER
    std_headers = ["Критерий", "Что проверяет нейросеть", "Маркер дефекта (0 баллов)"]
    
    for c_idx, h_text in enumerate(std_headers):
        cell = tbl_std.cell(0, c_idx)
        set_cell_shading(cell, HEX_BG_HEADER)
        set_cell_margins(cell, top=90, bottom=90, left=100, right=100)
        p = cell.paragraphs[0]
        r = p.add_run(h_text)
        r.font.name = "Calibri"
        r.font.size = Pt(9)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    criteria = [
        ("К1. Фиксация Next Step", "Назначена точная дата и время контакта", "«Я наберу на следующей неделе»"),
        ("К2. Инициатива в диалоге", "Менеджер ведет диалог вопросами", "Пассивно отвечает, клиент ведет разговор"),
        ("К3. Квалификация ЛПР", "Уточнена роль ЛПР и схема принятия решения", "Общение с секретарем/не ЛПР без квалификации"),
        ("К4. Выявление болей", "Озвучена финансовая или операционная боль", "Презентация продукта вслепую без вопросов"),
        ("К5. Защита цены", "Озвучена ценность до называния стоимости", "Назвал цену в первые 2 минуты без пользы"),
        ("К6. Отработка «Дорого»", "Возражение переведено в окупаемость/сравнение", "Сразу согласился дать скидку"),
        ("К7. Твердые кейсы", "Приведен пример аналогичного клиента с цифрами", "Абстрактные обещания без твердых фактов"),
        ("К8. Попытка закрытия", "Озвучен призыв к целевому действию (счет, встреча)", "Разговор закончился ничем"),
        ("К9. Регламент встречи", "Соблюден корпоративный стандарт приветствия", "Невнятное приветствие, не назвал компанию"),
        ("К10. Чистота речи", "Уверенный голос, отсутствие пауз > 5 секунд", "Слова-паразиты («э-э-э», «как бы»), неуверенность"),
        ("К11. Активное слушание", "Менеджер не перебивает, резюмирует слова клиента", "Перебивание клиента, спор"),
        ("К12. Подведение итогов", "В конце проговорены договоренности сторон", "Резкое завершение разговора"),
        ("К13. Защита маржи", "Скидка предоставлена только в обмен на объем/срок", "Скидка «просто так» по первой просьбе")
    ]
    
    for r_idx, (c_name, c_chk, c_def) in enumerate(criteria, 1):
        for c_idx, val in enumerate([c_name, c_chk, c_def]):
            cell = tbl_std.cell(r_idx, c_idx)
            set_cell_shading(cell, HEX_BG_LIGHT if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.name = "Calibri"
            r.font.size = Pt(8.5)
            r.font.color.rgb = COLOR_PRIMARY
            if c_idx == 0:
                r.bold = True
                
    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 5: Integration in 3 steps
    h5 = doc.add_heading(level=1)
    r_h5 = h5.add_run("5. Подключение amoCRM / Битрикс24 за 3 шага")
    r_h5.font.name = "Calibri"
    r_h5.font.color.rgb = COLOR_PRIMARY
    r_h5.font.size = Pt(16)
    
    int_steps = [
        ("Шаг 1: Активация корпоративной таблицы", "В меню Google Таблицы выберите «Файл» ➔ «Создать копию». Откройте лист «⚙️ Настройки» и укажите название вашей компании (B3) и плановую выручку на месяц (B9)."),
        ("Шаг 2: Включение штатного вебхука CRM (2 минуты)", "В amoCRM откройте «Настройки» ➔ «Интеграции» ➔ «Вебхуки». Добавьте URL сервера: https://ai-rop.ru/api/v1/crm/call-webhook. Отметьте событие: «Добавлен звонок к сделке». Все записи звонков начнут автоматически оцениваться ИИ."),
        ("Шаг 3: Подключение Telegram-супервайзера", "В Telegram найдите бота @revops_supervisor_bot, нажмите /start и введите ваш Tenant ID из ячейки E3 листа «⚙️ Настройки». РОП и руководство начнут мгновенно получать сигналы о срыве сделок.")
    ]
    
    for s_title, s_text in int_steps:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.space_before = Pt(6)
        p_st.paragraph_format.space_after = Pt(2)
        r_st = p_st.add_run(s_title)
        r_st.bold = True
        r_st.font.name = "Calibri"
        r_st.font.size = Pt(10.5)
        r_st.font.color.rgb = COLOR_ACCENT
        
        p_sx = doc.add_paragraph()
        p_sx.paragraph_format.space_after = Pt(6)
        r_sx = p_sx.add_run(s_text)
        r_sx.font.name = "Calibri"
        r_sx.font.size = Pt(9.5)
        r_sx.font.color.rgb = COLOR_PRIMARY

    # Section 6: Security & 152-FZ
    h6 = doc.add_heading(level=1)
    r_h6 = h6.add_run("6. Безопасность данных и соответствие 152-ФЗ РФ")
    r_h6.font.name = "Calibri"
    r_h6.font.color.rgb = COLOR_PRIMARY
    r_h6.font.size = Pt(16)
    
    p_sec = doc.add_paragraph()
    p_sec.paragraph_format.space_after = Pt(6)
    p_sec.add_run(
        "• Обработка аудиозаписей и стенограмм происходит в закрытом контуре защищенных серверов на территории Российской Федерации (Selectel / TimeWeb Cloud).\n"
        "• Перед отправкой в языковую модель персональные данные (номера телефонов, фамилии, реквизиты) автоматически маскируются алгоритмом PII-redaction.\n"
        "• Доступ к вашей рабочей таблице разграничен встроенной ролевой моделью RBAC: менеджеры видят только свои звонки, РОП видит отдел, генеральный директор видит сводную кассу."
    )
    
    # Footer Contacts
    p_foot = doc.add_paragraph()
    p_foot.paragraph_format.space_before = Pt(20)
    p_foot.paragraph_format.space_after = Pt(4)
    r_foot = p_foot.add_run("СЛУЖБА ЗАБОТЫ И ТЕХНИЧЕСКАЯ ПОДДЕРЖКА REVOPS OS:")
    r_foot.bold = True
    r_foot.font.name = "Calibri"
    r_foot.font.size = Pt(9.5)
    r_foot.font.color.rgb = COLOR_MUTED
    
    p_contact = doc.add_paragraph()
    r_c = p_contact.add_run("Telegram: @revops_support  |  Официальный портал: https://ai-rop.ru  |  E-mail: support@ai-rop.ru")
    r_c.font.name = "Calibri"
    r_c.font.size = Pt(9.5)
    r_c.font.color.rgb = COLOR_ACCENT

    # Save DOCX
    docx_path = os.path.abspath("docs/Руководство_Клиента_RevOps_OS_V18.docx")
    doc.save(docx_path)
    print(f"✅ DOCX saved successfully: {docx_path}")
    
    # Also save to presentation folder
    pres_docx_path = os.path.abspath("presentation/Руководство_Клиента_RevOps_OS_V18.docx")
    doc.save(pres_docx_path)
    
    return docx_path

def convert_docx_to_pdf(docx_path):
    pdf_path = docx_path.replace(".docx", ".pdf")
    ps_script = f'''
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {{
    $doc = $word.Documents.Open("{docx_path}")
    $doc.SaveAs([ref]"{pdf_path}", [ref]17)
    $doc.Close()
    Write-Host "SUCCESS: PDF created at {pdf_path}"
}} catch {{
    Write-Host "ERROR: $_"
}} finally {{
    $word.Quit()
}}
'''
    ps_file = "scripts/convert_manual_to_pdf.ps1"
    with open(ps_file, "w", encoding="utf-8") as f:
        f.write(ps_script)
        
    res = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_file], capture_output=True, text=True)
    print(res.stdout)
    if os.path.exists(pdf_path):
        print(f"✅ PDF successfully generated: {pdf_path} ({os.path.getsize(pdf_path):,} bytes)")
        # Copy to presentation folder as well
        pres_pdf = os.path.abspath("presentation/Руководство_Клиента_RevOps_OS_V18.pdf")
        with open(pdf_path, "rb") as f_in, open(pres_pdf, "wb") as f_out:
            f_out.write(f_in.read())
        print(f"✅ PDF copied to presentation folder: {pres_pdf}")
    else:
        print("⚠️ PDF conversion failed or file not found.")

if __name__ == "__main__":
    d_path = build_client_manual_doc()
    convert_docx_to_pdf(d_path)
