import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
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

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
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
    p.text = "REVOPS ENTERPRISE OS V18.0  |  БОЕВОЕ РУКОВОДСТВО ОСНОВАТЕЛЯ"
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.runs[0].font.size = Pt(8.5)
    p.runs[0].font.color.rgb = COLOR_MUTED
    p.runs[0].font.name = "Calibri"

def add_callout(doc, text, title="ПРАВИЛО ФАУНДЕРА", alert_type="info"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    
    bg_color = HEX_ACCENT_BG if alert_type == "info" else HEX_SUCCESS_BG
    border_color = "2563EB" if alert_type == "info" else "10B981"
    
    set_cell_shading(cell, bg_color)
    set_cell_margins(cell, top=120, bottom=120, left=160, right=160)
    
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:left w:val="single" w:sz="28" w:space="0" w:color="{border_color}"/><w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/></w:tcBorders>')
    tcPr.append(tcBorders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(4)
    r_title = p.add_run(f"⚡ {title}: ")
    r_title.bold = True
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(9.5)
    r_title.font.color.rgb = COLOR_ACCENT if alert_type == "info" else COLOR_SUCCESS
    
    r_text = p.add_run(text)
    r_text.font.name = "Calibri"
    r_text.font.size = Pt(9.5)
    r_text.font.color.rgb = COLOR_PRIMARY
    
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_after = Pt(4)

def build_founder_handbook_doc():
    doc = docx.Document()
    add_header(doc)
    
    # Title Block
    p_badge = doc.add_paragraph()
    p_badge.paragraph_format.space_before = Pt(6)
    p_badge.paragraph_format.space_after = Pt(2)
    r_badge = p_badge.add_run("CONFIDENTIAL  •  INTERNAL FOUNDER OPERATIONAL PLAYBOOK")
    r_badge.font.name = "Calibri"
    r_badge.font.size = Pt(9.5)
    r_badge.font.bold = True
    r_badge.font.color.rgb = COLOR_WARNING
    
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_after = Pt(6)
    r_title = p_title.add_run("Боевое руководство основателя RevOps OS V18.0")
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(24)
    r_title.font.bold = True
    r_title.font.color.rgb = COLOR_PRIMARY
    
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(12)
    r_sub = p_sub.add_run("Пошаговый регламент: продажи за 15 минут в Zoom, запуск клиентов за 15 минут, подключение amoCRM / Битрикс24 и экономика с маржой 97%")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(11.5)
    r_sub.font.color.rgb = COLOR_SECONDARY
    
    add_callout(
        doc,
        "Это ваше личное руководство к действию. Здесь собраны точные формулировки для разговора с собственниками, ссылки на боевые таблицы, алгоритм настройки CRM за 5 минут и финансовая модель, позволяющая зарабатывать от 500 000 до 1 500 000 ₽ чистой прибыли в месяц при себестоимости обслуживания всего ~1 100 ₽ на клиента.",
        title="СТРАТЕГИЧЕСКИЙ ФОКУС",
        alert_type="info"
    )

    # Chapter 1
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("1. Ваш рабочий арсенал: где что лежит")
    r_h1.font.name = "Calibri"
    r_h1.font.color.rgb = COLOR_PRIMARY
    r_h1.font.size = Pt(15)
    
    p_ch1 = doc.add_paragraph()
    p_ch1.paragraph_format.space_after = Pt(6)
    p_ch1.add_run("Вся ваша инфраструктура разделена на два контура — презентационный и клиентский:")
    
    tbl_tools = doc.add_table(rows=5, cols=3)
    tbl_tools.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_headers = ["Инструмент", "Ссылка / Путь к файлу", "Как использовать"]
    for c_idx, h_text in enumerate(t_headers):
        cell = tbl_tools.cell(0, c_idx)
        set_cell_shading(cell, HEX_BG_HEADER)
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        p = cell.paragraphs[0]
        r = p.add_run(h_text)
        r.font.name = "Calibri"
        r.font.size = Pt(9)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    tools_data = [
        ("1. Презентационная таблица (Showcase Master)", "ID: 1QnjrrbpqhYssofchee7G06szWrFvBCcOqGIVokjqdVc\n(лист: ⚡ Экспресс_3_Цифры)", "Открывать на демонстрациях в Zoom. 100% заполнена сделками, звонками и цифрами потерь (1.122М ₽)."),
        ("2. Чистый клиентский шаблон (Client Starter)", "ID: 1jBBotOfFh-XEJGScJyQi10jiFrna66OpJ91yPDrXy2A\n(ссылка /copy)", "Передавать клиенту после оплаты. 100% чистый шаблон (0 ₽, 0 сделок). Клиент жмет «Создать копию»."),
        ("3. Клиентское руководство (Word + PDF)", "docs/Руководство_Клиента_RevOps_OS_V18.pdf\n(и .docx в presentation/)", "Отправлять собственнику и РОПу в Telegram после созвона как подтверждение солидности решения."),
        ("4. Микросервис ИИ (AI Speech Proxy)", "services/ai_speech_blackbox_proxy.py", "Принимает аудиозвонки из CRM, транскрибирует через Whisper, скорит через DeepSeek и льет в raw_calls.")
    ]
    for r_idx, row_vals in enumerate(tools_data, 1):
        for c_idx, val in enumerate(row_vals):
            cell = tbl_tools.cell(r_idx, c_idx)
            set_cell_shading(cell, HEX_BG_LIGHT if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, top=70, bottom=70, left=100, right=100)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.name = "Calibri"
            r.font.size = Pt(8.5)
            r.font.color.rgb = COLOR_PRIMARY
            if c_idx == 0:
                r.bold = True
    
    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Chapter 2
    h2 = doc.add_heading(level=1)
    r_h2 = h2.add_run("2. Сценарий продающего созвона в Zoom (15 минут)")
    r_h2.font.name = "Calibri"
    r_h2.font.color.rgb = COLOR_PRIMARY
    r_h2.font.size = Pt(15)
    
    demo_steps = [
        ("Шаг 1 (00:00 – 03:00): Шок-контент на Экспресс-Калькуляторе",
         "Откройте лист «⚡ Экспресс_Калькулятор_3_Цифры» в Презентационной таблице.\n"
         "Слова слово в слово:\n"
         "«Иван Иванович, добрый день! Давайте без лишних презентаций за 30 секунд оцифруем ваш отдел продаж прямо на живой модели. Назовите 3 цифры: сколько лидов заходит в месяц, какой средний чек закрытой сделки и сколько менеджеров в штате?»\n"
         "Вбиваете его цифры в B5:B7 (например, 150 лидов, 400 000 ₽ чек, 4 менеджера). Показываете ячейку D19:\n"
         "«Смотрите, прямо сейчас на 4 узких местах вы ежемесячно теряете от 1 122 000 ₽. 28% улетает на звонках, когда менеджеры отпускают клиента без даты следующего шага. 18% зависает на этапе КП дольше 48 часов. Давайте покажу, как за 7 дней вернуть эти деньги в кассу»."),
         
        ("Шаг 2 (03:00 – 07:00): Пульт РОПа «15 минут в день»",
         "Перейдите на лист «📋 Пульт_РОПа_15_Минут».\n"
         "Слова:\n"
         "«Главная проблема РОПа — он тонет в рутине и не понимает, какие сделки горят. В нашей платформе рабочий день РОПа начинается в 09:00 ровно с этой вкладки. Система сама отбирает ТОП-5 сделок с риском срыва.\n"
         "Обратите внимание на колонку H: РОПу не надо гадать, что делать. Система уже сформировала персональный скрипт перехвата: например, направить клиенту в WhatsApp 2 конкретных слота встречи и спец-гарантию. Сделка на 850 000 ₽ спасается за 3 минуты»."),
         
        ("Шаг 3 (07:00 – 11:00): 100% ИИ-аудит звонков (Киллер-фича)",
         "Перейдите на лист «🎙️ ИИ_Аудит».\n"
         "Слова:\n"
         "«Обычно РОП успевает прослушать максимум 5% звонков. У нас нейросеть Whisper Large V3 слушает 100% звонков из вашей CRM и оценивает их по 13 жестким стандартам B2B.\n"
         "Посмотрите на строку 30: ИИ вытаскивает дословную цитату ошибки менеджера («ну... мы можем сделать скидочку 10%»). Нейросеть мгновенно ставит оценку и присылает алерт РОПу в Telegram в течение 30 секунд после завершения разговора»."),
         
        ("Шаг 4 (11:00 – 13:00): Диагностика утечек и окупаемость",
         "Перейдите на лист «💸 Диагностика_Утечек_ОП».\n"
         "Слова:\n"
         "«Внедрение RevOps OS окупается с первой же спасенной сделки. При стоимости спринта внедрения 250 000 ₽ чистый прирост кассы составляет от 900 000 ₽ в месяц. Срок окупаемости системы — всего 8 дней»."),
         
        ("Шаг 5 (13:00 – 15:00): Закрытие сделки (Оффер)",
         "Слова:\n"
         "«Мы не предлагаем верить на слово. Мы запускаем 7-дневный пилотный спринт: подключаем вашу CRM, настраиваем ИИ-контроль звонков, оцифровываем первые 100 диалогов и через 7 дней на планерке показываем вам точные точки сливов. Если вы не увидите ценности — мы возвращаем 100% оплаты по договору. Стартуем с понедельника?»")
    ]
    
    for s_title, s_text in demo_steps:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.space_before = Pt(6)
        p_st.paragraph_format.space_after = Pt(2)
        r_st = p_st.add_run(s_title)
        r_st.bold = True
        r_st.font.name = "Calibri"
        r_st.font.size = Pt(11)
        r_st.font.color.rgb = COLOR_ACCENT
        
        p_sx = doc.add_paragraph()
        p_sx.paragraph_format.space_after = Pt(6)
        r_sx = p_sx.add_run(s_text)
        r_sx.font.name = "Calibri"
        r_sx.font.size = Pt(9.5)
        r_sx.font.color.rgb = COLOR_PRIMARY

    # Chapter 3
    h3 = doc.add_heading(level=1)
    r_h3 = h3.add_run("3. Технический запуск нового клиента за 15 минут")
    r_h3.font.name = "Calibri"
    r_h3.font.color.rgb = COLOR_PRIMARY
    r_h3.font.size = Pt(15)
    
    onboard_steps = [
        ("Шаг 1. Передача рабочей копии клиенту (2 минуты)",
         "Отправьте клиенту ссылку автоматического копирования:\n"
         "https://docs.google.com/spreadsheets/d/1jBBotOfFh-XEJGScJyQi10jiFrna66OpJ91yPDrXy2A/copy\n"
         "Клиент нажимает «Создать копию». Таблица сохраняется на его личном Google Диске. Попросите клиента добавить ваш служебный email (из service_account.json) с правами «Редактор»."),
         
        ("Шаг 2. Конфигурация в листе ⚙️ Настройки (3 минуты)",
         "В открывшейся копии клиента перейдите на вкладку «⚙️ Настройки»:\n"
         "• Ячейка B3: Название компании (например: ООО «ТехноСнаб»);\n"
         "• Ячейка E3: Уникальный Tenant ID (например: CLIENT-TENANT-002);\n"
         "• Ячейка B8: Дата запуска (текущая дата);\n"
         "• Ячейка B9: Плановая выручка на месяц (например: 5 000 000 ₽);\n"
         "• Диапазон B21:B24: Имена и роли сотрудников отдела продаж."),
         
        ("Шаг 3. Подключение звонков из amoCRM / Битрикс24 (5 минут)",
         "Вам не нужны программисты. В amoCRM клиента зайдите: «Настройки» ➔ «Интеграции» ➔ «Вебхуки» ➔ «+ Добавить вебхук».\n"
         "• URL вебхука: https://ai-rop.ru/api/v1/crm/call-webhook?tenant_id=CLIENT-TENANT-002\n"
         "• Событие: ☑ «Сделка: Добавлено примечание (звонок)».\n"
         "Теперь каждый звонок из CRM автоматически отправляется на ваш сервер, оценивается за 15 секунд и появляется в отчетах."),
         
        ("Шаг 4. Подключение Telegram-супервайзера (5 минут)",
         "Создайте бота в @BotFather (или используйте единого @revops_supervisor_bot). Попросите РОПа нажать /start в боте. Узнайте его Chat ID через @userinfobot и впишите в лист «⚙️ Настройки» в ячейки B52:B54.")
    ]
    
    for s_title, s_text in onboard_steps:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.space_before = Pt(6)
        p_st.paragraph_format.space_after = Pt(2)
        r_st = p_st.add_run(s_title)
        r_st.bold = True
        r_st.font.name = "Calibri"
        r_st.font.size = Pt(10.5)
        r_st.font.color.rgb = COLOR_PRIMARY
        
        p_sx = doc.add_paragraph()
        p_sx.paragraph_format.space_after = Pt(6)
        r_sx = p_sx.add_run(s_text)
        r_sx.font.name = "Calibri"
        r_sx.font.size = Pt(9.5)
        r_sx.font.color.rgb = COLOR_SECONDARY

    # Chapter 4
    h4 = doc.add_heading(level=1)
    r_h4 = h4.add_run("4. Финансовая модель и ценообразование")
    r_h4.font.name = "Calibri"
    r_h4.font.color.rgb = COLOR_PRIMARY
    r_h4.font.size = Pt(15)
    
    tbl_pricing = doc.add_table(rows=4, cols=4)
    tbl_pricing.alignment = WD_TABLE_ALIGNMENT.CENTER
    p_headers = ["Тарифный пакет", "Цена для клиента", "Ваша себестоимость", "Ваша чистая маржа"]
    for c_idx, h_text in enumerate(p_headers):
        cell = tbl_pricing.cell(0, c_idx)
        set_cell_shading(cell, HEX_BG_HEADER)
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        p = cell.paragraphs[0]
        r = p.add_run(h_text)
        r.font.name = "Calibri"
        r.font.size = Pt(9)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    pricing_data = [
        ("1. Быстрый пилот (7 дней)", "49 000 – 75 000 ₽", "~300 ₽", "99% (48 700 – 74 700 ₽)"),
        ("2. Внедрение под ключ", "150 000 – 350 000 ₽", "~1 000 ₽", "99% (149 000 – 349 000 ₽)"),
        ("3. AI-супервайзер (подписка)", "35 000 – 50 000 ₽ / мес", "~1 100 ₽ / мес", "97% (33 900 – 48 900 ₽ / мес)")
    ]
    for r_idx, row_vals in enumerate(pricing_data, 1):
        for c_idx, val in enumerate(row_vals):
            cell = tbl_pricing.cell(r_idx, c_idx)
            set_cell_shading(cell, HEX_BG_LIGHT if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, top=70, bottom=70, left=100, right=100)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.name = "Calibri"
            r.font.size = Pt(8.5)
            r.font.color.rgb = COLOR_PRIMARY
            if c_idx == 3:
                r.bold = True
                r.font.color.rgb = COLOR_SUCCESS

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Chapter 5
    h5 = doc.add_heading(level=1)
    r_h5 = h5.add_run("5. Дорожная карта первых 7 дней ведения клиента")
    r_h5.font.name = "Calibri"
    r_h5.font.color.rgb = COLOR_PRIMARY
    r_h5.font.size = Pt(15)
    
    plan_days = [
        ("День 1 (Пн): Старт и интеграция", "Клонирование таблицы, включение вебхука в CRM клиента, отправка PDF-руководства собственнику и РОПу."),
        ("День 2 (Вт): Первый тестовый звонок", "Совершение 3–5 реальных звонков. Показать РОПу лист «🎙️ ИИ_Аудит»: звонки появились с баллами и цитатами ошибок."),
        ("День 3 (Ср): Запуск Telegram-бота", "Подключение Telegram-бота РОПу. Первый успешный перехват слитой сделки в режиме реального времени."),
        ("День 4–5 (Чт–Пт): Накопление базы диалогов", "Обработка 100–150 звонков. Автоматическое формирование рейтинга менеджеров и выявление системных сливов."),
        ("День 7 (Пн): Итоговая сессия с собственником", "Разбор листа «💸 Диагностика_Утечек_ОП» и спасенных сделок. Подтверждение окупаемости и подписание годового контракта.")
    ]
    
    for d_title, d_text in plan_days:
        p_dt = doc.add_paragraph()
        p_dt.paragraph_format.space_before = Pt(5)
        p_dt.paragraph_format.space_after = Pt(1)
        r_dt = p_dt.add_run(d_title)
        r_dt.bold = True
        r_dt.font.name = "Calibri"
        r_dt.font.size = Pt(10)
        r_dt.font.color.rgb = COLOR_PRIMARY
        
        p_dx = doc.add_paragraph()
        p_dx.paragraph_format.space_after = Pt(5)
        r_dx = p_dx.add_run(d_text)
        r_dx.font.name = "Calibri"
        r_dx.font.size = Pt(9.5)
        r_dx.font.color.rgb = COLOR_SECONDARY

    # Save DOCX
    docx_path = os.path.abspath("docs/Боевое_Руководство_Основателя_RevOps_OS_V18.docx")
    doc.save(docx_path)
    print(f"✅ DOCX saved successfully: {docx_path}")
    
    # Save to presentation folder
    pres_docx_path = os.path.abspath("presentation/Боевое_Руководство_Основателя_RevOps_OS_V18.docx")
    doc.save(pres_docx_path)
    
    return docx_path

def convert_founder_docx_to_pdf(docx_path):
    temp_docx = os.path.abspath("docs/founder_manual_temp.docx")
    temp_pdf = os.path.abspath("docs/founder_manual_temp.pdf")
    final_pdf = os.path.abspath("docs/Боевое_Руководство_Основателя_RevOps_OS_V18.pdf")
    
    import shutil
    shutil.copyfile(docx_path, temp_docx)
    
    ps_code = """
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open('""" + temp_docx + """')
    $doc.SaveAs([ref]'""" + temp_pdf + """', [ref]17)
    $doc.Close()
    Write-Host "SUCCESS_CONVERT"
} catch {
    Write-Host "ERROR: $_"
} finally {
    $word.Quit()
}
"""
    ps_path = os.path.abspath("scripts/run_founder_pdf_conv.ps1")
    with open(ps_path, "w", encoding="utf-8") as f:
        f.write(ps_code)
        
    res = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_path], capture_output=True, text=True)
    print("STDOUT:", res.stdout.strip())
    
    if os.path.exists(temp_pdf):
        shutil.copyfile(temp_pdf, final_pdf)
        pres_pdf = os.path.abspath("presentation/Боевое_Руководство_Основателя_RevOps_OS_V18.pdf")
        shutil.copyfile(temp_pdf, pres_pdf)
        os.remove(temp_docx)
        os.remove(temp_pdf)
        print(f"✅ Generated PDF successfully: {final_pdf} ({os.path.getsize(final_pdf):,} bytes)!")
        print(f"✅ Copied to presentation: {pres_pdf}!")
    else:
        print("❌ PDF conversion failed.")

if __name__ == "__main__":
    d_path = build_founder_handbook_doc()
    convert_founder_docx_to_pdf(d_path)
