# -*- coding: utf-8 -*-
"""
Generate instruction DOCX, README DOCX/MD, and copy full legal package
to target desktop folder: C:\\Users\\strel\\Desktop\\RevOps Platform\\Договор
"""

import os
import shutil
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=120, bottom=120, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def add_header_banner(doc, title, subtitle):
    section = doc.sections[0]
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.9)
    section.right_margin = Inches(0.9)

    tbl = doc.add_table(rows=1, cols=1)
    tbl.autofit = False
    tbl.columns[0].width = Inches(6.7)
    cell = tbl.cell(0, 0)
    set_cell_background(cell, "0F172A")
    set_cell_margins(cell, top=200, bottom=200, left=240, right=240)

    p1 = cell.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r1 = p1.add_run(f"REVOPS OS  |  {title.upper()}\n")
    r1.font.name = "Calibri"
    r1.font.size = Pt(10)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(56, 189, 248) # Sky blue

    r2 = p1.add_run(title)
    r2.font.name = "Calibri"
    r2.font.size = Pt(16)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(255, 255, 255)

    p2 = cell.add_paragraph()
    r3 = p2.add_run(subtitle)
    r3.font.name = "Calibri"
    r3.font.size = Pt(10.5)
    r3.font.color.rgb = RGBColor(203, 213, 225)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

def add_callout(doc, title, text, bg_hex="F8FAFC", border_hex="3B82F6"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.autofit = False
    tbl.columns[0].width = Inches(6.7)
    cell = tbl.cell(0, 0)
    set_cell_background(cell, bg_hex)
    set_cell_margins(cell, top=140, bottom=140, left=180, right=180)
    
    p = cell.paragraphs[0]
    r_title = p.add_run(f"📌 {title}\n")
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(11)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(15, 23, 42)
    
    r_text = p.add_run(text)
    r_text.font.name = "Calibri"
    r_text.font.size = Pt(10)
    r_text.font.color.rgb = RGBColor(51, 65, 85)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

def generate_instruction_docx(output_path):
    doc = docx.Document()
    add_header_banner(
        doc,
        "Инструкция: Работа как Самозанятый и Переход на ИП",
        "Юридический и налоговый регламент расчетов с B2B-клиентами (ООО и ИП) // RevOps OS"
    )

    # Section 1
    h1 = doc.add_heading(level=1)
    r = h1.add_run("1. Как работать с юрлицами (ООО) прямо сейчас как Самозанятый")
    r.font.name = "Calibri"
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = RGBColor(15, 23, 42)

    p = doc.add_paragraph()
    p.add_run(
        "По российскому законодательству (Федеральный закон от 27.11.2018 № 422-ФЗ) самозанятый гражданин "
        "(плательщик налога на профессиональный доход — НПД) имеет ПОЛНОЕ ЗАКОННОЕ ПРАВО оказывать услуги "
        "юридическим лицам (ООО) и другим ИП по безналичному расчету на банковский счет."
    )

    add_callout(
        doc,
        "Ключевые юридические факты для бухгалтерии Клиента (ООО):",
        "• Клиент НЕ платит за вас НДФЛ 13% и НЕ платит страховые взносы 30% (п. 8 ст. 15 422-ФЗ).\n"
        "• Клиент уменьшает налог на прибыль или УСН на сумму счета (Письмо Минфина РФ № 03-11-11/24004).\n"
        "• Закрывающие документы для клиента: Договор-оферта + Электронный чек ФНС «Мой налог» (Акт по желанию).\n"
        "• Налог для вас: ровно 6% от суммы платежа от юридического лица."
    )

    p_flow = doc.add_paragraph()
    p_flow.add_run("Пошаговый регламент каждой сделки:\n").bold = True
    p_flow.add_run(
        "1. Согласие клиента → отправляете Счет на оплату со ссылкой на оферту (ai-rop.ru/offer.html).\n"
        "2. Клиент оплачивает безналичным переводом на ваш счет в АО «ТИНЬКОФФ БАНК».\n"
        "3. В течение 24 часов выбиваете чек в приложении «Мой налог» (раздел «Юрлицу или ИП», вводите ИНН заказчика).\n"
        "4. Отправляете PDF-чек или ссылку на чек клиенту в Telegram или на Email.\n"
        "5. Клиент прикрепляет чек и оферту к своей бухгалтерии — сделка на 100% чистая."
    )

    # Section 2
    h2 = doc.add_heading(level=1)
    r2 = h2.add_run("2. Когда и почему нужно открывать статус ИП")
    r2.font.name = "Calibri"
    r2.font.size = Pt(13)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(15, 23, 42)

    p2 = doc.add_paragraph()
    p2.add_run(
        "Статус Самозанятого идеален на старте первых 10–20 клиентов. Переходить на статус ИП (УСН 6%) следует в следующих случаях:\n\n"
        "1. Приближение к лимиту 2.4 млн ₽ в год (при выручке от 150 000–200 000 ₽/мес).\n"
        "2. Подключение интернет-эквайринга с автоплатежами (Т-Касса, ЮKassa списывают ежемесячную подписку с бизнес-карт только на ИП/ООО).\n"
        "3. Крупные Enterprise-клиенты (от 89 000 ₽/мес), которым требуется ЭДО (Диадок / СБИС) строго с юридическим лицом."
    )

    # Section 3
    h3 = doc.add_heading(level=1)
    r3 = h3.add_run("3. Как открыть ИП за 3 дня бесплатно и без визита в налоговую")
    r3.font.name = "Calibri"
    r3.font.size = Pt(13)
    r3.font.bold = True
    r3.font.color.rgb = RGBColor(15, 23, 42)

    p3 = doc.add_paragraph()
    p3.add_run(
        "Регистрация проводится через Тинькофф Бизнес (tbank.ru/business/ip-registration/) полностью онлайн:\n"
        "• Банк сам готовит заявление (форма Р21001) и выпускает бесплатную КЭП через Госключ.\n"
        "• Госпошлина 800 ₽ при электронной подаче НЕ взимается.\n"
        "• Режим налогообложения: УСН 6% («Доходы»).\n"
        "• Страховые взносы ИП за себя (53 658 ₽ в год) в полном объеме уменьшают налог УСН вплоть до нуля!\n\n"
        "Рекомендуемые коды ОКВЭД для внесения при регистрации:\n"
        "  — 62.01 (Основной) — Разработка компьютерного программного обеспечения\n"
        "  — 62.02 — Деятельность консультативная и работы в области компьютерных технологий\n"
        "  — 63.11 — Обработка данных, предоставление услуг по размещению информации (SaaS-платформы)\n"
        "  — 62.09 — Деятельность, связанная с использованием вычислительной техники и IT\n"
        "  — 70.22 — Консультирование по вопросам коммерческой деятельности и управления"
    )

    # Section 4
    h4 = doc.add_heading(level=1)
    r4 = h4.add_run("4. Переключение документов на ИП")
    r4.font.name = "Calibri"
    r4.font.size = Pt(13)
    r4.font.bold = True
    r4.font.color.rgb = RGBColor(15, 23, 42)

    p4 = doc.add_paragraph()
    p4.add_run(
        "В наших договорах-офертах уже заранее предусмотрен блок «Вариант Б (Индивидуальный предприниматель)». "
        "После получения листа записи ЕГРИП из ФНС вам потребуется лишь вписать полученный ОГРНИП. "
        "Все действующие клиенты продолжают работу бесшовно, а новые счета выставляются уже от имени ИП Федотов Д."
    )

    doc.save(output_path)
    print(f"✓ Saved: {output_path}")

def generate_readme_navigator(docx_path, md_path):
    # Markdown Navigator
    md_content = """# 📁 Юридический Комплект RevOps OS (ai-rop.ru)
**Основатель:** Дмитрий Федотов (ИНН 731303201073)  
**Директория:** `C:\\Users\\strel\\Desktop\\RevOps Platform\\Договор`  
**Дата актуализации:** Октябрь 2026  

---

## 🎯 Навигатор по файлам: Что и когда использовать

| № | Файл | Для чего нужен | Кому и когда отправлять |
|---|---|---|---|
| **1** | `Договор_Оферта_RevOps_OS_Все_Тарифы.docx` / `.html` | **Генеральная публичная оферта** на все тарифы (Пилот 14 900 ₽, Старт 29 000 ₽, Рост 59 000 ₽, Корпоративный 89 000 ₽). Поддерживает статус Самозанятого и ИП. | Публикуется на сайте (`ai-rop.ru/offer.html`), ссылка указывается в счетах на оплату. |
| **2** | `Договор_Оферта_Пилот_14900.docx` / `.html` | **Специальная оферта на 7-дневный Пилот** с юридически зафиксированной 100% гарантией возврата денег при отсутствии оцифрованных точек роста выручки. | Отправляется клиенту при согласии на пилотное тестирование. |
| **3** | `Бланк_Счета_на_оплату_Пилота_14900.docx` | **Официальный счет на оплату для юрлиц (ООО и ИП)** с банковскими реквизитами Т-Банка, ссылкой на оферту и назначением платежа без НДС. | Выставляется клиенту после подтверждения участия в пилоте. |
| **4** | `Типовой_Акт_Оказанных_Услуг.docx` | **Акт сдачи-приемки оказанных услуг** с перечнем всех этапов (подключение CRM, транскрибация звонков, оцифровка потерь, отчет). Предусмотрен 3-дневный автоматический акцепт. | Подписывается по завершении пилота или календарного месяца подписки (по запросу бухгалтерии клиента). |
| **5** | `152-ФЗ_Комплект_Безопасности_RevOps_AI.docx` | **Регламент безопасности, NDA и соответствие 152-ФЗ** о персональных данных. Описывает алгоритмы обезличивания аудиопотока и шифрование. | Отправляется Службе Безопасности (СБ) и юристам средних и крупных заказчиков. Снимает 100% возражений по безопасности. |
| **6** | `Инструкция_Самозанятый_и_Переход_на_ИП.docx` / `.md` | **Пошаговое руководство для Дмитрия:** как законно принимать безнал от ООО на карту/счет самозанятого, выбивать чеки в «Мой налог», лимиты 2.4 млн ₽, и как за 3 дня бесплатно открыть ИП на УСН 6% через Т-Банк. | Внутренняя инструкция основателя. |
| **7** | `Партнерское_Предложение_Интеграторам_CRM.docx` | **Партнерская оферта и коммерческие условия** для интеграторов amoCRM и Битрикс24 (25–30% рекуррентной комиссии). | Отправляется CRM-интеграторам и IT-агентствам для привлечения клиентов через партнеров. |

---

## ⚡ Шпаргалка для быстрого выставления счета

### 1. Назначение платежа для клиента:
> *«Оплата по Договору-оферте от 02.10.2026 за услуги ИИ-аналитики продаж RevOps OS по Счету № [номер] от [дата]. Без налога (НДС) на основании ст. 15 Федерального закона от 27.11.2018 № 422-ФЗ»*.

### 2. Банковские реквизиты:
* **Получатель:** Федотов Дмитрий (плательщик НПД / Самозанятый)
* **ИНН:** `731303201073`
* **Банк:** АО «ТИНЬКОФФ БАНК»
* **Расчетный счет:** `40802810500003849120`
* **БИК:** `044525974`
* **Корр. счет:** `30101810145250000974`

### 3. Действия после получения денег:
1. Зайти в приложение **«Мой налог»** → «Новая продажа» → Сумма `14 900 ₽`.
2. Наименование: `Услуги аналитики и аудита звонков по счету №...`.
3. Тип покупателя: **«Юридическому лицу или ИП»** → ввести ИНН клиента.
4. Отправить ссылку на фискальный чек клиенту в Telegram / Email.
"""
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(md_content)
    print(f"✓ Saved: {md_path}")

    # DOCX Navigator
    doc = docx.Document()
    add_header_banner(
        doc,
        "Юридический Навигатор и Реестр Документов",
        "Полный юридический комплект RevOps OS для работы с B2B-клиентами (ООО и ИП) // Дмитрий Федотов"
    )

    doc.add_heading("Реестр файлов юридического пакета", level=1)
    
    table_data = [
        ("Файл документа", "Назначение", "Когда использовать"),
        ("Договор_Оферта_RevOps_OS_Все_Тарифы.docx / .html", "Генеральная публичная оферта на все тарифы (Пилот, Старт, Рост, Enterprise). Поддерживает Самозанятого и ИП.", "Опубликована на ai-rop.ru/offer.html, ссылка в счетах на оплату."),
        ("Договор_Оферта_Пилот_14900.docx / .html", "Оферта на быстрый пилот 14 900 ₽ со 100% гарантией возврата при отсутствии оцифрованных точек роста.", "Для быстрого старта с новыми заказчиками с гарантией результата."),
        ("Бланк_Счета_на_оплату_Пилота_14900.docx", "Официальный счет на оплату для юрлиц с реквизитами Т-Банка и назначением платежа без НДС.", "Выставляется клиенту после согласия на пилот."),
        ("Типовой_Акт_Оказанных_Услуг.docx", "Акт сдачи-приемки с перечнем всех оказанных услуг и 3-дневным молчаливым акцептом.", "Закрывающий документ для бухгалтерии заказчика."),
        ("152-ФЗ_Комплект_Безопасности_RevOps_AI.docx", "Положение о безопасности, защите персональных данных, регламент обезличивания аудио и NDA.", "Для юристов и Службы Безопасности корпоративных клиентов."),
        ("Инструкция_Самозанятый_и_Переход_на_ИП.docx", "Мануал для Дмитрия: как работать как самозанятый с ООО, выбивать чеки, и как за 3 дня открыть ИП на УСН 6%.", "Внутренняя настольная инструкция основателя."),
        ("Партнерское_Предложение_Интеграторам_CRM.docx", "Коммерческие и юридические условия партнерства с интеграторами amoCRM / Битрикс24 (25–30% комиссия).", "Для привлечения партнеров-интеграторов.")
    ]

    tbl = doc.add_table(rows=len(table_data), cols=3)
    tbl.autofit = False
    col_widths = [Inches(2.2), Inches(2.6), Inches(1.9)]

    for row_idx, row in enumerate(table_data):
        for col_idx, text in enumerate(row):
            cell = tbl.cell(row_idx, col_idx)
            cell.width = col_widths[col_idx]
            p = cell.paragraphs[0]
            r = p.add_run(text)
            r.font.name = "Calibri"
            if row_idx == 0:
                set_cell_background(cell, "0F172A")
                set_cell_margins(cell, top=140, bottom=140, left=120, right=120)
                r.font.size = Pt(9.5)
                r.font.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)
            else:
                bg = "F8FAFC" if row_idx % 2 == 1 else "FFFFFF"
                set_cell_background(cell, bg)
                set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
                r.font.size = Pt(9)
                r.font.color.rgb = RGBColor(30, 41, 59)
                if col_idx == 0:
                    r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    add_callout(
        doc,
        "Шпаргалка: Назначение платежа для счета от ООО:",
        "«Оплата по Договору-оферте от 02.10.2026 за услуги ИИ-аналитики продаж RevOps OS по Счету № [номер] от [дата]. Без налога (НДС) на основании ст. 15 Федерального закона от 27.11.2018 № 422-ФЗ»"
    )

    doc.save(docx_path)
    print(f"✓ Saved: {docx_path}")

def copy_all_to_desktop(desktop_dir, repo_docs_dir):
    os.makedirs(desktop_dir, exist_ok=True)
    
    files_to_copy = [
        "Договор_Оферта_RevOps_OS_Все_Тарифы.docx",
        "Договор_Оферта_RevOps_OS_Все_Тарифы.html",
        "Договор_Оферта_Пилот_14900.docx",
        "Договор_Оферта_Пилот_14900.html",
        "Бланк_Счета_на_оплату_Пилота_14900.docx",
        "Типовой_Акт_Оказанных_Услуг.docx",
        "152-ФЗ_Комплект_Безопасности_RevOps_AI.docx",
        "Инструкция_Самозанятый_и_Переход_на_ИП.docx",
        "Инструкция_Самозанятый_и_Переход_на_ИП.md",
        "Партнерское_Предложение_Интеграторам_CRM.docx",
        "README_Навигатор_Юридический_Комплект.docx",
        "README_Навигатор_Юридический_Комплект.md",
        "Боевой_Протокол_Клиент_За_7_Дней.docx",
        "Дорожная_Карта_Пилота_Для_Клиента_7_Дней.docx"
    ]

    copied = []
    for fname in files_to_copy:
        src = os.path.join(repo_docs_dir, fname)
        if os.path.exists(src):
            dst = os.path.join(desktop_dir, fname)
            shutil.copy2(src, dst)
            copied.append(fname)
        else:
            print(f"⚠ Warning: not found in docs: {src}")

    print(f"\nSuccessfully copied {len(copied)} files to:\n{desktop_dir}")
    for c in copied:
        size_kb = os.path.getsize(os.path.join(desktop_dir, c)) / 1024
        print(f"  • {c} ({size_kb:.1f} KB)")

if __name__ == "__main__":
    repo_docs = r"C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\docs"
    desktop_target = r"C:\Users\strel\Desktop\RevOps Platform\Договор"

    # 1. Generate DOCX instruction
    instr_docx = os.path.join(repo_docs, "Инструкция_Самозанятый_и_Переход_на_ИП.docx")
    generate_instruction_docx(instr_docx)

    # 2. Generate README/Navigator DOCX & MD
    readme_docx = os.path.join(repo_docs, "README_Навигатор_Юридический_Комплект.docx")
    readme_md = os.path.join(repo_docs, "README_Навигатор_Юридический_Комплект.md")
    generate_readme_navigator(readme_docx, readme_md)

    # 3. Copy everything to target desktop folder
    copy_all_to_desktop(desktop_target, repo_docs)
