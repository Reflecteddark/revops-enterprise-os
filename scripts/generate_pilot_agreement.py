import datetime
from pathlib import Path
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

def create_docx(output_path: Path):
    doc = docx.Document()

    # Set 0.7 inch margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("ДОГОВОР-ОФЕРТА № PLT-2026/01\nна проведение 7-дневного пилотного внедрения сервиса «RevOps OS»")
    run_title.font.name = "Arial"
    run_title.font.size = Pt(13)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(15, 23, 42)

    # City and Date
    p_date = doc.add_paragraph()
    p_date.paragraph_format.space_after = Pt(10)
    run_city = p_date.add_run("г. Москва / Дистанционно\t\t\t\t\t«____» ____________ 2026 г.")
    run_city.font.name = "Arial"
    run_city.font.size = Pt(9.5)
    run_city.font.color.rgb = RGBColor(100, 116, 139)

    # Preamble
    p_preamble = doc.add_paragraph()
    p_preamble.paragraph_format.space_after = Pt(8)
    p_preamble.paragraph_format.line_spacing = 1.15
    r_pre = p_preamble.add_run(
        "Индивидуальный предприниматель Федотов Дмитрий, именуемый в дальнейшем «Исполнитель» (сервис ai-rop.ru), "
        "публикует настоящую публичную Оферту на оказание услуг по тестированию и настройке программного комплекса ИИ-аналитики продаж "
        "в адрес любого юридического лица или индивидуального предпринимателя (далее — «Заказчик»)."
    )
    r_pre.font.name = "Arial"
    r_pre.font.size = Pt(9.5)

    def add_section(num_title: str, text: str):
        p_sec = doc.add_paragraph()
        p_sec.paragraph_format.space_before = Pt(8)
        p_sec.paragraph_format.space_after = Pt(3)
        r_sec = p_sec.add_run(num_title)
        r_sec.font.name = "Arial"
        r_sec.font.size = Pt(10)
        r_sec.font.bold = True
        r_sec.font.color.rgb = RGBColor(30, 41, 59)

        p_body = doc.add_paragraph()
        p_body.paragraph_format.space_after = Pt(6)
        p_body.paragraph_format.line_spacing = 1.15
        r_b = p_body.add_run(text)
        r_b.font.name = "Arial"
        r_b.font.size = Pt(9.5)
        r_b.font.color.rgb = RGBColor(51, 65, 85)

    add_section(
        "1. ПРЕДМЕТ ДОГОВОРА",
        "1.1. Исполнитель обязуется оказать Заказчику комплекс консультационных и технологических услуг по проведению 7-дневного пилотного внедрения системы ИИ-супервизии звонков «RevOps OS» в отделе продаж Заказчика, а Заказчик обязуется принять и оплатить оказанные услуги.\n"
        "1.2. В состав пилотного спринта входит:\n"
        "  • Быстрое подключение сервиса к действующей CRM Заказчика (amoCRM или Битрикс24) в течение 1 рабочего дня;\n"
        "  • 100% оцифровка и речевой аудит входящих и исходящих звонков менеджеров по 13 корпоративным критериям B2B;\n"
        "  • Настройка автоматических Telegram-алертов РОПу Заказчика (за 60 секунд при сливе сделки на возражении «Дорого» или потере следующего шага);\n"
        "  • Ежедневный утренний дайджест в 09:00 со списком приоритетных сделок под угрозой срыва;\n"
        "  • Итоговая PDF-карта скрытых сливов воронки с точной оцифровкой упущенной выручки и регламентом исправления."
    )

    add_section(
        "2. СТОИМОСТЬ УСЛУГ И ПОРЯДОК ОПЛАТЫ",
        "2.1. Фиксированная стоимость услуг по настоящему Договору за полный 7-дневный пилотный спринт составляет 14 900 (четырнадцать тысяч девятьсот) рублей 00 копеек. НДС не облагается в связи с применением Исполнителем УСН (п. 2 ст. 346.11 НК РФ).\n"
        "2.2. Оплата производится Заказчиком в форме 100% предоплаты на расчетный счет Исполнителя на основании выставленного счета или через защищенный корпоративный интернет-эквайринг.\n"
        "2.3. Никаких скрытых платежей, платных интеграций и абонентских плат за период пилота не взимается."
    )

    add_section(
        "3. БЕЗУСЛОВНАЯ 100% ГАРАНТИЯ ВОЗВРАТА СРЕДСТВ (MONEY-BACK GUARANTEE)",
        "3.1. Исполнитель гарантирует качество аналитики и окупаемость системы. Если по итогам 7 дней пилотного спринта система не обнаружит точек сливов выручки или Заказчик сочтет результаты пилота неудовлетворительными по любой причине, Исполнитель гарантирует 100% возврат всей оплаченной суммы (14 900 рублей).\n"
        "3.2. Возврат средств производится в течение 3 (трех) банковских дней на расчетный счет Заказчика на основании простого уведомления в свободной форме в Telegram (@dm1918) или по электронной почте, без удержаний, штрафов и бюрократических процедур."
    )

    add_section(
        "4. КОНФИДЕНЦИАЛЬНОСТЬ И БЕЗОПАСНОСТЬ ДАННЫХ (152-ФЗ РФ)",
        "4.1. Стороны обязуются соблюдать строгий режим конфиденциальности в отношении клиентской базы, аудиозаписей переговоров и коммерческих показателей Заказчика.\n"
        "4.2. Обработка данных осуществляется строго в соответствии с требованиями Федерального закона № 152-ФЗ «О персональных данных». Все серверные мощности и базы данных Исполнителя расположены исключительно на территории Российской Федерации (Selectel, г. Москва/г. Санкт-Петербург).\n"
        "4.3. Аудиозаписи и транскрипты автоматически проходят стадию токенизации и деперсонализации (ФИО клиентов, банковские реквизиты и номера телефонов маскируются криптографическими хэшами)."
    )

    add_section(
        "5. СРОК ДЕЙСТВИЯ И СДАЧА-ПРИЕМКА УСЛУГ",
        "5.1. Договор вступает в силу с момента поступления оплаты на расчетный счет Исполнителя и действует до полного исполнения Сторонами своих обязательств.\n"
        "5.2. Услуги считаются оказанными в полном объеме с момента передачи Заказчику итогового аналитического отчета за 7 дней пилота. Акт сдачи-приемки направляется в электронном виде (через Диадок/СБИС или скан-копией)."
    )

    # Signatures block
    p_sig_title = doc.add_paragraph()
    p_sig_title.paragraph_format.space_before = Pt(12)
    p_sig_title.paragraph_format.space_after = Pt(4)
    r_st = p_sig_title.add_run("6. РЕКВИЗИТЫ И ПОДПИСИ СТОРОН")
    r_st.font.name = "Arial"
    r_st.font.size = Pt(10)
    r_st.font.bold = True

    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    cell_isp = table.cell(0, 0)
    cell_zak = table.cell(0, 1)
    cell_isp.width = Inches(3.5)
    cell_zak.width = Inches(3.5)

    p_isp = cell_isp.paragraphs[0]
    p_isp.paragraph_format.space_after = Pt(2)
    r_ih = p_isp.add_run("ИСПОЛНИТЕЛЬ:\n")
    r_ih.font.bold = True
    r_ih.font.size = Pt(9)
    r_ib = p_isp.add_run(
        "ИП Федотов Дмитрий\n"
        "Сервис: RevOps OS (ai-rop.ru)\n"
        "ИНН: 731303201073\n"
        "Расчетный счет: 40802810500003849120\n"
        "Банк: АО «ТИНЬКОФФ БАНК»\n"
        "БИК: 044525974\n"
        "К/с: 30101810145250000974\n"
        "Telegram / MAX: @dm1918\n"
        "Email: founder@ai-rop.ru\n\n"
        "Подпись: _________________ / Федотов Д. /\n"
        "М.П."
    )
    r_ib.font.size = Pt(8.5)

    p_zak = cell_zak.paragraphs[0]
    p_zak.paragraph_format.space_after = Pt(2)
    r_zh = p_zak.add_run("ЗАКАЗЧИК:\n")
    r_zh.font.bold = True
    r_zh.font.size = Pt(9)
    r_zb = p_zak.add_run(
        "Наименование: _______________________\n"
        "ИНН / КПП: ___________________________\n"
        "ОГРН: ________________________________\n"
        "Юр. адрес: ___________________________\n"
        "Расчетный счет: ______________________\n"
        "Банк: ________________________________\n"
        "БИК: _________________________________\n"
        "Контактное лицо: _____________________\n"
        "Telegram / Тел: ______________________\n\n"
        "Подпись: _________________ / _________ /\n"
        "М.П."
    )
    r_zb.font.size = Pt(8.5)

    doc.save(str(output_path))
    print(f"✓ DOCX agreement generated: {output_path}")


def create_html(output_path: Path):
    html = """<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <title>Договор-Оферта на 7-дневный Пилот RevOps OS (14 900 ₽)</title>
    <style>
        body { font-family: 'Helvetica Neue', Arial, sans-serif; background: #f8fafc; color: #1e293b; padding: 30px; font-size: 13.5px; line-height: 1.5; }
        .page { max-width: 820px; margin: 0 auto; background: #fff; padding: 40px 50px; border-radius: 8px; box-shadow: 0 4px 20px rgba(0,0,0,0.06); border: 1px solid #e2e8f0; }
        h1 { font-size: 16px; text-align: center; text-transform: uppercase; margin-bottom: 5px; color: #0f172a; }
        .subhead { text-align: center; font-size: 13px; color: #64748b; margin-bottom: 25px; }
        .dates { display: flex; justify-content: space-between; font-size: 12px; color: #64748b; margin-bottom: 20px; border-bottom: 1px solid #e2e8f0; padding-bottom: 10px; }
        h2 { font-size: 13.5px; font-weight: bold; margin-top: 20px; margin-bottom: 6px; color: #1e293b; border-left: 3px solid #059669; padding-left: 8px; }
        p { margin: 6px 0; text-align: justify; }
        ul { margin: 6px 0 10px 20px; padding: 0; }
        li { margin-bottom: 4px; }
        .badge { background: #ecfdf5; color: #047857; font-weight: bold; padding: 2px 6px; border-radius: 4px; border: 1px solid #a7f3d0; }
        .sig-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 30px; margin-top: 30px; border-top: 2px solid #e2e8f0; padding-top: 20px; }
        .sig-box { font-size: 12px; }
        .sig-box h3 { font-size: 13px; margin-bottom: 8px; color: #0f172a; }
        @media print {
            body { background: #fff; padding: 0; }
            .page { box-shadow: none; border: none; padding: 0; }
        }
    </style>
</head>
<body>
<div class="page">
    <h1>Договор-Оферта № PLT-2026/01</h1>
    <div class="subhead">на проведение 7-дневного пилотного внедрения сервиса «RevOps OS»</div>
    
    <div class="dates">
        <span>г. Москва / Дистанционно</span>
        <span>Действительно с: 02.10.2026 г.</span>
    </div>

    <p><strong>Индивидуальный предприниматель Федотов Дмитрий</strong>, именуемый в дальнейшем «Исполнитель» (основатель сервиса <a href="https://ai-rop.ru">ai-rop.ru</a>), публикует настоящую публичную Оферту на оказание услуг по тестированию и настройке программного комплекса ИИ-аналитики продаж в адрес любого юридического лица или индивидуального предпринимателя (далее — «Заказчик»).</p>

    <h2>1. ПРЕДМЕТ ДОГОВОРА</h2>
    <p>1.1. Исполнитель обязуется оказать Заказчику комплекс консультационных и технологических услуг по проведению 7-дневного пилотного спринта внедрения системы ИИ-супервизии звонков «RevOps OS» в отделе продаж Заказчика.</p>
    <p>1.2. В состав пилотного спринта входит:</p>
    <ul>
        <li>Подключение сервиса к действующей CRM Заказчика (amoCRM или Битрикс24) за 1 рабочий день;</li>
        <li>100% оцифровка и речевой аудит входящих и исходящих звонков менеджеров по 13 корпоративным критериям B2B;</li>
        <li>Настройка автоматических Telegram-алертов РОПу Заказчика (за 60 секунд при сливе сделки на возражении «Дорого» или потере следующего шага);</li>
        <li>Ежедневный утренний дайджест в 09:00 со списком приоритетных сделок под угрозой срыва;</li>
        <li>Итоговая PDF-карта скрытых сливов воронки с точной оцифровкой упущенной выручки и регламентом исправления.</li>
    </ul>

    <h2>2. СТОИМОСТЬ УСЛУГ И ПОРЯДОК ОПЛАТЫ</h2>
    <p>2.1. Стоимость услуг за полный 7-дневный пилотный спринт составляет <span class="badge">14 900 (четырнадцать тысяч девятьсот) рублей</span>. НДС не облагается (УСН, п. 2 ст. 346.11 НК РФ).</p>
    <p>2.2. Оплата производится Заказчиком в форме 100% предоплаты по безналичному расчету на расчетный счет Исполнителя или корпоративной бизнес-картой.</p>

    <h2>3. БЕЗУСЛОВНАЯ 100% ГАРАНТИЯ ВОЗВРАТА СРЕДСТВ (MONEY-BACK GUARANTEE)</h2>
    <p>3.1. Исполнитель гарантирует окупаемость пилота. Если по итогам 7 дней система не выявит скрытых сливов выручки или Заказчик сочтет результаты пилота неудовлетворительными, <strong>Исполнитель возвращает 100% уплаченной суммы (14 900 ₽) в течение 3 банковских дней</strong>.</p>
    <p>3.2. Возврат производится по простому запросу в свободной форме в Telegram (@dm1918) или на email, без штрафов и удержаний.</p>

    <h2>4. КОНФИДЕНЦИАЛЬНОСТЬ И БЕЗОПАСНОСТЬ (152-ФЗ РФ)</h2>
    <p>4.1. Стороны соблюдают строгий режим конфиденциальности клиентской базы и записей звонков.</p>
    <p>4.2. Обработка данных ведется строго по 152-ФЗ РФ на серверах в Российской Федерации (Selectel, Москва/СПб). Персональные данные (ФИО, телефоны) автоматически токенизируются и деперсонализируются.</p>

    <h2>5. РЕКВИЗИТЫ СТОРОН</h2>
    <div class="sig-grid">
        <div class="sig-box">
            <h3>ИСПОЛНИТЕЛЬ:</h3>
            <p><strong>ИП Федотов Дмитрий</strong><br>
            Сервис RevOps OS (ai-rop.ru)<br>
            ИНН: 731303201073<br>
            Р/с: 40802810500003849120<br>
            Банк: АО «ТИНЬКОФФ БАНК»<br>
            БИК: 044525974<br>
            К/с: 30101810145250000974<br>
            Telegram / MAX: @dm1918<br>
            Email: founder@ai-rop.ru<br><br>
            Подпись: _______________ / Федотов Д. /
            </p>
        </div>
        <div class="sig-box">
            <h3>ЗАКАЗЧИК:</h3>
            <p>
            Организация: _________________________________<br>
            ИНН / КПП: ___________________________________<br>
            Адрес: _______________________________________<br>
            Р/с: _________________________________________<br>
            Банк: ________________________________________<br>
            БИК: _________________________________________<br>
            Контактное лицо: _____________________________<br>
            Telegram / Тел: ______________________________<br><br>
            Подпись: _______________ / ___________________ /
            </p>
        </div>
    </div>
</div>
</body>
</html>"""
    output_path.write_text(html, encoding="utf-8")
    print(f"✓ HTML agreement generated: {output_path}")

def main():
    docs_dir = Path("docs")
    docs_dir.mkdir(exist_ok=True)
    create_docx(docs_dir / "Договор_Оферта_Пилот_14900.docx")
    create_html(docs_dir / "Договор_Оферта_Пилот_14900.html")

if __name__ == "__main__":
    main()
