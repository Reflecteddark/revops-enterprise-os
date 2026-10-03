#!/usr/bin/env python3
"""
Automated B2B Billing & Accounting Document Generator for Self-Employed (НПД)
Generates:
1. Счет на оплату (Invoice) - HTML + DOCX
2. Акт сдачи-приемки услуг (Acceptance Act) - HTML + DOCX
Compliant with Federal Law No. 422-FZ and Russian accounting standards (402-FZ).
"""

import os
import sys
from datetime import datetime
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

sys.stdout.reconfigure(encoding='utf-8')

# Default Provider Data (Самозанятый)
PROVIDER = {
    "name": "Федотов Дмитрий Валерьевич",
    "status": "Плательщик налога на профессиональный доход (Самозанятый)",
    "inn": "731303201073",
    "bank_name": "Филиал № 6318 Банка ВТБ (ПАО)",
    "bank_address": "443030, г. Самара, ул. Спортивная, д.30",
    "bik": "043601968",
    "bank_inn": "7702070139",
    "bank_kpp": "631543002",
    "corr_account": "30101810422023601968",
    "account": "40817810635184001277",
    "email": "info@ai-rop.ru",
    "telegram": "@dm1918",
    "site": "https://ai-rop.ru",
    "offer_url": "https://ai-rop.ru/offer.html"
}

def number_to_words_ru(amount):
    """Simple Russian currency spelling for standard SaaS amounts."""
    amounts = {
        14900: "Четырнадцать тысяч девятьсот рублей 00 копеек",
        20000: "Двадцать тысяч рублей 00 копеек",
        35000: "Тридцать пять тысяч рублей 00 копеек",
        40000: "Сорок тысяч рублей 00 копеек",
        70000: "Семьдесят тысяч рублей 00 копеек",
        80000: "Восемьдесят тысяч рублей 00 копеек",
        120000: "Сто двадцать тысяч рублей 00 копеек",
    }
    return amounts.get(amount, f"{amount:,} рублей 00 копеек".replace(',', ' '))

def create_invoice_html(invoice_no, date_str, customer, item_name, amount, output_path):
    words = number_to_words_ru(amount)
    html = f"""<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <title>Счет на оплату № {invoice_no} от {date_str}</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            font-size: 13px;
            line-height: 1.4;
            color: #111827;
            background: #fff;
            margin: 0;
            padding: 30px;
        }}
        .invoice-box {{
            max-width: 800px;
            margin: 0 auto;
        }}
        .notice {{
            font-size: 11px;
            color: #4b5563;
            margin-bottom: 15px;
            border-bottom: 1px solid #e5e7eb;
            padding-bottom: 8px;
        }}
        table.bank-table {{
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 25px;
        }}
        table.bank-table td {{
            border: 1px solid #000;
            padding: 6px 8px;
            font-size: 12px;
            vertical-align: top;
        }}
        .header-title {{
            font-size: 18px;
            font-weight: bold;
            border-bottom: 2px solid #000;
            padding-bottom: 10px;
            margin-bottom: 20px;
        }}
        .meta-table {{
            width: 100%;
            margin-bottom: 20px;
        }}
        .meta-table td {{
            padding: 4px 0;
            vertical-align: top;
        }}
        .meta-label {{
            width: 120px;
            color: #4b5563;
        }}
        .items-table {{
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
        }}
        .items-table th, .items-table td {{
            border: 1px solid #000;
            padding: 8px;
            font-size: 12px;
        }}
        .items-table th {{
            background: #f3f4f6;
            text-align: center;
        }}
        .total-block {{
            text-align: right;
            margin-bottom: 20px;
            font-size: 13px;
        }}
        .words-block {{
            margin-bottom: 30px;
            font-size: 12px;
        }}
        .signatures {{
            margin-top: 40px;
            border-top: 1px solid #d1d5db;
            padding-top: 20px;
            display: flex;
            justify-content: space-between;
        }}
        .sign-col {{
            width: 45%;
        }}
        .sign-line {{
            margin-top: 35px;
            border-bottom: 1px solid #000;
            display: flex;
            justify-content: space-between;
        }}
        .sign-hint {{
            font-size: 10px;
            color: #6b7280;
            text-align: center;
            margin-top: 2px;
        }}
        .qr-badge {{
            display: inline-block;
            background: #ecfdf5;
            border: 1px solid #a7f3d0;
            color: #065f46;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: bold;
            margin-top: 10px;
        }}
        @media print {{
            body {{ padding: 0; }}
            .no-print {{ display: none; }}
        }}
    </style>
</head>
<body>
    <div class="invoice-box">
        <div class="notice">
            Внимание! Оплата данного счета означает полное и безоговорочное согласие с условиями Публичной оферты на сайте <a href="{PROVIDER['offer_url']}" target="_blank">{PROVIDER['offer_url']}</a>. В назначении платежа обязательно указывайте номер счета.
        </div>

        <!-- Банковские реквизиты -->
        <table class="bank-table">
            <tr>
                <td colspan="2" style="width: 60%;">
                    {PROVIDER['bank_name']}<br>
                    <span style="font-size: 10px; color: #4b5563;">Банк получателя</span>
                </td>
                <td style="width: 15%;">БИК</td>
                <td style="width: 25%; font-weight: bold;">{PROVIDER['bik']}</td>
            </tr>
            <tr>
                <td colspan="2"></td>
                <td>Сч. №</td>
                <td style="font-weight: bold;">{PROVIDER['corr_account']}</td>
            </tr>
            <tr>
                <td style="width: 25%;">ИНН {PROVIDER['inn']}</td>
                <td style="width: 35%;">КПП —</td>
                <td rowspan="2">Сч. №</td>
                <td rowspan="2" style="font-weight: bold; font-size: 14px; vertical-align: middle;">{PROVIDER['account']}</td>
            </tr>
            <tr>
                <td colspan="2">
                    {PROVIDER['status']} {PROVIDER['name']}<br>
                    <span style="font-size: 10px; color: #4b5563;">Получатель</span>
                </td>
            </tr>
        </table>

        <!-- Заголовок -->
        <div class="header-title">
            Счет на оплату № {invoice_no} от {date_str} г.
        </div>

        <!-- Стороны -->
        <table class="meta-table">
            <tr>
                <td class="meta-label">Поставщик (Исполнитель):</td>
                <td>
                    <strong>{PROVIDER['name']}</strong>, ИНН {PROVIDER['inn']}, {PROVIDER['status']}, Email: {PROVIDER['email']}, Сайт: {PROVIDER['site']}
                </td>
            </tr>
            <tr>
                <td class="meta-label">Покупатель (Заказчик):</td>
                <td>
                    <strong>{customer['name']}</strong>, ИНН {customer['inn']}{', КПП ' + customer.get('kpp', '') if customer.get('kpp') else ''}{', Адрес: ' + customer.get('address', '') if customer.get('address') else ''}
                </td>
            </tr>
            <tr>
                <td class="meta-label">Основание:</td>
                <td>
                    Публичный договор-оферта на предоставление сервиса RevOps OS (ai-rop.ru/offer.html)
                </td>
            </tr>
        </table>

        <!-- Табличная часть -->
        <table class="items-table">
            <thead>
                <tr>
                    <th style="width: 30px;">№</th>
                    <th>Товары (работы, услуги)</th>
                    <th style="width: 50px;">Кол-во</th>
                    <th style="width: 40px;">Ед.</th>
                    <th style="width: 90px;">Цена</th>
                    <th style="width: 100px;">Сумма</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td style="text-align: center;">1</td>
                    <td>{item_name}</td>
                    <td style="text-align: center;">1</td>
                    <td style="text-align: center;">усл.</td>
                    <td style="text-align: right;">{amount:,.2f} ₽</td>
                    <td style="text-align: right; font-weight: bold;">{amount:,.2f} ₽</td>
                </tr>
            </tbody>
        </table>

        <!-- Итоги -->
        <div class="total-block">
            <strong>Итого:</strong> {amount:,.2f} ₽<br>
            <strong>Без налога (НДС):</strong> —<br>
            <span style="font-size: 15px;"><strong>Всего к оплате:</strong> {amount:,.2f} ₽</span>
        </div>

        <div class="words-block">
            Всего наименований 1, на сумму <strong>{amount:,.2f} ₽</strong><br>
            <strong>Сумма прописью:</strong> {words}.<br>
            <span style="font-size: 11px; color: #4b5563;">НДС не облагается в связи с применением Исполнителем специального налогового режима «Налог на профессиональный доход» (Федеральный закон от 27.11.2018 № 422-ФЗ).</span>
        </div>

        <div style="background: #f8fafc; border: 1px dashed #cbd5e1; border-radius: 6px; padding: 10px; font-size: 11px; margin-bottom: 25px;">
            <strong>Назначение платежа для платежного поручения:</strong><br>
            <code style="background: #e2e8f0; padding: 2px 4px; border-radius: 3px; font-family: monospace;">Оплата по Счету № {invoice_no} от {date_str} за предоставление доступа к сервису ai-rop.ru по договору-оферте. НДС не облагается (НПД 422-ФЗ).</code>
        </div>

        <!-- Подписи -->
        <table style="width: 100%; margin-top: 30px;">
            <tr>
                <td style="width: 50%;">
                    Исполнитель: <strong>{PROVIDER['name']}</strong>
                    <div style="margin-top: 30px; border-bottom: 1px solid #000; width: 80%;"></div>
                    <div style="font-size: 9px; color: #6b7280; width: 80%; text-align: center;">(подпись)</div>
                </td>
                <td style="width: 50%; vertical-align: top; text-align: right;">
                    <div class="qr-badge">
                        🔒 Защищенный расчетный контур 422-ФЗ РФ
                    </div>
                </td>
            </tr>
        </table>
    </div>
</body>
</html>"""
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  [+] Invoice HTML created: {output_path}")

def create_act_html(act_no, date_str, customer, item_name, amount, output_path):
    words = number_to_words_ru(amount)
    html = f"""<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <title>Акт сдачи-приемки услуг № {act_no} от {date_str}</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            font-size: 13px;
            line-height: 1.4;
            color: #111827;
            background: #fff;
            margin: 0;
            padding: 30px;
        }}
        .act-box {{
            max-width: 800px;
            margin: 0 auto;
        }}
        .header-title {{
            font-size: 18px;
            font-weight: bold;
            border-bottom: 2px solid #000;
            padding-bottom: 10px;
            margin-bottom: 20px;
        }}
        .meta-table {{
            width: 100%;
            margin-bottom: 20px;
        }}
        .meta-table td {{
            padding: 4px 0;
            vertical-align: top;
        }}
        .meta-label {{
            width: 120px;
            color: #4b5563;
        }}
        .items-table {{
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
        }}
        .items-table th, .items-table td {{
            border: 1px solid #000;
            padding: 8px;
            font-size: 12px;
        }}
        .items-table th {{
            background: #f3f4f6;
            text-align: center;
        }}
        .total-block {{
            text-align: right;
            margin-bottom: 20px;
            font-size: 13px;
        }}
        .words-block {{
            margin-bottom: 25px;
            font-size: 12px;
        }}
        .act-clause {{
            font-size: 12px;
            line-height: 1.5;
            margin-bottom: 30px;
            background: #f8fafc;
            padding: 12px;
            border-left: 3px solid #10b981;
            border-radius: 4px;
        }}
        .signatures {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 40px;
        }}
        .signatures td {{
            width: 50%;
            vertical-align: top;
            padding: 0 10px;
        }}
        .sign-line {{
            margin-top: 40px;
            border-bottom: 1px solid #000;
            width: 90%;
        }}
        .sign-hint {{
            font-size: 9px;
            color: #6b7280;
            width: 90%;
            text-align: center;
            margin-top: 2px;
        }}
        @media print {{
            body {{ padding: 0; }}
        }}
    </style>
</head>
<body>
    <div class="act-box">
        <!-- Заголовок -->
        <div class="header-title">
            Акт сдачи-приемки оказанных услуг № {act_no} от {date_str} г.
        </div>

        <!-- Стороны -->
        <table class="meta-table">
            <tr>
                <td class="meta-label">Исполнитель:</td>
                <td>
                    <strong>{PROVIDER['name']}</strong>, ИНН {PROVIDER['inn']}, {PROVIDER['status']}, Email: {PROVIDER['email']}, Сайт: {PROVIDER['site']}
                </td>
            </tr>
            <tr>
                <td class="meta-label">Заказчик:</td>
                <td>
                    <strong>{customer['name']}</strong>, ИНН {customer['inn']}{', КПП ' + customer.get('kpp', '') if customer.get('kpp') else ''}{', Адрес: ' + customer.get('address', '') if customer.get('address') else ''}
                </td>
            </tr>
            <tr>
                <td class="meta-label">Основание:</td>
                <td>
                    Публичный договор-оферта на предоставление сервиса RevOps OS (ai-rop.ru/offer.html)
                </td>
            </tr>
        </table>

        <!-- Табличная часть -->
        <table class="items-table">
            <thead>
                <tr>
                    <th style="width: 30px;">№</th>
                    <th>Наименование оказанных услуг / объем работ</th>
                    <th style="width: 50px;">Кол-во</th>
                    <th style="width: 40px;">Ед.</th>
                    <th style="width: 90px;">Цена</th>
                    <th style="width: 100px;">Сумма</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td style="text-align: center;">1</td>
                    <td>{item_name}</td>
                    <td style="text-align: center;">1</td>
                    <td style="text-align: center;">усл.</td>
                    <td style="text-align: right;">{amount:,.2f} ₽</td>
                    <td style="text-align: right; font-weight: bold;">{amount:,.2f} ₽</td>
                </tr>
            </tbody>
        </table>

        <!-- Итоги -->
        <div class="total-block">
            <strong>Итого:</strong> {amount:,.2f} ₽<br>
            <strong>Без налога (НДС):</strong> —<br>
            <span style="font-size: 15px;"><strong>Всего оказано услуг на сумму:</strong> {amount:,.2f} ₽</span>
        </div>

        <div class="words-block">
            Всего оказано услуг 1, на сумму <strong>{amount:,.2f} ₽</strong><br>
            <strong>Сумма прописью:</strong> {words}.<br>
            <span style="font-size: 11px; color: #4b5563;">НДС не облагается в связи с применением Исполнителем специального налогового режима «Налог на профессиональный доход» (Федеральный закон от 27.11.2018 № 422-ФЗ).</span>
        </div>

        <!-- Утверждение приема -->
        <div class="act-clause">
            Вышеперечисленные услуги оказаны в полном объеме, своевременно и надлежащего качества. Стороны претензий по объему, качеству и срокам оказания услуг друг к другу не имеют.<br>
            К настоящему Акту прилагается электронный фискальный чек из мобильного приложения ФНС России «Мой налог» в соответствии с п. 1 ст. 15 Федерального закона от 27.11.2018 № 422-ФЗ.
        </div>

        <!-- Подписи Сторон -->
        <table class="signatures">
            <tr>
                <td>
                    <strong>ИСПОЛНИТЕЛЬ:</strong><br>
                    {PROVIDER['name']}<br>
                    (Самозанятый, ИНН {PROVIDER['inn']})
                    <div class="sign-line"></div>
                    <div class="sign-hint">/ {PROVIDER['name']} /</div>
                </td>
                <td>
                    <strong>ЗАКАЗЧИК:</strong><br>
                    {customer['name']}<br>
                    ИНН {customer['inn']}
                    <div class="sign-line"></div>
                    <div class="sign-hint">/ М.П. (при наличии) /</div>
                </td>
            </tr>
        </table>
    </div>
</body>
</html>"""
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  [+] Act HTML created: {output_path}")

def create_invoice_docx(invoice_no, date_str, customer, item_name, amount, output_path):
    doc = docx.Document()
    
    # Page Margins: 1.5 cm
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    p_notice = doc.add_paragraph()
    p_notice.add_run(f"Внимание! Оплата данного счета означает согласие с Публичной офертой на сайте {PROVIDER['offer_url']}.").font.size = Pt(8.5)
    p_notice.runs[0].font.color.rgb = RGBColor(100, 116, 139)

    # Bank Table
    bank_t = doc.add_table(rows=4, cols=4)
    bank_t.style = 'Table Grid'
    
    cells = bank_t.rows[0].cells
    cells[0].merge(cells[1])
    cells[0].text = f"{PROVIDER['bank_name']}\nБанк получателя"
    cells[2].text = "БИК"
    cells[3].text = PROVIDER['bik']
    
    cells_r1 = bank_t.rows[1].cells
    cells_r1[0].merge(cells_r1[1])
    cells_r1[2].text = "Сч. №"
    cells_r1[3].text = PROVIDER['corr_account']
    
    cells_r2 = bank_t.rows[2].cells
    cells_r2[0].text = f"ИНН {PROVIDER['inn']}"
    cells_r2[1].text = "КПП —"
    cells_r2[2].text = "Сч. №"
    cells_r2[3].text = PROVIDER['account']
    
    cells_r3 = bank_t.rows[3].cells
    cells_r3[0].merge(cells_r3[1])
    cells_r3[0].text = f"{PROVIDER['status']} {PROVIDER['name']}\nПолучатель"
    cells_r3[2].merge(cells_r2[2])
    cells_r3[3].merge(cells_r2[3])

    doc.add_paragraph()

    # Title
    p_title = doc.add_paragraph()
    run_t = p_title.add_run(f"Счет на оплату № {invoice_no} от {date_str} г.")
    run_t.bold = True
    run_t.font.size = Pt(14)
    p_title.paragraph_format.space_after = Pt(12)

    # Metadata
    p_meta = doc.add_paragraph()
    p_meta.add_run("Поставщик: ").bold = True
    p_meta.add_run(f"{PROVIDER['name']}, ИНН {PROVIDER['inn']}, {PROVIDER['status']}\n")
    p_meta.add_run("Покупатель: ").bold = True
    p_meta.add_run(f"{customer['name']}, ИНН {customer['inn']}\n")
    p_meta.add_run("Основание: ").bold = True
    p_meta.add_run("Публичный договор-оферта сервиса RevOps OS (ai-rop.ru/offer.html)")

    # Items table
    items_t = doc.add_table(rows=2, cols=6)
    items_t.style = 'Table Grid'
    hdr = items_t.rows[0].cells
    headers = ["№", "Товары (работы, услуги)", "Кол-во", "Ед.", "Цена", "Сумма"]
    for i, h in enumerate(headers):
        hdr[i].text = h
        hdr[i].paragraphs[0].runs[0].bold = True
        hdr[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    row1 = items_t.rows[1].cells
    row1[0].text = "1"
    row1[1].text = item_name
    row1[2].text = "1"
    row1[3].text = "усл."
    row1[4].text = f"{amount:,.2f} ₽"
    row1[5].text = f"{amount:,.2f} ₽"

    # Total
    p_tot = doc.add_paragraph()
    p_tot.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_tot.paragraph_format.space_before = Pt(10)
    p_tot.add_run(f"Итого: {amount:,.2f} ₽\nБез налога (НДС): —\n").bold = False
    run_bold_tot = p_tot.add_run(f"Всего к оплате: {amount:,.2f} ₽")
    run_bold_tot.bold = True
    run_bold_tot.font.size = Pt(12)

    # Words
    words = number_to_words_ru(amount)
    p_words = doc.add_paragraph()
    p_words.add_run(f"Всего наименований 1, на сумму {amount:,.2f} ₽\n")
    p_words.add_run(f"Сумма прописью: {words}.\n").bold = True
    p_words.add_run("НДС не облагается в связи с применением специального налогового режима НПД (422-ФЗ).").font.size = Pt(8.5)

    # Payment purpose box
    p_purp = doc.add_paragraph()
    p_purp.paragraph_format.space_before = Pt(10)
    run_purp = p_purp.add_run(f"Назначение платежа:\nОплата по Счету № {invoice_no} от {date_str} за предоставление доступа к сервису ai-rop.ru по договору-оферте. НДС не облагается (НПД 422-ФЗ).")
    run_purp.font.size = Pt(9)
    run_purp.italic = True

    # Signatures
    p_sign = doc.add_paragraph()
    p_sign.paragraph_format.space_before = Pt(25)
    p_sign.add_run(f"Исполнитель: ____________________ / {PROVIDER['name']} /")

    doc.save(output_path)
    print(f"  [+] Invoice DOCX created: {output_path}")

def create_act_docx(act_no, date_str, customer, item_name, amount, output_path):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    p_title = doc.add_paragraph()
    run_t = p_title.add_run(f"Акт сдачи-приемки оказанных услуг № {act_no} от {date_str} г.")
    run_t.bold = True
    run_t.font.size = Pt(14)
    p_title.paragraph_format.space_after = Pt(12)

    p_meta = doc.add_paragraph()
    p_meta.add_run("Исполнитель: ").bold = True
    p_meta.add_run(f"{PROVIDER['name']}, ИНН {PROVIDER['inn']}, {PROVIDER['status']}\n")
    p_meta.add_run("Заказчик: ").bold = True
    p_meta.add_run(f"{customer['name']}, ИНН {customer['inn']}\n")
    p_meta.add_run("Основание: ").bold = True
    p_meta.add_run("Публичный договор-оферта сервиса RevOps OS (ai-rop.ru/offer.html)")

    items_t = doc.add_table(rows=2, cols=6)
    items_t.style = 'Table Grid'
    hdr = items_t.rows[0].cells
    headers = ["№", "Наименование оказанных услуг", "Кол-во", "Ед.", "Цена", "Сумма"]
    for i, h in enumerate(headers):
        hdr[i].text = h
        hdr[i].paragraphs[0].runs[0].bold = True
        hdr[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    row1 = items_t.rows[1].cells
    row1[0].text = "1"
    row1[1].text = item_name
    row1[2].text = "1"
    row1[3].text = "усл."
    row1[4].text = f"{amount:,.2f} ₽"
    row1[5].text = f"{amount:,.2f} ₽"

    p_tot = doc.add_paragraph()
    p_tot.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_tot.paragraph_format.space_before = Pt(10)
    p_tot.add_run(f"Итого оказано услуг: {amount:,.2f} ₽\nБез налога (НДС): —\n").bold = True

    words = number_to_words_ru(amount)
    p_words = doc.add_paragraph()
    p_words.add_run(f"Всего оказано услуг 1, на сумму {amount:,.2f} ₽\n")
    p_words.add_run(f"Сумма прописью: {words}.\n").bold = True

    p_clause = doc.add_paragraph()
    p_clause.paragraph_format.space_before = Pt(10)
    p_clause.add_run(
        "Вышеперечисленные услуги оказаны в полном объеме, своевременно и надлежащего качества. "
        "Стороны претензий по объему, качеству и срокам оказания услуг друг к другу не имеют.\n"
        "К настоящему Акту прилагается электронный чек из приложения ФНС России «Мой налог» (п. 1 ст. 15 Федерального закона от 27.11.2018 № 422-ФЗ)."
    ).font.size = Pt(9.5)

    p_sign = doc.add_paragraph()
    p_sign.paragraph_format.space_before = Pt(30)
    p_sign.add_run(f"ИСПОЛНИТЕЛЬ:                              ЗАКАЗЧИК:\n\n")
    p_sign.add_run(f"__________ / {PROVIDER['name']} /               __________ / М.П. /")

    doc.save(output_path)
    print(f"  [+] Act DOCX created: {output_path}")

def main():
    os.makedirs("billing", exist_ok=True)
    today_str = datetime.now().strftime("%d.%m.%Y")
    
    # Example client: ООО "ТД Автожидкости"
    client_sample = {
        "name": 'ООО "ТД Автожидкости"',
        "inn": "7701984512",
        "kpp": "770101001",
        "address": "г. Москва, ул. Автозаводская, д. 14, оф. 302"
    }

    # 1. Invoice for 7-day pilot or month 1
    create_invoice_html(
        invoice_no="01-2026",
        date_str=today_str,
        customer=client_sample,
        item_name="Предоставление удаленного доступа к программному сервису речевой аналитики и контроля звонков ai-rop.ru по Тарифу PRO на 1 месяц",
        amount=40000,
        output_path="billing/Счет_01-2026_ТД_Автожидкости.html"
    )
    create_invoice_docx(
        invoice_no="01-2026",
        date_str=today_str,
        customer=client_sample,
        item_name="Предоставление удаленного доступа к программному сервису речевой аналитики и контроля звонков ai-rop.ru по Тарифу PRO на 1 месяц",
        amount=40000,
        output_path="billing/Счет_01-2026_ТД_Автожидкости.docx"
    )

    # 2. Acceptance Act
    create_act_html(
        act_no="01-2026",
        date_str=today_str,
        customer=client_sample,
        item_name="Услуги по предоставлению удаленного доступа к аналитическому сервису ai-rop.ru (RevOps OS Pro) по Тарифу PRO за отчетный период",
        amount=40000,
        output_path="billing/Акт_01-2026_ТД_Автожидкости.html"
    )
    create_act_docx(
        act_no="01-2026",
        date_str=today_str,
        customer=client_sample,
        item_name="Услуги по предоставлению удаленного доступа к аналитическому сервису ai-rop.ru (RevOps OS Pro) по Тарифу PRO за отчетный период",
        amount=40000,
        output_path="billing/Акт_01-2026_ТД_Автожидкости.docx"
    )

    print("\n[SUCCESS] Billing package ready in ./billing directory!")

if __name__ == '__main__':
    main()
