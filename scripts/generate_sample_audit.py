# -*- coding: utf-8 -*-
"""
Generate Sample Audit Report DOCX and HTML, then convert DOCX to PDF via Word COM.
"""

import os
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

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

def generate_sample_docx(output_path):
    doc = docx.Document()
    
    # Page setup
    for sec in doc.sections:
        sec.top_margin = Inches(0.7)
        sec.bottom_margin = Inches(0.7)
        sec.left_margin = Inches(0.8)
        sec.right_margin = Inches(0.8)

    # Header banner
    tbl = doc.add_table(rows=1, cols=1)
    tbl.autofit = False
    tbl.columns[0].width = Inches(6.9)
    cell = tbl.cell(0, 0)
    set_cell_background(cell, "0F172A")
    set_cell_margins(cell, top=200, bottom=200, left=240, right=240)

    p1 = cell.paragraphs[0]
    r1 = p1.add_run("REVOPS OS  |  ДЕМОНСТРАЦИОННЫЙ ОБРАЗЕЦ ИИ-АУДИТА\n")
    r1.font.name = "Calibri"
    r1.font.size = Pt(10)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(56, 189, 248)

    r2 = p1.add_run("Экспресс-Аудит 3 звонков отдела B2B-продаж")
    r2.font.name = "Calibri"
    r2.font.size = Pt(18)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(255, 255, 255)

    p2 = cell.add_paragraph()
    r3 = p2.add_run("Компания: ООО «ПромСнаб-Сервис»  •  Ниша: Оборудование и комплектующие  •  Дата: Октябрь 2026")
    r3.font.name = "Calibri"
    r3.font.size = Pt(10)
    r3.font.color.rgb = RGBColor(203, 213, 225)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Executive Summary Box
    tbl_sum = doc.add_table(rows=1, cols=1)
    tbl_sum.autofit = False
    tbl_sum.columns[0].width = Inches(6.9)
    c_sum = tbl_sum.cell(0, 0)
    set_cell_background(c_sum, "FEF2F2") # Light red
    set_cell_margins(c_sum, top=140, bottom=140, left=180, right=180)
    
    p_sum = c_sum.paragraphs[0]
    r_sum_title = p_sum.add_run("🚨 ГЛАВНЫЙ ИТОГ АУДИТА: В 2 ИЗ 3 ЗВОНКОВ ВЫЯВЛЕН СЛИВ СДЕЛКИ\n")
    r_sum_title.font.name = "Calibri"
    r_sum_title.font.size = Pt(11)
    r_sum_title.font.bold = True
    r_sum_title.font.color.rgb = RGBColor(153, 27, 27)

    r_sum_desc = p_sum.add_run(
        "• Сумма потенциального чека в разобранных звонках: 1 420 000 ₽\n"
        "• Объем денег под критическим риском слива: 980 000 ₽ (69% от суммы диалогов)\n"
        "• Главная причина слива: отсутствие фиксации Следующего Шага (Next Step) и капитуляция перед возражением «Дорого» без аргументации окупаемости."
    )
    r_sum_desc.font.name = "Calibri"
    r_sum_desc.font.size = Pt(9.5)
    r_sum_desc.font.color.rgb = RGBColor(127, 29, 29)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Matrix Table
    h2 = doc.add_heading("1. Сводная матрица оценки качества по 13 критериям", level=2)
    h2.paragraph_format.space_after = Pt(4)

    matrix_rows = [
        ("Критерий оценки ИИ-Супервайзера", "Звонок №1 (Алексей)", "Звонок №2 (Михаил)", "Звонок №3 (Елена)"),
        ("1. Приветствие по стандарту компании", "10 / 10", "10 / 10", "10 / 10"),
        ("2. Квалификация клиента (бюджет, ЛПР, сроки)", "7 / 10", "3 / 10 ⚠️", "9 / 10"),
        ("3. Выявление глубинной боли и задачи", "8 / 10", "2 / 10 ❌", "9 / 10"),
        ("4. Презентация решения через ценность", "8 / 10", "4 / 10 ⚠️", "10 / 10"),
        ("5. Отработка возражения «Дорого»", "0 / 10 ❌ (дал скидку)", "— (не возникло)", "9 / 10 (отработал)"),
        ("6. Фиксация дедлайна и Следующего Шага", "0 / 10 ❌ (слив)", "0 / 10 ❌ (слив)", "10 / 10 (назначен zoom)"),
        ("7. Итоговый балл качества диалога", "48% (Критично)", "34% (Критично)", "92% (Эталон)"),
        ("СУММА СДЕЛКИ ПОД РИСКОМ", "480 000 ₽", "500 000 ₽", "0 ₽ (Сделка в графике)")
    ]

    tbl_m = doc.add_table(rows=len(matrix_rows), cols=4)
    tbl_m.autofit = False
    col_w = [Inches(2.7), Inches(1.4), Inches(1.4), Inches(1.4)]

    for r_idx, row in enumerate(matrix_rows):
        for c_idx, val in enumerate(row):
            cell = tbl_m.cell(r_idx, c_idx)
            cell.width = col_w[c_idx]
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.name = "Calibri"
            if r_idx == 0:
                set_cell_background(cell, "0F172A")
                set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
                r.font.size = Pt(9)
                r.font.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)
            elif r_idx == len(matrix_rows) - 1:
                set_cell_background(cell, "FEF3C7")
                set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
                r.font.size = Pt(9)
                r.font.bold = True
                r.font.color.rgb = RGBColor(146, 64, 14)
            else:
                bg = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
                set_cell_background(cell, bg)
                set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
                r.font.size = Pt(8.5)
                r.font.color.rgb = RGBColor(30, 41, 59)
                if c_idx == 0:
                    r.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Breakdown Section
    h3 = doc.add_heading("2. Стенограмма и разбор критического момента (Звонок №1)", level=2)
    h3.paragraph_format.space_after = Pt(4)

    p_dial = doc.add_paragraph()
    p_dial.add_run("[03:14] Клиент: ").bold = True
    p_dial.add_run("«Слушайте, предложение интересное, но 480 тысяч сейчас дороговато для нас. В другом месте предлагали за 420.»\n")
    p_dial.add_run("[03:22] Менеджер Алексей: ").bold = True
    p_dial.add_run("«Ну... если для вас дорого, мы можем согласовать скидку 10%. Вы подумайте тогда, а на следующей неделе созвонимся.»\n")
    p_dial.add_run("[03:31] Клиент: ").bold = True
    p_dial.add_run("«Хорошо, если что — мы сами наберем. До свидания.»\n")

    # AI Alert Callout
    tbl_alert = doc.add_table(rows=1, cols=1)
    tbl_alert.autofit = False
    tbl_alert.columns[0].width = Inches(6.9)
    c_alert = tbl_alert.cell(0, 0)
    set_cell_background(c_alert, "F0FDF4")
    set_cell_margins(c_alert, top=120, bottom=120, left=160, right=160)
    
    p_al = c_alert.paragraphs[0]
    r_al_title = p_al.add_run("💡 АЛЕРТ REVOPS OS ДЛЯ РОПА В TELEGRAM (ОТПРАВЛЕН ЧЕРЕЗ 42 СЕКУНДЫ ПОСЛЕ ЗВОНКА):\n")
    r_al_title.font.name = "Calibri"
    r_al_title.font.size = Pt(10)
    r_al_title.font.bold = True
    r_al_title.font.color.rgb = RGBColor(22, 101, 52)

    r_al_body = p_al.add_run(
        "• Менеджер сразу сдал 48 000 ₽ маржи скидкой, не выяснив комплектацию конкурента.\n"
        "• Фраза «сами наберем» — классический слив сделки. Дата и время контакта не согласованы.\n"
        "• Действие РОПу: позвонить клиенту сегодня до 17:00, предложить сравнительную таблицу характеристик и зафиксировать встречу на вторник 11:00."
    )
    r_al_body.font.name = "Calibri"
    r_al_body.font.size = Pt(9)
    r_al_body.font.color.rgb = RGBColor(20, 83, 45)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Next Steps for Client
    h4 = doc.add_heading("3. Как получить такой же разбор по вашей команде", level=2)
    p_cta = doc.add_paragraph()
    p_cta.add_run(
        "Пришлите 3 вчерашних аудиозаписи любых ваших менеджеров в Telegram "
    )
    r_tg = p_cta.add_run("@dm1918")
    r_tg.bold = True
    r_tg.font.color.rgb = RGBColor(2, 132, 199)
    p_cta.add_run(" или оформите заявку на сайте ")
    r_site = p_cta.add_run("ai-rop.ru")
    r_site.bold = True
    r_site.font.color.rgb = RGBColor(5, 150, 105)
    p_cta.add_run(". Через 20 минут вы получите персональную PDF-карту сливов вашей выручки (0 ₽).")

    doc.save(output_path)
    print(f"✓ Saved DOCX: {output_path}")

def generate_sample_html(output_path):
    html_content = """<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Образец ИИ-Аудита 3 звонков // RevOps OS</title>
    <link rel="stylesheet" href="css/tailwind.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
    <style>
        body { font-family: 'Inter', -apple-system, sans-serif; }
    </style>
</head>
<body class="bg-slate-950 text-slate-100 min-h-screen py-10 px-4 sm:px-6">
    <div class="max-w-4xl mx-auto">
        <!-- Top Nav -->
        <div class="flex items-center justify-between pb-6 mb-6 border-b border-slate-800">
            <a href="/" class="flex items-center gap-2 text-xs text-slate-400 hover:text-white transition">
                <i class="fas fa-arrow-left text-emerald-400"></i> Вернуться на главную ai-rop.ru
            </a>
            <a href="samples/Образец_ИИ_Аудита_3_Звонков_RevOps_OS.pdf" download class="px-4 py-2 bg-gradient-to-r from-emerald-500 to-cyan-500 text-slate-950 font-bold rounded-xl text-xs flex items-center gap-2 shadow-lg shadow-emerald-500/20 hover:from-emerald-400 hover:to-cyan-400 transition">
                <i class="fas fa-download"></i> Скачать PDF (образцовый отчет)
            </a>
        </div>

        <!-- Document Preview Card -->
        <div class="rounded-3xl bg-slate-900 border border-slate-800 p-6 sm:p-10 shadow-2xl space-y-8">
            <!-- Header -->
            <div class="border-b border-slate-800 pb-6">
                <div class="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-cyan-500/10 text-cyan-400 font-mono text-xs font-semibold mb-3 border border-cyan-500/20">
                    // ДЕМОНСТРАЦИОННЫЙ ОБРАЗЕЦ ИИ-АУДИТА
                </div>
                <h1 class="text-2xl sm:text-3xl font-extrabold text-white">Экспресс-Аудит 3 звонков отдела B2B-продаж</h1>
                <p class="text-slate-400 text-xs sm:text-sm mt-1">Клиент: ООО «ПромСнаб-Сервис» • Ниша: Промышленное оборудование • Октябрь 2026</p>
            </div>

            <!-- Key Findings Box -->
            <div class="p-5 sm:p-6 rounded-2xl bg-rose-500/10 border border-rose-500/30 text-rose-200">
                <div class="flex items-center gap-2 font-bold text-rose-300 text-sm sm:text-base mb-2">
                    <i class="fas fa-triangle-exclamation text-rose-400"></i>
                    <span>В 2 ИЗ 3 ЗВОНКОВ ВЫЯВЛЕН КРИТИЧЕСКИЙ СЛИВ СДЕЛКИ</span>
                </div>
                <ul class="text-xs sm:text-sm space-y-1.5 text-slate-300">
                    <li>• Сумма чека в разобранных диалогах: <strong>1 420 000 ₽</strong></li>
                    <li>• Сумма под критическим риском слива: <strong class="text-rose-400">980 000 ₽</strong> (69% от суммы диалогов)</li>
                    <li>• Главная причина сбоя: менеджер сдал 48 000 ₽ маржи скидкой и отпустил клиента фразой «вы подумайте, потом созвонимся».</li>
                </ul>
            </div>

            <!-- Matrix -->
            <div>
                <h3 class="text-lg font-bold text-white mb-4">Сводная матрица оценки по 13 критериям качества</h3>
                <div class="overflow-x-auto rounded-xl border border-slate-800">
                    <table class="w-full text-xs text-left">
                        <thead class="bg-slate-950 text-slate-400 font-mono">
                            <tr>
                                <th class="p-3">Критерий оценки ИИ</th>
                                <th class="p-3">Звонок 1 (Алексей)</th>
                                <th class="p-3">Звонок 2 (Михаил)</th>
                                <th class="p-3">Звонок 3 (Елена)</th>
                            </tr>
                        </thead>
                        <tbody class="divide-y divide-slate-800/80 text-slate-300">
                            <tr>
                                <td class="p-3 font-medium text-white">Приветствие по регламенту</td>
                                <td class="p-3 text-emerald-400">10 / 10</td>
                                <td class="p-3 text-emerald-400">10 / 10</td>
                                <td class="p-3 text-emerald-400">10 / 10</td>
                            </tr>
                            <tr>
                                <td class="p-3 font-medium text-white">Квалификация ЛПР и бюджета</td>
                                <td class="p-3 text-amber-400">7 / 10</td>
                                <td class="p-3 text-rose-400 font-bold">3 / 10 ⚠️</td>
                                <td class="p-3 text-emerald-400">9 / 10</td>
                            </tr>
                            <tr>
                                <td class="p-3 font-medium text-white">Отработка возражения «Дорого»</td>
                                <td class="p-3 text-rose-400 font-bold">0 / 10 ❌ (дал скидку)</td>
                                <td class="p-3 text-slate-500">—</td>
                                <td class="p-3 text-emerald-400">9 / 10</td>
                            </tr>
                            <tr>
                                <td class="p-3 font-medium text-white">Фиксация Следующего Шага (Next Step)</td>
                                <td class="p-3 text-rose-400 font-bold">0 / 10 ❌ (слив)</td>
                                <td class="p-3 text-rose-400 font-bold">0 / 10 ❌ (слив)</td>
                                <td class="p-3 text-emerald-400">10 / 10 (zoom)</td>
                            </tr>
                            <tr class="bg-slate-950/80 font-bold">
                                <td class="p-3 text-white">Сумма сделки под риском</td>
                                <td class="p-3 text-rose-400">480 000 ₽</td>
                                <td class="p-3 text-rose-400">500 000 ₽</td>
                                <td class="p-3 text-emerald-400">0 ₽ (в графике)</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>

            <!-- Instant CTA -->
            <div class="p-6 rounded-2xl bg-emerald-500/10 border border-emerald-500/30 text-center">
                <h4 class="text-base sm:text-lg font-bold text-white mb-2">Хотите увидеть такой разбор по вашим менеджерам?</h4>
                <p class="text-xs sm:text-sm text-slate-300 max-w-xl mx-auto mb-4">
                    Пришлите 3 вчерашних аудиозаписи любых ваших сотрудников. Через 20 минут основатель RevOps OS Дмитрий Федотов выдаст персональную карту сливов (0 ₽).
                </p>
                <div class="flex flex-wrap gap-3 justify-center">
                    <a href="/#cta" class="px-6 py-3 bg-gradient-to-r from-emerald-500 to-cyan-500 hover:from-emerald-400 hover:to-cyan-400 text-slate-950 font-bold rounded-xl text-xs sm:text-sm shadow-lg transition">
                        Оставить заявку на бесплатный аудит (0 ₽)
                    </a>
                    <a href="https://t.me/dm1918" target="_blank" class="px-5 py-3 bg-slate-800 hover:bg-slate-700 text-white font-semibold rounded-xl text-xs sm:text-sm border border-slate-700 transition flex items-center gap-1.5">
                        <i class="fab fa-telegram text-cyan-400"></i> Отправить звонки в Telegram
                    </a>
                </div>
            </div>
        </div>
    </div>
</body>
</html>"""
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"✓ Saved HTML: {output_path}")

if __name__ == "__main__":
    repo = r"C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os"
    os.makedirs(os.path.join(repo, "docs", "samples"), exist_ok=True)
    os.makedirs(os.path.join(repo, "samples"), exist_ok=True)

    docx_path = os.path.join(repo, "docs", "samples", "Образец_ИИ_Аудита_3_Звонков_RevOps_OS.docx")
    generate_sample_docx(docx_path)

    html_path = os.path.join(repo, "sample-audit-report.html")
    generate_sample_html(html_path)
