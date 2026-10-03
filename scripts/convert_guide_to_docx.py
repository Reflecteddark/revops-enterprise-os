#!/usr/bin/env python3
"""
Convert B2B Accounting Cheat Sheet to beautifully formatted Word DOCX.
Also ensures all billing files are easily accessible.
"""

import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

sys.stdout.reconfigure(encoding='utf-8')

def build_guide_docx(output_path):
    doc = docx.Document()
    
    # 1. Page setup: 2 cm margins
    for s in doc.sections:
        s.top_margin = Inches(0.7)
        s.bottom_margin = Inches(0.7)
        s.left_margin = Inches(0.7)
        s.right_margin = Inches(0.7)

    # Header title
    title_p = doc.add_paragraph()
    r_title = title_p.add_run("📋 БОЕВАЯ ИНСТРУКЦИЯ ДЛЯ САМОЗАНЯТОГО В B2B\nЗакрывающие документы, чеки и ответы бухгалтерии клиентов (422-ФЗ)")
    r_title.bold = True
    r_title.font.size = Pt(16)
    r_title.font.color.rgb = RGBColor(15, 23, 42) # Dark Slate
    title_p.paragraph_format.space_after = Pt(12)

    sub_p = doc.add_paragraph()
    r_sub = sub_p.add_run("Практическое руководство для основателя платформы ai-rop.ru (RevOps OS Pro) Дмитрия Федотова.\nАктуально на 2026 год • Соответствует ФЗ № 422-ФЗ и ФЗ № 402-ФЗ.")
    r_sub.italic = True
    r_sub.font.size = Pt(10)
    r_sub.font.color.rgb = RGBColor(100, 116, 139)
    sub_p.paragraph_format.space_after = Pt(16)

    # Section 1
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("1. Главная юридическая база (что нужно знать наизусть)")
    r_h1.font.size = Pt(13)
    r_h1.font.color.rgb = RGBColor(14, 116, 144)

    p1 = doc.add_paragraph()
    p1.add_run("1. Закон № 422-ФЗ (ст. 15): ").bold = True
    p1.add_run("Юридическое лицо (ООО, АО) обязано иметь от самозанятого электронный чек из приложения «Мой налог» с указанием своего ИНН. На основании этого чека компания в полном объеме списывает сумму в расходы и уменьшает налог на прибыль (или налог при УСН «Доходы минус Расходы»).\n\n")
    p1.add_run("2. Налоги клиента: ").bold = True
    p1.add_run("Компания-клиент НЕ платит за вас НДФЛ 13% и НЕ платит страховые взносы 30% в Социальный фонд РФ (в отличие от обычных договоров ГПХ с физлицами без статуса самозанятого). Для компании это абсолютно чистый и законный B2B-платеж.\n\n")
    p1.add_run("3. Ваш налог: ").bold = True
    p1.add_run("Вы платите 6% с поступлений от юридических лиц. Налог рассчитывается автоматически в приложении «Мой налог» и оплачивается до 28-го числа следующего месяца.")

    # Section 2
    h2 = doc.add_heading(level=1)
    r_h2 = h2.add_run("2. Пошаговый регламент по 4 этапам работы с клиентом")
    r_h2.font.size = Pt(13)
    r_h2.font.color.rgb = RGBColor(14, 116, 144)

    # Step 1
    doc.add_heading("Этап 1. Бесплатный аудит и ДЕМО (0 ₽)", level=2)
    p_step1 = doc.add_paragraph()
    p_step1.add_run("• Документы бухгалтерии: ").bold = True
    p_step1.add_run("НЕ ТРЕБУЮТСЯ.\nПоскольку денег не передается, в бухгалтерии клиента операция не проводится. Согласие с условиями обработки звонков и политикой 152-ФЗ клиент дает на сайте ai-rop.ru при отправке аудиозаписей.")

    # Step 2
    doc.add_heading("Этап 2. Выставление счета на первую оплату (Тариф PRO — 40 000 ₽)", level=2)
    p_step2 = doc.add_paragraph()
    p_step2.add_run("1. Откройте готовый шаблон счета: ")
    p_step2.add_run("billing/Счет_01-2026_ТД_Автожидкости.docx.\n")
    p_step2.add_run("2. Проверьте номер счета, дату и название компании-клиента с её ИНН.\n")
    p_step2.add_run("3. Сохраните в PDF (Файл → Сохранить как PDF).\n")
    p_step2.add_run("4. Отправьте клиенту/бухгалтеру с сопроводительным письмом:")

    # Message quote box
    p_msg = doc.add_paragraph()
    p_msg.paragraph_format.left_indent = Inches(0.3)
    r_msg = p_msg.add_run(
        "«Добрый день! Во вложении направляю счет № 01-2026 на подключение сервиса речевой аналитики ai-rop.ru по Тарифу PRO.\n"
        "Оплата производится в безналичном порядке по договору публичной оферты (ai-rop.ru/offer.html). НДС не облагается (применение спецрежима НПД по 422-ФЗ).\n"
        "В день зачисления оплаты мы предоставим электронный фискальный чек ФНС РФ с вашим ИНН для бухгалтерии и закрывающий Акт.\n"
        "Реквизиты и назначение платежа указаны в счете. Спасибо за сотрудничество!»"
    )
    r_msg.italic = True
    r_msg.font.color.rgb = RGBColor(30, 41, 59)

    # Step 3
    doc.add_heading("Этап 3. Деньги пришли: Как сформировать чек в «Мой налог» (30 секунд)", level=2)
    p_step3 = doc.add_paragraph()
    p_step3.add_run("1. Откройте приложение «Мой налог» на смартфоне (или сайт lknpd.nalog.ru).\n")
    p_step3.add_run("2. Нажмите кнопку «Новая продажа» (+).\n")
    p_step3.add_run("3. Введите название: ").bold = True
    p_step3.add_run("Предоставление удаленного доступа к программному сервису речевой аналитики ai-rop.ru по Тарифу PRO.\n")
    p_step3.add_run("4. Введите сумму: ").bold = True
    p_step3.add_run("40 000 ₽.\n")
    p_step3.add_run("5. Переключите тумблер на ").bold = True
    p_step3.add_run("«Юридическому лицу или ИП».\n")
    p_step3.add_run("6. Введите ИНН клиента: ").bold = True
    p_step3.add_run("приложение само подтянет официальное наименование компании из базы ФНС!\n")
    p_step3.add_run("7. Нажмите «Выдать чек» → нажмите «Отправить» → скопируйте ссылку или отправьте PDF бухгалтеру.")

    # Step 4
    doc.add_heading("Этап 4. Закрывающий Акт (Конец месяца)", level=2)
    p_step4 = doc.add_paragraph()
    p_step4.add_run("1. Откройте файл: ")
    p_step4.add_run("billing/Акт_01-2026_ТД_Автожидкости.docx.\n")
    p_step4.add_run("2. Проверьте реквизиты, сумму и дату окончания периода.\n")
    p_step4.add_run("3. Сохраните в PDF и отправьте бухгалтеру по e-mail (или через Диадок/СБИС, если клиент работает через ЭДО).\n")
    p_step4.add_run("4. В оферте на ai-rop.ru/offer.html уже работает пункт о молчаливом акцепте: если в течение 3 рабочих дней мотивированных претензий не поступило, услуги считаются принятыми в полном объеме.")

    # Section 3
    h3 = doc.add_heading(level=1)
    r_h3 = h3.add_run("3. Готовые скрипты ответов на 3 главных вопроса бухгалтерии клиента")
    r_h3.font.size = Pt(13)
    r_h3.font.color.rgb = RGBColor(14, 116, 144)

    # Q&A Table
    qa_table = doc.add_table(rows=4, cols=2)
    qa_table.style = 'Table Grid'
    
    hdr = qa_table.rows[0].cells
    hdr[0].text = "Вопрос бухгалтера клиента"
    hdr[0].paragraphs[0].runs[0].bold = True
    hdr[1].text = "Как уверенно и профессионально ответить"
    hdr[1].paragraphs[0].runs[0].bold = True

    q1 = qa_table.rows[1].cells
    q1[0].text = "«А где в счете НДС?»"
    q1[1].text = '«Исполнитель применяет специальный налоговый режим НПД (налог на профессиональный доход) в соответствии с Федеральным законом № 422-ФЗ. Согласно п. 9 ст. 2 закона 422-ФЗ, операции НДС не облагаются. В платежном поручении указывается: "НДС не облагается"».'

    q2 = qa_table.rows[2].cells
    q2[0].text = "«А почему нет бумажного договора с синей печатью?»"
    q2[1].text = '«Договор заключается в электронной форме путем акцепта Публичной оферты (ст. 434 и 438 Гражданского кодекса РФ), размещенной на ai-rop.ru/offer.html. Оплата выставленного счета является полным и безоговорочным акцептом оферты. Это стандартная практика для IT-сервисов (так работают Яндекс, amoCRM, Контур)».'

    q3 = qa_table.rows[3].cells
    q3[0].text = "«Как мне закрыть расходы в налоговом учете?»"
    q3[1].text = '«В соответствии с п. 8–10 ст. 15 Федерального закона № 422-ФЗ и письмами Минфина РФ, единственным обязательным первичным документом для уменьшения налогооблагаемой базы является электронный фискальный чек ФНС из приложения "Мой налог" с указанием ИНН вашей организации. Чек выдается в день оплаты. Дополнительно мы предоставляем закрывающий Акт».'

    # Section 4: File directory
    doc.add_paragraph()
    h4 = doc.add_heading(level=1)
    r_h4 = h4.add_run("4. Где лежат ваши файлы на компьютере")
    r_h4.font.size = Pt(13)
    r_h4.font.color.rgb = RGBColor(14, 116, 144)

    p_files = doc.add_paragraph()
    p_files.add_run("Папка: ").bold = True
    p_files.add_run("C:\\Users\\strel\\.gemini\\antigravity\\scratch\\revops-enterprise-os\\billing\\\n")
    p_files.add_run("• Счет_01-2026_ТД_Автожидкости.docx — готовый счет в Word\n")
    p_files.add_run("• Акт_01-2026_ТД_Автожидкости.docx — готовый акт сдачи-приемки в Word\n")
    p_files.add_run("• Инструкция_Бухгалтерия_Самозанятого_B2B.docx — настоящая инструкция в Word\n")
    p_files.add_run("• Счет_01-2026_ТД_Автожидкости.html — веб-форма счета для быстрой печати в PDF\n")
    p_files.add_run("• Акт_01-2026_ТД_Автожидкости.html — веб-форма акта для быстрой печати в PDF\n")

    doc.save(output_path)
    print(f"  [+] Guide DOCX successfully created: {output_path}")

if __name__ == '__main__':
    os.makedirs("billing", exist_ok=True)
    os.makedirs("docs", exist_ok=True)
    
    # 1. Generate in docs/ and billing/
    build_guide_docx("docs/Инструкция_Бухгалтерия_Самозанятого_B2B.docx")
    build_guide_docx("billing/Инструкция_Бухгалтерия_Самозанятого_B2B.docx")
    
    # 2. Also copy everything to parent scratch directory for instant 1-click access
    scratch_billing = "C:/Users/strel/.gemini/antigravity/scratch/billing"
    os.makedirs(scratch_billing, exist_ok=True)
    import shutil
    for f in os.listdir("billing"):
        src = os.path.join("billing", f)
        dst = os.path.join(scratch_billing, f)
        if os.path.isfile(src):
            shutil.copy2(src, dst)
            print(f"  [+] Copied to scratch/billing/: {f}")

    print("\n[SUCCESS] All Word documents and billing files ready!")
