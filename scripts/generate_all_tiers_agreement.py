import datetime
from pathlib import Path
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

def set_cell_shading(cell, color_hex):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def create_docx(output_path: Path):
    doc = docx.Document()

    # Page Margins: 0.65 in
    for s in doc.sections:
        s.top_margin = Inches(0.65)
        s.bottom_margin = Inches(0.65)
        s.left_margin = Inches(0.75)
        s.right_margin = Inches(0.75)

    # Document Header
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(4)
    run_title = p_title.add_run("ПУБЛИЧНЫЙ ДОГОВОР-ОФЕРТА\nна предоставление доступа к программному сервису ИИ-аналитики продаж «RevOps OS»\n(Все тарифные планы: Пилот, Старт, Рост, Корпоративный)")
    run_title.font.name = "Arial"
    run_title.font.size = Pt(12.5)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(15, 23, 42)

    p_date = doc.add_paragraph()
    p_date.paragraph_format.space_after = Pt(10)
    run_city = p_date.add_run("г. Москва / Дистанционно\t\t\t\t\tРедакция: 02.10.2026 г.")
    run_city.font.name = "Arial"
    run_city.font.size = Pt(9)
    run_city.font.color.rgb = RGBColor(100, 116, 139)

    # Preamble with legal status dual-mode
    p_preamble = doc.add_paragraph()
    p_preamble.paragraph_format.space_after = Pt(8)
    p_preamble.paragraph_format.line_spacing = 1.15
    r_pre = p_preamble.add_run(
        "Настоящий Публичный Договор-Оферта (далее — «Договор») является официальным предложением Исполнителя:\n"
        "• В ТЕКУЩЕМ СТАТУСЕ: Гражданин РФ Федотов Дмитрий, применяющий специальный налоговый режим «Налог на профессиональный доход» (НПД / Самозанятый) в соответствии с Федеральным законом от 27.11.2018 № 422-ФЗ (ИНН: 731303201073);\n"
        "• ЛИБО С МОМЕНТА РЕГИСТРАЦИИ: Индивидуальный предприниматель Федотов Дмитрий (ОГРНИП: в процессе внесения в ЕГРИП, ИНН: 731303201073, УСН 6%),\n"
        "именуемый в дальнейшем «Исполнитель» (основатель сервиса ai-rop.ru), адресованным любому юридическому лицу или индивидуальному предпринимателю, "
        "зарегистрированному на территории Российской Федерации (далее — «Заказчик»), заключить договор на предоставление удаленного доступа "
        "к программному комплексу ИИ-супервизии звонков и сквозной аналитики продаж «RevOps OS» на изложенных ниже условиях."
    )
    r_pre.font.name = "Arial"
    r_pre.font.size = Pt(9)
    r_pre.font.color.rgb = RGBColor(51, 65, 85)

    def add_sec(title_text: str, content_text: str):
        p_s = doc.add_paragraph()
        p_s.paragraph_format.space_before = Pt(8)
        p_s.paragraph_format.space_after = Pt(2)
        r_s = p_s.add_run(title_text)
        r_s.font.name = "Arial"
        r_s.font.size = Pt(10)
        r_s.font.bold = True
        r_s.font.color.rgb = RGBColor(15, 23, 42)

        p_c = doc.add_paragraph()
        p_c.paragraph_format.space_after = Pt(6)
        p_c.paragraph_format.line_spacing = 1.15
        r_c = p_c.add_run(content_text)
        r_c.font.name = "Arial"
        r_c.font.size = Pt(9)
        r_c.font.color.rgb = RGBColor(51, 65, 85)

    add_sec(
        "1. ПРЕДМЕТ ДОГОВОРА",
        "1.1. Исполнитель обязуется предоставить Заказчику удаленный доступ к облачному программно-аппаратному комплексу «RevOps OS» (SaaS), "
        "обеспечивающему 100% распознавание, расшифровку, оценку по 13 критериям качества речи и выявление сливов сделок в переговорах менеджеров отдела продаж, "
        "а также отправку оперативных уведомлений в Telegram, а Заказчик обязуется принять и оплатить доступ в соответствии с выбранным Тарифным планом.\n"
        "1.2. Подключение сервиса осуществляется через официальные API к корпоративной CRM Заказчика (amoCRM или Битрикс24) либо посредством защищенной загрузки аудиофайлов."
    )

    # Section 2: Tariffs Table
    p_s2 = doc.add_paragraph()
    p_s2.paragraph_format.space_before = Pt(8)
    p_s2.paragraph_format.space_after = Pt(4)
    r_s2 = p_s2.add_run("2. ТАРИФНЫЕ ПЛАНЫ И СТОИМОСТЬ УСЛУГ")
    r_s2.font.name = "Arial"
    r_s2.font.size = Pt(10)
    r_s2.font.bold = True
    r_s2.font.color.rgb = RGBColor(15, 23, 42)

    # Tariffs Table
    tariffs = [
        ("Тарифный план", "Стоимость в месяц", "Объем менеджеров ОП", "Включенный функционал и состав услуг"),
        ("«7-дневный Пилотный спринт»", "14 900 ₽\n(разово за 7 дней)", "до 10 менеджеров", "100% оцифровка звонков за 7 дней, алерты РОПу за 60 секунд, утренние дайджесты в 09:00, PDF-карта сливов, 100% БЕЗУСЛОВНАЯ ГАРАНТИЯ ВОЗВРАТА."),
        ("Тариф «Старт» (Базовый)", "29 000 ₽ / мес\n(23 200 ₽ при оплате за год)", "до 5 менеджеров", "100% аудит всех звонков, интеграция с 1 CRM (amoCRM/Б24), утренний дайджест РОПу в Telegram, базовый чек-лист 13 критериев, техподдержка в чате."),
        ("Тариф «Рост» (Scale / Pro)", "59 000 ₽ / мес\n(47 200 ₽ при оплате за год)", "до 15 менеджеров", "Все из «Старт» + Telegram-алерты за 60 сек на срывы сделок, отработка возражения «Дорого», кастомные чек-листы под регламент Заказчика, еженедельный спринт с RevOps-экспертом."),
        ("Тариф «Корпоративный» (Enterprise)", "89 000 ₽ / мес\n(71 200 ₽ при оплате за год)", "до 30+ менеджеров", "Все из «Рост» + персональный выделенный аналитик, выделенный сервер 152-ФЗ, интеграция с несколькими CRM и филиалами, кастомные когортные дашборды, приоритетный SLA 99.9%.")
    ]

    t_table = doc.add_table(rows=len(tariffs), cols=4)
    t_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_table.autofit = False

    col_widths = [Inches(1.8), Inches(1.5), Inches(1.3), Inches(2.4)]
    for row_idx, row in enumerate(tariffs):
        for col_idx, text in enumerate(row):
            cell = t_table.cell(row_idx, col_idx)
            cell.width = col_widths[col_idx]
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.1
            run = p.add_run(text)
            run.font.name = "Arial"

            if row_idx == 0:
                set_cell_shading(cell, "1E2E4A")
                run.font.bold = True
                run.font.size = Pt(8.5)
                run.font.color.rgb = RGBColor(255, 255, 255)
            else:
                run.font.size = Pt(8)
                if row_idx % 2 == 1:
                    set_cell_shading(cell, "F8FAFC")
                if col_idx == 0:
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(15, 23, 42)
                elif col_idx == 1:
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(5, 150, 105)

    p_post_t = doc.add_paragraph()
    p_post_t.paragraph_format.space_before = Pt(4)
    p_post_t.paragraph_format.space_after = Pt(6)
    r_pt = p_post_t.add_run(
        "2.2. НДС не облагается в связи с применением Исполнителем специального налогового режима НПД (ст. 2 Федерального закона № 422-ФЗ) "
        "либо УСН (п. 2 ст. 346.11 НК РФ).\n"
        "2.3. При единовременной оплате любого регулярного тарифа («Старт», «Рост», «Корпоративный») за 12 месяцев Заказчику предоставляется скидка 20%."
    )
    r_pt.font.name = "Arial"
    r_pt.font.size = Pt(8.5)
    r_pt.font.color.rgb = RGBColor(100, 116, 139)

    add_sec(
        "3. УСЛОВИЯ БЕЗУСЛОВНОГО ВОЗВРАТА ПО ПИЛОТУ (MONEY-BACK GUARANTEE)",
        "3.1. Для Тарифа «7-дневный Пилотный спринт» действует 100% безусловная гарантия окупаемости и удовлетворенности результатами.\n"
        "3.2. Если по итогам 7 дней Заказчик посчитает, что система не выявила упущенную выручку или результаты пилота не принесли прямой ценности отделу продаж, "
        "Исполнитель обязуется вернуть 100% уплаченной суммы (14 900 рублей) на расчетный счет Заказчика в течение 3 (трех) банковских дней "
        "по первому письменному запросу Заказчика в Telegram (@dm1918) или на email founder@ai-rop.ru без штрафов и дополнительных условий."
    )

    add_sec(
        "4. ПОРЯДОК РАСЧЕТОВ И ЗАКРЫВАЮЩИЕ ДОКУМЕНТЫ (САМОЗАНЯТЫЙ / ИП)",
        "4.1. Оплата производится Заказчиком в форме 100% предоплаты по безналичному расчету на основании выставленного счета или корпоративной банковской картой.\n"
        "4.2. ПРИ ОПЛАТЕ САМОЗАНЯТОМУ (РЕЖИМ НПД):\n"
        "  • В соответствии с п. 1 ст. 14 и п. 8 ст. 15 Федерального закона от 27.11.2018 № 422-ФЗ Исполнитель в день поступления средств формирует официальный электронный фискальный чек через приложение «Мой налог» ФНС РФ с указанием ИНН и наименования Заказчика и направляет его Заказчику в электронном виде.\n"
        "  • На основании указанного чека Заказчик в полном объеме учитывает произведенные расходы при налогообложении (налог на прибыль или УСН «Доходы минус расходы» согласно Письму Минфина РФ № 03-11-11/24004). По требованию Заказчика формируется закрывающий Акт оказанных услуг.\n"
        "4.3. ПРИ ОПЛАТЕ НА РАСЧЕТНЫЙ СЧЕТ ИП (ПОСЛЕ РЕГИСТРАЦИИ ИП):\n"
        "  • Исполнитель выставляет счет на оплату от имени ИП Федотов Дмитрий (УСН 6%). По окончании расчетного месяца (или пилота) Заказчику направляется Акт сдачи-приемки через систему ЭДО (Диадок / СБИС) либо скан-копией по электронной почте. При отсутствии мотивированных возражений в течение 3 рабочих дней услуги считаются принятыми в полном объеме."
    )

    add_sec(
        "5. КОНФИДЕНЦИАЛЬНОСТЬ И БЕЗОПАСНОСТЬ ДАННЫХ (152-ФЗ РФ)",
        "5.1. Стороны обязуются соблюдать строгий режим конфиденциальности в отношении клиентской базы Заказчика, аудиозаписей переговоров и сумм сделок.\n"
        "5.2. Обработка данных осуществляется строго в соответствии с Федеральным законом № 152-ФЗ «О персональных данных». Серверная инфраструктура сервиса расположена исключительно в Российской Федерации (Selectel / Яндекс Облако, ЦОД Москва/Санкт-Петербург).\n"
        "5.3. Все персональные данные (ФИО абонентов, телефоны) автоматически токенизируются и деперсонализируются криптографическими алгоритмами до этапа распознавания речи."
    )

    add_sec(
        "6. СРОК ДЕЙСТВИЯ, ПРОЛОНГАЦИЯ И ПРЕКРАЩЕНИЕ",
        "6.1. Договор вступает в силу с момента акцепта (оплаты выбранного тарифа Заказчиком) и действует в течение оплаченного периода.\n"
        "6.2. Пролонгация на следующий месяц осуществляется путем своевременной оплаты Заказчиком выставленного счета на следующий расчетный период.\n"
        "6.3. Заказчик вправе в любой момент отказаться от продления подписки, письменно уведомив Исполнителя за 3 (три) рабочих дня до окончания оплаченного расчетного месяца."
    )

    # Section 7: Signatures & Details Block (Dual Status Mode)
    p_sig_title = doc.add_paragraph()
    p_sig_title.paragraph_format.space_before = Pt(10)
    p_sig_title.paragraph_format.space_after = Pt(4)
    r_st = p_sig_title.add_run("7. РЕКВИЗИТЫ СТОРОН (ВАРИАНТ САМОЗАНЯТОГО И ВАРИАНТ ИП)")
    r_st.font.name = "Arial"
    r_st.font.size = Pt(10)
    r_st.font.bold = True

    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_table.autofit = False

    c_isp = sig_table.cell(0, 0)
    c_zak = sig_table.cell(0, 1)
    c_isp.width = Inches(3.5)
    c_zak.width = Inches(3.5)

    p_i = c_isp.paragraphs[0]
    p_i.paragraph_format.space_after = Pt(2)
    r_ih = p_i.add_run("ИСПОЛНИТЕЛЬ (2 ВАРИАНТА РЕКВИЗИТОВ):\n\n")
    r_ih.font.bold = True
    r_ih.font.size = Pt(8.5)

    r_ib = p_i.add_run(
        "ВАРИАНТ А: САМОЗАНЯТЫЙ (ТЕКУЩИЙ):\n"
        "Плательщик НПД Федотов Дмитрий\n"
        "ИНН: 731303201073\n"
        "Банк: АО «ТИНЬКОФФ БАНК»\n"
        "Р/с: 40802810500003849120\n"
        "БИК: 044525974 | К/с: 30101810145250000974\n"
        "Основание: Чек ФНС «Мой налог» (п. 8 ст. 15 422-ФЗ)\n\n"
        "ВАРИАНТ Б: ИНДИВИДУАЛЬНЫЙ ПРЕДПРИНИМАТЕЛЬ:\n"
        "ИП Федотов Дмитрий\n"
        "ОГРНИП: [в процессе регистрации]\n"
        "ИНН: 731303201073 (УСН 6%)\n"
        "Р/с: 40802810500003849120\n"
        "Банк: АО «ТИНЬКОФФ БАНК»\n"
        "БИК: 044525974\n"
        "Telegram / MAX: @dm1918\n"
        "Сайт: ai-rop.ru | Email: founder@ai-rop.ru\n\n"
        "Подпись: _________________ / Федотов Д. /"
    )
    r_ib.font.size = Pt(8)

    p_z = c_zak.paragraphs[0]
    p_z.paragraph_format.space_after = Pt(2)
    r_zh = p_z.add_run("ЗАКАЗЧИК:\n\n")
    r_zh.font.bold = True
    r_zh.font.size = Pt(8.5)

    r_zb = p_z.add_run(
        "Наименование организации: ____________________\n"
        "ИНН / КПП: ____________________________________\n"
        "ОГРН / ОГРНИП: ________________________________\n"
        "Юридический адрес: ____________________________\n"
        "Расчетный счет: _______________________________\n"
        "Банк: _________________________________________\n"
        "БИК банка: ____________________________________\n"
        "Корр. счет: ___________________________________\n"
        "Контактное лицо (ЛПР): ________________________\n"
        "Telegram для алертов: _________________________\n"
        "Email для отчетов: ____________________________\n\n"
        "Подпись: _________________ / _________________ /\n"
        "М.П."
    )
    r_zb.font.size = Pt(8)

    doc.save(str(output_path))
    print(f"✓ DOCX agreement for all tiers generated: {output_path}")


def create_html(output_path: Path):
    html = """<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Публичный Договор-Оферта на сервис RevOps OS (Все тарифы)</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@fontsource/inter@5.0.8/index.css">
    <style>
        body { font-family: 'Inter', sans-serif; background: #0b0f19; color: #cbd5e1; }
        .glass-card { background: rgba(15, 23, 42, 0.85); backdrop-filter: blur(16px); border: 1px solid rgba(51, 65, 85, 0.5); }
    </style>
    <!-- Google tag (gtag.js) -->
    <script async src="https://www.googletagmanager.com/gtag/js?id=G-LHRE6Z6TBG"></script>
    <script>
      window.dataLayer = window.dataLayer || [];
      function gtag(){dataLayer.push(arguments);}
      gtag('js', new Date());
      gtag('config', 'G-LHRE6Z6TBG');
    </script>
    <!-- Yandex.Metrika counter -->
    <script type="text/javascript">
       (function(m,e,t,r,i,k,a){m[i]=m[i]||function(){(m[i].a=m[i].a||[]).push(arguments)};
       m[i].l=1*new Date();
       for (var j = 0; j < document.scripts.length; j++) {if (document.scripts[j].src === r) { return; }}
       k=e.createElement(t),a=e.getElementsByTagName(t)[0],k.async=1,k.src=r,a.parentNode.insertBefore(k,a)})
       (window, document, "script", "https://mc.yandex.ru/metrika/tag.js", "ym");
       ym(113329268, "init", {
            clickmap:true,
            trackLinks:true,
            accurateTrackBounce:true,
            webvisor:true
       });
    </script>
    <noscript><div><img src="https://mc.yandex.ru/watch/113329268" style="position:absolute; left:-9999px;" alt="" /></div></noscript>
    <!-- /Yandex.Metrika counter -->
</head>
<body class="min-h-screen py-8 px-4 sm:px-6 lg:px-8">
    <div class="max-w-4xl mx-auto glass-card rounded-2xl p-6 sm:p-10 shadow-2xl">
        
        <!-- Верхняя панель навигации -->
        <div class="flex items-center justify-between pb-6 border-b border-slate-800 mb-8">
            <a href="/" class="flex items-center gap-2 text-white font-bold text-lg hover:text-emerald-400 transition">
                <span class="w-8 h-8 rounded-lg bg-emerald-500/20 text-emerald-400 flex items-center justify-center border border-emerald-500/30 text-sm font-black">AI</span>
                <span>RevOps OS Pro</span>
            </a>
            <div class="flex items-center gap-3 text-xs">
                <span class="px-3 py-1 rounded-full bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 font-semibold">152-ФЗ РФ</span>
                <span class="text-slate-400">Редакция: 02.10.2026 г.</span>
            </div>
        </div>

        <!-- Заголовок -->
        <div class="text-center mb-8">
            <span class="text-xs uppercase font-bold tracking-widest text-emerald-400 bg-emerald-500/10 px-3 py-1 rounded-full border border-emerald-500/20">Официальный документ</span>
            <h1 class="text-2xl sm:text-3xl font-black text-white mt-3 mb-2">ПУБЛИЧНЫЙ ДОГОВОР-ОФЕРТА</h1>
            <p class="text-sm text-slate-400">на предоставление удаленного доступа к программному сервису ИИ-супервизии продаж «RevOps OS» (ai-rop.ru)</p>
        </div>

        <!-- Преамбула с переключателем статуса -->
        <div class="bg-slate-900/60 rounded-xl p-5 border border-slate-800 mb-8 text-xs sm:text-sm leading-relaxed text-slate-300">
            <p class="mb-3">
                Настоящий документ является публичной офертой <strong>Исполнителя (Федотова Дмитрия)</strong>, основателя сервиса <strong>ai-rop.ru</strong>, 
                адресованной юридическим лицам и индивидуальным предпринимателям РФ (далее — «Заказчик»), заключить договор на оказание услуг и предоставление доступа к сервису RevOps OS на изложенных ниже условиях.
            </p>
            <div class="p-3.5 bg-slate-950/80 rounded-lg border border-slate-800 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs">
                <div>
                    <span class="font-bold text-white">Правовой статус исполнителя:</span>
                    <div class="text-slate-400 mt-0.5">Официальный статус: <strong class="text-emerald-400">Плательщик налога на профессиональный доход (Самозанятый)</strong> по 422-ФЗ РФ</div>
                </div>
                <div class="text-slate-400 shrink-0">
                    ИНН: <span class="font-mono text-white">731303201073</span>
                </div>
            </div>
        </div>

        <!-- 1. Предмет -->
        <div class="mb-8">
            <h2 class="text-base font-bold text-white mb-2 flex items-center gap-2">
                <span class="text-emerald-400">1.</span> Предмет договора
            </h2>
            <p class="text-xs sm:text-sm text-slate-300 leading-relaxed mb-2">
                1.1. Исполнитель обязуется предоставить Заказчику удаленный облачный доступ к программному комплексу «RevOps OS», обеспечивающему 100% оцифровку, распознавание речи, оценку звонков менеджеров по 13 критериям B2B и отправку оперативных алертов в Telegram, а Заказчик обязуется принять и оплатить доступ в соответствии с выбранным Тарифным планом.
            </p>
            <p class="text-xs sm:text-sm text-slate-300 leading-relaxed">
                1.2. Подключение осуществляется через официальные API к amoCRM или Битрикс24 Заказчика в течение 1 рабочего дня.
            </p>
        </div>

        <!-- 2. Сетка Тарифов -->
        <div class="mb-8">
            <h2 class="text-base font-bold text-white mb-4 flex items-center gap-2">
                <span class="text-emerald-400">2.</span> Тарифные планы и стоимость услуг
            </h2>
            
            <div class="overflow-x-auto rounded-xl border border-slate-800">
                <table class="w-full text-left text-xs sm:text-sm">
                    <thead class="bg-slate-900/90 text-slate-300 font-bold border-b border-slate-800">
                        <tr>
                            <th class="p-3.5">Тарифный план</th>
                            <th class="p-3.5">Стоимость</th>
                            <th class="p-3.5">Объем ОП</th>
                            <th class="p-3.5">Включенный функционал</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-800/60 text-slate-300">
                        <tr class="bg-emerald-950/10 hover:bg-emerald-950/20">
                            <td class="p-3.5 font-bold text-white">
                                7-дневный Пилот
                                <span class="block text-[11px] text-emerald-400 font-normal">Тест с гарантией</span>
                            </td>
                            <td class="p-3.5 font-bold text-emerald-400 font-mono">14 900 ₽<br><span class="text-[10px] text-slate-400 font-normal">разово за 7 дней</span></td>
                            <td class="p-3.5">до 10 чел</td>
                            <td class="p-3.5 text-xs">100% аудит звонков за неделю, алерты за 60 сек, дайджест в 09:00, PDF-карта сливов, <strong class="text-emerald-300">100% гарантия возврата денег</strong>.</td>
                        </tr>
                        <tr class="hover:bg-slate-900/40">
                            <td class="p-3.5 font-bold text-white">Тариф «Старт»</td>
                            <td class="p-3.5 font-bold text-slate-200 font-mono">29 000 ₽ / мес<br><span class="text-[10px] text-slate-500 font-normal">23 200 ₽ при оплате за год</span></td>
                            <td class="p-3.5">до 5 чел</td>
                            <td class="p-3.5 text-xs">100% аудит всех звонков, интеграция с 1 CRM (amo/Б24), утренний дайджест РОПу в Telegram, базовый чек-лист 13 критериев.</td>
                        </tr>
                        <tr class="bg-indigo-950/10 hover:bg-indigo-950/20">
                            <td class="p-3.5 font-bold text-white">
                                Тариф «Рост» (Scale)
                                <span class="block text-[11px] text-indigo-400 font-normal">Самый популярный</span>
                            </td>
                            <td class="p-3.5 font-bold text-indigo-400 font-mono">59 000 ₽ / мес<br><span class="text-[10px] text-slate-500 font-normal">47 200 ₽ при оплате за год</span></td>
                            <td class="p-3.5">до 15 чел</td>
                            <td class="p-3.5 text-xs">Все из «Старт» + мгновенные Telegram-алерты за 60 сек на срыв сделок, отработка «Дорого», кастомный чеклист, еженедельный спринт с RevOps-экспертом.</td>
                        </tr>
                        <tr class="hover:bg-slate-900/40">
                            <td class="p-3.5 font-bold text-white">Тариф «Корпоративный»</td>
                            <td class="p-3.5 font-bold text-cyan-400 font-mono">89 000 ₽ / мес<br><span class="text-[10px] text-slate-500 font-normal">71 200 ₽ при оплате за год</span></td>
                            <td class="p-3.5">до 30+ чел</td>
                            <td class="p-3.5 text-xs">Все из «Рост» + персональный выделенный RevOps-аналитик, выделенный сервер 152-ФЗ, интеграция с несколькими CRM/филиалами, SLA 99.9%.</td>
                        </tr>
                    </tbody>
                </table>
            </div>
            <p class="text-xs text-slate-500 mt-2">
                * НДС не облагается в связи с применением Исполнителем специального налогового режима НПД (ст. 2 422-ФЗ) либо УСН (п. 2 ст. 346.11 НК РФ). При оплате за 12 месяцев предоставляется скидка 20%.
            </p>
        </div>

        <!-- 3. Гарантия возврата -->
        <div class="mb-8 bg-emerald-950/20 rounded-xl p-5 border border-emerald-500/30">
            <h2 class="text-base font-bold text-emerald-400 mb-2 flex items-center gap-2">
                <i class="fas fa-shield-halved"></i> 3. Безусловная 100% гарантия возврата по Пилоту (Money-Back Guarantee)
            </h2>
            <p class="text-xs sm:text-sm text-slate-300 leading-relaxed mb-2">
                3.1. Для Тарифа «7-дневный Пилотный спринт» действует 100% безусловная гарантия окупаемости.
            </p>
            <p class="text-xs sm:text-sm text-slate-300 leading-relaxed">
                3.2. Если по итогам 7 дней система не обнаружит упущенную выручку или Заказчик сочтет результаты пилота неудовлетворительными, <strong>Исполнитель возвращает 100% уплаченной суммы (14 900 ₽) на расчетный счет Заказчика в течение 3 банковских дней</strong> по первому запросу в Telegram (<a href="https://t.me/dm1918" class="text-emerald-400 underline">@dm1918</a>) или на email, без штрафов и удержаний.
            </p>
        </div>

        <!-- 4. Порядок оплаты и отчетность -->
        <div class="mb-8">
            <h2 class="text-base font-bold text-white mb-2 flex items-center gap-2">
                <span class="text-emerald-400">4.</span> Порядок расчетов и закрывающие документы
            </h2>
            <p class="text-xs sm:text-sm text-slate-300 leading-relaxed mb-2">
                4.1. Оплата производится Заказчиком в форме 100% предоплаты безналичным переводом на расчетный счет или корпоративной картой.
            </p>
            <p class="text-xs sm:text-sm text-slate-300 leading-relaxed mb-2">
                4.2. <strong>При расчетах в статусе Самозанятого (НПД):</strong> Исполнитель в день зачисления оплаты формирует фискальный чек через сервис ФНС «Мой налог» с указанием ИНН Заказчика. На основании чека Заказчик в полном объеме учитывает расходы для налога на прибыль / УСН (п. 8 ст. 15 422-ФЗ).
            </p>
            <p class="text-xs sm:text-sm text-slate-300 leading-relaxed">
                4.3. <strong>При расчетах в статусе ИП:</strong> Исполнитель выставляет счет и ежемесячный Акт сдачи-приемки через систему ЭДО (Диадок / СБИС) или на email.
            </p>
        </div>

        <!-- 5. 152-ФЗ Безопасность -->
        <div class="mb-8">
            <h2 class="text-base font-bold text-white mb-2 flex items-center gap-2">
                <span class="text-emerald-400">5.</span> Безопасность и соответствие 152-ФЗ РФ
            </h2>
            <p class="text-xs sm:text-sm text-slate-300 leading-relaxed mb-2">
                5.1. Стороны соблюдают конфиденциальность в отношении клиентской базы и записей звонков Заказчика.
            </p>
            <p class="text-xs sm:text-sm text-slate-300 leading-relaxed">
                5.2. Обработка данных осуществляется на серверах в Российской Федерации (Selectel, Москва/СПб). Все персональные данные (ФИО, телефоны) автоматически токенизируются и деперсонализируются.
            </p>
        </div>

        <!-- 6. Реквизиты Сторон -->
        <div class="pt-6 border-t border-slate-800">
            <h2 class="text-base font-bold text-white mb-4">6. Реквизиты Исполнителя</h2>
            
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs">
                <div class="p-4 rounded-xl bg-slate-900/60 border border-slate-800">
                    <span class="font-bold text-emerald-400 block mb-2">ВАРИАНТ А: САМОЗАНЯТЫЙ (ТЕКУЩИЙ)</span>
                    <p class="leading-relaxed text-slate-300">
                        <strong>Плательщик НПД:</strong> Федотов Дмитрий<br>
                        <strong>ИНН:</strong> 731303201073<br>
                        <strong>Банк:</strong> АО «ТИНЬКОФФ БАНК»<br>
                        <strong>Расчетный счет:</strong> 40802810500003849120<br>
                        <strong>БИК:</strong> 044525974<br>
                        <strong>К/с:</strong> 30101810145250000974<br>
                        <strong>Документ:</strong> Чек ФНС РФ (422-ФЗ)
                    </p>
                </div>
                <div class="p-4 rounded-xl bg-slate-900/60 border border-slate-800">
                    <span class="font-bold text-cyan-400 block mb-2">ВАРИАНТ Б: ИНДИВИДУАЛЬНЫЙ ПРЕДПРИНИМАТЕЛЬ</span>
                    <p class="leading-relaxed text-slate-300">
                        <strong>ИП:</strong> Федотов Дмитрий<br>
                        <strong>ОГРНИП:</strong> [в процессе регистрации]<br>
                        <strong>ИНН:</strong> 731303201073 (УСН 6%)<br>
                        <strong>Банк:</strong> АО «ТИНЬКОФФ БАНК»<br>
                        <strong>Расчетный счет:</strong> 40802810500003849120<br>
                        <strong>БИК:</strong> 044525974<br>
                        <strong>Telegram / MAX:</strong> <a href="https://t.me/dm1918" class="text-emerald-400">@dm1918</a><br>
                        <strong>Email:</strong> founder@ai-rop.ru
                    </p>
                </div>
            </div>
        </div>

        <div class="mt-8 pt-4 border-t border-slate-900 text-center text-xs text-slate-500">
            &copy; 2026 RevOps OS Pro (ai-rop.ru). Все права защищены.
        </div>
    </div>
</body>
</html>"""
    output_path.write_text(html, encoding="utf-8")
    print(f"✓ HTML agreement generated: {output_path}")

def main():
    docs_dir = Path("docs")
    docs_dir.mkdir(exist_ok=True)
    
    # 1. Generate DOCX in docs/
    create_docx(docs_dir / "Договор_Оферта_RevOps_OS_Все_Тарифы.docx")

    # 2. Generate HTML in docs/
    create_html(docs_dir / "Договор_Оферта_RevOps_OS_Все_Тарифы.html")

    # 3. Generate offer.html in root (for web live deployment at ai-rop.ru/offer.html)
    create_html(Path("offer.html"))

if __name__ == "__main__":
    main()
