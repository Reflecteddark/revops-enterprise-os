import os
import sys
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

sys.stdout.reconfigure(encoding="utf-8")

def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=140, bottom=140, left=180, right=180):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

doc = Document()

# Page Margins
sections = doc.sections
for s in sections:
    s.top_margin = Inches(0.7)
    s.bottom_margin = Inches(0.7)
    s.left_margin = Inches(0.7)
    s.right_margin = Inches(0.7)

# Title
title_p = doc.add_paragraph()
title_p.paragraph_format.space_before = Pt(0)
title_p.paragraph_format.space_after = Pt(4)
run_title = title_p.add_run("🎯 БОЕВОЙ СЦЕНАРИЙ ДЕМОНСТРАЦИИ ПО ЖИВОЙ ТАБЛИЦЕ")
run_title.font.name = "Segoe UI"
run_title.font.size = Pt(20)
run_title.font.bold = True
run_title.font.color.rgb = RGBColor(15, 23, 42)

sub_p = doc.add_paragraph()
sub_p.paragraph_format.space_after = Pt(14)
run_sub = sub_p.add_run("Пошаговый 20-минутный регламент спикера для продаж через Master-таблицу RevOps OS Pro (V17.6)")
run_sub.font.name = "Segoe UI"
run_sub.font.size = Pt(11)
run_sub.font.italic = True
run_sub.font.color.rgb = RGBColor(71, 85, 105)

# Banner Table
t_info = doc.add_table(rows=1, cols=1)
t_info.alignment = WD_TABLE_ALIGNMENT.CENTER
c_info = t_info.cell(0, 0)
set_cell_background(c_info, "0F172A")
set_cell_margins(c_info, 180, 180, 220, 220)
p_info = c_info.paragraphs[0]
p_info.paragraph_format.space_after = Pt(0)
r_info = p_info.add_run("🔗 Ссылка на демонстрационную таблицу: https://docs.google.com/spreadsheets/d/1QnjrrbpqhYssofchee7G06szWrFvBCcOqGIVokjqdVc\n"
                       "⏱️ Длительность демо: ровно 20 минут • Формат: показ экрана в Zoom/Google Meet • Масштаб: 90–100%")
r_info.font.name = "Segoe UI"
r_info.font.size = Pt(9.5)
r_info.font.bold = True
r_info.font.color.rgb = RGBColor(255, 255, 255)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Section 1: Pre-call setup
p_h1 = doc.add_paragraph()
r_h1 = p_h1.add_run("1. ПОДГОТОВКА К ДЕМОНСТРАЦИИ (ЗА 2 МИНУТЫ ДО СОЗВОНА)")
r_h1.font.name = "Segoe UI"
r_h1.font.size = Pt(13)
r_h1.font.bold = True
r_h1.font.color.rgb = RGBColor(30, 58, 138)

prep_points = [
    ("Открыть ссылку на таблицу:", "Убедитесь, что открыта первая вкладка: '⚡ Экспресс_Калькулятор_3_Цифры'."),
    ("Настроить масштаб экрана:", "Поставьте масштаб браузера 90% или 100%, чтобы дашборд помещался без горизонтальной полосы прокрутки."),
    ("Очистить поля калькулятора:", "В ячейках B5, B6, B7 должны стоять базовые значения (150 лидов, 400 000 ₽ чек, 4 менеджера)."),
    ("Закрыть лишние вкладки:", "Никаких личных мессенджеров и посторонних документов при расшаривании экрана.")
]
for title, desc in prep_points:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.left_indent = Inches(0.2)
    r1 = p.add_run(f"• {title} ")
    r1.font.name = "Segoe UI"
    r1.font.size = Pt(10)
    r1.font.bold = True
    r2 = p.add_run(desc)
    r2.font.name = "Segoe UI"
    r2.font.size = Pt(10)

doc.add_paragraph().paragraph_format.space_after = Pt(8)

# Section 2: 5 Steps walkthrough
p_h2 = doc.add_paragraph()
r_h2 = p_h2.add_run("2. ДРАМАТУРГИЯ ПОКАЗА: 5 ШАГОВ ОТ БОЛИ К ЗАКРЫТИЮ СДЕЛКИ")
r_h2.font.name = "Segoe UI"
r_h2.font.size = Pt(13)
r_h2.font.bold = True
r_h2.font.color.rgb = RGBColor(30, 58, 138)

steps = [
    {
        "tab": "ВКЛАДКА 1: ⚡ Экспресс_Калькулятор_3_Цифры (Минуты 0:00 – 4:00)",
        "goal": "Крючок внимания и оцифровка скрытых финансовых потерь в бизнесе клиента на лету.",
        "action": "Вбить в ячейки B5, B6, B7 реальные цифры клиента прямо во время разговора.",
        "speech": "«[Имя], чтобы не тратить время на абстрактные слайды, давайте сразу на ваших реальных цифрах за 30 секунд посмотрим математику вашего отдела продаж. Назовите всего 3 цифры:\n"
                  "1. Сколько входящих целевых лидов отдел получает в месяц?\n"
                  "2. Какой средний чек завершенной сделки?\n"
                  "3. Сколько менеджеров сейчас работает в отделе?\n\n"
                  "[Спикер вбивает цифры в B5, B6, B7 — таблица пересчитывается за 1 секунду].\n\n"
                  "Посмотрите на экран: вот что происходит с вашими деньгами прямо сейчас:\n"
                  "• 28% выручки сгорает на звонках (ячейка D15): менеджеры консультируют клиентов, но отпускают их без фиксации даты следующего шага.\n"
                  "• 18% выручки теряется на зависании КП (D16): предложения отправляются на 'кладбище почты' без звонка через 24 часа.\n"
                  "• Итого ваша упущенная прибыль — [сумма из D19, например 1 122 000 ₽] каждый месяц.\n"
                  "• А вот потенциал быстрого возврата в кассу (F19): устранив базовые ошибки, вы за первый месяц возвращаете от [сумма из F19, например 392 700 ₽].\n\n"
                  "А теперь давайте покажу, как ваш РОП за 15 минут в день видит всю эту картину и спасает сделки до того, как деньги утекли. Переходим на Пульт РОПа...»"
    },
    {
        "tab": "ВКЛАДКА 2: 📋 Пульт_РОПа_15_Минут (Минуты 4:00 – 8:00)",
        "goal": "Показать решение: как превратить многочасовую рутину прослушки звонков в 15 минут точечного контроля.",
        "action": "Показать строки 4-5 (сводка утра) и строки 8-13 (ТОП-5 горящих сделок). Кликнуть на ячейку H9 (готовый скрипт).",
        "speech": "«Вот реальный рабочий экран РОПа на каждое утро. Ему не нужно отслушивать 50 звонков подряд.\n"
                  "В строке 4 он сразу видит:\n"
                  "• За вчера прошло звонков: 6. Из них с браком: 2 диалога.\n"
                  "• Сумма в зоне риска прямо сейчас: 2 730 000 ₽.\n\n"
                  "Ниже — радар из 5 конкретных сделок, где клиенты готовы сорваться:\n"
                  "• Сделка D-104 (850 000 ₽): висит на этапе КП уже 12 дней без движения.\n"
                  "• Сделка D-109 (600 000 ₽): менеджер сорвал контакт и назвал неверные условия.\n\n"
                  "И самое главное — посмотрите на колонку H (ячейка H9). РОПу даже не нужно придумывать, что написать. Система уже составила готовое персональное сообщение для перехвата сделки! РОП отправляет его в WhatsApp клиенту или делает 3-минутный звонок. Результат: за 15 минут защищено 2.7 млн ₽ выручки.\n\n"
                  "Вы спросите: а как ИИ узнает, что происходит в звонках? Показываю ядро речевой аналитики — вкладка 'ИИ_Аудит'...»"
    },
    {
        "tab": "ВКЛАДКА 3: 🎙️ ИИ_Аудит (Минуты 8:00 – 13:00)",
        "goal": "Доказать надежность нейросети, показать 13 стандартов B2B-речи и снять страх саботажа менеджеров.",
        "action": "Показать лидерборд менеджеров (строки 7–13), кликнуть на цитаты срывов (строка 31) и объяснить Режим «Адвокат».",
        "speech": "«Здесь нейросеть Whisper Pro + DeepSeek расшифровывает 100% звонков вашего отдела без участия людей.\n"
                  "Сверху — рейтинг менеджеров:\n"
                  "• Мельников закрывает Next Step в 100% случаев, средний балл 12.2 из 13.\n"
                  "• Попов фиксирует дату контакта лишь в 25% звонков (балл 5.5). Именно он сжигает 75% вашего рекламного трафика!\n\n"
                  "Ниже — точные цитаты из реальных диалогов (строка 31):\n"
                  "Клиент спрашивает про смету, а менеджер говорит: 'Ну я не знаю, у нас такие цены'. РОП сразу видит цитату и тайминг аудио.\n\n"
                  "И критически важный момент: ПАКЕТ АНТИСАБОТАЖ.\n"
                  "Почему менеджеры не будут против системы? Потому что RevOps OS:\n"
                  "1. Сам заполняет карточки сделок в CRM за менеджера — экономит 40 минут рутины в день.\n"
                  "2. Работает как адвокат продавца: если клиент попался неадекватный или токсичный, ИИ подтвердит РОПу, что менеджер действовал по стандарту, и защитит от штрафа.\n"
                  "3. Присылает подсказки по дожиму в личные сообщения менеджеру, помогая ему перевыполнить план и заработать больше.\n\n"
                  "А теперь давайте взглянем на финансовый результат внедрения — вкладка 'Диагностика утечек'...»"
    },
    {
        "tab": "ВКЛАДКА 4: 💸 Диагностика_Утечек_ОП (Минуты 13:00 – 17:00)",
        "goal": "Обосновать окупаемость проекта в цифрах: 367% ROI и возврат инвестиций за 6–14 дней.",
        "action": "Показать 7 смертных грехов (строки 8–15) и финансовую экономику проекта (строки 18–24).",
        "speech": "«В этой таблице оцифрованы 7 смертных грехов отдела продаж: от 'черного ящика CRM' до зависшей дебиторки.\n"
                  "Но самое главное — блок экономики внедрения (строки 18–24):\n"
                  "• Стоимость проекта оптимизации ОП: 450 000 ₽ (спринт 'под ключ' с настройкой и сопровождением).\n"
                  "• Ожидаемый возврат выручки в кассу в первый месяц: 2 100 300 ₽.\n"
                  "• Чистая прибыль собственника за вычетом всех расходов: 1 650 300 ₽.\n"
                  "• Подтвержденный ROI: 367%.\n"
                  "• Срок полной окупаемости проекта: всего 6–14 рабочих дней отдела продаж!\n\n"
                  "Для первого лица компании у нас есть сводный пульт — Executive OnePager...»"
    },
    {
        "tab": "ВКЛАДКА 5: 📄 Executive_OnePager (Минуты 17:00 – 18:00)",
        "goal": "Показать собственнику и генеральному директору единое окно контроля всей коммерческой службы компании.",
        "action": "Продемонстрировать ключевые KPI: выручка факт, прогноз кассы, воронка конверсий и ТОП-риски.",
        "speech": "«Собственнику и генеральному директору не нужно открывать 10 разных отчетов. В одном экране:\n"
                  "• Выполнение месячного плана в процентах.\n"
                  "• Взвешенный прогноз кассы (SSOT) с учетом реальной вероятности закрытия сделок.\n"
                  "• Узкие горлышки воронки и главные угрозы кассового разрыва.\n"
                  "Все данные собираются автоматически и синхронизируются в реальном времени»."
    }
]

for s in steps:
    p_tab = doc.add_paragraph()
    p_tab.paragraph_format.space_before = Pt(8)
    p_tab.paragraph_format.space_after = Pt(2)
    r_tab = p_tab.add_run(s["tab"])
    r_tab.font.name = "Segoe UI"
    r_tab.font.size = Pt(11)
    r_tab.font.bold = True
    r_tab.font.color.rgb = RGBColor(5, 150, 105)

    p_g = doc.add_paragraph()
    p_g.paragraph_format.space_after = Pt(2)
    p_g.paragraph_format.left_indent = Inches(0.2)
    rg1 = p_g.add_run("🎯 Цель шага: ")
    rg1.font.bold = True
    rg1.font.size = Pt(9.5)
    rg2 = p_g.add_run(s["goal"])
    rg2.font.size = Pt(9.5)

    p_a = doc.add_paragraph()
    p_a.paragraph_format.space_after = Pt(4)
    p_a.paragraph_format.left_indent = Inches(0.2)
    ra1 = p_a.add_run("🖱️ Действие спикера: ")
    ra1.font.bold = True
    ra1.font.size = Pt(9.5)
    ra2 = p_a.add_run(s["action"])
    ra2.font.size = Pt(9.5)

    # Box for Speech Script
    t_box = doc.add_table(rows=1, cols=1)
    t_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_box = t_box.cell(0, 0)
    set_cell_background(c_box, "F8FAFC")
    set_cell_margins(c_box, 140, 140, 180, 180)
    p_box = c_box.paragraphs[0]
    p_box.paragraph_format.space_after = Pt(0)
    r_box_title = p_box.add_run("💬 РЕЧЕВОЙ СКРИПТ (СЛОВО В СЛОВО):\n")
    r_box_title.font.name = "Segoe UI"
    r_box_title.font.size = Pt(9)
    r_box_title.font.bold = True
    r_box_title.font.color.rgb = RGBColor(30, 58, 138)

    r_box_text = p_box.add_run(s["speech"])
    r_box_text.font.name = "Segoe UI"
    r_box_text.font.size = Pt(9)
    r_box_text.font.italic = True
    r_box_text.font.color.rgb = RGBColor(30, 41, 59)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

# Section 3: Closing Offer
p_h3 = doc.add_paragraph()
p_h3.paragraph_format.space_before = Pt(10)
r_h3 = p_h3.add_run("3. ФИНАЛЬНЫЙ ОФФЕР И ЗАКРЫТИЕ (МИНУТЫ 18:00 – 20:00)")
r_h3.font.name = "Segoe UI"
r_h3.font.size = Pt(13)
r_h3.font.bold = True
r_h3.font.color.rgb = RGBColor(30, 58, 138)

p_close = doc.add_paragraph()
p_close.paragraph_format.space_after = Pt(6)
p_close.paragraph_format.left_indent = Inches(0.2)
r_close = p_close.add_run(
    "«[Имя], мы не предлагаем верить цифрам на слово. Самый простой и надежный шаг прямо сейчас — проверить эту математику на реальных диалогах вашей компании.\n\n"
    "👉 БЕСПЛАТНЫЙ ЭКСПРЕСС-ТЕСТ НА 3 ВАШИХ ЗВОНКАХ (0 ₽):\n"
    "Вы скидываете в наш Telegram-бот 3 аудиозаписи любых вчерашних звонков ваших менеджеров.\n"
    "Нейросеть бесплатно разберет их по 13 критериям, покажет дефекты речи и оцифрует точную сумму финансового риска в рублях.\n"
    "Срок разбора — 15 минут. Никаких оплат и сложных интеграций.\n\n"
    "Готовы прямо сейчас после созвона отправить 3 звонка в Telegram?»"
)
r_close.font.name = "Segoe UI"
r_close.font.size = Pt(10)
r_close.font.bold = True
r_close.font.color.rgb = RGBColor(15, 23, 42)

# Section 4: Objections
p_h4 = doc.add_paragraph()
p_h4.paragraph_format.space_before = Pt(10)
r_h4 = p_h4.add_run("4. ОТРАБОТКА КАВЕРЗНЫХ ВОПРОСОВ ВО ВРЕМЯ ПОКАЗА ТАБЛИЦЫ")
r_h4.font.name = "Segoe UI"
r_h4.font.size = Pt(13)
r_h4.font.bold = True
r_h4.font.color.rgb = RGBColor(30, 58, 138)

objections = [
    ("«Почему это сделано в таблице Google / Excel, а не в SaaS-кабинете?»",
     "«Вся обработка нейросетью Whisper и DeepSeek идет в облачном микросервисе на бэкенде. Но витрину мы вывели в Google Sheets и Telegram, потому что это дает 100% гибкость: вы владеете данными, можете настроить любые формулы под себя без ожидания программистов, а РОП получает алерты в привычный Telegram без необходимости логиниться в очередной тяжелый софт»."),
    ("«Чем вы лучше сервисов речевой аналитики вроде Imot.io?»",
     "«1. Imot.io берет плату за минуты (от 54 до 400 тыс. ₽/мес) — чем дольше менеджеры говорят, тем выше счет. У нас — фиксированная абонплата и полный безлимит минут.\n"
     "2. Imot.io дает отчеты постфактум через 24–48 часов. Наш бот присылает алерт через 30 секунд со сценарием спасения сделки в карман РОПа.\n"
     "3. Imot.io воспринимается продавцами как надзиратель со штрафами. RevOps OS внедряется как адвокат, экономящий 40 минут в день на автозаполнении CRM»."),
    ("«Сложно ли подключить нашу CRM (amoCRM / Битрикс24)?»",
     "«Подключение занимает 15–20 минут по безопасному API-токену. Мы ничего не ломаем в вашей воронке. Система просто начинает считывать записи звонков и возвращать Executive Summary в карточку сделки»."),
    ("«Безопасно ли это по 152-ФЗ о персональных данных?»",
     "«Абсолютно. Вся обработка идет на серверах в РФ. Перед отправкой в языковую модель все аудиозаписи автоматически деперсонализируются: вырезаются паспортные данные, телефоны и персональные идентификаторы»."),
    ("«Что если менеджеры откажутся звонить под записью?»",
     "«В нашей системе активирован Режим «Адвокат»: менеджеры первыми выступают за систему, потому что она защищает их от неадекватных клиентов, освобождает от ручного заполнения полей и дает подсказки, как дожать сделку и поднять личную премию»")
]

for q, a in objections:
    p_q = doc.add_paragraph()
    p_q.paragraph_format.space_before = Pt(4)
    p_q.paragraph_format.space_after = Pt(2)
    p_q.paragraph_format.left_indent = Inches(0.2)
    rq = p_q.add_run(f"❓ {q}")
    rq.font.name = "Segoe UI"
    rq.font.size = Pt(9.5)
    rq.font.bold = True
    rq.font.color.rgb = RGBColor(185, 28, 28)

    p_a = doc.add_paragraph()
    p_a.paragraph_format.space_after = Pt(6)
    p_a.paragraph_format.left_indent = Inches(0.4)
    ra = p_a.add_run(f"💡 {a}")
    ra.font.name = "Segoe UI"
    ra.font.size = Pt(9.5)
    ra.font.color.rgb = RGBColor(51, 65, 85)

out_path = "docs/Шпаргалка_Спикера_Демо_по_Живой_Таблице.docx"
doc.save(out_path)
print(f"Successfully generated: {out_path}")
