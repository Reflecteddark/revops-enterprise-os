"""
Скрипт добавления листа '📢 Каналы_Продаж' в RevOps_Founder_Workspace_Planner.xlsx
Реализует Спринт 3: Платный трафик в Telegram для отделов продаж и CRM.
Включает:
- Топовые KPI медиаплана (бюджет, охват, заявки, ROI)
- Базу 20 отобранных каналов по 4 категориям с персональными UTM-ссылками
- 3 готовых креатива (Комбо 'Продажи+CRM', Кейс-разбор слива 340к, Переманивание с imot.io)
- Инструкцию по проверке каналов через TGStat и торг на скидку 10-20%
"""

import sys
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")


def add_channels_traffic_sheet(file_path: Path):
    if not file_path.exists():
        print(f"File not found: {file_path}")
        return

    wb = openpyxl.load_workbook(file_path)
    sheet_name = "📢 Каналы_Продаж"

    if sheet_name in wb.sheetnames:
        del wb[sheet_name]

    # Создаем лист после Спринтов или в удобной позиции
    ws = wb.create_sheet(title=sheet_name)
    ws.views.sheetView[0].showGridLines = True

    # Цветовая палитра
    c_navy_dark = "0F172A"       # #0F172A
    c_navy_header = "1E293B"     # #1E293B
    c_blue_primary = "1E3A8A"    # #1E3A8A
    c_blue_light = "EFF6FF"      # #EFF6FF
    c_emerald = "059669"         # #059669
    c_emerald_light = "ECFDF5"   # #ECFDF5
    c_amber = "D97706"           # #D97706
    c_amber_light = "FFFBEB"     # #FFFBEB
    c_rose = "E11D48"            # #E11D48
    c_rose_light = "FFF1F2"      # #FFF1F2
    c_purple = "7C3AED"          # #7C3AED
    c_purple_light = "F5F3FF"    # #F5F3FF
    c_card_bg = "F8FAFC"        # #F8FAFC
    c_border = "CBD5E1"          # #CBD5E1

    fill_navy = PatternFill(start_color=c_navy_dark, end_color=c_navy_dark, fill_type="solid")
    fill_header = PatternFill(start_color=c_blue_primary, end_color=c_blue_primary, fill_type="solid")
    fill_sub_hdr = PatternFill(start_color=c_navy_header, end_color=c_navy_header, fill_type="solid")
    fill_zebra = PatternFill(start_color=c_card_bg, end_color=c_card_bg, fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_blue_chip = PatternFill(start_color=c_blue_light, end_color=c_blue_light, fill_type="solid")
    fill_emerald_chip = PatternFill(start_color=c_emerald_light, end_color=c_emerald_light, fill_type="solid")
    fill_amber_chip = PatternFill(start_color=c_amber_light, end_color=c_amber_light, fill_type="solid")
    fill_purple_chip = PatternFill(start_color=c_purple_light, end_color=c_purple_light, fill_type="solid")

    font_title = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
    font_sub = Font(name="Segoe UI", size=9, italic=True, color="94A3B8")
    font_header = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
    font_sec = Font(name="Segoe UI", size=11, bold=True, color="0F172A")
    font_bold = Font(name="Segoe UI", size=9, bold=True, color="1E293B")
    font_reg = Font(name="Segoe UI", size=8.5, color="334155")
    font_code = Font(name="Consolas", size=8, color="0F172A")
    font_link = Font(name="Segoe UI", size=8.5, color="2563EB", underline="single")

    thin_border = Border(
        left=Side(style="thin", color=c_border),
        right=Side(style="thin", color=c_border),
        top=Side(style="thin", color=c_border),
        bottom=Side(style="thin", color=c_border)
    )

    # 1. Заголовок
    ws.merge_cells("A1:J1")
    ws["A1"] = "📢 RevOps Enterprise OS | Медиаплан Платного Трафика в Telegram (B2B Sales & CRM)"
    ws["A1"].font = font_title
    ws["A1"].fill = fill_navy
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 34

    ws.merge_cells("A2:J2")
    ws["A2"] = "База отобранных Telegram-каналов для посевов, форматы 1/24 и 2/48, UTM-генератор, 3 готовых креатива и трекер окупаемости (ROI)"
    ws["A2"].font = font_sub
    ws["A2"].fill = fill_navy
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 18

    # 2. Карточки KPI Закупки
    kpi_cards = [
        ("💰 БЮДЖЕТ ТЕСТА", "35 000 – 60 000 ₽", fill_blue_chip, "1E3A8A"),
        ("👁️ ЦЕЛЕВОЙ ОХВАТ", "15 000 – 35 000 Views", fill_purple_chip, "7C3AED"),
        ("🎯 КЛИКИ НА САЙТ (CTR ~2.5%)", "450 – 850 кликов", fill_blue_chip, "2563EB"),
        ("📥 ЗАЯВКИ НА АУДИТ (0 ₽)", "25 – 45 заявок", fill_amber_chip, "D97706"),
        ("🚀 ЗАКРЫТЫЕ ПИЛОТЫ (ROI)", "3 – 6 сделок (165к-360к ₽)", fill_emerald_chip, "059669")
    ]

    for idx, (t, val, fill_c, font_c) in enumerate(kpi_cards):
        col_start = idx * 2 + 1
        col_end = col_start + 1
        c_s_letter = get_column_letter(col_start)
        c_e_letter = get_column_letter(col_end)

        ws.merge_cells(f"{c_s_letter}4:{c_e_letter}4")
        cell_t = ws[f"{c_s_letter}4"]
        cell_t.value = t
        cell_t.font = Font(name="Segoe UI", size=8.5, bold=True, color="64748B")
        cell_t.fill = fill_c
        cell_t.alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells(f"{c_s_letter}5:{c_e_letter}5")
        cell_v = ws[f"{c_s_letter}5"]
        cell_v.value = val
        cell_v.font = Font(name="Segoe UI", size=10.5, bold=True, color=font_c)
        cell_v.fill = fill_white
        cell_v.alignment = Alignment(horizontal="center", vertical="center")

        ws.cell(4, col_start).border = thin_border
        ws.cell(4, col_end).border = thin_border
        ws.cell(5, col_start).border = thin_border
        ws.cell(5, col_end).border = thin_border

    ws.row_dimensions[4].height = 20
    ws.row_dimensions[5].height = 26

    # 3. Раздел 1: База Каналов
    ws.merge_cells("A7:J7")
    ws["A7"] = "🎯 РАЗДЕЛ 1: БАЗА 20 ОТОБРАННЫХ TELEGRAM-КАНАЛОВ ДЛЯ ЗАКУПА (B2B, CRM, РОПы, БИЗНЕС-РАЗБОРЫ)"
    ws["A7"].font = font_sec
    ws["A7"].fill = fill_blue_chip
    ws.row_dimensions[7].height = 24

    headers_ch = [
        "Код", "Категория", "Канал и Ссылка", "Целевая аудитория и специфика",
        "Формат", "Ориент. цена", "Креатив", "Персональная UTM-ссылка", "Статус контакта", "Переходы / Заявки"
    ]

    ws.row_dimensions[8].height = 26
    for c_i, h in enumerate(headers_ch, 1):
        cell = ws.cell(8, c_i, h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    channels_data = [
        # Категория 1: Продажи в B2B и РОПы
        ("TG-01", "B2B Продажи", "Заметки продавца B2B (@Salesnotes)", "Сложные продажи, квалификация лидов, РОПы и КД", "2/48", "8 000 – 12 000 ₽", "Пост 1 (CRM+Продажи)", "https://ai-rop.ru/?utm_source=tg&utm_medium=salesnotes", "В плане на контакт", "0 / 0"),
        ("TG-02", "B2B Продажи", "План продаж — М. Графский (@grafsky_b2b)", "Системные руководители ОП, регламенты и скрипты", "1/24", "10 000 – 15 000 ₽", "Пост 2 (Кейс 340к)", "https://ai-rop.ru/?utm_source=tg&utm_medium=grafsky_b2b", "В плане на контакт", "0 / 0"),
        ("TG-03", "B2B Продажи", "Продажник +1 (@floor_99)", "Практические кейсы, холодные звонки, дожим сделок", "2/48", "6 000 – 9 000 ₽", "Пост 1 (CRM+Продажи)", "https://ai-rop.ru/?utm_source=tg&utm_medium=floor_99", "В плане на контакт", "0 / 0"),
        ("TG-04", "B2B Продажи", "Ерохин про B2B продажи (@yerokhin_b2b)", "Топ-менеджмент B2B с длинными циклами и чеками", "2/48", "12 000 – 18 000 ₽", "Пост 2 (Кейс 340к)", "https://ai-rop.ru/?utm_source=tg&utm_medium=yerokhin_b2b", "В плане на контакт", "0 / 0"),
        ("TG-05", "B2B Продажи", "Охотники за продажами (@avsoln)", "А. Солнцева: разборы скриптов, возражение 'дорого'", "1/24", "7 000 – 10 000 ₽", "Пост 1 (CRM+Продажи)", "https://ai-rop.ru/?utm_source=tg&utm_medium=avsoln", "В плане на контакт", "0 / 0"),

        # Категория 2: CRM и Автоматизация
        ("TG-06", "CRM / Интеграторы", "50 оттенков CRM (@crm50)", "Юрий Николаев (Битрикс24), фанаты автоматизации ОП", "Натив / 2/48", "12 000 – 18 000 ₽", "Пост 1 (CRM+Продажи)", "https://ai-rop.ru/?utm_source=tg&utm_medium=crm50", "В плане на контакт", "0 / 0"),
        ("TG-07", "CRM / Интеграторы", "Блог RocketSales (@rocketsales_blog)", "Крупнейший интегратор amoCRM, активный B2B", "2/48", "8 000 – 14 000 ₽", "Пост 3 (Атака на минуты)", "https://ai-rop.ru/?utm_source=tg&utm_medium=rocketsales", "В плане на контакт", "0 / 0"),
        ("TG-08", "CRM / Интеграторы", "ИНТРОvert | CRM (@introvert_crm)", "Внедрение сложных IT-решений поверх amoCRM", "2/48", "9 000 – 15 000 ₽", "Пост 3 (Атака на минуты)", "https://ai-rop.ru/?utm_source=tg&utm_medium=introvert", "В плане на контакт", "0 / 0"),
        ("TG-09", "CRM / Интеграторы", "Секреты Битрикс24 и amoCRM (@crm_secrets)", "РОПы и предприниматели в поиске ИИ-виджетов", "2/48", "5 000 – 8 000 ₽", "Пост 1 (CRM+Продажи)", "https://ai-rop.ru/?utm_source=tg&utm_medium=crm_secrets", "В плане на контакт", "0 / 0"),
        ("TG-10", "CRM / Маркетинг", "Комплето B2B (@completo_ru)", "B2B Sales Performance: маркетинг + CRM + продажи", "2/48", "15 000 – 22 000 ₽", "Пост 1 (CRM+Продажи)", "https://ai-rop.ru/?utm_source=tg&utm_medium=completo", "В плане на контакт", "0 / 0"),
        ("TG-11", "CRM / Лидген", "Андрей Шишкин про B2B лидген (@andrew_shishkin)", "Автоматизация воронки, экономия времени РОПа", "1/24", "6 000 – 9 000 ₽", "Пост 2 (Кейс 340к)", "https://ai-rop.ru/?utm_source=tg&utm_medium=andrew_shishkin", "В плане на контакт", "0 / 0"),
        ("TG-12", "CRM / Аналитика", "B2B-медиа | Продажи (@b2bmedia)", "Исследования воронки, конверсии коммерческих служб", "2/48", "7 000 – 11 000 ₽", "Пост 1 (CRM+Продажи)", "https://ai-rop.ru/?utm_source=tg&utm_medium=b2bmedia", "В плане на контакт", "0 / 0"),

        # Категория 3: Федеральные Эксперты (Уровень Гребенюка / Высоцкого)
        ("TG-13", "Бизнес-Разборы", "Resulting | Отделы продаж (@resulting_ru)", "Агентство М. Гребенюка: 100% концентрация РОПов", "По запросу", "50 000 – 90 000 ₽", "Пост 2 (Кейс 340к)", "https://ai-rop.ru/?utm_source=tg&utm_medium=resulting", "В плане на контакт", "0 / 0"),
        ("TG-14", "Бизнес-Разборы", "Бизнес-Разборы | Кейсы (@biz_razbory)", "Ошибки управления, сливы лидов в реалити-шоу", "2/48", "10 000 – 16 000 ₽", "Пост 2 (Кейс 340к)", "https://ai-rop.ru/?utm_source=tg&utm_medium=biz_razbory", "В плане на контакт", "0 / 0"),
        ("TG-15", "Системный Бизнес", "Дневник Предпринимателя (@vysotsky_consulting)", "Александр Высоцкий: выход из операционки", "1/24", "25 000 – 45 000 ₽", "Пост 1 (CRM+Продажи)", "https://ai-rop.ru/?utm_source=tg&utm_medium=vysotsky", "В плане на контакт", "0 / 0"),
        ("TG-16", "Системный Бизнес", "Oy-li | Екатерина Уколова (@oyli_ru)", "Чек-листы, регламенты ОП, жесткая CRM-аналитика", "2/48", "30 000 – 50 000 ₽", "Пост 1 (CRM+Продажи)", "https://ai-rop.ru/?utm_source=tg&utm_medium=oyli", "В плане на контакт", "0 / 0"),
        ("TG-17", "Менеджмент B2B", "Максим Батырев | Комбат (@batyrev_maxim)", "Армия российских РОПов и коммерческих директоров", "1/24", "40 000 – 70 000 ₽", "Пост 2 (Кейс 340к)", "https://ai-rop.ru/?utm_source=tg&utm_medium=batyrev", "В плане на контакт", "0 / 0"),
        ("TG-18", "ИИ в Продажах", "Сергей Кошечкин | Продажи (@koshechkin_sales)", "Жесткий практик, автоматизация и ИИ в коммерции", "2/48", "8 000 – 14 000 ₽", "Пост 1 (CRM+Продажи)", "https://ai-rop.ru/?utm_source=tg&utm_medium=koshechkin", "В плане на контакт", "0 / 0"),

        # Категория 4: Управление и Сколково
        ("TG-19", "Тяжелый B2B", "MANAGED / Иван Чирков (@managed_ru)", "Операционная эффективность, оцифровка процессов", "2/48", "9 000 – 14 000 ₽", "Пост 1 (CRM+Продажи)", "https://ai-rop.ru/?utm_source=tg&utm_medium=managed", "В плане на контакт", "0 / 0"),
        ("TG-20", "ИИ в Бизнесе", "ИИ в бизнесе | Нейросети (@ai_business_ru)", "Практика внедрения нейросетей в коммерческие отделы", "2/48", "5 000 – 8 000 ₽", "Пост 3 (Атака на минуты)", "https://ai-rop.ru/?utm_source=tg&utm_medium=ai_business", "В плане на контакт", "0 / 0")
    ]

    current_r = 9
    for row_data in channels_data:
        ws.row_dimensions[current_r].height = 24
        bg = fill_zebra if current_r % 2 == 1 else fill_white

        for c_i, val in enumerate(row_data, 1):
            cell = ws.cell(current_r, c_i, val)
            cell.font = font_reg
            cell.fill = bg
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center")

            if c_i == 1:
                cell.font = font_bold
                cell.fill = fill_blue_chip
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_i == 2:
                cell.font = font_bold
            elif c_i == 6:
                cell.font = font_bold
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_i == 7:
                cell.font = font_bold
                cell.fill = fill_amber_chip
            elif c_i == 8:
                cell.font = font_link
            elif c_i == 9:
                cell.font = font_bold
                cell.fill = fill_emerald_chip
                cell.alignment = Alignment(horizontal="center", vertical="center")

        current_r += 1

    # 4. Раздел 2: Готовые Креативы (Посты 1, 2, 3)
    current_r += 1
    ws.merge_cells(f"A{current_r}:J{current_r}")
    ws.cell(current_r, 1, "📝 РАЗДЕЛ 2: ГОТОВЫЕ РЕКЛАМНЫЕ ПОСТЫ ДЛЯ ЗАКАЗА В TELEGRAM (FAST COPY-PASTE)")
    ws.cell(current_r, 1).font = font_sec
    ws.cell(current_r, 1).fill = fill_purple_chip
    ws.row_dimensions[current_r].height = 24
    current_r += 1

    posts = [
        (
            "Пост 1 (Комбо: Продажи + CRM)",
            "Для каналов: @crm50, @completo_ru, @Salesnotes, @crm_secrets, @introvert_crm",
            "Почему ваша CRM показывает прибыль, которой на самом деле нет? 📉\n\n"
            "Вы внедрили amoCRM или Битрикс24, настроили роботов, но менеджеры всё равно забивают на стандарты:\n"
            "• Отправляют КП на «кладбище почты» без фиксации даты следующего шага.\n"
            "• Сдаются при первом же возражении «Дорого» и сливают чек.\n"
            "• Пишут в комментариях абстрактное «клиент думает», маскируя уходящий лид.\n\n"
            "РОП физически успевает отслушать лишь 2–3% звонков, а остальные 98% воронки остаются в слепой зоне.\n\n"
            "Платформа RevOps OS PRO объединяет речевую аналитику ИИ и вашу CRM, полностью ликвидируя сливы:\n"
            "🟢 ИИ слушает 100% звонков: Нейросеть Whisper Pro расшифровывает каждый разговор и оценивает его по 13 жестким бизнес-критериям.\n"
            "🟢 Мгновенные алерты в Telegram: Если менеджер грубит, называет неверную цену или отпускает клиента — РОП получает уведомление со сценарием спасения сделки за 30 секунд.\n"
            "🟢 Автозаполнение карточек: ИИ сам пишет краткое резюме (Executive Summary) и выставляет договоренности прямо в CRM. Менеджеры больше не тратят время на рутину.\n\n"
            "Результат: Внедрение за 15 минут по API без остановки отдела. Спасение даже одной крупной сделки полностью окупает годовую подписку.\n\n"
            "Перестаньте гадать, сколько денег пролетает мимо кассы. Загрузите 3 вчерашних звонка ваших менеджеров — ИИ бесплатно оцифрует сумму денежного риска в рублях за 20 минут.\n\n"
            "👉 [Получить экспресс-аудит 3 звонков (0 ₽)] {ССЫЛКА_С_UTM}"
        ),
        (
            "Пост 2 (Кейс-разбор в стиле шоу Гребенюка)",
            "Для каналов: @resulting_ru, @biz_razbory, @floor_99, @batyrev_maxim, @grafsky_b2b",
            "«Ну ладно... Я просто скину вам КП на почту». Как менеджер своими руками похоронил сделку на 340 000 ₽. Пошаговый разбор 🔥\n\n"
            "Знакомая фраза? Руководитель думает, что отдел продаж работает, а ИИ-супервайзер RevOps OS видит реальность:\n"
            "1. Менеджер 40 минут рассказывал о продукте секретарю, так и не узнав, кто принимает финансовое решение (ЛПР).\n"
            "2. Услышав дежурное «Отправьте предложение, мы посмотрим», покорно согласился и положил трубку.\n"
            "3. Не назначил точную дату и время звонка. Сделка ушла на кладбище забытых лидов.\n\n"
            "Убыток компании: 340 000 ₽ чистой маржи + 6 500 ₽ сгоревшего бюджета Яндекс.Директа.\n\n"
            "РОП этого звонка не слышал, потому что физически успевает разобрать не более 5 диалогов в неделю.\n\n"
            "Платформа RevOps OS слушает 100% звонков менеджеров за секунды:\n"
            "• Находит фатальные ошибки по 13 критериям B2B.\n"
            "• Присылает РОПу алерт в Telegram за 30 секунд со сценарием: что сказать клиенту, чтобы вернуть его в сделку.\n"
            "• Показывает сумму упущенной прибыли прямо в рублях.\n\n"
            "Хотите узнать, где прямо сейчас сливают ваши менеджеры? Отправьте 3 вчерашних звонка — ИИ разберет их бесплатно за 20 минут.\n\n"
            "👉 [Проверить 3 звонка отдела продаж (0 ₽)] {ССЫЛКА_С_UTM}"
        ),
        (
            "Пост 3 (Переманивание клиентов с imot.io и Mango)",
            "Для каналов: @rocketsales_blog, @introvert_crm, @ai_business_ru, @crm50",
            "Платите за минуты речевой аналитики? Хватит считать секунды разговоров ваших менеджеров! ⏱️\n\n"
            "Классические системы аналитики продают «пакеты минут»: 50 000 ₽ за 10 000 минут. Если в отделе 8–10 менеджеров, лимит сгорает за 2 недели, а дальше вам выставляют счета на 120к–190к рублей.\n\n"
            "При этом вы получаете гору сложных графиков и облако тегов, в которых РОП всё равно не понимает, кому звонить.\n\n"
            "Переходите на RevOps OS Pro:\n"
            "✅ Фиксированная цена за команду без подсчета минут: контролируйте 100% разговоров без страха переплаты.\n"
            "✅ Деньги вместо графиков: ИИ показывает не «длительность пауз», а сумму в рублях, которую менеджер поставил под угрозу срыва.\n"
            "✅ Перенос настроек за 1 рабочий день: перенесем ваши словари и регламенты без стресса для IT-отдела.\n\n"
            "🎁 Спец-оффер при переходе: 1 месяц аналитики в подарок при переезде с поминутных тарифов.\n\n"
            "👉 [Узнать условия переезда и получить тест 3 звонков] {ССЫЛКА_С_UTM}"
        )
    ]

    for p_title, p_channels, p_text in posts:
        ws.row_dimensions[current_r].height = 22
        ws.merge_cells(f"A{current_r}:J{current_r}")
        ws.cell(current_r, 1, f"📌 {p_title} | {p_channels}").font = font_bold
        ws.cell(current_r, 1).fill = fill_blue_chip
        ws.cell(current_r, 1).border = thin_border
        current_r += 1

        ws.row_dimensions[current_r].height = 110
        ws.merge_cells(f"A{current_r}:J{current_r}")
        cell_body = ws.cell(current_r, 1, p_text)
        cell_body.font = font_code
        cell_body.fill = fill_white
        cell_body.border = thin_border
        cell_body.alignment = Alignment(vertical="center", wrap_text=True)
        current_r += 1

    # 5. Раздел 3: Инструкция по закупке без слива бюджета
    current_r += 1
    ws.merge_cells(f"A{current_r}:J{current_r}")
    ws.cell(current_r, 1, "🛡️ РАЗДЕЛ 3: ПРАВИЛА ЗАКУПА В TELEGRAM (КАК СБИТЬ ЦЕНУ И НЕ СЛИТЬ БЮДЖЕТ)")
    ws.cell(current_r, 1).font = font_sec
    ws.cell(current_r, 1).fill = fill_emerald_chip
    ws.row_dimensions[current_r].height = 24
    current_r += 1

    rules = [
        ("Правило 1: Проверка через TGStat", "Перед оплатой обязательно вбейте канал на tgstat.ru. Смотрите на ER (вовлеченность должна быть >15-20%) и график суточного охвата: если посты читают равномерно, а не скачком в 03:00 ночи — накрутки ботами нет."),
        ("Правило 2: Формат размещения 2/48", "Не берите 1/24 (быстро смоется лентой). Для длинных экспертных постов B2B просите формат 2/48 (2 часа в топе без перекрытия другими постами и 48 часов в ленте), либо формат 'без удаления'."),
        ("Правило 3: Скрипт торга со скидкой 10-20%", "Пишите админу прямо: «Тестируем новый ИИ-сервис для B2B отделов продаж. Если отдача будет хорошей — зайдем на постоянный пул повторных размещений раз в месяц. Какую скидку сможете дать на первый тестовый запуск?». В 80% случаев скидывают 15-20%."),
        ("Правило 4: Единый UTM-стандарт", "Для каждого канала строго своя ссылка: https://ai-rop.ru/?utm_source=tg&utm_medium={код_канала}&utm_campaign=audit3. В Яндекс.Метрике (ID: 113329268) сразу отслеживаем стоимость лида (CPL) и конверсию.")
    ]

    for r_title, r_desc in rules:
        ws.row_dimensions[current_r].height = 28
        ws.merge_cells(f"A{current_r}:C{current_r}")
        c1 = ws.cell(current_r, 1, r_title)
        c1.font = font_bold
        c1.fill = fill_zebra
        c1.border = thin_border
        c1.alignment = Alignment(vertical="center")

        ws.merge_cells(f"D{current_r}:J{current_r}")
        c2 = ws.cell(current_r, 4, r_desc)
        c2.font = font_reg
        c2.fill = fill_white
        c2.border = thin_border
        c2.alignment = Alignment(vertical="center", wrap_text=True)

        current_r += 1

    col_widths = {
        "A": 8,    # Код
        "B": 18,   # Категория
        "C": 30,   # Канал
        "D": 34,   # ЦА
        "E": 12,   # Формат
        "F": 16,   # Цена
        "G": 22,   # Креатив
        "H": 36,   # UTM
        "I": 18,   # Статус
        "J": 18    # Факт
    }
    for col_l, w in col_widths.items():
        ws.column_dimensions[col_l].width = w

    wb.save(file_path)
    print(f"Sheet '{sheet_name}' successfully added to: {file_path}")


if __name__ == "__main__":
    targets = [
        Path(r"C:\Users\strel\Desktop\RevOps_Founder_Workspace_Planner.xlsx"),
        Path(r"C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\docs\RevOps_Founder_Workspace_Planner.xlsx"),
        Path(r"C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\presentation\RevOps_Founder_Workspace_Planner.xlsx")
    ]
    for t in targets:
        add_channels_traffic_sheet(t)
