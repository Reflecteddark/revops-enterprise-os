"""
Генератор рабочей книги Фаундера: RevOps_Founder_Workspace_Planner.xlsx
Создает специализированный инструмент для заметок по ходу работы, планирования спринтов,
ведения базы знаний и трекинга партнеров/клиентов.
"""

from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def generate_founder_workspace(output_path: Path):
    wb = openpyxl.Workbook()

    # Палитра цветов для рабочего инструмента
    c_navy_dark = "0F172A"
    c_navy_light = "1E293B"
    c_indigo = "4F46E5"
    c_indigo_light = "EEF2FF"
    c_emerald = "10B981"
    c_emerald_light = "ECFDF5"
    c_amber_light = "FFFBEB"
    c_red_light = "FEF2F2"
    c_card_bg = "F8FAFC"
    c_border = "CBD5E1"

    fill_navy = PatternFill(start_color=c_navy_dark, end_color=c_navy_dark, fill_type="solid")
    fill_header = PatternFill(start_color=c_navy_light, end_color=c_navy_light, fill_type="solid")
    fill_card = PatternFill(start_color=c_card_bg, end_color=c_card_bg, fill_type="solid")
    fill_done = PatternFill(start_color=c_emerald_light, end_color=c_emerald_light, fill_type="solid")
    fill_in_progress = PatternFill(start_color=c_amber_light, end_color=c_amber_light, fill_type="solid")
    fill_backlog = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    fill_highlight = PatternFill(start_color=c_indigo_light, end_color=c_indigo_light, fill_type="solid")

    font_title = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
    font_sub = Font(name="Segoe UI", size=10, italic=True, color="94A3B8")
    font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    font_section = Font(name="Segoe UI", size=12, bold=True, color="1E293B")
    font_bold = Font(name="Segoe UI", size=10, bold=True, color="1E293B")
    font_reg = Font(name="Segoe UI", size=10, color="334155")
    font_p0 = Font(name="Segoe UI", size=10, bold=True, color="DC2626")
    font_p1 = Font(name="Segoe UI", size=10, bold=True, color="D97706")
    font_p2 = Font(name="Segoe UI", size=10, color="64748B")

    thin_border = Border(
        left=Side(style="thin", color=c_border),
        right=Side(style="thin", color=c_border),
        top=Side(style="thin", color=c_border),
        bottom=Side(style="thin", color=c_border)
    )

    # =========================================================================
    # Лист 1: 🎯 Спринты_и_Задачи
    # =========================================================================
    ws1 = wb.active
    ws1.title = "🎯 Спринты_и_Задачи"
    ws1.views.sheetView[0].showGridLines = True

    ws1.merge_cells("A1:H1")
    ws1["A1"] = "RevOps Enterprise OS — Мастер-Планнер Спринтов и Задач Фаундера"
    ws1["A1"].font = font_title
    ws1["A1"].fill = fill_navy
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 35

    task_headers = ["ID", "Стрим", "Задача / Инициатива", "Приоритет", "Дедлайн", "Статус", "Артефакт / Файл", "Заметки и результат"]
    for c_idx, h in enumerate(task_headers, start=1):
        cell = ws1.cell(3, c_idx, h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[3].height = 26

    tasks_data = [
        # Выполненные задачи
        ("T-01", "CRM Connectors", "Харденинг коннекторов amoCRM и Битрикс24 (защита от 429/503, батчинг, резолв дублей)", "P0", "01.10.2026", "✅ ВЫПОЛНЕНО", "amocrm_connector.py, bitrix24_connector.py", "105 тестов проходят (100% PASS), лимиты 6 и 2 rps соблюдены", fill_done),
        ("T-02", "Legal 152-ФЗ", "Сборка юридического комплекта безопасности для клиентов (5 шаблонов по 152-ФЗ)", "P0", "01.10.2026", "✅ ВЫПОЛНЕНО", "docs/152-ФЗ_Комплект_Безопасности.docx", "Включает согласие менеджера, приказ, регламент 30 дней", fill_done),
        ("T-03", "Partnerships", "Оффер для интеграторов CRM с 25-30% ревшаром (One-Pager DOCX + скрипты TG/Email)", "P0", "01.10.2026", "✅ ВЫПОЛНЕНО", "docs/Партнерское_Предложение_Интеграторам_CRM.docx", "Выплаты 12 250 - 36 000 ₽/мес на клиента, скрипты готовы", fill_done),
        ("T-04", "Speech Engine", "Модуль распознавания речи speech_engine.py по 152-ФЗ с токенизацией ПДн", "P0", "01.10.2026", "✅ ВЫПОЛНЕНО", "speech_engine.py, test_speech_engine.py", "Yandex SpeechKit Deferred (0.02 ₽/мин) + Groq fallback + Excel sync", fill_done),
        ("T-05", "Lead Magnet", "Экспресс-аудит 3 звонков (run_express_audit.py) для холодного/теплого выхода", "P0", "01.10.2026", "✅ ВЫПОЛНЕНО", "run_express_audit.py, карточка HTML на Desktop", "Выдает 1-page отчет + готовый текст в Telegram с расчетом потерь", fill_done),
        ("T-06", "Web / Landing", "Интерактивный B2B-лендинг с калиброванным калькулятором ROI", "P0", "01.10.2026", "✅ ВЫПОЛНЕНО", "presentation/landing.html", "18% SQL, 18% слив, возврат 35%, 1-клик запуск на Desktop", fill_done),
        ("T-07", "Product Demo", "Сборка 5-вкладочного Enterprise Showcase Excel (.xlsx) без скрытых формул", "P1", "01.10.2026", "✅ ВЫПОЛНЕНО", "RevOps_Platform_Demo_Sample.xlsx", "Вшит в сайт через Base64 + лежит на Рабочем столе", fill_done),
        # Текущие задачи
        ("T-08", "GTM / Outreach", "Первая рассылка оффера 10 интеграторам amoCRM и Битрикс24", "P0", "02.10.2026", "⏳ В РАБОТЕ", "Вкладка 'Пайплайн_Интеграторов'", "Цель: 2 партнерских созвона на этой неделе", fill_in_progress),
        ("T-09", "Deployment", "Развертывание веб-сайта на публичном домене / хостинге", "P1", "02.10.2026", "⏳ В РАБОТЕ", "example.qwen.site / Vercel", "Код готов, проверить мобильную верстку и форму", fill_in_progress),
        # Бэклог
        ("T-10", "Pilot Delivery", "Проведение первого платного пилотного спринта (29 000 ₽)", "P0", "07.10.2026", "📋 БЭКЛОГ", "Пайплайн клиентов", "Оцифровать 100% звонков первого клиента, выдать список спасения", fill_backlog),
        ("T-11", "Telegram Bot", "Подключение вебхука живой CRM к telegram_bot.py для realtime-алертов", "P1", "10.10.2026", "📋 БЭКЛОГ", "telegram_bot.py", "Утренний фокус в 08:30 + SOS-алерт при сливе Next Step", fill_backlog),
        ("T-12", "Legal NPD", "Подготовка формы договора-оферты и акта для Самозанятого через 'Мой Налог'", "P2", "12.10.2026", "📋 БЭКЛОГ", "docs/Договор_Оферта_Самозанятый.docx", "Статья 15 422-ФЗ, 0% НДС, безналичный расчет с юрлицами", fill_backlog),
    ]

    for r_idx, row in enumerate(tasks_data, start=4):
        ws1.row_dimensions[r_idx].height = 24
        for c_idx in range(1, 9):
            cell = ws1.cell(r_idx, c_idx, row[c_idx-1])
            cell.font = font_reg
            cell.border = thin_border
            if c_idx in [1, 2, 5]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx == 4:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = font_p0 if row[3] == "P0" else (font_p1 if row[3] == "P1" else font_p2)
            elif c_idx == 6:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.fill = row[8]
                cell.font = font_bold

    # =========================================================================
    # Лист 2: 💡 База_Знаний_и_УТП (Стратегические заметки)
    # =========================================================================
    ws2 = wb.create_sheet("💡 База_Знаний_и_УТП")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:G1")
    ws2["A1"] = "База Знаний Продукта: Позиционирование, УТП и 3 Ключевые Боли Клиентов"
    ws2["A1"].font = font_title
    ws2["A1"].fill = fill_navy
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 35

    # Карточка 1: Главное УТП (из A1 таблицы пользователя)
    ws2.merge_cells("A3:G3")
    ws2["A3"] = "🔥 ГЛАВНОЕ УНИКАЛЬНОЕ ТОРГОВОЕ ПРЕДЛОЖЕНИЕ (УТП)"
    ws2["A3"].font = font_section
    ws2["A3"].fill = fill_highlight

    ws2.merge_cells("A4:G6")
    utp_text = (
        '«Мы не просто слушаем ваши звонки. Мы находим, почему вы теряете деньги, и пишем вам пошаговый план, как это исправить».\n\n'
        'Это УТП четко отделяет продукт от конкурентов. Мы не продаем софт — мы продаем результат: повышение эффективности отдела продаж.'
    )
    ws2["A4"] = utp_text
    ws2["A4"].font = Font(name="Segoe UI", size=11, bold=True, color="1E3A8A")
    ws2["A4"].alignment = Alignment(vertical="center", wrap_text=True)
    ws2["A4"].border = thin_border

    # Карточка 2: 3 Боли целевой аудитории
    ws2.merge_cells("A8:G8")
    ws2["A8"] = "🎯 3 КЛЮЧЕВЫЕ БОЛИ ЦЕЛЕВОЙ АУДИТОРИИ (Коммерческие директора, РОПы, Собственники)"
    ws2["A8"].font = font_section
    ws2["A8"].fill = fill_highlight

    pains = [
        ("Боль №1", "«Я не знаю, почему мы теряем сделки»", "Руководитель видит, что план по выручке не выполняется, но причина скрыта в сотнях звонков (качество лидов? возражения? слив Next Step?). Наш продукт дает объективный математический ответ с точностью до рубля.", "Отчет CFO + Экспресс-аудит 3 звонков"),
        ("Боль №2", "«Мои менеджеры не растут»", "Руководитель тратит время на тренинги, но менеджеры продолжают отпускать клиентов фразой 'подумайте'. Наш продукт дает конкретные данные по 13 критериям для адресного коучинга каждого менеджера.", "Пульт РОПа + Рейтинг менеджеров"),
        ("Боль №3", "«Мне нужно быстро получить результат»", "Классический консалтинг занимает 2-3 месяца и стоит сотни тысяч. Наш продукт оцифровывает весь отдел за 7 дней пилотного спринта (29 000 ₽) с гарантией окупаемости.", "7-дневный пилотный спринт"),
    ]

    pain_headers = ["Боль", "Формулировка клиента", "Суть проблемы в бизнесе", "Наше решение в RevOps OS"]
    ws2.cell(9, 1, pain_headers[0]).font = font_header
    ws2.cell(9, 1).fill = fill_header
    ws2.cell(9, 2, pain_headers[1]).font = font_header
    ws2.cell(9, 2).fill = fill_header
    ws2.merge_cells("C9:E9")
    ws2["C9"] = pain_headers[2]
    ws2["C9"].font = font_header
    ws2["C9"].fill = fill_header
    ws2.merge_cells("F9:G9")
    ws2["F9"] = pain_headers[3]
    ws2["F9"].font = font_header
    ws2["F9"].fill = fill_header
    ws2.row_dimensions[9].height = 24

    for idx, (p_num, p_quote, p_desc, p_sol) in enumerate(pains, start=10):
        ws2.row_dimensions[idx].height = 36
        ws2.cell(idx, 1, p_num).font = font_bold
        ws2.cell(idx, 1).alignment = Alignment(horizontal="center", vertical="center")
        ws2.cell(idx, 1).border = thin_border

        ws2.cell(idx, 2, p_quote).font = Font(name="Segoe UI", size=10, italic=True, bold=True, color="DC2626")
        ws2.cell(idx, 2).alignment = Alignment(vertical="center", wrap_text=True)
        ws2.cell(idx, 2).border = thin_border

        ws2.merge_cells(f"C{idx}:E{idx}")
        ws2[f"C{idx}"] = p_desc
        ws2[f"C{idx}"].font = font_reg
        ws2[f"C{idx}"].alignment = Alignment(vertical="center", wrap_text=True)
        ws2[f"C{idx}"].border = thin_border

        ws2.merge_cells(f"F{idx}:G{idx}")
        ws2[f"F{idx}"] = p_sol
        ws2[f"F{idx}"].font = font_bold
        ws2[f"F{idx}"].fill = fill_done
        ws2[f"F{idx}"].alignment = Alignment(horizontal="center", vertical="center")
        ws2[f"F{idx}"].border = thin_border

    # Карточка 3: Позиционирование на рынке
    ws2.merge_cells("A14:G14")
    ws2["A14"] = "⚖️ МАТРИЦА ПОЗИЦИОНИРОВАНИЯ (ПОЧЕМУ МЫ ВЫИГРЫВАЕМ У КОНКУРЕНТОВ)"
    ws2["A14"].font = font_section
    ws2["A14"].fill = fill_highlight

    matrix_rows = [
        ("Классический консалтинг", "Слишком дорого (от 300 000 ₽) и долго (2-3 месяца). Дают PDF-отчет и уходят, внедрение виснет.", "RevOps OS: запуск за 7 дней, подписка от 49 000 ₽/мес, автоматический надзор 24/7."),
        ("Обычная речевая аналитика (SaaS)", "Слишком технично. Дают облако тегов и транскрипты, но не говорят РОПу, что делать и сколько денег теряется.", "RevOps OS: готовые алерты РОПу в Telegram и расчет потерь в рублях."),
        ("Зарубежный Revenue Intelligence (Gong)", "Заблокирован в РФ, не поддерживает amoCRM/Битрикс24, нарушает 152-ФЗ, стоит от $120/пользователь.", "RevOps OS: 100% серверы в РФ (152-ФЗ), нативная интеграция с amoCRM и Битрикс24."),
    ]

    for m_idx, (m_type, m_flaw, m_adv) in enumerate(matrix_rows, start=15):
        ws2.row_dimensions[m_idx].height = 30
        ws2.cell(m_idx, 1, m_type).font = font_bold
        ws2.cell(m_idx, 1).border = thin_border

        ws2.merge_cells(f"B{m_idx}:D{m_idx}")
        ws2[f"B{m_idx}"] = m_flaw
        ws2[f"B{m_idx}"].font = font_reg
        ws2[f"B{m_idx}"].alignment = Alignment(vertical="center", wrap_text=True)
        ws2[f"B{m_idx}"].border = thin_border

        ws2.merge_cells(f"E{m_idx}:G{m_idx}")
        ws2[f"E{m_idx}"] = m_adv
        ws2[f"E{m_idx}"].font = Font(name="Segoe UI", size=9, bold=True, color="16A34A")
        ws2[f"E{m_idx}"].fill = fill_done
        ws2[f"E{m_idx}"].alignment = Alignment(vertical="center", wrap_text=True)
        ws2[f"E{m_idx}"].border = thin_border

    # =========================================================================
    # Лист 3: 🤝 Пайплайн_Интеграторов
    # =========================================================================
    ws3 = wb.create_sheet("🤝 Пайплайн_Интеграторов")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:I1")
    ws3["A1"] = "Воронка Работы с Партнерами-Интеграторами CRM (Оффер 25–30% Recurring RevShare)"
    ws3["A1"].font = font_title
    ws3["A1"].fill = fill_navy
    ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 35

    partner_headers = ["№", "Компания-Интегратор", "CRM Стек", "Город / Сайт", "Контакт (TG / Email)", "Статус", "Условия RevShare", "Потенциал (клиентов)", "Следующий шаг"]
    for c_idx, h in enumerate(partner_headers, start=1):
        cell = ws3.cell(3, c_idx, h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[3].height = 26

    # 10 целевых партнеров (шаблон для заполнения)
    partners_sample = [
        (1, "Интегратор 1 (Топ amoCRM)", "amoCRM", "Москва", "@partner_ceo", "В плане на контакт", "25% (12 250 - 30 000 ₽/мес)", "15-20 активных", "Отправить оффер в Telegram"),
        (2, "Интегратор 2 (Золотой Битрикс24)", "Битрикс24", "СПб", "@bitrix_lead", "В плане на контакт", "30% (от 3 клиентов)", "25-30 активных", "Отправить One-Pager на почту"),
        (3, "Интегратор 3 (CRM + Телефония)", "amoCRM / Б24", "Екатеринбург", "ceo@crm-ural.ru", "В плане на контакт", "25%", "10 активных", "Созвон в Zoom на 15 минут"),
    ]

    for p_idx, p_data in enumerate(partners_sample, start=4):
        ws3.row_dimensions[p_idx].height = 24
        for c_i in range(1, 10):
            cell = ws3.cell(p_idx, c_i, p_data[c_i-1])
            cell.font = font_reg
            cell.border = thin_border
            if c_i in [1, 3]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_i == 6:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.fill = fill_in_progress
                cell.font = font_bold

    # Пустые строки для ведения базы (до 20 строк)
    for empty_idx in range(7, 21):
        ws3.row_dimensions[empty_idx].height = 22
        ws3.cell(empty_idx, 1, empty_idx - 3).alignment = Alignment(horizontal="center")
        for col_i in range(1, 10):
            ws3.cell(empty_idx, col_i).border = thin_border

    # =========================================================================
    # Лист 4: 🎙️ Клиентские_Пилоты
    # =========================================================================
    ws4 = wb.create_sheet("🎙️ Клиентские_Пилоты")
    ws4.views.sheetView[0].showGridLines = True

    ws4.merge_cells("A1:J1")
    ws4["A1"] = "Воронка Клиентов: Экспресс-Аудит (0 ₽) → 7-дневный Пилот (29 000 ₽) → Подписка"
    ws4["A1"].font = font_title
    ws4["A1"].fill = fill_navy
    ws4["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[1].height = 35

    client_headers = ["№", "Компания / Клиент", "Сфера бизнеса", "Менеджеров ОП", "Средний чек", "CRM", "Статус аудита", "Пилот (29к)", "Тариф подписки", "Заметки"]
    for c_idx, h in enumerate(client_headers, start=1):
        cell = ws4.cell(3, c_idx, h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[3].height = 26

    # Пустые строки для клиентов
    for c_row in range(4, 18):
        ws4.row_dimensions[c_row].height = 24
        ws4.cell(c_row, 1, c_row - 3).alignment = Alignment(horizontal="center")
        for col_j in range(1, 11):
            ws4.cell(c_row, col_j).border = thin_border

    # =========================================================================
    # Лист 5: 💬 Быстрые_Скрипты_и_Шаблоны
    # =========================================================================
    ws5 = wb.create_sheet("💬 Быстрые_Скрипты")
    ws5.views.sheetView[0].showGridLines = True

    ws5.merge_cells("A1:F1")
    ws5["A1"] = "Готовые Скрипты и Шаблоны Сообщений (Копировать / Вставить)"
    ws5["A1"].font = font_title
    ws5["A1"].fill = fill_navy
    ws5["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws5.row_dimensions[1].height = 35

    scripts = [
        ("Шаблон 1: Telegram партнеру-интегратору CRM", 
         "Приветствую! Вижу, вы плотно занимаетесь внедрением amoCRM/Битрикс24.\n\nМы разработали надстройку RevOps AI Supervisor: она решает главную боль клиентов после внедрения CRM — 'менеджеры не ведут сделки и не фиксируют Next Step'. Система слушает 100% звонков через Whisper, находит утечки и присылает РОПу готовый список спасения в Telegram.\n\nПредлагаем партнерство: 25–30% пожизненного ревшара с каждого чека (от 12 250 до 36 000 ₽/мес на клиента). Всю техподдержку и аналитику ведем мы.\n\nУдобно глянуть короткий One-Pager в Word на 1 страницу?"),
        
        ("Шаблон 2: Telegram собственнику после экспресс-аудита 3 звонков",
         "Приветствую! Прогнали 3 ваших тестовых звонка через ИИ-супервизор RevOps.\n\nРезультат экспресс-диагностики:\n• Средний балл качества: 7.3 из 13 (норматив B2B: ≥ 10.5)\n• Главная утечка: в 2 из 3 звонков менеджеры отпустили клиента без даты следующего контакта ('ну спишемся', 'подумайте').\n• Оценка потерь: при 150 лидах и чеке 400к вы теряете ~1.9М ₽ каждый месяц на брошенных сделках.\n\nИнтерактивную карточку аудита прикрепил файлом. Предлагаю запустить 7-дневный пилот за 29 000 ₽ — оцифруем 100% звонков и вернем от 400 000 ₽ зависших сделок уже на этой неделе.\n\nУдобно созвониться завтра в 11:30 на 10 минут?"),
         
        ("Шаблон 3: Ответ юристу / службе безопасности по 152-ФЗ",
         "Безопасность полностью закрыта по 152-ФЗ РФ:\n1. Все серверы и базы данных находятся в Москве (Yandex Cloud / Selectel, сертификаты ФСТЭК УЗ 1-2).\n2. До отправки в нейросеть работает криптографический модуль PIISanitizer: телефоны и email маскируются в [PHONE_TOKEN_X], персональные данные наружу не выходят.\n3. Аудиозаписи автоматически удаляются через 30 дней.\n4. Передаем полный юридический комплект: согласие сотрудника (ст. 86 ТК РФ), приказ о контроле качества и текст дисклеймера для АТС.\n5. Оплата по безналичному расчету (Самозанятый, ст. 15 422-ФЗ, 0% НДС, чек 'Мой Налог')."),
    ]

    curr_row = 3
    for s_title, s_body in scripts:
        ws5.merge_cells(f"A{curr_row}:F{curr_row}")
        ws5[f"A{curr_row}"] = s_title
        ws5[f"A{curr_row}"].font = font_section
        ws5[f"A{curr_row}"].fill = fill_highlight
        ws5.row_dimensions[curr_row].height = 24

        curr_row += 1
        ws5.merge_cells(f"A{curr_row}:F{curr_row+4}")
        ws5[f"A{curr_row}"] = s_body
        ws5[f"A{curr_row}"].font = font_reg
        ws5[f"A{curr_row}"].border = thin_border
        ws5[f"A{curr_row}"].alignment = Alignment(vertical="top", wrap_text=True)
        for r_offset in range(5):
            ws5.row_dimensions[curr_row + r_offset].height = 20

        curr_row += 6

    # Автоподбор ширин колонок
    for sheet in [ws1, ws2, ws3, ws4, ws5]:
        for col in sheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            sheet.column_dimensions[col_letter].width = max(12, min(max_len + 3, 50))

    ws1.column_dimensions["A"].width = 10
    ws1.column_dimensions["B"].width = 18
    ws1.column_dimensions["C"].width = 45
    ws1.column_dimensions["D"].width = 14
    ws1.column_dimensions["E"].width = 14
    ws1.column_dimensions["F"].width = 18
    ws1.column_dimensions["G"].width = 32
    ws1.column_dimensions["H"].width = 50

    ws2.column_dimensions["A"].width = 14
    ws2.column_dimensions["B"].width = 36
    ws2.column_dimensions["C"].width = 25
    ws2.column_dimensions["D"].width = 25
    ws2.column_dimensions["E"].width = 25
    ws2.column_dimensions["F"].width = 25
    ws2.column_dimensions["G"].width = 25

    ws3.column_dimensions["B"].width = 30
    ws3.column_dimensions["C"].width = 16
    ws3.column_dimensions["D"].width = 20
    ws3.column_dimensions["E"].width = 24
    ws3.column_dimensions["F"].width = 22
    ws3.column_dimensions["G"].width = 28
    ws3.column_dimensions["H"].width = 22
    ws3.column_dimensions["I"].width = 32

    ws5.column_dimensions["A"].width = 80

    wb.save(output_path)
    print(f"Founder Workspace Planner successfully created: {output_path}")


if __name__ == "__main__":
    target = Path("presentation/RevOps_Founder_Workspace_Planner.xlsx")
    generate_founder_workspace(target)
    import shutil
    shutil.copyfile(target, "docs/RevOps_Founder_Workspace_Planner.xlsx")
    shutil.copyfile(target, "C:/Users/strel/Desktop/RevOps_Founder_Workspace_Planner.xlsx")
    print("Copied to Desktop and docs!")
