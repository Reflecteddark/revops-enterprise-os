"""
RevOps Enterprise OS V17.6 — Генератор демонстрационного Excel-файла представительского класса.
Создает полноценный многовкладочный интерактивный документ (Showcase Edition),
который вызывает 'WOW'-эффект у Собственников, CFO и РОПов.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path


def create_enterprise_demo_excel(output_path: Path):
    wb = openpyxl.Workbook()

    # Стилизация палитры
    c_navy_dark = "0F172A"
    c_navy_light = "1E293B"
    c_indigo = "4F46E5"
    c_indigo_light = "EEF2FF"
    c_emerald = "10B981"
    c_emerald_light = "ECFDF5"
    c_red = "EF4444"
    c_red_light = "FEF2F2"
    c_amber = "F59E0B"
    c_amber_light = "FFFBEB"
    c_card_bg = "F8FAFC"
    c_border = "CBD5E1"

    fill_navy = PatternFill(start_color=c_navy_dark, end_color=c_navy_dark, fill_type="solid")
    fill_header = PatternFill(start_color=c_navy_light, end_color=c_navy_light, fill_type="solid")
    fill_indigo = PatternFill(start_color=c_indigo, end_color=c_indigo, fill_type="solid")
    fill_indigo_light = PatternFill(start_color=c_indigo_light, end_color=c_indigo_light, fill_type="solid")
    fill_input = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")  # Yellow highlight for inputs
    fill_card = PatternFill(start_color=c_card_bg, end_color=c_card_bg, fill_type="solid")
    fill_success = PatternFill(start_color=c_emerald_light, end_color=c_emerald_light, fill_type="solid")
    fill_alert = PatternFill(start_color=c_red_light, end_color=c_red_light, fill_type="solid")
    fill_warn = PatternFill(start_color=c_amber_light, end_color=c_amber_light, fill_type="solid")

    font_title = Font(name="Segoe UI", size=15, bold=True, color="FFFFFF")
    font_sub = Font(name="Segoe UI", size=10, italic=True, color="94A3B8")
    font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    font_section = Font(name="Segoe UI", size=12, bold=True, color="1E293B")
    font_bold = Font(name="Segoe UI", size=10, bold=True, color="1E293B")
    font_reg = Font(name="Segoe UI", size=10, color="334155")
    font_kpi_val = Font(name="Segoe UI", size=18, bold=True, color="1E293B")
    font_kpi_red = Font(name="Segoe UI", size=18, bold=True, color="DC2626")
    font_kpi_green = Font(name="Segoe UI", size=18, bold=True, color="16A34A")

    thin_border = Border(
        left=Side(style="thin", color=c_border),
        right=Side(style="thin", color=c_border),
        top=Side(style="thin", color=c_border),
        bottom=Side(style="thin", color=c_border)
    )
    thick_bottom = Border(
        left=Side(style="thin", color=c_border),
        right=Side(style="thin", color=c_border),
        top=Side(style="thin", color=c_border),
        bottom=Side(style="medium", color=c_indigo)
    )

    # =========================================================================
    # Вкладка 1: ⚡ Экспресс_Калькулятор_ROI (Интерактивная финансовая модель)
    # =========================================================================
    ws1 = wb.active
    ws1.title = "⚡ Экспресс_Калькулятор_ROI"
    ws1.views.sheetView[0].showGridLines = True

    ws1.merge_cells("B2:J2")
    ws1["B2"] = "RevOps Enterprise OS V17.6 — Интерактивный Калькулятор Окупаемости и Утечек"
    ws1["B2"].font = font_title
    ws1["B2"].fill = fill_navy
    ws1["B2"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[2].height = 38

    ws1.merge_cells("B3:J3")
    ws1["B3"] = "Измените желтые ячейки ниже, чтобы рассчитать потери вашего отдела продаж и возврат выручки"
    ws1["B3"].font = font_sub
    ws1["B3"].fill = fill_navy
    ws1["B3"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[3].height = 20

    # Заголовок секции ввода
    ws1["B5"] = "ПАРАМЕТРЫ ВАШЕГО БИЗНЕСА (ВВОД)"
    ws1["B5"].font = font_section
    ws1["B5"].alignment = Alignment(vertical="center")

    inputs = [
        ("B6", "C6", "Количество менеджеров по продажам (ОП)", 5, "чел", "0"),
        ("B7", "C7", "Входящий поток лидов в месяц", 150, "лидов", "#,##0"),
        ("B8", "C8", "Средний чек закрытой сделки", 400000, "₽", "#,##0 ₽"),
        ("B9", "C9", "Доля целевых квалифицированных лидов (SQL / B2B норма)", 0.18, "%", "0.0%"),
    ]

    for label_cell, val_cell, label, val, unit, num_fmt in inputs:
        ws1[label_cell] = label
        ws1[label_cell].font = font_bold
        ws1[label_cell].border = thin_border
        
        ws1[val_cell] = val
        ws1[val_cell].font = Font(name="Segoe UI", size=11, bold=True, color="1E3A8A")
        ws1[val_cell].fill = fill_input
        ws1[val_cell].alignment = Alignment(horizontal="right", vertical="center")
        ws1[val_cell].border = thin_border
        ws1[val_cell].number_format = num_fmt

        unit_cell = "D" + val_cell[1:]
        ws1[unit_cell] = unit
        ws1[unit_cell].font = font_sub
        ws1[unit_cell].alignment = Alignment(horizontal="left", vertical="center")

    ws1.row_dimensions[6].height = 24
    ws1.row_dimensions[7].height = 24
    ws1.row_dimensions[8].height = 24
    ws1.row_dimensions[9].height = 24

    # KPI Сводка (Справа)
    kpis = [
        ("F5", "G6", "Квалифицированный пайплайн", "=C7*C9*C8", font_kpi_val, fill_card, "#,##0 ₽"),
        ("H5", "I6", "Ежемесячные потери (слив Next Step)", "=F6*0.18", font_kpi_red, fill_alert, "#,##0 ₽"),
        ("F8", "G9", "Возврат выручки (Месяц 1, 35%)", "=H6*0.35", font_kpi_green, fill_success, "#,##0 ₽"),
        ("H8", "I9", "Срок окупаемости подписки", '=ROUNDUP(IF(C6<=5,49000,IF(C6<=10,79000,120000))/(F9/30),0)&" дня"', font_kpi_green, fill_success, "@"),
    ]

    for top_l, btm_r, label, formula, f_font, fill, n_fmt in kpis:
        tl_col = top_l[0]
        tl_row = int(top_l[1:])
        br_col = btm_r[0]
        br_row = int(btm_r[1:])
        ws1.merge_cells(f"{top_l}:{br_col}{tl_row}")
        ws1.merge_cells(f"{tl_col}{br_row}:{btm_r}")

        ws1[top_l] = label
        ws1[top_l].font = font_sub
        ws1[top_l].fill = fill
        ws1[top_l].alignment = Alignment(horizontal="center", vertical="center")
        ws1[top_l].border = thin_border

        val_coord = f"{tl_col}{br_row}"
        ws1[val_coord] = formula
        ws1[val_coord].font = f_font
        ws1[val_coord].fill = fill
        ws1[val_coord].alignment = Alignment(horizontal="center", vertical="center")
        ws1[val_coord].border = thin_border
        ws1[val_coord].number_format = n_fmt

    # Детализация экономики
    ws1["B12"] = "ФИНАНСОВЫЙ РАСЧЕТ И ЭКОНОМИЧЕСКИЙ ЭФФЕКТ СИСТЕМЫ"
    ws1["B12"].font = font_section

    calc_rows = [
        ("1. Объем целевых квалифицированных сделок в месяц", "=C7*C9", "#,##0", "сделок"),
        ("2. Общий объем квалифицированного пайплайна в работе", "=C13*C8", "#,##0 ₽", "рублей"),
        ("3. Доля сделок, теряемых из-за отсутствия жесткого Next Step", "18.0%", "0.0%", "стандарт B2B"),
        ("4. Суммарная потерянная выручка компании в месяц", "=C14*0.18", "#,##0 ₽", "рублей/мес"),
        ("5. Возврат зависших сделок в первый месяц внедрения (35%)", "=C16*0.35", "#,##0 ₽", "чистый приток"),
        ("6. Рекомендуемый тариф подписки RevOps AI Supervisor", '=IF(C6<=5,"Старт (49 000 ₽)",IF(C6<=10,"Бизнес (79 000 ₽)","Enterprise (120 000 ₽)"))', "@", "в месяц"),
        ("7. Чистая прибыль компании от внедрения за месяц", "=C17-IF(C6<=5,49000,IF(C6<=10,79000,120000))", "#,##0 ₽", "чистая выгода"),
        ("8. Окупаемость инвестиций (ROI)", "=C19/IF(C6<=5,49000,IF(C6<=10,79000,120000))", "0%", "рентабельность"),
    ]

    for idx, (label, formula, n_fmt, note) in enumerate(calc_rows, start=13):
        ws1[f"B{idx}"] = label
        ws1[f"B{idx}"].font = font_bold if idx in [16, 17, 19, 20] else font_reg
        ws1[f"B{idx}"].border = thin_border
        
        ws1[f"C{idx}"] = formula
        ws1[f"C{idx}"].font = font_bold
        ws1[f"C{idx}"].border = thin_border
        ws1[f"C{idx}"].alignment = Alignment(horizontal="right", vertical="center")
        ws1[f"C{idx}"].number_format = n_fmt
        if idx in [16]:
            ws1[f"C{idx}"].font = Font(name="Segoe UI", size=10, bold=True, color="DC2626")
        elif idx in [17, 19, 20]:
            ws1[f"C{idx}"].font = Font(name="Segoe UI", size=10, bold=True, color="16A34A")

        ws1[f"D{idx}"] = note
        ws1[f"D{idx}"].font = font_sub
        ws1[f"D{idx}"].border = thin_border
        ws1.row_dimensions[idx].height = 22

    # Баннер подключения
    ws1.merge_cells("B22:I23")
    ws1["B22"] = "🔒 Подключение автоматической интеграции с вашей CRM (amoCRM / Битрикс24) и телефонией: 7 дней под ключ. Пилотный спринт: 29 000 ₽ (гарантия возврата)."
    ws1["B22"].font = Font(name="Segoe UI", size=10, bold=True, color="1E3A8A")
    ws1["B22"].fill = fill_indigo_light
    ws1["B22"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # =========================================================================
    # Вкладка 2: 📋 Пульт_РОПа_15_Минут (Операционный надзор)
    # =========================================================================
    ws2 = wb.create_sheet("📋 Пульт_РОПа_15_Минут")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:H1")
    ws2["A1"] = "Ежедневный Пульт Руководителя Отдела Продаж (Контроль за 15 минут в день)"
    ws2["A1"].font = font_title
    ws2["A1"].fill = fill_navy
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 35

    rop_headers = ["Ранг", "Менеджер ОП", "Звонков за сутки", "Ср. балл речи (из 13)", "Next Step %", "Talk/Listen (Норма ≥45%)", "Утечка выручки (мес)", "Статус и фокус РОПа"]
    for c_idx, h in enumerate(rop_headers, start=1):
        cell = ws2.cell(3, c_idx, h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[3].height = 26

    managers_data = [
        ("1", "Алексей Мельников", 14, 11.4, 0.93, "48% / 52%", 0, "🏆 Лидер недели (эталонные звонки)", fill_success, "16A34A"),
        ("2", "Елена Васильева", 12, 10.1, 0.83, "51% / 49%", 160000, "✅ В нормативе (хороший темп)", fill_success, "16A34A"),
        ("3", "Анна Соколова", 10, 8.2, 0.60, "55% / 45%", 480000, "⚠️ Скидки без торга (проработать цену)", fill_warn, "D97706"),
        ("4", "Дмитрий Ковалев", 13, 6.4, 0.38, "72% / 28%", 860000, "🚨 Слив Next Step (много говорит, не закрывает)", fill_alert, "DC2626"),
        ("5", "Иван Попов", 9, 5.8, 0.33, "79% / 21%", 940000, "🚨 Режим справочной (нет квалификации ЛПР)", fill_alert, "DC2626"),
    ]

    for r_idx, row in enumerate(managers_data, start=4):
        ws2.row_dimensions[r_idx].height = 24
        for c_idx in range(1, 9):
            cell = ws2.cell(r_idx, c_idx, row[c_idx-1])
            cell.font = font_reg
            cell.border = thin_border
            if c_idx in [1, 3]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx == 4:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.number_format = "0.0"
            elif c_idx == 5:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.number_format = "0%"
                cell.font = Font(name="Segoe UI", size=10, bold=True, color=row[9])
            elif c_idx == 6:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx == 7:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0 ₽"
                if row[6] > 0:
                    cell.font = Font(name="Segoe UI", size=10, bold=True, color="DC2626")
            elif c_idx == 8:
                cell.fill = row[8]
                cell.font = Font(name="Segoe UI", size=10, bold=True, color=row[9])

    # Итог по отделу
    ws2.row_dimensions[10].height = 25
    ws2.cell(10, 2, "ИТОГО ПО ОТДЕЛУ ПРОДАЖ").font = font_bold
    ws2.cell(10, 3, "=SUM(C4:C8)").number_format = "#,##0"
    ws2.cell(10, 4, "=AVERAGE(D4:D8)").number_format = "0.0"
    ws2.cell(10, 5, "=AVERAGE(E4:E8)").number_format = "0%"
    ws2.cell(10, 7, "=SUM(G4:G8)").number_format = "#,##0 ₽"
    ws2.cell(10, 8, "Устранимо за 1 спринт").font = font_bold
    for c in range(1, 9):
        ws2.cell(10, c).border = thick_bottom
        ws2.cell(10, c).font = font_bold

    # Оперативные задачи РОПа на сегодня
    ws2.cell(12, 1, "ГОРЯЩИЕ ЗАДАЧИ РОПА ИЗ TELEGRAM-ШЕРИФА НА СЕГОДНЯ:").font = font_section
    actions = [
        "1. [SOS] Сделка D-702 (Дмитрий Ковалев) — чек 750 000 ₽. Клиент сказал 'я подумаю', менеджер ответил 'ну спишемся'. Задача: перезвонить до 12:00, назначить Zoom на четверг.",
        "2. [СКИДКА] Сделка D-705 (Анна Соколова) — чек 500 000 ₽. Менеджер предложила скидку 20% без запроса. Задача: отозвать скидку, предложить рассрочку на 2 платежа.",
        "3. [ЛПР] Сделка D-708 (Иван Попов) — чек 1 200 000 ₽. Диалог ведется с секретарем уже 3 недели. Задача: выйти напрямую на Генерального директора через скрипт обхода.",
    ]
    for idx, act in enumerate(actions, start=13):
        ws2.merge_cells(f"A{idx}:H{idx}")
        ws2[f"A{idx}"] = act
        ws2[f"A{idx}"].font = font_reg
        ws2[f"A{idx}"].fill = fill_card
        ws2[f"A{idx}"].border = thin_border
        ws2.row_dimensions[idx].height = 22

    # =========================================================================
    # Вкладка 3: 🎙️ ИИ_Аудит_25_Звонков (Полный реестр 13 критериев)
    # =========================================================================
    ws3 = wb.create_sheet("🎙️ ИИ_Аудит_25_Звонков")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:W1")
    ws3["A1"] = "Детализированный реестр 100% звонков с разметкой по 13 критериям Whisper Large V3 + LLM"
    ws3["A1"].font = font_title
    ws3["A1"].fill = fill_navy
    ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 35

    audit_cols = [
        "Call_ID", "Deal_ID", "Менеджер", "Длит.", "Тип звонка",
        "CR1 Next", "CR2 Иниц.", "CR3 ЛПР", "CR4 Боль", "CR5 Цена",
        "CR6 Подумаю", "CR7 Кейсы", "CR8 Закрыт.", "CR9 Привет.", "CR10 Чистота",
        "CR11 Слуш.", "CR12 Итог", "CR13 Маржа", "Сумма (13)", "Next Step Статус",
        "Talk/Listen", "Главный дефект диалога", "Рекомендация ИИ-супервайзера"
    ]

    for c_idx, col_name in enumerate(audit_cols, start=1):
        cell = ws3.cell(3, c_idx, col_name)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws3.row_dimensions[3].height = 32

    # Генерация 20 реалистичных звонков
    sample_calls_matrix = [
        ("C-801", "D-101", "Алексей М.", "05:12", "Квалификация", [1,1,1,1,1,1,1,1,1,1,1,1,1], "Зафиксирован", "46%/54%", "Без дефектов", "Эталонный звонок. Добавить в тренинги."),
        ("C-802", "D-104", "Дмитрий К.", "03:40", "Переговоры", [0,1,1,1,1,0,0,1,1,0,0,0,1], "СЛИВ NEXT STEP", "74%/26%", "Слив Next Step", "Назначить точный слот Zoom в четверг 14:00."),
        ("C-803", "D-106", "Анна С.", "04:15", "Отработка цены", [1,1,1,0,0,1,1,1,1,1,1,1,0], "Зафиксирован", "52%/48%", "Скидка без торга", "Защищать цену через окупаемость и рассрочку."),
        ("C-804", "D-108", "Иван П.", "02:50", "Входящий лид", [0,0,0,0,1,1,0,0,1,0,0,0,1], "СЛИВ NEXT STEP", "81%/19%", "Режим справочной", "Перехватывать инициативу 3 квалифицирующими вопросами."),
        ("C-805", "D-111", "Елена В.", "06:20", "Презентация КП", [1,1,1,1,1,1,1,1,1,1,1,1,1], "Зафиксирован", "44%/56%", "Без дефектов", "Отличная связка кейсов с болями клиента."),
        ("C-806", "D-114", "Дмитрий К.", "04:30", "Дожим договора", [0,1,1,1,0,0,1,0,1,1,0,0,1], "СЛИВ NEXT STEP", "68%/32%", "Слив Next Step", "Предложить подписание через ЭДО до пятницы."),
        ("C-807", "D-115", "Анна С.", "05:05", "Квалификация", [1,1,0,1,1,1,1,1,1,1,1,1,1], "Зафиксирован", "49%/51%", "Не выявлен ЛПР", "Уточнить состав комиссии по закупкам."),
        ("C-808", "D-119", "Алексей М.", "04:45", "Переговоры", [1,1,1,1,1,1,1,1,1,1,1,1,1], "Зафиксирован", "47%/53%", "Без дефектов", "Четкое резюме договоренностей перед завершением."),
        ("C-809", "D-122", "Иван П.", "03:10", "Повторный контакт", [0,0,0,1,1,0,0,0,1,0,0,0,1], "СЛИВ NEXT STEP", "76%/24%", "Нет фиксации даты", "Всегда завершать звонок альтернативой слотов."),
        ("C-810", "D-125", "Елена В.", "05:50", "Согласование счета", [1,1,1,1,1,1,1,1,1,1,1,1,1], "Зафиксирован", "48%/52%", "Без дефектов", "Счет направлен, встреча назначена на завтра."),
        ("C-811", "D-128", "Дмитрий К.", "02:40", "Входящий лид", [0,1,0,0,1,0,0,0,1,1,0,0,1], "СЛИВ NEXT STEP", "70%/30%", "Слив Next Step", "Не отпускать клиента в 'подумайте'."),
        ("C-812", "D-131", "Анна С.", "04:20", "Отработка цены", [1,1,1,1,1,0,1,1,1,1,1,1,1], "Зафиксирован", "50%/50%", "Слабое 'подумаю'", "Вскрыть истинное сомнение: цена или доверие?"),
        ("C-813", "D-135", "Алексей М.", "05:30", "Квалификация", [1,1,1,1,1,1,1,1,1,1,1,1,1], "Зафиксирован", "45%/55%", "Без дефектов", "Образцовая квалификация бюджета."),
        ("C-814", "D-138", "Елена В.", "04:00", "Презентация КП", [1,1,1,1,1,1,1,0,1,1,1,1,1], "Зафиксирован", "49%/51%", "Нет попытки закрытия", "В конце встречи сразу предлагать бронь слота."),
        ("C-815", "D-140", "Иван П.", "03:30", "Переговоры", [0,0,1,0,0,0,0,0,1,0,0,0,0], "СЛИВ NEXT STEP", "82%/18%", "Слив маржи и Next Step", "Пройти тренинг по базовому регламенту продаж."),
    ]

    for r_idx, (cid, did, mgr, dur, call_tp, cr_list, step_st, tl, def_name, rec) in enumerate(sample_calls_matrix, start=4):
        ws3.row_dimensions[r_idx].height = 21
        ws3.cell(r_idx, 1, cid).alignment = Alignment(horizontal="center")
        ws3.cell(r_idx, 2, did).alignment = Alignment(horizontal="center")
        ws3.cell(r_idx, 3, mgr)
        ws3.cell(r_idx, 4, dur).alignment = Alignment(horizontal="center")
        ws3.cell(r_idx, 5, call_tp)

        # 13 критериев
        for c_offset, score in enumerate(cr_list):
            cell = ws3.cell(r_idx, 6 + c_offset, score)
            cell.alignment = Alignment(horizontal="center")
            if score == 1:
                cell.font = Font(name="Segoe UI", size=9, bold=True, color="16A34A")
                cell.fill = fill_success
            else:
                cell.font = Font(name="Segoe UI", size=9, bold=True, color="DC2626")
                cell.fill = fill_alert

        # Формула суммы баллов
        sum_cell = ws3.cell(r_idx, 19, f"=SUM(F{r_idx}:R{r_idx})")
        sum_cell.font = font_bold
        sum_cell.alignment = Alignment(horizontal="center")

        # Next Step
        step_cell = ws3.cell(r_idx, 20, step_st)
        step_cell.alignment = Alignment(horizontal="center")
        if "СЛИВ" in step_st:
            step_cell.font = Font(name="Segoe UI", size=9, bold=True, color="DC2626")
            step_cell.fill = fill_alert
        else:
            step_cell.font = Font(name="Segoe UI", size=9, bold=True, color="16A34A")
            step_cell.fill = fill_success

        ws3.cell(r_idx, 21, tl).alignment = Alignment(horizontal="center")
        ws3.cell(r_idx, 22, def_name)
        ws3.cell(r_idx, 23, rec)

        for col_i in range(1, 24):
            ws3.cell(r_idx, col_i).border = thin_border

    # =========================================================================
    # Вкладка 4: 💸 Диагностика_Утечек (Водопад потерь)
    # =========================================================================
    ws4 = wb.create_sheet("💸 Диагностика_Утечек")
    ws4.views.sheetView[0].showGridLines = True

    ws4.merge_cells("A1:G1")
    ws4["A1"] = "Анализ ключевых зон потерь выручки отдела продаж (Водопад Утечек)"
    ws4["A1"].font = font_title
    ws4["A1"].fill = fill_navy
    ws4["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[1].height = 35

    leak_headers = ["№", "Критическая зона потерь", "Частота дефекта", "Механика потерь выручки", "Потери в месяц (₽)", "Решение RevOps OS", "Эффект за 30 дней"]
    for c_idx, h in enumerate(leak_headers, start=1):
        cell = ws4.cell(3, c_idx, h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[3].height = 28

    leaks_full = [
        (1, "Слив жесткого Next Step (нет даты/времени)", "38% звонков", "Сделка зависает со статусом 'клиент думает', забывается и остывает", 1944000, "Шериф в Telegram: алерт РОПу через 15 минут после звонка без даты", "+680 000 ₽ возврата"),
        (2, "Необоснованная скидка при первом 'дорого'", "24% звонков", "Менеджер снижает маржу компании на 3-5% вместо защиты ценности", 510000, "ИИ-детектор скидок: блокировка выставления КП в CRM без авторизации", "+510 000 ₽ маржи"),
        (3, "Презентация продукта не ЛПР (секретарь/инженер)", "19% звонков", "Затягивание цикла сделки на 3-4 недели, слив презентаций впустую", 380000, "Обязательная матрица квалификации ЛПР из 2 вопросов в CRM", "-14 дней к циклу"),
        (4, "Монолог менеджера (Talk/Listen > 70%)", "31% звонков", "Клиент не вовлечен, менеджер читает лекцию, не вскрывая боли", 260000, "Дневной срез баланса слушания РОПу в Telegram", "+15% к SQL конверсии"),
    ]

    for r_idx, row in enumerate(leaks_full, start=4):
        ws4.row_dimensions[r_idx].height = 25
        for c_idx, val in enumerate(row, start=1):
            cell = ws4.cell(r_idx, c_idx, val)
            cell.font = font_reg
            cell.border = thin_border
            if c_idx in [1, 3]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx == 5:
                cell.number_format = "#,##0 ₽"
                cell.font = Font(name="Segoe UI", size=10, bold=True, color="DC2626")
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c_idx == 7:
                cell.font = Font(name="Segoe UI", size=10, bold=True, color="16A34A")
                cell.fill = fill_success

    # Итоговая строка
    ws4.row_dimensions[8].height = 26
    ws4.cell(8, 2, "СУММАРНЫЕ ВЫЯВЛЕННЫЕ УТЕЧКИ В МЕСЯЦ").font = font_bold
    ws4.cell(8, 5, "=SUM(E4:E7)").number_format = "#,##0 ₽"
    ws4.cell(8, 5).font = font_kpi_red
    ws4.cell(8, 5).alignment = Alignment(horizontal="right", vertical="center")
    ws4.cell(8, 7, "от 1 190 000 ₽ / мес").font = font_bold
    ws4.cell(8, 7).fill = fill_success
    for col_i in range(1, 8):
        ws4.cell(8, col_i).border = thick_bottom

    # =========================================================================
    # Вкладка 5: 🔐 152-ФЗ_Безопасность_и_Договор
    # =========================================================================
    ws5 = wb.create_sheet("🔐 152-ФЗ_Контур_Безопасности")
    ws5.views.sheetView[0].showGridLines = True

    ws5.merge_cells("A1:F1")
    ws5["A1"] = "Юридический Контур Безопасности и Обезличивания Персональных Данных (152-ФЗ РФ)"
    ws5["A1"].font = font_title
    ws5["A1"].fill = fill_navy
    ws5["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws5.row_dimensions[1].height = 35

    sec_headers = ["Параметр безопасности", "Требование 152-ФЗ РФ", "Реализация в RevOps Enterprise OS", "Статус проверки"]
    for c_idx, h in enumerate(sec_headers, start=1):
        cell = ws5.cell(3, c_idx, h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws5.row_dimensions[3].height = 26

    sec_data = [
        ("Локализация баз данных (ст. 18.1)", "Хранение ПДн граждан РФ строго на территории РФ", "Серверы в Москве (Yandex Cloud / Selectel, аттестат ФСТЭК УЗ 1-2)", "✅ Соответствует 100%"),
        ("Обезличивание данных (ст. 152-ФЗ)", "Криптографическая токенизация до передачи во внешние API", "Модуль PIISanitizer: замена телефонов и email на токены [PHONE_TOKEN_X]", "✅ Автоматически"),
        ("Срок хранения записей (ст. 5)", "Удаление данных после достижения цели обработки", "Автоматическое удаление аудиозаписей через 30 дней", "✅ Настроено"),
        ("Согласие сотрудников (ст. 86 ТК РФ)", "Уведомление сотрудников о контроле качества звонков", "Готовый комплект документов: Приказ, Положение о контроле, Согласие менеджера", "✅ Комплект включен"),
        ("Формат работы и налоги", "Полная прозрачность для юридических лиц и бухгалтерии", "Плательщик НПД (Самозанятый), ст. 15 422-ФЗ. Оплата по безналичному расчету, 0% НДС, чеки", "✅ Безрисково для бухгалтерии"),
    ]

    for r_idx, row in enumerate(sec_data, start=4):
        ws5.row_dimensions[r_idx].height = 28
        for c_idx, val in enumerate(row, start=1):
            cell = ws5.cell(r_idx, c_idx, val)
            cell.font = font_reg
            cell.border = thin_border
            if c_idx == 4:
                cell.font = Font(name="Segoe UI", size=10, bold=True, color="16A34A")
                cell.fill = fill_success
                cell.alignment = Alignment(horizontal="center", vertical="center")

    # Автоподбор ширины всех колонок
    for sheet in [ws1, ws2, ws3, ws4, ws5]:
        for col in sheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            sheet.column_dimensions[col_letter].width = max(11, min(max_len + 3, 50))

    # Специфические настройки ширин
    ws1.column_dimensions["B"].width = 52
    ws1.column_dimensions["C"].width = 24
    ws1.column_dimensions["D"].width = 18

    ws2.column_dimensions["B"].width = 24
    ws2.column_dimensions["H"].width = 45

    ws3.column_dimensions["C"].width = 18
    ws3.column_dimensions["V"].width = 28
    ws3.column_dimensions["W"].width = 46

    ws4.column_dimensions["B"].width = 38
    ws4.column_dimensions["D"].width = 45
    ws4.column_dimensions["F"].width = 45

    ws5.column_dimensions["A"].width = 32
    ws5.column_dimensions["B"].width = 38
    ws5.column_dimensions["C"].width = 48
    ws5.column_dimensions["D"].width = 22

    wb.save(output_path)
    print(f"Enterprise Demo Excel successfully created at: {output_path}")


if __name__ == "__main__":
    out = Path("presentation/RevOps_Platform_Demo_Sample.xlsx")
    create_enterprise_demo_excel(out)
    # Копируем в docs и на Рабочий стол
    import shutil
    shutil.copyfile(out, "docs/RevOps_Platform_Demo_Sample.xlsx")
    shutil.copyfile(out, "C:/Users/strel/Desktop/RevOps_Platform_Demo_Sample.xlsx")
    print("Copied to docs and Desktop!")
