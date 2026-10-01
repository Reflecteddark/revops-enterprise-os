"""
Builder for the Presentation Excel Workbook (Презентационная_Таблица_RevOps.xlsx).
Mirrors the user's Golden Master (https://docs.google.com/spreadsheets/d/1QnjrrbpqhYssofchee7G06szWrFvBCcOqGIVokjqdVc)
and includes the complete 13 speech evaluation criteria for AI call audit.
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path


def build_presentation_workbook(target_path: Path):
    wb = openpyxl.Workbook()
    # remove default sheet later or rename

    # Color Palette matching Golden Master
    c_navy_dark = "0F172A"      # Main dark header
    c_navy_light = "1E293B"     # Card & sub-header
    c_slate_header = "334155"   # Navigation bar fill
    c_indigo = "4F46E5"         # Accent blue/indigo
    c_indigo_light = "EEF2FF"
    c_emerald = "10B981"        # Success green
    c_emerald_light = "ECFDF5"
    c_emerald_border = "A7F3D0"
    c_red = "EF4444"            # Critical alert
    c_red_light = "FEF2F2"
    c_red_border = "FECACA"
    c_amber = "F59E0B"          # Warning amber
    c_amber_light = "FFFBEB"
    c_amber_border = "FDE68A"
    c_input_yellow = "FEF9C3"   # 3 Numbers Input cell
    c_card_bg = "F8FAFC"        # Card background
    c_border_gray = "E2E8F0"
    c_border_cell = "CBD5E1"
    c_text_dark = "0F172A"
    c_text_muted = "64748B"

    # Fills
    fill_navy = PatternFill(start_color=c_navy_dark, end_color=c_navy_dark, fill_type="solid")
    fill_navy_light = PatternFill(start_color=c_navy_light, end_color=c_navy_light, fill_type="solid")
    fill_nav_bar = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    fill_table_header = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    fill_input = PatternFill(start_color=c_input_yellow, end_color=c_input_yellow, fill_type="solid")
    fill_card = PatternFill(start_color=c_card_bg, end_color=c_card_bg, fill_type="solid")
    fill_success = PatternFill(start_color=c_emerald_light, end_color=c_emerald_light, fill_type="solid")
    fill_alert = PatternFill(start_color=c_red_light, end_color=c_red_light, fill_type="solid")
    fill_warn = PatternFill(start_color=c_amber_light, end_color=c_amber_light, fill_type="solid")
    fill_indigo_light = PatternFill(start_color=c_indigo_light, end_color=c_indigo_light, fill_type="solid")
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    # Fonts
    font_main_title = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
    font_nav = Font(name="Segoe UI", size=9, bold=True, color="94A3B8")
    font_nav_active = Font(name="Segoe UI", size=9, bold=True, color="60A5FA")
    font_sec_hdr = Font(name="Segoe UI", size=11, bold=True, color="1E293B")
    font_tbl_hdr = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
    font_bold = Font(name="Segoe UI", size=9, bold=True, color="0F172A")
    font_reg = Font(name="Segoe UI", size=9, color="334155")
    font_muted = Font(name="Segoe UI", size=8, italic=True, color="64748B")
    font_card_lbl = Font(name="Segoe UI", size=8, bold=True, color="64748B")
    font_card_val = Font(name="Segoe UI", size=15, bold=True, color="0F172A")
    font_card_red = Font(name="Segoe UI", size=15, bold=True, color="DC2626")
    font_card_green = Font(name="Segoe UI", size=15, bold=True, color="16A34A")
    font_card_blue = Font(name="Segoe UI", size=15, bold=True, color="2563EB")
    font_input = Font(name="Segoe UI", size=12, bold=True, color="B45309")
    font_link = Font(name="Segoe UI", size=10, bold=True, color="2563EB", underline="single")

    # Borders
    thin_border = Border(
        left=Side(style="thin", color=c_border_cell),
        right=Side(style="thin", color=c_border_cell),
        top=Side(style="thin", color=c_border_cell),
        bottom=Side(style="thin", color=c_border_cell)
    )
    card_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )
    total_top_border = Border(
        left=Side(style="thin", color=c_border_cell),
        right=Side(style="thin", color=c_border_cell),
        top=Side(style="thin", color=c_border_cell),
        bottom=Side(style="double", color=c_navy_dark)
    )

    nav_items = [
        "📄 1. One-Pager",
        "⚡ 2. Пульс компании",
        "📋 3. Пульт РОПа",
        "💸 4. 7 Грехов ОП",
        "🎯 5. Action Center",
        "🎙️ 6. ИИ-Аудит",
        "⚡ 7. Экспресс 3 цифры",
        "🧪 8. QA Suite",
        "⚙️ 9. Настройки"
    ]

    def setup_header_and_nav(ws, title_text, active_idx=6, max_col="I"):
        ws.views.sheetView[0].showGridLines = True
        ws.merge_cells(f"A1:{max_col}1")
        c1 = ws["A1"]
        c1.value = title_text
        c1.font = font_main_title
        c1.fill = fill_navy
        c1.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 36

        # Nav bar on Row 2
        for col_idx, item in enumerate(nav_items, start=1):
            if col_idx > openpyxl.utils.column_index_from_string(max_col):
                break
            cell = ws.cell(row=2, column=col_idx)
            cell.value = item
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = fill_nav_bar
            if col_idx == active_idx + 1:
                cell.font = font_nav_active
            else:
                cell.font = font_nav
        ws.row_dimensions[2].height = 22

    # =========================================================================
    # SHEET 1: ⚡ Экспресс_Калькулятор_3_Цифры (gid 777000201 replica)
    # =========================================================================
    ws_calc = wb.active
    ws_calc.title = "⚡ Экспресс_Калькулятор_3_Цифры"
    setup_header_and_nav(ws_calc, "⚡ ЭКСПРЕСС-КАЛЬКУЛЯТОР УПУЩЕННОЙ ПРИБЫЛИ ОП (ДИАГНОСТИКА ЗА 30 СЕКУНД)", active_idx=6, max_col="I")

    # Row 4: Section Headers
    ws_calc.merge_cells("A4:B4")
    ws_calc["A4"] = "🎛️ ВВЕДИТЕ 3 ЦИФРЫ ВАШЕГО БИЗНЕСА:"
    ws_calc["A4"].font = font_sec_hdr

    ws_calc.merge_cells("C4:D4")
    ws_calc["C4"] = "📊 БАЗОВЫЕ ПОКАЗАТЕЛИ (AS-IS):"
    ws_calc["C4"].font = font_sec_hdr

    ws_calc.merge_cells("G4:H4")
    ws_calc["G4"] = "Категория капитала"
    ws_calc["G4"].font = font_tbl_hdr
    ws_calc["G4"].fill = fill_table_header
    ws_calc["G4"].alignment = Alignment(horizontal="center", vertical="center")
    ws_calc["H4"].fill = fill_table_header

    ws_calc.row_dimensions[4].height = 24

    # Inputs & As-Is rows (5, 6, 7, 8)
    # R5: Лиды
    ws_calc["A5"] = "1. Входящий поток лидов в месяц (шт):"
    ws_calc["A5"].font = font_bold
    ws_calc["A5"].border = thin_border
    ws_calc["B5"] = 150
    ws_calc["B5"].font = font_input
    ws_calc["B5"].fill = fill_input
    ws_calc["B5"].alignment = Alignment(horizontal="center", vertical="center")
    ws_calc["B5"].number_format = "#,##0"
    ws_calc["B5"].border = thin_border

    ws_calc["C5"] = "Емкость базы (при 100% конверсии):"
    ws_calc["C5"].font = font_reg
    ws_calc["C5"].border = thin_border
    ws_calc["D5"] = "=B5*B6"
    ws_calc["D5"].font = font_bold
    ws_calc["D5"].border = thin_border
    ws_calc["D5"].number_format = "#,##0 \"₽\""

    ws_calc["G5"] = "Факт выручки (As-Is)"
    ws_calc["G5"].font = font_reg
    ws_calc["G5"].border = thin_border
    ws_calc["H5"] = 1650000
    ws_calc["H5"].font = font_bold
    ws_calc["H5"].border = thin_border
    ws_calc["H5"].number_format = "#,##0 \"₽\""

    # R6: Чек
    ws_calc["A6"] = "2. Средний чек закрытой сделки (₽):"
    ws_calc["A6"].font = font_bold
    ws_calc["A6"].border = thin_border
    ws_calc["B6"] = 400000
    ws_calc["B6"].font = font_input
    ws_calc["B6"].fill = fill_input
    ws_calc["B6"].alignment = Alignment(horizontal="center", vertical="center")
    ws_calc["B6"].number_format = "#,##0 \"₽\""
    ws_calc["B6"].border = thin_border

    ws_calc["C6"] = "Текущая конверсия в оплату (факт):"
    ws_calc["C6"].font = font_reg
    ws_calc["C6"].border = thin_border
    ws_calc["D6"] = "=H5/D5"
    ws_calc["D6"].font = font_bold
    ws_calc["D6"].border = thin_border
    ws_calc["D6"].number_format = "0.0%"

    ws_calc["G6"] = "Упущенная прибыль ОП"
    ws_calc["G6"].font = font_bold
    ws_calc["G6"].border = thin_border
    ws_calc["H6"] = 1039500
    ws_calc["H6"].font = font_card_red
    ws_calc["H6"].border = thin_border
    ws_calc["H6"].number_format = "#,##0 \"₽\""

    # R7: Менеджеры
    ws_calc["A7"] = "3. Количество менеджеров в ОП (чел):"
    ws_calc["A7"].font = font_bold
    ws_calc["A7"].border = thin_border
    ws_calc["B7"] = 4
    ws_calc["B7"].font = font_input
    ws_calc["B7"].fill = fill_input
    ws_calc["B7"].alignment = Alignment(horizontal="center", vertical="center")
    ws_calc["B7"].number_format = "0"
    ws_calc["B7"].border = thin_border

    ws_calc["C7"] = "Текущая выручка месяца (факт):"
    ws_calc["C7"].font = font_reg
    ws_calc["C7"].border = thin_border
    ws_calc["D7"] = "=H5"
    ws_calc["D7"].font = font_bold
    ws_calc["D7"].border = thin_border
    ws_calc["D7"].number_format = "#,##0 \"₽\""

    ws_calc["G7"] = "Потенциал кассы (To-Be)"
    ws_calc["G7"].font = font_bold
    ws_calc["G7"].border = thin_border
    ws_calc["H7"] = 1961850
    ws_calc["H7"].font = font_card_green
    ws_calc["H7"].border = thin_border
    ws_calc["H7"].number_format = "#,##0 \"₽\""

    # R8: Отрасль и нагрузка
    ws_calc["A8"] = "Отраслевой сегмент компании:"
    ws_calc["A8"].font = font_reg
    ws_calc["A8"].border = thin_border
    ws_calc["B8"] = "B2B Услуги / Производство"
    ws_calc["B8"].font = font_bold
    ws_calc["B8"].border = thin_border

    ws_calc["C8"] = "Нагрузка на одного менеджера (лидов):"
    ws_calc["C8"].font = font_reg
    ws_calc["C8"].border = thin_border
    ws_calc["D8"] = "=ROUND(B5/B7, 0) & \" лид/мес\""
    ws_calc["D8"].font = font_bold
    ws_calc["D8"].border = thin_border

    # R10 & R11: 4 KPI Cards (Сценарии роста)
    kpis_r10 = [
        ("A", "B", "🔴 ВЫЯВЛЕННЫЙ РИСК ПОТЕРЬ", fill_alert, font_card_red, "=H6", "#,##0 \"₽\""),
        ("C", "D", "🟢 КОНСЕРВАТИВНЫЙ (+15%)", fill_success, font_card_green, "=H5*1.0945", "#,##0 \"₽\""),
        ("E", "F", "🚀 БАЗОВЫЙ ПЛАН (+30%)", fill_indigo_light, font_card_blue, "=H5*1.189", "#,##0 \"₽\""),
        ("G", "H", "⭐ ОПТИМИСТИЧНЫЙ (+50%)", fill_warn, font_card_green, "=H5*1.315", "#,##0 \"₽\""),
    ]

    for start_c, end_c, title, fill_c, font_v, formula_v, num_fmt in kpis_r10:
        c_title = f"{start_c}10"
        c_val = f"{start_c}11"
        ws_calc.merge_cells(f"{start_c}10:{end_c}10")
        ws_calc.merge_cells(f"{start_c}11:{end_c}11")
        ws_calc[c_title] = title
        ws_calc[c_title].font = font_bold
        ws_calc[c_title].fill = fill_c
        ws_calc[c_title].alignment = Alignment(horizontal="center", vertical="center")
        ws_calc[c_title].border = card_border

        ws_calc[c_val] = formula_v
        ws_calc[c_val].font = font_v
        ws_calc[c_val].fill = fill_c
        ws_calc[c_val].alignment = Alignment(horizontal="center", vertical="center")
        ws_calc[c_val].number_format = num_fmt
        ws_calc[c_val].border = card_border

    ws_calc.row_dimensions[10].height = 22
    ws_calc.row_dimensions[11].height = 32

    # R13: Header 4 Узких Места
    ws_calc.merge_cells("A13:F13")
    ws_calc["A13"] = "📊 РАСШИФРОВКА ФИНАНСОВЫХ УТЕЧЕК ПО 4 УЗКИМ МЕСТАМ"
    ws_calc["A13"].font = font_sec_hdr
    ws_calc.row_dimensions[13].height = 24

    # R14: Table Headers
    leaks_headers = [
        ("A14", "№"),
        ("B14", "Узкое место воронки продаж"),
        ("C14", "Отраслевой бенчмарк потерь"),
        ("D14", "Упущенная прибыль, ₽/мес"),
        ("E14", "Причина утечки"),
        ("F14", "Как закрываем за 14 дней")
    ]
    for cell_id, text in leaks_headers:
        ws_calc[cell_id] = text
        ws_calc[cell_id].font = font_tbl_hdr
        ws_calc[cell_id].fill = fill_table_header
        ws_calc[cell_id].alignment = Alignment(horizontal="center", vertical="center")
        ws_calc[cell_id].border = thin_border
    ws_calc.row_dimensions[14].height = 24

    leaks_data = [
        ("1", "Слив лидов на звонках (нет Next Step)", "30% выручки теряется из-за ошибок менеджеров", 412500, "Менеджеры консультируют, но не закрывают на дату", "ИИ-аудит 100% звонков + чек-лист Next Step"),
        ("2", "Зависание сделок на этапе КП > 48ч", "18% выручки теряется из-за просрочки follow-up", 297000, "Клиент остывает, пока менеджер ждет звонка", "Авторадар SLA 48ч + эскалация РОПу на 3-й день"),
        ("3", "Брошенные отказники без дожима (L1-L4)", "8% — потенциал реактивации списанных сделок", 132000, "Сделки списаны в архив без реактивации", "AI Recovery Engine: автодожим в WhatsApp/TG"),
        ("4", "Зависшая дебиторка и задержка оплат", "12% выставленных счетов с просрочкой 20+ дней", 198000, "Нет платежного календаря и контроля сроков", "Платежный календарь DSO + автонапоминания")
    ]

    for row_idx, data in enumerate(leaks_data, start=15):
        ws_calc[f"A{row_idx}"] = data[0]
        ws_calc[f"A{row_idx}"].alignment = Alignment(horizontal="center")
        ws_calc[f"B{row_idx}"] = data[1]
        ws_calc[f"C{row_idx}"] = data[2]
        ws_calc[f"D{row_idx}"] = data[3]
        ws_calc[f"D{row_idx}"].number_format = "#,##0 \"₽\""
        ws_calc[f"D{row_idx}"].font = font_bold
        ws_calc[f"E{row_idx}"] = data[4]
        ws_calc[f"F{row_idx}"] = data[5]

        for col_l in ["A", "B", "C", "D", "E", "F"]:
            ws_calc[f"{col_l}{row_idx}"].border = thin_border
            if col_l != "D":
                ws_calc[f"{col_l}{row_idx}"].font = font_reg
        ws_calc.row_dimensions[row_idx].height = 20

    # R19: Total Row
    ws_calc["B19"] = "ИТОГО УПУЩЕННОЙ ПРИБЫЛИ В МЕСЯЦ:"
    ws_calc["B19"].font = font_bold
    ws_calc["D19"] = "=SUM(D15:D18)"
    ws_calc["D19"].font = font_card_red
    ws_calc["D19"].number_format = "#,##0 \"₽\""
    ws_calc["E19"] = "ПОТЕНЦИАЛ БЫСТРОГО ВОЗВРАТА:"
    ws_calc["E19"].font = font_bold
    ws_calc["F19"] = "=ROUND(D19*0.35, 0) & \" ₽ в первый месяц\""
    ws_calc["F19"].font = font_card_green

    for col_l in ["A", "B", "C", "D", "E", "F"]:
        ws_calc[f"{col_l}{row_idx+1}"].border = total_top_border

    # R21-R27: CTA Banner
    ws_calc.merge_cells("A21:F21")
    ws_calc["A21"] = "🎁 СЛЕДУЮЩИЙ ШАГ: БЕСПЛАТНЫЙ ТЕСТ-ДРАЙВ НА 3 ВАШИХ ЗВОНКАХ"
    ws_calc["A21"].font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    ws_calc["A21"].fill = fill_navy
    ws_calc["A21"].alignment = Alignment(horizontal="center", vertical="center")
    ws_calc.row_dimensions[21].height = 26

    ws_calc.merge_cells("A22:F22")
    ws_calc["A22"] = "Хотите проверить эту математику на реальных данных вашей компании за 15 минут?"
    ws_calc["A22"].font = font_bold
    ws_calc["A22"].alignment = Alignment(horizontal="center")

    ws_calc.merge_cells("A23:F23")
    ws_calc["A23"] = "👉 Отправьте 3 аудиозаписи любых вчерашних звонков ваших менеджеров в Telegram бот."
    ws_calc["A23"].font = font_reg
    ws_calc["A23"].alignment = Alignment(horizontal="center")

    ws_calc.merge_cells("A24:F24")
    ws_calc["A24"] = "ИИ бесплатно разберет их по 13 критериям, покажет ошибки речи и рассчитает точную сумму под угрозой слива."
    ws_calc["A24"].font = font_reg
    ws_calc["A24"].alignment = Alignment(horizontal="center")

    ws_calc.merge_cells("A25:F25")
    ws_calc["A25"] = "🚀 ЗАПУСТИТЬ БЕСПЛАТНЫЙ ТЕСТ-ДРАЙВ: @RevOps_Super_Audit_Bot (https://t.me/RevOps_Super_Audit_Bot)"
    ws_calc["A25"].font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    ws_calc["A25"].fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    ws_calc["A25"].alignment = Alignment(horizontal="center", vertical="center")
    ws_calc.row_dimensions[25].height = 28

    ws_calc.merge_cells("A26:F26")
    ws_calc["A26"] = "Нужен комплексный аудит отдела продаж? 👉 Экономика спринта оптимизации (450 000 ₽) на вкладке «4. 7 Грехов ОП»"
    ws_calc["A26"].font = font_link
    ws_calc["A26"].alignment = Alignment(horizontal="center")

    ws_calc.merge_cells("A27:F27")
    ws_calc["A27"] = "🔒 Конфиденциально • Без передачи персональных данных • Экспресс-разбор за 15 минут (152-ФЗ РФ)"
    ws_calc["A27"].font = font_muted
    ws_calc["A27"].alignment = Alignment(horizontal="center")

    # Column widths
    ws_calc.column_dimensions["A"].width = 38
    ws_calc.column_dimensions["B"].width = 24
    ws_calc.column_dimensions["C"].width = 38
    ws_calc.column_dimensions["D"].width = 26
    ws_calc.column_dimensions["E"].width = 44
    ws_calc.column_dimensions["F"].width = 46
    ws_calc.column_dimensions["G"].width = 28
    ws_calc.column_dimensions["H"].width = 24
    ws_calc.column_dimensions["I"].width = 20

    # =========================================================================
    # SHEET 2: 📋 Пульт_РОПа_15_Минут (gid 777000301 replica)
    # =========================================================================
    ws_rop = wb.create_sheet("📋 Пульт_РОПа_15_Минут")
    setup_header_and_nav(ws_rop, "📋 СТРАТЕГИЧЕСКИЙ ПУЛЬТ СОБСТВЕННИКА: «5 РЕШЕНИЙ МЕСЯЦА»", active_idx=2, max_col="H")

    # Row 4 & 5: Top KPI cards
    rop_kpis = [
        ("A", "Звонков за вчера", "6", fill_card, font_card_val),
        ("B", "Диалогов с браком (<9)", "2", fill_alert, font_card_red),
        ("C", "Сделок с угрозой срыва", "2", fill_alert, font_card_red),
        ("D", "Сумма в зоне риска прямо сейчас", 2730000, fill_alert, font_card_red),
        ("E", "H", "Статус дня для РОПа", "🚨 ТРЕБУЕТСЯ 15 МИНУТ КОНТРОЛЯ РОПа", fill_alert, font_card_red),
    ]

    for item in rop_kpis[:4]:
        col_l, label, val, fill_c, font_v = item
        ws_rop[f"{col_l}4"] = label
        ws_rop[f"{col_l}4"].font = font_card_lbl
        ws_rop[f"{col_l}4"].fill = fill_c
        ws_rop[f"{col_l}4"].alignment = Alignment(horizontal="center", vertical="center")
        ws_rop[f"{col_l}4"].border = card_border

        ws_rop[f"{col_l}5"] = val
        ws_rop[f"{col_l}5"].font = font_v
        ws_rop[f"{col_l}5"].fill = fill_c
        ws_rop[f"{col_l}5"].alignment = Alignment(horizontal="center", vertical="center")
        ws_rop[f"{col_l}5"].border = card_border
        if isinstance(val, int):
            ws_rop[f"{col_l}5"].number_format = "#,##0 \"₽\""

    # Merge status E4:H4 and E5:H5
    ws_rop.merge_cells("E4:H4")
    ws_rop["E4"] = "Статус дня для РОПа"
    ws_rop["E4"].font = font_card_lbl
    ws_rop["E4"].fill = fill_alert
    ws_rop["E4"].alignment = Alignment(horizontal="center", vertical="center")
    ws_rop["E4"].border = card_border

    ws_rop.merge_cells("E5:H5")
    ws_rop["E5"] = "🚨 ТРЕБУЕТСЯ 15 МИНУТ КОНТРОЛЯ РОПа"
    ws_rop["E5"].font = Font(name="Segoe UI", size=13, bold=True, color="DC2626")
    ws_rop["E5"].fill = fill_alert
    ws_rop["E5"].alignment = Alignment(horizontal="center", vertical="center")
    ws_rop["E5"].border = card_border

    ws_rop.row_dimensions[4].height = 20
    ws_rop.row_dimensions[5].height = 32

    # Row 7: Section header
    ws_rop.merge_cells("A7:H7")
    ws_rop["A7"] = "🎯 ТОП-5 СТРАТЕГИЧЕСКИХ РЕШЕНИЙ ДЛЯ ЗАЩИТЫ КАПИТАЛА БИЗНЕСА"
    ws_rop["A7"].font = font_sec_hdr
    ws_rop.row_dimensions[7].height = 24

    # Row 8: Table Header
    rop_tbl_hdrs = [
        ("A8", "№"),
        ("B8", "Сделка / Контрагент"),
        ("C8", "Суть проблемы и симптом"),
        ("D8", "Сумма в риске"),
        ("E8", "Приоритет"),
        ("F8", "Срочное действие РОПа (Что сделать)"),
        ("G8", "Чекбокс"),
        ("H8", "🤖 Сгенерированный оффер ИИ-дожима (WhatsApp / Скрипт РОПа)")
    ]
    for cell_id, text in rop_tbl_hdrs:
        ws_rop[cell_id] = text
        ws_rop[cell_id].font = font_tbl_hdr
        ws_rop[cell_id].fill = fill_table_header
        ws_rop[cell_id].alignment = Alignment(horizontal="center", vertical="center")
        ws_rop[cell_id].border = thin_border
    ws_rop.row_dimensions[8].height = 26

    # 5 Decisions data
    rop_decisions = [
        (1, "Сделка D-104 (ООО «Вектор Плюс»)", "Зависание на этапе КП: 12 дней без ответа (норма 48ч)", 850000, "🔴 КРИТИЧЕСКИЙ", "Запросить статус у КАМа Петрова, согласовать скидку/рассрочку", "В РАБОТЕ", "Добрый день, Алексей! Вижу, наш расчет по договору согласован. Чтобы зафиксировать спец-условия сентября со скидкой 7%, предлагаю сегодня подписать спецификацию. Слот для звонка: 14:30. Удобно?"),
        (2, "Звонок C-505 / Сделка D-109 (ООО «ПромКомплект»)", "Менеджер сорвал контакт (балл речи 2/13), клиент ушел в отказ", 600000, "🔴 КРИТИЧЕСКИЙ", "Личный звонок РОПа ЛПР до 12:00: перехват сделки от лица руководства", "В РАБОТЕ", "Михаил Сергеевич, добрый день! Это Дмитрий Иванов, коммерческий директор. Прослушал вчерашний диалог с нашим менеджером — приношу извинения за неполную консультацию. Лично подготовил для вас проект с расширенной гарантией 24 мес. Уделите 5 минут?"),
        (3, "Звонок C-503 / Сделка D-106 (ООО «СнабСервис»)", "В диалоге не зафиксирована дата встречи (балл речи 6/13)", 280000, "🟡 СРЕДНИЙ", "Поставить задачу менеджеру Петрову: перезвонить и назначить слот встречи", "В РАБОТЕ", "Игорь, приветствую! Направил вам 2 слота на онлайн-демо платформы: сегодня в 16:00 или завтра в 11:30. За 15 минут покажем связку с вашей CRM. Какой слот забронировать?"),
        (4, "Сделка D-107 (ООО «ТехноПром»)", "Счет просрочен на 9 дней (сумма 150 000 ₽)", 150000, "🟡 СРЕДНИЙ", "Отправить триггерное напоминание в WhatsApp бухгалтеру клиента", "В РАБОТЕ", "Ольга, здравствуйте! Бухгалтерия напоминает: по счету INV-203 сегодня крайний день резерва оборудования по льготному курсу. Выставить ссылку на моментальную оплату через СБП?"),
        (5, "Сделка D-103 (ООО «ИнвестХолдинг»)", "Нагрузка 80%, 4 ключевые сделки висят на РОПе", 850000, "🟡 СРЕДНИЙ", "Делегировать 2 входящие сделки на КАМа (у КАМа загрузка всего 4%)", "В РАБОТЕ", "Регламентное действие: передать права ответственного по сделке D-102 на КАМа Петрова (свободная емкость 75%). РОПу оставить контроль ключевых контрольных точек.")
    ]

    for idx, row_data in enumerate(rop_decisions, start=9):
        ws_rop[f"A{idx}"] = row_data[0]
        ws_rop[f"A{idx}"].alignment = Alignment(horizontal="center", vertical="center")
        ws_rop[f"B{idx}"] = row_data[1]
        ws_rop[f"C{idx}"] = row_data[2]
        ws_rop[f"D{idx}"] = row_data[3]
        ws_rop[f"D{idx}"].number_format = "#,##0 \"₽\""
        ws_rop[f"D{idx}"].font = font_bold
        ws_rop[f"E{idx}"] = row_data[4]
        ws_rop[f"E{idx}"].alignment = Alignment(horizontal="center", vertical="center")
        ws_rop[f"F{idx}"] = row_data[5]
        ws_rop[f"G{idx}"] = row_data[6]
        ws_rop[f"G{idx}"].alignment = Alignment(horizontal="center", vertical="center")
        ws_rop[f"G{idx}"].font = font_bold
        ws_rop[f"H{idx}"] = row_data[7]
        ws_rop[f"H{idx}"].font = Font(name="Segoe UI", size=8, italic=True, color="334155")

        for col_l in ["A", "B", "C", "D", "E", "F", "G", "H"]:
            ws_rop[f"{col_l}{idx}"].border = thin_border
            if col_l not in ["D", "G", "H"]:
                ws_rop[f"{col_l}{idx}"].font = font_reg
        ws_rop.row_dimensions[idx].height = 42

    # R14: Total Row
    ws_rop["B14"] = "ИТОГО ДЕНЕГ В ЗОНЕ РИСКА СЕГОДНЯ:"
    ws_rop["B14"].font = font_bold
    ws_rop["D14"] = "=SUM(D9:D13)"
    ws_rop["D14"].font = font_card_red
    ws_rop["D14"].number_format = "#,##0 \"₽\""
    ws_rop["E14"] = "СПАСЕННАЯ ВЫРУЧКА:"
    ws_rop["E14"].font = font_bold
    ws_rop["F14"] = "При выполнении всех 5 действий"
    ws_rop["F14"].font = font_reg
    ws_rop["G14"] = "100% защита"
    ws_rop["G14"].font = font_card_green

    for col_l in ["A", "B", "C", "D", "E", "F", "G", "H"]:
        ws_rop[f"{col_l}14"].border = total_top_border
    ws_rop.row_dimensions[14].height = 24

    # R17-R22: 15-Minute ROP Morning Protocol
    ws_rop.merge_cells("E17:H17")
    ws_rop["E17"] = "⏱️ ЭКСПРЕСС-РЕГЛАМЕНТ РОПа (15 МИНУТ/ДЕНЬ)"
    ws_rop["E17"].font = font_sec_hdr
    ws_rop["E17"].alignment = Alignment(horizontal="center", vertical="center")
    ws_rop["E17"].fill = fill_indigo_light
    ws_rop.row_dimensions[17].height = 24

    reglament_rows = [
        ("09:00 - 09:05", "Проверка ночного ИИ-аудита: 2 звонка с браком (<9 баллов)"),
        ("09:05 - 09:10", "Перехват сделки D-109 (600 т.р.) и звонок ЛПР до 12:00"),
        ("09:10 - 09:13", "Согласование спецусловий по D-104 (850 т.р.)"),
        ("09:13 - 09:15", "Делегирование 2 сделок с РОПа на свободного КАМа"),
        ("Результат:", "2 730 000 ₽ защищено от слива за 15 минут")
    ]

    for idx, (t_range, t_desc) in enumerate(reglament_rows, start=18):
        ws_rop[f"E{idx}"] = t_range
        ws_rop[f"E{idx}"].font = font_bold
        ws_rop[f"E{idx}"].alignment = Alignment(horizontal="center", vertical="center")
        ws_rop[f"E{idx}"].border = thin_border

        ws_rop.merge_cells(f"F{idx}:H{idx}")
        ws_rop[f"F{idx}"] = t_desc
        ws_rop[f"F{idx}"].font = font_bold if idx == 22 else font_reg
        if idx == 22:
            ws_rop[f"F{idx}"].font = font_card_green
            ws_rop[f"E{idx}"].fill = fill_success
            ws_rop[f"F{idx}"].fill = fill_success
        ws_rop[f"F{idx}"].border = thin_border
        ws_rop[f"G{idx}"].border = thin_border
        ws_rop[f"H{idx}"].border = thin_border
        ws_rop.row_dimensions[idx].height = 20

    ws_rop.column_dimensions["A"].width = 6
    ws_rop.column_dimensions["B"].width = 36
    ws_rop.column_dimensions["C"].width = 46
    ws_rop.column_dimensions["D"].width = 20
    ws_rop.column_dimensions["E"].width = 22
    ws_rop.column_dimensions["F"].width = 42
    ws_rop.column_dimensions["G"].width = 16
    ws_rop.column_dimensions["H"].width = 65

    # =========================================================================
    # SHEET 3: 🎙️ ИИ_Аудит (gid 852624872 replica + 13 EVALUATION PARAMETERS!)
    # =========================================================================
    ws_audit = wb.create_sheet("🎙️ ИИ_Аудит")
    setup_header_and_nav(ws_audit, "🎙️ РЕЧЕВАЯ ИИ-АНАЛИТИКА: 13 ПАРАМЕТРОВ WHISPER + LLM", active_idx=5, max_col="J")

    # Row 3 & 4: Top KPI Cards
    audit_kpis = [
        ("A", "Звонков проанализировано", "6", fill_card, font_card_val),
        ("B", "C", "Средний балл качества", "9.5 из 13", fill_indigo_light, font_card_blue),
        ("D", "Доля фиксации Next Step", "66.7%", fill_card, font_card_val),
        ("E", "Диалогов с браком (<9)", "2", fill_alert, font_card_red),
        ("F", "G", "Сумма пайплайна под угрозой", 880000, fill_alert, font_card_red),
        ("H", "J", "Финансовый вердикт ИИ", "🚨 ВЫСОКИЙ РИСК СЛИВА VIP-КЛИЕНТА", fill_alert, font_card_red),
    ]

    # Render audit KPIs
    ws_audit["A3"] = "Звонков проанализировано"
    ws_audit["A3"].font = font_card_lbl
    ws_audit["A3"].fill = fill_card
    ws_audit["A3"].alignment = Alignment(horizontal="center", vertical="center")
    ws_audit["A3"].border = card_border
    ws_audit["A4"] = 6
    ws_audit["A4"].font = font_card_val
    ws_audit["A4"].fill = fill_card
    ws_audit["A4"].alignment = Alignment(horizontal="center", vertical="center")
    ws_audit["A4"].border = card_border

    ws_audit.merge_cells("B3:C3")
    ws_audit["B3"] = "Средний балл качества"
    ws_audit["B3"].font = font_card_lbl
    ws_audit["B3"].fill = fill_indigo_light
    ws_audit["B3"].alignment = Alignment(horizontal="center", vertical="center")
    ws_audit["B3"].border = card_border
    ws_audit.merge_cells("B4:C4")
    ws_audit["B4"] = "9.5 из 13"
    ws_audit["B4"].font = font_card_blue
    ws_audit["B4"].fill = fill_indigo_light
    ws_audit["B4"].alignment = Alignment(horizontal="center", vertical="center")
    ws_audit["B4"].border = card_border

    ws_audit["D3"] = "Доля фиксации Next Step"
    ws_audit["D3"].font = font_card_lbl
    ws_audit["D3"].fill = fill_card
    ws_audit["D3"].alignment = Alignment(horizontal="center", vertical="center")
    ws_audit["D3"].border = card_border
    ws_audit["D4"] = "66.7%"
    ws_audit["D4"].font = font_card_val
    ws_audit["D4"].fill = fill_card
    ws_audit["D4"].alignment = Alignment(horizontal="center", vertical="center")
    ws_audit["D4"].border = card_border

    ws_audit["E3"] = "Диалогов с браком (<9)"
    ws_audit["E3"].font = font_card_lbl
    ws_audit["E3"].fill = fill_alert
    ws_audit["E3"].alignment = Alignment(horizontal="center", vertical="center")
    ws_audit["E3"].border = card_border
    ws_audit["E4"] = 2
    ws_audit["E4"].font = font_card_red
    ws_audit["E4"].fill = fill_alert
    ws_audit["E4"].alignment = Alignment(horizontal="center", vertical="center")
    ws_audit["E4"].border = card_border

    ws_audit.merge_cells("F3:G3")
    ws_audit["F3"] = "Сумма пайплайна под угрозой"
    ws_audit["F3"].font = font_card_lbl
    ws_audit["F3"].fill = fill_alert
    ws_audit["F3"].alignment = Alignment(horizontal="center", vertical="center")
    ws_audit["F3"].border = card_border
    ws_audit.merge_cells("F4:G4")
    ws_audit["F4"] = 880000
    ws_audit["F4"].font = font_card_red
    ws_audit["F4"].fill = fill_alert
    ws_audit["F4"].alignment = Alignment(horizontal="center", vertical="center")
    ws_audit["F4"].number_format = "#,##0 \"₽\""
    ws_audit["F4"].border = card_border

    ws_audit.merge_cells("H3:J3")
    ws_audit["H3"] = "Финансовый вердикт ИИ"
    ws_audit["H3"].font = font_card_lbl
    ws_audit["H3"].fill = fill_alert
    ws_audit["H3"].alignment = Alignment(horizontal="center", vertical="center")
    ws_audit["H3"].border = card_border
    ws_audit.merge_cells("H4:J4")
    ws_audit["H4"] = "🚨 ВЫСОКИЙ РИСК СЛИВА VIP-КЛИЕНТА"
    ws_audit["H4"].font = Font(name="Segoe UI", size=11, bold=True, color="DC2626")
    ws_audit["H4"].fill = fill_alert
    ws_audit["H4"].alignment = Alignment(horizontal="center", vertical="center")
    ws_audit["H4"].border = card_border

    ws_audit.row_dimensions[3].height = 20
    ws_audit.row_dimensions[4].height = 30

    # Row 6: Section 1 Header (Call Registry)
    ws_audit.merge_cells("A6:J6")
    ws_audit["A6"] = "📋 РЕЕСТР ПРОБЛЕМНЫХ ЗВОНКОВ И РЕКОМЕНДАЦИИ ИИ (СИНХРОНИЗАЦИЯ С CRM)"
    ws_audit["A6"].font = font_sec_hdr
    ws_audit.row_dimensions[6].height = 24

    audit_tbl_hdrs = [
        ("A7", "ID звонка"),
        ("B7", "ID сделки"),
        ("C7", "Менеджер ID"),
        ("D7", "Длительность"),
        ("E7", "Балл /13"),
        ("F7", "Ошибка диалога"),
        ("G7", "ИИ-рекомендация"),
        ("H7", "Сумма сделки под угрозой, ₽"),
        ("I7", "Уровень фин. риска"),
        ("J7", "Срочное действие РОПа")
    ]
    for cell_id, text in audit_tbl_hdrs:
        ws_audit[cell_id] = text
        ws_audit[cell_id].font = font_tbl_hdr
        ws_audit[cell_id].fill = fill_table_header
        ws_audit[cell_id].alignment = Alignment(horizontal="center", vertical="center")
        ws_audit[cell_id].border = thin_border
    ws_audit.row_dimensions[7].height = 26

    audit_calls_data = [
        ("C-505", "D-109", "102 (Петров)", "210 сек", 2, "Срыв контакта / грубость", "Личный звонок РОПа с извинениями", 600000, "🔴 ВЫСОКИЙ (VIP-чек)", "🚨 Личный перезвон РОПа ЛПР"),
        ("C-503", "D-106", "102 (Петров)", "180 сек", 6, "Нет фиксации даты встречи", "Направить 2 слота в мессенджер", 280000, "🟡 СРЕДНИЙ", "Тренинг менеджера по скрипту"),
        ("C-502", "D-104", "101 (Сидоров)", "310 сек", 11, "Слабая отработка возражения цены", "Предложить расчет окупаемости и рассрочку", 850000, "🟢 КОНТРОЛЬ", "Перепроверить отправку КП"),
        ("C-504", "D-107", "103 (Иванова)", "540 сек", 12, "Не выявлен дедлайн бюджета", "Задать прямой вопрос о сроках согласования", 150000, "🟢 КОНТРОЛЬ", "Контроль выставления счета"),
        ("C-501", "D-101", "102 (Петров)", "420 сек", 13, "Без дефектов (Эталонный диалог)", "Использовать в базе знаний компании", 700000, "⭐ ЭТАЛОН", "Похвала сотрудника на летучке"),
        ("C-506", "D-108", "103 (Иванова)", "290 сек", 13, "Без дефектов (Эталонный диалог)", "Быстрый квалификационный переход", 320000, "⭐ ЭТАЛОН", "Перевод сделки на этап демо")
    ]

    for idx, cdata in enumerate(audit_calls_data, start=8):
        ws_audit[f"A{idx}"] = cdata[0]
        ws_audit[f"A{idx}"].alignment = Alignment(horizontal="center")
        ws_audit[f"B{idx}"] = cdata[1]
        ws_audit[f"B{idx}"].alignment = Alignment(horizontal="center")
        ws_audit[f"C{idx}"] = cdata[2]
        ws_audit[f"D{idx}"] = cdata[3]
        ws_audit[f"D{idx}"].alignment = Alignment(horizontal="center")

        ws_audit[f"E{idx}"] = cdata[4]
        ws_audit[f"E{idx}"].alignment = Alignment(horizontal="center")
        ws_audit[f"E{idx}"].font = font_card_red if cdata[4] < 9 else font_bold

        ws_audit[f"F{idx}"] = cdata[5]
        ws_audit[f"G{idx}"] = cdata[6]

        ws_audit[f"H{idx}"] = cdata[7]
        ws_audit[f"H{idx}"].number_format = "#,##0 \"₽\""
        ws_audit[f"H{idx}"].font = font_bold

        ws_audit[f"I{idx}"] = cdata[8]
        ws_audit[f"I{idx}"].alignment = Alignment(horizontal="center")
        ws_audit[f"J{idx}"] = cdata[9]

        for col_l in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]:
            ws_audit[f"{col_l}{idx}"].border = thin_border
            if col_l not in ["E", "H"]:
                ws_audit[f"{col_l}{idx}"].font = font_reg
        ws_audit.row_dimensions[idx].height = 22

    # Total Row for Calls
    ws_audit["B14"] = "ИТОГО ПОТЕРЬ ИЗ-ЗА БРАКА ЗВОНКОВ:"
    ws_audit["B14"].font = font_bold
    ws_audit["H14"] = "=SUM(H8:H9)"
    ws_audit["H14"].font = font_card_red
    ws_audit["H14"].number_format = "#,##0 \"₽\""
    ws_audit["I14"] = "ИТОГО В РИСКЕ"
    ws_audit["I14"].font = font_bold
    ws_audit["I14"].alignment = Alignment(horizontal="center")
    ws_audit["J14"] = "Срочный разбор с менеджерами"
    ws_audit["J14"].font = font_reg

    for col_l in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]:
        ws_audit[f"{col_l}14"].border = total_top_border
    ws_audit.row_dimensions[14].height = 24

    # -------------------------------------------------------------------------
    # TABLE 2: EXACT 13 SPEECH EVALUATION PARAMETERS (Requested by User!)
    # -------------------------------------------------------------------------
    ws_audit.merge_cells("A16:J16")
    ws_audit["A16"] = "🎙️ ЭТАЛОННАЯ СИСТЕМА 13 КРИТЕРИЕВ ОЦЕНКИ ЗВОНКОВ (WHISPER LARGE V3 + LLM SUPERVISOR)"
    ws_audit["A16"].font = font_sec_hdr
    ws_audit["A16"].fill = fill_indigo_light
    ws_audit["A16"].alignment = Alignment(horizontal="center", vertical="center")
    ws_audit.row_dimensions[16].height = 26

    crit_headers = [
        ("A17", "№"),
        ("B17", "Параметр оценки диалога"),
        ("C17", "Вес"),
        ("D17", "Приоритет"),
        ("E17", "Что проверяет алгоритм ИИ (Суть критерия)"),
        ("F17", "G17", "Маркеры брака и речевые дефекты"),
        ("H17", "J17", "Coaching Tip: как обучить менеджера и предотвратить слив")
    ]

    for item in crit_headers:
        if len(item) == 2:
            cell_id, text = item
            ws_audit[cell_id] = text
            ws_audit[cell_id].fill = fill_table_header
            ws_audit[cell_id].font = font_tbl_hdr
            ws_audit[cell_id].alignment = Alignment(horizontal="center", vertical="center")
            ws_audit[cell_id].border = thin_border
        else:
            s_cell, e_cell, text = item
            ws_audit.merge_cells(f"{s_cell}:{e_cell}")
            ws_audit[s_cell] = text
            ws_audit[s_cell].fill = fill_table_header
            ws_audit[s_cell].font = font_tbl_hdr
            ws_audit[s_cell].alignment = Alignment(horizontal="center", vertical="center")
            ws_audit[s_cell].border = thin_border
            # border on covered cells
            start_col = s_cell[0]
            end_col = e_cell[0]
            r_num = s_cell[1:]
            for c_code in range(ord(start_col), ord(end_col) + 1):
                ws_audit[f"{chr(c_code)}{r_num}"].border = thin_border
                ws_audit[f"{chr(c_code)}{r_num}"].fill = fill_table_header

    ws_audit.row_dimensions[17].height = 26

    criteria_13 = [
        (1, "Жёсткий Next Step", "2.0x", "🔴 P0 (Критический)", "Зафиксирована точная дата и время следующего контакта (календарный слот)", "«Я вам перезвоню на днях», «Подумайте, наберите», отпускание без времени", "Внедрить обязательный вопрос в конце диалога: «Во вторник в 14:30 или в среду в 11:00 удобно?»"),
        (2, "Инициатива диалога", "1.0x", "🟡 P1 (Высокий)", "Менеджер ведёт диалог вопросами, держит сценарий, а не работает автоответчиком", "Клиент задает все вопросы сам, менеджер только отвечает короткими репликами", "Правило «Ответ + Встречный вопрос». Завершать каждую реплику квалифицирующим вопросом."),
        (3, "Квалификация ЛПР / ЛВПР", "1.5x", "🔴 P0 (Критический)", "Установлен статус собеседника: принимает ли решения по бюджету и согласованию", "Трата 40 минут на презентацию инженеру или секретарю без полномочий подписания", "«Кто кроме вас участвует в согласовании проекта и утверждении финансового бюджета?»"),
        (4, "Выявление болей и срочности", "1.2x", "🟡 P1 (Высокий)", "Понятна ли бизнес-задача клиента, текущие потери и дедлайн решения проблемы", "Менеджер сразу начинает презентовать продукт, не узнав текущую ситуацию клиента", "Сначала диагностика по методологии SPIN/BANT, затем презентация под конкретную боль."),
        (5, "Отработка возражения «Дорого»", "1.5x", "🔴 P0 (Критический)", "Обоснование ценности, расчет ROI/окупаемости, рассрочка (без моментальной скидки)", "Моментальная скидка при первом сомнении клиента или согласие («Да, у нас недешево»)", "Обосновать окупаемость: показать, сколько клиент теряет каждый день без внедрения решения."),
        (6, "Отработка «Я подумаю»", "1.0x", "🟡 P1 (Высокий)", "Вскрытие истинного сомнения клиента, а не молчаливое отпускание в архив", "«Хорошо, думайте, ждем вашего звонка», завершение диалога на неопределенности", "«Алексей, обычно говорят 'подумаю', когда смущает цена или функционал. Что именно из этого?»"),
        (7, "Презентация через кейсы и цифры", "1.0x", "🟡 P1 (Высокий)", "Приведение аналогичных проектов, оцифрованных результатов и сроков возврата", "Абстрактные слова «мы надежные», «у нас высокое качество» без конкретных фактов", "Показать схожий кейс из отрасли клиента: «Для завода X мы сократили потери на 1.4 млн ₽ за 21 день»."),
        (8, "Попытка закрытия сделки", "1.5x", "🔴 P0 (Критический)", "Прямое предложение заключить договор, забронировать слот аудита или выставить счет", "Разговор закончился тепло, но менеджер не предложил сделать следующий шаг к оплате", "«Если все параметры подходят, предлагаю сегодня запустить пилот. Куда отправить договор?»"),
        (9, "Регламент приветствия и повестка", "0.8x", "🟢 P2 (Стандарт)", "Представление имени, компании, уточнение времени и согласование цели разговора", "Невнятное приветствие, вопрос «Вам удобно говорить?» без обозначения повестки", "«Добрый день! Дмитрий, компания RevOps. Звоню по вашей заявке на аудит воронки. Уделите 5 минут?»"),
        (10, "Чистота речи и уверенность", "0.8x", "🟢 P2 (Стандарт)", "Отсутствие слов-паразитов («как бы», «э-э-э»), пауз и извиняющихся интонаций при цене", "Дрожащий голос при озвучивании чека, слова сомнения («может быть», «наверное»)", "Отработать блок цены на тренажере: уверенная пауза после озвучивания стоимости."),
        (11, "Слушание клиента (Talk/Listen)", "1.2x", "🟡 P1 (Высокий)", "Менеджер слушает клиента не менее 45–50% общего времени разговора", "Монолог менеджера > 70% времени: клиент молчит и теряет фокус внимания", "Держать баланс 50/50. Задавать открытые вопросы и давать клиенту выговориться о проблеме."),
        (12, "Фиксация договорённостей", "1.0x", "🟢 P2 (Стандарт)", "Резюмирование договоренностей перед окончанием звонка: кто, что и когда делает", "Диалог оборвался, у клиента и менеджера разное понимание следующих действий", "«Итак, резюмирую: я отправляю КП в WhatsApp до 15:00, а завтра в 11:00 мы созваниваемся по условиям»."),
        (13, "Защита маржинальности", "1.5x", "🔴 P0 (Критический)", "Отсутствие необоснованных скидок без встречных уступок по объему или предоплате", "Раздача скидок из маржи компании без требования 100% предоплаты или большего объема", "Скидка возможна только в обмен на ценность: «Готовы дать 5% при оплате счета в течение 24 часов».")
    ]

    for idx, c in enumerate(criteria_13, start=18):
        c_num, c_name, c_weight, c_pri, c_desc, c_defect, c_tip = c
        ws_audit[f"A{idx}"] = c_num
        ws_audit[f"A{idx}"].alignment = Alignment(horizontal="center", vertical="center")
        ws_audit[f"B{idx}"] = c_name
        ws_audit[f"B{idx}"].font = font_bold
        ws_audit[f"C{idx}"] = c_weight
        ws_audit[f"C{idx}"].alignment = Alignment(horizontal="center", vertical="center")
        ws_audit[f"D{idx}"] = c_pri
        ws_audit[f"D{idx}"].alignment = Alignment(horizontal="center", vertical="center")
        if "🔴" in c_pri:
            ws_audit[f"D{idx}"].fill = fill_alert
            ws_audit[f"D{idx}"].font = font_bold
        elif "🟡" in c_pri:
            ws_audit[f"D{idx}"].fill = fill_warn
            ws_audit[f"D{idx}"].font = font_bold
        else:
            ws_audit[f"D{idx}"].fill = fill_success
            ws_audit[f"D{idx}"].font = font_reg

        ws_audit[f"E{idx}"] = c_desc
        ws_audit.merge_cells(f"F{idx}:G{idx}")
        ws_audit[f"F{idx}"] = c_defect
        ws_audit.merge_cells(f"H{idx}:J{idx}")
        ws_audit[f"H{idx}"] = c_tip

        for col_l in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]:
            ws_audit[f"{col_l}{idx}"].border = thin_border
            if col_l not in ["B", "D"]:
                ws_audit[f"{col_l}{idx}"].font = font_reg
        ws_audit.row_dimensions[idx].height = 32

    # Column dimensions for audit sheet
    ws_audit.column_dimensions["A"].width = 12
    ws_audit.column_dimensions["B"].width = 28
    ws_audit.column_dimensions["C"].width = 14
    ws_audit.column_dimensions["D"].width = 16
    ws_audit.column_dimensions["E"].width = 12
    ws_audit.column_dimensions["F"].width = 32
    ws_audit.column_dimensions["G"].width = 38
    ws_audit.column_dimensions["H"].width = 26
    ws_audit.column_dimensions["I"].width = 22
    ws_audit.column_dimensions["J"].width = 34

    # =========================================================================
    # SHEET 4: 💸 Диагностика_Утечек_ОП (gid 777000101 replica)
    # =========================================================================
    ws_sins = wb.create_sheet("💸 Диагностика_Утечек_ОП")
    setup_header_and_nav(ws_sins, "💸 ДИАГНОСТИЧЕСКИЙ АУДИТ УПУЩЕННОЙ ПРИБЫЛИ И УЗКИХ МЕСТ ОТДЕЛА ПРОДАЖ", active_idx=3, max_col="H")

    # Row 4 & 5: Top KPI cards
    sins_kpis = [
        ("A", "Фактическая выручка месяца", 1650000, fill_card, font_card_val, "#,##0 \"₽\""),
        ("B", "Взвешенный пайплайн в работе", 2754000, fill_indigo_light, font_card_blue, "#,##0 \"₽\""),
        ("C", "Выявленная упущенная прибыль", 7001000, fill_alert, font_card_red, "#,##0 \"₽\""),
        ("D", "Потенциал выручки (To-Be)", 8651000, fill_success, font_card_green, "#,##0 \"₽\""),
        ("E", "F", "Ожидаемый чистый ROI", "367%", fill_success, font_card_green, None),
        ("G", "H", "Срок окупаемости проекта", "6 дн.", fill_success, font_card_green, None),
    ]

    for col_l, label, val, fill_c, font_v, num_fmt in sins_kpis[:4]:
        ws_sins[f"{col_l}4"] = label
        ws_sins[f"{col_l}4"].font = font_card_lbl
        ws_sins[f"{col_l}4"].fill = fill_c
        ws_sins[f"{col_l}4"].alignment = Alignment(horizontal="center", vertical="center")
        ws_sins[f"{col_l}4"].border = card_border

        ws_sins[f"{col_l}5"] = val
        ws_sins[f"{col_l}5"].font = font_v
        ws_sins[f"{col_l}5"].fill = fill_c
        ws_sins[f"{col_l}5"].alignment = Alignment(horizontal="center", vertical="center")
        ws_sins[f"{col_l}5"].border = card_border
        if num_fmt:
            ws_sins[f"{col_l}5"].number_format = num_fmt

    # Merge ROI & Payback
    ws_sins.merge_cells("E4:F4")
    ws_sins["E4"] = "Ожидаемый чистый ROI"
    ws_sins["E4"].font = font_card_lbl
    ws_sins["E4"].fill = fill_success
    ws_sins["E4"].alignment = Alignment(horizontal="center", vertical="center")
    ws_sins["E4"].border = card_border
    ws_sins.merge_cells("E5:F5")
    ws_sins["E5"] = "367%"
    ws_sins["E5"].font = font_card_green
    ws_sins["E5"].fill = fill_success
    ws_sins["E5"].alignment = Alignment(horizontal="center", vertical="center")
    ws_sins["E5"].border = card_border

    ws_sins.merge_cells("G4:H4")
    ws_sins["G4"] = "Срок окупаемости проекта"
    ws_sins["G4"].font = font_card_lbl
    ws_sins["G4"].fill = fill_success
    ws_sins["G4"].alignment = Alignment(horizontal="center", vertical="center")
    ws_sins["G4"].border = card_border
    ws_sins.merge_cells("G5:H5")
    ws_sins["G5"] = "6 дн."
    ws_sins["G5"].font = font_card_green
    ws_sins["G5"].fill = fill_success
    ws_sins["G5"].alignment = Alignment(horizontal="center", vertical="center")
    ws_sins["G5"].border = card_border

    ws_sins.row_dimensions[4].height = 20
    ws_sins.row_dimensions[5].height = 32

    # Row 7: Section Header 7 Sins
    ws_sins.merge_cells("A7:H7")
    ws_sins["A7"] = "🔴 7 СМЕРТНЫХ ГРЕХОВ ОТДЕЛА ПРОДАЖ (АНАТОМИЯ СБОЕВ И ИХ УСТРАНЕНИЕ)"
    ws_sins["A7"].font = font_sec_hdr
    ws_sins.row_dimensions[7].height = 24

    sins_headers = [
        ("A8", "№"),
        ("B8", "Смертный грех ОП (Системный сбой)"),
        ("C8", "Симптом и диагностика в компании"),
        ("D8", "Объем базы в риске"),
        ("E8", "Упущенная выгода / Риск, ₽"),
        ("F8", "Приоритет"),
        ("G8", "Как закрывает наша платформа"),
        ("H8", "Ожидаемый эффект")
    ]
    for cell_id, text in sins_headers:
        ws_sins[cell_id] = text
        ws_sins[cell_id].font = font_tbl_hdr
        ws_sins[cell_id].fill = fill_table_header
        ws_sins[cell_id].alignment = Alignment(horizontal="center", vertical="center")
        ws_sins[cell_id].border = thin_border
    ws_sins.row_dimensions[8].height = 26

    sins_data = [
        ("1", "«Черный ящик» CRM: иллюзорный пайплайн", "В CRM висит 5.3 млн ₽, но реальный взвешенный объем всего 2.6 млн ₽. Фаундер живет в иллюзии плана", 5530000, 2776000, "🔴 P0 (Критический)", "Взвешенный прогноз воронки (SSOT) + исключение зомби-сделок", "100% точность прогноза Run-Rate без иллюзий"),
        ("2", "«Кладбище сделок»: срыв регламентов SLA", "Сделки маринуются на этапе КП по 12–14 дней (норма 48ч). Вероятность закрытия падает на 45%", 3850000, 1732500, "🔴 P0 (Критический)", "Авторадар зависания сделок + эскалация РОПу на 3-й день", "+45% к конверсии этапа КП в оплату"),
        ("3", "РОП — «играющий тренер» и узкое горлышко", "РОП ведет 80% сделок и перегружен (нагрузка 80%), а КАМы простаивают с 4% загрузки. Сделки тухнут без внимания", 850000, 255000, "🔴 P0 (Критический)", "Матрица емкости (Capacity Planning) + автоперелив сделок на КАМов", "Разгрузка РОПа на 60% для управления и коучинга"),
        ("4", "Менеджеры-автоответчики и слив звонков", "В 33% звонков не зафиксирован Next Step. Средний балл речи < 9/13. Лиды уходят в «подумаю»", 5530000, 1382500, "🔴 P0 (Критический)", "13-факторная речевая аналитика + чек-листы BANT + мотивация от речи", "+30% к конверсии из первого звонка во встречу"),
        ("5", "«Кладбище отказников» без реактивации", "1 млн ₽ списан в отказ (L1 «дорого», L2 «подумаю»). База брошена в архиве без повторного прогрева", 1000000, 440000, "🟡 P1 (Высокий)", "Agentic AI Recovery Engine: сегментация отказников + офферы даунсейла", "Возврат до 20–30% списанной выручки из архива"),
        ("6", "Зависшая дебиторка и кассовые разрывы", "Счета выставлены, но не оплачены (DSO > 25 дней). 330к просрочки грозит кассовым разрывом", 330000, 330000, "🟡 P1 (Высокий)", "Платежный календарь DSO + триггерные AI-уведомления бухгалтерии клиента", "Сокращение срока дебиторки DSO с 27 до 7 дней"),
        ("7", "Слив маркетинга и война с отделом продаж", "Google Ads слил 85 000 ₽ с ROMI -100% (0 оплат). Деньги сожжены без сквозного контроля окупаемости", 850000, 85000, "🟡 P1 (Высокий)", "Сквозная RevOps-атрибуция W-Shaped: автоотключение убыточных связок", "Экономия рекламного бюджета + перелив в прибыльный Контур/Директ")
    ]

    for idx, s in enumerate(sins_data, start=9):
        ws_sins[f"A{idx}"] = s[0]
        ws_sins[f"A{idx}"].alignment = Alignment(horizontal="center", vertical="center")
        ws_sins[f"B{idx}"] = s[1]
        ws_sins[f"B{idx}"].font = font_bold
        ws_sins[f"C{idx}"] = s[2]
        ws_sins[f"D{idx}"] = s[3]
        ws_sins[f"D{idx}"].number_format = "#,##0 \"₽\""
        ws_sins[f"D{idx}"].font = font_reg
        ws_sins[f"E{idx}"] = s[4]
        ws_sins[f"E{idx}"].number_format = "#,##0 \"₽\""
        ws_sins[f"E{idx}"].font = font_bold
        ws_sins[f"F{idx}"] = s[5]
        ws_sins[f"F{idx}"].alignment = Alignment(horizontal="center", vertical="center")
        if "🔴" in s[5]:
            ws_sins[f"F{idx}"].fill = fill_alert
        else:
            ws_sins[f"F{idx}"].fill = fill_warn
        ws_sins[f"G{idx}"] = s[6]
        ws_sins[f"H{idx}"] = s[7]

        for col_l in ["A", "B", "C", "D", "E", "F", "G", "H"]:
            ws_sins[f"{col_l}{idx}"].border = thin_border
            if col_l not in ["B", "E", "F"]:
                ws_sins[f"{col_l}{idx}"].font = font_reg
        ws_sins.row_dimensions[idx].height = 36

    # Total Row
    ws_sins["B16"] = "ИТОГО ВЫЯВЛЕННЫХ ПОТЕРЬ И РИСКОВ В МЕСЯЦ:"
    ws_sins["B16"].font = font_bold
    ws_sins["D16"] = "=SUM(D9:D15)"
    ws_sins["D16"].font = font_bold
    ws_sins["D16"].number_format = "#,##0 \"₽\""
    ws_sins["E16"] = "=SUM(E9:E15)"
    ws_sins["E16"].font = font_card_red
    ws_sins["E16"].number_format = "#,##0 \"₽\""
    ws_sins["F16"] = "ИТОГО ПОТЕРЬ"
    ws_sins["F16"].font = font_bold
    ws_sins["F16"].alignment = Alignment(horizontal="center")
    ws_sins["G16"] = "Потенциал возврата в кассу"
    ws_sins["G16"].font = font_reg
    ws_sins["H16"] = "Окупаемость за счет устранения узких мест"
    ws_sins["H16"].font = font_card_green

    for col_l in ["A", "B", "C", "D", "E", "F", "G", "H"]:
        ws_sins[f"{col_l}16"].border = total_top_border
    ws_sins.row_dimensions[16].height = 24

    # Project Economics Block
    ws_sins.merge_cells("A18:H18")
    ws_sins["A18"] = "🚀 ЭКОНОМИКА ПРОЕКТА ОПТИМИЗАЦИИ И СРОК ОКУПАЕМОСТИ ДЛЯ КЛИЕНТА"
    ws_sins["A18"].font = font_sec_hdr
    ws_sins["A18"].fill = fill_indigo_light
    ws_sins["A18"].alignment = Alignment(horizontal="center", vertical="center")
    ws_sins.row_dimensions[18].height = 26

    econ_hdrs = [
        ("A19", "C19", "Параметр экономики проекта"),
        ("D19", "Значение"),
        ("E19", "H19", "Комментарий и логика расчета")
    ]
    for s_c, *rest in econ_hdrs:
        if rest:
            e_c, text = rest[0], rest[1] if len(rest) > 1 else rest[0]
            if len(rest) == 1:
                text = e_c
                e_c = s_c
            ws_sins.merge_cells(f"{s_c}:{e_c}")
            ws_sins[s_c] = text
            ws_sins[s_c].font = font_tbl_hdr
            ws_sins[s_c].fill = fill_table_header
            ws_sins[s_c].alignment = Alignment(horizontal="center", vertical="center")
        else:
            cell_id, text = s_c, rest[0]
            ws_sins[cell_id] = text
            ws_sins[cell_id].font = font_tbl_hdr
            ws_sins[cell_id].fill = fill_table_header
            ws_sins[cell_id].alignment = Alignment(horizontal="center", vertical="center")
    ws_sins.row_dimensions[19].height = 24

    econ_data = [
        ("Стоимость проекта оптимизации ОП", 450000, "#,##0 \"₽\"", "Фиксированная стоимость внедрения (4 недели спринтов)"),
        ("Ожидаемый возврат выручки (M1 факт)", 2100300, "#,##0 \"₽\"", "Консервативный ориентир: устранение всего 30% найденных узких мест"),
        ("Чистая прибыль собственника за вычетом услуг", "=D21-D20", "#,##0 \"₽\"", "Чистый денежный плюс в кассе компании в первый же месяц"),
        ("Чистый ROI инвестиций в проект", "=(D22/D20)*100 & \"%\"", None, "Процент отдачи на каждый вложенный рубль"),
        ("Срок полной окупаемости проекта", "=ROUND(D20/(D21/30), 0) & \" дн.\"", None, "Количество дней работы ОП до полной окупаемости чека внедрения")
    ]

    for idx, (param, val, num_fmt, comm) in enumerate(econ_data, start=20):
        ws_sins.merge_cells(f"A{idx}:C{idx}")
        ws_sins[f"A{idx}"] = param
        ws_sins[f"A{idx}"].font = font_bold
        ws_sins[f"D{idx}"] = val
        ws_sins[f"D{idx}"].font = font_card_green if idx in [21, 22, 23, 24] else font_bold
        if num_fmt:
            ws_sins[f"D{idx}"].number_format = num_fmt
        ws_sins[f"D{idx}"].alignment = Alignment(horizontal="center", vertical="center")

        ws_sins.merge_cells(f"E{idx}:H{idx}")
        ws_sins[f"E{idx}"] = comm
        ws_sins[f"E{idx}"].font = font_reg

        for c_code in range(ord("A"), ord("H") + 1):
            ws_sins[f"{chr(c_code)}{idx}"].border = thin_border
        ws_sins.row_dimensions[idx].height = 22

    # CTAs
    ws_sins.merge_cells("A26:H26")
    ws_sins["A26"] = "👉 УТВЕРДИТЬ ПРОЕКТ ОПТИМИЗАЦИИ И СФОРМИРОВАТЬ ИНВОЙС (450 000 ₽)"
    ws_sins["A26"].font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    ws_sins["A26"].fill = PatternFill(start_color="16A34A", end_color="16A34A", fill_type="solid")
    ws_sins["A26"].alignment = Alignment(horizontal="center", vertical="center")
    ws_sins.row_dimensions[26].height = 28

    ws_sins.merge_cells("A27:H27")
    ws_sins["A27"] = "Не готовы к спринту? 👉 Бесплатный тест-драйв на 3 звонках на вкладке «⚡ Экспресс_Калькулятор_3_Цифры»"
    ws_sins["A27"].font = font_link
    ws_sins["A27"].alignment = Alignment(horizontal="center")

    ws_sins.merge_cells("A28:H28")
    ws_sins["A28"] = "🔒 Безопасный B2B-расчет: безналичный платеж по счету (УПД / Диадок) • 152-ФЗ РФ • Старт спринта за 24 часа"
    ws_sins["A28"].font = font_muted
    ws_sins["A28"].alignment = Alignment(horizontal="center")

    ws_sins.column_dimensions["A"].width = 6
    ws_sins.column_dimensions["B"].width = 38
    ws_sins.column_dimensions["C"].width = 44
    ws_sins.column_dimensions["D"].width = 22
    ws_sins.column_dimensions["E"].width = 26
    ws_sins.column_dimensions["F"].width = 22
    ws_sins.column_dimensions["G"].width = 44
    ws_sins.column_dimensions["H"].width = 38

    # =========================================================================
    # SHEET 5: 📄 Executive_OnePager (gid 1377308911 replica)
    # =========================================================================
    ws_one = wb.create_sheet("📄 Executive_OnePager")
    setup_header_and_nav(ws_one, "📄 REVOPS ENTERPRISE EXECUTIVE ONE-PAGER (УПРАВЛЕНЧЕСКИЙ ДАЙДЖЕСТ)", active_idx=0, max_col="I")

    # Row 4 & 5: 8 Core Executive KPIs
    one_kpis = [
        ("A", "Выручка (Факт Closed-Won)", 1650000, font_card_val, "#,##0 \"₽\""),
        ("B", "% Выполнения плана", "33.0%", font_card_val, None),
        ("C", "Взвешенный прогноз (SSOT)", 2754000, font_card_blue, "#,##0 \"₽\""),
        ("D", "Run-Rate касса (Прогноз)", 1706897, font_card_val, "#,##0 \"₽\""),
        ("E", "Активный пайплайн в работе", 5530000, font_card_val, "#,##0 \"₽\""),
        ("F", "Сделок под угрозой SLA", "4", font_card_red, None),
        ("G", "Качество речи ОП (ИИ)", "73.1%", font_card_val, None),
        ("H", "I", "Data Quality Index", "100.0%", font_card_green, None)
    ]

    for item in one_kpis[:7]:
        col_l, label, val, font_v, num_fmt = item
        ws_one[f"{col_l}4"] = label
        ws_one[f"{col_l}4"].font = font_card_lbl
        ws_one[f"{col_l}4"].fill = fill_card
        ws_one[f"{col_l}4"].alignment = Alignment(horizontal="center", vertical="center")
        ws_one[f"{col_l}4"].border = card_border

        ws_one[f"{col_l}5"] = val
        ws_one[f"{col_l}5"].font = font_v
        ws_one[f"{col_l}5"].fill = fill_card
        ws_one[f"{col_l}5"].alignment = Alignment(horizontal="center", vertical="center")
        ws_one[f"{col_l}5"].border = card_border
        if num_fmt:
            ws_one[f"{col_l}5"].number_format = num_fmt

    # Merge H4:I4 and H5:I5 for Data Quality
    ws_one.merge_cells("H4:I4")
    ws_one["H4"] = "Data Quality Index"
    ws_one["H4"].font = font_card_lbl
    ws_one["H4"].fill = fill_success
    ws_one["H4"].alignment = Alignment(horizontal="center", vertical="center")
    ws_one["H4"].border = card_border

    ws_one.merge_cells("H5:I5")
    ws_one["H5"] = "100.0%"
    ws_one["H5"].font = font_card_green
    ws_one["H5"].fill = fill_success
    ws_one["H5"].alignment = Alignment(horizontal="center", vertical="center")
    ws_one["H5"].border = card_border

    ws_one.row_dimensions[4].height = 20
    ws_one.row_dimensions[5].height = 30

    # Row 7: Section Headers
    ws_one.merge_cells("A7:D7")
    ws_one["A7"] = "📊 ВОРОНКА ПРОДАЖ & УЗКИЕ ГОРЛЫШКИ (AS-IS)"
    ws_one["A7"].font = font_sec_hdr

    ws_one.merge_cells("F7:I7")
    ws_one["F7"] = "🚨 ТОП-5 КРИТИЧЕСКИХ РИСКОВ ВЫРУЧКИ (ACTION CENTER)"
    ws_one["F7"].font = font_sec_hdr

    ws_one.row_dimensions[7].height = 24

    # Row 8: Headers
    left_hdrs = [("A8", "Этап сделки"), ("B8", "Сделок"), ("C8", "Объем, ₽"), ("D8", "SLA норматив (ч)")]
    for cell_id, text in left_hdrs:
        ws_one[cell_id] = text
        ws_one[cell_id].font = font_tbl_hdr
        ws_one[cell_id].fill = fill_table_header
        ws_one[cell_id].alignment = Alignment(horizontal="center", vertical="center")
        ws_one[cell_id].border = thin_border

    right_hdrs = [("F8", "Сделка / Клиент"), ("G8", "Проблема & Симптом"), ("H8", "Сумма в риске, ₽"), ("I8", "Срочное действие РОПа")]
    for cell_id, text in right_hdrs:
        ws_one[cell_id] = text
        ws_one[cell_id].font = font_tbl_hdr
        ws_one[cell_id].fill = fill_table_header
        ws_one[cell_id].alignment = Alignment(horizontal="center", vertical="center")
        ws_one[cell_id].border = thin_border

    ws_one.row_dimensions[8].height = 24

    funnel_stages = [
        ("1. Новый лид", 1, 950000, 24),
        ("2. Квалификация / ЛПР", 1, 120000, 48),
        ("3. Встреча / Демо", 1, 280000, 120),
        ("4. КП и согласование", 2, 3850000, 48),
        ("5. Счет выставлен", 2, 330000, 72),
        ("6. Успешно реализовано", 2, 1650000, 0)
    ]

    action_risks = [
        ("D-104 (ООО Вектор)", "КП без ответа > 48ч", 850000, "Повторный контакт с ЛПР + звонок РОПа"),
        ("D-103 (ИнвестХолдинг)", "Просрочен счёт на оплату", 150000, "Дожим счета через AI-бота + автозвонок"),
        ("D-109 (ПромКомплект)", "Критический брак речи (<9)", 600000, "Личный перезвон РОПа ЛПР до 12:00"),
        ("D-106 (СнабСервис)", "Зависание на этапе КП", 280000, "Эскалация РОПу, согласование слота"),
        ("D-107 (ТехноПром)", "Крупная сделка без движения", 3000000, "Подключение генерального директора")
    ]

    for idx in range(6):
        r_num = 9 + idx
        # Funnel stage
        f_stage, f_cnt, f_sum, f_sla = funnel_stages[idx]
        ws_one[f"A{r_num}"] = f_stage
        ws_one[f"A{r_num}"].font = font_reg
        ws_one[f"A{r_num}"].border = thin_border
        ws_one[f"B{r_num}"] = f_cnt
        ws_one[f"B{r_num}"].alignment = Alignment(horizontal="center")
        ws_one[f"B{r_num}"].font = font_reg
        ws_one[f"B{r_num}"].border = thin_border
        ws_one[f"C{r_num}"] = f_sum
        ws_one[f"C{r_num}"].number_format = "#,##0 \"₽\""
        ws_one[f"C{r_num}"].font = font_bold
        ws_one[f"C{r_num}"].border = thin_border
        ws_one[f"D{r_num}"] = f_sla
        ws_one[f"D{r_num}"].alignment = Alignment(horizontal="center")
        ws_one[f"D{r_num}"].font = font_reg
        ws_one[f"D{r_num}"].border = thin_border

        # Action risks (5 rows)
        if idx < len(action_risks):
            a_deal, a_symp, a_risk, a_act = action_risks[idx]
            ws_one[f"F{r_num}"] = a_deal
            ws_one[f"F{r_num}"].font = font_bold
            ws_one[f"F{r_num}"].border = thin_border
            ws_one[f"G{r_num}"] = a_symp
            ws_one[f"G{r_num}"].font = font_reg
            ws_one[f"G{r_num}"].border = thin_border
            ws_one[f"H{r_num}"] = a_risk
            ws_one[f"H{r_num}"].number_format = "#,##0 \"₽\""
            ws_one[f"H{r_num}"].font = font_card_red
            ws_one[f"H{r_num}"].border = thin_border
            ws_one[f"I{r_num}"] = a_act
            ws_one[f"I{r_num}"].font = font_reg
            ws_one[f"I{r_num}"].border = thin_border

        ws_one.row_dimensions[r_num].height = 22

    # Row 16: Chief Management Focus
    ws_one.merge_cells("A16:I16")
    ws_one["A16"] = "🎯 ГЛАВНЫЙ ФОКУС РУКОВОДСТВА ДО КОНЦА МЕСЯЦА: Перехват сделки D-104 (850 000 ₽) и перераспределение 3 лидов с перегруженного РОПа на КАМов."
    ws_one["A16"].font = Font(name="Segoe UI", size=10, bold=True, color="1E293B")
    ws_one["A16"].fill = fill_indigo_light
    ws_one["A16"].alignment = Alignment(horizontal="center", vertical="center")
    ws_one["A16"].border = card_border
    ws_one.row_dimensions[16].height = 28

    ws_one.column_dimensions["A"].width = 26
    ws_one.column_dimensions["B"].width = 12
    ws_one.column_dimensions["C"].width = 20
    ws_one.column_dimensions["D"].width = 18
    ws_one.column_dimensions["E"].width = 6
    ws_one.column_dimensions["F"].width = 26
    ws_one.column_dimensions["G"].width = 28
    ws_one.column_dimensions["H"].width = 20
    ws_one.column_dimensions["I"].width = 44

    # =========================================================================
    # SHEET 6: 🔐 152-ФЗ_Контур_Безопасности (Enterprise Compliance)
    # =========================================================================
    ws_sec = wb.create_sheet("🔐 152-ФЗ_Контур_Безопасности")
    setup_header_and_nav(ws_sec, "🔐 ЮРИДИЧЕСКИЙ КОНТУР БЕЗОПАСНОСТИ И ОБЕЗЛИЧИВАНИЯ ПДН (152-ФЗ РФ)", active_idx=8, max_col="F")

    ws_sec.merge_cells("A4:F4")
    ws_sec["A4"] = "🛡️ ГАРАНТИЯ СООТВЕТСТВИЯ ЗАКОНОДАТЕЛЬСТВУ И РЕГЛАМЕНТАМ БЕЗОПАСНОСТИ РФ"
    ws_sec["A4"].font = font_sec_hdr
    ws_sec.row_dimensions[4].height = 24

    sec_hdrs = [
        ("A5", "№"),
        ("B5", "Параметр безопасности"),
        ("C5", "Требование законодательства РФ"),
        ("D5", "Архитектурная реализация платформы RevOps"),
        ("E5", "Статус проверки"),
        ("F5", "Юридическое обоснование")
    ]
    for cell_id, text in sec_hdrs:
        ws_sec[cell_id] = text
        ws_sec[cell_id].font = font_tbl_hdr
        ws_sec[cell_id].fill = fill_table_header
        ws_sec[cell_id].alignment = Alignment(horizontal="center", vertical="center")
        ws_sec[cell_id].border = thin_border
    ws_sec.row_dimensions[5].height = 26

    sec_data = [
        (1, "Локализация баз данных", "Хранение ПДн граждан РФ строго на серверах в РФ", "Дата-центры в Москве (Yandex Cloud / Selectel), аттестация ФСТЭК УЗ 1-2", "✅ 100% Соответствие", "ст. 18.1 Федерального закона № 152-ФЗ"),
        (2, "Криптографическая токенизация", "Запрет передачи персональных данных во внешние контуры", "Модуль PIISanitizer: замена телефонов, ФИО и карт на токены [PHONE_TOKEN_X]", "✅ Автоматически", "ст. 6 Федерального закона № 152-ФЗ"),
        (3, "Срок хранения аудиозаписей", "Удаление данных после достижения цели обработки", "Автоматическое безвозвратное удаление исходных аудио через 30 дней", "✅ Настроено в cron", "ст. 5 Федерального закона № 152-ФЗ"),
        (4, "Согласие сотрудников ОП", "Уведомление персонала о контроле качества переговоров", "Готовый пакет документов: Приказ, Положение о контроле, Согласие", "✅ Комплект включен", "ст. 86 Трудового Кодекса РФ"),
        (5, "Юридический статус исполнителя", "Прозрачность расчетов с юридическими лицами", "Плательщик НПД (Самозанятый), чек приложения «Мой Налог», договор, акт", "✅ Безопасно для бухгалтерии", "ст. 15 Федерального закона № 422-ФЗ"),
        (6, "Гарантия конфиденциальности NDA", "Защита коммерческой тайны и клиентской базы компании", "Подписание двустороннего соглашения о неразглашении конфиденциальных данных", "✅ Подписание в ЭДО", "Федеральный закон № 98-ФЗ «О коммерческой тайне»")
    ]

    for idx, s in enumerate(sec_data, start=6):
        ws_sec[f"A{idx}"] = s[0]
        ws_sec[f"A{idx}"].alignment = Alignment(horizontal="center", vertical="center")
        ws_sec[f"B{idx}"] = s[1]
        ws_sec[f"B{idx}"].font = font_bold
        ws_sec[f"C{idx}"] = s[2]
        ws_sec[f"D{idx}"] = s[3]
        ws_sec[f"E{idx}"] = s[4]
        ws_sec[f"E{idx}"].alignment = Alignment(horizontal="center", vertical="center")
        ws_sec[f"E{idx}"].font = font_card_green
        ws_sec[f"E{idx}"].fill = fill_success
        ws_sec[f"F{idx}"] = s[5]
        ws_sec[f"F{idx}"].font = font_reg

        for col_l in ["A", "B", "C", "D", "E", "F"]:
            ws_sec[f"{col_l}{idx}"].border = thin_border
            if col_l not in ["B", "E"]:
                ws_sec[f"{col_l}{idx}"].font = font_reg
        ws_sec.row_dimensions[idx].height = 28

    ws_sec.column_dimensions["A"].width = 6
    ws_sec.column_dimensions["B"].width = 30
    ws_sec.column_dimensions["C"].width = 44
    ws_sec.column_dimensions["D"].width = 54
    ws_sec.column_dimensions["E"].width = 24
    ws_sec.column_dimensions["F"].width = 36

    # Set active sheet to ⚡ Экспресс_Калькулятор_3_Цифры
    wb.active = ws_calc

    # Save to target
    target_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(target_path)
    print(f"Successfully generated: {target_path}")


if __name__ == "__main__":
    desktop_file = Path(r"C:\Users\strel\Desktop\RevOps Platform\Презентация\Презентационная_Таблица_RevOps.xlsx")
    repo_file1 = Path(r"presentation\RevOps_Platform_Demo_Sample.xlsx")
    repo_file2 = Path(r"docs\RevOps_Platform_Demo_Sample.xlsx")

    build_presentation_workbook(desktop_file)
    build_presentation_workbook(repo_file1)
    build_presentation_workbook(repo_file2)
