#!/usr/bin/env python3
"""
Patch and upgrade RevOps Platform workbooks to Enterprise V18.1:
1. Fix BOT_TOKEN in ⚙️ Настройки (mask with secure ScriptProperties notice, update Apps Script).
2. Standardize Navigation Bar across all 17 visual showcase dashboards.
3. Eliminate all 114 merged cells on 🎙️ ИИ_Аудит (replace with proper column widths & wrap_text).
4. Upgrade Empty States from '0.0' and '—' to smart onboarding hints.
5. Upgrade 305 hardcoded $1000 formula ranges to $50000 enterprise scale.
"""

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
import os
import sys

def patch_workbook(filepath):
    print(f"\n==================================================")
    print(f"PATCHING: {filepath}")
    print(f"==================================================")
    if not os.path.exists(filepath):
        print(f"  [!] File not found: {filepath}")
        return False

    wb = openpyxl.load_workbook(filepath, data_only=False)
    sheetnames = wb.sheetnames
    print(f"  Total sheets: {len(sheetnames)}")

    # ----------------------------------------------------
    # FIX 1: BOT_TOKEN in ⚙️ Настройки
    # ----------------------------------------------------
    settings_sheets = [s for s in sheetnames if 'Настройки' in s]
    if settings_sheets:
        ws_set = wb[settings_sheets[0]]
        # Find BOT_TOKEN row
        found_token_row = None
        for r in range(1, min(ws_set.max_row + 1, 100)):
            val = str(ws_set.cell(row=r, column=1).value or '')
            if 'BOT_TOKEN' in val:
                found_token_row = r
                break
        if found_token_row:
            old_val = ws_set.cell(row=found_token_row, column=2).value
            ws_set.cell(row=found_token_row, column=2).value = (
                "🔒 Хранится в защищенном Script Properties (настройка: меню RevOps -> 🔐 Telegram Bot)"
            )
            ws_set.cell(row=found_token_row, column=2).font = Font(name="Arial", size=10, bold=True, color="10B981")
            print(f"  [FIX 1] ⚙️ Настройки row {found_token_row}: BOT_TOKEN masked securely (was: '{old_val}').")

    # Update Code_Archive if present
    code_sheets = [s for s in sheetnames if 'Code_Archive' in s]
    if code_sheets:
        ws_code = wb[code_sheets[0]]
        for r in range(1, ws_code.max_row + 1):
            cell_val = str(ws_code.cell(row=r, column=1).value or '')
            if "const botToken = String(settingsSheet.getRange('B52').getValue()).trim();" in cell_val:
                ws_code.cell(row=r, column=1).value = (
                    "  // Защищенное получение токена: сначала ScriptProperties, фоллбэк на ячейку только при ручной отладке\n"
                    "  let botToken = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');\n"
                    "  if (!botToken) {\n"
                    "    const rawVal = String(settingsSheet.getRange('B52').getValue()).trim();\n"
                    "    if (rawVal && rawVal.indexOf('placeholder') === -1 && rawVal.indexOf('🔒') === -1) botToken = rawVal;\n"
                    "  }"
                )
                print(f"  [FIX 1] 🔐 Code_Archive row {r}: patched botToken lookup to use PropertiesService.")
                break

    # ----------------------------------------------------
    # FIX 2 & FIX 3: 🎙️ ИИ_Аудит - Unmerge all 114 cells & format properly
    # ----------------------------------------------------
    audit_sheets = [s for s in sheetnames if 'ИИ_Аудит' in s]
    if audit_sheets:
        ws_audit = wb[audit_sheets[0]]
        merged_ranges = list(ws_audit.merged_cells.ranges)
        count_merged = len(merged_ranges)
        print(f"  [FIX 3] 🎙️ ИИ_Аудит: Found {count_merged} merged ranges. Unmerging data cells...")
        
        # We unmerge all ranges to guarantee flat tabular filtering
        for mrange in merged_ranges:
            ws_audit.unmerge_cells(str(mrange))
        print(f"  [FIX 3] 🎙️ ИИ_Аудит: All {count_merged} merged ranges unmerged successfully.")

        # Set optimal column widths & wrap text for clean readability
        col_widths = {
            1: 12,  # Call ID / №
            2: 24,  # Сделка / Контрагент
            3: 16,  # Менеджер
            4: 16,  # Сумма в риске
            5: 28,  # Дефект речи / Параметр
            6: 45,  # Цитата Whisper / Суть критерия
            7: 45,  # Решение РОПа / Маркеры
            8: 14,  # Оценка
            9: 14,  # Статус
            10: 16, # Длительность
            11: 16, # Дата
            12: 18, # CRM Ссылка
        }
        for col_idx, width in col_widths.items():
            col_letter = get_column_letter(col_idx)
            ws_audit.column_dimensions[col_letter].width = width

        # Apply wrap text to text columns in table area
        for r in range(1, min(ws_audit.max_row + 1, 120)):
            for c in range(1, 15):
                cell = ws_audit.cell(row=r, column=c)
                if cell.value:
                    align = cell.alignment
                    if align:
                        cell.alignment = Alignment(
                            horizontal=align.horizontal or 'left',
                            vertical='center',
                            wrap_text=True
                        )
                    else:
                        cell.alignment = Alignment(vertical='center', wrap_text=True)

    # ----------------------------------------------------
    # FIX 2: Standardize Navigation Bar across all 17 Visual Dashboards
    # ----------------------------------------------------
    visual_dashboards = [
        ('📄 Executive_OnePager', '📄 1. One-Pager'),
        ('⚡ Пульс_Компании', '⚡ 2. Пульс компании'),
        ('📋 Пульт_РОПа_15_Минут', '📋 3. Пульт РОПа'),
        ('💸 Диагностика_Утечек_ОП', '💸 4. 7 Грехов ОП'),
        ('🎯 Action_Center', '🎯 5. Action Center'),
        ('🎙️ ИИ_Аудит', '🎙️ 6. ИИ-Аудит'),
        ('⚡ Экспресс_Калькулятор_3_Цифры', '⚡ 7. Экспресс 3 цифры'),
        ('🧪 QA_Suite', '🧪 8. QA Suite'),
        ('⚙️ Настройки', '⚙️ 9. Настройки'),
        ('🌐 Мультиканальная_Атрибуция', '🌐 10. Атрибуция'),
        ('💳 Финансы_и_AI_Дожим', '💳 11. Финансы'),
        ('👥 Мотивация_ОП', '👥 12. Мотивация'),
        ('🔮 Симулятор_Роста', '🔮 13. Симулятор'),
        ('🗄️ DWH_и_Безопасность', '🗄️ 14. DWH / RBAC'),
        ('🌐 Сквозная_RevOps_Аналитика', '🌐 15. Сквозная RevOps'),
        ('🌐 Маркетинг_и_Трафик', '🌐 16. Маркетинг'),
        ('🎯 Воронка_и_SLA', '🎯 17. Воронка & SLA'),
        ('📊 Юнит_Экономика', '📊 18. Юнит-Экономика'),
        ('🚨 Радар_Алертов', '🚨 19. Радар Алертов'),
        ('📈 Когорты_LTV', '📈 20. Когорты LTV'),
        ('👥 Ресурсный_План', '👥 21. Ресурсный План'),
    ]

    # Navbar core buttons (top 9 primary jump links)
    nav_buttons = [
        ("📄 Executive_OnePager", "📄 1. One-Pager"),
        ("⚡ Пульс_Компании", "⚡ 2. Пульс компании"),
        ("📋 Пульт_РОПа_15_Минут", "📋 3. Пульт РОПа"),
        ("💸 Диагностика_Утечек_ОП", "💸 4. 7 Грехов ОП"),
        ("🎯 Action_Center", "🎯 5. Action Center"),
        ("🎙️ ИИ_Аудит", "🎙️ 6. ИИ-Аудит"),
        ("⚡ Экспресс_Калькулятор_3_Цифры", "⚡ 7. Экспресс 3 цифры"),
        ("🧪 QA_Suite", "🧪 8. QA Suite"),
        ("⚙️ Настройки", "⚙️ 9. Настройки"),
    ]

    nav_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    nav_font = Font(name="Arial", size=9, bold=True, color="38BDF8")
    nav_border = Border(
        bottom=Side(style='thin', color="334155"),
        top=Side(style='thin', color="334155"),
        left=Side(style='thin', color="1E293B"),
        right=Side(style='thin', color="1E293B")
    )
    nav_align = Alignment(horizontal="center", vertical="center", wrap_text=False)

    nav_added_count = 0
    for sheet_tuple in visual_dashboards:
        sh_name = sheet_tuple[0]
        matching = [s for s in sheetnames if sh_name in s or s in sh_name]
        if not matching:
            continue
        ws = wb[matching[0]]
        
        # Unmerge any cells touching row 2 on this sheet first to avoid MergedCell read-only error
        row2_merged = [str(mr) for mr in list(ws.merged_cells.ranges) if mr.min_row <= 2 <= mr.max_row]
        for rstr in row2_merged:
            ws.unmerge_cells(rstr)

        # We place the navigation bar into row 2
        ws.row_dimensions[2].height = 24
        for idx, (target_sheet, label) in enumerate(nav_buttons, start=1):
            # Check if target sheet actually exists in this workbook
            target_match = [s for s in sheetnames if target_sheet in s or s in target_sheet]
            actual_target = target_match[0] if target_match else target_sheet
            
            # Excel internal hyperlink syntax: =HYPERLINK("#'SheetName'!A1", "Label")
            formula = f"=HYPERLINK(\"#'{actual_target}'!A1\", \"{label}\")"
            cell = ws.cell(row=2, column=idx)
            cell.value = formula
            cell.fill = nav_fill
            cell.font = nav_font
            cell.alignment = nav_align
            cell.border = nav_border
        nav_added_count += 1
        print(f"  [FIX 2] Standardized internal Navbar on: '{matching[0]}'")

    print(f"  [FIX 2] Total visual sheets updated with Navbar: {nav_added_count}")

    # ----------------------------------------------------
    # FIX 4: Upgrade Empty States (replace '0.0' and '—' with helpful onboarding prompts)
    # ----------------------------------------------------
    empty_states_fixed = 0
    for name in sheetnames:
        if name.startswith('raw_') or name.startswith('calc_') or name in ('changelog', 'Code_Archive'):
            continue
        ws = wb[name]
        for r in range(1, min(ws.max_row + 1, 120)):
            for c in range(1, min(ws.max_column + 1, 30)):
                cell = ws.cell(row=r, column=c)
                if type(cell).__name__ == 'MergedCell':
                    continue
                val = cell.value
                if val is not None:
                    sval = str(val).strip()
                    # Handle raw string empty states
                    if sval in ('—', '--', '- -', '---'):
                        # If in AI Audit
                        if 'ИИ_Аудит' in name and r == 31:
                            cell.value = "⚡ Ожидает 3 звонков"
                            cell.font = Font(name="Arial", size=9, italic=True, color="94A3B8")
                            empty_states_fixed += 1
                    elif sval in ('0.0', '0.00', '0.0%') and cell.data_type == 's':
                        cell.value = "0.0 (Ожидание данных)"
                        cell.font = Font(name="Arial", size=9, italic=True, color="94A3B8")
                        empty_states_fixed += 1

    print(f"  [FIX 4] Empty states upgraded: {empty_states_fixed} placeholder cells improved.")

    # ----------------------------------------------------
    # FIX 5: Upgrade 305 Hardcoded Ranges ($1000 / 1000) to 50000 Enterprise Scale
    # ----------------------------------------------------
    formulas_expanded = 0
    # Regular expression to match range references ending in 1000, e.g.:
    # $A$2:$A$1000, $D$2:$D1000, raw_deals!$A$2:$A1000, raw_calls!$H$2:$H1000
    pattern_1000 = re.compile(r'(\$?[A-Z]+)(\$?[0-9]+):(\$?[A-Z]+)(\$?)1000\b')

    for name in sheetnames:
        ws = wb[name]
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(row=r, column=c)
                if type(cell).__name__ == 'MergedCell':
                    continue
                val = str(cell.value or '')
                if val.startswith('=') and '1000' in val:
                    # Check if it has range syntax
                    new_val = pattern_1000.sub(r'\g<1>\g<2>:\g<3>\g<4>50000', val)
                    if new_val != val:
                        cell.value = new_val
                        formulas_expanded += 1

    print(f"  [FIX 5] Formulas expanded from 1,000 to 50,000 rows: {formulas_expanded} formulas updated.")

    # Add Watchdog Alert cell on QA Suite / Пульт РОПа
    qa_sheets = [s for s in sheetnames if 'QA_Suite' in s]
    if qa_sheets:
        ws_qa = wb[qa_sheets[0]]
        watchdog_row = ws_qa.max_row + 2
        ws_qa.cell(row=watchdog_row, column=1).value = "WATCHDOG_ROW_CAPACITY"
        ws_qa.cell(row=watchdog_row, column=2).value = "Контроль емкости строк (50 000)"
        ws_qa.cell(row=watchdog_row, column=4).value = (
            '=IF(OR(COUNTA(raw_deals!$A$2:$A$50000)>48000, COUNTA(raw_calls!$A$2:$A$50000)>48000), '
            '"🚨 ВНИМАНИЕ: Заполнено 96% емкости (48 000+). Создайте архивный срез.", "🟢 Лимит в норме (до 50 000 строк)")'
        )
        ws_qa.cell(row=watchdog_row, column=5).value = True
        ws_qa.cell(row=watchdog_row, column=6).value = f"=D{watchdog_row}"
        print(f"  [FIX 5] Added Watchdog Capacity Alert to 🧪 QA_Suite row {watchdog_row}.")

    # Save upgraded workbook
    wb.save(filepath)
    print(f"  [SUCCESS] Saved upgraded workbook to: {filepath}\n")
    return True

if __name__ == '__main__':
    targets = [
        'docs/RevOps_Platform_V18.0_Clean_Client_Starter.xlsx',
        'docs/RevOps_Platform_V18.0_Showcase_Master.xlsx',
        'web/RevOps_Platform_Demo_Sample.xlsx',
        'C:/Users/strel/.gemini/antigravity/scratch/RevOps Platform V17.6 (RBAC Production Suite).xlsx'
    ]
    
    # Also search for any client-specific files like "Автожидкости" in scratch
    scratch_dir = 'C:/Users/strel/.gemini/antigravity/scratch'
    for f in os.listdir(scratch_dir):
        if 'жидкост' in f.lower() or 'авто' in f.lower():
            full_p = os.path.join(scratch_dir, f)
            if full_p.endswith('.xlsx'):
                targets.append(full_p)

    print(f"Target files to patch: {len(targets)}")
    for t in targets:
        try:
            patch_workbook(t)
        except Exception as e:
            print(f"  [ERROR] Failed to patch {t}: {e}")
