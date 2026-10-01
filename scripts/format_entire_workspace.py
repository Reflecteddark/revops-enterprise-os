"""
RevOps Enterprise OS V17.6 — Master Formatter for Founder's Workspace Planner.
Полная калибровка размеров: ширина всех колонок, высота всех строк, перенос текста (wrap_text),
выравнивание и фиксация заголовков (freeze rows) как в локальном файле Excel,
так и в облачной Google Таблице.
"""

from __future__ import annotations

import datetime
import os
import shutil
import sys
from pathlib import Path

if sys.stdout.encoding != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import gspread

LOCAL_PLANNER = Path("presentation/RevOps_Founder_Workspace_Planner.xlsx")
DOCS_PLANNER = Path("docs/RevOps_Founder_Workspace_Planner.xlsx")
DESKTOP_PLANNER = Path(os.path.expanduser(r"~\Desktop\RevOps_Founder_Workspace_Planner.xlsx"))

SPREADSHEET_ID = "1AbaQGHhzar7Qsb9U9TByN57rAmOn36zw4HOD2XVrmTc"
CREDENTIALS_FILE = Path("service_account.json")


def format_local_excel():
    """Калибрует размеры строк, колонок и перенос текста во всех листах Excel."""
    print("1. Форматирование локального файла Excel...")
    wb = openpyxl.load_workbook(LOCAL_PLANNER)

    # ── Лист 1: 🎯 Спринты_и_Задачи ──
    if "🎯 Спринты_и_Задачи" in wb.sheetnames:
        ws = wb["🎯 Спринты_и_Задачи"]
        ws.row_dimensions[1].height = 36
        ws.row_dimensions[2].height = 12
        ws.row_dimensions[3].height = 28

        col_widths_ws1 = {
            "A": 9,   # ID
            "B": 18,  # Стрим
            "C": 46,  # Задача
            "D": 13,  # Приоритет
            "E": 14,  # Дедлайн
            "F": 18,  # Статус
            "G": 36,  # Артефакт
            "H": 52,  # Заметки
        }
        for col, width in col_widths_ws1.items():
            ws.column_dimensions[col].width = width

        for r in range(4, ws.max_row + 1):
            ws.row_dimensions[r].height = 34
            # Включаем wrap_text на длинных колонках
            for col_idx in [3, 7, 8]:
                cell = ws.cell(r, col_idx)
                cell.alignment = Alignment(vertical="center", wrap_text=True)

    # ── Лист 2: 💡 База_Знаний_и_УТП ──
    if "💡 База_Знаний_и_УТП" in wb.sheetnames:
        ws = wb["💡 База_Знаний_и_УТП"]
        ws.row_dimensions[1].height = 36
        ws.row_dimensions[2].height = 12
        ws.row_dimensions[3].height = 26
        ws.row_dimensions[4].height = 65  # УТП карточка
        ws.row_dimensions[8].height = 26
        ws.row_dimensions[9].height = 26

        # Боли
        for r in [10, 11, 12]:
            ws.row_dimensions[r].height = 42

        ws.row_dimensions[14].height = 26
        # Матрица
        for r in [15, 16, 17]:
            ws.row_dimensions[r].height = 42

        col_widths_ws2 = {
            "A": 12, "B": 32, "C": 26, "D": 26, "E": 26, "F": 26, "G": 26
        }
        for col, width in col_widths_ws2.items():
            ws.column_dimensions[col].width = width

        for r in range(4, 18):
            for c in range(1, 8):
                cell = ws.cell(r, c)
                if cell.value:
                    align = cell.alignment
                    ws.cell(r, c).alignment = Alignment(
                        horizontal=align.horizontal or "left",
                        vertical="center",
                        wrap_text=True
                    )

    # ── Лист 3: 📑 Инсайты_2_Исследований ──
    if "📑 Инсайты_2_Исследований" in wb.sheetnames:
        ws = wb["📑 Инсайты_2_Исследований"]
        ws.row_dimensions[1].height = 36
        ws.row_dimensions[2].height = 12
        ws.row_dimensions[3].height = 28

        col_widths_ws3 = {
            "A": 8,   # №
            "B": 26,  # Тема
            "C": 42,  # Тезис
            "D": 42,  # Боль
            "E": 42,  # Решение
            "F": 46,  # Правило
        }
        for col, width in col_widths_ws3.items():
            ws.column_dimensions[col].width = width

        for r in range(4, ws.max_row + 1):
            ws.row_dimensions[r].height = 68
            for c in range(1, 7):
                cell = ws.cell(r, c)
                align = cell.alignment
                cell.alignment = Alignment(
                    horizontal=align.horizontal or "left",
                    vertical="center",
                    wrap_text=True
                )

    # ── Лист 4: 🤝 Пайплайн_Интеграторов ──
    if "🤝 Пайплайн_Интеграторов" in wb.sheetnames:
        ws = wb["🤝 Пайплайн_Интеграторов"]
        ws.row_dimensions[1].height = 36
        ws.row_dimensions[2].height = 12
        ws.row_dimensions[3].height = 28

        col_widths_ws4 = {
            "A": 8,   # №
            "B": 30,  # Компания
            "C": 16,  # Стек
            "D": 18,  # Город
            "E": 24,  # Контакт
            "F": 22,  # Статус
            "G": 28,  # RevShare
            "H": 22,  # Потенциал
            "I": 36,  # Следующий шаг
        }
        for col, width in col_widths_ws4.items():
            ws.column_dimensions[col].width = width

        for r in range(4, ws.max_row + 1):
            ws.row_dimensions[r].height = 25
            cell_i = ws.cell(r, 9)
            cell_i.alignment = Alignment(vertical="center", wrap_text=True)

    # ── Лист 5: 🎙️ Клиентские_Пилоты ──
    if "🎙️ Клиентские_Пилоты" in wb.sheetnames:
        ws = wb["🎙️ Клиентские_Пилоты"]
        ws.row_dimensions[1].height = 36
        ws.row_dimensions[2].height = 12
        ws.row_dimensions[3].height = 28

        col_widths_ws5 = {
            "A": 8,   # №
            "B": 28,  # Клиент
            "C": 22,  # Сфера
            "D": 16,  # Менеджеров
            "E": 18,  # Чек
            "F": 16,  # CRM
            "G": 22,  # Статус
            "H": 16,  # Пилот
            "I": 20,  # Тариф
            "J": 38,  # Заметки
        }
        for col, width in col_widths_ws5.items():
            ws.column_dimensions[col].width = width

        for r in range(4, ws.max_row + 1):
            ws.row_dimensions[r].height = 25
            cell_j = ws.cell(r, 10)
            cell_j.alignment = Alignment(vertical="center", wrap_text=True)

    # ── Лист 6: 💬 Быстрые_Скрипты ──
    if "💬 Быстрые_Скрипты" in wb.sheetnames:
        ws = wb["💬 Быстрые_Скрипты"]
        ws.row_dimensions[1].height = 36
        for c_letter in ["A", "B", "C", "D", "E", "F"]:
            ws.column_dimensions[c_letter].width = 22

        # Настраиваем ячейки текстов скриптов
        script_rows = [(3, 4), (10, 11), (17, 18)]
        for h_row, body_row in script_rows:
            ws.row_dimensions[h_row].height = 26
            ws.row_dimensions[body_row].height = 100
            cell = ws.cell(body_row, 1)
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    wb.save(LOCAL_PLANNER)
    shutil.copyfile(LOCAL_PLANNER, DOCS_PLANNER)
    if DESKTOP_PLANNER.parent.exists():
        shutil.copyfile(LOCAL_PLANNER, DESKTOP_PLANNER)

    print("✓ Локальный Excel файл откалиброван и обновлен на Рабочем столе!")


def format_google_sheets():
    """Применяет автоматическое выравнивание и перенос текста в облачной Google Таблице."""
    if not CREDENTIALS_FILE.exists():
        print("service_account.json не найден, пропуск Google Sheets.")
        return

    print("2. Калибровка облачной Google Таблицы через API...")
    gc = gspread.service_account(filename=str(CREDENTIALS_FILE))
    sh = gc.open_by_key(SPREADSHEET_ID)

    requests = []

    # Конфигурация пиксельных ширин для каждого листа
    sheet_configs = {
        "🎯 Спринты_и_Задачи": {
            "cols": [75, 140, 360, 100, 110, 140, 280, 420],
            "freeze_rows": 3,
        },
        "💡 База_Знаний_и_УТП": {
            "cols": [90, 240, 200, 200, 200, 200, 200],
            "freeze_rows": 1,
        },
        "📑 Инсайты_2_Исследований": {
            "cols": [65, 200, 320, 320, 320, 360],
            "freeze_rows": 3,
        },
        "🤝 Пайплайн_Интеграторов": {
            "cols": [65, 240, 120, 140, 180, 170, 220, 170, 280],
            "freeze_rows": 3,
        },
        "🎙️ Клиентские_Пилоты": {
            "cols": [65, 220, 170, 120, 140, 120, 170, 120, 150, 300],
            "freeze_rows": 3,
        },
        "💬 Быстрые_Скрипты": {
            "cols": [160, 160, 160, 160, 160, 160],
            "freeze_rows": 1,
        },
    }

    for ws in sh.worksheets():
        sheet_id = ws.id
        title = ws.title

        # Включаем автоматический перенос текста (WRAP) и вертикальное выравнивание (MIDDLE)
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": 40,
                    "startColumnIndex": 0,
                    "endColumnIndex": 12,
                },
                "cell": {
                    "userEnteredFormat": {
                        "wrapStrategy": "WRAP",
                        "verticalAlignment": "MIDDLE",
                    }
                },
                "fields": "userEnteredFormat(wrapStrategy,verticalAlignment)",
            }
        })

        # Применяем специфические настройки колонок
        if title in sheet_configs:
            cfg = sheet_configs[title]
            # Заморозка шапки
            if cfg.get("freeze_rows"):
                requests.append({
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": sheet_id,
                            "gridProperties": {
                                "frozenRowCount": cfg["freeze_rows"]
                            }
                        },
                        "fields": "gridProperties.frozenRowCount"
                    }
                })

            # Размеры колонок в пикселях
            for col_idx, width_px in enumerate(cfg["cols"]):
                requests.append({
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": col_idx,
                            "endIndex": col_idx + 1,
                        },
                        "properties": {
                            "pixelSize": width_px,
                        },
                        "fields": "pixelSize",
                    }
                })

    if requests:
        body = {"requests": requests}
        sh.batch_update(body)
        print("✓ Google Таблица успешно отформатирована: ширины колонок, перенос текста и фиксация шапок применены!")


if __name__ == "__main__":
    format_local_excel()
    format_google_sheets()
