"""
RevOps Enterprise OS V17.6 — Live Google Sheets Synchronizer.
Синхронизация рабочей книги фаундера напрямую с облачной Google Таблицей:
https://docs.google.com/spreadsheets/d/1AbaQGHhzar7Qsb9U9TByN57rAmOn36zw4HOD2XVrmTc/
Сервисный аккаунт: sheets-agent@antigravity-sheets-509519.iam.gserviceaccount.com
"""

from __future__ import annotations

import datetime
import os
import sys
from pathlib import Path
from typing import List, Optional

if sys.stdout.encoding != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

import gspread

SPREADSHEET_ID = "1AbaQGHhzar7Qsb9U9TByN57rAmOn36zw4HOD2XVrmTc"
CREDENTIALS_FILE = Path("service_account.json")


def get_client() -> gspread.Client:
    """Возвращает аутентифицированный gspread клиент."""
    if not CREDENTIALS_FILE.exists():
        raise FileNotFoundError(f"Файл ключа {CREDENTIALS_FILE} не найден в корне проекта.")
    return gspread.service_account(filename=str(CREDENTIALS_FILE))


def get_spreadsheet() -> gspread.Spreadsheet:
    """Открывает мастер-таблицу фаундера."""
    client = get_client()
    return client.open_by_key(SPREADSHEET_ID)


def update_task_in_google_sheets(task_id: str, new_status: str, notes: Optional[str] = None) -> bool:
    """
    Обновляет статус и заметки задачи в Google Таблице по ее ID (например, T-08).
    """
    try:
        sh = get_spreadsheet()
        ws = sh.worksheet("🎯 Спринты_и_Задачи")
        records = ws.get_all_values()

        found_row = None
        for i, row in enumerate(records):
            if row and row[0].strip().upper() == task_id.strip().upper():
                found_row = i + 1  # 1-indexed
                break

        if not found_row:
            print(f"Задача {task_id} не найдена в Google Таблице.")
            return False

        # Колонка F (6) — Статус
        ws.update_cell(found_row, 6, new_status)
        if notes:
            # Колонка H (8) — Заметки
            ws.update_cell(found_row, 8, notes)

        # Обновляем таймстемп синхронизации в шапке
        now_str = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
        ws.update(range_name="H1", values=[[f"ИИ Antigravity • {now_str}"]])

        print(f"[Google Sheets] Задача {task_id} успешно обновлена: {new_status}")
        return True

    except Exception as e:
        print(f"[Google Sheets Error] Ошибка обновления задачи {task_id}: {e}")
        return False


def add_task_to_google_sheets(
    stream: str,
    title: str,
    priority: str = "P1",
    deadline: str = "",
    status: str = "⏳ В РАБОТЕ",
    artifact: str = "",
    notes: str = "",
) -> Optional[str]:
    """
    Добавляет новую задачу в список спринтов Google Таблицы.
    """
    try:
        sh = get_spreadsheet()
        ws = sh.worksheet("🎯 Спринты_и_Задачи")
        records = ws.get_all_values()

        max_id = 0
        for r in records[3:]:
            if r and r[0].startswith("T-"):
                try:
                    num = int(r[0].split("-")[1])
                    max_id = max(max_id, num)
                except ValueError:
                    pass

        new_id = f"T-{max_id + 1:02d}"
        if not deadline:
            deadline = datetime.datetime.now().strftime("%d.%m.%Y")

        new_row = [new_id, stream, title, priority, deadline, status, artifact, notes]
        ws.append_row(new_row)

        now_str = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
        ws.update(range_name="H1", values=[[f"ИИ Antigravity • {now_str}"]])

        print(f"[Google Sheets] Добавлена новая задача {new_id}: {title} [{status}]")
        return new_id

    except Exception as e:
        print(f"[Google Sheets Error] Ошибка добавления задачи: {e}")
        return None


def add_partner_to_google_sheets(
    company: str,
    stack: str = "amoCRM",
    city: str = "Москва",
    contact: str = "",
    status: str = "В плане на контакт",
    revshare: str = "25%",
    potential: str = "10 активных",
    next_step: str = "Отправить оффер",
) -> bool:
    """
    Добавляет партнера-интегратора в воронку в Google Таблице.
    """
    try:
        sh = get_spreadsheet()
        ws = sh.worksheet("🤝 Пайплайн_Интеграторов")
        records = ws.get_all_values()

        # Ищем первую строку, где колонка B пустая (начиная с 4-й строки)
        empty_row = None
        for i in range(3, len(records)):
            if len(records[i]) < 2 or not records[i][1].strip():
                empty_row = i + 1
                break

        if empty_row:
            idx = empty_row - 3
            row_data = [idx, company, stack, city, contact, status, revshare, potential, next_step]
            ws.update(range_name=f"A{empty_row}:I{empty_row}", values=[row_data])
        else:
            idx = len(records) - 2
            row_data = [idx, company, stack, city, contact, status, revshare, potential, next_step]
            ws.append_row(row_data)

        print(f"[Google Sheets] Партнер {company} записан в пайплайн.")
        return True

    except Exception as e:
        print(f"[Google Sheets Error] Ошибка добавления партнера: {e}")
        return False


def add_client_to_google_sheets(
    company: str,
    niche: str,
    managers: int,
    check: int,
    crm: str,
    status: str = "Лид на аудит",
    notes: str = "",
) -> bool:
    """
    Добавляет клиента в воронку пилотов в Google Таблице.
    """
    try:
        sh = get_spreadsheet()
        ws = sh.worksheet("🎙️ Клиентские_Пилоты")
        records = ws.get_all_values()

        empty_row = None
        for i in range(3, len(records)):
            if len(records[i]) < 2 or not records[i][1].strip():
                empty_row = i + 1
                break

        tier = "Старт (49к)" if managers <= 5 else ("Бизнес (79к)" if managers <= 10 else "Enterprise (120к)")
        check_str = f"{check:,}".replace(",", " ") + " ₽"

        if empty_row:
            idx = empty_row - 3
            row_data = [idx, company, niche, managers, check_str, crm, status, "29 000 ₽", tier, notes]
            ws.update(range_name=f"A{empty_row}:J{empty_row}", values=[row_data])
        else:
            idx = len(records) - 2
            row_data = [idx, company, niche, managers, check_str, crm, status, "29 000 ₽", tier, notes]
            ws.append_row(row_data)

        print(f"[Google Sheets] Клиент {company} записан в воронку пилотов.")
        return True

    except Exception as e:
        print(f"[Google Sheets Error] Ошибка добавления клиента: {e}")
        return False


if __name__ == "__main__":
    if sys.stdout.encoding != "utf-8":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass

    print("Проверка подключения к Google Таблице фаундера...")
    try:
        spreadsheet = get_spreadsheet()
        print(f"✓ Успешное подключение к: '{spreadsheet.title}'")
        for ws in spreadsheet.worksheets():
            print(f"  - Вкладка: {ws.title}")
    except Exception as err:
        print(f"Ошибка: {err}")
