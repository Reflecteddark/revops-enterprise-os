"""
Модуль аудита генерации отчетов RevOps Enterprise OS V17.6.
Фиксирует факты доступа и экспорта финансовых данных в соответствии с требованиями 152-ФЗ РФ и корпоративного комплаенса.

Использует исключительно стандартную библиотеку Python (csv, os, time, getpass, datetime, shutil).
"""

import os
import sys
import csv
import time
import getpass
from datetime import datetime
import shutil

AUDIT_LOG_FILENAME = "audit_log.csv"
MAX_LOG_ENTRIES = 10000

FIELDNAMES = [
    "timestamp",
    "user_login",
    "role",
    "report_type",
    "duration_sec",
    "status",
    "error_message",
    "pdf_size_kb"
]


def _get_current_user() -> str:
    """Определяет логин текущего пользователя в кроссплатформенном режиме."""
    try:
        return os.environ.get("USERNAME") or os.environ.get("USER") or getpass.getuser() or "unknown"
    except Exception:
        return "unknown"


def _rotate_if_needed(csv_path: str) -> None:
    """
    Архивирует audit_log.csv в audit_log_YYYY-MM.csv при превышении 10 000 строк.
    """
    if not os.path.exists(csv_path):
        return

    try:
        with open(csv_path, "r", encoding="utf-8-sig", errors="ignore") as f:
            line_count = sum(1 for _ in f)

        # 1 строка заголовка + 10 000 записей
        if line_count > MAX_LOG_ENTRIES:
            dir_name = os.path.dirname(csv_path)
            now = datetime.now()
            archive_name = f"audit_log_{now.strftime('%Y-%m')}.csv"
            archive_path = os.path.join(dir_name, archive_name)

            # Если архивный файл с таким именем уже есть, добавляем точный таймстемп
            if os.path.exists(archive_path):
                archive_name = f"audit_log_{now.strftime('%Y-%m-%d_%H%M%S')}.csv"
                archive_path = os.path.join(dir_name, archive_name)

            shutil.move(csv_path, archive_path)
            print(f"[AUDIT_LOG] Журнал превысил {MAX_LOG_ENTRIES} записей и архивирован в {archive_name}")
    except Exception as e:
        print(f"[AUDIT_LOG_WARNING] Не удалось выполнить ротацию журнала: {e}", file=sys.stderr)


def log_generation(
    status: str,
    duration_sec: float,
    role: str = "full",
    report_type: str = "Role_Report",
    error_message: str = "",
    pdf_size_kb: float = 0.0,
    log_dir: str | None = None
) -> bool:
    """
    Записывает событие генерации отчета в audit_log.csv.

    Параметры:
        status (str): "success" или "error"
        duration_sec (float): Время выполнения генерации в секундах
        role (str): Роль пользователя / срез отчета ("full", "ceo", "rop", "cfo" и т.д.)
        report_type (str): Тип отчета ("Executive_Summary" или "Role_Report" / "Full_Report")
        error_message (str): Текст ошибки (пусто при success)
        pdf_size_kb (float): Размер сформированного PDF в килобайтах
        log_dir (str, optional): Папка для размещения audit_log.csv. По умолчанию папка модуля.

    Возвращает:
        bool: True, если запись прошла успешно, False при ошибке (исключение не выбрасывается).
    """
    try:
        base_dir = log_dir if log_dir else os.path.dirname(os.path.abspath(__file__))
        csv_path = os.path.join(base_dir, AUDIT_LOG_FILENAME)

        # Проверка и ротация перед записью
        _rotate_if_needed(csv_path)

        file_exists = os.path.exists(csv_path)

        timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        user_login = _get_current_user()

        record = {
            "timestamp": timestamp_str,
            "user_login": user_login,
            "role": str(role).strip().lower(),
            "report_type": str(report_type).strip(),
            "duration_sec": round(float(duration_sec), 2),
            "status": "success" if str(status).strip().lower() == "success" else "error",
            "error_message": str(error_message).strip() if status != "success" else "",
            "pdf_size_kb": round(float(pdf_size_kb), 1)
        }

        # Открываем с utf-8-sig (BOM), чтобы Microsoft Excel открывал кириллицу корректно
        with open(csv_path, mode="a", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=FIELDNAMES, delimiter=";")
            if not file_exists:
                writer.writeheader()
            writer.writerow(record)

        return True

    except Exception as e:
        # Не блокируем бизнес-процесс генерации при сбое аудита
        print(f"[AUDIT_LOG_WARNING] Ошибка записи в {AUDIT_LOG_FILENAME}: {e}", file=sys.stderr)
        return False
