"""
Централизованная конфигурация путей и параметров RevOps Enterprise OS V17.6.
"""
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
XLSX_PATTERN = "RevOps Platform V17*.xlsx"
AUDIT_LOG_FILENAME = "audit_log.csv"
