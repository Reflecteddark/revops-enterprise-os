"""
Bitrix24 → RevOps Enterprise OS V17.6: Коннектор синхронизации сделок.

Ключевое отличие от AmoCRM:
  Bitrix24 использует ВЕБХУКИ — не нужен OAuth, не нужен token refresh.
  Один URL вида https://mycompany.bitrix24.ru/rest/1/abc123/ — и всё.

Что делает:
  1. Через webhook URL обращается к Bitrix24 REST API
  2. Скачивает сделки, контакты, пользователей, этапы воронок
  3. Маппит этапы Bitrix24 → RevOps (1-6) по bitrix24_stage_mapping.json
  4. UPSERT в лист raw_deals: обновляет существующие, добавляет новые
  5. Пишет результат в audit_log.csv

Запуск:
  python bitrix24_connector.py                  — полная синхронизация
  python bitrix24_connector.py --days 7         — только за 7 дней
  python bitrix24_connector.py --dry-run        — предпросмотр
  python bitrix24_connector.py --setup          — интерактивная настройка
  python bitrix24_connector.py --show-stages    — список этапов воронок

Настройка вебхука в Bitrix24:
  Приложения → Вебхуки → Добавить вебхук (входящий)
  Разрешения: CRM, Пользователи
  Скопировать URL вида: https://ДОМЕН/rest/USER_ID/TOKEN/

Файлы:
  bitrix24_credentials.json     — webhook URL (в .gitignore!)
  bitrix24_stage_mapping.json   — маппинг stage_id → этапы 1-6
"""

import json
import sys
import os
import time
import glob
import argparse
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional
from dataclasses import dataclass, field, asdict

import requests
import openpyxl
from openpyxl.styles import PatternFill

# ─────────────────────────────────────────────
# Импорт конфига (graceful fallback)
# ─────────────────────────────────────────────
try:
    from config import BASE_DIR, XLSX_PATTERN
except ImportError:
    BASE_DIR = Path(__file__).parent
    XLSX_PATTERN = "RevOps Platform V17*.xlsx"

try:
    from audit_log import log_generation
    _HAS_AUDIT = True
except ImportError:
    _HAS_AUDIT = False
    def log_generation(**kwargs): pass  # noqa: stub

# ─────────────────────────────────────────────
# Константы
# ─────────────────────────────────────────────
CREDENTIALS_FILE = BASE_DIR / "bitrix24_credentials.json"
STAGE_MAPPING_FILE = BASE_DIR / "bitrix24_stage_mapping.json"

# Битрикс возвращает max 50 записей за вызов
B24_PAGE_SIZE = 50

# STAGE_SEMANTIC_ID: "S"=Success(WON), "F"=Failure(LOSE), "P"=Process(активный)
SEMANTIC_TO_STAGE = {"S": 6, "F": 0}

# Веса этапов для weighted_val
STAGE_WEIGHTS = {1: 0.05, 2: 0.15, 3: 0.30, 4: 0.50, 5: 0.75, 6: 1.0, 0: 0.0}

# Поля сделок для SELECT (чтобы не тащить всё)
DEAL_SELECT_FIELDS = [
    "ID", "TITLE", "OPPORTUNITY", "STAGE_ID", "STAGE_SEMANTIC_ID",
    "ASSIGNED_BY_ID", "CONTACT_ID", "COMPANY_ID", "CATEGORY_ID",
    "DATE_CREATE", "DATE_MODIFY", "CLOSEDATE", "PROBABILITY",
    "UTM_SOURCE", "UTM_MEDIUM", "UTM_CAMPAIGN", "YCLID", "GCLID",
    "SOURCE_ID", "SOURCE_DESCRIPTION", "COMMENTS",
    "UF_CRM_LEAD_SOURCE",  # частый кастомный поль источника
]

# Колонки raw_deals (строго по схеме Excel)
RAW_DEALS_HEADER = [
    "deal_id", "client_name", "amount", "stage_id", "stage_name",
    "created_date", "stage_changed_date", "manager_id", "segment_code",
    "lead_source", "client_id", "scoring_tier", "loss_code", "synced_at",
    "is_won", "is_lost", "days_in_stage", "aging_cohort", "weighted_val",
    "cycle_days", "utm_source", "utm_medium", "utm_campaign", "yclid",
    "gclid", "data_snapshot_date", "data_version", "uploaded_by",
    "upload_hash", "next_action_type", "next_action_date",
    "next_action_owner", "next_action_channel", "next_action_health",
    "deal_health_score",
]


# ─────────────────────────────────────────────
# Модели данных
# ─────────────────────────────────────────────
@dataclass
class B24Creds:
    webhook_url: str   # https://domain.bitrix24.ru/rest/1/token/
    domain: str = ""   # автовычисляется из webhook_url

    def __post_init__(self):
        if not self.domain:
            # Извлекаем домен из вебхука: https://X.bitrix24.ru/rest/...
            try:
                self.domain = self.webhook_url.split("/")[2]
            except (IndexError, AttributeError):
                self.domain = "unknown"

    def url(self, method: str) -> str:
        """Формирует полный URL метода."""
        base = self.webhook_url.rstrip("/")
        return f"{base}/{method}"

    @classmethod
    def load(cls, path: Path) -> "B24Creds":
        if not path.exists():
            print(f"[ОШИБКА] Файл учётных данных не найден: {path}")
            print("  Запустите: python bitrix24_connector.py --setup")
            sys.exit(1)
        data = json.loads(path.read_text(encoding="utf-8"))
        return cls(**data)

    def save(self, path: Path) -> None:
        path.write_text(
            json.dumps(asdict(self), ensure_ascii=False, indent=2),
            encoding="utf-8"
        )


@dataclass
class SyncResult:
    added: int = 0
    updated: int = 0
    skipped: int = 0
    errors: int = 0
    total_fetched: int = 0
    duration_sec: float = 0.0
    workbook_path: str = ""
    error_details: list = field(default_factory=list)

    def __str__(self) -> str:
        return (
            f"Загружено из Bitrix24: {self.total_fetched} | "
            f"Добавлено: {self.added} | Обновлено: {self.updated} | "
            f"Пропущено (проигрыш/без маппинга): {self.skipped} | "
            f"Ошибок: {self.errors} | Время: {self.duration_sec:.1f}с"
        )


# ─────────────────────────────────────────────
# HTTP-клиент Bitrix24
# ─────────────────────────────────────────────
class Bitrix24Client:
    """
    REST-клиент для Bitrix24 через вебхук.
    Автоматически обрабатывает пагинацию и rate-limit.
    """

    def __init__(self, creds: B24Creds):
        self.creds = creds
        self.session = requests.Session()
        self.session.headers.update({
            "Content-Type": "application/json",
            "User-Agent": "RevOps-Enterprise-OS/17.6",
        })

    def call(self, method: str, params: dict | None = None,
             retries: int = 3) -> dict:
        """POST к методу Bitrix24 REST API."""
        url = self.creds.url(method)
        payload = params or {}

        for attempt in range(retries):
            try:
                resp = self.session.post(url, json=payload, timeout=30)

                # Bitrix24 rate limit: 2 req/sec, при превышении → 503
                if resp.status_code == 503:
                    wait = 2 ** attempt
                    print(f"  [RATE LIMIT] Ждём {wait}с...")
                    time.sleep(wait)
                    continue

                if resp.status_code == 401:
                    print("[ОШИБКА] Вебхук недействителен или истёк.")
                    print("  Создайте новый вебхук и обновите credentials.")
                    sys.exit(1)

                resp.raise_for_status()
                data = resp.json()

                # Bitrix возвращает {"error": "...", "error_description": "..."}
                if "error" in data:
                    err = data.get("error_description", data["error"])
                    raise RuntimeError(f"Bitrix24 API error: {err}")

                return data

            except requests.RequestException as e:
                if attempt == retries - 1:
                    raise
                time.sleep(2 ** attempt)

        return {}

    def list_all(self, method: str, params: dict | None = None,
                 result_key: str = "result", label: str = "") -> list[dict]:
        """
        Постранично загружает все записи.
        Bitrix24 пагинация: параметр 'start' (0, 50, 100, ...).
        """
        results = []
        start = 0
        base_params = dict(params or {})

        while True:
            paged = {**base_params, "start": start}
            data = self.call(method, paged)

            items = data.get(result_key, [])
            if not items:
                break

            results.extend(items)
            page_num = start // B24_PAGE_SIZE + 1
            suffix = f" {label}" if label else ""
            print(f"  Страница {page_num}: загружено {len(items)}{suffix} "
                  f"(всего: {len(results)} / {data.get('total', '?')})")

            # "next" присутствует если есть следующая страница
            if "next" not in data:
                break

            start = data["next"]
            time.sleep(0.5)  # Bitrix24: не более 2 req/sec

        return results


# ─────────────────────────────────────────────
# Загрузка справочников
# ─────────────────────────────────────────────
def fetch_deals(client: Bitrix24Client, days: int | None = None) -> list[dict]:
    """Загружает сделки из Bitrix24."""
    params = {
        "select": DEAL_SELECT_FIELDS,
        "order": {"DATE_MODIFY": "DESC"},
    }
    if days:
        since = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%dT00:00:00")
        params["filter"] = {">DATE_MODIFY": since}

    return client.list_all("crm.deal.list", params, label="сделок")


def fetch_contacts(client: Bitrix24Client) -> dict[int, dict]:
    """Контакты → {id: contact}."""
    params = {"select": ["ID", "NAME", "LAST_NAME", "SECOND_NAME", "COMPANY_ID"]}
    contacts = client.list_all("crm.contact.list", params, label="контактов")
    result = {}
    for c in contacts:
        cid = int(c["ID"])
        first = c.get("NAME") or ""
        last = c.get("LAST_NAME") or ""
        name = f"{last} {first}".strip() or f"contact_{cid}"
        result[cid] = {"name": name, "id": cid}
    return result


def fetch_companies(client: Bitrix24Client) -> dict[int, str]:
    """Компании → {id: title}."""
    params = {"select": ["ID", "TITLE"]}
    companies = client.list_all("crm.company.list", params, label="компаний")
    return {int(c["ID"]): c.get("TITLE", "") for c in companies}


def fetch_users(client: Bitrix24Client) -> dict[int, str]:
    """Пользователи → {id: 'Фамилия И.О.'}."""
    params = {
        "select": ["ID", "NAME", "LAST_NAME", "SECOND_NAME"],
        "filter": {"ACTIVE": True},
    }
    users = client.list_all("user.get", params, label="пользователей")
    result = {}
    for u in users:
        uid = int(u["ID"])
        last = u.get("LAST_NAME") or ""
        first = (u.get("NAME") or "")[:1]
        second = (u.get("SECOND_NAME") or "")[:1]
        name = f"{last} {first}.{second}.".rstrip(".")
        result[uid] = name or f"user_{uid}"
    return result


def fetch_stages(client: Bitrix24Client) -> dict[str, dict]:
    """
    Этапы всех воронок → {stage_id: {name, pipeline_name, semantic}}.
    Bitrix24: стандартные этапы + этапы по категориям.
    stage_id формат: "NEW" (стандарт) или "C{category_id}:STAGE_NAME"
    """
    stages: dict[str, dict] = {}

    # 1. Стандартные этапы (воронка 0)
    data = client.call("crm.status.list", {"filter": {"ENTITY_ID": "DEAL_STAGE"}})
    for s in data.get("result", []):
        stages[s["STATUS_ID"]] = {
            "name": s.get("NAME", s["STATUS_ID"]),
            "pipeline_name": "Основная воронка",
            "category_id": 0,
            "semantic": s.get("SEMANTICS", "P"),
        }
    # WON/LOSE — стандартные
    stages.setdefault("WON",  {"name": "Успешно", "pipeline_name": "Основная воронка",
                                "category_id": 0, "semantic": "S"})
    stages.setdefault("LOSE", {"name": "Провал",  "pipeline_name": "Основная воронка",
                                "category_id": 0, "semantic": "F"})

    # 2. Воронки (категории сделок)
    cats_data = client.call("crm.dealcategory.list")
    for cat in cats_data.get("result", []):
        cid = cat["ID"]
        cat_name = cat.get("NAME", f"Воронка {cid}")
        # Этапы этой воронки
        s_data = client.call("crm.dealcategory.stage.list", {"id": cid})
        for s in s_data.get("result", []):
            sid = s["STATUS_ID"]  # формат "C{cid}:STAGENAME"
            stages[sid] = {
                "name": s.get("NAME", sid),
                "pipeline_name": cat_name,
                "category_id": int(cid),
                "semantic": s.get("SEMANTICS", "P"),
            }
        # WON/LOSE для этой воронки
        stages[f"C{cid}:WON"]  = {"name": "Успешно", "pipeline_name": cat_name,
                                    "category_id": int(cid), "semantic": "S"}
        stages[f"C{cid}:LOSE"] = {"name": "Провал",  "pipeline_name": cat_name,
                                    "category_id": int(cid), "semantic": "F"}

    return stages


# ─────────────────────────────────────────────
# Маппинг этапов
# ─────────────────────────────────────────────
def load_stage_mapping(path: Path) -> dict[str, int]:
    """
    Загружает маппинг Bitrix24 stage_id → RevOps stage_id (1-6, 0=проигрыш).
    Поддерживает форматы: "NEW" и "C1:NEW" (с категорией).
    Если файла нет — создаёт шаблон.
    """
    if not path.exists():
        default = {
            "_comment": (
                "Ключи — stage_id из Bitrix24. Стандартная воронка: 'NEW', 'IN_PROCESS' и т.д. "
                "Пользовательские воронки: 'C1:STAGE_NAME' (C + category_id + :stage). "
                "Запустите --show-stages чтобы увидеть все этапы. "
                "RevOps этапы: 1=Новый, 2=Квалификация, 3=Переговоры, 4=КП, 5=Подписание, "
                "6=Закрыт(выиграш), 0=Закрыт(проигрыш)"
            ),
            "_semantic_fallback": (
                "Если stage_id не найден в маппинге, используется STAGE_SEMANTIC_ID: "
                "S=6(WON), F=0(LOSE), P=3(активный этап по умолчанию)"
            ),
            "WON": 6,
            "LOSE": 0,
            "NEW": 1,
            "IN_PROCESS": 2,
            "PREPARATION": 3,
            "PREPAYMENT_INVOICE": 4,
            "EXECUTING": 5,
            "FINAL_INVOICE": 5,
            "EXAMPLE_C1:NEW": 1,
            "EXAMPLE_C1:WON": 6,
        }
        path.write_text(
            json.dumps(default, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )
        print(f"\n[ВНИМАНИЕ] Создан шаблон маппинга: {path}")
        print("  Запустите: python bitrix24_connector.py --show-stages\n")

    raw = json.loads(path.read_text(encoding="utf-8"))
    mapping = {}
    for k, v in raw.items():
        if k.startswith("_"):
            continue
        try:
            mapping[str(k)] = int(v)
        except (ValueError, TypeError):
            pass
    return mapping


def resolve_stage(
    b24_stage_id: str,
    semantic: str,
    stage_mapping: dict[str, int],
) -> int:
    """
    Определяет RevOps stage_id с несколькими fallback-уровнями:
    1. Точный маппинг по stage_id
    2. Маппинг по semantic (S→6, F→0)
    3. Default для активных этапов (P → 3)
    """
    if b24_stage_id in stage_mapping:
        return stage_mapping[b24_stage_id]
    # Fallback по semantic
    if semantic == "S":
        return 6
    if semantic == "F":
        return 0
    # Активный этап без маппинга → считаем как "Переговоры"
    return 3


# ─────────────────────────────────────────────
# Трансформация сделок
# ─────────────────────────────────────────────
def _parse_b24_date(value: str | None) -> str:
    """Парсит дату Bitrix24 → 'YYYY-MM-DD'."""
    if not value:
        return ""
    try:
        # Bitrix24 возвращает ISO 8601 с TZ: "2026-09-15T14:30:00+03:00"
        dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
        return dt.strftime("%Y-%m-%d")
    except (ValueError, AttributeError):
        return str(value)[:10]


def _calc_days_since(date_str: str) -> int:
    """Количество дней с указанной даты до сегодня."""
    if not date_str:
        return 0
    try:
        d = datetime.strptime(date_str, "%Y-%m-%d").date()
        return (datetime.now().date() - d).days
    except ValueError:
        return 0


def _calc_health_score(stage_id: int, days_in_stage: int,
                        amount: float, cycle_days: int) -> int:
    """Скоринг сделки 0-100."""
    score = 100
    if days_in_stage > 30:
        score -= 40
    elif days_in_stage > 14:
        score -= 20
    elif days_in_stage > 7:
        score -= 10
    if cycle_days > 90:
        score -= 20
    elif cycle_days > 45:
        score -= 10
    if stage_id >= 5:
        score += 10
    elif stage_id <= 1:
        score -= 5
    return max(0, min(100, score))


def map_deal_to_row(
    deal: dict,
    contacts_map: dict[int, dict],
    companies_map: dict[int, str],
    users_map: dict[int, str],
    stages_map: dict[str, dict],
    stage_mapping: dict[str, int],
    snapshot_date: str,
) -> dict | None:
    """
    Трансформирует сделку Bitrix24 → строку raw_deals.
    Возвращает None для WON/LOSE (не в активном пайплайне).
    """
    b24_stage_id = deal.get("STAGE_ID", "")
    semantic = deal.get("STAGE_SEMANTIC_ID", "P")
    stage_id = resolve_stage(b24_stage_id, semantic, stage_mapping)

    # Информация о этапе и воронке
    stage_info = stages_map.get(b24_stage_id, {})
    stage_name = stage_info.get("name") or b24_stage_id
    pipeline_name = stage_info.get("pipeline_name", "")

    # Клиент: контакт или компания
    client_name = ""
    client_id = ""
    contact_id = deal.get("CONTACT_ID")
    company_id = deal.get("COMPANY_ID")

    if contact_id:
        cid = int(contact_id)
        client_id = str(cid)
        client_name = contacts_map.get(cid, {}).get("name", "")

    if not client_name and company_id:
        coid = int(company_id)
        client_name = companies_map.get(coid, "")
        if not client_id:
            client_id = f"company_{coid}"

    # Менеджер
    assigned_id = deal.get("ASSIGNED_BY_ID", 0)
    try:
        assigned_id_int = int(assigned_id)
    except (ValueError, TypeError):
        assigned_id_int = 0
    manager_name = users_map.get(assigned_id_int, f"user_{assigned_id}")

    # Суммы
    amount_raw = deal.get("OPPORTUNITY") or 0
    try:
        amount = float(amount_raw)
    except (ValueError, TypeError):
        amount = 0.0

    # Даты
    created_date = _parse_b24_date(deal.get("DATE_CREATE"))
    modified_date = _parse_b24_date(deal.get("DATE_MODIFY"))
    days_in_stage = _calc_days_since(modified_date)
    cycle_days = _calc_days_since(created_date)

    # Aging cohort
    if days_in_stage <= 3:
        aging = "Fresh"
    elif days_in_stage <= 7:
        aging = "Active"
    elif days_in_stage <= 14:
        aging = "Aging"
    elif days_in_stage <= 30:
        aging = "Stale"
    else:
        aging = "Dead"

    # UTM (в Bitrix24 встроены в сделку)
    utm_source = deal.get("UTM_SOURCE") or ""
    utm_medium = deal.get("UTM_MEDIUM") or ""
    utm_campaign = deal.get("UTM_CAMPAIGN") or ""
    yclid = deal.get("YCLID") or ""
    gclid = deal.get("GCLID") or ""

    # Источник лида
    lead_source = (
        deal.get("SOURCE_ID") or
        deal.get("UF_CRM_LEAD_SOURCE") or
        utm_source or ""
    )

    # Weighted pipeline
    weight = STAGE_WEIGHTS.get(stage_id, 0)
    weighted_val = round(amount * weight)

    # Health score
    health = _calc_health_score(stage_id, days_in_stage, amount, cycle_days)

    # Причина проигрыша
    loss_code = ""
    if stage_id == 0:
        loss_code = deal.get("SOURCE_DESCRIPTION") or "Не указана"

    return {
        "deal_id": str(deal["ID"]),
        "client_name": client_name,
        "amount": amount,
        "stage_id": stage_id,
        "stage_name": stage_name,
        "created_date": created_date,
        "stage_changed_date": modified_date,
        "manager_id": manager_name,
        "segment_code": pipeline_name,
        "lead_source": lead_source,
        "client_id": client_id,
        "scoring_tier": "A" if amount > 1_000_000 else ("B" if amount > 300_000 else "C"),
        "loss_code": loss_code,
        "synced_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "is_won": 1 if stage_id == 6 else 0,
        "is_lost": 1 if stage_id == 0 else 0,
        "days_in_stage": days_in_stage,
        "aging_cohort": aging,
        "weighted_val": weighted_val,
        "cycle_days": cycle_days,
        "utm_source": utm_source,
        "utm_medium": utm_medium,
        "utm_campaign": utm_campaign,
        "yclid": yclid,
        "gclid": gclid,
        "data_snapshot_date": snapshot_date,
        "data_version": "17.6",
        "uploaded_by": "bitrix24_connector",
        "upload_hash": str(deal["ID"]),
        "next_action_type": "",
        "next_action_date": "",
        "next_action_owner": manager_name,
        "next_action_channel": "",
        "next_action_health": "OK" if health >= 50 else "РИСК",
        "deal_health_score": health,
    }


# ─────────────────────────────────────────────
# Запись в Excel (UPSERT — идентична AmoCRM)
# ─────────────────────────────────────────────
def find_workbook_path() -> str:
    pattern = str(BASE_DIR / XLSX_PATTERN)
    candidates = [
        f for f in glob.glob(pattern)
        if not os.path.basename(f).startswith("~$")
    ]
    if not candidates:
        print(f"[ОШИБКА] Не найден файл {XLSX_PATTERN} в {BASE_DIR}")
        sys.exit(1)
    return max(candidates, key=os.path.getmtime)


def write_deals_to_excel(
    deals: list[dict],
    wb_path: str,
    dry_run: bool = False,
) -> SyncResult:
    """UPSERT сделок в raw_deals. Совместим с AmoCRM коннектором."""
    result = SyncResult(workbook_path=wb_path)
    print(f"\n  Открываем Excel: {wb_path}")

    wb = openpyxl.load_workbook(wb_path)
    ws = wb["raw_deals"]

    # Индекс колонок
    header_row = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    col_index = {}
    for col_name in RAW_DEALS_HEADER:
        try:
            col_index[col_name] = header_row.index(col_name) + 1
        except ValueError:
            col_index[col_name] = None

    deal_id_col = col_index.get("deal_id", 1)

    # Индекс существующих сделок
    existing = {}
    for row in range(2, ws.max_row + 1):
        val = ws.cell(row, deal_id_col).value
        if val is not None:
            existing[str(val)] = row

    print(f"  Существующих строк: {ws.max_row - 1}")
    print(f"  Из Bitrix24 к записи: {len(deals)}")

    if dry_run:
        print("\n  [DRY RUN] Изменения НЕ применяются.")
        for d in deals[:5]:
            print(f"    ID={d['deal_id']} stage={d['stage_id']} "
                  f"amount={d['amount']:,.0f} manager={d['manager_id']}")
        if len(deals) > 5:
            print(f"    ... и ещё {len(deals) - 5}")
        result.total_fetched = len(deals)
        return result

    # Стили
    b24_fill  = PatternFill(start_color="FFF8F0", end_color="FFF8F0", fill_type="solid")
    won_fill  = PatternFill(start_color="E6F4EA", end_color="E6F4EA", fill_type="solid")
    lost_fill = PatternFill(start_color="FCE8E8", end_color="FCE8E8", fill_type="solid")

    # Находим первую свободную строку (учитываем шаблоны с пустыми строками до 1000)
    last_filled = 1
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, deal_id_col).value
        if v is not None and str(v).strip():
            last_filled = r
    next_row = last_filled + 1

    for deal in deals:
        deal_id = deal["deal_id"]
        row = existing.get(deal_id)

        if row is None:
            row = next_row
            next_row += 1
            result.added += 1
        else:
            result.updated += 1

        for col_name, value in deal.items():
            col_num = col_index.get(col_name)
            if col_num is None:
                continue
            cell = ws.cell(row, col_num)
            cell.value = value
            stage = deal.get("stage_id", 0)
            cell.fill = (won_fill if stage == 6
                         else lost_fill if stage == 0
                         else b24_fill)

    result.total_fetched = len(deals)
    wb.save(wb_path)
    wb.close()
    print(f"  Excel сохранён: {wb_path}")
    return result


# ─────────────────────────────────────────────
# CLI команды
# ─────────────────────────────────────────────
def cmd_show_stages(creds: B24Creds) -> None:
    """Выводит все этапы всех воронок — для настройки маппинга."""
    client = Bitrix24Client(creds)
    print(f"\n  Воронки и этапы в {creds.domain}:\n")
    stages = fetch_stages(client)

    by_pipeline: dict[str, list] = {}
    for sid, info in stages.items():
        p = info["pipeline_name"]
        by_pipeline.setdefault(p, []).append((sid, info))

    for p_name, items in by_pipeline.items():
        print(f"  📊 Воронка: {p_name}")
        for sid, info in items:
            sem = info.get("semantic", "P")
            hint = ""
            if sem == "S":
                hint = " ← ПОБЕДА (stage_id=6)"
            elif sem == "F":
                hint = " ← ПРОИГРЫШ (stage_id=0)"
            print(f"      {sid:<30}  [{info['name']}]{hint}")
        print()

    print(f"  Добавьте нужные stage_id в: {STAGE_MAPPING_FILE}")


def cmd_setup() -> None:
    """Интерактивная настройка вебхука."""
    print("\n" + "=" * 60)
    print("  Настройка Bitrix24 коннектора")
    print("=" * 60)
    print()
    print("1. Войдите в Bitrix24")
    print("2. Перейдите: Приложения → Разработчикам → Другое →")
    print("             Входящий вебхук → Добавить")
    print("3. Отметьте разрешения: CRM, Пользователи")
    print("4. Скопируйте URL вебхука (вида https://ДОМЕН/rest/1/TOKEN/)")
    print()

    webhook_url = input("Вставьте URL вебхука: ").strip()
    if not webhook_url.endswith("/"):
        webhook_url += "/"

    # Проверка соединения
    print("\nПроверяем соединение...")
    creds = B24Creds(webhook_url=webhook_url)
    client = Bitrix24Client(creds)
    try:
        data = client.call("app.info")
        print(f"✓ Соединение установлено. Аккаунт: {creds.domain}")
    except Exception as e:
        try:
            # app.info может быть недоступен, пробуем user.current
            data = client.call("user.current")
            print(f"✓ Соединение установлено.")
        except Exception as e2:
            print(f"[ОШИБКА] Не удалось подключиться: {e2}")
            print("  Проверьте URL вебхука и разрешения.")
            sys.exit(1)

    creds.save(CREDENTIALS_FILE)
    print(f"\n✓ Учётные данные сохранены: {CREDENTIALS_FILE}")
    print(f"\nТеперь запустите: python bitrix24_connector.py --show-stages")
    print("Чтобы настроить маппинг этапов под вашу воронку.")


# ─────────────────────────────────────────────
# Основная синхронизация
# ─────────────────────────────────────────────
def sync(
    wb_path: str | None = None,
    days: int | None = None,
    dry_run: bool = False,
) -> SyncResult:
    """
    Полная синхронизация Bitrix24 → raw_deals.

    Args:
        wb_path:  Путь к Excel. Если None — ищет по маске.
        days:     Только сделки за последние N дней. None = все.
        dry_run:  Не писать в Excel.

    Returns:
        SyncResult с метриками.
    """
    start_time = time.time()
    snapshot_date = datetime.now().strftime("%Y-%m-%d")

    # Учётные данные
    creds = B24Creds.load(CREDENTIALS_FILE)
    client = Bitrix24Client(creds)

    # Маппинг этапов
    stage_mapping = load_stage_mapping(STAGE_MAPPING_FILE)

    print(f"\nЗагрузка данных из {creds.domain}...")

    # Справочники
    print("  Этапы воронок...")
    stages_map = fetch_stages(client)
    print(f"  Найдено этапов: {len(stages_map)}")

    print("  Пользователи...")
    users_map = fetch_users(client)
    print(f"  Найдено пользователей: {len(users_map)}")

    print("  Контакты...")
    contacts_map = fetch_contacts(client)
    print(f"  Найдено контактов: {len(contacts_map)}")

    print("  Компании...")
    companies_map = fetch_companies(client)
    print(f"  Найдено компаний: {len(companies_map)}")

    print(f"  Сделки (период: {'все' if not days else f'последние {days} дн.'})...")
    raw_deals = fetch_deals(client, days=days)
    print(f"  Загружено сделок: {len(raw_deals)}")

    # Трансформация
    deals = []
    skipped = 0
    for deal in raw_deals:
        row = map_deal_to_row(
            deal, contacts_map, companies_map, users_map,
            stages_map, stage_mapping, snapshot_date
        )
        if row is None:
            skipped += 1
            continue
        deals.append(row)

    print(f"  После трансформации: {len(deals)} сделок "
          f"({skipped} пропущено)")

    # Запись
    if wb_path is None:
        wb_path = find_workbook_path()

    result = write_deals_to_excel(deals, wb_path, dry_run=dry_run)
    result.skipped = skipped
    result.duration_sec = round(time.time() - start_time, 1)

    # Аудит
    if not dry_run and _HAS_AUDIT:
        log_generation(
            status="success",
            duration_sec=result.duration_sec,
            role="sync",
            report_type="Bitrix24_Sync",
            pdf_size_kb=0,
        )

    return result


# ─────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────
def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(
        description="RevOps Bitrix24 Connector — синхронизация сделок в Excel",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры:
  python bitrix24_connector.py                  — полная синхронизация
  python bitrix24_connector.py --days 30        — только за 30 дней
  python bitrix24_connector.py --dry-run        — предпросмотр
  python bitrix24_connector.py --setup          — первоначальная настройка
  python bitrix24_connector.py --show-stages    — список этапов воронок
        """
    )
    parser.add_argument("--setup", action="store_true",
                        help="Интерактивная настройка вебхука")
    parser.add_argument("--show-stages", action="store_true",
                        help="Показать этапы воронок для маппинга")
    parser.add_argument("--days", type=int, default=None,
                        help="Синхронизировать за последние N дней")
    parser.add_argument("--dry-run", action="store_true",
                        help="Показать что будет, не писать в Excel")
    parser.add_argument("--workbook", type=str, default=None,
                        help="Явный путь к Excel-файлу")

    args = parser.parse_args()

    if args.setup:
        cmd_setup()
        return 0

    if args.show_stages:
        creds = B24Creds.load(CREDENTIALS_FILE)
        cmd_show_stages(creds)
        return 0

    print("=" * 60)
    print("  RevOps Bitrix24 Connector V17.6")
    print("=" * 60)

    try:
        result = sync(
            wb_path=args.workbook,
            days=args.days,
            dry_run=args.dry_run,
        )
        print(f"\n{'='*60}")
        print(f"  РЕЗУЛЬТАТ: {result}")
        print(f"{'='*60}\n")
        return 0

    except KeyboardInterrupt:
        print("\nПрервано пользователем.")
        return 1
    except Exception as e:
        print(f"\n[КРИТИЧЕСКАЯ ОШИБКА] {e}", file=sys.stderr)
        if _HAS_AUDIT:
            log_generation(
                status="error",
                duration_sec=0,
                role="sync",
                report_type="Bitrix24_Sync",
                error_message=str(e),
            )
        raise


if __name__ == "__main__":
    sys.exit(main())
