"""
AmoCRM → RevOps Enterprise OS V17.6: Коннектор синхронизации сделок.

Что делает:
  1. Авторизуется в AmoCRM через OAuth 2.0 (auto-refresh токенов)
  2. Скачивает сделки, контакты, пользователей, пайплайны
  3. Маппит статусы AmoCRM → этапы RevOps (1-6) по stage_mapping.json
  4. UPSERT в лист raw_deals: обновляет существующие, добавляет новые
  5. Пишет результат в audit_log.csv

Запуск:
  python amocrm_connector.py               — полная синхронизация
  python amocrm_connector.py --days 7      — только за последние 7 дней
  python amocrm_connector.py --setup       — интерактивная настройка токенов
  python amocrm_connector.py --dry-run     — только показать, не писать в Excel

Настройка:
  1. Создайте интеграцию в AmoCRM: Настройки → Интеграции → + Создать интеграцию
  2. Запустите: python amocrm_connector.py --setup
  3. Следуйте инструкциям

Файлы:
  amocrm_credentials.json   — токены и домен (в .gitignore!)
  amocrm_stage_mapping.json — маппинг статусов AmoCRM → этапы 1-6
"""

import json
import sys
import os
import time
import glob
import argparse
import webbrowser
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Optional
from dataclasses import dataclass, field, asdict
from urllib.parse import urlencode
import re
import random

import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

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
# Константы и пути
# ─────────────────────────────────────────────
CREDENTIALS_FILE = BASE_DIR / "amocrm_credentials.json"
STAGE_MAPPING_FILE = BASE_DIR / "amocrm_stage_mapping.json"

AMO_API_VERSION = "v4"
AMO_TOKEN_URL = "https://{domain}/oauth2/access_token"
AMO_LEADS_URL = "https://{domain}/api/v4/leads"
AMO_CONTACTS_URL = "https://{domain}/api/v4/contacts"
AMO_USERS_URL = "https://{domain}/api/v4/users"
AMO_PIPELINES_URL = "https://{domain}/api/v4/leads/pipelines"

# Этапы 1-6 RevOps (стандарт продукта)
REVOPS_STAGES = {
    1: "Новый лид",
    2: "Квалификация",
    3: "Переговоры",
    4: "Коммерческое предложение",
    5: "Согласование/Подписание",
    6: "Успешно реализовано",
    0: "Закрыто (проигрыш)",   # не попадает в активный пайплайн
}

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

# Веса этапов для weighted_val (weighted pipeline)
STAGE_WEIGHTS = {1: 0.05, 2: 0.15, 3: 0.30, 4: 0.50, 5: 0.75, 6: 1.0, 0: 0.0}


# ─────────────────────────────────────────────
# Модели данных
# ─────────────────────────────────────────────
@dataclass
class AmoCreds:
    domain: str          # example.amocrm.ru
    client_id: str
    client_secret: str
    access_token: str
    refresh_token: str
    token_expires_at: float = 0.0  # unix timestamp

    @classmethod
    def load(cls, path: Path) -> "AmoCreds":
        if not path.exists():
            print(f"[ОШИБКА] Файл учётных данных не найден: {path}")
            print("  Запустите: python amocrm_connector.py --setup")
            sys.exit(1)
        data = json.loads(path.read_text(encoding="utf-8"))
        return cls(**data)

    def save(self, path: Path) -> None:
        path.write_text(
            json.dumps(asdict(self), ensure_ascii=False, indent=2),
            encoding="utf-8"
        )

    def is_token_expired(self) -> bool:
        # Считаем истёкшим за 5 минут до реального истечения
        return time.time() > (self.token_expires_at - 300)


@dataclass
class SyncResult:
    added: int = 0
    updated: int = 0
    skipped: int = 0  # stage_id == 0 (проигранные) — не в пайплайне
    errors: int = 0
    total_fetched: int = 0
    duration_sec: float = 0.0
    workbook_path: str = ""
    error_details: list = field(default_factory=list)

    def __str__(self) -> str:
        return (
            f"Загружено из AmoCRM: {self.total_fetched} | "
            f"Добавлено: {self.added} | Обновлено: {self.updated} | "
            f"Пропущено (проигрыш): {self.skipped} | "
            f"Ошибок: {self.errors} | "
            f"Время: {self.duration_sec:.1f}с"
        )


# ─────────────────────────────────────────────
# Авторизация, HTTP и Rate Limiting (Защита от лимитов API)
# ─────────────────────────────────────────────
class RateLimiter:
    """
    Клиентский ограничитель скорости запросов (Token Bucket / Interval Pacing).
    Защищает от превышения лимитов amoCRM API (~7 req/sec) и ошибок 429 Too Many Requests.
    """
    def __init__(self, max_per_sec: float = 6.0):
        self.min_interval = 1.0 / max_per_sec if max_per_sec > 0 else 0.0
        self.last_call_ts = 0.0

    def wait(self) -> None:
        elapsed = time.time() - self.last_call_ts
        if elapsed < self.min_interval:
            time.sleep(self.min_interval - elapsed)
        self.last_call_ts = time.time()


class AmoCRMClient:
    """HTTP-клиент для AmoCRM API v4 с auto-refresh токенов и защитой от лимитов."""

    def __init__(self, creds: AmoCreds, max_requests_per_sec: float = 6.0):
        self.creds = creds
        self.rate_limiter = RateLimiter(max_per_sec=max_requests_per_sec)
        self.session = requests.Session()
        self.session.headers.update({
            "Content-Type": "application/json",
            "User-Agent": "RevOps-Enterprise-OS/17.6",
        })
        self._cache: dict[str, dict] = {}

    def _refresh_token(self) -> None:
        """Обновляет access_token через refresh_token."""
        print("  [AUTH] Обновление access_token...")
        url = AMO_TOKEN_URL.format(domain=self.creds.domain)
        payload = {
            "client_id": self.creds.client_id,
            "client_secret": self.creds.client_secret,
            "grant_type": "refresh_token",
            "refresh_token": self.creds.refresh_token,
            "redirect_uri": "https://example.com/oauth",
        }
        resp = requests.post(url, json=payload, timeout=30)
        if resp.status_code != 200:
            print(f"  [AUTH ERROR] {resp.status_code}: {resp.text[:200]}")
            raise RuntimeError(f"Не удалось обновить токен AmoCRM: {resp.status_code}")

        data = resp.json()
        self.creds.access_token = data["access_token"]
        self.creds.refresh_token = data["refresh_token"]
        self.creds.token_expires_at = time.time() + data.get("expires_in", 86400)
        self.creds.save(CREDENTIALS_FILE)
        print("  [AUTH] Токен обновлён и сохранён.")

    def _ensure_token(self) -> None:
        if self.creds.is_token_expired():
            self._refresh_token()
        self.session.headers["Authorization"] = f"Bearer {self.creds.access_token}"

    def get(self, url: str, params: dict | None = None,
            retries: int = 5) -> dict:
        """GET с клиентским rate-limiting и экспоненциальным backoff при 401/429/5xx."""
        self._ensure_token()
        for attempt in range(retries):
            self.rate_limiter.wait()
            try:
                resp = self.session.get(url, params=params, timeout=30)
                if resp.status_code == 401:
                    self._refresh_token()
                    self.session.headers["Authorization"] = (
                        f"Bearer {self.creds.access_token}"
                    )
                    continue
                if resp.status_code in (429, 503):
                    header_retry = resp.headers.get("Retry-After")
                    base_wait = float(header_retry) if header_retry else (2 ** attempt)
                    wait = base_wait + random.uniform(0.2, 0.8)
                    print(f"  [RATE LIMIT {resp.status_code}] Ждём {wait:.1f}с (попытка {attempt+1}/{retries})...")
                    time.sleep(wait)
                    continue
                if resp.status_code == 204:
                    return {}  # No content — пустая страница
                resp.raise_for_status()
                return resp.json()
            except requests.RequestException as e:
                if attempt == retries - 1:
                    raise
                backoff = (2 ** attempt) + random.uniform(0.2, 0.8)
                time.sleep(backoff)
        return {}

    def post(self, url: str, json_data: dict | None = None,
             retries: int = 5) -> dict:
        """POST с клиентским rate-limiting и защитой от перегрузок."""
        self._ensure_token()
        for attempt in range(retries):
            self.rate_limiter.wait()
            try:
                resp = self.session.post(url, json=json_data, timeout=30)
                if resp.status_code == 401:
                    self._refresh_token()
                    self.session.headers["Authorization"] = (
                        f"Bearer {self.creds.access_token}"
                    )
                    continue
                if resp.status_code in (429, 503):
                    header_retry = resp.headers.get("Retry-After")
                    base_wait = float(header_retry) if header_retry else (2 ** attempt)
                    wait = base_wait + random.uniform(0.2, 0.8)
                    print(f"  [RATE LIMIT {resp.status_code}] Ждём {wait:.1f}с (попытка {attempt+1}/{retries})...")
                    time.sleep(wait)
                    continue
                resp.raise_for_status()
                return resp.json() if resp.content else {}
            except requests.RequestException as e:
                if attempt == retries - 1:
                    raise
                backoff = (2 ** attempt) + random.uniform(0.2, 0.8)
                time.sleep(backoff)
        return {}


# ─────────────────────────────────────────────
# Нормализация телефонов и умный матчинг звонков (Защита от дублей)
# ─────────────────────────────────────────────
def normalize_phone(phone: str | None) -> str:
    """
    Нормализует телефон к формату E.164 (+7XXXXXXXXXX для РФ/СНГ).
    Удаляет пробелы, тире, скобки.
    89XXXXXXXXX -> +79XXXXXXXXX
    79XXXXXXXXX -> +79XXXXXXXXX
    9XXXXXXXXX  -> +79XXXXXXXXX
    """
    if not phone:
        return ""
    digits = re.sub(r"\D", "", str(phone))
    if len(digits) == 11 and digits.startswith("8"):
        return "+7" + digits[1:]
    elif len(digits) == 11 and digits.startswith("7"):
        return "+" + digits
    elif len(digits) == 10 and digits.startswith("9"):
        return "+7" + digits
    elif digits:
        return "+" + digits if not str(phone).strip().startswith("+") else "+" + digits
    return ""


def extract_contact_phones(contact: dict) -> list[str]:
    """
    Извлекает все телефоны контакта amoCRM из custom_fields_values (field_code='PHONE').
    Возвращает список нормализованных номеров.
    """
    phones = []
    cf_values = contact.get("custom_fields_values") or []
    for cf in cf_values:
        if cf.get("field_code") == "PHONE":
            for item in cf.get("values", []):
                val = item.get("value")
                norm = normalize_phone(val)
                if norm and norm not in phones:
                    phones.append(norm)
    return phones


def resolve_target_deal(
    candidate_deals: list[dict],
    stage_mapping: dict | None = None,
) -> dict | None:
    """
    Устранение коллизий псевдо-дублей: сопоставляет звонок с наиболее релевантной сделкой.
    
    Правила приоритизации (из исследований рынка RevOps):
      1. Приоритет активных стадий: открытые сделки (RevOps 1-5, не выиграны 6 и не проиграны 0).
      2. Приоритет суммы (amount / price): при наличии нескольких открытых сделок выбирается сделка с максимальной суммой.
      3. Приоритет свежести (updated_at): при равенстве сумм выбирается сделка с самой свежей активностью.
      4. Фоллбэк: если все сделки закрыты - выбирается наиболее свежая закрытая (кандидат на повторную продажу).
    """
    if not candidate_deals:
        return None
    if len(candidate_deals) == 1:
        return candidate_deals[0]

    def _is_active(d: dict) -> bool:
        status_id = d.get("status_id", 0)
        if stage_mapping and status_id in stage_mapping:
            stage_id = stage_mapping[status_id]
            return stage_id not in (0, 6)
        # Дефолтные статусы amoCRM: 142=won, 143=lost
        return status_id not in (142, 143)

    active_deals = [d for d in candidate_deals if _is_active(d)]
    pool = active_deals if active_deals else candidate_deals

    def _deal_priority_key(d: dict):
        price = d.get("price") or d.get("amount") or 0
        try:
            price = float(price)
        except (ValueError, TypeError):
            price = 0.0
        updated = d.get("updated_at") or d.get("updated_ts") or 0
        try:
            updated = float(updated)
        except (ValueError, TypeError):
            updated = 0.0
        return (price, updated)

    return max(pool, key=_deal_priority_key)


def find_deals_by_phone(
    phone: str,
    all_deals: list[dict],
    contacts_map: dict[int, dict],
) -> list[dict]:
    """
    Ищет все сделки, привязанные к контактам с данным телефонным номером.
    """
    norm_target = normalize_phone(phone)
    if not norm_target:
        return []

    matched_contact_ids = set()
    for cid, cdata in contacts_map.items():
        phones = extract_contact_phones(cdata)
        if norm_target in phones:
            matched_contact_ids.add(cid)

    if not matched_contact_ids:
        return []

    matched_deals = []
    for d in all_deals:
        main_cid = d.get("main_contact_id")
        if main_cid in matched_contact_ids:
            matched_deals.append(d)
            continue
        emb_contacts = d.get("_embedded", {}).get("contacts", [])
        for ec in emb_contacts:
            if ec.get("id") in matched_contact_ids:
                matched_deals.append(d)
                break

    return matched_deals


def match_call_to_deal(
    call_phone: str,
    all_deals: list[dict],
    contacts_map: dict[int, dict],
    stage_mapping: dict | None = None,
) -> dict | None:
    """
    Сквозной матчинг аудиозвонка к целевой сделке в amoCRM с защитой от псевдо-дублей.
    """
    candidate_deals = find_deals_by_phone(call_phone, all_deals, contacts_map)
    return resolve_target_deal(candidate_deals, stage_mapping)


# ─────────────────────────────────────────────
# Загрузка данных из AmoCRM
# ─────────────────────────────────────────────
def _fetch_all_pages(client: AmoCRMClient, url: str,
                     entity_key: str, params: dict | None = None,
                     label: str = "") -> list[dict]:
    """Постранично скачивает все записи из AmoCRM (лимит 250/страница)."""
    results = []
    page = 1
    base_params = {"limit": 250, **(params or {})}

    while True:
        paged_params = {**base_params, "page": page}
        data = client.get(url, params=paged_params)

        if not data or "_embedded" not in data:
            break

        items = data["_embedded"].get(entity_key, [])
        if not items:
            break

        results.extend(items)
        suffix = f" {label}" if label else ""
        print(f"  Страница {page}: загружено {len(items)}{suffix} "
              f"(всего: {len(results)})")

        # Проверяем наличие следующей страницы
        links = data.get("_links", {})
        if "next" not in links:
            break
        page += 1
        time.sleep(0.2)  # Небольшая пауза для rate-limit

    return results


def fetch_leads(client: AmoCRMClient, days: int | None = None) -> list[dict]:
    """Загружает сделки из AmoCRM."""
    url = AMO_LEADS_URL.format(domain=client.creds.domain)
    params = {"with": "contacts,loss_reason,tags"}

    if days:
        since = int((datetime.now() - timedelta(days=days)).timestamp())
        params["filter[updated_at][from]"] = since

    return _fetch_all_pages(client, url, "leads", params, "сделок")


def fetch_contacts(client: AmoCRMClient) -> dict[int, dict]:
    """Загружает контакты → словарь {id: contact}."""
    url = AMO_CONTACTS_URL.format(domain=client.creds.domain)
    contacts = _fetch_all_pages(client, url, "contacts", label="контактов")
    return {c["id"]: c for c in contacts}


def fetch_users(client: AmoCRMClient) -> dict[int, str]:
    """Загружает пользователей → словарь {id: 'Фамилия И.О.'}."""
    url = AMO_USERS_URL.format(domain=client.creds.domain)
    data = client.get(url)
    users = {}
    if data and "_embedded" in data:
        for u in data["_embedded"].get("users", []):
            name = u.get("name") or f"user_{u['id']}"
            users[u["id"]] = name
    return users


def fetch_pipelines(client: AmoCRMClient) -> dict[int, dict]:
    """Загружает пайплайны и их статусы → {status_id: {name, pipeline_name}}."""
    url = AMO_PIPELINES_URL.format(domain=client.creds.domain)
    data = client.get(url)
    status_map = {}
    if data and "_embedded" in data:
        for pipeline in data["_embedded"].get("pipelines", []):
            p_name = pipeline.get("name", "")
            for status in pipeline.get("_embedded", {}).get("statuses", []):
                status_map[status["id"]] = {
                    "name": status.get("name", ""),
                    "pipeline_name": p_name,
                    "pipeline_id": pipeline["id"],
                    "type": status.get("type", 0),
                    # type 142 = won, 143 = lost в AmoCRM
                }
    return status_map


# ─────────────────────────────────────────────
# Маппинг и трансформация
# ─────────────────────────────────────────────
def load_stage_mapping(path: Path) -> dict[int, int]:
    """
    Загружает маппинг AmoCRM status_id → RevOps stage_id (1-6, 0=проигрыш).
    Если файла нет — создаёт дефолтный и просит настроить.
    """
    if not path.exists():
        default = {
            "_comment": (
                "Замените ключи на реальные status_id из вашего AmoCRM. "
                "Запустите python amocrm_connector.py --show-stages чтобы "
                "увидеть список статусов. RevOps этапы: 1=Новый, 2=Квалификация, "
                "3=Переговоры, 4=КП, 5=Подписание, 6=Закрыт(выиграш), 0=Закрыт(проигрыш)"
            ),
            "142": 6,
            "143": 0,
            "EXAMPLE_STATUS_ID_1": 1,
            "EXAMPLE_STATUS_ID_2": 2,
            "EXAMPLE_STATUS_ID_3": 3,
            "EXAMPLE_STATUS_ID_4": 4,
            "EXAMPLE_STATUS_ID_5": 5,
        }
        path.write_text(
            json.dumps(default, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )
        print(f"\n[ВНИМАНИЕ] Создан шаблон маппинга этапов: {path}")
        print("  Необходимо настроить маппинг под ваш AmoCRM.")
        print("  Запустите: python amocrm_connector.py --show-stages\n")

    raw = json.loads(path.read_text(encoding="utf-8"))
    mapping = {}
    for k, v in raw.items():
        if k.startswith("_"):
            continue
        try:
            mapping[int(k)] = int(v)
        except (ValueError, TypeError):
            pass
    return mapping


def _extract_custom_field(lead: dict, field_code: str) -> str:
    """Извлекает значение custom field по коду."""
    for cf in lead.get("custom_fields_values") or []:
        if cf.get("field_code") == field_code:
            vals = cf.get("values", [])
            if vals:
                return str(vals[0].get("value", ""))
    return ""


def map_lead_to_deal(
    lead: dict,
    contacts_map: dict,
    users_map: dict,
    pipelines_map: dict,
    stage_mapping: dict,
    snapshot_date: str,
) -> dict | None:
    """
    Трансформирует сделку AmoCRM → строку raw_deals.
    Возвращает None если stage_id не найден в маппинге (игнорируем).
    """
    status_id = lead.get("status_id", 0)
    stage_id = stage_mapping.get(status_id)

    if stage_id is None:
        # Статус не настроен в маппинге — пропускаем
        return None

    # Информация о статусе и пайплайне
    status_info = pipelines_map.get(status_id, {})
    stage_name = status_info.get("name") or REVOPS_STAGES.get(stage_id, "")
    pipeline_name = status_info.get("pipeline_name", "")

    # Клиент
    client_name = ""
    client_id = ""
    main_contact_id = lead.get("main_contact_id")
    if main_contact_id and main_contact_id in contacts_map:
        c = contacts_map[main_contact_id]
        client_name = c.get("name", "")
        client_id = str(c.get("id", ""))

    # Менеджер
    resp_user_id = lead.get("responsible_user_id", 0)
    manager_name = users_map.get(resp_user_id, f"user_{resp_user_id}")

    # Даты
    created_ts = lead.get("created_at", 0)
    updated_ts = lead.get("updated_at", 0)
    created_date = datetime.fromtimestamp(created_ts).strftime("%Y-%m-%d") if created_ts else ""
    updated_date = datetime.fromtimestamp(updated_ts).strftime("%Y-%m-%d") if updated_ts else ""

    # Deal velocity
    now_ts = time.time()
    days_in_stage = int((now_ts - updated_ts) / 86400) if updated_ts else 0
    cycle_days = int((now_ts - created_ts) / 86400) if created_ts else 0

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

    # UTM / источник
    utm_source = _extract_custom_field(lead, "UTM_SOURCE")
    utm_medium = _extract_custom_field(lead, "UTM_MEDIUM")
    utm_campaign = _extract_custom_field(lead, "UTM_CAMPAIGN")
    yclid = _extract_custom_field(lead, "YCLID")
    gclid = _extract_custom_field(lead, "GCLID")

    # Источник сделки (lead_source)
    lead_source = utm_source or _extract_custom_field(lead, "LEAD_SOURCE") or ""

    # Weighted value
    amount = lead.get("price") or 0
    weight = STAGE_WEIGHTS.get(stage_id, 0)
    weighted_val = round(amount * weight)

    # Теги
    tags_raw = lead.get("tags_values") or lead.get("_embedded", {}).get("tags", [])
    tags = ";".join(t.get("name", "") for t in tags_raw if t.get("name"))

    # Deal health score (простая формула)
    health = _calc_health_score(stage_id, days_in_stage, amount, cycle_days)

    # Причина проигрыша
    loss_code = ""
    loss_data = lead.get("_embedded", {}).get("loss_reason", [])
    if loss_data and isinstance(loss_data, list) and loss_data[0]:
        loss_code = str(loss_data[0].get("name", ""))

    return {
        "deal_id": str(lead["id"]),
        "client_name": client_name,
        "amount": amount,
        "stage_id": stage_id,
        "stage_name": stage_name,
        "created_date": created_date,
        "stage_changed_date": updated_date,
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
        "uploaded_by": "amocrm_connector",
        "upload_hash": str(lead["id"]),
        "next_action_type": "",
        "next_action_date": "",
        "next_action_owner": manager_name,
        "next_action_channel": "",
        "next_action_health": "OK" if health >= 50 else "РИСК",
        "deal_health_score": health,
    }


def _calc_health_score(stage_id: int, days_in_stage: int,
                       amount: float, cycle_days: int) -> int:
    """
    Простой скоринг сделки 0-100.
    100 = идеальная сделка, 0 = мёртвая.
    """
    score = 100
    # Штраф за зависание в этапе
    if days_in_stage > 30:
        score -= 40
    elif days_in_stage > 14:
        score -= 20
    elif days_in_stage > 7:
        score -= 10
    # Штраф за долгий цикл
    if cycle_days > 90:
        score -= 20
    elif cycle_days > 45:
        score -= 10
    # Бонус за поздние этапы (ближе к закрытию)
    if stage_id >= 5:
        score += 10
    elif stage_id <= 1:
        score -= 5
    return max(0, min(100, score))


# ─────────────────────────────────────────────
# Запись в Excel (UPSERT)
# ─────────────────────────────────────────────
def find_workbook_path() -> str:
    """Ищет Excel по маске от BASE_DIR."""
    pattern = str(BASE_DIR / XLSX_PATTERN)
    candidates = [
        f for f in glob.glob(pattern)
        if not os.path.basename(f).startswith("~$")
    ]
    if not candidates:
        print(f"[ОШИБКА] Не найден файл {XLSX_PATTERN} в {BASE_DIR}")
        sys.exit(1)
    wb_path = max(candidates, key=os.path.getmtime)
    return wb_path


def write_deals_to_excel(
    deals: list[dict],
    wb_path: str,
    dry_run: bool = False,
) -> SyncResult:
    """
    UPSERT сделок в raw_deals:
    - Если deal_id уже есть → обновляем строку
    - Если нового → добавляем в конец
    - Существующие строки без deal_id из AmoCRM — не трогаем
    """
    result = SyncResult(workbook_path=wb_path)
    print(f"\n  Открываем Excel: {wb_path}")

    wb = openpyxl.load_workbook(wb_path)
    ws = wb["raw_deals"]

    # Читаем текущий заголовок и индекс существующих deal_id
    header_row = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]

    # Определяем позиции нужных колонок в существующем файле
    col_index = {}
    for col_name in RAW_DEALS_HEADER:
        try:
            col_index[col_name] = header_row.index(col_name) + 1
        except ValueError:
            col_index[col_name] = None  # колонки нет — добавим в конец

    # deal_id колонка
    deal_id_col = col_index.get("deal_id", 1)

    # Строим индекс существующих сделок: {deal_id: row_number}
    existing = {}
    for row in range(2, ws.max_row + 1):
        cell_val = ws.cell(row, deal_id_col).value
        if cell_val is not None:
            existing[str(cell_val)] = row

    print(f"  Существующих строк в raw_deals: {ws.max_row - 1}")
    print(f"  Из AmoCRM к записи: {len(deals)}")

    if dry_run:
        print("\n  [DRY RUN] Изменения НЕ применяются.")
        for d in deals[:5]:
            print(f"    deal_id={d['deal_id']} stage={d['stage_id']} "
                  f"amount={d['amount']:,} manager={d['manager_id']}")
        if len(deals) > 5:
            print(f"    ... и ещё {len(deals)-5} сделок")
        result.total_fetched = len(deals)
        return result

    # Стиль для строк AmoCRM (лёгкая заливка для отличия от ручных)
    amo_fill = PatternFill(start_color="F0F7FF", end_color="F0F7FF", fill_type="solid")
    won_fill = PatternFill(start_color="E6F4EA", end_color="E6F4EA", fill_type="solid")
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
        row = existing.get(deal_id, None)

        if row is None:
            # Новая сделка — добавляем в первую свободную строку
            row = next_row
            next_row += 1
            result.added += 1
        else:
            result.updated += 1

        # Записываем значения по колонкам
        for col_name, value in deal.items():
            col_num = col_index.get(col_name)
            if col_num is None:
                continue
            cell = ws.cell(row, col_num)
            cell.value = value

            # Стилизация по этапу
            stage = deal.get("stage_id", 0)
            if stage == 6:
                cell.fill = won_fill
            elif stage == 0:
                cell.fill = lost_fill
            else:
                cell.fill = amo_fill

    result.total_fetched = len(deals)
    wb.save(wb_path)
    wb.close()
    print(f"  Excel сохранён: {wb_path}")
    return result


# ─────────────────────────────────────────────
# Вспомогательные команды CLI
# ─────────────────────────────────────────────
def cmd_show_stages(creds: AmoCreds) -> None:
    """Показывает все статусы пайплайна — для настройки маппинга."""
    client = AmoCRMClient(creds)
    print(f"\n  Пайплайны и статусы в {creds.domain}:\n")
    pipelines_map = fetch_pipelines(client)

    pipeline_groups: dict[str, list] = {}
    for status_id, info in sorted(pipelines_map.items()):
        p_name = info["pipeline_name"]
        pipeline_groups.setdefault(p_name, []).append((status_id, info))

    for p_name, statuses in pipeline_groups.items():
        print(f"  📊 Пайплайн: {p_name}")
        for status_id, info in statuses:
            type_label = ""
            if info.get("type") == 142:
                type_label = " ← ПОБЕДА (используй stage_id=6)"
            elif info.get("type") == 143:
                type_label = " ← ПРОИГРЫШ (используй stage_id=0)"
            print(f"      status_id={status_id:>8}  [{info['name']}]{type_label}")
        print()

    print(f"  Добавьте нужные status_id в: {STAGE_MAPPING_FILE}")


def cmd_setup() -> None:
    """Интерактивная настройка учётных данных AmoCRM."""
    print("\n" + "=" * 60)
    print("  Настройка AmoCRM коннектора")
    print("=" * 60)
    print()
    print("1. Войдите в AmoCRM → Настройки → Интеграции")
    print("2. Нажмите '+ Создать интеграцию' → выберите 'Внешняя интеграция'")
    print("3. Укажите Redirect URI: https://example.com/oauth")
    print("4. Сохраните client_id и client_secret")
    print()

    domain = input("Домен AmoCRM (например: mycompany.amocrm.ru): ").strip()
    client_id = input("Client ID: ").strip()
    client_secret = input("Client Secret: ").strip()

    # Получение authorization code через браузер
    auth_url = (
        f"https://{domain}/oauth?client_id={client_id}"
        f"&state=revops&mode=popup"
    )
    print(f"\nОткрываем браузер для авторизации...")
    print(f"URL: {auth_url}")
    try:
        webbrowser.open(auth_url)
    except Exception:
        pass

    print("\nПосле авторизации скопируйте code= из URL редиректа:")
    auth_code = input("Authorization code: ").strip()

    # Обмен кода на токены
    token_url = AMO_TOKEN_URL.format(domain=domain)
    payload = {
        "client_id": client_id,
        "client_secret": client_secret,
        "grant_type": "authorization_code",
        "code": auth_code,
        "redirect_uri": "https://example.com/oauth",
    }

    print("\nПолучаем токены...")
    resp = requests.post(token_url, json=payload, timeout=30)
    if resp.status_code != 200:
        print(f"[ОШИБКА] {resp.status_code}: {resp.text}")
        sys.exit(1)

    data = resp.json()
    creds = AmoCreds(
        domain=domain,
        client_id=client_id,
        client_secret=client_secret,
        access_token=data["access_token"],
        refresh_token=data["refresh_token"],
        token_expires_at=time.time() + data.get("expires_in", 86400),
    )
    creds.save(CREDENTIALS_FILE)
    print(f"\n✓ Учётные данные сохранены в: {CREDENTIALS_FILE}")
    print(f"\nТеперь запустите: python amocrm_connector.py --show-stages")
    print("Чтобы настроить маппинг этапов под ваш пайплайн.")


# ─────────────────────────────────────────────
# Основная функция синхронизации
# ─────────────────────────────────────────────
def sync(
    wb_path: str | None = None,
    days: int | None = None,
    dry_run: bool = False,
) -> SyncResult:
    """
    Полная синхронизация AmoCRM → raw_deals.

    Args:
        wb_path: Путь к Excel. Если None — ищет по маске.
        days:    Только сделки за последние N дней. None = все.
        dry_run: Не писать в Excel, только показать что будет.

    Returns:
        SyncResult с метриками синхронизации.
    """
    start_time = time.time()
    snapshot_date = datetime.now().strftime("%Y-%m-%d")

    # 1. Учётные данные
    creds = AmoCreds.load(CREDENTIALS_FILE)
    client = AmoCRMClient(creds)

    # 2. Маппинг этапов
    stage_mapping = load_stage_mapping(STAGE_MAPPING_FILE)
    if not any(v > 0 for v in stage_mapping.values()):
        print("[ВНИМАНИЕ] Маппинг этапов содержит только дефолтные значения.")
        print(f"  Настройте: {STAGE_MAPPING_FILE}")

    # 3. Справочники
    print(f"\nЗагрузка данных из {creds.domain}...")
    print("  Пайплайны и статусы...")
    pipelines_map = fetch_pipelines(client)

    print("  Пользователи...")
    users_map = fetch_users(client)
    print(f"  Найдено менеджеров: {len(users_map)}")

    print("  Контакты...")
    contacts_map = fetch_contacts(client)
    print(f"  Найдено контактов: {len(contacts_map)}")

    # 4. Сделки
    print(f"  Сделки (период: {'все' if not days else f'последние {days} дн.'})...")
    leads = fetch_leads(client, days=days)
    print(f"  Загружено сделок: {len(leads)}")

    # 5. Трансформация
    deals = []
    skipped = 0
    for lead in leads:
        deal = map_lead_to_deal(
            lead, contacts_map, users_map, pipelines_map,
            stage_mapping, snapshot_date
        )
        if deal is None:
            skipped += 1
            continue
        if deal["stage_id"] == 0:  # проигрыш — пропускаем в активный пайплайн
            skipped += 1
            continue
        deals.append(deal)

    print(f"  После маппинга: {len(deals)} сделок в пайплайне "
          f"({skipped} пропущено/проигрыш)")

    # 6. Запись в Excel
    if wb_path is None:
        wb_path = find_workbook_path()

    result = write_deals_to_excel(deals, wb_path, dry_run=dry_run)
    result.skipped = skipped
    result.duration_sec = round(time.time() - start_time, 1)

    # 7. Аудит
    if not dry_run and _HAS_AUDIT:
        log_generation(
            status="success",
            duration_sec=result.duration_sec,
            role="sync",
            report_type="AmoCRM_Sync",
            pdf_size_kb=0,
        )

    return result


# ─────────────────────────────────────────────
# CLI точка входа
# ─────────────────────────────────────────────
def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(
        description="RevOps AmoCRM Connector — синхронизация сделок в Excel",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры:
  python amocrm_connector.py                 — полная синхронизация
  python amocrm_connector.py --days 30       — только за 30 дней
  python amocrm_connector.py --dry-run       — предпросмотр без записи
  python amocrm_connector.py --setup         — первоначальная настройка
  python amocrm_connector.py --show-stages   — список статусов пайплайна
        """
    )
    parser.add_argument("--setup", action="store_true",
                        help="Интерактивная настройка токенов AmoCRM")
    parser.add_argument("--show-stages", action="store_true",
                        help="Показать статусы пайплайна для настройки маппинга")
    parser.add_argument("--days", type=int, default=None,
                        help="Синхронизировать только за последние N дней")
    parser.add_argument("--dry-run", action="store_true",
                        help="Показать что будет синхронизировано, не записывать")
    parser.add_argument("--workbook", type=str, default=None,
                        help="Явный путь к Excel-файлу")

    args = parser.parse_args()

    if args.setup:
        cmd_setup()
        return 0

    if args.show_stages:
        creds = AmoCreds.load(CREDENTIALS_FILE)
        cmd_show_stages(creds)
        return 0

    print("=" * 60)
    print("  RevOps AmoCRM Connector V17.6")
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
                report_type="AmoCRM_Sync",
                error_message=str(e),
            )
        raise


if __name__ == "__main__":
    sys.exit(main())
