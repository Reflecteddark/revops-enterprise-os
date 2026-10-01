"""
Тесты Bitrix24 коннектора — работают без реального Bitrix24.
Тестируем: трансформацию, маппинг этапов, UPSERT, fallback-логику.
"""
import json
import pytest
import openpyxl
from pathlib import Path

from bitrix24_connector import (
    B24Creds,
    SyncResult,
    map_deal_to_row,
    load_stage_mapping,
    write_deals_to_excel,
    resolve_stage,
    _calc_health_score,
    _parse_b24_date,
    _calc_days_since,
    RAW_DEALS_HEADER,
    STAGE_WEIGHTS,
)


# ─────────────────────────────────────────────
# Фикстуры
# ─────────────────────────────────────────────

@pytest.fixture
def stage_mapping():
    return {
        "NEW": 1,
        "IN_PROCESS": 2,
        "PREPARATION": 3,
        "PROPOSAL": 4,
        "CONTRACT": 5,
        "WON": 6,
        "LOSE": 0,
        "C1:NEW": 1,
        "C1:WON": 6,
    }


@pytest.fixture
def stages_map():
    return {
        "NEW":       {"name": "Новый", "pipeline_name": "Основная", "category_id": 0, "semantic": "P"},
        "IN_PROCESS":{"name": "В работе", "pipeline_name": "Основная", "category_id": 0, "semantic": "P"},
        "WON":       {"name": "Успешно", "pipeline_name": "Основная", "category_id": 0, "semantic": "S"},
        "LOSE":      {"name": "Провал",  "pipeline_name": "Основная", "category_id": 0, "semantic": "F"},
        "C1:NEW":    {"name": "Входящий", "pipeline_name": "VIP-воронка", "category_id": 1, "semantic": "P"},
        "C1:WON":    {"name": "Закрыт", "pipeline_name": "VIP-воронка", "category_id": 1, "semantic": "S"},
    }


@pytest.fixture
def users_map():
    return {1: "Иванов А.В.", 2: "Петрова К.М.", 3: "Сидоров Н."}


@pytest.fixture
def contacts_map():
    return {
        10: {"id": 10, "name": "Алексеев Петр"},
        20: {"id": 20, "name": "ООО Бета"},
    }


@pytest.fixture
def companies_map():
    return {100: "ООО Альфа", 200: "АО Гамма"}


@pytest.fixture
def sample_deal():
    """Стандартная сделка Bitrix24."""
    return {
        "ID": "7777",
        "TITLE": "Тестовая сделка",
        "OPPORTUNITY": "850000",
        "STAGE_ID": "IN_PROCESS",
        "STAGE_SEMANTIC_ID": "P",
        "ASSIGNED_BY_ID": "1",
        "CONTACT_ID": "10",
        "COMPANY_ID": "100",
        "CATEGORY_ID": "0",
        "DATE_CREATE": "2026-08-01T10:00:00+03:00",
        "DATE_MODIFY": "2026-09-20T15:30:00+03:00",
        "UTM_SOURCE": "yandex",
        "UTM_MEDIUM": "cpc",
        "UTM_CAMPAIGN": "revops_promo",
        "YCLID": "12345",
        "GCLID": "",
        "SOURCE_ID": "WEB",
        "SOURCE_DESCRIPTION": "",
    }


@pytest.fixture
def excel_with_raw_deals(tmp_path):
    """Excel с листом raw_deals по схеме продукта."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "raw_deals"
    for col, name in enumerate(RAW_DEALS_HEADER, start=1):
        ws.cell(1, col).value = name
    # Существующая ручная строка
    ws.cell(2, 1).value = "MANUAL-999"
    ws.cell(2, 3).value = 250_000
    ws.cell(2, 4).value = 3
    path = tmp_path / "RevOps Platform V17.6 B24Test.xlsx"
    wb.save(path)
    wb.close()
    return str(path)


# ─────────────────────────────────────────────
# Тесты B24Creds
# ─────────────────────────────────────────────

def test_b24creds_domain_extracted():
    """Домен автоматически извлекается из webhook URL."""
    creds = B24Creds(webhook_url="https://mycompany.bitrix24.ru/rest/1/token123/")
    assert creds.domain == "mycompany.bitrix24.ru"


def test_b24creds_url_method():
    """Метод корректно добавляется к webhook URL."""
    creds = B24Creds(webhook_url="https://example.bitrix24.ru/rest/1/abc/")
    url = creds.url("crm.deal.list")
    assert url == "https://example.bitrix24.ru/rest/1/abc/crm.deal.list"


def test_b24creds_save_load(tmp_path):
    """Сохранение и загрузка credentials."""
    path = tmp_path / "test_creds.json"
    creds = B24Creds(webhook_url="https://test.bitrix24.ru/rest/5/xyz/")
    creds.save(path)
    loaded = B24Creds.load(path)
    assert loaded.webhook_url == creds.webhook_url
    assert loaded.domain == "test.bitrix24.ru"


# ─────────────────────────────────────────────
# Тесты resolve_stage
# ─────────────────────────────────────────────

def test_resolve_stage_exact_mapping(stage_mapping):
    """Точный маппинг по stage_id работает."""
    assert resolve_stage("WON", "S", stage_mapping) == 6
    assert resolve_stage("LOSE", "F", stage_mapping) == 0
    assert resolve_stage("NEW", "P", stage_mapping) == 1


def test_resolve_stage_semantic_fallback(stage_mapping):
    """Если stage_id нет в маппинге — используется semantic."""
    # Неизвестный этап с semantic=S → WON
    assert resolve_stage("UNKNOWN_WON_STAGE", "S", stage_mapping) == 6
    # Неизвестный этап с semantic=F → LOSE
    assert resolve_stage("UNKNOWN_LOSE_STAGE", "F", stage_mapping) == 0


def test_resolve_stage_active_default(stage_mapping):
    """Активный этап без маппинга → 3 (Переговоры по умолчанию)."""
    assert resolve_stage("SOME_CUSTOM_STAGE", "P", stage_mapping) == 3


def test_resolve_stage_category_format(stage_mapping):
    """Этапы с категорией (C1:STAGE) маппятся корректно."""
    assert resolve_stage("C1:NEW", "P", stage_mapping) == 1
    assert resolve_stage("C1:WON", "S", stage_mapping) == 6


# ─────────────────────────────────────────────
# Тесты трансформации сделок
# ─────────────────────────────────────────────

def test_map_deal_basic_fields(sample_deal, contacts_map, companies_map,
                                users_map, stages_map, stage_mapping):
    """Базовые поля трансформируются корректно."""
    row = map_deal_to_row(
        sample_deal, contacts_map, companies_map, users_map,
        stages_map, stage_mapping, "2026-09-30"
    )
    assert row is not None
    assert row["deal_id"] == "7777"
    assert row["amount"] == 850_000.0
    assert row["stage_id"] == 2  # IN_PROCESS → 2
    assert row["client_name"] == "Алексеев Петр"
    assert row["manager_id"] == "Иванов А.В."


def test_map_deal_utm_fields(sample_deal, contacts_map, companies_map,
                              users_map, stages_map, stage_mapping):
    """UTM-метки (встроенные в Bitrix24) корректно извлекаются."""
    row = map_deal_to_row(
        sample_deal, contacts_map, companies_map, users_map,
        stages_map, stage_mapping, "2026-09-30"
    )
    assert row["utm_source"] == "yandex"
    assert row["utm_medium"] == "cpc"
    assert row["utm_campaign"] == "revops_promo"
    assert row["yclid"] == "12345"


def test_map_deal_won_flags(sample_deal, contacts_map, companies_map,
                              users_map, stages_map, stage_mapping):
    """WON сделка: is_won=1, is_lost=0, stage_id=6."""
    sample_deal["STAGE_ID"] = "WON"
    sample_deal["STAGE_SEMANTIC_ID"] = "S"
    row = map_deal_to_row(
        sample_deal, contacts_map, companies_map, users_map,
        stages_map, stage_mapping, "2026-09-30"
    )
    assert row["stage_id"] == 6
    assert row["is_won"] == 1
    assert row["is_lost"] == 0


def test_map_deal_lose_flags(sample_deal, contacts_map, companies_map,
                               users_map, stages_map, stage_mapping):
    """LOSE сделка: is_lost=1, stage_id=0."""
    sample_deal["STAGE_ID"] = "LOSE"
    sample_deal["STAGE_SEMANTIC_ID"] = "F"
    row = map_deal_to_row(
        sample_deal, contacts_map, companies_map, users_map,
        stages_map, stage_mapping, "2026-09-30"
    )
    assert row["stage_id"] == 0
    assert row["is_lost"] == 1
    assert row["is_won"] == 0


def test_map_deal_company_fallback(sample_deal, contacts_map, companies_map,
                                    users_map, stages_map, stage_mapping):
    """Если контакт не найден — используется название компании."""
    sample_deal["CONTACT_ID"] = "9999"  # не существует
    row = map_deal_to_row(
        sample_deal, contacts_map, companies_map, users_map,
        stages_map, stage_mapping, "2026-09-30"
    )
    assert row["client_name"] == "ООО Альфа"  # из companies_map


def test_map_deal_scoring_tiers(sample_deal, contacts_map, companies_map,
                                  users_map, stages_map, stage_mapping):
    """Scoring tier: A > 1М, B > 300k, C остальные."""
    # 850k → B
    row = map_deal_to_row(
        sample_deal, contacts_map, companies_map, users_map,
        stages_map, stage_mapping, "2026-09-30"
    )
    assert row["scoring_tier"] == "B"

    sample_deal["OPPORTUNITY"] = "1_500_000"
    row_a = map_deal_to_row(
        sample_deal, contacts_map, companies_map, users_map,
        stages_map, stage_mapping, "2026-09-30"
    )
    # "1_500_000" парсится как 1_500_000 = 1500000 в Python float → A
    # (Python float() поддерживает underscore начиная с 3.6)
    assert row_a["scoring_tier"] == "A"

    sample_deal["OPPORTUNITY"] = "50000"
    row_c = map_deal_to_row(
        sample_deal, contacts_map, companies_map, users_map,
        stages_map, stage_mapping, "2026-09-30"
    )
    assert row_c["scoring_tier"] == "C"


def test_map_deal_weighted_val(sample_deal, contacts_map, companies_map,
                                 users_map, stages_map, stage_mapping):
    """weighted_val = amount × weight для этапа 2 (0.15)."""
    row = map_deal_to_row(
        sample_deal, contacts_map, companies_map, users_map,
        stages_map, stage_mapping, "2026-09-30"
    )
    expected = round(850_000 * STAGE_WEIGHTS[2])
    assert row["weighted_val"] == expected


# ─────────────────────────────────────────────
# Тесты вспомогательных функций
# ─────────────────────────────────────────────

def test_parse_b24_date_iso():
    """Парсинг ISO-даты Bitrix24."""
    assert _parse_b24_date("2026-09-15T14:30:00+03:00") == "2026-09-15"


def test_parse_b24_date_empty():
    assert _parse_b24_date("") == ""
    assert _parse_b24_date(None) == ""


def test_calc_days_since_recent():
    """Недавняя дата → небольшое количество дней."""
    from datetime import datetime, timedelta
    recent = (datetime.now() - timedelta(days=5)).strftime("%Y-%m-%d")
    assert _calc_days_since(recent) == 5


def test_calc_health_score_bounds():
    """Health score всегда 0-100."""
    for stage in range(0, 7):
        for days in [0, 7, 14, 30, 60, 120]:
            score = _calc_health_score(stage, days, 100_000, days)
            assert 0 <= score <= 100


# ─────────────────────────────────────────────
# Тесты маппинга
# ─────────────────────────────────────────────

def test_load_stage_mapping_creates_template(tmp_path):
    """Если файла нет — создаётся шаблон с WON/LOSE."""
    path = tmp_path / "b24_mapping.json"
    mapping = load_stage_mapping(path)
    assert path.exists()
    assert mapping.get("WON") == 6
    assert mapping.get("LOSE") == 0
    assert mapping.get("NEW") == 1


def test_load_stage_mapping_custom(tmp_path):
    """Кастомный маппинг загружается корректно."""
    path = tmp_path / "b24_mapping.json"
    custom = {
        "WON": 6, "LOSE": 0,
        "C2:QUALIFIED": 2,
        "C2:PROPOSAL": 4,
        "C2:WON": 6,
    }
    path.write_text(json.dumps(custom), encoding="utf-8")
    mapping = load_stage_mapping(path)
    assert mapping["C2:QUALIFIED"] == 2
    assert mapping["C2:PROPOSAL"] == 4


def test_load_stage_mapping_ignores_comments(tmp_path):
    """Ключи начинающиеся с _ игнорируются."""
    path = tmp_path / "b24_mapping.json"
    data = {"_comment": "ignore me", "WON": 6, "LOSE": 0}
    path.write_text(json.dumps(data), encoding="utf-8")
    mapping = load_stage_mapping(path)
    assert "_comment" not in mapping
    assert "WON" in mapping


# ─────────────────────────────────────────────
# Тесты UPSERT в Excel
# ─────────────────────────────────────────────

def _make_deal_row(deal_id: str = "B24-001", stage_id: int = 2,
                   amount: float = 500_000) -> dict:
    """Фабрика тестовых строк raw_deals."""
    return {
        "deal_id": deal_id,
        "client_name": "Тест Клиент",
        "amount": amount,
        "stage_id": stage_id,
        "stage_name": "В работе",
        "created_date": "2026-08-01",
        "stage_changed_date": "2026-09-20",
        "manager_id": "Иванов А.В.",
        "segment_code": "Основная",
        "lead_source": "WEB",
        "client_id": "10",
        "scoring_tier": "B",
        "loss_code": "",
        "synced_at": "2026-09-30 12:00",
        "is_won": 0, "is_lost": 0,
        "days_in_stage": 10,
        "aging_cohort": "Aging",
        "weighted_val": round(amount * STAGE_WEIGHTS.get(stage_id, 0)),
        "cycle_days": 60,
        "utm_source": "yandex", "utm_medium": "cpc",
        "utm_campaign": "", "yclid": "", "gclid": "",
        "data_snapshot_date": "2026-09-30",
        "data_version": "17.6",
        "uploaded_by": "bitrix24_connector",
        "upload_hash": deal_id,
        "next_action_type": "", "next_action_date": "",
        "next_action_owner": "Иванов А.В.",
        "next_action_channel": "",
        "next_action_health": "OK",
        "deal_health_score": 70,
    }


def test_write_adds_new_deal(excel_with_raw_deals):
    """Новая сделка добавляется в конец листа."""
    deal = _make_deal_row("B24-001", amount=600_000)
    result = write_deals_to_excel([deal], excel_with_raw_deals, dry_run=False)

    assert result.added == 1
    assert result.updated == 0

    wb = openpyxl.load_workbook(excel_with_raw_deals)
    ws = wb["raw_deals"]
    assert ws.cell(3, 1).value == "B24-001"
    assert ws.cell(3, 3).value == 600_000
    wb.close()


def test_write_updates_existing_deal(excel_with_raw_deals):
    """Существующая сделка обновляется (не дублируется)."""
    deal = _make_deal_row("B24-002", amount=700_000)
    write_deals_to_excel([deal], excel_with_raw_deals)

    # Обновляем сумму
    deal["amount"] = 900_000
    result = write_deals_to_excel([deal], excel_with_raw_deals)

    assert result.updated == 1
    assert result.added == 0

    wb = openpyxl.load_workbook(excel_with_raw_deals)
    ws = wb["raw_deals"]
    amount_col = RAW_DEALS_HEADER.index("amount") + 1
    # Строка 3 — B24-002
    assert ws.cell(3, amount_col).value == 900_000
    wb.close()


def test_dry_run_no_changes(excel_with_raw_deals):
    """Dry run не изменяет файл."""
    wb_before = openpyxl.load_workbook(excel_with_raw_deals)
    rows_before = wb_before["raw_deals"].max_row
    wb_before.close()

    deal = _make_deal_row("B24-DRY")
    write_deals_to_excel([deal], excel_with_raw_deals, dry_run=True)

    wb_after = openpyxl.load_workbook(excel_with_raw_deals)
    assert wb_after["raw_deals"].max_row == rows_before
    wb_after.close()


def test_manual_row_preserved(excel_with_raw_deals):
    """Ручные строки (не Bitrix24) не затираются."""
    deal = _make_deal_row("B24-NEW")
    write_deals_to_excel([deal], excel_with_raw_deals)

    wb = openpyxl.load_workbook(excel_with_raw_deals)
    ws = wb["raw_deals"]
    assert ws.cell(2, 1).value == "MANUAL-999"
    assert ws.cell(2, 3).value == 250_000
    wb.close()


def test_multiple_deals_all_added(excel_with_raw_deals):
    """Несколько новых сделок — все добавляются."""
    deals = [_make_deal_row(f"B24-{i:03d}", amount=i * 100_000)
             for i in range(1, 6)]
    result = write_deals_to_excel(deals, excel_with_raw_deals)
    assert result.added == 5
    assert result.updated == 0

    wb = openpyxl.load_workbook(excel_with_raw_deals)
    ws = wb["raw_deals"]
    # Строки 2 (ручная) + 5 новых = max_row 6
    assert ws.max_row == 7
    wb.close()


# ─────────────────────────────────────────────
# Тесты RateLimiter и защиты от лимитов Bitrix24
# ─────────────────────────────────────────────
import time
from unittest.mock import MagicMock
from bitrix24_connector import (
    RateLimiter,
    normalize_phone,
    extract_contact_phones,
    resolve_target_deal,
    find_deals_by_phone,
    match_call_to_deal,
    Bitrix24Client,
)


def test_b24_rate_limiter_pacing():
    limiter = RateLimiter(max_per_sec=10.0)
    t0 = time.time()
    limiter.wait()
    limiter.wait()
    t1 = time.time()
    assert (t1 - t0) >= 0.08


def test_b24_call_batch_mock():
    creds = B24Creds(webhook_url="https://test.bitrix24.ru/rest/1/abc/")
    client = Bitrix24Client(creds)
    client.call = MagicMock(return_value={"result": {"result": {"cmd1": {"ID": 10}, "cmd2": {"ID": 20}}}})

    cmds = {"cmd1": "crm.deal.get?id=10", "cmd2": "crm.deal.get?id=20"}
    res = client.call_batch(cmds)
    assert "cmd1" in res
    assert res["cmd1"]["ID"] == 10
    client.call.assert_called_once_with("batch", {"halt": 0, "cmd": cmds})


def test_b24_call_batch_limit_validation():
    creds = B24Creds(webhook_url="https://test.bitrix24.ru/rest/1/abc/")
    client = Bitrix24Client(creds)
    cmds = {f"cmd_{i}": f"crm.deal.get?id={i}" for i in range(51)}
    with pytest.raises(ValueError, match="максимум 50 команд"):
        client.call_batch(cmds)


# ─────────────────────────────────────────────
# Тесты нормализации телефонов и извлечения
# ─────────────────────────────────────────────
@pytest.mark.parametrize("raw,expected", [
    ("8 (800) 555-35-35", "+78005553535"),
    ("+7 495 123-45-67", "+74951234567"),
    ("9161234567", "+79161234567"),
    ("", ""),
    (None, ""),
])
def test_b24_normalize_phone(raw, expected):
    assert normalize_phone(raw) == expected


def test_b24_extract_contact_phones():
    contact = {
        "ID": "10",
        "NAME": "Иван",
        "PHONE": [
            {"VALUE": "+7 (495) 111-22-33", "VALUE_TYPE": "WORK"},
            {"VALUE": "89162223344", "VALUE_TYPE": "MOBILE"},
            {"VALUE": "+74951112233"}, # дубль
        ],
    }
    phones = extract_contact_phones(contact)
    assert len(phones) == 2
    assert "+74951112233" in phones
    assert "+79162223344" in phones


# ─────────────────────────────────────────────
# Тесты SmartDealMatcher (Защита от псевдо-дублей в Bitrix24)
# ─────────────────────────────────────────────
def test_b24_resolve_target_deal_prefers_active():
    """Открытая сделка (semantic P) всегда приоритетнее выигранных (S) или проигранных (F)."""
    deals = [
        {"ID": "1", "STAGE_SEMANTIC_ID": "S", "OPPORTUNITY": "5000000", "DATE_MODIFY": "2026-09-01"},
        {"ID": "2", "STAGE_SEMANTIC_ID": "F", "OPPORTUNITY": "3000000", "DATE_MODIFY": "2026-09-05"},
        {"ID": "3", "STAGE_SEMANTIC_ID": "P", "OPPORTUNITY": "450000",  "DATE_MODIFY": "2026-08-01"},
    ]
    chosen = resolve_target_deal(deals)
    assert chosen["ID"] == "3"


def test_b24_resolve_target_deal_prefers_higher_amount():
    """При нескольких открытых сделках выбирается сделка с наибольшей суммой."""
    deals = [
        {"ID": "10", "STAGE_SEMANTIC_ID": "P", "OPPORTUNITY": "250000", "DATE_MODIFY": "2026-09-10"},
        {"ID": "11", "STAGE_SEMANTIC_ID": "P", "OPPORTUNITY": "1800000", "DATE_MODIFY": "2026-09-01"},
        {"ID": "12", "STAGE_SEMANTIC_ID": "P", "OPPORTUNITY": "90000",  "DATE_MODIFY": "2026-09-15"},
    ]
    chosen = resolve_target_deal(deals)
    assert chosen["ID"] == "11"


def test_b24_resolve_target_deal_prefers_latest_modify_on_tie():
    """При равенстве сумм выбирается более свежая сделка."""
    deals = [
        {"ID": "20", "STAGE_SEMANTIC_ID": "P", "OPPORTUNITY": "500000", "DATE_MODIFY": "2026-09-01T10:00:00"},
        {"ID": "21", "STAGE_SEMANTIC_ID": "P", "OPPORTUNITY": "500000", "DATE_MODIFY": "2026-09-20T12:00:00"},
    ]
    chosen = resolve_target_deal(deals)
    assert chosen["ID"] == "21"


def test_b24_match_call_to_deal_end_to_end():
    """Сквозной матчинг звонка к сделке Bitrix24."""
    contacts_map = {
        "77": {
            "ID": "77",
            "NAME": "Клиент",
            "PHONE": [{"VALUE": "+7 (916) 555-44-33"}],
        }
    }
    all_deals = [
        {"ID": "501", "CONTACT_ID": "77", "STAGE_SEMANTIC_ID": "F", "OPPORTUNITY": "100000"},
        {"ID": "502", "CONTACT_ID": "77", "STAGE_SEMANTIC_ID": "P", "OPPORTUNITY": "950000"},
    ]
    res = match_call_to_deal("89165554433", all_deals, contacts_map)
    assert res is not None
    assert res["ID"] == "502"

