"""
Тесты AmoCRM коннектора — работают без реального AmoCRM.
Тестируем трансформацию данных, маппинг, UPSERT-логику.
"""
import json
import time
import pytest
import openpyxl
from pathlib import Path
from unittest.mock import patch, MagicMock

from amocrm_connector import (
    AmoCreds,
    SyncResult,
    map_lead_to_deal,
    load_stage_mapping,
    write_deals_to_excel,
    _calc_health_score,
    _extract_custom_field,
    RAW_DEALS_HEADER,
    STAGE_WEIGHTS,
)


# ─────────────────────────────────────────────
# Фикстуры
# ─────────────────────────────────────────────

@pytest.fixture
def stage_mapping():
    """Простой маппинг для тестов: 3 статуса → этапы 1, 3, 6."""
    return {101: 1, 303: 3, 142: 6, 143: 0}


@pytest.fixture
def pipelines_map():
    return {
        101: {"name": "Входящий", "pipeline_name": "Основной", "pipeline_id": 1, "type": 0},
        303: {"name": "Переговоры", "pipeline_name": "Основной", "pipeline_id": 1, "type": 0},
        142: {"name": "Успешно реализовано", "pipeline_name": "Основной", "pipeline_id": 1, "type": 142},
        143: {"name": "Закрыто и не реализовано", "pipeline_name": "Основной", "pipeline_id": 1, "type": 143},
    }


@pytest.fixture
def users_map():
    return {7: "Иванов А.", 8: "Петров К.", 9: "Сидорова М."}


@pytest.fixture
def contacts_map():
    return {
        1001: {"id": 1001, "name": "ООО Альфа"},
        1002: {"id": 1002, "name": "ИП Бетов"},
    }


@pytest.fixture
def sample_lead():
    """Минимальная структура сделки AmoCRM."""
    return {
        "id": 5555,
        "name": "Сделка с ООО Альфа",
        "price": 750_000,
        "status_id": 303,
        "responsible_user_id": 7,
        "main_contact_id": 1001,
        "created_at": int(time.time()) - 86400 * 10,
        "updated_at": int(time.time()) - 86400 * 3,
        "custom_fields_values": [
            {"field_code": "UTM_SOURCE", "values": [{"value": "google"}]},
            {"field_code": "UTM_MEDIUM", "values": [{"value": "cpc"}]},
        ],
        "tags_values": [{"name": "B2B"}, {"name": "Крупный"}],
        "_embedded": {},
    }


@pytest.fixture
def excel_with_raw_deals(tmp_path):
    """Создаёт Excel с листом raw_deals по схеме продукта."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "raw_deals"
    # Заголовок строго по схеме
    for col, name in enumerate(RAW_DEALS_HEADER, start=1):
        ws.cell(1, col).value = name
    # Одна существующая строка (ручная, не из AmoCRM)
    ws.cell(2, 1).value = "MANUAL-001"
    ws.cell(2, 3).value = 100_000
    ws.cell(2, 4).value = 2

    path = tmp_path / "RevOps Platform V17.6 Test.xlsx"
    wb.save(path)
    wb.close()
    return str(path)


# ─────────────────────────────────────────────
# Тесты трансформации
# ─────────────────────────────────────────────

def test_map_lead_basic(sample_lead, contacts_map, users_map,
                        pipelines_map, stage_mapping):
    """Базовая трансформация: поля заполнены корректно."""
    deal = map_lead_to_deal(
        sample_lead, contacts_map, users_map, pipelines_map,
        stage_mapping, "2026-09-30"
    )
    assert deal is not None
    assert deal["deal_id"] == "5555"
    assert deal["amount"] == 750_000
    assert deal["stage_id"] == 3
    assert deal["client_name"] == "ООО Альфа"
    assert deal["manager_id"] == "Иванов А."


def test_map_lead_unknown_status_returns_none(sample_lead, contacts_map,
                                              users_map, pipelines_map, stage_mapping):
    """Статус не в маппинге → None (игнорируем сделку)."""
    sample_lead["status_id"] = 9999  # не в маппинге
    deal = map_lead_to_deal(
        sample_lead, contacts_map, users_map, pipelines_map,
        stage_mapping, "2026-09-30"
    )
    assert deal is None


def test_map_lead_closed_won(sample_lead, contacts_map, users_map,
                              pipelines_map, stage_mapping):
    """Статус 142 → stage_id=6, is_won=1."""
    sample_lead["status_id"] = 142
    deal = map_lead_to_deal(
        sample_lead, contacts_map, users_map, pipelines_map,
        stage_mapping, "2026-09-30"
    )
    assert deal is not None
    assert deal["stage_id"] == 6
    assert deal["is_won"] == 1
    assert deal["is_lost"] == 0


def test_map_lead_utm_extraction(sample_lead, contacts_map, users_map,
                                  pipelines_map, stage_mapping):
    """UTM-метки корректно извлекаются из custom_fields_values."""
    deal = map_lead_to_deal(
        sample_lead, contacts_map, users_map, pipelines_map,
        stage_mapping, "2026-09-30"
    )
    assert deal["utm_source"] == "google"
    assert deal["utm_medium"] == "cpc"


def test_map_lead_weighted_val(sample_lead, contacts_map, users_map,
                                pipelines_map, stage_mapping):
    """weighted_val = amount × weight этапа."""
    deal = map_lead_to_deal(
        sample_lead, contacts_map, users_map, pipelines_map,
        stage_mapping, "2026-09-30"
    )
    expected_weight = STAGE_WEIGHTS[3]   # этап 3 = 0.30
    expected_val = round(750_000 * expected_weight)
    assert deal["weighted_val"] == expected_val


def test_map_lead_scoring_tier(sample_lead, contacts_map, users_map,
                                 pipelines_map, stage_mapping):
    """Scoring tier: A > 1М, B > 300k, C остальные."""
    # 750k → B
    deal = map_lead_to_deal(
        sample_lead, contacts_map, users_map, pipelines_map,
        stage_mapping, "2026-09-30"
    )
    assert deal["scoring_tier"] == "B"

    # Меняем на A
    sample_lead["price"] = 2_000_000
    deal_a = map_lead_to_deal(
        sample_lead, contacts_map, users_map, pipelines_map,
        stage_mapping, "2026-09-30"
    )
    assert deal_a["scoring_tier"] == "A"


# ─────────────────────────────────────────────
# Тесты health score
# ─────────────────────────────────────────────

def test_health_score_fresh_deal():
    """Свежая сделка на позднем этапе — высокий скор."""
    score = _calc_health_score(stage_id=5, days_in_stage=1,
                                amount=500_000, cycle_days=15)
    assert score >= 90


def test_health_score_stale_deal():
    """Зависшая сделка (35 дней без движения) — низкий скор."""
    score = _calc_health_score(stage_id=2, days_in_stage=35,
                                amount=200_000, cycle_days=60)
    assert score <= 50


def test_health_score_bounds():
    """Скор всегда в диапазоне 0-100."""
    for stage in range(0, 7):
        for days in [0, 7, 14, 30, 60, 120]:
            score = _calc_health_score(stage, days, 100_000, days)
            assert 0 <= score <= 100


# ─────────────────────────────────────────────
# Тесты маппинга этапов
# ─────────────────────────────────────────────

def test_load_stage_mapping_creates_default(tmp_path):
    """Если файл маппинга отсутствует — создаётся шаблон."""
    mapping_path = tmp_path / "stage_mapping.json"
    mapping = load_stage_mapping(mapping_path)
    assert mapping_path.exists()
    # Стандартные AmoCRM статусы должны быть
    assert 142 in mapping  # Closed Won
    assert mapping[142] == 6


def test_load_stage_mapping_custom(tmp_path):
    """Пользовательский маппинг загружается корректно."""
    mapping_path = tmp_path / "stage_mapping.json"
    custom = {"142": 6, "143": 0, "9001": 2, "9002": 4}
    mapping_path.write_text(json.dumps(custom), encoding="utf-8")
    mapping = load_stage_mapping(mapping_path)
    assert mapping[9001] == 2
    assert mapping[9002] == 4
    assert mapping[143] == 0


# ─────────────────────────────────────────────
# Тесты записи в Excel (UPSERT)
# ─────────────────────────────────────────────

def test_write_deals_adds_new_row(excel_with_raw_deals, stage_mapping,
                                   pipelines_map, users_map, contacts_map,
                                   sample_lead):
    """Новая сделка добавляется в конец raw_deals."""
    deal = map_lead_to_deal(
        sample_lead, contacts_map, users_map, pipelines_map,
        stage_mapping, "2026-09-30"
    )
    wb_path = excel_with_raw_deals
    result = write_deals_to_excel([deal], wb_path, dry_run=False)
    assert result.added == 1
    assert result.updated == 0

    # Проверяем что строка записалась
    wb = openpyxl.load_workbook(wb_path)
    ws = wb["raw_deals"]
    # Строка 2 = существующая MANUAL-001, строка 3 = новая из AmoCRM
    assert ws.cell(3, 1).value == "5555"
    assert ws.cell(3, 3).value == 750_000
    wb.close()


def test_write_deals_updates_existing(excel_with_raw_deals, stage_mapping,
                                       pipelines_map, users_map, contacts_map,
                                       sample_lead):
    """Существующая сделка (по deal_id) обновляется, не дублируется."""
    # Добавляем сделку первый раз
    deal = map_lead_to_deal(
        sample_lead, contacts_map, users_map, pipelines_map,
        stage_mapping, "2026-09-30"
    )
    wb_path = excel_with_raw_deals
    write_deals_to_excel([deal], wb_path, dry_run=False)

    # Меняем сумму и записываем снова
    sample_lead["price"] = 999_000
    deal2 = map_lead_to_deal(
        sample_lead, contacts_map, users_map, pipelines_map,
        stage_mapping, "2026-09-30"
    )
    result = write_deals_to_excel([deal2], wb_path, dry_run=False)
    assert result.updated == 1
    assert result.added == 0

    # Проверяем обновление
    wb = openpyxl.load_workbook(wb_path)
    ws = wb["raw_deals"]
    # Должна быть строка 3 с новой суммой
    amount_col = RAW_DEALS_HEADER.index("amount") + 1
    assert ws.cell(3, amount_col).value == 999_000
    # Строк должно быть ровно 3 (заголовок + MANUAL + AmoCRM)
    rows_with_data = sum(
        1 for r in range(2, ws.max_row + 1)
        if any(ws.cell(r, c).value is not None for c in range(1, 5))
    )
    assert rows_with_data == 2
    wb.close()


def test_write_deals_dry_run_no_changes(excel_with_raw_deals, stage_mapping,
                                         pipelines_map, users_map, contacts_map,
                                         sample_lead):
    """Dry run не изменяет Excel."""
    deal = map_lead_to_deal(
        sample_lead, contacts_map, users_map, pipelines_map,
        stage_mapping, "2026-09-30"
    )
    wb_path = excel_with_raw_deals

    # Запоминаем состояние до
    wb_before = openpyxl.load_workbook(wb_path)
    rows_before = wb_before["raw_deals"].max_row
    wb_before.close()

    write_deals_to_excel([deal], wb_path, dry_run=True)

    # После dry_run — без изменений
    wb_after = openpyxl.load_workbook(wb_path)
    rows_after = wb_after["raw_deals"].max_row
    wb_after.close()

    assert rows_before == rows_after


def test_manual_rows_not_overwritten(excel_with_raw_deals, stage_mapping,
                                      pipelines_map, users_map, contacts_map,
                                      sample_lead):
    """Ручные строки (не из AmoCRM) не затираются при UPSERT."""
    deal = map_lead_to_deal(
        sample_lead, contacts_map, users_map, pipelines_map,
        stage_mapping, "2026-09-30"
    )
    wb_path = excel_with_raw_deals
    write_deals_to_excel([deal], wb_path, dry_run=False)

    wb = openpyxl.load_workbook(wb_path)
    ws = wb["raw_deals"]
    # Строка 2 = MANUAL-001 должна остаться нетронутой
    assert ws.cell(2, 1).value == "MANUAL-001"
    assert ws.cell(2, 3).value == 100_000
    wb.close()


# ─────────────────────────────────────────────
# Тест custom field экстракции
# ─────────────────────────────────────────────

def test_extract_custom_field_found():
    lead = {
        "custom_fields_values": [
            {"field_code": "UTM_SOURCE", "values": [{"value": "yandex"}]},
        ]
    }
    assert _extract_custom_field(lead, "UTM_SOURCE") == "yandex"


def test_extract_custom_field_missing():
    lead = {"custom_fields_values": []}
    assert _extract_custom_field(lead, "UTM_SOURCE") == ""


def test_extract_custom_field_none_values():
    lead = {"custom_fields_values": None}
    assert _extract_custom_field(lead, "UTM_SOURCE") == ""


# ─────────────────────────────────────────────
# Тесты RateLimiter и защиты от лимитов
# ─────────────────────────────────────────────
from amocrm_connector import (
    RateLimiter,
    normalize_phone,
    extract_contact_phones,
    resolve_target_deal,
    find_deals_by_phone,
    match_call_to_deal,
)


def test_rate_limiter_pacing():
    limiter = RateLimiter(max_per_sec=20.0)
    t0 = time.time()
    limiter.wait()
    limiter.wait()
    t1 = time.time()
    assert (t1 - t0) >= 0.04


# ─────────────────────────────────────────────
# Тесты нормализации телефонов
# ─────────────────────────────────────────────
@pytest.mark.parametrize("raw,expected", [
    ("8 (999) 123-45-67", "+79991234567"),
    ("+7 999 123 45 67", "+79991234567"),
    ("79991234567", "+79991234567"),
    ("9991234567", "+79991234567"),
    ("+375 29 123-45-67", "+375291234567"),
    ("", ""),
    (None, ""),
])
def test_normalize_phone(raw, expected):
    assert normalize_phone(raw) == expected


def test_extract_contact_phones():
    contact = {
        "id": 1001,
        "custom_fields_values": [
            {
                "field_code": "PHONE",
                "values": [
                    {"value": "8 (999) 111-22-33"},
                    {"value": "+7 999 444-55-66"},
                    {"value": "89991112233"}, # дубль
                ],
            }
        ],
    }
    phones = extract_contact_phones(contact)
    assert len(phones) == 2
    assert "+79991112233" in phones
    assert "+79994445566" in phones


# ─────────────────────────────────────────────
# Тесты SmartDealMatcher (Защита от псевдо-дублей)
# ─────────────────────────────────────────────
def test_resolve_target_deal_prefers_active():
    """Активная сделка всегда имеет приоритет над закрытыми."""
    deals = [
        {"id": 1, "status_id": 142, "price": 5_000_000, "updated_at": 1000}, # Won
        {"id": 2, "status_id": 143, "price": 2_000_000, "updated_at": 2000}, # Lost
        {"id": 3, "status_id": 303, "price": 500_000, "updated_at": 500},    # Active (In progress)
    ]
    mapping = {303: 3, 142: 6, 143: 0}
    chosen = resolve_target_deal(deals, stage_mapping=mapping)
    assert chosen["id"] == 3


def test_resolve_target_deal_prefers_higher_amount():
    """При нескольких активных сделках выбирается сделка с максимальной суммой."""
    deals = [
        {"id": 10, "status_id": 303, "price": 300_000, "updated_at": 2000},
        {"id": 11, "status_id": 303, "price": 1_200_000, "updated_at": 1000},
        {"id": 12, "status_id": 303, "price": 150_000, "updated_at": 3000},
    ]
    mapping = {303: 3}
    chosen = resolve_target_deal(deals, stage_mapping=mapping)
    assert chosen["id"] == 11


def test_resolve_target_deal_prefers_latest_update_on_tie():
    """При равенстве сумм выбирается наиболее свежая сделка."""
    deals = [
        {"id": 21, "status_id": 303, "price": 500_000, "updated_at": 1000},
        {"id": 22, "status_id": 303, "price": 500_000, "updated_at": 5000},
    ]
    mapping = {303: 3}
    chosen = resolve_target_deal(deals, stage_mapping=mapping)
    assert chosen["id"] == 22


def test_resolve_target_deal_fallback_when_all_closed():
    """Если все сделки закрыты, выбирается наиболее свежая закрытая."""
    deals = [
        {"id": 31, "status_id": 143, "price": 100_000, "updated_at": 1000},
        {"id": 32, "status_id": 142, "price": 200_000, "updated_at": 3000},
    ]
    mapping = {142: 6, 143: 0}
    chosen = resolve_target_deal(deals, stage_mapping=mapping)
    assert chosen["id"] == 32


def test_match_call_to_deal_end_to_end():
    """Сквозной тест матчинга звонка по телефону к сделке."""
    contacts_map = {
        501: {
            "id": 501,
            "custom_fields_values": [
                {"field_code": "PHONE", "values": [{"value": "+7 (999) 777-88-99"}]}
            ],
        }
    }
    all_deals = [
        {"id": 101, "main_contact_id": 501, "status_id": 143, "price": 50_000, "updated_at": 100},
        {"id": 102, "main_contact_id": 501, "status_id": 303, "price": 750_000, "updated_at": 200},
    ]
    mapping = {303: 3, 143: 0}
    res = match_call_to_deal("89997778899", all_deals, contacts_map, mapping)
    assert res is not None
    assert res["id"] == 102

