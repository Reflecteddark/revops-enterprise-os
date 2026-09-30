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
