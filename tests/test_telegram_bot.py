"""Тесты Telegram бота — без реального Telegram и Excel."""
import pytest
import openpyxl
from unittest.mock import patch, MagicMock
from telegram_bot import (
    calc_kpi, build_morning_brief, build_risk_msg,
    build_hot_msg, build_plan_msg, build_help_msg,
    find_latest_report_pdf, tg_send, tg_send_document, tg_get_updates,
    _fmt_m, _col,
)


@pytest.fixture
def excel_wb(tmp_path):
    wb = openpyxl.Workbook()

    # Настройки
    ws_s = wb.active
    ws_s.title = "⚙️ Настройки"
    ws_s["B3"] = "ООО Тест"
    ws_s["B9"] = 10_000_000

    # raw_deals
    ws = wb.create_sheet("raw_deals")
    header = ["deal_id", "client_name", "amount", "stage_id", "stage_name",
              "created_date", "stage_changed_date", "manager_id", "segment_code",
              "lead_source", "client_id", "scoring_tier", "loss_code", "synced_at",
              "is_won", "is_lost", "days_in_stage", "aging_cohort", "weighted_val",
              "cycle_days", "utm_source", "utm_medium", "utm_campaign", "yclid",
              "gclid", "data_snapshot_date", "data_version", "uploaded_by",
              "upload_hash", "next_action_type", "next_action_date",
              "next_action_owner", "next_action_channel", "next_action_health",
              "deal_health_score"]
    ws.append(header)
    # Закрытые сделки (stage=6)
    ws.append(["D001", "ООО Альфа", 3_000_000, 6, "WON", "2026-09-01", "2026-09-20",
               "Иванов А.", "", "", "", "A", "", "", 1, 0, 0, "Fresh", 3_000_000,
               30, "", "", "", "", "", "2026-09-30", "17.6", "test", "D001",
               "", "", "Иванов А.", "", "OK", 90])
    # Активные сделки
    ws.append(["D002", "ООО Бета", 800_000, 4, "КП", "2026-08-15", "2026-09-15",
               "Петров К.", "", "", "", "B", "", "", 0, 0, 15, "Aging", 400_000,
               45, "", "", "", "", "", "2026-09-30", "17.6", "test", "D002",
               "", "", "Петров К.", "", "РИСК", 35])
    ws.append(["D003", "ИП Гамма", 1_200_000, 5, "Подписание", "2026-09-01", "2026-09-28",
               "Иванов А.", "", "", "", "A", "", "", 0, 0, 2, "Fresh", 900_000,
               29, "", "", "", "", "", "2026-09-30", "17.6", "test", "D003",
               "", "", "Иванов А.", "", "OK", 85])
    ws.append(["D004", "АО Дельта", 200_000, 2, "Квалификация", "2026-08-01", "2026-08-30",
               "Сидоров Н.", "", "", "", "C", "", "", 0, 0, 31, "Stale", 30_000,
               60, "", "", "", "", "", "2026-09-30", "17.6", "test", "D004",
               "", "", "Сидоров Н.", "", "РИСК", 20])

    p = tmp_path / "RevOps Platform V17.6 Test.xlsx"
    wb.save(p)
    wb.close()
    return str(p)


# ── KPI ──────────────────────────────────────────────────────────────

def test_calc_kpi_won_sum(excel_wb):
    kpi = calc_kpi(excel_wb)
    assert kpi["won_sum"] == 3_000_000


def test_calc_kpi_plan(excel_wb):
    kpi = calc_kpi(excel_wb)
    assert kpi["plan"] == 10_000_000


def test_calc_kpi_plan_pct(excel_wb):
    kpi = calc_kpi(excel_wb)
    assert kpi["plan_pct"] == 30.0  # 3M / 10M


def test_calc_kpi_active_count(excel_wb):
    kpi = calc_kpi(excel_wb)
    assert kpi["active_count"] == 3  # D002, D003, D004


def test_calc_kpi_at_risk(excel_wb):
    """Сделки с days_in_stage > 14 или health < 40 → в риске."""
    kpi = calc_kpi(excel_wb)
    risk_ids = {d["name"] for d in kpi["at_risk"]}
    assert "ООО Бета" in risk_ids    # 15 дней, health 35
    assert "АО Дельта" in risk_ids   # 31 день, health 20


def test_calc_kpi_hot(excel_wb):
    """Этапы 4-5 → горячие."""
    kpi = calc_kpi(excel_wb)
    hot_names = {d["name"] for d in kpi["hot"]}
    assert "ИП Гамма" in hot_names   # этап 5
    assert "ООО Бета" in hot_names   # этап 4


def test_calc_kpi_top_mgrs(excel_wb):
    """Топ менеджеров по закрытым сделкам."""
    kpi = calc_kpi(excel_wb)
    assert kpi["top_mgrs"][0][0] == "Иванов А."
    assert kpi["top_mgrs"][0][1] == 3_000_000


def test_calc_kpi_org_name(excel_wb):
    kpi = calc_kpi(excel_wb)
    assert kpi["org"] == "ООО Тест"


# ── Форматирование ────────────────────────────────────────────────────

def test_fmt_m():
    assert "750" in _fmt_m(750_000)
    assert "₽" in _fmt_m(1_000_000)


def test_morning_brief_contains_plan(excel_wb):
    kpi = calc_kpi(excel_wb)
    msg = build_morning_brief(kpi)
    assert "30.0%" in msg
    assert "В риске" in msg
    assert "Горячие" in msg


def test_risk_msg_empty():
    kpi = {"at_risk": [], "date_str": "30.09.2026"}
    msg = build_risk_msg(kpi)
    assert "нет" in msg.lower()


def test_risk_msg_with_deals():
    kpi = {
        "at_risk": [{"name": "ООО X", "amt": 500_000, "stage": 3,
                     "mgr": "Иванов", "days": 20}],
        "date_str": "30.09.2026",
    }
    msg = build_risk_msg(kpi)
    assert "ООО X" in msg
    assert "Иванов" in msg


def test_hot_msg_empty():
    kpi = {"hot": [], "date_str": "30.09.2026"}
    assert "Нет" in build_hot_msg(kpi)


def test_plan_msg(excel_wb):
    kpi = calc_kpi(excel_wb)
    msg = build_plan_msg(kpi)
    assert "30.0%" in msg
    assert "10" in msg   # план 10М присутствует


def test_build_help_msg():
    msg = build_help_msg()
    assert "/status" in msg
    assert "/summary" in msg
    assert "/risk" in msg
    assert "/help" in msg


def test_morning_brief_commands():
    kpi = {
        "org": "Тест", "date_str": "30.09.2026", "plan_pct": 50.0,
        "won_sum": 5_000_000, "plan": 10_000_000, "weighted": 2_000_000,
        "active_count": 5, "at_risk": [], "hot": [], "top_mgrs": []
    }
    msg = build_morning_brief(kpi)
    assert "/summary" in msg
    assert "/help" in msg


def test_find_latest_report_pdf(tmp_path, monkeypatch):
    import telegram_bot
    f1 = tmp_path / "RevOps_Enterprise_OS_V17.6_Full_Report_1.pdf"
    f2 = tmp_path / "RevOps_Enterprise_OS_V17.6_Full_Report_2.pdf"
    f1.write_text("dummy pdf 1", encoding="utf-8")
    f2.write_text("dummy pdf 2", encoding="utf-8")
    monkeypatch.setattr(telegram_bot, "BASE_DIR", tmp_path)
    res = find_latest_report_pdf()
    assert res is not None
    assert "RevOps" in res


def test_find_latest_report_pdf_none(tmp_path, monkeypatch):
    from pathlib import Path
    import telegram_bot
    monkeypatch.setattr(telegram_bot, "BASE_DIR", tmp_path)
    monkeypatch.setattr(Path, "home", lambda: tmp_path)
    res = find_latest_report_pdf()
    assert res is None


def test_tg_send_success():
    with patch("requests.post") as mock_post:
        mock_post.return_value = MagicMock(ok=True)
        ok = tg_send("token123", "chat123", "Hello")
        assert ok is True
        mock_post.assert_called_once()
        assert "sendMessage" in mock_post.call_args[0][0]


def test_tg_send_failure():
    with patch("requests.post") as mock_post:
        mock_post.return_value = MagicMock(ok=False, status_code=400, text="Bad Request")
        ok = tg_send("token123", "chat123", "Hello")
        assert ok is False


def test_tg_send_document_file_not_found():
    ok = tg_send_document("token123", "chat123", "non_existent_file.pdf")
    assert ok is False


def test_tg_send_document_success(tmp_path):
    dummy_pdf = tmp_path / "test.pdf"
    dummy_pdf.write_bytes(b"%PDF-1.4 test")
    with patch("requests.post") as mock_post:
        mock_post.return_value = MagicMock(ok=True)
        ok = tg_send_document("token123", "chat123", str(dummy_pdf), caption="Test Report")
        assert ok is True
        mock_post.assert_called_once()
        assert "sendDocument" in mock_post.call_args[0][0]


def test_tg_get_updates():
    with patch("requests.get") as mock_get:
        mock_get.return_value = MagicMock(ok=True, json=lambda: {"result": [{"update_id": 100}]})
        updates = tg_get_updates("token123")
        assert len(updates) == 1
        assert updates[0]["update_id"] == 100
