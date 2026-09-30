"""
Тесты генерации ролевых отчетов RevOps Enterprise OS V17.6.
Проверяют конфигурации ролей, формирование HTML, номера страниц и команды Telegram.
"""
import pytest
import os
import json
import jinja2

from presentation.roles_config import ROLE_CONFIGS, get_role_context
from presentation.build import TEMPLATE_FILE, DATA_FILE
from telegram_bot import find_role_report_pdf, build_help_msg


# ── Проверка конфигурации ролей ──────────────────────────────────────

def test_role_configs_exist():
    expected_roles = {'all', 'ceo', 'rop', 'cfo'}
    assert set(ROLE_CONFIGS.keys()) == expected_roles


def test_role_expected_pages():
    assert ROLE_CONFIGS['all']['expected_pages'] == 10
    assert ROLE_CONFIGS['ceo']['expected_pages'] == 5
    assert ROLE_CONFIGS['rop']['expected_pages'] == 5
    assert ROLE_CONFIGS['cfo']['expected_pages'] == 5


def test_role_slide_counts_match_expected_pages():
    for role, cfg in ROLE_CONFIGS.items():
        assert len(cfg['slides']) == cfg['expected_pages'], f"Role {role} slide count mismatch"


def test_get_role_context_valid():
    ctx = get_role_context('ceo')
    assert ctx['total_pages'] == 5
    assert len(ctx['active_slides']) == 5
    assert 'slide_1' in ctx['active_slides']
    assert 'slide_2' in ctx['active_slides']
    assert 'slide_3' in ctx['active_slides']
    assert 'slide_9' in ctx['active_slides']
    assert 'slide_10' in ctx['active_slides']
    assert ctx['page_numbers']['slide_1'] == 1
    assert ctx['page_numbers']['slide_2'] == 2
    assert ctx['page_numbers']['slide_10'] == 5


def test_get_role_context_invalid():
    with pytest.raises(ValueError):
        get_role_context('unknown_role')


# ── Проверка шаблонизации Jinja2 для ролей ──────────────────────────

@pytest.fixture
def presentation_data():
    with open(DATA_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)


@pytest.fixture
def template_content():
    with open(TEMPLATE_FILE, 'r', encoding='utf-8') as f:
        return f.read()


def test_render_ceo_html(presentation_data, template_content):
    ctx = get_role_context('ceo')
    full_ctx = {**presentation_data, **ctx}
    template = jinja2.Template(template_content)
    rendered = template.render(**full_ctx)

    assert "EXECUTIVE BRIEFING • CEO" in rendered
    assert "Стратегический срез" in rendered
    assert "Executive Dashboard" in rendered
    assert "из 5" in rendered
    # В CEO отчете не должно быть пульта РОПа (slide_4) или аудита речи (slide_7)
    assert "Топ-5 Рисков Выручки" not in rendered
    assert "Whisper AI Speech Engine" not in rendered


def test_render_rop_html(presentation_data, template_content):
    ctx = get_role_context('rop')
    full_ctx = {**presentation_data, **ctx}
    template = jinja2.Template(template_content)
    rendered = template.render(**full_ctx)

    assert "SALES OPERATIONS PULSE • РОП" in rendered
    assert "Топ-5 Рисков Выручки" in rendered
    assert "Диагностика 7 Смертных Грехов Отдела Продаж" in rendered
    assert "Whisper AI Speech Engine" in rendered
    assert "из 5" in rendered
    # В отчете РОПа не должно быть финотчета DSO (slide_8) или One-Pager (slide_2)
    assert "Cash Flow &amp; AI Recovery" not in rendered
    assert "Executive Dashboard" not in rendered


def test_render_cfo_html(presentation_data, template_content):
    ctx = get_role_context('cfo')
    full_ctx = {**presentation_data, **ctx}
    template = jinja2.Template(template_content)
    rendered = template.render(**full_ctx)

    assert "CASH-FLOW &amp; DSO RECOVERY • CFO" in rendered or "CASH-FLOW & DSO RECOVERY • CFO" in rendered
    assert "Финансовый Контур и Agentic AI-Дожим" in rendered
    assert "Cash Flow &amp; AI Recovery" in rendered
    assert "из 5" in rendered
    # В отчете CFO не должно быть пульта РОПа (slide_4)
    assert "Топ-5 Рисков Выручки" not in rendered


def test_render_all_html(presentation_data, template_content):
    ctx = get_role_context('all')
    full_ctx = {**presentation_data, **ctx}
    template = jinja2.Template(template_content)
    rendered = template.render(**full_ctx)

    assert "RELEASE V17.6 ENTERPRISE" in rendered
    assert "Executive Dashboard" in rendered
    assert "Топ-5 Рисков Выручки" in rendered
    assert "Финансовый Контур и Agentic AI-Дожим" in rendered
    assert "из 10" in rendered


# ── Проверка поиска ролевых PDF ──────────────────────────────────────

def test_find_role_report_pdf():
    ceo_pdf = find_role_report_pdf('ceo')
    assert ceo_pdf is not None
    assert 'CEO' in os.path.basename(ceo_pdf).upper()

    rop_pdf = find_role_report_pdf('rop')
    assert rop_pdf is not None
    assert 'ROP' in os.path.basename(rop_pdf).upper()

    cfo_pdf = find_role_report_pdf('cfo')
    assert cfo_pdf is not None
    assert 'CFO' in os.path.basename(cfo_pdf).upper()

    all_pdf = find_role_report_pdf('all')
    assert all_pdf is not None


def test_telegram_help_contains_roles():
    help_msg = build_help_msg()
    assert "/ceo" in help_msg
    assert "/rop" in help_msg
    assert "/cfo" in help_msg
