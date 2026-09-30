r"""
RevOps Telegram Bot — ежедневный брифинг отдела продаж.

Два режима:
  1. ОТПРАВКА (основной): читает Excel → считает KPI → шлёт сообщение
     python telegram_bot.py --send
     python telegram_bot.py --send --days-left 5

  2. POLLING (опционально): бот отвечает на /status, /risk, /hot
     python telegram_bot.py --polling

Настройка:
  python telegram_bot.py --setup

Автозапуск (Windows Task Scheduler):
  Программа: python
  Аргументы: C:\...\telegram_bot.py --send
  Триггер: ежедневно в 07:30

Зависимости: только requests (уже установлен).
"""

import sys
import os
import json
import glob
import time
import argparse
from datetime import datetime, date
from pathlib import Path

import requests
import openpyxl

try:
    from config import BASE_DIR, XLSX_PATTERN
except ImportError:
    BASE_DIR = Path(__file__).parent
    XLSX_PATTERN = "RevOps Platform V17*.xlsx"

CREDS_FILE = BASE_DIR / "telegram_credentials.json"

# ── Telegram API (без лишних зависимостей) ──────────────────────────

def tg_send(token: str, chat_id: str, text: str,
            parse_mode: str = "HTML") -> bool:
    """Отправляет сообщение через Bot API. Возвращает True при успехе."""
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    resp = requests.post(url, json={
        "chat_id": chat_id,
        "text": text,
        "parse_mode": parse_mode,
        "disable_web_page_preview": True,
    }, timeout=15)
    if not resp.ok:
        print(f"[TG ERROR] {resp.status_code}: {resp.text[:200]}")
    return resp.ok


def tg_send_document(token: str, chat_id: str, file_path: str,
                     caption: str = "", timeout: int = 40) -> bool:
    """Отправляет PDF-документ через Telegram Bot API."""
    url = f"https://api.telegram.org/bot{token}/sendDocument"
    if not os.path.exists(file_path):
        print(f"[TG ERROR] Файл для отправки не найден: {file_path}")
        return False
    try:
        with open(file_path, "rb") as f:
            resp = requests.post(
                url,
                data={"chat_id": chat_id, "caption": caption},
                files={"document": f},
                timeout=timeout
            )
        if not resp.ok:
            print(f"[TG ERROR] {resp.status_code}: {resp.text[:200]}")
        return resp.ok
    except Exception as e:
        print(f"[TG ERROR] Ошибка при отправке документа: {e}")
        return False


def tg_get_updates(token: str, offset: int = 0) -> list[dict]:
    url = f"https://api.telegram.org/bot{token}/getUpdates"
    resp = requests.get(url, params={"offset": offset, "timeout": 20}, timeout=25)
    if resp.ok:
        return resp.json().get("result", [])
    return []


def tg_get_me(token: str) -> dict | None:
    """Запрашивает информацию о боте через getMe. Возвращает dict с данными бота или None."""
    url = f"https://api.telegram.org/bot{token}/getMe"
    try:
        resp = requests.get(url, timeout=10)
        if resp.ok:
            return resp.json().get("result")
    except Exception as e:
        print(f"[TG ERROR] getMe failed: {e}")
    return None


def check_crm_status() -> str:
    """Проверяет наличие учетных данных CRM-коннекторов."""
    amo_creds = BASE_DIR / "amocrm_credentials.json"
    b24_creds = BASE_DIR / "bitrix24_credentials.json"
    lines = ["🔄 <b>Статус CRM-интеграций:</b>", ""]
    if amo_creds.exists():
        lines.append("• <b>AmoCRM:</b> настроена ✅")
    else:
        lines.append("• <b>AmoCRM:</b> не настроена ⚠️ (нужен amocrm_credentials.json)")
    if b24_creds.exists():
        lines.append("• <b>Bitrix24:</b> настроена ✅")
    else:
        lines.append("• <b>Bitrix24:</b> не настроена ⚠️ (нужен bitrix24_credentials.json)")
    lines.append("")
    lines.append("Для синхронизации запустите соответствующий коннектор.")
    return "\n".join(lines)


# ── Учётные данные ──────────────────────────────────────────────────

def load_creds() -> dict:
    if not CREDS_FILE.exists():
        print(f"[ОШИБКА] {CREDS_FILE} не найден. Запустите: python telegram_bot.py --setup")
        sys.exit(1)
    return json.loads(CREDS_FILE.read_text(encoding="utf-8"))


def save_creds(data: dict) -> None:
    CREDS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


# ── Аналитика из Excel ──────────────────────────────────────────────

def find_wb() -> str:
    candidates = [f for f in glob.glob(str(BASE_DIR / XLSX_PATTERN))
                  if not os.path.basename(f).startswith("~$")]
    if not candidates:
        print(f"[ОШИБКА] Excel не найден: {XLSX_PATTERN}")
        sys.exit(1)
    return max(candidates, key=os.path.getmtime)


def find_role_report_pdf(role: str = "all") -> str | None:
    """Находит свежий PDF-отчёт для указанной роли (all, ceo, rop, cfo)."""
    role = role.lower()
    pattern_map = {
        "all": ["*Executive_Summary*.pdf", "*Full_Report*.pdf", "RevOps*.pdf"],
        "ceo": ["*Report_CEO*.pdf", "*CEO*.pdf"],
        "rop": ["*Report_ROP*.pdf", "*ROP*.pdf", "*Пульт_РОПа*.pdf"],
        "cfo": ["*Report_CFO*.pdf", "*CFO*.pdf", "*Финансы*.pdf"],
    }
    patterns = pattern_map.get(role, ["*Executive_Summary*.pdf", "RevOps*.pdf"])

    candidates = []
    for loc in [BASE_DIR, BASE_DIR / "presentation", Path.home() / "Desktop"]:
        for pat in patterns:
            try:
                candidates.extend(glob.glob(str(loc / pat)))
            except Exception:
                pass

    valid = [
        f for f in candidates
        if os.path.isfile(f) and os.path.getsize(f) > 0 and not os.path.basename(f).startswith("~$")
    ]
    if valid:
        return max(valid, key=os.path.getmtime)

    # Фоллбек: любые PDF внутри BASE_DIR
    fallback = [
        f for f in glob.glob(str(BASE_DIR / "*.pdf"))
        if os.path.isfile(f) and os.path.getsize(f) > 0 and not os.path.basename(f).startswith("~$")
    ]
    if fallback:
        return max(fallback, key=os.path.getmtime)
    return None


def find_latest_report_pdf() -> str | None:
    """Находит самый свежий PDF-отчёт (для обратной совместимости)."""
    return find_role_report_pdf("all")


def _col(header: list, name: str, default: int) -> int:
    try:
        return header.index(name) + 1
    except ValueError:
        return default


def calc_kpi(wb_path: str) -> dict:
    """
    Читает raw_deals и ⚙️ Настройки → возвращает KPI-словарь.
    Все вычисления здесь — бот только форматирует.
    """
    wb = openpyxl.load_workbook(wb_path, data_only=True)

    # План из настроек
    plan = 0.0
    org_name = "RevOps"
    try:
        ws_s = wb["⚙️ Настройки"]
        org_name = str(ws_s["B3"].value or "RevOps").strip()
        plan = float(ws_s["B9"].value or 0)
    except Exception:
        pass

    # raw_deals
    ws = wb["raw_deals"]
    hdr = [str(ws.cell(1, c).value or "").strip() for c in range(1, ws.max_column + 1)]

    c_id    = _col(hdr, "deal_id",         1)
    c_name  = _col(hdr, "client_name",     2)
    c_amt   = _col(hdr, "amount",          3)
    c_stg   = _col(hdr, "stage_id",        4)
    c_mgr   = _col(hdr, "manager_id",      8)
    c_days  = _col(hdr, "days_in_stage",   17)
    c_wval  = _col(hdr, "weighted_val",    19)
    c_hlth  = _col(hdr, "deal_health_score", 35)

    active, won_sum, weighted, at_risk, hot = [], 0.0, 0.0, [], []
    mgr_won: dict[str, float] = {}

    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, min(ws.max_column + 1, 6))]
        if not any(v is not None for v in vals):
            continue

        try:
            stage = int(float(ws.cell(r, c_stg).value or 0))
        except (ValueError, TypeError):
            continue

        amt   = float(ws.cell(r, c_amt).value or 0)
        mgr   = str(ws.cell(r, c_mgr).value or "—")
        deal  = str(ws.cell(r, c_id).value or "")
        cname = str(ws.cell(r, c_name).value or "")

        try:
            days = int(ws.cell(r, c_days).value or 0)
        except (ValueError, TypeError):
            days = 0
        try:
            wval = float(ws.cell(r, c_wval).value or 0)
        except (ValueError, TypeError):
            wval = amt * 0.3
        try:
            health = int(ws.cell(r, c_hlth).value or 50)
        except (ValueError, TypeError):
            health = 50

        if stage == 6:
            won_sum += amt
            mgr_won[mgr] = mgr_won.get(mgr, 0) + amt

        elif 1 <= stage <= 5:
            active.append({"id": deal, "name": cname, "amt": amt,
                           "stage": stage, "mgr": mgr, "days": days,
                           "wval": wval, "health": health})
            weighted += wval

            if days > 14 or health < 40:
                at_risk.append({"name": cname, "amt": amt, "stage": stage,
                                "mgr": mgr, "days": days})

            if stage >= 4 and amt > 0:
                hot.append({"name": cname, "amt": amt, "stage": stage, "mgr": mgr})

    wb.close()

    # Дни до конца месяца
    today = date.today()
    import calendar
    last_day = calendar.monthrange(today.year, today.month)[1]
    days_left = last_day - today.day

    # Топ менеджеров по выручке
    top_mgrs = sorted(mgr_won.items(), key=lambda x: x[1], reverse=True)[:3]

    # Сортировки
    at_risk.sort(key=lambda x: x["days"], reverse=True)
    hot.sort(key=lambda x: x["amt"], reverse=True)

    return {
        "org": org_name,
        "plan": plan,
        "won_sum": won_sum,
        "plan_pct": round(won_sum / plan * 100, 1) if plan else 0,
        "active_count": len(active),
        "weighted": weighted,
        "at_risk": at_risk[:5],
        "hot": hot[:3],
        "top_mgrs": top_mgrs,
        "days_left": days_left,
        "date_str": today.strftime("%d.%m.%Y"),
    }


# ── Форматирование сообщений ────────────────────────────────────────

def _fmt_m(v: float) -> str:
    """750000 → '750 000 ₽'"""
    return f"{v:,.0f} ₽".replace(",", " ")


def build_morning_brief(kpi: dict) -> str:
    pct = kpi["plan_pct"]
    pct_emoji = "🔴" if pct < 60 else ("🟡" if pct < 85 else "🟢")
    dl = kpi.get("days_left", 0)
    dl_str = ("⚠️ СЕГОДНЯ ПОСЛЕДНИЙ ДЕНЬ!" if dl == 0
              else f"📅 До конца месяца: {dl} дн.")

    lines = [
        f"<b>🌅 RevOps Пульс | {kpi['date_str']}</b>",
        f"<i>{kpi['org']}</i>",
        "",
        f"{pct_emoji} <b>Выполнение плана: {pct}%</b>",
        f"   Закрыто: {_fmt_m(kpi['won_sum'])} / план {_fmt_m(kpi['plan'])}",
        f"📊 Пайплайн: {kpi['active_count']} сделок | weighted {_fmt_m(kpi['weighted'])}",
        f"{dl_str}",
    ]

    if kpi["at_risk"]:
        lines += ["", "🔴 <b>В риске (нужна работа сегодня):</b>"]
        for d in kpi["at_risk"]:
            lines.append(f"  • {d['mgr']} — {d['name'] or 'б/н'} "
                         f"({_fmt_m(d['amt'])}, {d['days']} дн. на эт.{d['stage']})")

    if kpi["hot"]:
        lines += ["", "🟢 <b>Горячие (закрыть в приоритете):</b>"]
        for d in kpi["hot"]:
            stg_label = {4: "КП отправлено", 5: "Подписание"}.get(d["stage"], f"эт.{d['stage']}")
            lines.append(f"  • {d['mgr']} — {d['name'] or 'б/н'} "
                         f"({_fmt_m(d['amt'])}, {stg_label})")

    if kpi["top_mgrs"]:
        lines += ["", "🏆 <b>Топ по выручке (месяц):</b>"]
        for i, (mgr, amt) in enumerate(kpi["top_mgrs"], 1):
            lines.append(f"  {i}. {mgr} — {_fmt_m(amt)}")

    lines += ["", "─────────────────", "Команды: /risk /hot /plan /summary /help"]
    return "\n".join(lines)


def build_risk_msg(kpi: dict) -> str:
    if not kpi["at_risk"]:
        return "✅ Зависших сделок нет. Отдел в порядке."
    lines = [f"<b>🔴 Сделки в риске — {kpi['date_str']}</b>", ""]
    for d in kpi["at_risk"]:
        lines.append(f"<b>{d['mgr']}</b>: {d['name'] or 'б/н'}")
        lines.append(f"  Сумма: {_fmt_m(d['amt'])} | Этап: {d['stage']} | {d['days']} дн.")
        lines.append("")
    return "\n".join(lines)


def build_hot_msg(kpi: dict) -> str:
    if not kpi["hot"]:
        return "ℹ️ Нет сделок на этапах КП/Подписание."
    lines = [f"<b>🟢 Горячие сделки — {kpi['date_str']}</b>", ""]
    for d in kpi["hot"]:
        lines.append(f"<b>{d['mgr']}</b>: {d['name'] or 'б/н'}")
        lines.append(f"  {_fmt_m(d['amt'])} | Этап: {d['stage']}")
        lines.append("")
    return "\n".join(lines)


def build_plan_msg(kpi: dict) -> str:
    return (
        f"<b>📊 Статус плана — {kpi['date_str']}</b>\n\n"
        f"Выполнено: <b>{kpi['plan_pct']}%</b>\n"
        f"Закрыто: {_fmt_m(kpi['won_sum'])}\n"
        f"План: {_fmt_m(kpi['plan'])}\n"
        f"Осталось: {_fmt_m(max(0, kpi['plan'] - kpi['won_sum']))}\n\n"
        f"Weighted pipeline: {_fmt_m(kpi['weighted'])}\n"
        f"Активных сделок: {kpi['active_count']}"
    )


def build_help_msg() -> str:
    return (
        "🤖 <b>RevOps Enterprise Bot — Команды:</b>\n\n"
        "• /status или /start — Утренний брифинг отдела продаж\n"
        "• /risk — Сделки в зоне риска (зависшие > 14 дней или health < 40)\n"
        "• /hot — Горячие сделки на стадиях КП и Подписание договора\n"
        "• /plan — Детализация выполнения финансового плана\n"
        "• /ceo — 👔 Стратегический отчёт для CEO / Собственника (PDF, 5 стр.)\n"
        "• /rop — 📋 Операционный пульт РОПа & аудит звонков (PDF, 5 стр.)\n"
        "• /cfo — 💳 Финансовый срез: платежный календарь и DSO (PDF, 5 стр.)\n"
        "• /summary или /report — 📄 Полная мастер-презентация C-Level (PDF, 10 стр.)\n"
        "• /sync — Проверить статус CRM-интеграций (AmoCRM / Bitrix24)\n"
        "• /help — Справка по доступным командам"
    )


# ── Polling (опциональный командный режим) ──────────────────────────

def run_polling(token: str, chat_id: str, wb_path: str) -> None:
    """Простой polling-бот. Отвечает на /start /status /risk /hot /plan /summary /report /ceo /rop /cfo /sync /help."""
    print("Polling... (Ctrl+C для выхода)")
    offset = 0
    COMMANDS = {"/risk", "/hot", "/plan", "/status", "/start", "/summary", "/report", "/help", "/sync", "/ceo", "/rop", "/cfo"}

    while True:
        try:
            updates = tg_get_updates(token, offset)
            for upd in updates:
                offset = upd["update_id"] + 1
                msg = upd.get("message", {})
                from_id = str(msg.get("chat", {}).get("id", ""))
                raw_text = msg.get("text", "").strip()
                text = raw_text.lower().split()[0] if raw_text else ""

                # Принимаем только из разрешённого chat_id
                if from_id != str(chat_id) or text not in COMMANDS:
                    continue

                if text in ("/start", "/status"):
                    kpi = calc_kpi(wb_path)
                    reply = build_morning_brief(kpi)
                    tg_send(token, chat_id, reply)
                elif text == "/risk":
                    kpi = calc_kpi(wb_path)
                    reply = build_risk_msg(kpi)
                    tg_send(token, chat_id, reply)
                elif text == "/hot":
                    kpi = calc_kpi(wb_path)
                    reply = build_hot_msg(kpi)
                    tg_send(token, chat_id, reply)
                elif text == "/plan":
                    kpi = calc_kpi(wb_path)
                    reply = build_plan_msg(kpi)
                    tg_send(token, chat_id, reply)
                elif text == "/sync":
                    reply = check_crm_status()
                    tg_send(token, chat_id, reply)
                elif text == "/help":
                    reply = build_help_msg()
                    tg_send(token, chat_id, reply)
                elif text in ("/summary", "/report", "/ceo", "/rop", "/cfo"):
                    parts = raw_text.lower().split()
                    role = "all"
                    if text == "/ceo" or (len(parts) > 1 and parts[1] == "ceo"):
                        role = "ceo"
                    elif text == "/rop" or (len(parts) > 1 and parts[1] == "rop"):
                        role = "rop"
                    elif text == "/cfo" or (len(parts) > 1 and parts[1] == "cfo"):
                        role = "cfo"

                    role_captions = {
                        "all": "📄 RevOps Executive Master Report (10 страниц)",
                        "ceo": "👔 RevOps Отчет для CEO / Собственника (5 страниц)",
                        "rop": "📋 RevOps Пульт РОПа: Сделки в риске и аудит звонков (5 страниц)",
                        "cfo": "💳 RevOps Финансовый срез: Платежный календарь и DSO (5 страниц)",
                    }
                    tg_send(token, chat_id, f"⏳ Подготавливаю отчет ({role.upper()})...")
                    pdf = find_role_report_pdf(role)
                    if not pdf:
                        try:
                            cmd_gen = [sys.executable, str(BASE_DIR / "presentation" / "build.py"), "--role", role, "--skip-presteps"]
                            subprocess.run(cmd_gen, cwd=BASE_DIR, timeout=45)
                            pdf = find_role_report_pdf(role)
                        except Exception as e:
                            print(f"[TG ERROR] Auto-generation failed: {e}")

                    if pdf:
                        caption = role_captions.get(role, f"📄 RevOps Report: {os.path.basename(pdf)}")
                        ok = tg_send_document(token, chat_id, pdf, caption=caption)
                        if not ok:
                            tg_send(token, chat_id, "❌ Не удалось отправить документ через Telegram API.")
                    else:
                        tg_send(token, chat_id, f"⚠️ PDF-отчет для роли '{role}' не найден. Запустите генерацию через launchers/Generate_Report_{role.upper()}.bat")
                else:
                    continue

                print(f"  Ответил на {text} → {from_id}")

            time.sleep(1)
        except KeyboardInterrupt:
            print("\nПолинг остановлен.")
            break
        except Exception as e:
            print(f"[POLLING ERROR] {e}")
            time.sleep(5)


# ── Настройка ───────────────────────────────────────────────────────

def cmd_setup() -> None:
    print("\n" + "=" * 55)
    print("  Настройка Telegram бота RevOps")
    print("=" * 55)
    print()
    print("1. Откройте @BotFather в Telegram")
    print("2. Напишите /newbot → получите токен")
    print("3. Напишите боту любое сообщение")
    print("4. Откройте: https://api.telegram.org/bot<TOKEN>/getUpdates")
    print("   Найдите 'chat': {'id': ЧИСЛО} — это ваш chat_id")
    print()
    token = input("Bot token: ").strip()
    chat_id = input("Chat ID (число): ").strip()

    # Тест отправки
    ok = tg_send(token, chat_id,
                 "✅ RevOps Bot подключён! Используйте /status для брифинга.")
    if ok:
        save_creds({"bot_token": token, "chat_id": chat_id})
        print(f"\n✓ Готово! Credentials: {CREDS_FILE}")
        print("\nДля утреннего авторассылки добавьте в Task Scheduler:")
        print(f"  python \"{BASE_DIR / 'telegram_bot.py'}\" --send")
        print("  Триггер: ежедневно в 07:30")
    else:
        print("\n[ОШИБКА] Проверьте токен и chat_id.")
        sys.exit(1)


def cmd_test(token: str, chat_id: str, wb_path: str) -> bool:
    """Проверяет подключение к Telegram API, Excel и отправляет тестовый пинг."""
    print("\n" + "=" * 55)
    print("  Диагностика RevOps Telegram Bot")
    print("=" * 55)

    # 1. Проверка getMe
    bot_info = tg_get_me(token)
    if not bot_info:
        print("❌ Ошибка авторизации: неверный bot_token")
        return False
    bot_username = bot_info.get("username", "Unknown")
    print(f"✓ Бот авторизован: @{bot_username} (ID: {bot_info.get('id')})")

    # 2. Проверка книги Excel
    if os.path.exists(wb_path):
        print(f"✓ Книга данных найдена: {os.path.basename(wb_path)}")
    else:
        print(f"❌ Книга данных не найдена: {wb_path}")
        return False

    # 3. Проверка PDF
    pdf = find_latest_report_pdf()
    if pdf:
        print(f"✓ Найден свежий PDF-отчёт: {os.path.basename(pdf)}")
    else:
        print("ℹ️ Свежий PDF-отчёт пока не сформирован")

    # 4. Проверка тестовой отправки
    print(f"Отправка тестового пинга в chat_id {chat_id}...")
    ok = tg_send(
        token,
        chat_id,
        "🔔 <b>Тест связи RevOps Bot</b>\n\nДиагностика пройдена успешно. Бот готов к рассылке и приёму команд."
    )
    if ok:
        print("✓ Тестовое сообщение успешно доставлено в чат!")
    else:
        print("❌ Не удалось доставить сообщение в чат. Проверьте chat_id и права бота.")
        return False

    print("\n🎉 Все проверки успешно пройдены!")
    return True


# ── CLI ─────────────────────────────────────────────────────────────

def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(
        description="RevOps Telegram Bot — ежедневный брифинг отдела продаж"
    )
    parser.add_argument("--setup",    action="store_true", help="Настройка бота")
    parser.add_argument("--test",     action="store_true", help="Проверить подключение и отправить тестовый пинг")
    parser.add_argument("--preview",  action="store_true", help="Предпросмотр утреннего дайджеста в консоли (без отправки)")
    parser.add_argument("--dry-run",  action="store_true", help="Алиас для --preview")
    parser.add_argument("--send",     action="store_true", help="Отправить утренний брифинг")
    parser.add_argument("--summary",  action="store_true", help="Отправить свежий PDF-отчёт")
    parser.add_argument("--polling",  action="store_true", help="Запустить polling-бот")
    parser.add_argument("--workbook", type=str, default=None, help="Путь к Excel")
    args = parser.parse_args()

    if args.setup:
        cmd_setup()
        return 0

    wb_path = args.workbook or find_wb()

    if args.preview or args.dry_run:
        print(f"\n[PREVIEW] Чтение данных: {wb_path}")
        kpi = calc_kpi(wb_path)
        msg = build_morning_brief(kpi)
        print("\n" + "=" * 55)
        print("  ПРЕДПРОСМОТР УТРЕННЕГО БРИФИНГА (БЕЗ ОТПРАВКИ)")
        print("=" * 55 + "\n")
        clean_text = msg.replace("<b>", "").replace("</b>", "").replace("<i>", "").replace("</i>", "")
        print(clean_text)
        print("\n" + "=" * 55)
        pdf = find_latest_report_pdf()
        print(f"Свежий PDF-отчёт: {pdf if pdf else 'Не найден'}")
        print("✓ Предпросмотр завершён успешно.")
        return 0

    creds = load_creds()
    token    = creds["bot_token"]
    chat_id  = creds["chat_id"]

    if args.test:
        ok = cmd_test(token, chat_id, wb_path)
        return 0 if ok else 1

    if args.summary:
        pdf = find_latest_report_pdf()
        if pdf:
            print(f"Отправляем PDF: {pdf}")
            ok = tg_send_document(token, chat_id, pdf, caption=f"📄 RevOps Executive Report: {os.path.basename(pdf)}")
            print("✓ PDF-отчет отправлен." if ok else "[ОШИБКА] Не удалось отправить PDF.")
            return 0 if ok else 1
        else:
            print("[ОШИБКА] PDF-отчет не найден.")
            return 1

    if args.send:
        print(f"Читаем данные: {wb_path}")
        kpi = calc_kpi(wb_path)
        msg = build_morning_brief(kpi)
        ok = tg_send(token, chat_id, msg)
        print("✓ Брифинг отправлен." if ok else "[ОШИБКА] Не отправлено.")
        return 0 if ok else 1

    if args.polling:
        run_polling(token, chat_id, wb_path)
        return 0

    # По умолчанию — отправить
    kpi = calc_kpi(wb_path)
    tg_send(token, chat_id, build_morning_brief(kpi))
    return 0


if __name__ == "__main__":
    sys.exit(main())
