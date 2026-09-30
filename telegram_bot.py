"""
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


def tg_get_updates(token: str, offset: int = 0) -> list[dict]:
    url = f"https://api.telegram.org/bot{token}/getUpdates"
    resp = requests.get(url, params={"offset": offset, "timeout": 20}, timeout=25)
    if resp.ok:
        return resp.json().get("result", [])
    return []


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
    dl = kpi["days_left"]
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

    lines += ["", "─────────────────", "Команды: /risk /hot /plan"]
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


# ── Polling (опциональный командный режим) ──────────────────────────

def run_polling(token: str, chat_id: str, wb_path: str) -> None:
    """Простой polling-бот. Отвечает на /start /status /risk /hot /plan."""
    print(f"Polling... (Ctrl+C для выхода)")
    offset = 0
    COMMANDS = {"/risk", "/hot", "/plan", "/status", "/start"}

    while True:
        try:
            updates = tg_get_updates(token, offset)
            for upd in updates:
                offset = upd["update_id"] + 1
                msg = upd.get("message", {})
                from_id = str(msg.get("chat", {}).get("id", ""))
                text = msg.get("text", "").strip().lower().split()[0] if msg.get("text") else ""

                # Принимаем только из разрешённого chat_id
                if from_id != str(chat_id) or text not in COMMANDS:
                    continue

                kpi = calc_kpi(wb_path)
                if text in ("/start", "/status"):
                    reply = build_morning_brief(kpi)
                elif text == "/risk":
                    reply = build_risk_msg(kpi)
                elif text == "/hot":
                    reply = build_hot_msg(kpi)
                elif text == "/plan":
                    reply = build_plan_msg(kpi)
                else:
                    continue

                tg_send(token, chat_id, reply)
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


# ── CLI ─────────────────────────────────────────────────────────────

def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(
        description="RevOps Telegram Bot — ежедневный брифинг отдела продаж"
    )
    parser.add_argument("--setup",    action="store_true", help="Настройка бота")
    parser.add_argument("--send",     action="store_true", help="Отправить утренний брифинг")
    parser.add_argument("--polling",  action="store_true", help="Запустить polling-бот")
    parser.add_argument("--workbook", type=str, default=None, help="Путь к Excel")
    args = parser.parse_args()

    if args.setup:
        cmd_setup()
        return 0

    creds = load_creds()
    token    = creds["bot_token"]
    chat_id  = creds["chat_id"]
    wb_path  = args.workbook or find_wb()

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
