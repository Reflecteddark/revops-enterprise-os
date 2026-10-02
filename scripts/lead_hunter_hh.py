"""
Парсер целевых лидов с HeadHunter (hh.ru) для Outbound-воронки RevOps Platform.
Собирает компании с открытыми вакансиями 'РОП', 'Руководитель отдела продаж', 'Менеджер B2B',
формирует Excel-таблицу с контактами, ссылками и ГОТОВЫМИ ПЕРСОНАЛИЗИРОВАННЫМИ СООБЩЕНИЯМИ для первого касания.
"""

import sys
import time
import re
from pathlib import Path
import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")


HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
    "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
}


def clean_url(url: str) -> str:
    if not url:
        return ""
    if url.startswith("/"):
        return f"https://hh.ru{url}"
    # Remove tracking click url if possible
    m = re.search(r"vacancy/(\d+)", url)
    if m:
        return f"https://hh.ru/vacancy/{m.group(1)}"
    return url.split("?")[0] if "?" in url else url


def fetch_vacancies(query="Руководитель отдела продаж", pages=2, area=113):
    """
    area=113 (Россия), area=1 (Москва), area=2 (СПб)
    """
    vacancies = []
    print(f"[*] Запуск сбора вакансий по запросу '{query}' (страниц: {pages})...")

    for page in range(pages):
        url = f"https://hh.ru/search/vacancy?text={requests.utils.quote(query)}&area={area}&page={page}&items_on_page=20"
        try:
            r = requests.get(url, headers=HEADERS, timeout=12)
            if r.status_code != 200:
                print(f"[!] Ошибка запроса страницы {page}: статус {r.status_code}")
                continue

            soup = BeautifulSoup(r.text, "html.parser")
            cards = soup.find_all("div", attrs={"data-qa": "vacancy-serp__vacancy"})
            if not cards:
                cards = soup.find_all("div", class_=lambda c: c and "vacancy-card" in c)

            for card in cards:
                title_elem = card.find("a", attrs={"data-qa": "serp-item__title"}) or card.find("span", attrs={"data-qa": "serp-item__title"})
                if not title_elem:
                    continue

                title = title_elem.get_text(strip=True)
                raw_link = title_elem.get("href", "")
                link = clean_url(raw_link)

                # Skip promo/ads
                if not link or "adsrv" in link:
                    # try to find real vacancy id
                    m = re.search(r"vacancy/(\d+)", raw_link)
                    if m:
                        link = f"https://hh.ru/vacancy/{m.group(1)}"
                    else:
                        continue

                company_elem = card.find("a", attrs={"data-qa": "vacancy-serp__vacancy-employer"}) or card.find("span", attrs={"data-qa": "vacancy-serp__vacancy-employer"})
                company = company_elem.get_text(strip=True) if company_elem else "Компания не указана"

                salary_elem = card.find("span", attrs={"data-qa": "vacancy-serp__vacancy-compensation"}) or card.find("span", class_=lambda c: c and "compensation" in c)
                salary = salary_elem.get_text(strip=True) if salary_elem else "Не указана"

                city_elem = card.find("span", attrs={"data-qa": "vacancy-serp__vacancy-address"}) or card.find("div", attrs={"data-qa": "vacancy-serp__vacancy-address"})
                city = city_elem.get_text(strip=True) if city_elem else "РФ"

                # Генерируем триггерное персонализированное сообщение
                if "руководител" in title.lower() or "роп" in title.lower():
                    trigger_angle = "Ищут РОПа (рутина контроля / дыра в аналитике)"
                    pitch_msg = (
                        f"«Приветствую! Вижу, в {company} сейчас открыта вакансия '{title}'. "
                        f"Пока идет найм сильного РОПа (а это обычно 1-2 месяца), мы можем за 24 часа подключить нейросетевой аудит 100% звонков вашего отдела: "
                        f"ИИ выявляет сливы сделок, контролирует стандарты и дает ежедневный светофор собственнику без раздувания ФОТ. "
                        f"Готовы бесплатно разобрать 3 любых ваших звонка за вчера, чтобы показать точки роста. Куда удобнее скинуть пример отчета?»"
                    )
                else:
                    trigger_angle = "Расширяют отдел продаж (слив лидов новичками)"
                    pitch_msg = (
                        f"«Приветствую! Вижу, вы активно нанимаете менеджеров по продажам в {company}. "
                        f"Главная проблема при масштабировании — новички сливают до 40% входящих лидов из-за забытых следующих шагов и невыявленных ЛПР. "
                        f"Наш сервис ai-rop.ru слушает 100% звонков менеджеров и выявляет ошибки в первый же день их работы. "
                        f"Предлагаю бесплатно протестировать систему на 3 ваших реальных звонках. Интересно взглянуть на формат отчета?»"
                    )

                vacancies.append({
                    "title": title,
                    "company": company,
                    "salary": salary,
                    "city": city,
                    "link": link,
                    "trigger_angle": trigger_angle,
                    "pitch_msg": pitch_msg
                })

            time.sleep(1.0)  # Gentle delay
        except Exception as e:
            print(f"[!] Ошибка при парсинге страницы {page}: {e}")

    # Remove duplicates by company + title
    unique_vacancies = []
    seen = set()
    for v in vacancies:
        key = (v["company"].lower(), v["title"].lower())
        if key not in seen:
            seen.add(key)
            unique_vacancies.append(v)

    print(f"[+] Собрано {len(unique_vacancies)} уникальных целевых компаний!")
    return unique_vacancies


def export_to_excel(vacancies, output_path: Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "🔥 Горячие_Лиды_HH"
    ws.views.sheetView[0].showGridLines = True

    c_navy_dark = "0F172A"
    c_blue_primary = "1E3A8A"
    c_blue_light = "EFF6FF"
    c_emerald_light = "ECFDF5"
    c_amber_light = "FFFBEB"
    c_card_bg = "F8FAFC"
    c_border = "CBD5E1"

    fill_navy = PatternFill(start_color=c_navy_dark, end_color=c_navy_dark, fill_type="solid")
    fill_header = PatternFill(start_color=c_blue_primary, end_color=c_blue_primary, fill_type="solid")
    fill_zebra = PatternFill(start_color=c_card_bg, end_color=c_card_bg, fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_blue_chip = PatternFill(start_color=c_blue_light, end_color=c_blue_light, fill_type="solid")

    font_title = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
    font_sub = Font(name="Segoe UI", size=9, italic=True, color="94A3B8")
    font_header = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
    font_bold = Font(name="Segoe UI", size=9, bold=True, color="1E293B")
    font_reg = Font(name="Segoe UI", size=8.5, color="334155")
    font_code = Font(name="Consolas", size=8, color="0F172A")
    font_link = Font(name="Segoe UI", size=8.5, color="2563EB", underline="single")

    thin_border = Border(
        left=Side(style="thin", color=c_border),
        right=Side(style="thin", color=c_border),
        top=Side(style="thin", color=c_border),
        bottom=Side(style="thin", color=c_border)
    )

    # 1. Заголовок
    ws.merge_cells("A1:H1")
    ws["A1"] = "🔥 RevOps Lead Hunter | База Горячих Лидов с HH.ru (Компании с открытыми вакансиями ОП)"
    ws["A1"].font = font_title
    ws["A1"].fill = fill_navy
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:H2")
    ws["A2"] = "Компании, которые прямо сейчас нанимают РОПов и продавцов. Готовые скрипты первого касания для Telegram/WhatsApp."
    ws["A2"].font = font_sub
    ws["A2"].fill = fill_navy
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 18

    # 2. Шапка таблицы
    headers = [
        "№", "Компания / Работодатель", "Открытая вакансия", "Зарплатная вилка",
        "Город", "Ссылка на вакансию", "Триггер выхода на ЛПР", "Готовое сообщение для отправки в Telegram (Copy-Paste)"
    ]

    ws.row_dimensions[4].height = 26
    for col_idx, h_text in enumerate(headers, start=1):
        cell = ws.cell(4, col_idx, h_text)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # 3. Данные
    current_row = 5
    for idx, v in enumerate(vacancies, start=1):
        ws.row_dimensions[current_row].height = 54
        bg = fill_zebra if idx % 2 == 1 else fill_white

        c1 = ws.cell(current_row, 1, idx)
        c1.font = font_bold
        c1.fill = fill_blue_chip
        c1.alignment = Alignment(horizontal="center", vertical="center")

        c2 = ws.cell(current_row, 2, v["company"])
        c2.font = font_bold
        c2.fill = bg
        c2.alignment = Alignment(vertical="center", wrap_text=True)

        c3 = ws.cell(current_row, 3, v["title"])
        c3.font = font_reg
        c3.fill = bg
        c3.alignment = Alignment(vertical="center", wrap_text=True)

        c4 = ws.cell(current_row, 4, v["salary"])
        c4.font = font_bold
        c4.fill = bg
        c4.alignment = Alignment(horizontal="center", vertical="center")

        c5 = ws.cell(current_row, 5, v["city"])
        c5.font = font_reg
        c5.fill = bg
        c5.alignment = Alignment(horizontal="center", vertical="center")

        c6 = ws.cell(current_row, 6, v["link"])
        c6.font = font_link
        c6.fill = bg
        c6.alignment = Alignment(vertical="center")

        c7 = ws.cell(current_row, 7, v["trigger_angle"])
        c7.font = font_bold
        c7.fill = bg
        c7.alignment = Alignment(vertical="center", wrap_text=True)

        c8 = ws.cell(current_row, 8, v["pitch_msg"])
        c8.font = font_code
        c8.fill = bg
        c8.alignment = Alignment(vertical="center", wrap_text=True)

        for col_idx in range(1, 9):
            ws.cell(current_row, col_idx).border = thin_border

        current_row += 1

    col_widths = {
        "A": 6,    # №
        "B": 24,   # Компания
        "C": 26,   # Вакансия
        "D": 18,   # Зарплата
        "E": 14,   # Город
        "F": 30,   # Ссылка
        "G": 28,   # Триггер
        "H": 50    # Сообщение
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    print(f"[✓] Таблица с горячими лидами сохранена: {output_path}")


def main():
    queries = ["Руководитель отдела продаж", "Менеджер по продажам B2B"]
    all_leads = []
    for q in queries:
        leads = fetch_vacancies(query=q, pages=2, area=113)
        all_leads.extend(leads)

    # Unique
    unique_leads = []
    seen = set()
    for item in all_leads:
        key = (item["company"].lower(), item["title"].lower())
        if key not in seen:
            seen.add(key)
            unique_leads.append(item)

    dest_desktop = Path(r"C:\Users\strel\Desktop\RevOps Platform\Каналы продаж\Горячие_Лиды_HH_RevOps.xlsx")
    dest_repo = Path(r"C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\docs\Горячие_Лиды_HH_RevOps.xlsx")

    export_to_excel(unique_leads, dest_desktop)
    export_to_excel(unique_leads, dest_repo)


if __name__ == "__main__":
    main()
