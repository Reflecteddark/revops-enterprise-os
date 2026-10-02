import gspread

def main():
    gc = gspread.service_account(filename="service_account.json")
    sh = gc.open_by_key("1AbaQGHhzar7Qsb9U9TByN57rAmOn36zw4HOD2XVrmTc")
    ws = sh.worksheet("🎙️ Клиентские_Пилоты")
    sheet_id = ws.id

    # Unmerge any previous merges
    sh.batch_update({"requests": [{
        "unmergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 0,
                "endRowIndex": 50,
                "startColumnIndex": 0,
                "endColumnIndex": 15
            }
        }
    }]})
    ws.clear()
    print("Worksheet cleared and unmerged.")

    title_block = [
        ["🎙️ RevOps OS // Пайплайн Клиентских Пилотов и Экспресс-Аудитов Звонков", "", "", "", "", "", "", "", "", "", "", ""],
        ["Воронка движения лидов: Входящий контакт / Аутрич → Экспресс-Аудит 3 звонков (0 ₽) → 7-дневный Пилот (14 900 ₽) → Подписка (29k / 59k / 89k ₽/мес). Актуализировано: 02.10.2026.", "", "", "", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", "", "", ""]
    ]

    headers = [
        "№ / Код",
        "Компания / Бренд",
        "Контактное лицо (ЛПР) / Telegram",
        "Отрасль (Сегмент ICP)",
        "Сейлзов в ОП",
        "CRM-система",
        "Этап сделки",
        "Статус аудита 3 звонков (0 ₽)",
        "Пилот 7 дней (14 900 ₽)",
        "Потенциальный LTV (Тариф)",
        "Сумма сливов в звонках (₽)",
        "Следующий шаг и дедлайн"
    ]

    demo_leads = [
        [
            "PLT-01",
            "ООО ТД «Металл-Комплект»",
            "Сергей Волков (Собственник)\n@volkov_metal",
            "B2B Опт / Металлопрокат",
            "8 чел",
            "amoCRM",
            "🔥 4. Пилот 14.9k запущен",
            "✅ Аудит сдан (слив 2.4 млн ₽)",
            "💳 Оплачен 14 900 ₽ (День 3 из 7)",
            "59 000 ₽ / мес (Scale)",
            "2 450 000 ₽",
            "Промежуточный отчет РОПу в Telegram в четверг 14:00"
        ],
        [
            "PLT-02",
            "Завод «СпецСтанок-М»",
            "Игорь Баринов (Гендир)\n@barinov_stank",
            "Промпроизводство оборудования",
            "5 чел",
            "Битрикс24",
            "⚡ 3. Отчёт сдан (дожим на пилот)",
            "✅ Аудит сдан (слив на ТЗ 4.1 млн ₽)",
            "⏳ Выставлен счет на 14 900 ₽",
            "29 000 ₽ / мес (Старт)",
            "4 120 000 ₽",
            "Созвон по согласованию счета в пятницу 11:30"
        ],
        [
            "PLT-03",
            "ГК «Монолит Девелопмент»",
            "Артем Ковалев (Директор ОП)\n@kovalev_monolit",
            "Девелопмент / Застройщик",
            "14 чел",
            "amoCRM",
            "💬 2. Получено 3 аудиофайла",
            "⚙️ В процессе разбора нейросетью",
            "📋 Ожидает результат аудита",
            "89 000 ₽ / мес (Enterprise)",
            "В расчете (~8.5 млн ₽)",
            "Отправить PDF-карту сливов сегодня до 18:00"
        ],
        [
            "PLT-04",
            "«ТракЛизинг Групп»",
            "Михаил Савельев (РОП)\n@savelev_truck",
            "Спецтехника и Лизинг",
            "6 чел",
            "Битрикс24",
            "📩 1. Отправлен аутрич (SCR-04)",
            "⏳ Запрошены 3 звонка",
            "📋 Презентован оффер 14.9k",
            "59 000 ₽ / мес (Scale)",
            "Оценка (~6.0 млн ₽)",
            "Follow-up #1 по скрипту SCR-09 завтра в 10:00"
        ]
    ]

    # Blank rows for incoming leads
    blank_rows = []
    for i in range(5, 16):
        blank_rows.append([f"PLT-{i:02d}", "", "", "", "", "", "1. Новый контакт", "Ожидание записей", "Не выставлен", "", "", ""])

    all_rows = title_block + [headers] + demo_leads + blank_rows
    ws.update(values=all_rows, range_name="A1")
    print(f"Data written successfully! ({len(all_rows)} rows)")

    # Formatting requests
    requests = [
        # Merge Title block
        {"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 12}, "mergeType": "MERGE_ALL"}},
        {"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": 12}, "mergeType": "MERGE_ALL"}},

        # Title formatting
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 12},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 0.06, "green": 0.09, "blue": 0.16},
                        "textFormat": {"foregroundColor": {"red": 0.06, "green": 0.72, "blue": 0.51}, "fontSize": 14, "bold": True},
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE"
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
            }
        },
        # Subtitle formatting
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": 12},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 0.09, "green": 0.13, "blue": 0.22},
                        "textFormat": {"foregroundColor": {"red": 0.58, "green": 0.64, "blue": 0.72}, "fontSize": 10, "italic": True},
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE"
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
            }
        },
        # Header Row (Row 4)
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 3, "endRowIndex": 4, "startColumnIndex": 0, "endColumnIndex": 12},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 0.12, "green": 0.18, "blue": 0.29},
                        "textFormat": {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "fontSize": 10, "bold": True},
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                        "wrapStrategy": "WRAP"
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)"
            }
        },
        # Data rows formatting
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 4, "endRowIndex": len(all_rows), "startColumnIndex": 0, "endColumnIndex": 12},
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {"fontSize": 10, "foregroundColor": {"red": 0.08, "green": 0.11, "blue": 0.18}},
                        "verticalAlignment": "TOP",
                        "wrapStrategy": "WRAP"
                    }
                },
                "fields": "userEnteredFormat(textFormat,verticalAlignment,wrapStrategy)"
            }
        },
        # Column A (Code): center & bold emerald
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 4, "endRowIndex": len(all_rows), "startColumnIndex": 0, "endColumnIndex": 1},
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {"bold": True, "fontSize": 10, "foregroundColor": {"red": 0.05, "green": 0.45, "blue": 0.35}},
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE"
                    }
                },
                "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment)"
            }
        },
        # Column Borders
        {
            "updateBorders": {
                "range": {"sheetId": sheet_id, "startRowIndex": 3, "endRowIndex": len(all_rows), "startColumnIndex": 0, "endColumnIndex": 12},
                "top": {"style": "SOLID", "color": {"red": 0.7, "green": 0.75, "blue": 0.82}},
                "bottom": {"style": "SOLID", "color": {"red": 0.7, "green": 0.75, "blue": 0.82}},
                "left": {"style": "SOLID", "color": {"red": 0.7, "green": 0.75, "blue": 0.82}},
                "right": {"style": "SOLID", "color": {"red": 0.7, "green": 0.75, "blue": 0.82}},
                "innerHorizontal": {"style": "SOLID", "color": {"red": 0.85, "green": 0.88, "blue": 0.92}},
                "innerVertical": {"style": "SOLID", "color": {"red": 0.85, "green": 0.88, "blue": 0.92}}
            }
        },
        # Column Widths
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 90}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2}, "properties": {"pixelSize": 210}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 3}, "properties": {"pixelSize": 220}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 3, "endIndex": 4}, "properties": {"pixelSize": 200}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 4, "endIndex": 5}, "properties": {"pixelSize": 100}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 5, "endIndex": 6}, "properties": {"pixelSize": 110}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 6, "endIndex": 7}, "properties": {"pixelSize": 200}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 7, "endIndex": 8}, "properties": {"pixelSize": 200}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 8, "endIndex": 9}, "properties": {"pixelSize": 190}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 9, "endIndex": 10}, "properties": {"pixelSize": 180}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 10, "endIndex": 11}, "properties": {"pixelSize": 160}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 11, "endIndex": 12}, "properties": {"pixelSize": 250}, "fields": "pixelSize"}},

        # Fixed Row Heights for headers
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 50}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": 1, "endIndex": 2}, "properties": {"pixelSize": 30}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": 2, "endIndex": 3}, "properties": {"pixelSize": 16}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": 3, "endIndex": 4}, "properties": {"pixelSize": 46}, "fields": "pixelSize"}},

        # Freeze top 4 rows
        {"updateSheetProperties": {"properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 4}}, "fields": "gridProperties.frozenRowCount"}}
    ]

    # Zebra shading for data rows
    for r_idx in range(4, len(all_rows)):
        bg = {"red": 0.97, "green": 0.98, "blue": 0.99} if (r_idx % 2 == 1) else {"red": 1.0, "green": 1.0, "blue": 1.0}
        requests.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": r_idx, "endRowIndex": r_idx + 1, "startColumnIndex": 0, "endColumnIndex": 12}, "cell": {"userEnteredFormat": {"backgroundColor": bg}}, "fields": "userEnteredFormat(backgroundColor)"}})
        # Dynamic row height
        lines = 1
        for cell_val in all_rows[r_idx]:
            lines = max(lines, str(cell_val).count("\n") + 1)
        h = max(45, int(lines * 19 + 20))
        requests.append({"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": r_idx, "endIndex": r_idx + 1}, "properties": {"pixelSize": h}, "fields": "pixelSize"}})

    sh.batch_update({"requests": requests})
    print("✓ Клиентские_Пилоты sheet updated successfully with 14 900 ₽ pricing and executive styling!")

if __name__ == "__main__":
    main()
