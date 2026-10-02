import gspread

def main():
    gc = gspread.service_account(filename="service_account.json")
    sh = gc.open_by_key("1AbaQGHhzar7Qsb9U9TByN57rAmOn36zw4HOD2XVrmTc")
    ws = sh.worksheet("📑 Инсайты_2_Исследований")
    sheet_id = ws.id

    # Unmerge pre-existing merges
    sh.batch_update({"requests": [{
        "unmergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 0,
                "endRowIndex": 50,
                "startColumnIndex": 0,
                "endColumnIndex": 10
            }
        }
    }]})
    ws.clear()
    print("Worksheet cleared and unmerged.")

    title_block = [
        ["📑 RevOps OS Pro // Главные Тезисы и Инсайты 2-х Исследований Рынка РФ (43 стр. + 15 стр.)", "", "", "", ""],
        ["Ключевые боли B2B-продаж, unit-экономика распознавания речи, архитектура интеграции с CRM и юридический контур 152-ФЗ. Актуализировано: 02.10.2026.", "", "", "", ""],
        ["", "", "", "", ""]
    ]

    headers = [
        "№ / Код",
        "Тема / Направление",
        "Ключевой инсайт исследования (Тезис)",
        "Почему это критически важно (Боль бизнеса)",
        "Решение в RevOps OS (Что делаем)"
    ]

    insights_data = [
        [
            "INS-1.1",
            "Кризис B2B-продаж и рост CAC",
            "Стоимость привлечения лида (CAC) в РФ выросла в 2.5–4 раза. Яндекс Директ перегрет, холодные звонки дают <1% конверсии.",
            "Бизнес больше не может расти экстенсивно за счет вливания бюджета в рекламу. Слив каждого лида обходится в 2 000 – 18 000 ₽ чистых потерь.",
            "Фокус RevOps OS: не заливать новый дорогой трафик, а ликвидировать скрытые дыры в воронке, возвращая до 35% зависших сделок без роста рекламного бюджета."
        ],
        [
            "INS-1.2",
            "Уровень утечек 1: «Справочное бюро»",
            "35% лидов сливаются на 1-м звонке: менеджеры подробно отвечают на вопросы клиента, но не берут инициативу и не закрывают на дату.",
            "Клиент получает бесплатную консультацию и уходит фразой «я подумаю». Менеджер ставит статус 'думает' и навсегда забывает о лиде.",
            "ИИ-Супервайзер по 13 критериям: жесткая проверка Next Step. Если точная дата и время контакта не согласованы — алерт РОПу через 60 секунд."
        ],
        [
            "INS-1.3",
            "Уровень утечек 2: «Слепота РОПа»",
            "РОП слушает менее 2% звонков физически (при объеме 300–500 звонков в неделю на отдел). Выводы строятся на субъективных ощущениях.",
            "Тренинги по продажам не работают, потому что РОП не видит системных ошибок и не контролирует закрепление навыков сейлзов в реальных звонках.",
            "Telegram-Супервайзер: утренний дайджест в 09:00 (приоритеты дня) + оперативные алерты на критические сливы. РОП тратит 15 мин/день вместо 4 часов в наушниках."
        ],
        [
            "INS-1.4",
            "Уровень утечек 3: «Кладбище зомби-сделок»",
            "В CRM висят миллионы на этапах «КП отправлено» или «В работе». При реальном аудите 65–75% из них брошены без касания 14+ дней.",
            "Отдел продаж создает иллюзию огромного пайплайна, а фактическая касса пуста. Собственник не видит реального прогноза выручки.",
            "Автоматический Радар Выручки: маркировка зомби-сделок, расчет упущенной выгоды и выгрузка списка реанимации для оперативного дожима."
        ],
        [
            "INS-1.5",
            "Позиционирование: RevOps-консультант",
            "SaaS-платформы (MANGO, UIS) дают сырые транскрипты без действий. Классический консалтинг стоит от 300к ₽ и длится 2-3 месяца.",
            "Клиенту не нужен голый софт и не нужны длинные отчеты — ему нужны деньги в кассе и внедренные регламенты без затягивания.",
            "Гибрид «Софт + Консалтинг»: 7-дневный пилотный спринт за 14 900 ₽ (со 100% гарантией возврата) + регулярная супервизия с гарантией окупаемости."
        ],
        [
            "INS-2.1",
            "Революция Unit-экономики STT",
            "Yandex SpeechKit Deferred (Batch STT) стоит ~0.02 ₽/мин ($0.0003), что в 3x дешевле Groq Whisper (~0.06 ₽/мин) и на 100% в РФ.",
            "Себестоимость распознавания всех звонков компании составляет всего 300–600 ₽ в месяц! Маржинальность нашего сервиса превышает 90%.",
            "Двухуровневая архитектура: Yandex Deferred для ночной 100% обработки звонков + локальные модели для быстрого (<5 сек) экспресс-аудита."
        ],
        [
            "INS-2.2",
            "Лимиты API amoCRM и Битрикс24",
            "amoCRM режет запросы при > 7 rps. Облачный Битрикс24 жестко блокирует при > 2 rps (пауза 0.5s). Вебхуки Битрикс24 лагают на 5-20 мин.",
            "Без встроенного рейт-лимитера и очереди выгрузка 500 звонков приводит к ошибкам 429/503 и полному падению синхронизации.",
            "Внедрен RateLimiter (6 rps amo, 2 rps Б24), экспоненциальный бэкофф с джиттером и Batch API (до 50 операций в 1 запросе к Б24)."
        ],
        [
            "INS-2.3",
            "Защита от псевдодублей в CRM",
            "Клиенты звонят с разных номеров. Метод crm.duplicate.findbycomm в Битрикс24 возвращает максимум 20 записей. Менеджеры плодят дубли.",
            "Если прикрепить звонок к неверной закрытой сделке или старому дублю, РОП видит искаженную картину и теряет нить переговоров.",
            "Алгоритм resolve_target_deal: нормализация номеров по E.164, приоритет открытых стадий воронки > наибольшая сумма > свежая дата контакта."
        ],
        [
            "INS-2.4",
            "Партнерство с интеграторами CRM",
            "Интеграторы amoCRM и Битрикс24 зарабатывают на разовой настройке и страдают от оттока. Им жизненно необходим recurring-доход.",
            "Интеграторы уже имеют доверие сотен лояльных B2B-клиентов. Партнерская сеть — главный бесплатный канал лидогенерации.",
            "Оффер интеграторам: 25–30% пожизненного ревшара (от 12 250 до 36 000 ₽/мес на каждого клиента). Мы ведем 100% клиентского саппорта."
        ],
        [
            "INS-2.5",
            "Контур безопасности 152-ФЗ РФ",
            "Штрафы Роскомнадзора за нарушение 152-ФЗ и хранение ПДн за рубежом достигают 6–18 млн ₽. Любая служба безопасности блокирует сервис без документов.",
            "Клиенты боятся проверок, а менеджеры могут пожаловаться в трудовую инспекцию на прослушку звонков.",
            "Модуль PIISanitizer (токенизация телефонов/email) + комплект из 5 документов (согласие сотрудника, приказ, регламент 30 дней, чек НПД)."
        ]
    ]

    all_rows = title_block + [headers] + insights_data
    ws.update(values=all_rows, range_name="A1")
    print(f"Data written successfully! ({len(all_rows)} rows)")

    requests = [
        # Merge Title block
        {"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 5}, "mergeType": "MERGE_ALL"}},
        {"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": 5}, "mergeType": "MERGE_ALL"}},

        # Title formatting
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 5},
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
                "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": 5},
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
                "range": {"sheetId": sheet_id, "startRowIndex": 3, "endRowIndex": 4, "startColumnIndex": 0, "endColumnIndex": 5},
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
                "range": {"sheetId": sheet_id, "startRowIndex": 4, "endRowIndex": len(all_rows), "startColumnIndex": 0, "endColumnIndex": 5},
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
        # Table Borders
        {
            "updateBorders": {
                "range": {"sheetId": sheet_id, "startRowIndex": 3, "endRowIndex": len(all_rows), "startColumnIndex": 0, "endColumnIndex": 5},
                "top": {"style": "SOLID", "color": {"red": 0.7, "green": 0.75, "blue": 0.82}},
                "bottom": {"style": "SOLID", "color": {"red": 0.7, "green": 0.75, "blue": 0.82}},
                "left": {"style": "SOLID", "color": {"red": 0.7, "green": 0.75, "blue": 0.82}},
                "right": {"style": "SOLID", "color": {"red": 0.7, "green": 0.75, "blue": 0.82}},
                "innerHorizontal": {"style": "SOLID", "color": {"red": 0.85, "green": 0.88, "blue": 0.92}},
                "innerVertical": {"style": "SOLID", "color": {"red": 0.85, "green": 0.88, "blue": 0.92}}
            }
        },
        # Column Widths
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 95}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2}, "properties": {"pixelSize": 240}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 3}, "properties": {"pixelSize": 380}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 3, "endIndex": 4}, "properties": {"pixelSize": 400}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 4, "endIndex": 5}, "properties": {"pixelSize": 420}, "fields": "pixelSize"}},

        # Fixed Row Heights for headers
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 52}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": 1, "endIndex": 2}, "properties": {"pixelSize": 30}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": 2, "endIndex": 3}, "properties": {"pixelSize": 16}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": 3, "endIndex": 4}, "properties": {"pixelSize": 46}, "fields": "pixelSize"}},

        # Freeze top 4 rows
        {"updateSheetProperties": {"properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 4}}, "fields": "gridProperties.frozenRowCount"}}
    ]

    # Zebra striping & dynamic row heights
    for idx, item in enumerate(insights_data):
        r_num = 4 + idx
        bg = {"red": 0.97, "green": 0.98, "blue": 0.99} if (r_num % 2 == 1) else {"red": 1.0, "green": 1.0, "blue": 1.0}
        requests.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": r_num, "endRowIndex": r_num + 1, "startColumnIndex": 0, "endColumnIndex": 5}, "cell": {"userEnteredFormat": {"backgroundColor": bg}}, "fields": "userEnteredFormat(backgroundColor)"}})
        
        # Calculate visual lines
        lines_b = (len(item[1]) + 24) // 24
        lines_c = (len(item[2]) + 40) // 40
        lines_d = (len(item[3]) + 42) // 42
        lines_e = (len(item[4]) + 44) // 44
        max_l = max(lines_b, lines_c, lines_d, lines_e, 3)
        h = max(80, int(max_l * 19 + 25))
        requests.append({"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": r_num, "endIndex": r_num + 1}, "properties": {"pixelSize": h}, "fields": "pixelSize"}})

    sh.batch_update({"requests": requests})
    print("✓ Инсайты_2_Исследований updated with 14 900 ₽ and dynamic heights!")

if __name__ == "__main__":
    main()
