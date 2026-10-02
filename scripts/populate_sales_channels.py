import gspread
from pathlib import Path

def main():
    gc = gspread.service_account(filename="service_account.json")
    sh = gc.open_by_key("1AbaQGHhzar7Qsb9U9TByN57rAmOn36zw4HOD2XVrmTc")
    
    ws_title = "📢 Каналы_Продаж"
    try:
        ws = sh.worksheet(ws_title)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=ws_title, rows=120, cols=10)

    sheet_id = ws.id

    # 1. Unmerge any existing merges to avoid cell swallowing
    sh.batch_update({"requests": [{
        "unmergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 0,
                "endRowIndex": 120,
                "startColumnIndex": 0,
                "endColumnIndex": 10
            }
        }
    }]})
    ws.clear()
    print("Worksheet cleared and unmerged.")

    # Data arrays
    # -------------------------------------------------------------
    # Title Block
    title_block = [
        ["📢 RevOps OS Pro // Анализ Целевого Рынка B2B (ICP) и База Telegram-Каналов для Лидогенерации", "", "", "", "", "", "", ""],
        ["Глубокая сегментация B2B-клиентов, средние чеки, точки сливов + 36 проверенных Telegram-каналов и сообществ для аутрича и посевов. Актуализировано: 02.10.2026.", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", ""]
    ]

    # Section 1 Header
    sec1_banner = [
        ["🎯 РАЗДЕЛ 1: АНАЛИЗ РЫНКА B2B И ПРОФИЛИ ИДЕАЛЬНОГО КЛИЕНТА (ICP — КОМУ ПРОДАВАТЬ)", "", "", "", "", "", "", ""]
    ]
    sec1_headers = [
        "Код",
        "Отрасль / Сегмент рынка",
        "Портрет ЛПР (Кто решает)",
        "Ср. чек и цикл сделки",
        "Главные боли и точки сливов в звонках",
        "Цена ошибки (Потери в месяц)",
        "Убойный оффер для первого контакта",
        "Приоритет сегмента"
    ]

    sec1_data = [
        [
            "SEG-01",
            "B2B Оптовая торговля и Дистрибуция\n(Химия, сырье, металлопрокат, упаковка, кабель, стройматериалы оптом)",
            "Собственник, Генеральный директор, Коммерческий директор (Head of Sales)",
            "300 000 – 3 500 000 ₽\nЦикл: 14–45 дней",
            "• Менеджеры отправляют КП/счет и забывают перезвонить («ждут ответа»).\n• Слив на первом же «Дорого / у конкурентов дешевле».\n• Нет жесткой фиксации следующего шага с датой и временем.\n• CRM заполнена на 30%, сделки зависают в статусе «Думают».",
            "1.5 – 4.0 млн ₽ маржи в месяц\n(Слив всего 2-3 крупных оптовых поставок)",
            "Экспресс-аудит 3 вчерашних звонков по оптовым сделкам + оцифровка упущенной маржи за 20 минут. Пилот 7 дней за 14 900 ₽.",
            "🔥 Tier-1 (Максимальный)"
        ],
        [
            "SEG-02",
            "Заводы и Промышленное производство\n(Станки, оборудование, промтара, металлоконструкции, комплектующие)",
            "Владелец завода, Генеральный директор, Директор по сбыту",
            "500 000 – 10 000 000 ₽\nЦикл: 30–90 дней",
            "• Менеджеры-инженеры отлично разбираются в железе, но боятся продавать и называть цену.\n• После долгого согласования ТЗ пауза на 3 недели, клиент уходит к более активным конкурентам.\n• Не выявляют реальный бюджет и срок ввода в эксплуатацию.",
            "3.0 – 8.0 млн ₽ в месяц\n(Потеря одного контракта на поставку перечеркивает квартальный план)",
            "Контроль 100% звонков менеджеров-инженеров: алерт РОПу за 60 секунд, если клиент ушел с ТЗ к конкуренту. Разбор 3 сложных переговоров.",
            "🔥 Tier-1 (Максимальный)"
        ],
        [
            "SEG-03",
            "Недвижимость, Девелопмент и Застройщики\n(Жилые комплексы, коммерческая недвижимость, загородные поселки)",
            "Директор по маркетингу и продажам (CMO/CRO), Коммерческий директор, РОП",
            "8 000 000 – 35 000 000 ₽\n(Лид стоит 6 000 – 18 000 ₽)",
            "• Менеджеры дежурно «консультируют» вместо продажи («вся инфа на сайте»).\n• Не квалифицируют ипотечную готовность и бюджет покупки.\n• Не продают встречу/показ объекта на объекте строительства.\n• Слив звонка = выброшенные деньги на рекламу в Яндекс/Таргете.",
            "Сотни миллионов выручки\n(Слив даже 1-2 покупателей квартир/помещений)",
            "Мгновенный перехват лидов: ИИ находит слитые звонки с рекламы до того, как клиент ушел к соседнему застройщику. Спасение рекламного бюджета.",
            "🔥 Tier-1 (Максимальный)"
        ],
        [
            "SEG-04",
            "Спецтехника, Коммерческий транспорт и Лизинг\n(Грузовики, экскаваторы, тракторы, тягачи, погрузчики)",
            "Собственник дилерского центра, Генеральный директор, РОП",
            "2 000 000 – 25 000 000 ₽\nЦикл: 20–60 дней",
            "• Менеджеры не умеют продавать лизинговые программы и расчет окупаемости.\n• Теряют контакт после первого звонка («клиент сказал, что пока выбирает»).\n• Не отрабатывают возражения по китайским аналогам и срокам поставки.",
            "5.0 – 15.0 млн ₽ с каждой упущенной единицы техники",
            "Всего 1 спасенная сделка окупает подписку на 5 лет вперед. Бесплатный разбор 3 записей переговоров по технике за 20 минут.",
            "🔥 Tier-1 (Максимальный)"
        ],
        [
            "SEG-05",
            "B2B Услуги, IT-разработка и Digital-агентства\n(Заказная разработка ПО, интеграторы, маркетинг под ключ, консалтинг)",
            "Основатель (Founder / CEO), Руководитель отдела продаж",
            "150 000 – 1 500 000 ₽\nЦикл: 14–35 дней",
            "• Менеджеры не квалифицируют платежеспособность и тратят часы на пустое составление смет.\n• Сливают лиды на этапе «согласования договора».\n• Нет дисциплины ведения CRM: теряется история коммуникаций.",
            "800 000 – 2.0 млн ₽ в месяц\n(Потеря 2-4 контрактов на разработку/услуги)",
            "100% автозаполнение CRM и резюме переговоров без рутины менеджеров + оцифровка конверсии каждого сейлза. Демо за 15 минут.",
            "⚡ Tier-2 (Высокий)"
        ],
        [
            "SEG-06",
            "Сертификация, Лицензирование и Юруслуги B2B\n(Охрана труда, СОУТ, пожарная безопасность, банкротство юрлиц)",
            "Генеральный директор, Коммерческий директор",
            "80 000 – 500 000 ₽\nЦикл: 7–21 день",
            "• Менеджеры нарушают регламенты и обязательные юридические скрипты.\n• Забывают предложить сопутствующие обязательные аудиты и лицензии.\n• Риск штрафов и проверок регуляторов из-за некорректных консультаций.",
            "500 000 – 1.5 млн ₽ упущенной выгоды на допродажах",
            "Автоматический чек-лист из 12 параметров по каждому звонку: контроль соблюдения скрипта и стандартов 152-ФЗ. Пилот за 14 900 ₽.",
            "⚡ Tier-2 (Высокий)"
        ],
        [
            "SEG-07",
            "EdTech B2B и Корпоративное обучение\n(Обучение персонала, бизнес-школы, корпоративные тренинги)",
            "Директор по продажам, РОП, Фаундер",
            "70 000 – 400 000 ₽\nОтделы продаж: 10–40 чел",
            "• Огромный поток звонков (сотни в день); РОП физически слышит не более 2% разговоров.\n• Менеджеры не дожимают участников бесплатных вебинаров и мастер-классов.\n• Высокая текучка среди менеджеров, долгий ввод новичков.",
            "1.2 – 3.0 млн ₽ в месяц на брошенных заявках",
            "100% охват всех звонков 24/7 без найма дорогостоящего отдела контроля качества (ОКК) — чистая экономия от 120 000 ₽/мес на ФОТ контролеров.",
            "⚡ Tier-2 (Высокий)"
        ]
    ]

    # Section 2 Header
    sec2_banner = [
        ["", "", "", "", "", "", "", ""],
        ["📱 РАЗДЕЛ 2: БАЗА TELEGRAM-КАНАЛОВ И СООБЩЕСТВ ДЛЯ ПРИВЛЕЧЕНИЯ КЛИЕНТОВ (36 КАНАЛОВ И ЧАТОВ)", "", "", "", "", "", "", ""]
    ]
    sec2_headers = [
        "Код",
        "Категория ресурса",
        "Название канала / Чата",
        "Ссылка / Handle (Telegram)",
        "Кто целевая аудитория (ЛПР)",
        "Формат работы (Стратегия захода)",
        "Рекомендуемый скрипт / Оффер",
        "Потенциал канала"
    ]

    sec2_data = [
        # Категория 1: Профессиональные сообщества РОПов и Продаж
        [
            "TG-01",
            "1. Сообщество РОПов и Продаж",
            "Гребенюк | Результат чужими руками",
            "https://t.me/grebenuk_m",
            "РОПы, коммерческие директора, собственники B2B (160k+ подп.)",
            "Экспертный комментинг под постами о проблемах ОП и нехватке времени у РОПа. Парсинг активных комментаторов.",
            "SCR-01 (Собственнику)\nSCR-02 (РОПу)",
            "🔥 Сверхвысокий"
        ],
        [
            "TG-02",
            "1. Сообщество РОПов и Продаж",
            "Чат Школы РОПов Гребенюка",
            "https://t.me/grebenuk_chat",
            "Действующие РОПы и руководители отделов продаж (12k+ участников)",
            "Точечный нетворкинг, ответы на вопросы о контроле звонков в CRM, предложение разобрать 3 вчерашних звонка.",
            "SCR-02 (Разгрузка 4ч рутины)\nSCR-05 (Краш-тест звонков)",
            "🔥 Сверхвысокий"
        ],
        [
            "TG-03",
            "1. Сообщество РОПов и Продаж",
            "Максим Батырев | Комбат",
            "https://t.me/batyrev_maxim",
            "Топ-менеджеры, РОПы, собственники с филиальными сетями (85k+ подп.)",
            "Комментинг постов про стандарты работы, дисциплину менеджеров и регламенты звонков.",
            "SCR-01 (Оцифровка сливов)\nSCR-03 (Ультра-блиц)",
            "🔥 Сверхвысокий"
        ],
        [
            "TG-04",
            "1. Сообщество РОПов и Продаж",
            "Екатерина Уколова | Продажи в B2B",
            "https://t.me/ukolova_pro",
            "Владельцы отделов продаж, коммерческие директора (45k+ подп.)",
            "Посевы, комментинг разборов воронок продаж, предложение оцифровки упущенной прибыли.",
            "SCR-02 (РОПу)\nSCR-04 (Отраслевой)",
            "⭐ Высокий"
        ],
        [
            "TG-05",
            "1. Сообщество РОПов и Продаж",
            "Олег Шевелев | ПораРасти",
            "https://t.me/shevelev_oleg",
            "РОПы, тренеры по продажам, менеджеры B2B (30k+ подп.)",
            "Экспертные кейсы с реальными аудиозаписями сливов на возражении «Дорого».",
            "SCR-05 (Вызов команде)\nSCR-07 (Прием файлов)",
            "⭐ Высокий"
        ],
        [
            "TG-06",
            "1. Сообщество РОПов и Продаж",
            "Сергей Азимов | Продажи и Переговоры",
            "https://t.me/azimov_sergey",
            "РОПы, переговорщики, сейлзы с большими чеками (25k+ подп.)",
            "Кейсы отработки возражений и правильного назначения следующего шага.",
            "SCR-11 (Возражение РОПа)\nSCR-01 (Собственнику)",
            "⭐ Высокий"
        ],
        [
            "TG-07",
            "1. Сообщество РОПов и Продаж",
            "Клуб РОПов России (Чат руководителей)",
            "https://t.me/rop_community",
            "Действующие РОПы и руководители коммерческих служб (8k+ участников)",
            "Прямой аутрич, обсуждение болей контроля менеджеров и нехватки времени на прослушку.",
            "SCR-02 (Экономия 4 часов)\nSCR-08 (Презентация аудита)",
            "🔥 Сверхвысокий"
        ],

        # Категория 2: Сообщества Собственников бизнеса, Фаундеров и CEO
        [
            "TG-08",
            "2. Клубы Собственников и CEO",
            "Т-Бизнес | Бизнес-секреты",
            "https://t.me/t_business",
            "Владельцы малого и среднего B2B-бизнеса (120k+ подп.)",
            "Мониторинг обсуждений предпринимателей, точечный аутрич собственников в личку.",
            "SCR-01 (Удар по деньгам)\nSCR-03 (Блиц)",
            "🔥 Сверхвысокий"
        ],
        [
            "TG-09",
            "2. Клубы Собственников и CEO",
            "VC.ru — Главное о бизнесе и SaaS",
            "https://t.me/vc_ru",
            "Фаундеры, IT-руководители, C-level менеджеры (190k+ подп.)",
            "Статьи-разборы на vc.ru + ссылка на Дмитрия Федотова в профиле, экспертные ветки.",
            "SCR-01 (Собственнику)\nSCR-06 (CRM-фокус)",
            "🔥 Сверхвысокий"
        ],
        [
            "TG-10",
            "2. Клубы Собственников и CEO",
            "Аркадий Морейнис | Тёмная сторона",
            "https://t.me/temno",
            "Фаундеры, инвесторы, B2B-бизнесмены (90k+ подп.)",
            "Комментинг бизнес-моделей, поиск фаундеров с масштабируемыми отделами продаж.",
            "SCR-01 (ROI и потери)\nSCR-03 (3 строки)",
            "⭐ Высокий"
        ],
        [
            "TG-11",
            "2. Клубы Собственников и CEO",
            "Бизнес-клуб Terra (Терра)",
            "https://t.me/terra_community",
            "Предприниматели с оборотом от 5 до 100 млн ₽ (50k+ участников)",
            "Личный нетворкинг Дмитрия Федотова, бесплатный экспресс-разбор 3 звонков для резидентов.",
            "SCR-01 (Собственнику)\nSCR-07 (Прием аудио)",
            "🔥 Сверхвысокий"
        ],
        [
            "TG-12",
            "2. Клубы Собственников и CEO",
            "Сколково Стартапы & Бизнес",
            "https://t.me/skolkovo_startups",
            "Технологические фаундеры, B2B SaaS компании (22k+ участников)",
            "Нетворкинг, обмен B2B-лидами, предложение автозаполнения CRM.",
            "SCR-03 (Блиц)\nSCR-06 (CRM-фокус)",
            "⭐ Высокий"
        ],
        [
            "TG-13",
            "2. Клубы Собственников и CEO",
            "Деловая среда (СберБизнес)",
            "https://t.me/delovayasreda",
            "Владельцы опта, услуг и производства (40k+ предпринимателей)",
            "Посевы полезного контента, комментинг тем роста выручки и конверсий.",
            "SCR-01 (Оцифровка потерь)\nSCR-04 (Отраслевой)",
            "⭐ Высокий"
        ],
        [
            "TG-14",
            "2. Клубы Собственников и CEO",
            "Опора России | Малый и средний бизнес",
            "https://t.me/opora_russia",
            "Региональные производители, оптовики, дистрибьюторы (35k+ подп.)",
            "Отраслевой аутрич по региональным коммерческим директорам.",
            "SCR-04 (Крупный B2B-чек)\nSCR-01 (Собственнику)",
            "⭐ Высокий"
        ],

        # Категория 3: CRM-интеграторы, IT-директора и внедренцы
        [
            "TG-15",
            "3. CRM-интеграторы и IT",
            "amoCRM | Официальный канал",
            "https://t.me/amocrm_official",
            "Пользователи CRM, РОПы, интеграторы (45k+ подп.)",
            "Мониторинг комментариев о проблемах телефонии и контроля работы менеджеров.",
            "SCR-06 (Автозаполнение CRM)\nSCR-13 (Партнерство)",
            "🔥 Сверхвысокий"
        ],
        [
            "TG-16",
            "3. CRM-интеграторы и IT",
            "Чат Партнёров и Интеграторов amoCRM",
            "https://t.me/amocrm_partners",
            "Руководители digital-агентств и внедренцы amoCRM (5k+ интеграторов)",
            "Партнёрский оффер: 25–30% рекуррентной комиссии с каждого внедрения клиенту.",
            "SCR-13 (Партнёрский аутрич)\nSCR-06 (CRM-директору)",
            "🔥 Сверхвысокий"
        ],
        [
            "TG-17",
            "3. CRM-интеграторы и IT",
            "Битрикс24 для бизнеса",
            "https://t.me/bitrix24_official",
            "Руководители компаний на Битрикс24, ИТ-директора (60k+ подп.)",
            "Экспертные статьи по автоматизации контроля качества звонков поверх Б24.",
            "SCR-06 (100% автозаполнение)\nSCR-12 (152-ФЗ)",
            "🔥 Сверхвысокий"
        ],
        [
            "TG-18",
            "3. CRM-интеграторы и IT",
            "Чат интеграторов Битрикс24",
            "https://t.me/b24_integrators",
            "Внедренцы корпоративных порталов и CRM (7k+ специалистов)",
            "Партнёрство: допродажа RevOps OS действующей клиентской базе интегратора.",
            "SCR-13 (Партнёрский оффер)\nSCR-06 (CRM-интеграция)",
            "🔥 Сверхвысокий"
        ],
        [
            "TG-19",
            "3. CRM-интеграторы и IT",
            "MANGO OFFICE | Телефония и CRM",
            "https://t.me/mango_office_b2b",
            "Компании с активными исходящими и входящими звонками (18k+ пользователей)",
            "Оффер поверх телефонии Mango: мгновенная аналитика без тяжелой настройки.",
            "SCR-06 (CRM и телефония)\nSCR-08 (Презентация)",
            "⭐ Высокий"
        ],
        [
            "TG-20",
            "3. CRM-интеграторы и IT",
            "UIS & CoMagic | Телефония и маркетинг",
            "https://t.me/uis_telecom",
            "Директора по продажам и маркетингу (15k+ специалистов)",
            "Показ преимуществ RevOps OS: 14 900 ₽ пилот против сотен тысяч за тяжелые системы.",
            "SCR-02 (РОПу)\nSCR-11 (Возражения)",
            "⭐ Высокий"
        ],

        # Категория 4: Отраслевые B2B-сообщества
        [
            "TG-21",
            "4. Отраслевой B2B-сектор",
            "B2B Оптовики России & Поставщики",
            "https://t.me/b2b_opt_russia",
            "Коммерческие директора оптовых компаний, дистрибьюторы (28k+ участников)",
            "Прямой аутрич коммерческим директорам оптовых баз и дистрибьюторов.",
            "SCR-04 (Крупный чек B2B)\nSCR-01 (Собственнику)",
            "🔥 Сверхвысокий"
        ],
        [
            "TG-22",
            "4. Отраслевой B2B-сектор",
            "Промышленность и Заводы РФ",
            "https://t.me/prom_rus_industry",
            "Директора заводов, начальники отделов сбыта (19k+ руководителей)",
            "Разборы сливов сложных производственных сделок при согласовании ТЗ.",
            "SCR-04 (Производство)\nSCR-01 (Собственнику)",
            "🔥 Сверхвысокий"
        ],
        [
            "TG-23",
            "4. Отраслевой B2B-сектор",
            "Спецтехника и Коммерческий транспорт",
            "https://t.me/specteh_b2b",
            "Владельцы автосалонов спецтехники, РОПы по грузовикам (16k+ дилеров)",
            "Аутрич с упором на высокий чек (>2 млн ₽) и окупаемость пилота с 1 звонка.",
            "SCR-04 (Отраслевой чек)\nSCR-07 (Прием записей)",
            "🔥 Сверхвысокий"
        ],
        [
            "TG-24",
            "4. Отраслевой B2B-сектор",
            "Девелопмент и Недвижимость | Застройщики",
            "https://t.me/proptech_dev",
            "Директора по продажам застройщиков, РОПы агентств (24k+ специалистов)",
            "Аудит звонков с рекламы, спасение лидов за 60 секунд до ухода к конкуренту.",
            "SCR-01 (Оцифровка сливов)\nSCR-08 (Презентация)",
            "🔥 Сверхвысокий"
        ],
        [
            "TG-25",
            "4. Отраслевой B2B-сектор",
            "Логистика и ВЭД Россия",
            "https://t.me/logistics_russia_b2b",
            "Руководители логистических операторов и брокеров (32k+ специалистов)",
            "Аутрич: спасение брошенных заявок на контейнерные перевозки и автодоставку.",
            "SCR-04 (Отраслевой)\nSCR-02 (РОПу)",
            "⭐ Высокий"
        ],
        [
            "TG-26",
            "4. Отраслевой B2B-сектор",
            "Стройматериалы Оптом & Закупки",
            "https://t.me/stroy_opt_b2b",
            "Поставщики цемента, металлопроката, сухих смесей (21k+ директоров)",
            "Экспресс-тест 3 вчерашних звонков прорабам и девелоперам.",
            "SCR-04 (Опт и стройка)\nSCR-07 (Разбор)",
            "⭐ Высокий"
        ],
        [
            "TG-27",
            "4. Отраслевой B2B-сектор",
            "Ассоциация Digital-агентств и IT-услуг",
            "https://t.me/digital_agencies_ru",
            "Владельцы IT-агентств, веб-студий, заказной разработки (14k+ участников)",
            "100% автозаполнение CRM, формирование резюме встреч для проджектов.",
            "SCR-06 (CRM и IT)\nSCR-03 (Блиц)",
            "⭐ Высокий"
        ],

        # Категория 5: Каналы вакансий РОПов и Сейлзов (Сверхгорячие лиды!)
        [
            "TG-28",
            "5. Горячие Вакансии РОПов",
            "Вакансии РОП и Коммерческий директор",
            "https://t.me/job_sales_rop",
            "Собственники бизнеса и HRD, прямо сейчас ищущие РОПа (35k+ подп.)",
            "Прямой контакт автору вакансии: «Пока ищете РОПа, подключите ИИ за 14.9k, чтобы отдел не сливал лидов». Конверсия во встречу >30%!",
            "SCR-01 (Собственнику)\nSCR-03 (Ультра-блиц)",
            "🔥 Сверхвысокий (HOT)"
        ],
        [
            "TG-29",
            "5. Горячие Вакансии РОПов",
            "Работа в продажах | Headhunter Sales",
            "https://t.me/sales_jobs_russia",
            "Компании, расширяющие отдел продаж прямо сейчас (48k+ подп.)",
            "Аутрич фаундерам, нанимающим 3+ менеджеров: контроль адаптации новичков.",
            "SCR-01 (Собственнику)\nSCR-02 (РОПу)",
            "🔥 Сверхвысокий (HOT)"
        ],
        [
            "TG-30",
            "5. Горячие Вакансии РОПов",
            "Топ-менеджеры | Вакансии C-level",
            "https://t.me/clevel_jobs_ru",
            "Крупные компании, нанимающие директоров по продажам (25k+ подп.)",
            "Выход на совет директоров и собственников с предложением аудита звонков.",
            "SCR-01 (Удар по деньгам)\nSCR-06 (CRM-директору)",
            "⭐ Высокий"
        ],
        [
            "TG-31",
            "5. Горячие Вакансии РОПов",
            "GeekJobs Sales & Marketing",
            "https://t.me/geekjobs_sales",
            "Технологические бизнесы в стадии быстрого роста (19k+ подп.)",
            "Выход на фаундеров технологических бизнесов с предложением AI-ROP.",
            "SCR-06 (CRM-фокус)\nSCR-03 (Блиц)",
            "⭐ Высокий"
        ],
        [
            "TG-32",
            "5. Горячие Вакансии РОПов",
            "Вакансии amoCRM и Битрикс24",
            "https://t.me/crm_jobs_ru",
            "Компании, ищущие CRM-интеграторов и аналитиков (15k+ подп.)",
            "Предложение готового ИИ-супервайзера без переделки CRM и найма аналитиков.",
            "SCR-06 (100% автозаполнение)\nSCR-12 (152-ФЗ)",
            "⭐ Высокий"
        ],

        # Категория 6: B2B Нетворкинг и Лидогенерация
        [
            "TG-33",
            "6. B2B Нетворкинг и Лидген",
            "B2B Нетворкинг | Партнеры и Клиенты",
            "https://t.me/b2b_networking_ru",
            "Предприниматели в поиске партнеров и клиентов (42k+ участников)",
            "Еженедельная самопрезентация Дмитрия Федотова с оффером 3 бесплатных аудитов.",
            "SCR-01 (Собственнику)\nSCR-07 (Прием файлов)",
            "⭐ Высокий"
        ],
        [
            "TG-34",
            "6. B2B Нетворкинг и Лидген",
            "Бизнес-чат Предпринимателей Москвы и СПб",
            "https://t.me/biz_chat_msk_spb",
            "Владельцы действующего бизнеса двух столиц (28k+ участников)",
            "Ответы на запросы предпринимателей о падении конверсий и контроле звонков.",
            "SCR-01 (Оцифровка потерь)\nSCR-03 (Блиц)",
            "⭐ Высокий"
        ],
        [
            "TG-35",
            "6. B2B Нетворкинг и Лидген",
            "Клуб Директоров по маркетингу и продажам",
            "https://t.me/cmo_cro_club",
            "CRO, CMO, коммерческие директора (17k+ участников)",
            "Публикация кейсов снижения CAC и роста LTV через контроль звонков нейросетью.",
            "SCR-02 (РОПу)\nSCR-09 (Кейс +18%)",
            "⭐ Высокий"
        ],
        [
            "TG-36",
            "6. B2B Нетворкинг и Лидген",
            "Лидогенерация и B2B Трафик",
            "https://t.me/leadgen_b2b_pro",
            "Эксперты по воронкам и сквозной аналитике (22k+ участников)",
            "Партнерские кросс-промо, обмен лидами, совместные вебинары по сквозной аналитике.",
            "SCR-06 (CRM и аналитика)\nSCR-13 (Партнерка)",
            "⭐ Высокий"
        ]
    ]

    # Section 3: Actionable Strategy Block
    sec3_banner = [
        ["", "", "", "", "", "", "", ""],
        ["🚀 РАЗДЕЛ 3: ПОШАГОВЫЙ ПЛАН ЛИДОГЕНЕРАЦИИ В TELEGRAM (5–15 ЛИДОВ В НЕДЕЛЮ БЕЗ БАНА)", "", "", "", "", "", "", ""]
    ]
    sec3_headers = [
        "№",
        "Название стратегии",
        "Суть механики и триггер действия",
        "Каналы для работы",
        "Конверсионное действие",
        "Ожидаемая конверсия (KPI)",
        "Рекомендуемый шаблон",
        "Статус запуска"
    ]
    sec3_data = [
        [
            "STR-01",
            "🔥 Перехват вакансий РОПа\n(Самый горячий лидген)",
            "Мониторим каналы TG-28...TG-32. Когда собственник ищет РОПа, пишем ему в ЛС: «Пока ищете РОПа (1–2 мес), подключите ИИ за 14.9k, чтобы отдел не сливал сделки».",
            "TG-28, TG-29, TG-30, TG-32",
            "Предложение бесплатного аудита 3 вчерашних звонков кандидатов или менеджеров",
            "25–35% конверсия в диалог,\n2–3 пилота в неделю",
            "SCR-01 (Собственнику)\nSCR-03 (Блиц)",
            "🚀 Готово к запуску"
        ],
        [
            "STR-02",
            "🎯 Экспертный комментинг\n(Входящий поток на личный бренд)",
            "Включаем уведомления на каналы TG-01 (Гребенюк), TG-03 (Батырев), TG-04 (Уколова). Пишем экспертные разборы с цифрами под новыми постами в первые 10 минут.",
            "TG-01, TG-03, TG-04, TG-08",
            "Переход в профиль Дмитрия Федотова (в описании: «Оцифрую сливы в отделе продаж за 20 мин -> @dm1918»)",
            "10–20 переходов в профиль в день,\n3–5 входящих лидов в неделю",
            "Экспертный разбор темы поста",
            "🚀 Готово к запуску"
        ],
        [
            "STR-03",
            "💬 Точечный аутрич участников чатов\n(Прямой исходящий контакт)",
            "Парсим активных участников профильных чатов (TG-02, TG-07, TG-11). Пишем персонализированное сообщение с предложением бесплатного теста 3 звонков.",
            "TG-02, TG-07, TG-11, TG-21",
            "Отправка 3 аудиозаписей прямо в Telegram-чат Дмитрию (@dm1918)",
            "10–15% конверсия в аудит\n(25 диалогов/день = 2-3 аудита)",
            "SCR-01 / SCR-02\nSCR-07 (Прием файлов)",
            "🚀 Готово к запуску"
        ],
        [
            "STR-04",
            "🤝 Партнёрская сеть интеграторов\n(Масштабируемый LTV-канал)",
            "Пишем руководителям интеграторов amoCRM / Bitrix24 (TG-16, TG-18). Предлагаем 25–30% рекуррентной комиссии с каждого внедрения их клиентам.",
            "TG-15, TG-16, TG-17, TG-18",
            "Отправка 1-страничной партнерской презентации + звонок на 15 мин",
            "5 интеграторов = 15–25 пилотов в месяц на автопилоте",
            "SCR-13 (Партнёрский аутрич)\nSCR-06 (CRM-интеграторам)",
            "🚀 Готово к запуску"
        ]
    ]

    all_rows = (
        title_block + 
        sec1_banner + [sec1_headers] + sec1_data + 
        sec2_banner + [sec2_headers] + sec2_data + 
        sec3_banner + [sec3_headers] + sec3_data
    )

    ws.update(values=all_rows, range_name="A1")
    print(f"Data written successfully! Total rows: {len(all_rows)}")

    # Calculate row indices for formatting
    row_title = 0
    row_subtitle = 1
    row_sec1_banner = 3
    row_sec1_header = 4
    sec1_start = 5
    sec1_end = sec1_start + len(sec1_data)

    row_sec2_spacer = sec1_end
    row_sec2_banner = sec1_end + 1
    row_sec2_header = sec1_end + 2
    sec2_start = row_sec2_header + 1
    sec2_end = sec2_start + len(sec2_data)

    row_sec3_spacer = sec2_end
    row_sec3_banner = sec2_end + 1
    row_sec3_header = sec2_end + 2
    sec3_start = row_sec3_header + 1
    sec3_end = sec3_start + len(sec3_data)

    requests = [
        # Merges for Title Block & Banners
        {"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 8}, "mergeType": "MERGE_ALL"}},
        {"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": 8}, "mergeType": "MERGE_ALL"}},
        {"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": row_sec1_banner, "endRowIndex": row_sec1_banner + 1, "startColumnIndex": 0, "endColumnIndex": 8}, "mergeType": "MERGE_ALL"}},
        {"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": row_sec2_banner, "endRowIndex": row_sec2_banner + 1, "startColumnIndex": 0, "endColumnIndex": 8}, "mergeType": "MERGE_ALL"}},
        {"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": row_sec3_banner, "endRowIndex": row_sec3_banner + 1, "startColumnIndex": 0, "endColumnIndex": 8}, "mergeType": "MERGE_ALL"}},

        # 1. Main Title formatting
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 8},
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
        # 2. Subtitle formatting
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": 8},
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

        # 3. Section Banners formatting (Dark Navy with Emerald accent)
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": row_sec1_banner, "endRowIndex": row_sec1_banner + 1, "startColumnIndex": 0, "endColumnIndex": 8},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 0.08, "green": 0.14, "blue": 0.25},
                        "textFormat": {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "fontSize": 11, "bold": True},
                        "horizontalAlignment": "LEFT",
                        "verticalAlignment": "MIDDLE"
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
            }
        },
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": row_sec2_banner, "endRowIndex": row_sec2_banner + 1, "startColumnIndex": 0, "endColumnIndex": 8},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 0.08, "green": 0.14, "blue": 0.25},
                        "textFormat": {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "fontSize": 11, "bold": True},
                        "horizontalAlignment": "LEFT",
                        "verticalAlignment": "MIDDLE"
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
            }
        },
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": row_sec3_banner, "endRowIndex": row_sec3_banner + 1, "startColumnIndex": 0, "endColumnIndex": 8},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 0.08, "green": 0.14, "blue": 0.25},
                        "textFormat": {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "fontSize": 11, "bold": True},
                        "horizontalAlignment": "LEFT",
                        "verticalAlignment": "MIDDLE"
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
            }
        },

        # 4. Table Header formatting (Navy Slate with bold white text)
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": row_sec1_header, "endRowIndex": row_sec1_header + 1, "startColumnIndex": 0, "endColumnIndex": 8},
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
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": row_sec2_header, "endRowIndex": row_sec2_header + 1, "startColumnIndex": 0, "endColumnIndex": 8},
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
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": row_sec3_header, "endRowIndex": row_sec3_header + 1, "startColumnIndex": 0, "endColumnIndex": 8},
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

        # 5. Default text formatting for all data rows
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": sec1_start, "endRowIndex": sec1_end, "startColumnIndex": 0, "endColumnIndex": 8},
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
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": sec2_start, "endRowIndex": sec2_end, "startColumnIndex": 0, "endColumnIndex": 8},
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
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": sec3_start, "endRowIndex": sec3_end, "startColumnIndex": 0, "endColumnIndex": 8},
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

        # Column A formatting (Codes: bold emerald, centered)
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": sec1_start, "endRowIndex": sec1_end, "startColumnIndex": 0, "endColumnIndex": 1},
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
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": sec2_start, "endRowIndex": sec2_end, "startColumnIndex": 0, "endColumnIndex": 1},
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
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": sec3_start, "endRowIndex": sec3_end, "startColumnIndex": 0, "endColumnIndex": 1},
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

        # Column Widths (A to H)
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 95}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2}, "properties": {"pixelSize": 220}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 3}, "properties": {"pixelSize": 240}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 3, "endIndex": 4}, "properties": {"pixelSize": 220}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 4, "endIndex": 5}, "properties": {"pixelSize": 300}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 5, "endIndex": 6}, "properties": {"pixelSize": 320}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 6, "endIndex": 7}, "properties": {"pixelSize": 260}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 7, "endIndex": 8}, "properties": {"pixelSize": 180}, "fields": "pixelSize"}},

        # Fixed Row Heights for headers and banners
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 52}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": 1, "endIndex": 2}, "properties": {"pixelSize": 30}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": 2, "endIndex": 3}, "properties": {"pixelSize": 16}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": row_sec1_banner, "endIndex": row_sec1_banner + 1}, "properties": {"pixelSize": 36}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": row_sec1_header, "endIndex": row_sec1_header + 1}, "properties": {"pixelSize": 46}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": row_sec2_banner, "endIndex": row_sec2_banner + 1}, "properties": {"pixelSize": 36}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": row_sec2_header, "endIndex": row_sec2_header + 1}, "properties": {"pixelSize": 46}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": row_sec3_banner, "endIndex": row_sec3_banner + 1}, "properties": {"pixelSize": 36}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": row_sec3_header, "endIndex": row_sec3_header + 1}, "properties": {"pixelSize": 46}, "fields": "pixelSize"}},

        # Freeze top 5 rows (Title + Sec 1 banner + headers)
        {"updateSheetProperties": {"properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 5}}, "fields": "gridProperties.frozenRowCount"}}
    ]

    # Zebra shading for data rows
    for r_idx in range(sec1_start, sec1_end):
        bg = {"red": 0.97, "green": 0.98, "blue": 0.99} if (r_idx % 2 == 1) else {"red": 1.0, "green": 1.0, "blue": 1.0}
        requests.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": r_idx, "endRowIndex": r_idx + 1, "startColumnIndex": 0, "endColumnIndex": 8}, "cell": {"userEnteredFormat": {"backgroundColor": bg}}, "fields": "userEnteredFormat(backgroundColor)"}})

    for r_idx in range(sec2_start, sec2_end):
        bg = {"red": 0.97, "green": 0.98, "blue": 0.99} if (r_idx % 2 == 1) else {"red": 1.0, "green": 1.0, "blue": 1.0}
        requests.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": r_idx, "endRowIndex": r_idx + 1, "startColumnIndex": 0, "endColumnIndex": 8}, "cell": {"userEnteredFormat": {"backgroundColor": bg}}, "fields": "userEnteredFormat(backgroundColor)"}})

    for r_idx in range(sec3_start, sec3_end):
        bg = {"red": 0.97, "green": 0.98, "blue": 0.99} if (r_idx % 2 == 1) else {"red": 1.0, "green": 1.0, "blue": 1.0}
        requests.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": r_idx, "endRowIndex": r_idx + 1, "startColumnIndex": 0, "endColumnIndex": 8}, "cell": {"userEnteredFormat": {"backgroundColor": bg}}, "fields": "userEnteredFormat(backgroundColor)"}})

    # Borders for each section table
    requests.extend([
        {
            "updateBorders": {
                "range": {"sheetId": sheet_id, "startRowIndex": row_sec1_header, "endRowIndex": sec1_end, "startColumnIndex": 0, "endColumnIndex": 8},
                "top": {"style": "SOLID", "color": {"red": 0.7, "green": 0.75, "blue": 0.82}},
                "bottom": {"style": "SOLID", "color": {"red": 0.7, "green": 0.75, "blue": 0.82}},
                "left": {"style": "SOLID", "color": {"red": 0.7, "green": 0.75, "blue": 0.82}},
                "right": {"style": "SOLID", "color": {"red": 0.7, "green": 0.75, "blue": 0.82}},
                "innerHorizontal": {"style": "SOLID", "color": {"red": 0.85, "green": 0.88, "blue": 0.92}},
                "innerVertical": {"style": "SOLID", "color": {"red": 0.85, "green": 0.88, "blue": 0.92}}
            }
        },
        {
            "updateBorders": {
                "range": {"sheetId": sheet_id, "startRowIndex": row_sec2_header, "endRowIndex": sec2_end, "startColumnIndex": 0, "endColumnIndex": 8},
                "top": {"style": "SOLID", "color": {"red": 0.7, "green": 0.75, "blue": 0.82}},
                "bottom": {"style": "SOLID", "color": {"red": 0.7, "green": 0.75, "blue": 0.82}},
                "left": {"style": "SOLID", "color": {"red": 0.7, "green": 0.75, "blue": 0.82}},
                "right": {"style": "SOLID", "color": {"red": 0.7, "green": 0.75, "blue": 0.82}},
                "innerHorizontal": {"style": "SOLID", "color": {"red": 0.85, "green": 0.88, "blue": 0.92}},
                "innerVertical": {"style": "SOLID", "color": {"red": 0.85, "green": 0.88, "blue": 0.92}}
            }
        },
        {
            "updateBorders": {
                "range": {"sheetId": sheet_id, "startRowIndex": row_sec3_header, "endRowIndex": sec3_end, "startColumnIndex": 0, "endColumnIndex": 8},
                "top": {"style": "SOLID", "color": {"red": 0.7, "green": 0.75, "blue": 0.82}},
                "bottom": {"style": "SOLID", "color": {"red": 0.7, "green": 0.75, "blue": 0.82}},
                "left": {"style": "SOLID", "color": {"red": 0.7, "green": 0.75, "blue": 0.82}},
                "right": {"style": "SOLID", "color": {"red": 0.7, "green": 0.75, "blue": 0.82}},
                "innerHorizontal": {"style": "SOLID", "color": {"red": 0.85, "green": 0.88, "blue": 0.92}},
                "innerVertical": {"style": "SOLID", "color": {"red": 0.85, "green": 0.88, "blue": 0.92}}
            }
        }
    ])

    # Dynamic row heights calculation for Section 1 (ICP rows)
    for idx, item in enumerate(sec1_data):
        r_num = sec1_start + idx
        # measure lines across columns
        lines_b = item[1].count("\n") + 1
        lines_d = item[3].count("\n") + 1
        lines_e = sum(max(1, len(l)//32 + 1) for l in item[4].split("\n"))
        lines_f = item[5].count("\n") + 1
        lines_g = sum(max(1, len(l)//28 + 1) for l in item[6].split("\n"))
        max_l = max(lines_b, lines_d, lines_e, lines_f, lines_g)
        h = max(80, int(max_l * 19 + 25))
        requests.append({"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": r_num, "endIndex": r_num + 1}, "properties": {"pixelSize": h}, "fields": "pixelSize"}})

    # Dynamic row heights for Section 2 (Telegram channels)
    for idx, item in enumerate(sec2_data):
        r_num = sec2_start + idx
        lines_f = sum(max(1, len(l)//35 + 1) for l in item[5].split("\n"))
        lines_e = sum(max(1, len(l)//30 + 1) for l in item[4].split("\n"))
        max_l = max(lines_f, lines_e, 2)
        h = max(60, int(max_l * 19 + 20))
        requests.append({"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": r_num, "endIndex": r_num + 1}, "properties": {"pixelSize": h}, "fields": "pixelSize"}})

    # Dynamic row heights for Section 3 (Strategies)
    for idx, item in enumerate(sec3_data):
        r_num = sec3_start + idx
        lines_c = sum(max(1, len(l)//32 + 1) for l in item[2].split("\n"))
        lines_e = sum(max(1, len(l)//26 + 1) for l in item[4].split("\n"))
        max_l = max(lines_c, lines_e, 3)
        h = max(75, int(max_l * 19 + 22))
        requests.append({"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": r_num, "endIndex": r_num + 1}, "properties": {"pixelSize": h}, "fields": "pixelSize"}})

    sh.batch_update({"requests": requests})
    print("✓ All formatting, column widths, row heights, and styling applied successfully!")

if __name__ == "__main__":
    main()
