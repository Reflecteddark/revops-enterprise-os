"""
Конфигурация ролевых отчетов RevOps Enterprise OS V17.6.
Определяет состав слайдов, метаданные обложки и целевые файлы для каждой роли.
"""

ROLE_CONFIGS = {
    'all': {
        'role_key': 'all',
        'name': 'Полная презентация (Executive Master Suite)',
        'filename': 'RevOps_Executive_Summary_2026-09-30.pdf',
        'root_copy': 'RevOps_Enterprise_OS_V17.6_Executive_Summary.pdf',
        'role_badge': 'RELEASE V17.6 ENTERPRISE',
        'role_subtitle': 'Управленческий дайджест руководителя',
        'role_title': 'RevOps Enterprise OS V17.6',
        'role_desc': 'Архитектура управления выручкой, речевой ИИ-аудит звонков Whisper+LLM и ликвидация 7 узких мест воронки продаж.',
        'role_audience': 'Адресат: Совет Директоров & C-Level Команда • Конфиденциально',
        'expected_pages': 10,
        'slides': [
            'slide_1', 'slide_2', 'slide_3', 'slide_4', 'slide_5',
            'slide_6', 'slide_7', 'slide_8', 'slide_9', 'slide_10'
        ]
    },
    'ceo': {
        'role_key': 'ceo',
        'name': 'Отчет для Генерального Директора (CEO / Собственник)',
        'filename': 'RevOps_Report_CEO.pdf',
        'root_copy': 'RevOps_Report_CEO.pdf',
        'role_badge': 'EXECUTIVE BRIEFING • CEO',
        'role_subtitle': 'Стратегический срез • Совет Директоров & Собственник',
        'role_title': 'RevOps Executive One-Pager & Strategic Revenue Pulse',
        'role_desc': 'Консолидированный дайджест выполнения финансового плана, здоровье пайплайна, сценарное моделирование Монте-Карло и окупаемость спринта внедрения.',
        'role_audience': 'Адресат: Генеральный директор / Акционеры • Конфиденциально',
        'expected_pages': 5,
        'slides': [
            'slide_1', 'slide_2', 'slide_3', 'slide_9', 'slide_10'
        ]
    },
    'rop': {
        'role_key': 'rop',
        'name': 'Пульт РОПа (Head of Sales / Коммерческий директор)',
        'filename': 'RevOps_Report_ROP.pdf',
        'root_copy': 'RevOps_Report_ROP.pdf',
        'role_badge': 'SALES OPERATIONS PULSE • РОП',
        'role_subtitle': 'Операционный Пульт • Управление продажами 15 минут в день',
        'role_title': 'RevOps Sales Control: Пульт РОПа & Аудит Звонков',
        'role_desc': 'Сделки в зоне критического риска (>14 дней), контроль регламентов SLA 48ч, дефекты скриптов по ИИ-аналитике звонков Whisper+LLM.',
        'role_audience': 'Адресат: Руководитель Отдела Продаж (РОП) / Коммерческий Директор',
        'expected_pages': 5,
        'slides': [
            'slide_1', 'slide_4', 'slide_5', 'slide_6', 'slide_7'
        ]
    },
    'cfo': {
        'role_key': 'cfo',
        'name': 'Финансовый срез (CFO / Главный бухгалтер)',
        'filename': 'RevOps_Report_CFO.pdf',
        'root_copy': 'RevOps_Report_CFO.pdf',
        'role_badge': 'CASH-FLOW & DSO RECOVERY • CFO',
        'role_subtitle': 'Финансовый контроль • Cash Flow & DSO Engine',
        'role_title': 'RevOps Financial Recovery: Дебиторка & Платежный Календарь',
        'role_desc': 'Оперативный контроль платежей, динамика DSO, автоматизированный AI-дожим просроченных инвойсов и сценарный стресс-тест ликвидности.',
        'role_audience': 'Адресат: Финансовый Директор (CFO) / Главный Бухгалтер',
        'expected_pages': 5,
        'slides': [
            'slide_1', 'slide_2', 'slide_8', 'slide_9', 'slide_10'
        ]
    }
}


def get_role_context(role_key: str) -> dict:
    """Формирует контекст переменных Jinja2 для заданной роли."""
    if role_key not in ROLE_CONFIGS:
        raise ValueError(f"Неизвестная роль: {role_key}. Допустимые: {list(ROLE_CONFIGS.keys())}")

    cfg = ROLE_CONFIGS[role_key]
    active_slides = cfg['slides']
    total_pages = len(active_slides)

    # Карта номеров страниц (1-based)
    page_numbers = {}
    for idx, slide_id in enumerate(active_slides, 1):
        page_numbers[slide_id] = idx

    return {
        'active_slides': active_slides,
        'total_pages': total_pages,
        'page_numbers': page_numbers,
        'role_key': role_key,
        'role_badge': cfg['role_badge'],
        'role_subtitle': cfg['role_subtitle'],
        'role_title': cfg['role_title'],
        'role_desc': cfg['role_desc'],
        'role_audience': cfg['role_audience'],
    }
