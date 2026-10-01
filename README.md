# RevOps Enterprise OS V17.6

<div align="center">

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Excel](https://img.shields.io/badge/Excel-openpyxl-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
[![Website](https://img.shields.io/badge/Live_Site-ai--rop.ru-2563EB?style=for-the-badge&logo=googlechrome&logoColor=white)](https://ai-rop.ru)
![Security](https://img.shields.io/badge/Compliance-152--ФЗ_Audit_Log-16A34A?style=for-the-badge)
![WCAG](https://img.shields.io/badge/Accessibility-WCAG_2.1_AAA-4F46E5?style=for-the-badge)
![License](https://img.shields.io/badge/License-MIT-blue?style=for-the-badge)

**Автономная операционная система управления выручкой (Revenue Operations), финансовым аудитом и автоматической генерацией C-Level PDF-отчетов.**

[🌐 Живой сайт: ai-rop.ru](https://ai-rop.ru) • [Обзор](#-обзор-системы) • [Архитектура](#-архитектура-решения) • [Веб-раздел](#-веб-сайт-и-маркетинговая-витрина-ai-ropru) • [Компоненты](#-ключевые-компоненты) • [Быстрый старт](#-быстрый-старт) • [Структура репозитория](#-структура-репозитория)

</div>

---

## 🚀 Обзор системы

**RevOps Enterprise OS V17.6** — комплексное программное решение корпоративного класса, трансформирующее сырые данные CRM, телефонии и финансовых транзакций в динамические аналитические витрины, исполнительские дашборды и презентационные PDF-отчеты полиграфического качества.

### Ключевые возможности:
- 📊 **15 витринных аналитических листов** в Excel-ядре: от экспресс-калькулятора до мультиканальной сквозной атрибуции и симулятора Монте-Карло.
- 🛡️ **Role-Based Access Control (RBAC)**: строгое разграничение прав доступа (CEO, РОП, CFO, CMO, TeamLead, Auditor, Admin) на уровне формул Excel и скриптов Google Apps Script.
- 🔍 **Pre-flight Валидатор Данных (`data_validator.py`)**: мгновенная проверка 8 критических ячеек и формул на ошибки (`#REF!`, `#VALUE!`, `#DIV/0!`), аномальные диапазоны и целостность пайплайна перед рендером.
- 🖨️ **Автономный рендеринг PDF через Chrome Headless**: печать векторных SVG-графиков, премиальной типографики (`Playfair Display`, `Inter`) и альбомной верстки формата A4 без водяных знаков и зависимости от сторонних платных API.
- 📋 **Корпоративный Audit Log (`audit_log.py`)**: соответствие 152-ФЗ РФ, разделитель `;`, кодировка `UTF-8-SIG` (BOM), автоматическая ротация архива при превышении 10 000 записей.
- 💳 **Stripe B2B Checkout Design System**: эргономика оформления сделок с фокусным Primary CTA (`#4F46E5`), математическим контрастом WCAG AAA (15.8:1) и просторными строками.

---

## 🏛️ Архитектура решения

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                       REVOPS ENTERPRISE OS V17.6                            │
├─────────────────────────────────────────────────────────────────────────────┤
│ 1. DATA SOURCE LAYER                                                        │
│    • RevOps Platform V17.6 (RBAC Production Suite).xlsx                     │
│    • CRM Raw Deals, Call Recordings, Unit Economics Data                    │
├─────────────────────────────────────────────────────────────────────────────┤
│ 2. PRE-FLIGHT VALIDATION & AUDIT ENGINE                                      │
│    • data_validator.py (OpenPyXL, Trapping #REF!/#VALUE!, Pipeline Integrity│
│    • audit_log.py (Compliance 152-FZ, UTF-8-BOM, Timestamp, Role Audit)    │
├─────────────────────────────────────────────────────────────────────────────┤
│ 3. TEMPLATE & VISUAL PRESENTATION ENGINE                                    │
│    • presentation/build.py & generate_presentation_html.py                  │
│    • Dynamic SVG Generation (Pipeline funnel, Waterfall, Monte Carlo)       │
│    • HTML5 / CSS Paged Media Layout (297mm x 210mm Landscape)               │
├─────────────────────────────────────────────────────────────────────────────┤
│ 4. HEADLESS EXPORT & RUNTIME DISTRIBUTION                                   │
│    • Chrome Headless (--print-to-pdf, Fallback Registry/Path Discovery)     │
│    • Automatic Artifact Mirroring (Desktop, Brain Persistent Cache)         │
│    • Batch Launchers (.bat for C-Level instant execution)                   │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## 📦 Ключевые компоненты

### 1. Pre-flight Data Validator (`data_validator.py`)
Гарантирует целостность входных данных перед запуском рендера PDF:
- **Лист `⚙️ Настройки`**: непустое название компании (B3), положительный реалистичный план ($0 < B9 < 10^9$), валидный диапазон дат ($B6 < B7$).
- **Лист `⚡ Экспресс_Калькулятор_3_Цифры`**: лиды целое $1..100\,000$ (B5), чек $10\,000..10\,000\,000$ ₽ (B6), штат менеджеров $1..100$ (B7).
- **Лист `raw_deals`**: минимум 1 закрытая сделка на этапе 6 (Closed-Won), минимум 1 сделка в активном пайплайне (этапы 1–5), отсутствие отрицательных сумм и пустых идентификаторов.
- **Лист `📄 Executive_OnePager`**: выручка Closed-Won $A5 \ge 0$ с поддержкой незакэшированных формул openpyxl.
- **Детекция битых формул**: перехват `#REF!`, `#VALUE!`, `#DIV/0!`, `#N/A`, `#NAME?` на лету.

### 2. Модуль Аудита Генераций (`audit_log.py`)
Фиксирует факты доступа к финансовым отчетам в `audit_log.csv`:
- Поля: `timestamp`, `user_login`, `role`, `report_type`, `duration_sec`, `status`, `error_message`, `pdf_size_kb`.
- Гарантированная совместимость с MS Excel в русской локали (BOM `\xef\xbb\xbf`, разделитель `;`).
- Отказоустойчивость: сбой аудита логируется в `stderr`, не прерывая генерацию клиенту.

### 3. Генераторы отчетов PDF (`generate_full_report_pdf.py` & `presentation/build.py`)
- **Full Report (15 витрин)**: полный управленческий срез компании (OnePager, Пульс, Пульт РОПа, 7 грехов ОП, Экспресс, ИИ-Аудит, Атрибуция, Финансы, Мотивация, Симулятор, Воронка, Юнит-экономика, Радар алертов, Ресурсный план).
- **Executive Summary (10 слайдов)**: лаконичный дайджест для генерального директора и совета директоров с векторной инфографикой.
- **Динамический футер**: `Данные на {DD.MM.YYYY} {HH:MM} MSK • Сгенерировано за {duration} сек`.

---

## ⚡ Быстрый старт

### Требования
- Python 3.10+
- Google Chrome (установлен в стандартном каталоге Windows или доступен в `PATH`)

### Установка

```bash
git clone https://github.com/Reflecteddark/revops-enterprise-os.git
cd revops-enterprise-os
pip install -r requirements.txt
```

### Запуск генерации отчетов

1. **Сформировать Полный Отчет (15 витрин)**:
   ```bash
   python generate_full_report_pdf.py
   ```
   *Готовый PDF появится в папке проекта и на Рабочем столе.*

2. **Сформировать Executive Summary (10 слайдов)**:
   ```bash
   python presentation/build.py
   ```

3. **Сформировать ВСЕ отчеты одной командой**:
   ```bash
   python generate_all.py
   ```

4. **Запуск через ярлыки без консоли (Windows)**:
   В папке `launchers/` подготовлены готовые `.bat`-файлы запуска в один клик:
   - `📊 Сформировать Executive Summary (10 слайдов).bat`
   - `📑 Сформировать Полный Отчет (15 витрин).bat`
   - `🚀 Сформировать ВСЕ Отчеты RevOps.bat`

---

## 🌐 Веб-сайт и Маркетинговая Витрина ([ai-rop.ru](https://ai-rop.ru))

Публичный веб-сайт платформы полностью вынесен в **отдельный изолированный раздел [`web/`](web/)**:
- 🔗 **Официальный домен**: **[https://ai-rop.ru](https://ai-rop.ru)**
- 🎛️ **Интерактивный экспресс-калькулятор**: клиентское моделирование упущенной прибыли и сценариев роста (+15%, +30%, +50%) за 30 секунд.
- 📥 **Скачивание презентационного отчета**: прямая выдача эталонной таблицы `RevOps_Platform_Demo_Sample.xlsx` с **13 критериями речевой аналитики** и «5 решениями месяца».
- 🤖 **Лид-магнит тест-драйва**: интеграция с Telegram-ботом [@RevOps_Super_Audit_Bot](https://t.me/RevOps_Super_Audit_Bot) для разбора 3 звонков.
- 🚀 **Автодеплой и CI/CD**: автоматическая публикация раздела `web/` на GitHub Pages через GitHub Actions при push в `main`.

---

## 🔐 Безопасность и Аудит (152-ФЗ)

- **Zero-Cloud Reporting**: все вычисления, парсинг Excel и рендер PDF происходят строго локально на машине или в защищенном контуре. Финансовые данные не передаются на сторонние серверы.
- **Ролевая матрица (RBAC)**: жесткая изоляция закрытых листов для непривилегированных ролей через макросы Google Apps Script (`revops_rbac_v176_hardened.js`).
- **Журналирование действий**: факт каждой генерации с точностью до секунды протоколируется для ИБ-комплаенса.

---

## 📂 Структура репозитория

```
revops-enterprise-os/
│
├── web/                                              # 🌐 ИЗОЛИРОВАННЫЙ РАЗДЕЛ ВЕБ-САЙТА (ai-rop.ru)
│   ├── index.html                                    # Главная страница (лендинг платформы с ROI-калькулятором)
│   ├── CNAME                                         # Привязка собственного домена ai-rop.ru
│   ├── .nojekyll                                     # Отключение Jekyll для чистого HTML/JS
│   ├── RevOps_Platform_Demo_Sample.xlsx              # Презентационный Excel-образец (13 критериев ИИ)
│   └── README.md                                     # Документация раздела веб-сайта
│
├── speech_engine.py                                  # Модуль речевой аналитики Whisper + 152-ФЗ Sanitizer
├── amocrm_connector.py                               # Промышленный коннектор к amoCRM (SLA, батчинг)
├── bitrix24_connector.py                             # Промышленный коннектор к Битрикс24
├── telegram_bot.py                                   # Telegram-бот супервайзера @RevOps_Super_Audit_Bot
├── run_express_audit.py                              # Движок экспресс-аудита 3 звонков
├── google_sheets_sync.py                             # Двусторонняя синхронизация с Google Sheets
│
├── RevOps Platform V17.6 (RBAC Production Suite).xlsx # Мастер-книга с витринами
├── generate_full_report_pdf.py                       # Движок рендера полного отчета
├── data_validator.py                                 # Pre-flight валидация данных
├── audit_log.py                                      # Модуль аудита генераций
├── generate_all.py                                   # Оркестратор сборки всех отчетов
├── generate_presentation_html.py                     # Генератор HTML-слайдов
├── revops_rbac_v176_hardened.js                      # Защищенный скрипт RBAC для Sheets
│
├── presentation/                                     # Модуль сборки Executive Summary
│   ├── build.py                                      # Сборщик презентации
│   ├── build_presentation.py                         # Интеграционный скрипт
│   └── templates/                                    # HTML/CSS-шаблоны слайдов
│
├── launchers/                                        # Готовые батники для Windows
│   ├── 📊 Сформировать Executive Summary (10 слайдов).bat
│   ├── 📑 Сформировать Полный Отчет (15 витрин).bat
│   └── 🚀 Сформировать ВСЕ Отчеты RevOps.bat
│
├── docs/                                             # Юридический контур, регламенты и ТЗ
│   ├── 152-ФЗ_Комплект_Безопасности_RevOps_AI.docx  # 5 юридических документов по 152-ФЗ
│   ├── Боевой_Протокол_Клиент_За_7_Дней.docx        # Внутренний регламент фаундера
│   ├── Дорожная_Карта_Пилота_Для_Клиента_7_Дней.docx # Клиентская 1-page памятка
│   ├── Партнерское_Предложение_Интеграторам_CRM.docx # Оффер для интеграторов CRM
│   └── ТЗ_Архитектура_RevOps_Bot_Reporting_Engine.docx # Официальное ТЗ на внедрение
│
├── samples/                                          # Примеры сгенерированных отчетов
│   ├── RevOps_Enterprise_OS_V17.6_Executive_Summary.pdf
│   └── RevOps_Enterprise_OS_V17.6_Full_Report.pdf
│
├── .github/workflows/                                # CI/CD Автодеплой
│   └── deploy-pages.yml                              # Автоматический деплой web/ на ai-rop.ru
│
├── requirements.txt                                  # Зависимости Python
├── .gitignore                                        # Исключения Git
├── LICENSE                                           # Лицензия MIT
└── README.md                                         # Документация проекта
```

---

## 📄 Лицензия

Проект распространяется под лицензией [MIT](LICENSE).

Автор: **Dmitriy Fedotov** ([@Reflecteddark](https://github.com/Reflecteddark))
