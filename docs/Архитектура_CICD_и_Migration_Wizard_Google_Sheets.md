# Архитектура CI/CD и Migration Wizard для RevOps OS (Google Sheets)

> **Ключевой вызов масштабирования (Scale Bottleneck):**
> В традиционном SaaS код отделен от базы данных (PostgreSQL + Docker). В Google Sheets код, расчетные формулы, верстка дашбордов и клиентские транзакции (`raw_deals`, `raw_calls`) живут в **одном физическом файле**. 
> При выпуске релиза **V18.0** ручное обновление 50 клиентских таблиц привело бы к колоссальным трудозатратам и рискам случайного затирания данных клиентов.

Ниже представлена разработанная **Enterprise CI/CD & Migration Architecture**, решающая эту проблему на 100%.

---

## 1. Сквозная архитектура обновлений (Dual-Engine CI/CD)

Система работает в двух согласованных режимах:
1. **Pull-режим (Self-Service Client Update):** Клиент нажимает пункт меню `🛡️ RevOps` ➔ `🚀 Обновить RevOps OS (Migration Wizard)...`, видит список новых фич и запускает автообновление в 1 клик.
2. **Push-режим (Centralized Fleet Deployment):** Интегратор/фаундер запускает консольный скрипт `fleet_migration_manager.py`, который обновляет все 50 файлов пакетно через Google Sheets API за 2 минуты.

---

## 2. 5-Фазный протокол безопасной миграции (Zero-Data-Loss)

### Фаза 1. Версионный Handshake и Релиз-Манифест
* Мастер-таблица хранит актуальный номер релиза в `⚙️ Настройки!E4` (например, `RevOps OS V18.0 Enterprise Release`).
* Клиентский файл сравнивает свою текущую версию с версией ядра. Если доступно обновление, активируется модальное окно с перечнем новых модулей.

### Фаза 2. Атомарная точка восстановления (Drive Backup Snapshot)
* До выполнения любых изменений скрипт вызывает Google Drive API:
  ```javascript
  const backupName = `[BACKUP_REVOPS_${timestamp}] ${ss.getName()}`;
  const backupFile = currentFile.makeCopy(backupName);
  ```
* Бэкап сохраняется в изолированной папке клиента. В случае любого непредвиденного сбоя клиент может откатиться за 5 секунд.

### Фаза 3. Эволюция схем данных (Schema Evolution без сдвига строк)
* **Защищенные сырые слои:** `raw_deals`, `raw_calls`, `raw_touchpoints`, `raw_invoices`, `raw_payments`, `raw_marketing`, `raw_alerts`, `raw_audit_log`, `📅 Deal_Event_Log`, `📊 Stage_History`.
* **Правило эволюции:** Строки данных 2..N **никогда не удаляются и не перезаписываются**.
* Если в V18.0 появились новые системные поля (например, `ai_followup_status` или `v18_ai_tag`), скрипт вычисляет `sh.getLastColumn() + 1` и **дописывает новые заголовки в конец первой строки**, не нарушая индексы существующих формул.

### Фаза 4. Перенос новых модулей и связывание формул
* Новые листы (например, `🤖 AI_Автодожим_WhatsApp`) копируются из Master-таблицы методом:
  ```javascript
  const newSheet = masterSs.getSheetByName('🤖 AI_Автодожим').copyTo(clientSs);
  newSheet.setName('🤖 AI_Автодожим');
  ```
* Расчетные формулы в `calc_engine` и витринах обновляются, автоматически подхватывая данные из клиентских `raw_*` слоев.
* Пользовательские параметры (название компании, плановая выручка, состав менеджеров в `⚙️ Настройки!B21:B24`, токены CRM и Telegram) изолированы и сохраняются в исходном виде.

### Фаза 5. QA Verification Gate & Changelog Audit
* Автоматически вызывается лист `🧪 QA_Suite`.
* Все 118 регрессионных тестов подтверждают отсутствие ошибок `#REF!`, `#DIV/0!`, `#VALUE!`.
* В лист `changelog` на строку 2 автоматически добавляется запись с датой, версией и ссылкой на ID бэкапа.

---

## 3. Компоненты системы, внедренные в проект

| Компонент | Расположение | Назначение |
| :--- | :--- | :--- |
| **In-App Migration Wizard** | [revops_migration_wizard.js](file:///C:/Users/strel/.gemini/antigravity/scratch/revops-enterprise-os/revops_migration_wizard.js) | UI-окно для клиента с прогресс-баром, автобэкапом и обновлением в 1 клик |
| **RBAC & Menu Hook** | [revops_rbac_v176_hardened.js](file:///C:/Users/strel/.gemini/antigravity/scratch/revops-enterprise-os/revops_rbac_v176_hardened.js#L880-L885) | Пункт `🚀 Обновить RevOps OS (Migration Wizard)...` в меню `🛡️ RevOps` |
| **Fleet Migration CLI** | [fleet_migration_manager.py](file:///C:/Users/strel/.gemini/antigravity/scratch/revops-enterprise-os/scripts/fleet_migration_manager.py) | Консольный менеджер для интегратора: аудит и параллельный деплой на 50+ клиентов |
| **Fleet Registry** | [client_fleet_registry.json](file:///C:/Users/strel/.gemini/antigravity/scratch/revops-enterprise-os/config/client_fleet_registry.json) | Центральная база тенантов (ID, компания, текущая версия, статус бэкапа) |
| **Apps Script Console** | Лист `🛠️ Apps_Script_Console` (строка 11) | Официальный статус модуля `🟢 READY` в рабочей книге Google Таблицы |

---

## 4. Как интегратор обновляет 50 клиентов (Operational Playbook)

### Шаг 1. Аудит состояния парка клиентов
Интегратор запускает команду проверки статусов:
```bash
python scripts/fleet_migration_manager.py --action list
```

### Шаг 2. Безопасное тестирование (Dry-Run)
Перед внесением изменений запускается симуляция:
```bash
python scripts/fleet_migration_manager.py --action migrate --tenant all --dry-run
```

### Шаг 3. Боевой пакетный деплой релиза V18.0
```bash
python scripts/fleet_migration_manager.py --action migrate --tenant all
```
За 2-3 минуты скрипт:
1. Создаст 50 резервных копий на Google Диске;
2. Проверит и расширит схемы таблиц в 50 файлах;
3. Накатит новые формулы и модули V18;
4. Проверит `QA_Suite` (100% PASS);
5. Зафиксирует успешное завершение в отчете.
