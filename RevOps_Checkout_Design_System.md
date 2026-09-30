# RevOps Checkout & Revenue Conversion Design System (V17.6 Master Specification)

> **Статус документа**: Финальный нормативный стандарт (Production Ready)  
> **Основа системы**: Принципы B2B-конверсии Stripe, Linear, Mercury Bank  
> **Стандарты доступности**: Соответствие WCAG 2.1 & 2.2 (Уровни AA и AAA)  
> **Целевая аудитория**: Enterprise-заказчики, C-Level, финдиректора, коммерческие директора  

---

## 1. Математический аудит контрастности WCAG 2.1 / 2.2

Аудит проведен программным расчетом относительной яркости (Relative Luminance) по алгоритму W3C WCAG:
$$\text{Luminance } L = 0.2126 \cdot R_{lin} + 0.7152 \cdot G_{lin} + 0.0722 \cdot B_{lin}$$
$$\text{Contrast Ratio } CR = \frac{L_{lighter} + 0.05}{L_{darker} + 0.05}$$

### Нормативы WCAG:
- **Уровень AA (Обычный текст < 18pt / < 14pt bold)**: $CR \ge 4.5:1$
- **Уровень AA (Крупный текст $\ge$ 18pt / $\ge$ 14pt bold)**: $CR \ge 3.0:1$
- **Уровень AA (Графические элементы и индикаторы фокуса)**: $CR \ge 3.0:1$
- **Уровень AAA (Обычный текст)**: $CR \ge 7.0:1$

### Сводная матрица аудита всех пар цветов системы:

| Название элемента | Цвет переднего плана | Фон | Контраст | WCAG AA Normal | WCAG AA Large / UI | WCAG AAA | Вердикт и статус |
| :--- | :--- | :--- | :---: | :---: | :---: | :---: | :--- |
| **Основной текст** (`text-primary`) | `#0F172A` | `#FFFFFF` | **17.85:1** | **PASS** | **PASS** | **PASS** | Превосходит AAA более чем в 2.5 раза |
| **Основной текст на холсте** | `#0F172A` | `#F8FAFC` | **17.06:1** | **PASS** | **PASS** | **PASS** | Идеальная читаемость на светлом фоне |
| **Основной текст на подложке** | `#0F172A` | `#F1F5F9` | **16.30:1** | **PASS** | **PASS** | **PASS** | Четкое считывание в строках таблиц |
| **Второстепенный текст** (`text-secondary`) | `#64748B` | `#FFFFFF` | **4.76:1** | **PASS** | **PASS** | FAIL | Полный Pass AA для меток и подсказок |
| **Второстепенный текст на холсте** | `#64748B` | `#F8FAFC` | **4.55:1** | **PASS** | **PASS** | FAIL | Соответствует порогу 4.5:1 |
| **Неактивный текст** (`text-disabled`) | `#94A3B8` | `#FFFFFF` | **2.56:1** | Исключение | Исключение | FAIL | WCAG 1.4.3 исключает disabled-поля |
| **Текст кнопки CTA** (`text-inverse`) | `#FFFFFF` | `#4F46E5` | **6.28:1** | **PASS** | **PASS** | FAIL | **AA+**. Идеально для кнопки «Оплатить» |
| **CTA Hover** | `#FFFFFF` | `#4338CA` | **7.90:1** | **PASS** | **PASS** | **PASS** | **AAA**. Повышение контраста при наведении |
| **CTA Pressed** | `#FFFFFF` | `#3730A3` | **9.93:1** | **PASS** | **PASS** | **PASS** | **AAA**. Четкий отклик при нажатии |
| **Текст титульного баннера** | `#FFFFFF` | `#0F172A` | **17.85:1** | **PASS** | **PASS** | **PASS** | Абсолютный контраст титула книги |
| **Текст супер-заголовков** | `#FFFFFF` | `#1E293B` | **14.63:1** | **PASS** | **PASS** | **PASS** | Премиальный вид разделов |
| **Иконка успеха** (`feedback-success`) | `#16A34A` | `#FFFFFF` | **3.30:1** | FAIL | **PASS** | FAIL | **Только для иконок!** (Pass для UI $\ge 3:1$) |
| **Иконка успеха на плашке** | `#16A34A` | `#ECFDF5` | **3.13:1** | FAIL | **PASS** | FAIL | Pass для UI-компонентов |
| **Текст успеха** (`feedback-success-text`) | **`#15803D`** | `#FFFFFF` | **5.02:1** | **PASS** | **PASS** | FAIL | **Pass AA для обычного текста!** |
| **Текст успеха на плашке** | **`#15803D`** | `#ECFDF5` | **4.76:1** | **PASS** | **PASS** | FAIL | Pass AA для бейджей статусов |
| **Текст выручки/KPI в таблицах** | **`#047857`** | `#FFFFFF` | **5.48:1** | **PASS** | **PASS** | FAIL | Изумрудный для подтвержденных оплат |
| **Текст ошибки валидации** | `#DC2626` | `#FFFFFF` | **4.83:1** | **PASS** | **PASS** | FAIL | **Pass AA при размещении на белом фоне** |
| **Ошибка на розовой плашке** | `#DC2626` | `#FFF1F2` | **4.40:1** | FAIL | **PASS** | FAIL | Доказывает правило: **фон инпута НЕ красить!** |
| **Усиленный текст ошибки** | `#B91C1C` | `#FFFFFF` | **6.47:1** | **PASS** | **PASS** | FAIL | Рекомендуется для критических алертов |
| **Иконка предупреждения** | `#D97706` | `#FFFFFF` | **3.19:1** | FAIL | **PASS** | FAIL | Только для пиктограмм $\ge 3.0:1$ |
| **Текст предупреждения** | **`#B45309`** | `#FFFFFF` | **5.02:1** | **PASS** | **PASS** | FAIL | **Pass AA для текста задержек и холдов** |
| **Текст предупреждения на плашке** | **`#B45309`** | `#FFFBEB` | **4.84:1** | **PASS** | **PASS** | FAIL | Pass AA для плашек дебиторки |
| **Ссылка навигации** | `#1E3A8A` | `#EEF2FF` | **9.26:1** | **PASS** | **PASS** | **PASS** | **AAA**. Максимальная читаемость в строке 2 |
| **Рамка активного поля** (`border-focus`) | `#4F46E5` | `#FFFFFF` | **6.29:1** | **PASS** | **PASS** | FAIL | Превосходит требование 3:1 для индикатора фокуса |

---

## 2. Оптимизация типографики: Playfair Display vs Sans-Serif

### Оптический анализ шрифта Playfair Display
Шрифт **Playfair Display** относится к классу высококонтрастных переходных антикв (Modern Serif / Didone-influence):
- **Сильные стороны**: Ярко выраженный контраст между основными и соединительными штрихами (hairlines), утонченные засечки, округлые каплевидные окончания (ball terminals). Вызывает подсознательные ассоциации с швейцарским private banking, отчетами Financial Times, консалтингом McKinsey и люксовыми B2B-продуктами с чеком от 1 000 000 ₽.
- **Ограничения и зоны риска**:
  1. *Исчезновение волосных линий*: При кегле меньше 12pt на мониторах со стандартным DPI тонкие засечки размываются («визуальное мерцание»).
  2. *Непригодность для плотных таблиц*: Засечки сцепляются визуально при чтении рядов чисел, замедляя считывание.
  3. *Пропорциональные цифры*: В Playfair Display цифры не являются моноширинными табличными (`tnum`), что приводит к визуальному «гулянию» разрядов в финансовых колонках.
  4. *Снижение кликабельности в кнопках*: В кнопках CTA антиква выглядит как журнальный заголовок, а не как интерактивный элемент. Кнопка должна ощущаться монолитной и физически нажимаемой.

### Матрица оптимизации: где оставить Playfair, а где заменить на Sans-Serif

| Элемент интерфейса | Рекомендуемый шрифт | Кегль и начертание | Обоснование решения |
| :--- | :--- | :--- | :--- |
| **Строка 1: Главный баннер книги/листа** | **Playfair Display** | **15–16pt Bold** | **Оставить**. Задает статус премиального SaaS-решения с первой секунды. |
| **Супер-заголовки разделов дашборда** | **Playfair Display** | **13–14pt Bold** | **Оставить**. Создает четкое зонирование экранов без тяжелых рамок. |
| **Стратегические вердикты СЕО / РОПа** | **Playfair Display** | **12–13pt Bold** | **Оставить**. Выделяет управленческий вывод (например, «🎯 ВЕРДИКТ: ПЕРЕХВАТ СДЕЛКИ D-104»). |
| **Крупные KPI-суммы витрин (Topline)** | **Playfair Display** | **20–24pt Bold** | **Оставить (только одиночные)**. В крупном размере (20pt+) цифры Playfair выглядят как отчеканенная валюта. |
| **Итоговая сумма заказа в Checkout** | **Playfair Display** | **28pt Bold** | **Оставить**. Сумма «450 000 ₽» должна выглядеть весомо и монументально. |
| **Колонки таблиц с цифрами и кассой** | **Inter / Segoe UI** | **11pt Medium (tnum)** | **Заменить на Sans-serif**. Моноширинные табличные цифры выравниваются строго по разрядам. |
| **Заголовки столбцов таблиц (TH)** | **Inter / Segoe UI** | **11pt Bold** | **Заменить на Sans-serif**. Обеспечивает мгновенное горизонтальное сканирование глазами. |
| **Строки описаний, причин, регламентов** | **Inter / Segoe UI** | **11pt Regular** | **Заменить на Sans-serif**. Высокая читаемость многострочного русского текста. |
| **Поля ввода формы (Input / Placeholder)** | **Inter** | **15pt Regular** | **Заменить на Sans-serif**. Исключает ошибки при вводе ИНН, БИК, Email. |
| **Кнопка «Оплатить / Сформировать счет»** | **Inter** | **16pt Bold** | **Заменить на Sans-serif**. Геометрический гротеск максимизирует конверсию клика. |
| **Бейджи статусов (`ОПЛАЧЕНО`, `ПРОСРОЧКА`)** | **Inter / Segoe UI** | **10.5pt Bold** | **Заменить на Sans-serif**. Компактность и мгновенная различимость. |
| **Строка 2: Навигационные пилюли** | **Inter / Segoe UI** | **10.5pt Bold** | **Заменить на Sans-serif**. Интерактивные табы должны читаться мгновенно. |

---

## 3. Токены дизайн-системы (Tokens Specification)

### 3.1. CSS Custom Properties (`:root`)
```css
:root {
  /* Surfaces (Stripe Principle: depth via tone, not heavy shadows) */
  --surface-base: #F8FAFC;            /* Холодный светлый холст */
  --surface-elevated: #FFFFFF;        /* Карточки счетов, форма checkout */
  --surface-sunken: #F1F5F9;          /* Заголовки таблиц, подложки */

  /* Primary Action CTA (Sacred Token: Single element on screen) */
  --action-primary: #4F46E5;          /* Индиго Stripe / Fintech */
  --action-primary-hover: #4338CA;    /* Контраст 7.9:1 (AAA) */
  --action-primary-pressed: #3730A3;  /* Контраст 9.9:1 (AAA) */
  --action-primary-soft: #EEF2FF;     /* Подложки навигации, не спорящие с CTA */
  --action-primary-success: #15803D;  /* 5.02:1 (AA) - завершенное состояние действия */

  /* Typography Colors */
  --text-primary: #0F172A;            /* 17.85:1 (AAA) - максимальный контраст */
  --text-secondary: #64748B;          /* 4.76:1 (AA) - пояснения, ссылки без конкуренции */
  --text-disabled: #94A3B8;           /* Неактивные элементы */
  --text-inverse: #FFFFFF;            /* 6.29:1 (AA+) на кнопке CTA */

  /* Feedback Semantics */
  --feedback-success-icon: #16A34A;   /* 3.3:1 - графические чекмарки */
  --feedback-success-text: #15803D;   /* 5.02:1 (AA) - текст успеха */
  --feedback-success-bg: #ECFDF5;     /* Мягкий фон бейджа успеха */
  --feedback-error: #DC2626;          /* 4.83:1 (AA) - только текст ошибки! */
  --feedback-error-bg: #FFF1F2;
  --feedback-warning-icon: #D97706;   /* Иконка внимания */
  --feedback-warning-text: #B45309;   /* 5.02:1 (AA) - текст задержек и холдов */
  --feedback-warning-bg: #FFFBEB;

  /* Borders & Focus Indicators */
  --border-default: #E2E8F0;          /* Тонкая граница карточек */
  --border-focus: #4F46E5;            /* Фокус 2px цвета CTA */
  --action-primary-border: #C7D2FE;   /* Мягкая индиго-рамка для акцентных блоков и карточек */
  --focus-ring: 0 0 0 3px rgba(79, 70, 229, 0.22);

  /* Typography Font Families */
  --font-display: 'Playfair Display', Georgia, serif;
  --font-body: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
}
```

### 3.2. SCSS Map
```scss
$revops-tokens: (
  surface: (
    base: #F8FAFC,
    elevated: #FFFFFF,
    sunken: #F1F5F9
  ),
  action: (
    primary: #4F46E5,
    hover: #4338CA,
    pressed: #3730A3,
    soft: #EEF2FF,
    success: #15803D
  ),
  text: (
    primary: #0F172A,
    secondary: #64748B,
    disabled: #94A3B8,
    inverse: #FFFFFF
  ),
  feedback: (
    success-icon: #16A34A,
    success-text: #15803D,
    success-bg: #ECFDF5,
    error: #DC2626,
    error-bg: #FFF1F2,
    warning-icon: #D97706,
    warning-text: #B45309,
    warning-bg: #FFFBEB
  ),
  border: (
    default: #E2E8F0,
    focus: #4F46E5,
    action-primary-border: #C7D2FE
  )
);
```

### 3.3. JSON Tokens (Для CI/CD и Design Systems)
```json
{
  "color": {
    "surface": {
      "base": { "value": "#F8FAFC", "type": "color" },
      "elevated": { "value": "#FFFFFF", "type": "color" },
      "sunken": { "value": "#F1F5F9", "type": "color" }
    },
    "action": {
      "primary": { "value": "#4F46E5", "type": "color" },
      "hover": { "value": "#4338CA", "type": "color" },
      "pressed": { "value": "#3730A3", "type": "color" },
      "soft": { "value": "#EEF2FF", "type": "color" },
      "success": { "value": "#15803D", "type": "color" }
    },
    "text": {
      "primary": { "value": "#0F172A", "type": "color" },
      "secondary": { "value": "#64748B", "type": "color" },
      "disabled": { "value": "#94A3B8", "type": "color" },
      "inverse": { "value": "#FFFFFF", "type": "color" }
    },
    "feedback": {
      "successText": { "value": "#15803D", "type": "color" },
      "successIcon": { "value": "#16A34A", "type": "color" },
      "error": { "value": "#DC2626", "type": "color" },
      "warningText": { "value": "#B45309", "type": "color" }
    },
    "border": {
      "default": { "value": "#E2E8F0", "type": "color" },
      "focus": { "value": "#4F46E5", "type": "color" }
    }
  }
}
```

---

### 3.4. Архитектура состояний взаимодействия (Post-Action & Interactive States)

| Состояние | Поведение UI / Классы | Токены и Стили | WCAG & UX Стандарты |
| :--- | :--- | :--- | :--- |
| **Default (Готов к клику)** | Кнопка активна, фокус по Tab | `background: var(--action-primary)`, `color: #FFFFFF`, размер: **16px (1rem) Semibold (600)**, высота 52px | WCAG AA+ (6.28:1), Tap target 52px |
| **Hover / Active** | Мягкое затемнение, `transform: scale(0.99)` | `var(--action-primary-hover)` / `var(--action-primary-pressed)` | Контраст 7.9:1 / 9.9:1 (AAA) |
| **Focus-Visible** | Внешнее кольцо фокуса | `box-shadow: 0 0 0 3px rgba(79, 70, 229, 0.35)` | WCAG 2.4.7 (Focus Visible) |
| **Loading (В процессе)** | Кнопка заблокирована, отображается спиннер | `.btn-checkout:disabled` (`opacity: 0.82; background: var(--action-primary)`). Спиннер 18px | Визуальная стабильность: кнопка не схлопывается и не теряет геометрию |
| **Motion Reduction** | Отключение вращения спиннера | `@media (prefers-reduced-motion: reduce) { .spinner { animation: none; } }` | WCAG 2.2 SC 2.3.3 (Animation from Interactions) |
| **Success (Пост-действие)** | Кнопка фиксирует успех, отображается квитанция | `.btn-checkout.is-success` (`background: var(--action-primary-success)`), текст `✓ Счет сформирован` | Контраст 5.02:1 AA. Переключение классом, а не inline-стилями |
| **Screen Reader Announcement** | Озвучивание результата | Статичный контейнер `<div id="liveAnnouncer" class="sr-only" role="status" aria-live="polite">` | Гарантия считывания скринридерами (NVDA, VoiceOver) без сбоев из-за `display:none` |
| **Дисциплина Checkout (No Exit Ramps)**| **Строго 0 альтернативных CTA** | Под кнопкой **запрещены** любые soft-CTA («Не готовы? Бесплатный тест-драйв»). Вся страховка переносится в **пост-конверсионный блок** | Защита конверсии: ноль альтернативных путей в точке максимального намерения |

---

## 4. Результаты проверки HTML-прототипа `RevOps_B2B_Enterprise_Checkout.html`

Прототип расположен по адресу:  
[`RevOps_B2B_Enterprise_Checkout.html`](file:///C:/Users/strel/.gemini/antigravity/brain/6a2f6569-0174-4758-bbb9-311ec84e9bf3/RevOps_B2B_Enterprise_Checkout.html)

### Проведенная верификация:
1. **Семантическая разметка HTML5**:
   - Форма переведена на теги `<form novalidate>`, способы оплаты оформлены в `<fieldset>` с легендой `<legend class="section-title">` и `role="radiogroup"`.
   - Сайдбар заказа вынесен в семантический `<aside class="summary-card" aria-label="Детали заказа">`.
2. **Явные связи меток и полей (WCAG 1.3.1 & 4.1.2)**:
   - Все поля ввода имеют жесткие атрибуты `for="innInput"`, `for="kppInput"`, `for="bikInput"`, `for="emailInput"`, `for="phoneInput"`.
3. **Обработка ошибок и доступность скринридеров (WCAG 3.3.1 & 3.3.2)**:
   - Сообщения об ошибках снабжены атрибутом `role="alert"`.
   - Поля ввода связываются с подсказками и ошибками через `aria-describedby="innError innHint"`.
   - При невалидном вводе динамически проставляется `aria-invalid="true"`.
4. **Правило заливки полей (Критический фикс контрастности)**:
   - При ошибке фон поля ввода **остается чистым белым (`#FFFFFF`)**, меняется только цвет границы на `#DC2626` и выводится текст ошибки `#DC2626` (контраст **4.83:1 AA**).
5. **Клавиатурная навигация и фокус (WCAG 2.4.7 & 1.4.11)**:
   - Фокус поля оформлен 2px рамкой цвета CTA (`#4F46E5`, контраст 6.29:1) и мягким кольцом свечения `0 0 0 3px rgba(79, 70, 229, 0.22)`.
   - Кнопка CTA имеет `:focus-visible` индикатор для навигации клавишей Tab.
6. **Единственный Primary CTA и дисциплина Exit Ramp (Stripe & Linear Benchmark)**:
   - На всей странице цвет `#4F46E5` используется **только для кнопки «Сформировать счет на 450 000 ₽»** (стандарт: `Inter 16px (1rem), Weight 600 Semibold, height 52px`).
   - Исключены любые альтернативные пути или ссылки под CTA («Не готовы к спринту? Бесплатный тест-драйв»). Вся страховка unready-трафика перенесена строго в пост-конверсионный блок `#checkoutSuccessBanner` (`.post-conversion-hint`), предотвращая утечку конверсии в точке максимального намерения покупки.

---

## 5. Чек-лист внедрения для команды (Spreadsheets & Frontend)

### 5.1. Чек-лист по витринным листам книги RevOps Platform V17.6

| Лист книги | Элемент листа | Координаты | Применяемый токен | Шрифт / Начертание | Высота / Размеры |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **`📄 Executive_OnePager`** | Главный баннер | Строка 1 | `navy-hero` (`#0F172A`) | Playfair Display 15pt Bold | 46–56 px |
| | Панель навигации | Строка 2 | `action-primary-soft` (`#EEF2FF`) | Inter/Segoe UI 10.5pt Bold | 32–38 px |
| | Карточки Выручки/Плана | A4:B5 | `feedback-success-bg` (`#ECFDF5`) | Playfair Display 20pt Bold | Значения: 44–54 px |
| | Карточки Прогноза | C4:D5 | `surface-sunken` (`#F1F5F9`) | Playfair Display 18pt Bold | Значения: 44–54 px |
| | Карточка Риска SLA | F4:F5 | `feedback-error-bg` (`#FFF1F2`) | Playfair Display 20pt Bold | Значения: 44–54 px |
| | Заголовки таблиц | A7, F7 | `navy-sub` (`#1E293B`) | Playfair Display 13pt Bold | 32–40 px |
| | Строки данных воронки | A9:I14 | `surface-elevated` / `surface-base` | Inter/Segoe UI 11pt Regular | 26–32 px |
| | Вердикт руководства | A16 | `feedback-warning-bg` (`#FEF3C7`) | Playfair Display 12pt Bold | 36–46 px |
| **`⚡ Пульс_Компании`** | Двухуровневые KPI | Строки 4–7 | Изумрудный, Янтарный, Розовый | Playfair Display 20pt Bold | Метки: 24px, Цифры: 44px |
| | Таблица распределения | Строки 9–16 | Зебра `#FFFFFF` / `#F8FAFC` | Inter/Segoe UI 11pt Regular | 26–32 px |
| **`📋 Пульт_РОПа_15_Минут`** | Карточка денег в риске | D4:D5 | `feedback-warning-bg` (`#FFFBEB`) | Playfair Display 20pt Bold | 44–54 px |
| | Карточка дефектов речи | B4:B5 | `feedback-error-bg` (`#FFF1F2`) | Playfair Display 20pt Bold | 44–54 px |
| | Срочные действия РОПа | A9:H13 | Зебра `#FFFFFF` / `#F8FAFC` | Inter/Segoe UI 11pt Bold | **38–48 px** (под текст) |
| **`💸 Диагностика_Утечек_ОП`** | 7 Грехов (Аудит As-Is) | A10:G15 | Зебра `#FFFFFF` / `#F8FAFC` | Inter/Segoe UI 11pt Regular | **38–48 px** (под текст) |
| | Строка Итого потерь | A16:G16 | Розовая плашка (`#FEE2E2`) | Playfair Display 13pt Bold | 34–42 px |
| | Таблица экономики спринта | A19:E24 | Белоснежная карточка | Inter/Segoe UI 11pt, факт 12pt | 28–34 px |
| | **PRIMARY CTA: Утвердить** | **A26** | **`action-primary` (`#4F46E5`)** | **Inter/Segoe UI 12pt Bold** | **42–52 px** |
| | Гарантии и 152-ФЗ | A27 | `text-secondary` (`#64748B`) | Inter/Segoe UI 10pt Regular | 24–30 px |
| **`⚡ Экспресс_Калькулятор_3_Цифры`** | Поля ввода 3 цифр | B5:B7 | Золотистый инпут (`#FEF3C7`) | Playfair Display 13pt Bold | 28–34 px |
| | Карточки 4 сценариев | A10:D10 | Розовый, Мятный, Изумрудный | Playfair Display 12pt Bold | 42–52 px |
| | Таблица 4 узких мест | A15:G18 | Белоснежная карточка | Inter/Segoe UI 11pt Regular | 36–44 px |
| | **PRIMARY CTA: Тест-драйв** | **A25** | **`action-primary` (`#4F46E5`)** | **Inter/Segoe UI 12pt Bold** | **42–52 px** |
| | Гарантия конфиденциальности | A26 | `text-secondary` (`#64748B`) | Inter/Segoe UI 10pt Regular | 24–30 px |
| **`💳 Финансы_и_AI_Дожим`** | Платежный календарь | A8:I14 | Белоснежная карточка | Inter/Segoe UI 11pt Regular | 28–32 px |
| | Статусы инвойсов | H8:H14 | Зеленый `#15803D` / Красный `#DC2626` | Inter/Segoe UI 10.5pt Bold | Центр, Middle |
| **`🔮 Симулятор_Роста`** | Коридор Монте-Карло | A15:G18 | P10 (Rose), P50 (Amber), P90 (Emerald)| Playfair 13pt / Inter 11pt | 30–36 px |
| **`🧪 QA_Suite`** | Вердикт 100% готовности | A3 | Изумрудная рамка (`#A7F3D0`) | Playfair Display 16pt Bold | 36–46 px |
| | Сводные счетчики (118/118) | B4:D4 | Белоснежные карточки | Playfair Display 20pt Bold | 42–52 px |
| | Таблица 118 регресс-тестов | A7:G124 | Зебра `#FFFFFF` / `#F8FAFC` | Inter/Segoe UI 10pt Regular | 24–28 px |

---

### 5.2. Памятка для разработчиков и верстальщиков (Engineering Rules)

1. **Никакого `#4F46E5` в навигации**: Навигационные плашки строки 2 оформляются исключительно токеном `action-primary-soft` (`#EEF2FF`) с цветом текста `#1E3A8A`. Индиго зарезервирован только под целевые кнопки действия.
2. **Белые фоны для инпутов**: При валидации формы никогда не заливать фон инпута бледно-красным цветом. Рамка становится `#DC2626`, текст ошибки `#DC2626`, фон инпута — строго `#FFFFFF`.
3. **Моноширинные цифры в таблицах**: В веб-интерфейсе для таблиц всегда задавать `font-variant-numeric: tabular-nums;`.
4. **Многострочный текст в таблицах**: В Google Sheets и Excel для столбцов комментариев, симптомов и действий обязательно включен атрибут `wrapText = True` и установлена достаточная высота строки (от 36pt).
5. **Сохранение формул QA Suite**: Любое обновление стилей Excel должно сопровождаться верификацией через `python check_target_files.py` с сохранением статуса 118 / 118 PASS (100.0%).
