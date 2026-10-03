import os
import subprocess
import shutil

html_content = """<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<title>RevOps Enterprise OS V18.0 — Презентация</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

@page {
    size: 16in 9in;
    margin: 0;
}

* {
    box-sizing: border-box;
    margin: 0;
    padding: 0;
}

body {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    background: #E2E8F0;
    color: #0F172A;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
}

.slide {
    width: 16in;
    height: 9in;
    page-break-after: always;
    position: relative;
    overflow: hidden;
    padding: 60px 80px;
    background: #F1F5F9;
    display: flex;
    flex-direction: column;
}

/* Background Variations */
.slide.dark-bg {
    background: #0B1329;
    color: #FFFFFF;
    justify-content: center;
    padding: 80px 100px;
}

/* Decorative leaf / modern background element */
.dark-bg .bg-glow {
    position: absolute;
    right: -100px;
    bottom: -100px;
    width: 700px;
    height: 700px;
    background: radial-gradient(circle, rgba(37, 99, 235, 0.25) 0%, rgba(15, 23, 42, 0) 70%);
    border-radius: 50%;
    pointer-events: none;
}

.dark-bg .bg-pattern {
    position: absolute;
    right: 50px;
    top: 50px;
    opacity: 0.08;
    pointer-events: none;
}

.header-logo {
    display: flex;
    align-items: center;
    gap: 12px;
    margin-bottom: 40px;
}

.logo-icon {
    width: 38px;
    height: 38px;
    background: #2563EB;
    border-radius: 10px;
    display: flex;
    align-items: center;
    justify-content: center;
    color: white;
    font-weight: 800;
    font-size: 20px;
}

.logo-text {
    font-size: 26px;
    font-weight: 700;
    letter-spacing: -0.5px;
    color: #FFFFFF;
}

.slide-header {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    margin-bottom: 35px;
}

.slide-title {
    font-size: 38px;
    font-weight: 800;
    color: #0F172A;
    letter-spacing: -0.5px;
    line-height: 1.2;
}

.slide-subtitle {
    font-size: 18px;
    color: #64748B;
    margin-top: 6px;
    font-weight: 500;
}

.corner-logo {
    display: flex;
    align-items: center;
    gap: 8px;
    opacity: 0.85;
}

.corner-logo-text {
    font-size: 18px;
    font-weight: 700;
    color: #2563EB;
}

/* Card Styles */
.card {
    background: #FFFFFF;
    border-radius: 20px;
    padding: 30px;
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.04);
    border: 1px solid #E2E8F0;
    display: flex;
    flex-direction: column;
}

/* Slide 1 Styles */
.hero-title {
    font-size: 64px;
    font-weight: 800;
    line-height: 1.15;
    margin-bottom: 24px;
    letter-spacing: -1px;
    max-width: 1100px;
}

.hero-subtitle {
    font-size: 24px;
    color: #94A3B8;
    line-height: 1.5;
    max-width: 950px;
    margin-bottom: 50px;
    font-weight: 400;
}

.hero-badge {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: rgba(37, 99, 235, 0.2);
    border: 1px solid rgba(59, 130, 246, 0.4);
    padding: 8px 18px;
    border-radius: 30px;
    font-size: 14px;
    font-weight: 600;
    color: #60A5FA;
    margin-bottom: 25px;
    width: fit-content;
}

/* Slide 2: Ecosystem Architecture */
.top-hub-card {
    background: #FFFFFF;
    border-radius: 18px;
    padding: 22px 30px;
    text-align: center;
    margin-bottom: 24px;
    border: 1px solid #CBD5E1;
    box-shadow: 0 4px 15px rgba(0,0,0,0.03);
}

.top-hub-title {
    font-size: 22px;
    font-weight: 800;
    color: #0F172A;
}

.top-hub-sub {
    font-size: 15px;
    color: #64748B;
    margin-top: 4px;
}

.quad-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 20px;
    flex: 1;
}

.quad-card {
    background: #FFFFFF;
    border-radius: 18px;
    padding: 26px;
    border: 1px solid #E2E8F0;
    display: flex;
    flex-direction: column;
}

.quad-tag {
    display: inline-block;
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    color: #2563EB;
    background: #EFF6FF;
    padding: 4px 10px;
    border-radius: 6px;
    margin-bottom: 14px;
    width: fit-content;
}

.quad-title {
    font-size: 19px;
    font-weight: 700;
    color: #0F172A;
    margin-bottom: 12px;
    line-height: 1.3;
}

.quad-desc {
    font-size: 13.5px;
    color: #64748B;
    line-height: 1.5;
    flex: 1;
}

.quad-metric {
    margin-top: 16px;
    padding-top: 14px;
    border-top: 1px solid #F1F5F9;
    font-size: 13px;
    font-weight: 600;
    color: #0F172A;
}

/* Slide 3: 3 Cards What we do */
.tri-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 28px;
    flex: 1;
}

.tri-card {
    background: #FFFFFF;
    border-radius: 22px;
    padding: 45px 35px;
    border: 1px solid #E2E8F0;
    display: flex;
    flex-direction: column;
}

.icon-box {
    width: 64px;
    height: 64px;
    border-radius: 16px;
    background: #EFF6FF;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 30px;
    margin-bottom: 28px;
    color: #2563EB;
}

.tri-title {
    font-size: 23px;
    font-weight: 700;
    color: #0F172A;
    margin-bottom: 16px;
    line-height: 1.3;
}

.tri-desc {
    font-size: 16px;
    color: #64748B;
    line-height: 1.6;
}

/* Slide 4: 2 Columns Audience */
.dual-col-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 30px;
    flex: 1;
}

.audience-card {
    background: #FFFFFF;
    border-radius: 22px;
    padding: 35px;
    border: 1px solid #E2E8F0;
    display: flex;
    flex-direction: column;
}

.audience-header {
    border-bottom: 2px solid #F1F5F9;
    padding-bottom: 18px;
    margin-bottom: 22px;
}

.audience-role {
    font-size: 22px;
    font-weight: 800;
    color: #0F172A;
}

.audience-sub {
    font-size: 14px;
    color: #64748B;
    margin-top: 4px;
}

.task-list {
    display: flex;
    flex-direction: column;
    gap: 14px;
    flex: 1;
}

.task-item {
    display: flex;
    align-items: flex-start;
    gap: 14px;
}

.task-num {
    width: 26px;
    height: 26px;
    border-radius: 50%;
    background: #EFF6FF;
    color: #2563EB;
    font-weight: 700;
    font-size: 13px;
    display: flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
    margin-top: 2px;
}

.task-text {
    font-size: 14.5px;
    color: #334155;
    line-height: 1.45;
}

/* Slides 5, 6, 7: How it Works (3 Steps with huge numbers) */
.steps-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 26px;
    flex: 1;
}

.step-card {
    background: #FFFFFF;
    border-radius: 22px;
    padding: 38px 32px;
    border: 1px solid #E2E8F0;
    display: flex;
    flex-direction: column;
    position: relative;
    overflow: hidden;
}

.step-watermark {
    position: absolute;
    top: -20px;
    right: 20px;
    font-size: 120px;
    font-weight: 900;
    color: #F1F5F9;
    z-index: 1;
    pointer-events: none;
    line-height: 1;
}

.step-content {
    position: relative;
    z-index: 2;
    display: flex;
    flex-direction: column;
    flex: 1;
}

.step-num-badge {
    width: 44px;
    height: 44px;
    border-radius: 12px;
    background: #2563EB;
    color: #FFFFFF;
    font-weight: 800;
    font-size: 20px;
    display: flex;
    align-items: center;
    justify-content: center;
    margin-bottom: 22px;
}

.step-title {
    font-size: 21px;
    font-weight: 700;
    color: #0F172A;
    margin-bottom: 16px;
    line-height: 1.3;
}

.step-points {
    list-style: none;
    display: flex;
    flex-direction: column;
    gap: 12px;
}

.step-point {
    font-size: 14.5px;
    color: #475569;
    line-height: 1.5;
    position: relative;
    padding-left: 20px;
}

.step-point::before {
    content: "•";
    position: absolute;
    left: 4px;
    color: #2563EB;
    font-weight: 900;
    font-size: 18px;
}

/* Slide 8: Numbers Grid */
.numbers-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    grid-template-rows: repeat(2, 1fr);
    gap: 22px;
    flex: 1;
}

.num-card {
    background: #FFFFFF;
    border-radius: 20px;
    padding: 30px 24px;
    border: 1px solid #E2E8F0;
    display: flex;
    flex-direction: column;
    justify-content: center;
}

.num-val {
    font-size: 40px;
    font-weight: 900;
    color: #2563EB;
    letter-spacing: -1px;
    line-height: 1.1;
    margin-bottom: 10px;
}

.num-label {
    font-size: 14px;
    color: #475569;
    line-height: 1.45;
}

/* Slide 9: 8 Facts with checkmarks */
.facts-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    grid-template-rows: repeat(2, 1fr);
    gap: 20px;
    flex: 1;
}

.fact-card {
    background: #FFFFFF;
    border-radius: 18px;
    padding: 24px;
    border: 1px solid #E2E8F0;
    display: flex;
    flex-direction: column;
}

.check-circle {
    width: 32px;
    height: 32px;
    border-radius: 50%;
    background: #10B981;
    color: white;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 16px;
    font-weight: 800;
    margin-bottom: 16px;
}

.fact-title {
    font-size: 16px;
    font-weight: 700;
    color: #0F172A;
    margin-bottom: 8px;
    line-height: 1.35;
}

.fact-desc {
    font-size: 13.5px;
    color: #64748B;
    line-height: 1.45;
}

/* Slide 10: 3 Principles */
.principles-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 28px;
    flex: 1;
}

.principle-card {
    background: #FFFFFF;
    border-radius: 22px;
    padding: 42px 34px;
    border: 1px solid #E2E8F0;
    display: flex;
    flex-direction: column;
}

.principle-title {
    font-size: 22px;
    font-weight: 800;
    color: #0F172A;
    margin-bottom: 20px;
    line-height: 1.3;
}

.principle-body {
    font-size: 15.5px;
    color: #475569;
    line-height: 1.6;
    margin-bottom: 20px;
}

.principle-highlight {
    margin-top: auto;
    padding: 14px 18px;
    background: #EFF6FF;
    border-radius: 12px;
    font-size: 13.5px;
    color: #1E40AF;
    font-weight: 600;
    line-height: 1.45;
}

/* Slide 11: Call to Action */
.cta-box {
    background: rgba(255, 255, 255, 0.05);
    border: 1px solid rgba(255, 255, 255, 0.12);
    border-radius: 24px;
    padding: 45px 50px;
    max-width: 1050px;
    margin-bottom: 45px;
}

.cta-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 30px;
    margin-top: 30px;
}

.cta-step {
    border-left: 2px solid #3B82F6;
    padding-left: 18px;
}

.cta-step-num {
    font-size: 13px;
    text-transform: uppercase;
    color: #60A5FA;
    font-weight: 700;
    margin-bottom: 4px;
}

.cta-step-title {
    font-size: 17px;
    font-weight: 700;
    color: #FFFFFF;
    margin-bottom: 6px;
}

.cta-step-desc {
    font-size: 14px;
    color: #94A3B8;
    line-height: 1.4;
}

.contacts-row {
    display: flex;
    gap: 40px;
    font-size: 17px;
    color: #E2E8F0;
    font-weight: 500;
}
</style>
</head>
<body>

<!-- СЛАЙД 1: ОБЛОЖКА -->
<div class="slide dark-bg">
    <div class="bg-glow"></div>
    <div class="header-logo">
        <div class="logo-icon">R</div>
        <div class="logo-text">RevOps Enterprise OS</div>
    </div>
    <div class="hero-badge">ENTERPRISE RELEASE V18.0 • REVOPS ARCHITECTURE</div>
    <h1 class="hero-title">Добро пожаловать в RevOps Enterprise OS!</h1>
    <p class="hero-subtitle">
        Интеллектуальная платформа сквозного контроля выручки, 100% речевого ИИ-аудита диалогов и оперативного управления отделом продаж для B2B-компаний.
    </p>
    <div style="font-size: 16px; color: #64748B; font-weight: 600;">
        Готовое решение для Собственников (CEO) и Руководителей отделов продаж (РОП)
    </div>
</div>

<!-- СЛАЙД 2: СТРУКТУРА ЭКОСИСТЕМЫ -->
<div class="slide">
    <div class="slide-header">
        <div>
            <div class="slide-title">Архитектура экосистемы RevOps</div>
            <div class="slide-subtitle">Единая технологическая среда для полной прозрачности коммерческого блока</div>
        </div>
        <div class="corner-logo">
            <span class="corner-logo-text">RevOps OS</span>
        </div>
    </div>

    <div class="top-hub-card">
        <div class="top-hub-title">RevOps Core Platform (Единый Источник Правды — SSOT)</div>
        <div class="top-hub-sub">Центральное вычислительное ядро и дашборд, агрегирующий сделки, звонки, финансы и KPI</div>
    </div>

    <div class="quad-grid">
        <div class="quad-card">
            <div class="quad-tag">ИИ-Модуль</div>
            <div class="quad-title">AI Speech Inspector</div>
            <div class="quad-desc">
                Собственный сервер речевой аналитики. Автоматически транскрибирует 100% звонков менеджеров и оценивает их по 13 жестким стандартам продаж.
            </div>
            <div class="quad-metric">100% звонков без человека</div>
        </div>

        <div class="quad-card">
            <div class="quad-tag">CRM-Модуль</div>
            <div class="quad-title">amoCRM / B24 Reverse ETL</div>
            <div class="quad-desc">
                Двусторонний бесшовный коннектор. За 15 секунд возвращает в карточку сделки текстовый аудит, оценку, теги и задачи Next Step.
            </div>
            <div class="quad-metric">Запись в сделку за 15 сек</div>
        </div>

        <div class="quad-card">
            <div class="quad-tag">Менеджмент</div>
            <div class="quad-title">Пульт РОПа & Воронка SLA</div>
            <div class="quad-desc">
                Рабочий стол руководителя: 15-минутная утренняя планерка, контроль зависших КП (>48ч), платежный календарь дебиторки (DSO).
            </div>
            <div class="quad-metric">Планерка за 15 минут</div>
        </div>

        <div class="quad-card">
            <div class="quad-tag">Security & Bot</div>
            <div class="quad-title">Telegram War Room Bot</div>
            <div class="quad-desc">
                Мгновенный радар критических инцидентов. Оповещает РОПа в течение 15 минут о сливе крупной сделки, формирует утренние и вечерние сводки.
            </div>
            <div class="quad-metric">Алерт при срыве сделки</div>
        </div>
    </div>
</div>

<!-- СЛАЙД 3: ЧЕМ ЗАНИМАЕТСЯ КОМПАНИЯ -->
<div class="slide">
    <div class="slide-header">
        <div>
            <div class="slide-title">Чем занимается RevOps Enterprise OS?</div>
            <div class="slide-subtitle">Мы превращаем непредсказуемый хаос в продажах в точную математическую систему</div>
        </div>
        <div class="corner-logo"><span class="corner-logo-text">RevOps OS</span></div>
    </div>

    <div class="tri-grid">
        <div class="tri-card">
            <div class="icon-box">🛡️</div>
            <div class="tri-title">Оцифровывает и ликвидирует финансовые потери</div>
            <div class="tri-desc">
                Находит 4 главные системные утечки отдела продаж (слив на звонках, зависание КП >48ч, брошенные отказники, задержка оплат) и возвращает до 28% упущенной выручки компании.
            </div>
        </div>

        <div class="tri-card">
            <div class="icon-box">🎙️</div>
            <div class="tri-title">Слушает и объективно оценивает 100% звонков</div>
            <div class="tri-desc">
                Полностью освобождает РОПа от рутинного прослушивания сотен аудиозаписей. Нейросеть беспристрастно аудирует каждый диалог и дает персональные Coaching Tips менеджерам.
            </div>
        </div>

        <div class="tri-card">
            <div class="icon-box">📊</div>
            <div class="tri-title">Дает руководству управляемость и прогноз кассы</div>
            <div class="tri-desc">
                Заменяет субъективные отчеты на единый пульт с прогнозом закрытия месяца (Run-Rate), контролем регламентов SLA и прозрачной формулой динамической мотивации (Payroll).
            </div>
        </div>
    </div>
</div>

<!-- СЛАЙД 4: КТО НАШИ КЛИЕНТЫ -->
<div class="slide">
    <div class="slide-header">
        <div>
            <div class="slide-title">Кто наши клиенты и какие задачи мы решаем?</div>
            <div class="slide-subtitle">Специализированное решение для двух ключевых лидеров коммерческого блока</div>
        </div>
        <div class="corner-logo"><span class="corner-logo-text">RevOps OS</span></div>
    </div>

    <div class="dual-col-grid">
        <!-- Собственники -->
        <div class="audience-card">
            <div class="audience-header">
                <div class="audience-role">Собственники и Генеральные директора (CEO)</div>
                <div class="audience-sub">Компании сегмента B2B, оптовой торговли, производства и услуг</div>
            </div>
            <div class="task-list">
                <div class="task-item">
                    <div class="task-num">1</div>
                    <div class="task-text">Прекращаем слив рекламного бюджета на неквалифицированных звонках менеджеров.</div>
                </div>
                <div class="task-item">
                    <div class="task-num">2</div>
                    <div class="task-text">Обеспечиваем математически точный прогноз кассы к концу месяца (Run-Rate).</div>
                </div>
                <div class="task-item">
                    <div class="task-num">3</div>
                    <div class="task-text">Оцифровываем скрытые финансовые потери коммерческого блока в рублях.</div>
                </div>
                <div class="task-item">
                    <div class="task-num">4</div>
                    <div class="task-text">Исключаем зависимость бизнеса от «незаменимых» и токсичных менеджеров-звезд.</div>
                </div>
                <div class="task-item">
                    <div class="task-num">5</div>
                    <div class="task-text">Внедряем честный расчет зарплат: выплата бонусов только за соблюдение регламентов.</div>
                </div>
                <div class="task-item">
                    <div class="task-num">6</div>
                    <div class="task-text">Защищаем компанию от кассовых разрывов через контроль дебиторской задолженности (DSO).</div>
                </div>
            </div>
        </div>

        <!-- РОПы -->
        <div class="audience-card">
            <div class="audience-header">
                <div class="audience-role">Руководители отделов продаж (РОП / Коммерческий директор)</div>
                <div class="audience-sub">Управление командами от 3 до 50+ продавцов в CRM</div>
            </div>
            <div class="task-list">
                <div class="task-item">
                    <div class="task-num">1</div>
                    <div class="task-text">Сокращаем утреннюю планерку с 1.5 часов до 15 минут строго по фактам и цифрам.</div>
                </div>
                <div class="task-item">
                    <div class="task-num">2</div>
                    <div class="task-text">Экономим до 3 часов в день за счет автоматического 100% аудита звонков нейросетью.</div>
                </div>
                <div class="task-item">
                    <div class="task-num">3</div>
                    <div class="task-text">Мгновенно узнаем о срыве крупных сделок и успеваем перехватить клиента за 15 минут.</div>
                </div>
                <div class="task-item">
                    <div class="task-num">4</div>
                    <div class="task-text">Гарантируем 100% контроль железобетонного Next Step (дата и время контакта).</div>
                </div>
                <div class="task-item">
                    <div class="task-num">5</div>
                    <div class="task-text">Получаем готовые Coaching Tips — персональные фразы-подсказки для каждого менеджера.</div>
                </div>
                <div class="task-item">
                    <div class="task-num">6</div>
                    <div class="task-text">Устраняем рутину ручного контроля заполнения карточек и задач в amoCRM.</div>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- СЛАЙД 5: КАК ЭТО РАБОТАЕТ (МОДУЛЬ 1 - АУДИТ ЗВОНКОВ) -->
<div class="slide">
    <div class="slide-header">
        <div>
            <div class="slide-title">Как это работает: Речевой ИИ-Аудит Звонков</div>
            <div class="slide-subtitle">Технология объективного контроля 100% телефонных переговоров без участия человека</div>
        </div>
        <div class="corner-logo"><span class="corner-logo-text">RevOps OS</span></div>
    </div>

    <div class="steps-grid">
        <div class="step-card">
            <div class="step-watermark">1</div>
            <div class="step-content">
                <div class="step-num-badge">1</div>
                <div class="step-title">Бесшовный захват аудио из CRM</div>
                <ul class="step-points">
                    <li class="step-point">Интеграция перехватывает аудиозапись из любой телефонии (Манго, Телфин, Sipuni, Мои Звонки).</li>
                    <li class="step-point">Захватываются звонки 100% менеджеров по всей компании.</li>
                    <li class="step-point">0 ручных действий: менеджер просто говорит по телефону.</li>
                    <li class="step-point">Шифрованная передача аудиофайла в защищенный контур.</li>
                </ul>
            </div>
        </div>

        <div class="step-card">
            <div class="step-watermark">2</div>
            <div class="step-content">
                <div class="step-num-badge">2</div>
                <div class="step-title">Скоринг по 13 критериям продаж</div>
                <ul class="step-points">
                    <li class="step-point">ИИ транскрибирует речь и сопоставляет диалог с жестким регламентом компании.</li>
                    <li class="step-point">Бинарная проверка (0/1): приветствие, ЛПР, боли, бюджет, дедлайны КП.</li>
                    <li class="step-point">Железобетонный контроль: зафиксирована ли точная дата и время следующего шага.</li>
                    <li class="step-point">Формирование итогового балла качества от 0 до 100.</li>
                </ul>
            </div>
        </div>

        <div class="step-card">
            <div class="step-watermark">3</div>
            <div class="step-content">
                <div class="step-num-badge">3</div>
                <div class="step-title">Формирование Coaching Tip</div>
                <ul class="step-points">
                    <li class="step-point">Определение зоны риска: красный (&lt;70), желтый (70-84) или зеленый (85+).</li>
                    <li class="step-point">Генерация конкретной рекомендации: какую фразу нужно было сказать вместо ошибки.</li>
                    <li class="step-point">Анализ эмоционального тона диалога и детекция раздражения клиента.</li>
                    <li class="step-point">Фиксация данных в единой аналитической базе.</li>
                </ul>
            </div>
        </div>
    </div>
</div>

<!-- СЛАЙД 6: КАК ЭТО РАБОТАЕТ (МОДУЛЬ 2 - AMOCRM) -->
<div class="slide">
    <div class="slide-header">
        <div>
            <div class="slide-title">Как это работает: Двусторонняя Автоматизация amoCRM</div>
            <div class="slide-subtitle">Робот возвращает аналитику, задачи и теги обратно в CRM за 15 секунд</div>
        </div>
        <div class="corner-logo"><span class="corner-logo-text">RevOps OS</span></div>
    </div>

    <div class="steps-grid">
        <div class="step-card">
            <div class="step-watermark">1</div>
            <div class="step-content">
                <div class="step-num-badge">1</div>
                <div class="step-title">Развернутое саммари в сделку</div>
                <ul class="step-points">
                    <li class="step-point">Через 15 секунд после звонка в ленте карточки появляется структурированное примечание.</li>
                    <li class="step-point">Краткий пересказ сути договоренностей (чтение за 20 секунд).</li>
                    <li class="step-point">Баллы по 13 критериям и совет РОПа.</li>
                    <li class="step-point">Заполнение числового поля «Оценка ИИ (0–100)» для фильтрации воронки.</li>
                </ul>
            </div>
        </div>

        <div class="step-card">
            <div class="step-watermark">2</div>
            <div class="step-content">
                <div class="step-num-badge">2</div>
                <div class="step-title">Автопостановка задач Next Step</div>
                <ul class="step-points">
                    <li class="step-point">Если менеджер забыл назначить дату перезвона, ИИ сам ставит жесткую задачу до конца дня.</li>
                    <li class="step-point">Если клиент попросил КП к четвергу в 14:00, робот создаст задачу с точным дедлайном.</li>
                    <li class="step-point">100% защита от «забытых» и потерянных клиентов.</li>
                    <li class="step-point">Авто-заполнение полей бюджета и потребностей из разговора.</li>
                </ul>
            </div>
        </div>

        <div class="step-card">
            <div class="step-watermark">3</div>
            <div class="step-content">
                <div class="step-num-badge">3</div>
                <div class="step-title">Умное тегирование и Follow-Up</div>
                <ul class="step-points">
                    <li class="step-point">Авто-теги: #Брак_Речи, #Слив_Клиента, #Эталонный_Звонок, #Возражение_Дорого.</li>
                    <li class="step-point">Генерация готового текста сообщения в WhatsApp/Telegram для отправки клиенту.</li>
                    <li class="step-point">Менеджеру остается нажать 1 кнопку для отправки резюме встречи.</li>
                    <li class="step-point">Мгновенная сегментация воронки по реальным возражениям.</li>
                </ul>
            </div>
        </div>
    </div>
</div>

<!-- СЛАЙД 7: КАК ЭТО РАБОТАЕТ (МОДУЛЬ 3 - УПРАВЛЕНИЕ И БЕЗОПАСНОСТЬ) -->
<div class="slide">
    <div class="slide-header">
        <div>
            <div class="slide-title">Как это работает: Операционный Пульт и Безопасность</div>
            <div class="slide-subtitle">Инструменты оперативного контроля выручки и соответствие стандартам безопасности РФ</div>
        </div>
        <div class="corner-logo"><span class="corner-logo-text">RevOps OS</span></div>
    </div>

    <div class="steps-grid">
        <div class="step-card">
            <div class="step-watermark">1</div>
            <div class="step-content">
                <div class="step-num-badge">1</div>
                <div class="step-title">Утренний Пульт РОПа (15 минут)</div>
                <ul class="step-points">
                    <li class="step-point">Оценка темпа месяца (Pacing): идем ли по графику плана выручки.</li>
                    <li class="step-point">Светофор воронки: автоматическое выявление зависших сделок на этапе КП (&gt;48ч).</li>
                    <li class="step-point">ТОП-5 критических рисков выручки: список сделок, требующих немедленного дожима.</li>
                    <li class="step-point">Ежедневная постановка фокусов команде за 15 минут.</li>
                </ul>
            </div>
        </div>

        <div class="step-card">
            <div class="step-watermark">2</div>
            <div class="step-content">
                <div class="step-num-badge">2</div>
                <div class="step-title">Telegram War Room & Алерты</div>
                <ul class="step-points">
                    <li class="step-point">Красная кнопка РОПу: при срыве сделки с чеком &gt;300 000 ₽ — сигнал в Telegram за 15 минут.</li>
                    <li class="step-point">РОП успевает лично перезвонить клиенту и спасти контракт.</li>
                    <li class="step-point">Утренняя сводка в 09:00: рейтинг менеджеров и средний балл речи отдела.</li>
                    <li class="step-point">Вечерний отчет собственнику в 19:00: касса за день и прогноз месяца.</li>
                </ul>
            </div>
        </div>

        <div class="step-card">
            <div class="step-watermark">3</div>
            <div class="step-content">
                <div class="step-num-badge">3</div>
                <div class="step-title">Безопасность и 152-ФЗ РФ</div>
                <ul class="step-points">
                    <li class="step-point">Полное соблюдение Федерального закона 152-ФЗ: сервера расположены на территории РФ.</li>
                    <li class="step-point">Изоляция данных Multi-Tenant: каждый клиент работает в закрытом контуре Tenant UUID.</li>
                    <li class="step-point">Защита IP и коммерческой тайны: закрытое ядро формул и промптов.</li>
                    <li class="step-point">Шифрование трафика и разграничение прав доступа.</li>
                </ul>
            </div>
        </div>
    </div>
</div>

<!-- СЛАЙД 8: КОМПАНИЯ В ЦИФРАХ -->
<div class="slide">
    <div class="slide-header">
        <div>
            <div class="slide-title">Платформа RevOps OS в цифрах</div>
            <div class="slide-subtitle">Ключевые измеримые метрики надежности и финансовой эффективности системы</div>
        </div>
        <div class="corner-logo"><span class="corner-logo-text">RevOps OS</span></div>
    </div>

    <div class="numbers-grid">
        <div class="num-card">
            <div class="num-val">100%</div>
            <div class="num-label">звонков аудируются искусственным интеллектом без участия человека</div>
        </div>

        <div class="num-card">
            <div class="num-val">15 минут</div>
            <div class="num-label">длительность утренней планерки РОПа по фактам и цифрам вместо 1.5 часов</div>
        </div>

        <div class="num-card">
            <div class="num-val">+18–34%</div>
            <div class="num-label">средний прирост чистой выручки клиентов в первые 60 дней внедрения</div>
        </div>

        <div class="num-card">
            <div class="num-val">&lt; 48 часов</div>
            <div class="num-label">жесткий норматив нахождения сделки на этапе КП под контролем SLA</div>
        </div>

        <div class="num-card">
            <div class="num-val">13 стандартов</div>
            <div class="num-label">объективной оценки диалогов в эталонной матрице речевого аудита</div>
        </div>

        <div class="num-card">
            <div class="num-val">118 тестов</div>
            <div class="num-label">автоматической проверки целостности математики в модуле QA Suite</div>
        </div>

        <div class="num-card">
            <div class="num-val">4 утечки</div>
            <div class="num-label">финансовых потерь ликвидируются (Next Step, КП, отказники, дебиторка)</div>
        </div>

        <div class="num-card">
            <div class="num-val">15 минут</div>
            <div class="num-label">техническое время полного развертывания платформы для нового клиента</div>
        </div>
    </div>
</div>

<!-- СЛАЙД 9: ФАКТЫ О REVOPS OS -->
<div class="slide">
    <div class="slide-header">
        <div>
            <div class="slide-title">Факты о RevOps Enterprise OS</div>
            <div class="slide-subtitle">Технологические и юридические стандарты промышленного решения</div>
        </div>
        <div class="corner-logo"><span class="corner-logo-text">RevOps OS</span></div>
    </div>

    <div class="facts-grid">
        <div class="fact-card">
            <div class="check-circle">✓</div>
            <div class="fact-title">Единый контур SSOT</div>
            <div class="fact-desc">Единственная система, объединяющая речевой аудит, CRM-автоматизацию и пульт РОПа в один экран.</div>
        </div>

        <div class="fact-card">
            <div class="check-circle">✓</div>
            <div class="fact-title">Двусторонний обмен CRM</div>
            <div class="fact-desc">Прямая интеграция с amoCRM и Битрикс24 без установки сторонних расширений в браузеры сотрудников.</div>
        </div>

        <div class="fact-card">
            <div class="check-circle">✓</div>
            <div class="fact-title">100% 152-ФЗ РФ Compliant</div>
            <div class="fact-desc">Серверная инфраструктура и базы данных размещены в аккредитованных дата-центрах на территории РФ.</div>
        </div>

        <div class="fact-card">
            <div class="check-circle">✓</div>
            <div class="fact-title">Закрытый AI Black Box</div>
            <div class="fact-desc">Промпты и скоринговые модели защищены на выделенном сервере и постоянно дообучаются.</div>
        </div>

        <div class="fact-card">
            <div class="check-circle">✓</div>
            <div class="fact-title">Быстрый старт за 1 день</div>
            <div class="fact-desc">Запуск пилота без остановки работы отдела продаж. Первые данные доступны уже на следующий день.</div>
        </div>

        <div class="fact-card">
            <div class="check-circle">✓</div>
            <div class="fact-title">Адаптация под B2B-нишу</div>
            <div class="fact-desc">Калибровка 13 стандартов речи под специфику опта, сложных услуг, дистрибуции или производства.</div>
        </div>

        <div class="fact-card">
            <div class="check-circle">✓</div>
            <div class="fact-title">Динамический Payroll</div>
            <div class="fact-desc">Автоматический расчет премий с дисконтом 30% за систематический брак переговоров.</div>
        </div>

        <div class="fact-card">
            <div class="check-circle">✓</div>
            <div class="fact-title">Полная окупаемость пилота</div>
            <div class="fact-desc">Возврат инвестиций в пилотный запуск в 3–5 раз за счет спасения всего 1–2 зависших сделок.</div>
        </div>
    </div>
</div>

<!-- СЛАЙД 10: ПОЧЕМУ МЫ ДОСТИГЛИ РЕЗУЛЬТАТОВ -->
<div class="slide">
    <div class="slide-header">
        <div>
            <div class="slide-title">Почему мы достигаем таких результатов?</div>
            <div class="slide-subtitle">Три фундаментальных принципа архитектуры RevOps Enterprise OS</div>
        </div>
        <div class="corner-logo"><span class="corner-logo-text">RevOps OS</span></div>
    </div>

    <div class="principles-grid">
        <div class="principle-card">
            <div class="principle-title">1. Цифры и факты вместо субъективных отчетов</div>
            <div class="principle-body">
                Менеджеры больше не могут прикрывать неудачи фразами «клиент думает» или «он не берет трубку». Руководство видит объективную стенограмму разговора, соблюдение этапов и точные таймкоды ключевых договоренностей.
            </div>
            <div class="principle-highlight">
                Управленческие решения принимаются на базе реальных данных, а не догадок.
            </div>
        </div>

        <div class="principle-card">
            <div class="principle-title">2. Фокус на живой выручке, а не звонках ради галочки</div>
            <div class="principle-body">
                Мы оцениваем не бессмысленную длительность звонков, а продвижение сделки к деньгам: квалификацию ЛПР, выявление бюджета, сокращение времени согласования КП и соблюдение сроков оплаты счетов (DSO).
            </div>
            <div class="principle-highlight">
                Каждый диалог менеджера измеряется его влиянием на итоговую кассу месяца.
            </div>
        </div>

        <div class="principle-card">
            <div class="principle-title">3. Обучение команды без стресса и микроменеджмента</div>
            <div class="principle-body">
                Искусственный интеллект подсказывает ошибки тактично и сразу после звонка через персональные Coaching Tips. Менеджеры растут в квалификации самостоятельно, а РОП становится наставником, а не контролером.
            </div>
            <div class="principle-highlight">
                Здоровая атмосфера в отделе продаж и рост конверсий без текучки кадров.
            </div>
        </div>
    </div>
</div>

<!-- СЛАЙД 11: ФИНАЛЬНЫЙ СЛАЙД / CTA -->
<div class="slide dark-bg">
    <div class="bg-glow"></div>
    <div class="header-logo">
        <div class="logo-icon">R</div>
        <div class="logo-text">RevOps Enterprise OS</div>
    </div>
    
    <h1 class="hero-title" style="font-size: 52px; margin-bottom: 20px;">
        Готовы оцифровать и усилить ваш отдел продаж?
    </h1>
    
    <p class="hero-subtitle" style="font-size: 20px; margin-bottom: 35px;">
        Запустите 7-дневный тест-драйв: мы подключим вашу CRM и покажем полный ИИ-аудит первых 20 звонков вашей команды уже завтра!
    </p>

    <div class="cta-box">
        <div style="font-size: 19px; font-weight: 700; color: #FFFFFF; margin-bottom: 8px;">
            Простой порядок запуска пилотного проекта:
        </div>
        <div class="cta-grid">
            <div class="cta-step">
                <div class="cta-step-num">Шаг 1 • 15 минут</div>
                <div class="cta-step-title">Подключение CRM</div>
                <div class="cta-step-desc">Предоставление доступа к amoCRM и запуск штатного коннектора.</div>
            </div>
            <div class="cta-step">
                <div class="cta-step-num">Шаг 2 • 24 часа</div>
                <div class="cta-step-title">Аудит звонков</div>
                <div class="cta-step-desc">Нейросеть слушает диалоги и выявляет скрытые финансовые утечки.</div>
            </div>
            <div class="cta-step">
                <div class="cta-step-num">Шаг 3 • Результат</div>
                <div class="cta-step-title">Разбор на Пульте РОПа</div>
                <div class="cta-step-desc">Презентация карты потерь и запуск ежедневных планерок.</div>
            </div>
        </div>
    </div>

    <div class="contacts-row">
        <div>🌐 <strong>Платформа:</strong> revops-enterprise.ru</div>
        <div>📱 <strong>Telegram для связи:</strong> @revops_founder</div>
        <div>⏱️ <strong>Время развертывания:</strong> 15 минут</div>
    </div>
</div>

</body>
</html>
"""

# Write HTML
html_path = os.path.abspath("docs/Презентация_RevOps_Enterprise_OS_11_Слайдов.html")
with open(html_path, "w", encoding="utf-8") as f:
    f.write(html_content)

print(f"HTML presentation written to: {html_path}")

# Render to PDF via Edge
edge_bin = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
pdf_path = os.path.abspath("docs/Презентация_RevOps_Enterprise_OS_11_Слайдов.pdf")
pres_pdf = os.path.abspath("presentation/Презентация_RevOps_Enterprise_OS_11_Слайдов.pdf")

cmd = [
    edge_bin,
    "--headless",
    "--disable-gpu",
    "--no-pdf-header-footer",
    f"--print-to-pdf={pdf_path}",
    html_path
]

res = subprocess.run(cmd, capture_output=True, text=True)
print("Edge return code:", res.returncode)

if os.path.exists(pdf_path):
    shutil.copyfile(pdf_path, pres_pdf)
    print(f"SUCCESS: Generated 16:9 PDF Presentation: {pdf_path} ({os.path.getsize(pdf_path):,} bytes)")
    print(f"SUCCESS: Copied to presentation folder: {pres_pdf}")
else:
    print("FAILED to generate PDF via Edge")
