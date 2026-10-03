# -*- coding: utf-8 -*-
"""
Generate privacy.html (152-ФЗ Privacy Policy) and synchronize to root and docs/
"""

import os
import shutil

html_content = """<!DOCTYPE html>
<html lang="ru" class="scroll-smooth">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Политика конфиденциальности и обработки персональных данных (152-ФЗ) // RevOps OS</title>
    <meta name="description" content="Политика обработки персональных данных и защиты конфиденциальной информации сервиса RevOps OS (ai-rop.ru) в строгом соответствии с 152-ФЗ РФ.">
    <link rel="canonical" href="https://ai-rop.ru/privacy.html">

    <!-- Fonts & Icons -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600;700&family=Plus+Jakarta+Sans:wght@600;700;800&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
    
    <!-- Static Production CSS -->
    <link rel="stylesheet" href="css/tailwind.min.css">

    <!-- Yandex.Metrika -->
    <script type="text/javascript">
       (function(m,e,t,r,i,k,a){m[i]=m[i]||function(){(m[i].a=m[i].a||[]).push(arguments)};
       m[i].l=1*new Date();
       for (var j = 0; j < document.scripts.length; j++) {if (document.scripts[j].src === r) { return; }}
       k=e.createElement(t),a=e.getElementsByTagName(t)[0],k.async=1,k.src=r,a.parentNode.insertBefore(k,a)})
       (window, document, "script", "https://mc.yandex.ru/metrika/tag.js", "ym");
       ym(100714777, "init", {
            clickmap:true,
            trackLinks:true,
            accurateTrackBounce:true,
            webvisor:true
       });
    </script>
    <noscript><div><img src="https://mc.yandex.ru/watch/100714777" style="position:absolute; left:-9999px;" alt="" /></div></noscript>

    <style>
        body { font-family: 'Inter', sans-serif; }
        .font-display { font-family: 'Plus Jakarta Sans', sans-serif; }
        .font-mono { font-family: 'JetBrains Mono', monospace; }
        .glass-card {
            background: rgba(15, 23, 42, 0.7);
            backdrop-filter: blur(16px);
            border: 1px solid rgba(255, 255, 255, 0.08);
        }
    </style>
</head>
<body class="bg-slate-950 text-slate-100 min-h-screen flex flex-col antialiased selection:bg-emerald-500 selection:text-slate-950">

    <!-- HEADER / NAVIGATION -->
    <header class="sticky top-0 z-40 bg-slate-950/85 backdrop-blur-md border-b border-slate-800">
        <div class="max-w-5xl mx-auto px-4 sm:px-6 h-16 flex items-center justify-between">
            <a href="/" class="flex items-center gap-3 group">
                <div class="w-9 h-9 rounded-xl bg-gradient-to-tr from-emerald-500 to-cyan-500 flex items-center justify-center text-slate-950 font-black text-lg shadow-md shadow-emerald-500/20 group-hover:scale-105 transition-transform">
                    R
                </div>
                <div class="flex flex-col">
                    <span class="font-display font-bold text-white text-base tracking-tight leading-tight">RevOps OS</span>
                    <span class="text-[10px] font-mono text-emerald-400 font-semibold tracking-wider">AI SUPERVISOR</span>
                </div>
            </a>
            <div class="flex items-center gap-4 text-xs font-medium">
                <a href="/" class="text-slate-300 hover:text-white transition flex items-center gap-1.5">
                    <i class="fas fa-arrow-left text-[11px] text-emerald-400"></i>
                    <span>На главную</span>
                </a>
                <a href="/offer.html" class="text-slate-300 hover:text-white transition hidden sm:inline-block">
                    Договор-оферта
                </a>
                <a href="https://t.me/dm1918" target="_blank" rel="noopener noreferrer" class="px-3 py-1.5 rounded-lg bg-emerald-500/15 border border-emerald-500/30 text-emerald-300 hover:bg-emerald-500/25 transition">
                    Связь с основателем
                </a>
            </div>
        </div>
    </header>

    <!-- MAIN CONTENT -->
    <main class="flex-1 max-w-4xl mx-auto px-4 sm:px-6 py-12 lg:py-16 w-full">
        <!-- Badge & Title -->
        <div class="mb-10 text-center sm:text-left">
            <div class="inline-flex items-center gap-2 px-3.5 py-1.5 rounded-full bg-emerald-500/10 text-emerald-400 text-xs font-mono font-semibold border border-emerald-500/20 mb-4">
                <i class="fas fa-shield-halved"></i>
                <span>СООТВЕТСТВИЕ 152-ФЗ РФ И СТАНДАРТАМ NDA</span>
            </div>
            <h1 class="font-display text-3xl sm:text-4xl font-extrabold text-white tracking-tight leading-tight mb-4">
                Политика обработки персональных данных и конфиденциальности
            </h1>
            <p class="text-slate-400 text-sm sm:text-base leading-relaxed">
                Редакция действует с 02 октября 2026 г. Сервис RevOps OS (домен <strong>ai-rop.ru</strong>) гарантирует полную конфиденциальность и юридическую безопасность при обработке данных клиентов и партнеров.
            </p>
        </div>

        <div class="space-y-8 text-sm sm:text-base text-slate-300 leading-relaxed">

            <!-- Card 1 -->
            <section class="glass-card rounded-2xl p-6 sm:p-8 border border-slate-800">
                <h2 class="font-display text-lg sm:text-xl font-bold text-white mb-3 flex items-center gap-2.5">
                    <span class="w-7 h-7 rounded-lg bg-emerald-500/10 text-emerald-400 text-xs flex items-center justify-center font-mono">1</span>
                    Общие положения и Оператор данных
                </h2>
                <p class="mb-3">
                    1.1. Настоящая Политика обработки персональных данных (далее — «Политика») составлена в соответствии с Федеральным законом от 27.07.2006 № 152-ФЗ «О персональных данных» и определяет порядок сбора, хранения, обработки и защиты персональных данных пользователей сайта <strong>ai-rop.ru</strong> (далее — «Сайт»).
                </p>
                <p class="mb-3">
                    1.2. Оператором персональных данных является: <strong>Федотов Дмитрий</strong> (плательщик налога на профессиональный доход по 422-ФЗ / Индивидуальный предприниматель), ИНН: <strong>731303201073</strong>, контактный e-mail: <a href="mailto:info@ai-rop.ru" class="text-emerald-400 underline">info@ai-rop.ru</a>, официальный Telegram: <a href="https://t.me/dm1918" class="text-emerald-400 underline">@dm1918</a>.
                </p>
                <p>
                    1.3. Использование функционала Сайта, заполнение веб-форм, отправка файлов в чат-бот или заказ услуг означает безоговорочное согласие Пользователя с условиями настоящей Политики.
                </p>
            </section>

            <!-- Card 2 -->
            <section class="glass-card rounded-2xl p-6 sm:p-8 border border-slate-800">
                <h2 class="font-display text-lg sm:text-xl font-bold text-white mb-3 flex items-center gap-2.5">
                    <span class="w-7 h-7 rounded-lg bg-emerald-500/10 text-emerald-400 text-xs flex items-center justify-center font-mono">2</span>
                    Категории обрабатываемых данных
                </h2>
                <p class="mb-3">
                    2.1. В рамках оказания сервиса и демонстрации возможностей ИИ-аналитики Оператор может обрабатывать следующие категории данных:
                </p>
                <ul class="list-disc list-inside space-y-2 text-slate-300 pl-2 mb-3">
                    <li><strong>Контактные данные:</strong> фамилия, имя, номер мобильного телефона, никнейм/ID в мессенджерах Telegram и MAX, адрес электронной почты;</li>
                    <li><strong>Корпоративные данные:</strong> наименование организации, должность представителя, используемая CRM-система (amoCRM, 1С-Битрикс24 и др.) и используемая IP-телефония;</li>
                    <li><strong>Аудиозаписи звонков:</strong> записи телефонных переговоров, добровольно предоставляемые Пользователем для проведения бесплатного экспресс-аудита или в рамках Пилотного проекта;</li>
                    <li><strong>Технические данные:</strong> IP-адрес, данные файлов Cookie, географическое положение (страна/город), технические параметры браузера и устройства, реферер, данные сервисов веб-аналитики (Яндекс.Метрика, Google Analytics 4, Вебвизор).</li>
                </ul>
            </section>

            <!-- Card 3 -->
            <section class="glass-card rounded-2xl p-6 sm:p-8 border border-slate-800">
                <h2 class="font-display text-lg sm:text-xl font-bold text-white mb-3 flex items-center gap-2.5">
                    <span class="w-7 h-7 rounded-lg bg-emerald-500/10 text-emerald-400 text-xs flex items-center justify-center font-mono">3</span>
                    Цели обработки персональных данных
                </h2>
                <p class="mb-3">
                    3.1. Обработка персональных данных осуществляется исключительно в следующих законных целях:
                </p>
                <ul class="list-disc list-inside space-y-2 text-slate-300 pl-2">
                    <li>Предоставление Пользователю бесплатного экспресс-аудита 3 звонков и расчет карты потерь выручки;</li>
                    <li>Заключение, исполнение и сопровождение Договора-оферты на пилотное тестирование или подписку RevOps OS;</li>
                    <li>Организация обратной связи, проведение консультаций и презентаций сервиса;</li>
                    <li>Анализ посещаемости Сайта и оптимизация его быстродействия и интерфейса.</li>
                </ul>
            </section>

            <!-- Card 4 -->
            <section class="glass-card rounded-2xl p-6 sm:p-8 border border-slate-800">
                <h2 class="font-display text-lg sm:text-xl font-bold text-white mb-3 flex items-center gap-2.5">
                    <span class="w-7 h-7 rounded-lg bg-emerald-500/10 text-emerald-400 text-xs flex items-center justify-center font-mono">4</span>
                    Безопасность, локализация в РФ и деперсонализация аудио (PII-Sanitizer)
                </h2>
                <div class="p-4 rounded-xl bg-emerald-500/10 border border-emerald-500/30 mb-4 text-emerald-300 text-xs sm:text-sm">
                    <i class="fas fa-lock mr-1.5"></i>
                    <strong>Требование ч. 5 ст. 18 152-ФЗ РФ:</strong> Все серверные мощности и базы данных RevOps OS размещены исключительно на территории Российской Федерации (дата-центры Selectel и Yandex Cloud, Москва/Санкт-Петербург).
                </div>
                <p class="mb-3">
                    4.1. При анализе звонков модулем нейросети применяется встроенный регламент <strong>PII-Sanitizer</strong>: персональные данные конечных клиентов (паспортные данные, реквизиты карт, персональные адреса) автоматически маскируются и удаляются из текстовых транскриптов до формирования отчетов.
                </p>
                <p class="mb-3">
                    4.2. Передача данных осуществляется по защищенным криптографическим протоколам с шифрованием TLS 1.3. Доступ к рабочим интерфейсам защищен двухфакторной аутентификацией и ролевой моделью разграничения прав (RBAC).
                </p>
                <p>
                    4.3. Данные третьим лицам не передаются и не продаются, за исключением случаев, прямо предусмотренных действующим законодательством РФ.
                </p>
            </section>

            <!-- Card 5 -->
            <section class="glass-card rounded-2xl p-6 sm:p-8 border border-slate-800">
                <h2 class="font-display text-lg sm:text-xl font-bold text-white mb-3 flex items-center gap-2.5">
                    <span class="w-7 h-7 rounded-lg bg-emerald-500/10 text-emerald-400 text-xs flex items-center justify-center font-mono">5</span>
                    Использование файлов Cookie и веб-аналитики
                </h2>
                <p class="mb-3">
                    5.1. Сайт использует файлы Cookie (куки) и счетчики веб-аналитики (Яндекс.Метрика с Вебвизором и Google Analytics 4) для обеспечения базовой функциональности, запоминания предпочтений пользователя (включая тему оформления) и учета конверсий рекламы.
                </p>
                <p class="mb-3">
                    5.2. Пользователь может в любой момент ограничить или полностью отключить сохранение файлов Cookie в настройках своего браузера. При этом некоторые функции Сайта могут работать с ограничениями.
                </p>
                <p>
                    5.3. При первом посещении Сайта Пользователю демонстрируется уведомление (Cookie-баннер), продолжая использование Сайта или нажимая «Принять», Пользователь дает информированное согласие на обработку cookie-данных.
                </p>
            </section>

            <!-- Card 6 -->
            <section class="glass-card rounded-2xl p-6 sm:p-8 border border-slate-800">
                <h2 class="font-display text-lg sm:text-xl font-bold text-white mb-3 flex items-center gap-2.5">
                    <span class="w-7 h-7 rounded-lg bg-emerald-500/10 text-emerald-400 text-xs flex items-center justify-center font-mono">6</span>
                    Сроки обработки и порядок отзыва согласия
                </h2>
                <p class="mb-3">
                    6.1. Персональные данные обрабатываются до достижения целей обработки либо до момента отзыва согласия Пользователем.
                </p>
                <p class="mb-3">
                    6.2. Пользователь вправе в любой момент отозвать свое согласие на обработку персональных данных, направив письменное уведомление на электронный адрес: <a href="mailto:info@ai-rop.ru" class="text-emerald-400 underline font-medium">info@ai-rop.ru</a> или через мессенджер Telegram: <a href="https://t.me/dm1918" class="text-emerald-400 underline font-medium">@dm1918</a> с темой «Отзыв согласия на обработку персональных данных».
                </p>
                <p>
                    6.3. Оператор прекращает обработку и уничтожает персональные данные в срок, не превышающий 30 календарных дней с момента получения отзыва.
                </p>
            </section>

        </div>

        <div class="mt-12 pt-8 border-t border-slate-800 flex flex-col sm:flex-row items-center justify-between gap-4 text-xs text-slate-500">
            <div>© 2026 RevOps OS (ai-rop.ru). Все права защищены.</div>
            <div class="flex items-center gap-4">
                <a href="/offer.html" class="hover:text-slate-300 underline">Публичная оферта</a>
                <a href="/pilot-roadmap.html" class="hover:text-slate-300 underline">Дорожная карта пилота</a>
                <a href="/" class="hover:text-slate-300 underline">Главная</a>
            </div>
        </div>
    </main>

</body>
</html>
"""

repo = r"C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os"
root_path = os.path.join(repo, "privacy.html")
docs_path = os.path.join(repo, "docs", "privacy.html")

with open(root_path, "w", encoding="utf-8") as f:
    f.write(html_content)

with open(docs_path, "w", encoding="utf-8") as f:
    f.write(html_content)

print(f"✓ Generated privacy.html in root ({os.path.getsize(root_path)} bytes)")
print(f"✓ Generated privacy.html in docs/ ({os.path.getsize(docs_path)} bytes)")
