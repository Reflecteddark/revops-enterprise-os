import base64
from pathlib import Path

def update_landing():
    xlsx_path = Path("presentation/RevOps_Platform_Demo_Sample.xlsx")
    with open(xlsx_path, "rb") as f:
        b64_str = base64.b64encode(f.read()).decode("utf-8")

    landing_path = Path("presentation/landing.html")
    with open(landing_path, "r", encoding="utf-8") as f:
        html = f.read()

    # 1. Add Open Graph tags
    og_tags = """    <meta name="description" content="ИИ-Супервайзер звонков отдела продаж и Радар Выручки: контроль 100% звонков в amoCRM / Битрикс24 по 13 критериям. Ликвидация утечек от 1.2M ₽/мес.">
    <meta property="og:title" content="RevOps Enterprise OS V17.6 — ИИ-Супервайзер звонков и Радар Выручки">
    <meta property="og:description" content="Контроль 100% звонков в amoCRM / Битрикс24 по 13 критериям. Ликвидация утечек выручки от 1.2M ₽/мес. 152-ФЗ РФ.">
    <meta property="og:type" content="website">"""

    if '<meta property="og:title"' not in html:
        html = html.replace(
            "<title>RevOps Enterprise OS V17.6 — ИИ-Супервайзер звонков и Радар Выручки</title>",
            "<title>RevOps Enterprise OS V17.6 — ИИ-Супервайзер звонков и Радар Выручки</title>\n" + og_tags
        )

    # 2. Add Excel download button in Hero
    hero_old = """                    <div class="flex flex-col sm:flex-row gap-4 mb-10">
                        <a href="#cta" class="inline-flex items-center justify-center gap-2 px-7 py-3.5 bg-gradient-to-r from-indigo-600 to-indigo-500 hover:from-indigo-500 hover:to-indigo-400 text-white font-semibold rounded-xl transition-all shadow-xl shadow-indigo-500/25 text-sm sm:text-base">
                            <i class="fas fa-microphone-alt"></i>
                            Бесплатный аудит 3 звонков
                        </a>
                        <a href="#calculator" class="inline-flex items-center justify-center gap-2 px-7 py-3.5 bg-slate-800/80 hover:bg-slate-700/80 text-slate-200 font-semibold rounded-xl border border-slate-700 transition-all text-sm sm:text-base">
                            <i class="fas fa-calculator"></i>
                            Рассчитать потери отдела
                        </a>
                    </div>"""

    hero_new = """                    <div class="flex flex-col sm:flex-row gap-3 mb-10">
                        <a href="#cta" class="inline-flex items-center justify-center gap-2 px-6 py-3.5 bg-gradient-to-r from-indigo-600 to-indigo-500 hover:from-indigo-500 hover:to-indigo-400 text-white font-semibold rounded-xl transition-all shadow-xl shadow-indigo-500/25 text-sm sm:text-base">
                            <i class="fas fa-microphone-alt"></i>
                            Бесплатный аудит 3 звонков
                        </a>
                        <button onclick="downloadDemoExcel()" class="inline-flex items-center justify-center gap-2 px-6 py-3.5 bg-emerald-600/20 hover:bg-emerald-600/30 text-emerald-300 font-semibold rounded-xl border border-emerald-500/30 transition-all text-sm sm:text-base">
                            <i class="fas fa-file-excel text-emerald-400"></i>
                            Скачать демо Excel (.xlsx)
                        </button>
                        <a href="#calculator" class="inline-flex items-center justify-center gap-2 px-5 py-3.5 bg-slate-800/80 hover:bg-slate-700/80 text-slate-300 font-semibold rounded-xl border border-slate-700 transition-all text-sm sm:text-base">
                            <i class="fas fa-calculator"></i>
                            Калькулятор
                        </a>
                    </div>"""

    if hero_old in html:
        html = html.replace(hero_old, hero_new)

    # 3. Add Excel download button in Product Component 1
    comp1_anchor = """                    </div>
                    <div class="mt-5 pt-5 border-t border-slate-700/50">
                        <div class="flex items-center justify-between">
                            <span class="text-xs text-slate-500">Формулы и движок</span>"""

    comp1_replacement = """<button onclick="downloadDemoExcel()" class="mt-3 w-full py-2 px-3 bg-emerald-600/20 hover:bg-emerald-600/30 border border-emerald-500/30 text-emerald-300 rounded-lg text-xs font-semibold flex items-center justify-center gap-2 transition">
                            <i class="fas fa-download"></i> Скачать демо-шаблон Excel (.xlsx)
                        </button>
                    </div>
                    <div class="mt-4 pt-4 border-t border-slate-700/50">
                        <div class="flex items-center justify-between">
                            <span class="text-xs text-slate-500">Формулы и движок</span>"""

    if comp1_anchor in html:
        html = html.replace(comp1_anchor, comp1_replacement, 1)

    # 4. Add 152-FZ consent checkbox in form
    consent_checkbox = """                    <div class="flex items-start gap-2 text-left pt-1">
                        <input type="checkbox" id="formConsent" required checked class="mt-1 rounded bg-slate-800 border-slate-700 text-indigo-600 focus:ring-indigo-500">
                        <label for="formConsent" class="text-xs text-slate-400 leading-tight">
                            Я даю согласие на обработку персональных данных в соответствии с <span class="text-indigo-400 underline cursor-pointer">152-ФЗ РФ</span>. Гарантия конфиденциальности.
                        </label>
                    </div>
                    <button type="submit\""""

    if '<input type="checkbox" id="formConsent"' not in html:
        html = html.replace('<button type="submit"', consent_checkbox, 1)

    # 5. Add JS download function
    js_code = f"""
        // Embedded Base64 Demo Excel (.xlsx)
        const DEMO_EXCEL_B64 = "{b64_str}";

        function downloadDemoExcel() {{
            try {{
                const byteCharacters = atob(DEMO_EXCEL_B64);
                const byteNumbers = new Array(byteCharacters.length);
                for (let i = 0; i < byteCharacters.length; i++) {{
                    byteNumbers[i] = byteCharacters.charCodeAt(i);
                }}
                const byteArray = new Uint8Array(byteNumbers);
                const blob = new Blob([byteArray], {{ type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }});
                const link = document.createElement('a');
                link.href = URL.createObjectURL(blob);
                link.download = 'RevOps_Platform_Demo_Sample.xlsx';
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            }} catch (err) {{
                console.error('Ошибка скачивания файла:', err);
                alert('Не удалось скачать файл. Попробуйте еще раз.');
            }}
        }}
"""

    if "function downloadDemoExcel()" not in html:
        html = html.replace("function calculate() {", js_code + "\n        function calculate() {")

    # 6. Update formSuccess card to have Telegram button
    old_success = """        <!-- Form Success State -->
        <div id="formSuccess" class="hidden text-center py-8">
            <div class="w-16 h-16 rounded-full bg-emerald-500/20 text-emerald-400 flex items-center justify-center mx-auto mb-4 text-2xl">
                <i class="fas fa-check"></i>
            </div>
            <h3 class="font-display text-2xl font-bold text-white mb-2">Заявка принята!</h3>
            <p class="text-slate-400 text-sm max-w-md mx-auto">Мы свяжемся с вами в течение 15 минут в Telegram или по телефону для получения 3 аудиозаписей.</p>
        </div>"""

    new_success = """        <!-- Form Success State -->
        <div id="formSuccess" class="hidden text-center py-8">
            <div class="w-16 h-16 rounded-full bg-emerald-500/20 text-emerald-400 flex items-center justify-center mx-auto mb-4 text-2xl">
                <i class="fas fa-check"></i>
            </div>
            <h3 class="font-display text-2xl font-bold text-white mb-2">Заявка принята!</h3>
            <p class="text-slate-400 text-sm max-w-md mx-auto mb-6">Мы свяжемся с вами в течение 15 минут в Telegram или по телефону для получения 3 аудиозаписей.</p>
            <a id="tgDirectBtn" href="https://t.me/Reflecteddark" target="_blank" class="inline-flex items-center gap-2 px-6 py-3 bg-sky-600 hover:bg-sky-500 text-white rounded-xl font-bold text-sm transition shadow-lg shadow-sky-600/30">
                <i class="fab fa-telegram-plane"></i> Отправить звонки в Telegram прямо сейчас
            </a>
        </div>"""

    if old_success in html:
        html = html.replace(old_success, new_success)

    # 7. Update form submission handler
    old_submit = """            leadForm.style.display = 'none';
            formSuccess.classList.remove('hidden');"""

    new_submit = """            leadForm.style.display = 'none';
            formSuccess.classList.remove('hidden');
            
            const tgMsg = encodeURIComponent(`Здравствуйте! Хочу получить экспресс-аудит 3 звонков.\\nИмя: ${name}\\nТелефон: ${phone}\\nTelegram: ${tg}\\nCRM: ${crm}`);
            const tgUrl = `https://t.me/Reflecteddark?text=${tgMsg}`;
            const tgBtn = document.getElementById('tgDirectBtn');
            if (tgBtn) {
                tgBtn.href = tgUrl;
            }"""

    if old_submit in html:
        html = html.replace(old_submit, new_submit, 1)

    with open(landing_path, "w", encoding="utf-8") as f:
        f.write(html)

    print("presentation/landing.html updated successfully with demo Excel download and features!")

if __name__ == "__main__":
    update_landing()
