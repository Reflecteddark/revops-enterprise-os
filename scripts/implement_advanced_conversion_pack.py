# -*- coding: utf-8 -*-
"""
Implements 5 advanced conversion optimizations on index.html:
1. Sample PDF Audit Report CTA button & links
2. CRM & Team size 2-click interactive qualifier in #leadForm
3. Floating founder quick-contact widget (Telegram / MAX)
4. Dynamic ROI calculator-to-form bridge
5. UTM & CRM parameter dynamic header / multilanding logic
"""

import os
import re

repo = r"C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os"
index_path = os.path.join(repo, "index.html")

with open(index_path, "r", encoding="utf-8") as f:
    content = f.read()

# ----------------------------------------------------
# 1. SAMPLE AUDIT REPORT BUTTONS
# ----------------------------------------------------
# Under audio player
audio_sample_btn = """
                <!-- Sample Audit Report Link -->
                <div class="mt-4 pt-4 border-t border-slate-800 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs">
                    <span class="text-slate-400">Хотите посмотреть, как выглядит итоговый PDF-отчет аудита?</span>
                    <a href="sample-audit-report.html" target="_blank" class="text-emerald-400 hover:text-emerald-300 font-bold underline flex items-center gap-1.5 transition">
                        <i class="fas fa-file-pdf text-rose-400"></i>
                        <span>Посмотреть образец отчёта по 3 звонкам (PDF) →</span>
                    </a>
                </div>
"""

if 'sample-audit-report.html' not in content:
    # insert before end of audio player card
    content = content.replace(
        '<!-- Transcript lines -->',
        audio_sample_btn + '\n                <!-- Transcript lines -->'
    )

# Under form
form_sample_link = """
                        <!-- Sample Audit Link under Form -->
                        <div class="pt-2 text-center">
                            <a href="sample-audit-report.html" target="_blank" class="text-xs text-slate-400 hover:text-emerald-400 underline font-mono flex items-center justify-center gap-1.5 transition">
                                <i class="fas fa-file-pdf text-rose-400"></i>
                                <span>Посмотреть образец готового PDF-разбора 3 звонков (0 ₽)</span>
                            </a>
                        </div>
"""
if 'Посмотреть образец готового PDF-разбора' not in content:
    content = content.replace(
        '<!-- Соцдоказательство у формы -->',
        form_sample_link + '\n                        <!-- Соцдоказательство у формы -->'
    )

# ----------------------------------------------------
# 2. CRM & TEAM SIZE QUALIFIERS IN #leadForm
# ----------------------------------------------------
qualifier_html = """
                        <!-- CRM QUALIFIER (1-CLICK) -->
                        <div>
                            <label class="block text-xs font-semibold text-slate-300 uppercase tracking-wider mb-2">Какая CRM установлена в вашей компании?</label>
                            <div class="grid grid-cols-2 sm:grid-cols-4 gap-2" id="crmPills">
                                <button type="button" onclick="selectCrmPill('amoCRM')" id="crmPill_amoCRM" class="crm-pill px-2.5 py-2 rounded-xl border text-xs font-bold text-center transition bg-emerald-500/15 border-emerald-500 text-emerald-300 shadow-md shadow-emerald-500/10">
                                    <span>amoCRM</span>
                                </button>
                                <button type="button" onclick="selectCrmPill('Битрикс24')" id="crmPill_Битрикс24" class="crm-pill px-2.5 py-2 rounded-xl border text-xs font-medium text-center transition bg-slate-900 border-slate-800 text-slate-300 hover:border-slate-700">
                                    <span>Битрикс24</span>
                                </button>
                                <button type="button" onclick="selectCrmPill('МойСклад')" id="crmPill_МойСклад" class="crm-pill px-2.5 py-2 rounded-xl border text-xs font-medium text-center transition bg-slate-900 border-slate-800 text-slate-300 hover:border-slate-700">
                                    <span>МойСклад</span>
                                </button>
                                <button type="button" onclick="selectCrmPill('Другая / Excel')" id="crmPill_Другая_Excel" class="crm-pill px-2.5 py-2 rounded-xl border text-xs font-medium text-center transition bg-slate-900 border-slate-800 text-slate-300 hover:border-slate-700">
                                    <span>Другая / Excel</span>
                                </button>
                            </div>
                            <input type="hidden" id="selectedCrmInput" value="amoCRM">
                        </div>

                        <!-- TEAM SIZE QUALIFIER (1-CLICK) -->
                        <div>
                            <label class="block text-xs font-semibold text-slate-300 uppercase tracking-wider mb-2">Сколько менеджеров звонит клиентам?</label>
                            <div class="grid grid-cols-3 gap-2" id="teamPills">
                                <button type="button" onclick="selectTeamPill('2-5 менеджеров')" id="teamPill_2-5" class="team-pill px-2.5 py-2 rounded-xl border text-xs font-bold text-center transition bg-emerald-500/15 border-emerald-500 text-emerald-300">
                                    <span>2–5 чел</span>
                                </button>
                                <button type="button" onclick="selectTeamPill('6-15 менеджеров')" id="teamPill_6-15" class="team-pill px-2.5 py-2 rounded-xl border text-xs font-medium text-center transition bg-slate-900 border-slate-800 text-slate-300 hover:border-slate-700">
                                    <span>6–15 чел</span>
                                </button>
                                <button type="button" onclick="selectTeamPill('16+ менеджеров')" id="teamPill_16+" class="team-pill px-2.5 py-2 rounded-xl border text-xs font-medium text-center transition bg-slate-900 border-slate-800 text-slate-300 hover:border-slate-700">
                                    <span>16+ чел</span>
                                </button>
                            </div>
                            <input type="hidden" id="selectedTeamInput" value="2-5 менеджеров">
                        </div>
"""

# Dynamic calculation callout banner inside form (for feature 4)
calc_callout_html = """
                        <!-- DYNAMIC CALC LOSS CALLOUT -->
                        <div id="dynamicLossCallout" class="hidden p-3.5 rounded-2xl bg-amber-500/10 border border-amber-500/30 text-amber-300 text-xs">
                            <div class="flex items-center gap-2 font-bold mb-1">
                                <i class="fas fa-calculator text-amber-400"></i>
                                <span>ПЕРСОНАЛЬНЫЙ АУДИТ ПО ВАШЕМУ РАСЧЕТУ</span>
                            </div>
                            <p class="text-slate-300 text-[11px] leading-tight" id="dynamicLossCalloutText">
                                Оцифруем точки слива рассчитанных ~480 000 ₽/мес на 3 реальных звонках вашей команды.
                            </p>
                        </div>
"""

if 'id="crmPills"' not in content:
    content = content.replace(
        '<div class="space-y-4">',
        calc_callout_html + '\n' + qualifier_html + '\n                        <div class="space-y-4">'
    )

# ----------------------------------------------------
# 3. FLOATING FOUNDER QUICK CONTACT WIDGET
# ----------------------------------------------------
floating_widget_html = """
    <!-- ========================================================
         FLOATING FOUNDER QUICK CONTACT WIDGET (CONVERSION ENGINE)
         ======================================================== -->
    <div id="founderFloatWidget" class="fixed bottom-6 right-6 z-40 flex flex-col items-end">
        <!-- Floating Speech Tooltip Bubble -->
        <div id="founderTooltip" class="mb-3 max-w-xs p-3.5 rounded-2xl bg-slate-900/95 border border-emerald-500/50 shadow-2xl backdrop-blur-xl text-xs text-white transition-all duration-300 transform translate-y-2 opacity-0 pointer-events-none relative">
            <button type="button" id="closeFounderTooltip" class="absolute top-2 right-2 text-slate-400 hover:text-white text-[10px]">
                <i class="fas fa-times"></i>
            </button>
            <div class="font-bold text-emerald-400 flex items-center gap-1.5 mb-1">
                <span class="w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></span>
                <span>Дмитрий Федотов</span>
            </div>
            <p class="text-slate-300 text-[11px] leading-tight pr-3">
                Пришлите 3 звонка в Telegram — через 20 минут выдам оцифровку сливов выручки (0 ₽).
            </p>
        </div>

        <!-- Floating Action Button -->
        <div class="flex items-center gap-2.5">
            <div class="hidden sm:flex px-3 py-1.5 rounded-full bg-slate-900/90 border border-slate-700/80 text-[11px] font-mono font-semibold text-slate-300 shadow-xl backdrop-blur-md">
                Основатель онлайн
            </div>
            <button type="button" id="founderFloatBtn" class="relative group w-14 h-14 rounded-2xl bg-gradient-to-tr from-emerald-500 to-cyan-500 p-0.5 shadow-2xl shadow-emerald-500/30 active:scale-95 transition-all hover:scale-105" title="Написать Дмитрию напрямую">
                <div class="w-full h-full rounded-[14px] bg-slate-950 flex items-center justify-center overflow-hidden relative">
                    <img src="images/founder-dmitry.jpg" alt="Дмитрий" class="w-full h-full object-cover object-top group-hover:scale-110 transition-transform">
                    <span class="absolute bottom-1 right-1 w-3.5 h-3.5 rounded-full bg-emerald-400 border-2 border-slate-950 animate-pulse"></span>
                </div>
            </button>
        </div>

        <!-- Floating Quick Menu Popup (Hidden by default) -->
        <div id="founderQuickMenu" class="hidden mt-3 p-4 rounded-3xl bg-slate-900/95 border-2 border-emerald-500/50 shadow-2xl backdrop-blur-2xl w-72 text-left transition-all">
            <div class="flex items-center justify-between pb-3 mb-3 border-b border-slate-800">
                <div class="flex items-center gap-2.5">
                    <img src="images/founder-dmitry.jpg" alt="Дмитрий" class="w-9 h-9 rounded-xl object-cover object-top border border-emerald-400">
                    <div>
                        <div class="text-xs font-bold text-white leading-tight">Дмитрий Федотов</div>
                        <div class="text-[10px] text-emerald-400 font-mono">Основатель RevOps OS</div>
                    </div>
                </div>
                <button type="button" id="closeFounderQuickMenu" class="text-slate-400 hover:text-white text-xs">
                    <i class="fas fa-times"></i>
                </button>
            </div>
            <p class="text-slate-300 text-xs leading-relaxed mb-3">
                Куда вам удобнее прислать 3 вчерашних звонка на бесплатный аудит?
            </p>
            <div class="space-y-2 text-xs">
                <a href="https://t.me/dm1918?text=%D0%94%D0%BC%D0%B8%D1%82%D1%80%D0%B8%D0%B9%2C%20%D0%BF%D1%80%D0%B8%D0%B2%D0%B5%D1%82%D1%81%D1%82%D0%B2%D1%83%D1%8E!%20%D0%A5%D0%BE%D1%87%D1%83%20%D0%B1%D0%B5%D1%81%D0%BF%D0%BB%D0%B0%D1%82%D0%BD%D1%8B%D0%B9%20%D0%B0%D1%83%D0%B4%D0%B8%D1%82%203%20%D0%B7%D0%B2%D0%BE%D0%BD%D0%BA%D0%BE%D0%B2%20%D0%BE%D1%82%D0%B4%D0%B5%D0%BB%D0%B0%20%D0%BF%D1%80%D0%BE%D0%B4%D0%B0%D0%B6." target="_blank" rel="noopener noreferrer" class="w-full py-2.5 px-3 rounded-xl bg-sky-500 hover:bg-sky-400 text-slate-950 font-bold flex items-center justify-center gap-2 shadow-lg transition">
                    <i class="fab fa-telegram text-sm"></i>
                    <span>В Telegram (@dm1918)</span>
                </a>
                <a href="https://t.me/RevOps_Super_Audit_Bot" target="_blank" rel="noopener noreferrer" class="w-full py-2.5 px-3 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 font-semibold flex items-center justify-center gap-2 border border-slate-700 transition">
                    <i class="fas fa-robot text-cyan-400"></i>
                    <span>В Telegram-бот аудита</span>
                </a>
                <a href="https://max.ru/u/f9LHodD0cOLRcNKRakz94FpZjmUJ25lYzJUMzPBJyLKwxM1Tzzai1aF-dTg" target="_blank" rel="noopener noreferrer" class="w-full py-2.5 px-3 rounded-xl bg-purple-600/30 hover:bg-purple-600/50 text-purple-200 font-semibold flex items-center justify-center gap-2 border border-purple-500/40 transition">
                    <i class="fas fa-comment-dots text-purple-400"></i>
                    <span>В мессенджер MAX</span>
                </a>
            </div>
        </div>
    </div>
"""

if 'id="founderFloatWidget"' not in content:
    content = content.replace('</body>', floating_widget_html + '\n</body>')

# ----------------------------------------------------
# 4. JAVASCRIPT CONTROLLERS: QUALIFIERS, ROI BRIDGE & MULTILANDING
# ----------------------------------------------------
new_js_logic = """
        // ========================================================
        // 1. CRM & TEAM SIZE QUALIFIER CONTROLLER
        // ========================================================
        window.selectCrmPill = function(crmName) {
            const input = document.getElementById('selectedCrmInput');
            if (input) input.value = crmName;
            const pills = document.querySelectorAll('.crm-pill');
            pills.forEach(p => {
                p.className = 'crm-pill px-2.5 py-2 rounded-xl border text-xs font-medium text-center transition bg-slate-900 border-slate-800 text-slate-300 hover:border-slate-700';
            });
            const selectedBtn = document.getElementById('crmPill_' + crmName.replace(/\\s+|\\//g, '_'));
            if (selectedBtn) {
                selectedBtn.className = 'crm-pill px-2.5 py-2 rounded-xl border text-xs font-bold text-center transition bg-emerald-500/15 border-emerald-500 text-emerald-300 shadow-md shadow-emerald-500/10';
            }
        };

        window.selectTeamPill = function(teamSize) {
            const input = document.getElementById('selectedTeamInput');
            if (input) input.value = teamSize;
            const pills = document.querySelectorAll('.team-pill');
            pills.forEach(p => {
                p.className = 'team-pill px-2.5 py-2 rounded-xl border text-xs font-medium text-center transition bg-slate-900 border-slate-800 text-slate-300 hover:border-slate-700';
            });
            let btnId = 'teamPill_2-5';
            if (teamSize.startsWith('6')) btnId = 'teamPill_6-15';
            else if (teamSize.startsWith('16')) btnId = 'teamPill_16+';
            const selectedBtn = document.getElementById(btnId);
            if (selectedBtn) {
                selectedBtn.className = 'team-pill px-2.5 py-2 rounded-xl border text-xs font-bold text-center transition bg-emerald-500/15 border-emerald-500 text-emerald-300';
            }
        };

        // ========================================================
        // 2. DYNAMIC ROI CALCULATOR BRIDGE TO FORM
        // ========================================================
        window.claimLossAudit = function() {
            const ctaSection = document.getElementById('cta');
            if (ctaSection) {
                ctaSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
            const lossValElem = document.getElementById('lossVal');
            const callout = document.getElementById('dynamicLossCallout');
            const calloutText = document.getElementById('dynamicLossCalloutText');
            if (lossValElem && callout && calloutText) {
                const lossText = lossValElem.textContent.replace('-', '').trim();
                calloutText.innerHTML = `Оцифруем точки слива рассчитанных ~<strong>${lossText}</strong> потерь на 3 реальных звонках вашей команды.`;
                callout.classList.remove('hidden');
            }
            const phoneInput = document.getElementById('formContact');
            if (phoneInput) {
                setTimeout(() => { phoneInput.focus(); }, 600);
            }
        };

        // ========================================================
        // 3. FLOATING FOUNDER WIDGET CONTROLLER
        // ========================================================
        (function() {
            const btn = document.getElementById('founderFloatBtn');
            const menu = document.getElementById('founderQuickMenu');
            const closeMenuBtn = document.getElementById('closeFounderQuickMenu');
            const tooltip = document.getElementById('founderTooltip');
            const closeTooltip = document.getElementById('closeFounderTooltip');

            // Show tooltip after 12s
            setTimeout(() => {
                if (tooltip && !sessionStorage.getItem('founder_tip_dismissed')) {
                    tooltip.classList.remove('translate-y-2', 'opacity-0', 'pointer-events-none');
                    tooltip.classList.add('translate-y-0', 'opacity-100', 'pointer-events-auto');
                }
            }, 12000);

            if (closeTooltip) {
                closeTooltip.addEventListener('click', (e) => {
                    e.stopPropagation();
                    tooltip.classList.add('opacity-0', 'pointer-events-none');
                    sessionStorage.setItem('founder_tip_dismissed', 'true');
                });
            }

            if (btn && menu) {
                btn.addEventListener('click', () => {
                    menu.classList.toggle('hidden');
                    if (tooltip) tooltip.classList.add('opacity-0', 'pointer-events-none');
                });
            }

            if (closeMenuBtn && menu) {
                closeMenuBtn.addEventListener('click', () => {
                    menu.classList.add('hidden');
                });
            }
        })();

        // ========================================================
        // 4. MULTILANDING & DYNAMIC UTM / CRM HEADER
        // ========================================================
        (function() {
            const urlParams = new URLSearchParams(window.location.search);
            const crmParam = (urlParams.get('crm') || urlParams.get('utm_term') || urlParams.get('utm_campaign') || '').toLowerCase();
            const heroBadge = document.querySelector('section.hero-glow .inline-flex span:last-child');
            const heroH1 = document.querySelector('section.hero-glow h1');

            if (crmParam.includes('amocrm') || crmParam.includes('amo') || crmParam.includes('амо')) {
                if (heroBadge) heroBadge.textContent = '⚡ Специализированный модуль для amoCRM (API v4) • Радар Выручки';
                if (heroH1) heroH1.innerHTML = 'Найдём сделки в amoCRM, которые отдел продаж<br><span class="text-transparent bg-clip-text bg-gradient-to-r from-emerald-400 via-teal-300 to-cyan-400">теряет прямо сейчас.</span>';
                if (typeof window.selectCrmPill === 'function') window.selectCrmPill('amoCRM');
            } else if (crmParam.includes('bitrix') || crmParam.includes('b24') || crmParam.includes('битрикс')) {
                if (heroBadge) heroBadge.textContent = '⚡ Специализированный модуль для 1С-Битрикс24 (REST API) • Радар Выручки';
                if (heroH1) heroH1.innerHTML = 'Найдём сделки в Битрикс24, которые отдел продаж<br><span class="text-transparent bg-clip-text bg-gradient-to-r from-emerald-400 via-teal-300 to-cyan-400">теряет прямо сейчас.</span>';
                if (typeof window.selectCrmPill === 'function') window.selectCrmPill('Битрикс24');
            } else if (crmParam.includes('telecom') || crmParam.includes('звонк') || crmParam.includes('uis') || crmParam.includes('mango')) {
                if (heroBadge) heroBadge.textContent = '⚡ Прямой поток из облачной АТС (UIS, Mango, Novofon, Asterisk)';
                if (heroH1) heroH1.innerHTML = 'Найдём звонки, на которых менеджеры<br><span class="text-transparent bg-clip-text bg-gradient-to-r from-emerald-400 via-teal-300 to-cyan-400">сливают горячих клиентов.</span>';
            }
        })();
"""

# Insert JS before </body>
content = content.replace('</body>', '<script>' + new_js_logic + '</script>\n</body>')

# Update ROI calculator button to call window.claimLossAudit()
content = content.replace(
    '<a href="#roi" class="w-full py-2.5 bg-slate-800 hover:bg-slate-700 text-white rounded-lg text-xs font-semibold transition text-center border border-slate-700 flex items-center justify-center gap-1.5">\n                                <i class="fas fa-calculator text-emerald-400"></i>\n                                Посчитать свои потери\n                            </a>',
    """<button type="button" onclick="window.claimLossAudit()" class="w-full py-2.5 bg-gradient-to-r from-emerald-500/20 to-cyan-500/20 hover:from-emerald-500/30 hover:to-cyan-500/30 text-emerald-300 rounded-lg text-xs font-bold transition text-center border border-emerald-500/40 flex items-center justify-center gap-1.5 shadow-lg">
                                <i class="fas fa-calculator text-emerald-400"></i>
                                <span id="roiClaimBtnText">Найти эти потери на бесплатном аудите 3 звонков →</span>
                            </button>"""
)

# Connect ROI calculator update function to update the button text dynamically
content = content.replace(
    "lossVal.textContent = '- ' + totalLoss.toLocaleString('ru-RU') + ' ₽ / мес';",
    """lossVal.textContent = '- ' + totalLoss.toLocaleString('ru-RU') + ' ₽ / мес';
                const roiClaimBtnText = document.getElementById('roiClaimBtnText');
                if (roiClaimBtnText) {
                    roiClaimBtnText.textContent = `Найти эти ${totalLoss.toLocaleString('ru-RU')} ₽ потерь на аудите →`;
                }"""
)

# In form submission, include CRM and Team Size in lead data
content = content.replace(
    "const selectedPkgName = packageNames[pkgKey] || 'Экспресс-Аудит (0 ₽)';",
    """const selectedPkgName = packageNames[pkgKey] || 'Экспресс-Аудит (0 ₽)';
                    const selectedCrm = document.getElementById('selectedCrmInput')?.value || 'Не указана';
                    const selectedTeam = document.getElementById('selectedTeamInput')?.value || 'Не указан';"""
)

content = content.replace(
    "successSub.innerText = `Вы выбрали: ${selectedPkgName}. Заявка передана основателю Дмитрию Федотову (@dm1918). Он ответит вам в течение 10–15 минут.`;",
    "successSub.innerText = `Вы выбрали: ${selectedPkgName} (CRM: ${selectedCrm}, отдел: ${selectedTeam}). Заявка передана основателю Дмитрию Федотову (@dm1918). Он ответит вам в течение 10–15 минут.`;"
)

with open(index_path, "w", encoding="utf-8") as f:
    f.write(content)

print(f"✓ index.html successfully upgraded with all 5 conversion engines! ({len(content)} bytes)")
