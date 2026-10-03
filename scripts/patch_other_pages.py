# -*- coding: utf-8 -*-
"""
Add footer links and cookie banner to offer.html and pilot-roadmap.html
"""

import os
import re

repo = r"C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os"

cookie_component = """
    <!-- COOKIE CONSENT BANNER (152-ФЗ РФ) -->
    <div id="cookieBanner" class="fixed bottom-4 left-4 right-4 sm:left-auto sm:right-6 sm:max-w-md z-50 p-4 rounded-2xl bg-slate-900/95 border border-slate-700/80 shadow-2xl backdrop-blur-xl transition-all duration-500 transform translate-y-20 opacity-0 pointer-events-none">
        <div class="flex items-start gap-3">
            <div class="w-9 h-9 rounded-xl bg-emerald-500/10 border border-emerald-500/30 flex items-center justify-center text-emerald-400 shrink-0 text-base">
                <i class="fas fa-cookie-bite"></i>
            </div>
            <div class="flex-1">
                <p class="text-xs text-slate-300 leading-relaxed">
                    Мы используем cookie и сервисы аналитики для защиты данных по <a href="privacy.html" class="text-emerald-400 underline hover:text-emerald-300 font-medium">152-ФЗ РФ</a>.
                </p>
                <div class="mt-3 flex items-center gap-2">
                    <button type="button" id="acceptCookieBtn" class="px-4 py-2 bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-bold rounded-xl text-xs transition shadow-md shadow-emerald-500/20">
                        Принять все
                    </button>
                    <a href="privacy.html" class="px-3 py-2 text-slate-400 hover:text-white text-xs transition underline">
                        Подробнее
                    </a>
                </div>
            </div>
            <button type="button" id="closeCookieBtn" class="text-slate-500 hover:text-white p-1 text-xs" title="Закрыть">
                <i class="fas fa-times"></i>
            </button>
        </div>
    </div>
    <script>
        (function() {
            const banner = document.getElementById('cookieBanner');
            const acceptBtn = document.getElementById('acceptCookieBtn');
            const closeBtn = document.getElementById('closeCookieBtn');
            const CONSENT_KEY = 'revops_cookie_consent_v1';
            if (!banner) return;
            try {
                if (!localStorage.getItem(CONSENT_KEY)) {
                    setTimeout(() => {
                        banner.classList.remove('translate-y-20', 'opacity-0', 'pointer-events-none');
                        banner.classList.add('translate-y-0', 'opacity-100', 'pointer-events-auto');
                    }, 1200);
                }
            } catch(e) {}
            function acceptConsent() {
                try { localStorage.setItem(CONSENT_KEY, 'true'); } catch(e) {}
                banner.classList.add('translate-y-20', 'opacity-0', 'pointer-events-none');
                banner.classList.remove('translate-y-0', 'opacity-100', 'pointer-events-auto');
            }
            if (acceptBtn) acceptBtn.addEventListener('click', acceptConsent);
            if (closeBtn) closeBtn.addEventListener('click', acceptConsent);
        })();
    </script>
"""

# 1. Update offer.html
offer_path = os.path.join(repo, "offer.html")
with open(offer_path, "r", encoding="utf-8") as f:
    o_content = f.read()

o_content = o_content.replace(
    '<div class="mt-8 pt-4 border-t border-slate-900 text-center text-xs text-slate-500">\n            &copy; 2026 RevOps OS Pro (ai-rop.ru). Все права защищены.\n        </div>',
    """<div class="mt-8 pt-4 border-t border-slate-900 text-center text-xs text-slate-500">
            <div class="flex items-center justify-center gap-4 mb-2">
                <a href="/" class="hover:text-slate-300 underline">Главная</a>
                <a href="/pilot-roadmap.html" class="hover:text-slate-300 underline">Дорожная карта пилота</a>
                <a href="/privacy.html" class="hover:text-slate-300 underline">Политика 152-ФЗ</a>
            </div>
            &copy; 2026 RevOps OS Pro (ai-rop.ru). Все права защищены.
        </div>"""
)

if 'id="cookieBanner"' not in o_content:
    o_content = o_content.replace('</body>', cookie_component + '\n</body>')

with open(offer_path, "w", encoding="utf-8") as f:
    f.write(o_content)
print("✓ offer.html updated with footer links & cookie banner")

# 2. Update pilot-roadmap.html
roadmap_path = os.path.join(repo, "pilot-roadmap.html")
with open(roadmap_path, "r", encoding="utf-8") as f:
    r_content = f.read()

if 'id="cookieBanner"' not in r_content:
    r_content = r_content.replace('</body>', cookie_component + '\n</body>')

with open(roadmap_path, "w", encoding="utf-8") as f:
    f.write(r_content)
print("✓ pilot-roadmap.html updated with cookie banner")
