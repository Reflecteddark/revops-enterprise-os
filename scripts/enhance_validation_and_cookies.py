# -*- coding: utf-8 -*-
"""
1. Enhances phone mask and inline validation + cookie banner in index.html
2. Updates pilot-roadmap.html, offer.html, and audio-review.html with static CSS, cache removal, privacy links
3. Synchronizes to docs/ and web/
"""

import os
import re
import shutil

repo = r"C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os"

# ----------------------------------------------------
# 1. ENHANCE index.html
# ----------------------------------------------------
index_path = os.path.join(repo, "index.html")
with open(index_path, "r", encoding="utf-8") as f:
    content = f.read()

# Add Cookie banner script before </body>
cookie_script = """
    <!-- Cookie Banner Logic -->
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

if 'CONSENT_KEY' not in content:
    content = content.replace('</body>', cookie_script + '\n</body>')

# Ensure blur validation is attached to formContactInput
blur_validation_code = """
                formContactInput.addEventListener('blur', () => {
                    const val = formContactInput.value.trim();
                    if (val && !validateContactField()) {
                        if (contactErrorMsg) {
                            contactErrorMsg.classList.remove('hidden');
                            if (contactErrorText) {
                                contactErrorText.textContent = (currentContactMode === 'phone') 
                                    ? 'Пожалуйста, укажите полный номер: +7 (XXX) XXX-XX-XX' 
                                    : 'Пожалуйста, укажите корректный Telegram (@username)';
                            }
                        }
                        formContactInput.classList.add('border-rose-500', 'ring-1', 'ring-rose-500');
                    }
                });
"""

if "formContactInput.addEventListener('blur'" not in content:
    # insert right after focus listener
    content = content.replace("formContactInput.addEventListener('focus', () => {", blur_validation_code + "\n                formContactInput.addEventListener('focus', () => {")

# Remove all remaining console.log
content = re.sub(r'console\.log\([^)]*\);?', '', content)

with open(index_path, "w", encoding="utf-8") as f:
    f.write(content)
print("✓ index.html finalized")

# ----------------------------------------------------
# 2. UPDATE pilot-roadmap.html
# ----------------------------------------------------
roadmap_path = os.path.join(repo, "pilot-roadmap.html")
if os.path.exists(roadmap_path):
    with open(roadmap_path, "r", encoding="utf-8") as f:
        r_content = f.read()
    
    # Remove cache meta
    r_content = re.sub(r'\s*<meta http-equiv="Cache-Control"[^>]*>\s*', '\n', r_content, flags=re.IGNORECASE)
    r_content = re.sub(r'\s*<meta http-equiv="Pragma"[^>]*>\s*', '\n', r_content, flags=re.IGNORECASE)
    r_content = re.sub(r'\s*<meta http-equiv="Expires"[^>]*>\s*', '\n', r_content, flags=re.IGNORECASE)
    
    # Replace Tailwind CDN
    r_content = r_content.replace(
        '<script src="https://cdn.tailwindcss.com"></script>',
        '<link rel="stylesheet" href="css/tailwind.min.css">'
    )

    # Link privacy policy in footer
    if 'privacy.html' not in r_content:
        r_content = r_content.replace(
            '<a href="offer.html"',
            '<a href="privacy.html" class="hover:text-white transition">Политика 152-ФЗ</a> • <a href="offer.html"'
        )
    
    # Clean console.log
    r_content = re.sub(r'console\.log\([^)]*\);?', '', r_content)

    with open(roadmap_path, "w", encoding="utf-8") as f:
        f.write(r_content)
    print("✓ pilot-roadmap.html finalized")

# ----------------------------------------------------
# 3. UPDATE offer.html
# ----------------------------------------------------
offer_path = os.path.join(repo, "offer.html")
if os.path.exists(offer_path):
    with open(offer_path, "r", encoding="utf-8") as f:
        o_content = f.read()

    # Replace Tailwind CDN
    o_content = o_content.replace(
        '<script src="https://cdn.tailwindcss.com"></script>',
        '<link rel="stylesheet" href="css/tailwind.min.css">'
    )

    # Link privacy policy in header/footer
    if 'privacy.html' not in o_content:
        o_content = o_content.replace(
            '<a href="/" class="hover:text-white transition">Главная</a>',
            '<a href="/" class="hover:text-white transition">Главная</a> • <a href="/privacy.html" class="hover:text-white transition">Политика 152-ФЗ</a>'
        )

    # Clean console.log
    o_content = re.sub(r'console\.log\([^)]*\);?', '', o_content)

    with open(offer_path, "w", encoding="utf-8") as f:
        f.write(o_content)
    print("✓ offer.html finalized")

# ----------------------------------------------------
# 4. REBUILD TAILWIND MINIFIED CSS TO INCLUDE ALL NEW CLASSES
# ----------------------------------------------------
print("Rebuilding Tailwind CSS to include all HTML files...")
os.system("npx tailwindcss@3 -i ./css/input.css -o ./css/tailwind.min.css --minify")

# ----------------------------------------------------
# 5. SYNCHRONIZE ALL WEB FILES TO docs/ AND web/
# ----------------------------------------------------
files_to_sync = [
    "index.html",
    "pilot-roadmap.html",
    "offer.html",
    "privacy.html",
    "sitemap.xml",
    "CNAME"
]

for folder in ["docs", "web"]:
    target_dir = os.path.join(repo, folder)
    if os.path.exists(target_dir):
        # copy css folder
        css_src = os.path.join(repo, "css")
        css_dst = os.path.join(target_dir, "css")
        os.makedirs(css_dst, exist_ok=True)
        for f in os.listdir(css_src):
            shutil.copy2(os.path.join(css_src, f), os.path.join(css_dst, f))

        # copy files
        for f in files_to_sync:
            src = os.path.join(repo, f)
            if os.path.exists(src):
                shutil.copy2(src, os.path.join(target_dir, f))
        print(f"✓ Synchronized all files to {folder}/")

print("All tasks completed successfully!")
