import re
import sys

sys.stdout.reconfigure(encoding="utf-8")

with open("deployed_page_raw.html", "r", encoding="utf-8") as f:
    html = f.read()

print("=== META & HEAD ===")
title_match = re.search(r"<title>(.*?)</title>", html, re.DOTALL)
print("Title:", title_match.group(1).strip() if title_match else "None")

og_title = re.search(r'<meta property="og:title" content="(.*?)"', html)
print("OG Title:", og_title.group(1) if og_title else "None")

og_desc = re.search(r'<meta property="og:description" content="(.*?)"', html)
print("OG Desc:", og_desc.group(1) if og_desc else "None")

print("\n=== KEY FEATURES IN CODE ===")
print("Has downloadDemoExcel():", "downloadDemoExcel" in html)
print("Has Base64 Excel embedded:", "DEMO_EXCEL_B64" in html)
print("Has Telegram link (t.me):", "t.me" in html)
print("Has 152-FZ consent checkbox:", "formConsent" in html)
print("Mobile menu fix present:", ".mobile-menu {" in html and "display: none" in html)

print("\n=== CALCULATOR FORMULA ===")
calc_idx = html.find("function calculate()")
if calc_idx != -1:
    print(html[calc_idx:calc_idx + 800])
else:
    print("calculate() not found")

print("\n=== LINKS ===")
links = re.findall(r'<a\s+[^>]*href=["\']([^"\']*)["\'][^>]*>(.*?)</a>', html, re.DOTALL)
print(f"Total <a> links: {len(links)}")
for h, text in links[:15]:
    clean_t = re.sub(r"<[^>]+>", "", text).strip()
    print(f" - {h:35s} | {clean_t[:40]}")

print("\n=== FORM SUBMISSION SCRIPT ===")
sub_idx = html.find("leadForm.addEventListener('submit'")
if sub_idx != -1:
    print(html[sub_idx:sub_idx + 600])

print("\n=== EXTERNAL DEPENDENCIES ===")
scripts = re.findall(r'<script[^>]*src=["\']([^"\']*)["\']', html)
styles = re.findall(r'<link[^>]*href=["\']([^"\']*)["\']', html)
print("External scripts:", scripts)
print("External stylesheets:", styles)
