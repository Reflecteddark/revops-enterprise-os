#!/usr/bin/env python3
"""
Apply Dmitry Valeryevich Fedotov's official VTB bank requisites across the entire system.
"""

import os
import sys
import shutil

sys.stdout.reconfigure(encoding='utf-8')

VTB_DATA = {
    "name": "Федотов Дмитрий Валерьевич",
    "status": "Плательщик налога на профессиональный доход (Самозанятый)",
    "inn": "731303201073",
    "bank_name": "Филиал № 6318 Банка ВТБ (ПАО)",
    "bank_address": "443030, г. Самара, ул. Спортивная, д.30",
    "bik": "043601968",
    "bank_inn": "7702070139",
    "bank_kpp": "631543002",
    "corr_account": "30101810422023601968",
    "account": "40817810635184001277",
    "email": "info@ai-rop.ru",
    "telegram": "@dm1918",
    "site": "https://ai-rop.ru",
    "offer_url": "https://ai-rop.ru/offer.html"
}

print("==================================================")
print("APPLYING OFFICIAL VTB REQUISITES")
print("==================================================")
print(f"Recipient: {VTB_DATA['name']}")
print(f"INN: {VTB_DATA['inn']}")
print(f"Bank: {VTB_DATA['bank_name']}")
print(f"BIK: {VTB_DATA['bik']}")
print(f"Corr Acc: {VTB_DATA['corr_account']}")
print(f"Account: {VTB_DATA['account']}\n")

# 1. Update scripts/generate_billing_docs.py
billing_script = "scripts/generate_billing_docs.py"
with open(billing_script, "r", encoding="utf-8") as f:
    b_code = f.read()

# Replace PROVIDER dict
old_provider_match = """PROVIDER = {
    "name": "Федотов Дмитрий",
    "status": "Плательщик налога на профессиональный доход (Самозанятый)",
    "inn": "731303201073",
    "bank_name": "АО «ТИНЬКОФФ БАНК»",
    "bik": "044525974",
    "corr_account": "30101810145250000974",
    "account": "40802810500003849120",
    "email": "info@ai-rop.ru",
    "telegram": "@dm1918",
    "site": "https://ai-rop.ru",
    "offer_url": "https://ai-rop.ru/offer.html"
}"""

new_provider_code = f"""PROVIDER = {{
    "name": "{VTB_DATA['name']}",
    "status": "{VTB_DATA['status']}",
    "inn": "{VTB_DATA['inn']}",
    "bank_name": "{VTB_DATA['bank_name']}",
    "bank_address": "{VTB_DATA['bank_address']}",
    "bik": "{VTB_DATA['bik']}",
    "bank_inn": "{VTB_DATA['bank_inn']}",
    "bank_kpp": "{VTB_DATA['bank_kpp']}",
    "corr_account": "{VTB_DATA['corr_account']}",
    "account": "{VTB_DATA['account']}",
    "email": "{VTB_DATA['email']}",
    "telegram": "{VTB_DATA['telegram']}",
    "site": "{VTB_DATA['site']}",
    "offer_url": "{VTB_DATA['offer_url']}"
}}"""

if old_provider_match in b_code:
    b_code = b_code.replace(old_provider_match, new_provider_code)
else:
    # generic regex replace for PROVIDER
    import re
    b_code = re.sub(r'PROVIDER = \{[^\}]+\}', new_provider_code, b_code)

with open(billing_script, "w", encoding="utf-8") as f:
    f.write(b_code)
print(f"  [+] Updated {billing_script}")

# 2. Update offer.html requisites
for offer_p in ["web/offer.html", "docs/offer.html", "offer.html"]:
    if os.path.exists(offer_p):
        with open(offer_p, "r", encoding="utf-8") as f:
            c = f.read()
        
        # Replace requisites block
        c = c.replace("АО «ТИНЬКОФФ БАНК»", VTB_DATA['bank_name'])
        c = c.replace("40802810500003849120", VTB_DATA['account'])
        c = c.replace("044525974", VTB_DATA['bik'])
        c = c.replace("30101810145250000974", VTB_DATA['corr_account'])
        c = c.replace("Федотов Дмитрий<br>", f"{VTB_DATA['name']}<br>")
        
        with open(offer_p, "w", encoding="utf-8") as f:
            f.write(c)
        print(f"  [+] Updated {offer_p}")

# 3. Regenerate all billing documents
print("\n--- Regenerating Billing Docs ---")
import generate_billing_docs
generate_billing_docs.PROVIDER = VTB_DATA
generate_billing_docs.main()

import convert_guide_to_docx
convert_guide_to_docx.build_guide_docx("docs/Инструкция_Бухгалтерия_Самозанятого_B2B.docx")
convert_guide_to_docx.build_guide_docx("billing/Инструкция_Бухгалтерия_Самозанятого_B2B.docx")

# 4. Mirror all files to User Desktop folder!
desktop_dir = r"C:\Users\strel\Desktop\RevOps Platform\Закрывающие документы"
desktop_dogovor = r"C:\Users\strel\Desktop\RevOps Platform\Договор"
scratch_billing = r"C:\Users\strel\.gemini\antigravity\scratch\billing"

if os.path.exists(desktop_dir):
    print(f"\n--- Copying updated documents to: {desktop_dir} ---")
    for f in os.listdir("billing"):
        src = os.path.join("billing", f)
        dst = os.path.join(desktop_dir, f)
        if os.path.isfile(src):
            shutil.copy2(src, dst)
            print(f"  [+] Mirrored to Desktop: {f}")

if os.path.exists(desktop_dogovor):
    shutil.copy2("web/offer.html", os.path.join(desktop_dogovor, "Публичная_Оферта_RevOps_OS.html"))
    print(f"  [+] Mirrored offer to Desktop/Договор")

if os.path.exists(scratch_billing):
    for f in os.listdir("billing"):
        src = os.path.join("billing", f)
        dst = os.path.join(scratch_billing, f)
        if os.path.isfile(src):
            shutil.copy2(src, dst)

print("\n==================================================")
print("SUCCESS: VTB REQUISITES DEPLOYED EVERYWHERE!")
print("==================================================")
