# -*- coding: utf-8 -*-
"""
Comprehensive Automated Website Diagnostic Audit for ai-rop.ru
"""

import os
import re
import urllib.request
import ssl
import json
import xml.etree.ElementTree as ET

repo = r"C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os"
index_file = os.path.join(repo, "index.html")

with open(index_file, "r", encoding="utf-8") as f:
    content = f.read()

report = {}

# 1. HTML STRUCTURE & HEADINGS
h1_tags = re.findall(r'<h1[^>]*>(.*?)</h1>', content, re.DOTALL | re.IGNORECASE)
h2_tags = re.findall(r'<h2[^>]*>(.*?)</h2>', content, re.DOTALL | re.IGNORECASE)
h3_tags = re.findall(r'<h3[^>]*>(.*?)</h3>', content, re.DOTALL | re.IGNORECASE)

report['headings'] = {
    'h1_count': len(h1_tags),
    'h1_text': [re.sub(r'<[^>]+>', ' ', h).strip() for h in h1_tags],
    'h2_count': len(h2_tags),
    'h3_count': len(h3_tags)
}

# 2. META & SEO CHECKS
title_match = re.search(r'<title>(.*?)</title>', content, re.IGNORECASE)
desc_match = re.search(r'<meta name="description" content="([^"]*)"', content, re.IGNORECASE)
canonical_match = re.search(r'<link rel="canonical" href="([^"]*)"', content, re.IGNORECASE)

report['seo'] = {
    'title': title_match.group(1) if title_match else None,
    'title_len': len(title_match.group(1)) if title_match else 0,
    'desc': desc_match.group(1) if desc_match else None,
    'desc_len': len(desc_match.group(1)) if desc_match else 0,
    'canonical': canonical_match.group(1) if canonical_match else None,
    'has_og_image': 'og:image' in content,
    'has_schema_ldjson': 'application/ld+json' in content
}

# 3. ASSETS INTEGRITY
css_links = re.findall(r'<link[^>]+href=["\']([^"\']+\.css)["\']', content)
images = re.findall(r'<img[^>]+src=["\']([^"\']+)["\']', content)

missing_imgs = []
img_sizes = {}
for img in images:
    if not img.startswith('http') and not img.startswith('data:'):
        p = os.path.join(repo, img.replace('/', os.sep))
        if not os.path.exists(p):
            missing_imgs.append(img)
        else:
            img_sizes[img] = os.path.getsize(p)

report['assets'] = {
    'css_links': css_links,
    'total_images': len(images),
    'missing_images': missing_imgs,
    'image_sizes_kb': {k: f"{v/1024:.1f} KB" for k, v in img_sizes.items()}
}

# 4. CONVERSION & TRACKING
report['conversion'] = {
    'has_lead_form': 'id="leadForm"' in content,
    'has_phone_mask': 'formatPhoneNumber' in content,
    'has_crm_qualifier': 'crmPill' in content,
    'has_team_qualifier': 'teamPill' in content,
    'has_sample_pdf_button': 'sample-audit-report.html' in content,
    'has_founder_widget': 'founderFloatWidget' in content,
    'has_cookie_banner': 'cookieBanner' in content,
    'has_exit_intent': 'exitIntentModal' in content or 'exitModal' in content,
    'yandex_metrika': '100714777' in content,
    'google_analytics': 'G-H8N72E9D4E' in content,
    'no_cache_removed': not any(k in content.lower() for k in ['cache-control', 'pragma', 'expires']),
    'tailwind_cdn_removed': 'cdn.tailwindcss.com' not in content
}

# 5. LIVE HTTP RESPONSE TEST
live_status = {}
try:
    ctx = ssl.create_default_context()
    req = urllib.request.Request("https://ai-rop.ru", headers={"User-Agent": "Mozilla/5.0 (Audit-Bot)"})
    with urllib.request.urlopen(req, context=ctx, timeout=8) as res:
        live_status['code'] = res.status
        live_status['headers'] = dict(res.headers)
        live_status['url'] = res.geturl()
except Exception as e:
    live_status['error'] = str(e)

report['live_http'] = live_status

print(json.dumps(report, ensure_ascii=False, indent=2))
