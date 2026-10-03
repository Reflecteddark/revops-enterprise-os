# -*- coding: utf-8 -*-
"""
Generate and validate sitemap.xml across repo
"""

import os
import xml.etree.ElementTree as ET

sitemap_content = """<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
  <url>
    <loc>https://ai-rop.ru/</loc>
    <lastmod>2026-10-03</lastmod>
    <changefreq>daily</changefreq>
    <priority>1.0</priority>
  </url>
  <url>
    <loc>https://ai-rop.ru/pilot-roadmap.html</loc>
    <lastmod>2026-10-03</lastmod>
    <changefreq>weekly</changefreq>
    <priority>0.8</priority>
  </url>
  <url>
    <loc>https://ai-rop.ru/offer.html</loc>
    <lastmod>2026-10-03</lastmod>
    <changefreq>monthly</changefreq>
    <priority>0.7</priority>
  </url>
  <url>
    <loc>https://ai-rop.ru/privacy.html</loc>
    <lastmod>2026-10-03</lastmod>
    <changefreq>monthly</changefreq>
    <priority>0.6</priority>
  </url>
</urlset>
"""

# Validate XML first
root = ET.fromstring(sitemap_content)
print(f"✓ XML is 100% valid! Found {len(root)} URL entries.")

repo = r"C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os"
targets = [
    os.path.join(repo, "sitemap.xml"),
    os.path.join(repo, "docs", "sitemap.xml"),
    os.path.join(repo, "web", "sitemap.xml"),
]

for t in targets:
    if os.path.exists(os.path.dirname(t)):
        with open(t, "w", encoding="utf-8") as f:
            f.write(sitemap_content.strip() + "\n")
        print(f"✓ Saved: {t}")
