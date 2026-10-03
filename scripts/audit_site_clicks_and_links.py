#!/usr/bin/env python3
"""Comprehensive clickability, navigation, and links audit across the entire site."""
import re
import sys
from bs4 import BeautifulSoup
import os

sys.stdout.reconfigure(encoding='utf-8')

pages = [
    'web/index.html',
    'web/offer.html',
    'web/privacy.html',
    'web/pilot-roadmap.html'
]

print("==================================================")
print("AUDITING SITE CLICKABILITY, NAVIGATION & LINKS")
print("==================================================\n")

for p in pages:
    if not os.path.exists(p):
        print(f"[!] File not found: {p}")
        continue

    with open(p, 'r', encoding='utf-8') as f:
        html = f.read()

    soup = BeautifulSoup(html, 'html.parser')
    print(f"=== {p} ===")
    
    # Title
    title = soup.find('title')
    print(f"  Title: {title.text if title else 'NO TITLE'}")

    # Subpage navigation check
    if p != 'web/index.html':
        back_links = []
        for a in soup.find_all('a'):
            txt = a.get_text().strip()
            href = a.get('href', '')
            if href in ('/', '/index.html', 'index.html', '#') or 'главн' in txt.lower() or 'назад' in txt.lower():
                back_links.append((txt, href))
        print(f"  Back to main links found: {len(back_links)}")
        for txt, href in back_links:
            print(f"    - \"{txt}\" -> {href}")

    # Anchor links check
    all_a = soup.find_all('a')
    all_buttons = soup.find_all('button')
    print(f"  Total <a> links: {len(all_a)} | Total <button> elements: {len(all_buttons)}")

    broken_anchors = []
    empty_links = []
    external_links = []
    
    # Collect all element IDs in document
    element_ids = set()
    for tag in soup.find_all(True):
        if tag.get('id'):
            element_ids.add(tag['id'])

    for a in all_a:
        href = a.get('href')
        if not href or href == '#' or href.startswith('javascript:'):
            # check if it has an onclick or role=button or id
            onclick = a.get('onclick')
            aid = a.get('id')
            aclass = a.get('class', [])
            empty_links.append((a.get_text(strip=True)[:30], href, aid, onclick))
        elif href.startswith('#'):
            target_id = href[1:]
            if target_id not in element_ids:
                broken_anchors.append((a.get_text(strip=True)[:30], href))
        elif href.startswith('http'):
            external_links.append(href)

    if broken_anchors:
        print(f"  [!] BROKEN ANCHOR LINKS ({len(broken_anchors)}):")
        for txt, href in broken_anchors:
            print(f"      - \"{txt}\" points to non-existent ID {href}")
    else:
        print(f"  [+] All internal anchor #hash links exist in DOM!")

    if empty_links:
        print(f"  [i] Links with href='#' or empty ({len(empty_links)}):")
        for txt, href, aid, onclick in empty_links[:8]:
            print(f"      - \"{txt}\" (id={aid}, onclick={onclick})")
        if len(empty_links) > 8:
            print(f"      ... and {len(empty_links)-8} more")

    # JS getElementById check
    ids_in_js = re.findall(r'getElementById\([\'\"]([^\'\"]+)[\'\"]\)', html)
    missing_ids = [gid for gid in set(ids_in_js) if gid not in element_ids]
    if missing_ids:
        print(f"  [!] Missing DOM IDs referenced in JS ({len(missing_ids)}): {missing_ids}")
    else:
        print(f"  [+] All {len(set(ids_in_js))} IDs referenced in getElementById exist in DOM!")

    print()

print("==================================================")
print("AUDIT COMPLETE")
print("==================================================")
