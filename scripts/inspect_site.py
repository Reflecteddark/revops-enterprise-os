# -*- coding: utf-8 -*-
import os
import re

repo = r"C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os"
index_path = os.path.join(repo, "index.html")

with open(index_path, "r", encoding="utf-8") as f:
    lines = f.readlines()

print("Total lines:", len(lines))

print("\n--- Sections found ---")
for i, l in enumerate(lines):
    if "<section" in l or "id=\"hero\"" in l or "<!-- HERO" in l.upper():
        print(f"Line {i+1}: {l.strip()[:100]}")

print("\n--- Hero area (around line 900-1400) ---")
for i, l in enumerate(lines):
    if "hero" in l.lower() and "<section" in l:
        print(f"Hero section starts at line {i+1}")
        for j in range(i, min(i+120, len(lines))):
            print(f"{j+1}: {lines[j].rstrip()}")
        break
