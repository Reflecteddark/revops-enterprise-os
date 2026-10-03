# -*- coding: utf-8 -*-
import os

repo = r"C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os"
index_path = os.path.join(repo, "index.html")

with open(index_path, "r", encoding="utf-8") as f:
    lines = f.readlines()

for i, l in enumerate(lines):
    if "<form" in l or "id=\"cta\"" in l or "type=\"tel\"" in l:
        print(f"Line {i+1}: {l.strip()[:100]}")
        for j in range(max(0, i-5), min(len(lines), i+35)):
            print(f"  {j+1}: {lines[j].rstrip()}")
        print("="*60)
