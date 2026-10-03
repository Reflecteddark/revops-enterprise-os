# -*- coding: utf-8 -*-
"""
Sync all files to docs/ and web/
"""

import os
import shutil

repo = r"C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os"

files_to_sync = [
    "index.html",
    "pilot-roadmap.html",
    "offer.html",
    "privacy.html",
    "sample-audit-report.html",
    "sitemap.xml",
    "CNAME"
]

directories_to_sync = [
    "css",
    "samples"
]

for folder in ["docs", "web"]:
    target_dir = os.path.join(repo, folder)
    if os.path.exists(target_dir):
        # copy directories
        for d in directories_to_sync:
            src_d = os.path.join(repo, d)
            dst_d = os.path.join(target_dir, d)
            if os.path.exists(src_d):
                os.makedirs(dst_d, exist_ok=True)
                for f in os.listdir(src_d):
                    src_f = os.path.join(src_d, f)
                    if os.path.isfile(src_f):
                        shutil.copy2(src_f, os.path.join(dst_d, f))

        # copy files
        for f in files_to_sync:
            src_f = os.path.join(repo, f)
            if os.path.exists(src_f):
                shutil.copy2(src_f, os.path.join(target_dir, f))
        print(f"✓ Synchronized all files and folders to {folder}/")

print("Sync completed!")
