"""
Скрипт синхронизации рабочих файлов RevOps Enterprise OS из рабочей папки в Git-репозиторий.
"""
import os
import shutil
import sys

sys.stdout.reconfigure(encoding='utf-8')

scratch_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
repo_dir = os.path.dirname(__file__)

# Если скрипт запущен из самой папки scratch
if os.path.basename(repo_dir) != "revops-enterprise-os":
    scratch_dir = repo_dir
    repo_dir = os.path.join(scratch_dir, "revops-enterprise-os")

if not os.path.exists(scratch_dir) or not os.path.exists(repo_dir):
    print("[SYNC] Папки не найдены, пропускаем синхронизацию.")
    sys.exit(0)

core_files = [
    'generate_full_report_pdf.py',
    'data_validator.py',
    'audit_log.py',
    'generate_all.py',
    'generate_presentation_html.py',
    'generate_all.bat',
    'RevOps Platform V17.6 (RBAC Production Suite).xlsx',
    'revops_rbac_v176_hardened.js',
    'sync_enterprise_system.py',
    'sync_local_xlsx_fixes.py',
    'apply_full_audit_fixes_v176.py',
    'RevOps_Enterprise_Executive_Summary.html',
    'RevOps_Enterprise_Full_Report.html'
]

updated = 0
for f in core_files:
    s = os.path.join(scratch_dir, f)
    d = os.path.join(repo_dir, f)
    if os.path.exists(s) and s != d:
        if not os.path.exists(d) or os.path.getmtime(s) > os.path.getmtime(d):
            shutil.copy2(s, d)
            updated += 1

pres_src = os.path.join(scratch_dir, 'presentation')
pres_dst = os.path.join(repo_dir, 'presentation')
if os.path.exists(pres_src) and pres_src != pres_dst:
    shutil.copytree(pres_src, pres_dst, dirs_exist_ok=True)

print(f"[SYNC] Синхронизировано измененных файлов: {updated}")
