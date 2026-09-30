@echo off
chcp 65001 > nul
title Push to GitHub: Reflecteddark/revops-enterprise-os

echo ======================================================================
echo       REVOPS ENTERPRISE OS V17.6: СИНХРОНИЗАЦИЯ И PUSH НА GITHUB
echo       Репозиторий: https://github.com/Reflecteddark/revops-enterprise-os
echo ======================================================================
echo.

set "SCRATCH_DIR=C:\Users\strel\.gemini\antigravity\scratch"
set "REPO_DIR=C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os"

echo [1/4] Проверка и синхронизация файлов из рабочей директории...
cd /d "%SCRATCH_DIR%"
python -c "
import os, shutil
src = r'%SCRATCH_DIR%'
dst = r'%REPO_DIR%'
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
    s = os.path.join(src, f)
    d = os.path.join(dst, f)
    if os.path.exists(s):
        if not os.path.exists(d) or os.path.getmtime(s) > os.path.getmtime(d):
            shutil.copy2(s, d)
            updated += 1
# Sync presentation folder
pres_src = os.path.join(src, 'presentation')
pres_dst = os.path.join(dst, 'presentation')
if os.path.exists(pres_src):
    shutil.copytree(pres_src, pres_dst, dirs_exist_ok=True)
print(f'Синхронизировано измененных файлов: {updated}')
"

cd /d "%REPO_DIR%"

echo.
echo [2/4] Проверка статуса Git...
git status -s

echo.
set "COMMIT_MSG="
set /p COMMIT_MSG="[3/4] Введите комментарий к коммиту (Enter для авто-сообщения): "
if "%COMMIT_MSG%"=="" (
    for /f "tokens=1-2 delims= " %%a in ('date /t') do set "CURR_DATE=%%a"
    for /f "tokens=1-2 delims= " %%a in ('time /t') do set "CURR_TIME=%%a"
    set "COMMIT_MSG=chore: update RevOps Enterprise OS %CURR_DATE% %CURR_TIME%"
)

echo.
echo Добавление изменений в индекс...
git add .

git diff-index --quiet HEAD --
if %ERRORLEVEL% EQU 0 (
    echo.
    echo [ИНФО] Локальных изменений для коммита не обнаружено.
) else (
    echo Создание коммита: "%COMMIT_MSG%"...
    git commit -m "%COMMIT_MSG%"
)

echo.
echo [4/4] Отправка изменений на GitHub (git push origin main)...
git push origin main

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ======================================================================
    echo    ✓ УСПЕХ! Проект успешно синхронизирован с GitHub:
    echo    https://github.com/Reflecteddark/revops-enterprise-os
    echo ======================================================================
) else (
    echo.
    echo ======================================================================
    echo    [ОШИБКА] Не удалось выполнить push. Проверьте подключение к сети.
    echo ======================================================================
)

echo.
pause
