@echo off
chcp 65001 > nul
echo =======================================================================
echo   RevOps Enterprise OS: Генерация Единой Мастер-Дорожной Карты
echo =======================================================================
echo.

cd /d "%~dp0\.."

python scripts\generate_unified_roadmap_docx.py

if errorlevel 1 (
    echo.
    echo [ERROR] Ошибка генерации документа!
    pause
    exit /b 1
)

echo.
echo [OK] Единый Мастер-документ успешно обновлен на Рабочем столе и в docs!
echo.
pause
