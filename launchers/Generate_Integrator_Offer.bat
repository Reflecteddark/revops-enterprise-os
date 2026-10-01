@echo off
chcp 65001 > nul
echo =======================================================================
echo   RevOps Enterprise OS: Партнерский Оффер для Интеграторов CRM
echo =======================================================================
echo.

cd /d "%~dp0\.."

python scripts\generate_integrator_offer_docx.py

if errorlevel 1 (
    echo.
    echo [ERROR] Ошибка генерации документа!
    pause
    exit /b 1
)

echo.
echo [OK] Партнерский оффер успешно обновлен на Рабочем столе и в docs!
echo.
pause
