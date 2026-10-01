@echo off
chcp 65001 > nul
echo =======================================================================
echo   RevOps Enterprise OS: Генерация 152-ФЗ Комплекта Безопасности
echo =======================================================================
echo.

cd /d "%~dp0\.."

python scripts\generate_152fz_compliance_docx.py

if errorlevel 1 (
    echo.
    echo [ERROR] Ошибка генерации документа!
    pause
    exit /b 1
)

echo.
echo [OK] Документ успешно обновлен на Рабочем столе и в docs!
echo.
pause
