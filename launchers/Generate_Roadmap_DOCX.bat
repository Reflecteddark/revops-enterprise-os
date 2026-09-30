@echo off
chcp 65001 > nul
echo ===================================================
echo   RevOps AI-Супервайзер: Генерация Дорожной Карты
echo ===================================================
echo.

cd /d "%~dp0\.."

python scripts\generate_roadmap_docx.py

if errorlevel 1 (
    echo.
    echo [ERROR] Ошибка генерации документа!
    pause
    exit /b 1
)

echo.
echo [OK] Документ Word успешно обновлен в docs\ и скопирован на Рабочий стол!
echo.
pause
