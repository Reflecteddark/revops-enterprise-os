@echo off
chcp 65001 > nul
title RevOps Telegram Bot Manager

if exist "telegram_bot.py" (
    rem Already in project directory
) else if exist "%~dp0\telegram_bot.py" (
    cd /d "%~dp0"
) else if exist "%~dp0\..\telegram_bot.py" (
    cd /d "%~dp0\.."
) else if exist "C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\telegram_bot.py" (
    cd /d "C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os"
) else if exist "C:\Users\strel\.gemini\antigravity\scratch\telegram_bot.py" (
    cd /d "C:\Users\strel\.gemini\antigravity\scratch"
)

:menu
cls
echo ======================================================================
echo          REVOPS ENTERPRISE OS V17.6: TELEGRAM BOT MANAGER
echo ======================================================================
echo.
echo   1. [Шаг 1] Настройка токена и Chat ID (--setup)
echo   2. [Шаг 2] Диагностика и проверка связи (--test)
echo   3. [Шаг 3] Предпросмотр брифинга в консоли (--preview)
echo   4. [Шаг 4] Отправить утренний брифинг в Telegram (--send)
echo   5. [Шаг 4] Отправить свежий PDF-отчет в Telegram (--summary)
echo   6. [Шаг 5] Запустить интерактивный polling-бот (--polling)
echo   7. Выход
echo.
echo ======================================================================
set /p opt="Выберите действие [1-7]: "

if "%opt%"=="1" (
    echo.
    python telegram_bot.py --setup
    echo.
    pause
    goto menu
)
if "%opt%"=="2" (
    echo.
    python telegram_bot.py --test
    echo.
    pause
    goto menu
)
if "%opt%"=="3" (
    echo.
    python telegram_bot.py --preview
    echo.
    pause
    goto menu
)
if "%opt%"=="4" (
    echo.
    python telegram_bot.py --send
    echo.
    pause
    goto menu
)
if "%opt%"=="5" (
    echo.
    python telegram_bot.py --summary
    echo.
    pause
    goto menu
)
if "%opt%"=="6" (
    echo.
    echo Запуск polling-бота (нажмите Ctrl+C для остановки)...
    python telegram_bot.py --polling
    echo.
    pause
    goto menu
)
if "%opt%"=="7" exit /b 0

goto menu
