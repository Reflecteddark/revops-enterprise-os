@echo off
chcp 1251 > nul
title RevOps Telegram Bot Manager

set "PYTHON_EXE=python"
where python >nul 2>nul
if errorlevel 1 (
    if exist "C:\Python314\python.exe" set "PYTHON_EXE=C:\Python314\python.exe"
)

set "SCRIPT_DIR=C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os"
if exist "%SCRIPT_DIR%\telegram_bot.py" (
    cd /d "%SCRIPT_DIR%"
) else if exist "%~dp0telegram_bot.py" (
    cd /d "%~dp0"
) else if exist "%~dp0..\telegram_bot.py" (
    cd /d "%~dp0.."
) else (
    echo [ОШИБКА] Не найден файл telegram_bot.py!
    pause
    exit /b 1
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
set "opt="
set /p opt="Выберите действие [1-7]: "

if "%opt%"=="1" goto do_setup
if "%opt%"=="2" goto do_test
if "%opt%"=="3" goto do_preview
if "%opt%"=="4" goto do_send
if "%opt%"=="5" goto do_summary
if "%opt%"=="6" goto do_polling
if "%opt%"=="7" goto do_exit
goto menu

:do_setup
echo.
"%PYTHON_EXE%" telegram_bot.py --setup
echo.
pause
goto menu

:do_test
echo.
"%PYTHON_EXE%" telegram_bot.py --test
echo.
pause
goto menu

:do_preview
echo.
"%PYTHON_EXE%" telegram_bot.py --preview
echo.
pause
goto menu

:do_send
echo.
"%PYTHON_EXE%" telegram_bot.py --send
echo.
pause
goto menu

:do_summary
echo.
"%PYTHON_EXE%" telegram_bot.py --summary
echo.
pause
goto menu

:do_polling
echo.
echo Запуск polling-бота (нажмите Ctrl+C для остановки)...
"%PYTHON_EXE%" telegram_bot.py --polling
echo.
pause
goto menu

:do_exit
exit /b 0
