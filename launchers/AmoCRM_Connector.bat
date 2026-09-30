@echo off
chcp 1251 > nul
title AmoCRM Connector Manager

set "PYTHON_EXE=python"
where python >nul 2>nul
if errorlevel 1 (
    if exist "C:\Python314\python.exe" set "PYTHON_EXE=C:\Python314\python.exe"
)

set "SCRIPT_DIR=C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os"
if exist "%SCRIPT_DIR%\amocrm_connector.py" (
    cd /d "%SCRIPT_DIR%"
) else if exist "%~dp0amocrm_connector.py" (
    cd /d "%~dp0"
) else if exist "%~dp0..\amocrm_connector.py" (
    cd /d "%~dp0.."
) else (
    echo [ОШИБКА] Не найден файл amocrm_connector.py!
    pause
    exit /b 1
)

:menu
cls
echo ======================================================================
echo          REVOPS ENTERPRISE OS V17.6: AMOCRM CONNECTOR
echo ======================================================================
echo.
echo   1. [Шаг 1] Первоначальная настройка токенов OAuth 2.0 (--setup)
echo   2. [Шаг 2] Посмотреть этапы воронки и их ID (--show-stages)
echo   3. [Шаг 3] Открыть файл маппинга этапов (amocrm_stage_mapping.json)
echo   4. [Шаг 4] Тестовый запуск без записи в Excel (--dry-run --days 30)
echo   5. [Шаг 5] Боевая синхронизация сделок в raw_deals (Excel)
echo   6. Выход
echo.
echo ======================================================================
set "opt="
set /p opt="Выберите действие [1-6]: "

if "%opt%"=="1" goto do_setup
if "%opt%"=="2" goto do_stages
if "%opt%"=="3" goto do_mapping
if "%opt%"=="4" goto do_dryrun
if "%opt%"=="5" goto do_sync
if "%opt%"=="6" goto do_exit
goto menu

:do_setup
echo.
"%PYTHON_EXE%" amocrm_connector.py --setup
echo.
pause
goto menu

:do_stages
echo.
"%PYTHON_EXE%" amocrm_connector.py --show-stages
echo.
pause
goto menu

:do_mapping
if not exist amocrm_stage_mapping.json (
    echo Файл маппинга еще не создан. Запустите сначала Шаг 2 (--show-stages).
) else (
    start notepad amocrm_stage_mapping.json
)
echo.
pause
goto menu

:do_dryrun
echo.
"%PYTHON_EXE%" amocrm_connector.py --dry-run --days 30
echo.
pause
goto menu

:do_sync
echo.
"%PYTHON_EXE%" amocrm_connector.py
echo.
pause
goto menu

:do_exit
exit /b 0
