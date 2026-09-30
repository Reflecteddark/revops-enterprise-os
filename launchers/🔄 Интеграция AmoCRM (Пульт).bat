@echo off
chcp 65001 > nul
title AmoCRM Connector Manager
cd /d "%~dp0\.."

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
set /p opt="Выберите действие [1-6]: "

if "%opt%"=="1" (
    echo.
    python amocrm_connector.py --setup
    echo.
    pause
    goto menu
)
if "%opt%"=="2" (
    echo.
    python amocrm_connector.py --show-stages
    echo.
    pause
    goto menu
)
if "%opt%"=="3" (
    if not exist amocrm_stage_mapping.json (
        echo Файл маппинга еще не создан. Запустите сначала Шаг 2 (--show-stages).
    ) else (
        start notepad amocrm_stage_mapping.json
    )
    echo.
    pause
    goto menu
)
if "%opt%"=="4" (
    echo.
    python amocrm_connector.py --dry-run --days 30
    echo.
    pause
    goto menu
)
if "%opt%"=="5" (
    echo.
    python amocrm_connector.py
    echo.
    pause
    goto menu
)
if "%opt%"=="6" exit /b 0

goto menu
