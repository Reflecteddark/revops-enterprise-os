@echo off
chcp 65001 > nul
title Bitrix24 Connector Manager
cd /d "%~dp0\.."

:menu
cls
echo ======================================================================
echo          REVOPS ENTERPRISE OS V17.6: BITRIX24 CONNECTOR
echo ======================================================================
echo.
echo   1. [Шаг 1] Настройка входящего вебхука REST API (--setup)
echo   2. [Шаг 2] Посмотреть этапы воронки и их ID (--show-stages)
echo   3. [Шаг 3] Открыть файл маппинга этапов (bitrix24_stage_mapping.json)
echo   4. [Шаг 4] Тестовый запуск без записи в Excel (--dry-run --days 30)
echo   5. [Шаг 5] Боевая синхронизация сделок в raw_deals (Excel)
echo   6. Выход
echo.
echo ======================================================================
set /p opt="Выберите действие [1-6]: "

if "%opt%"=="1" (
    echo.
    python bitrix24_connector.py --setup
    echo.
    pause
    goto menu
)
if "%opt%"=="2" (
    echo.
    python bitrix24_connector.py --show-stages
    echo.
    pause
    goto menu
)
if "%opt%"=="3" (
    if not exist bitrix24_stage_mapping.json (
        echo Файл маппинга еще не создан. Запустите сначала Шаг 2 (--show-stages).
    ) else (
        start notepad bitrix24_stage_mapping.json
    )
    echo.
    pause
    goto menu
)
if "%opt%"=="4" (
    echo.
    python bitrix24_connector.py --dry-run --days 30
    echo.
    pause
    goto menu
)
if "%opt%"=="5" (
    echo.
    python bitrix24_connector.py
    echo.
    pause
    goto menu
)
if "%opt%"=="6" exit /b 0

goto menu
