@echo off
chcp 65001 > nul
title RevOps: Генерация всех отчетов (Master + CEO + РОП + CFO)

cd /d "%~dp0.."

echo ======================================================================
echo    REVOPS ENTERPRISE OS V17.6: ПАКЕТНАЯ ГЕНЕРАЦИЯ ВСЕХ ОТЧЕТОВ
echo ======================================================================
echo.
python generate_all.py --role all-roles
if errorlevel 1 (
    echo.
    echo [ОШИБКА] Произошел сбой при генерации отчетов.
    pause
    exit /b %errorlevel%
)

echo.
pause
