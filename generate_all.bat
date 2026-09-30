@echo off
chcp 65001 > nul
echo ===================================================
echo   RevOps Enterprise OS V17.6 - Генератор PDF Отчетов
echo ===================================================
echo.

cd /d "%~dp0"

echo [1/3] Обновление векторных SVG-графиков...
python presentation\generate_charts.py
if %ERRORLEVEL% NEQ 0 (
    echo Ошибка при генерации графиков!
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [2/3] Сборка Executive Summary PDF (10 слайдов C-Level)...
python presentation\build.py
if %ERRORLEVEL% NEQ 0 (
    echo Ошибка при сборке Executive Summary!
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [3/3] Сборка Full Report PDF (15 аналитических витрин)...
python generate_full_report_pdf.py
if %ERRORLEVEL% NEQ 0 (
    echo Ошибка при сборке Full Report!
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo ===================================================
echo   УСПЕШНО! Все PDF-отчеты сформированы:
echo   1. RevOps_Enterprise_OS_V17.6_Executive_Summary.pdf
echo   2. RevOps_Enterprise_OS_V17.6_Full_Report.pdf
echo ===================================================
echo.
pause
