@echo off
chcp 1251 > nul
title RevOps All Reports
echo =========================================================
echo   RevOps Enterprise OS V17.6: Генерация Всех Отчетов
echo =========================================================
echo.
cd /d "C:\Users\strel\.gemini\antigravity\scratch"
python generate_all.py
if errorlevel 1 goto error

echo.
echo =========================================================
echo   УСПЕШНО! Все отчеты лежат на вашем Рабочем столе:
echo   1. %USERPROFILE%\Desktop\RevOps_Enterprise_OS_V17.6_Executive_Summary.pdf
echo   2. %USERPROFILE%\Desktop\RevOps_Enterprise_OS_V17.6_Full_Report.pdf
echo =========================================================
echo.
echo Открытие Executive Summary...
start "" "%USERPROFILE%\Desktop\RevOps_Enterprise_OS_V17.6_Executive_Summary.pdf"
goto end

:error
echo.
echo [ОШИБКА] Произошел сбой при генерации отчетов!
pause

:end
ping -n 4 127.0.0.1 > nul
exit /b 0
