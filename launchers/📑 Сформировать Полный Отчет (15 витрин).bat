@echo off
chcp 1251 > nul
title RevOps Full Report
echo =========================================================
echo   RevOps Enterprise OS V17.6: Генерация Полного Отчета
echo =========================================================
echo.
echo [1/2] Экспорт 15 витрин и запуск Chrome Headless...
cd /d "C:\Users\strel\.gemini\antigravity\scratch"
python generate_full_report_pdf.py
if errorlevel 1 goto error

echo.
echo [2/2] Отчет успешно сохранен на Рабочий стол:
echo       %USERPROFILE%\Desktop\RevOps_Enterprise_OS_V17.6_Full_Report.pdf
echo.
echo Открытие готового PDF...
start "" "%USERPROFILE%\Desktop\RevOps_Enterprise_OS_V17.6_Full_Report.pdf"
goto end

:error
echo.
echo [ОШИБКА] Произошел сбой при генерации отчета!
pause

:end
ping -n 3 127.0.0.1 > nul
exit /b 0
