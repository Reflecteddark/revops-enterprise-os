@echo off
chcp 65001 > nul
echo ===================================================
echo   RevOps Enterprise OS V17.6 - Express AI Audit
echo   Запуск Экспресс-Аудита 3 звонков (Лид-магнит)
echo ===================================================
cd /d "%~dp0.."
python run_express_audit.py --demo
echo.
echo Аудит успешно завершен! Открываем отчет в браузере...
start "" "%~dp0..\docs\Экспресс_ИИ_Аудит_Звонков.html"
pause
exit /b 0
