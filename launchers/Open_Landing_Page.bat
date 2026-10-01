@echo off
chcp 65001 > nul
echo Opening RevOps AI Landing Page in default browser...
start "" "%~dp0..\presentation\landing.html"
echo Opened successfully.
exit /b 0
