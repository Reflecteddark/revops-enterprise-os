@echo off
chcp 65001 > nul
echo Opening RevOps AI Landing Page in default browser...

if exist "%~dp0..\presentation\landing.html" (
    start "" "%~dp0..\presentation\landing.html"
) else if exist "C:\Users\strel\Desktop\RevOps_AI_Landing_Page.html" (
    start "" "C:\Users\strel\Desktop\RevOps_AI_Landing_Page.html"
) else if exist "C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\presentation\landing.html" (
    start "" "C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\presentation\landing.html"
) else (
    echo Error: Landing page HTML file not found!
    pause
)
exit /b 0
