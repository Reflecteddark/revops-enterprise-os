@echo off
chcp 65001 > nul
title AI-ROP MAX Messenger Bot (@se14526668_bot)
echo ========================================================
echo   AI-ROP MAX Bot Service
echo   Bot: @se14526668_bot (https://max.ru/se14526668_bot)
echo   Mini App: https://ai-rop.ru
echo ========================================================
echo.
cd /d "%~dp0"
python services\max_bot.py
pause
