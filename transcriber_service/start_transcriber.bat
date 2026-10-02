@echo off
title Voice Transcriber (Port 8000)
cd /d "%~dp0"
"C:\Users\strel\.gemini\antigravity\scratch\game_vision_analytics\.venv\Scripts\python.exe" -m uvicorn server:app --host 127.0.0.1 --port 8000
pause
