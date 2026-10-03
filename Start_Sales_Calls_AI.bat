@echo off
title Sales Calls AI - Complete Starter

echo ===================================================
echo [1/3] Check and start Ollama (DeepSeek-R1)...
echo ===================================================
set "OLLAMA_STATUS=000"
for /f "tokens=*" %%a in ('curl -s -m 2 -o nul -w "%%{http_code}" http://127.0.0.1:11434/api/tags 2^>nul') do set "OLLAMA_STATUS=%%a"

if "%OLLAMA_STATUS%"=="200" (
    echo Ollama is already running on port 11434.
    goto check_whisper
)

echo Starting Ollama in background...
start "" /b ollama serve >nul 2>&1
ping -n 3 127.0.0.1 >nul
echo Ollama started.

:check_whisper
echo.
echo ===================================================
echo [2/3] Check and start Whisper Transcriber (Port 8000)...
echo ===================================================
set "WHISPER_STATUS=000"
for /f "tokens=*" %%a in ('curl -s -m 2 -o nul -w "%%{http_code}" http://127.0.0.1:8000/docs 2^>nul') do set "WHISPER_STATUS=%%a"

if "%WHISPER_STATUS%"=="200" (
    echo Transcriber is already running on port 8000.
    goto check_n8n
)

echo Starting Faster-Whisper microservice...
start "" /b "C:\Users\strel\.gemini\antigravity\scratch\game_vision_analytics\.venv\Scripts\python.exe" -m uvicorn server:app --app-dir "C:\Users\strel\.gemini\antigravity\scratch\transcriber_service" --host 127.0.0.1 --port 8000 >nul 2>&1
ping -n 4 127.0.0.1 >nul
echo Whisper started on port 8000.

:check_n8n
echo.
echo ===================================================
echo [3/3] Check and start n8n (Port 5678)...
echo ===================================================
set "N8N_STATUS=000"
for /f "tokens=*" %%a in ('curl -s -m 2 -o nul -w "%%{http_code}" http://localhost:5678/healthz 2^>nul') do set "N8N_STATUS=%%a"

if "%N8N_STATUS%"=="200" (
    echo n8n is already running on port 5678.
    goto open_ui
)

echo Starting n8n service...
schtasks /Run /TN "JobHunter_N8N" >nul 2>&1
start "" "C:\Python314\pythonw.exe" "C:\Users\strel\.gemini\antigravity\scratch\run_n8n_silent.py" >nul 2>&1

set attempts=0
:wait_n8n
ping -n 3 127.0.0.1 >nul
set "N8N_STATUS=000"
for /f "tokens=*" %%a in ('curl -s -m 2 -o nul -w "%%{http_code}" http://localhost:5678/healthz 2^>nul') do set "N8N_STATUS=%%a"

if "%N8N_STATUS%"=="200" (
    echo n8n is up on port 5678!
    goto open_ui
)

set /a attempts+=1
if %attempts% geq 12 (
    echo [!] n8n is starting slower than usual, opening UI...
    goto open_ui
)
goto wait_n8n

:open_ui
echo.
echo ===================================================
echo  All services are active!
echo  Opening workflow in browser...
echo ===================================================
if /i not "%1"=="--silent" (
    start http://localhost:5678/workflow/Sh4JkQtMKVdeRJn5
)

ping -n 4 127.0.0.1 >nul
exit
