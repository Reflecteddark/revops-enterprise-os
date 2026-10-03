@echo off
title AI & Sales Services - Complete Stopper

echo ===================================================
echo [1/3] Stopping n8n (Port 5678)...
echo ===================================================
schtasks /End /TN "JobHunter_N8N" >nul 2>&1
for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":5678" ^| findstr "LISTENING"') do (
    taskkill /F /PID %%P >nul 2>&1
)
taskkill /F /IM node.exe >nul 2>&1
echo n8n stopped, port 5678 is free.

echo.
echo ===================================================
echo [2/3] Stopping Ollama AI (Port 11434)...
echo ===================================================
taskkill /F /IM ollama.exe >nul 2>&1
taskkill /F /IM "ollama app.exe" >nul 2>&1
echo Ollama unloaded from memory (VRAM released).

echo.
echo ===================================================
echo [3/3] Stopping Whisper Transcriber (Port 8000)...
echo ===================================================
for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":8000" ^| findstr "LISTENING"') do (
    taskkill /F /PID %%P >nul 2>&1
)
echo Transcriber stopped, port 8000 is free.

echo.
echo ===================================================
echo === SUMMARY STATUS ===
echo ===================================================
netstat -ano | findstr ":5678" | findstr "LISTENING" >nul && (echo [!] n8n :5678 is STILL BUSY) || (echo [OK] n8n :5678 is STOPPED)
netstat -ano | findstr ":11434" | findstr "LISTENING" >nul && (echo [!] Ollama :11434 is STILL BUSY) || (echo [OK] Ollama :11434 is STOPPED)
netstat -ano | findstr ":8000" | findstr "LISTENING" >nul && (echo [!] Whisper :8000 is STILL BUSY) || (echo [OK] Whisper :8000 is STOPPED)
echo ===================================================
echo All CPU, RAM and GPU resources released.
echo.
ping -n 4 127.0.0.1 >nul
exit
