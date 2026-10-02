@echo off
title AI & Sales Services - Complete Stopper
chcp 65001 >nul

echo ===================================================
echo [1/3] Остановка n8n (Порт 5678)...
echo ===================================================
schtasks /End /TN "JobHunter_N8N" >nul 2>&1
for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":5678" ^| findstr "LISTENING"') do (
    taskkill /F /PID %%P >nul 2>&1
)
taskkill /F /IM node.exe >nul 2>&1
echo n8n остановлен, порт 5678 свободен.

echo.
echo ===================================================
echo [2/3] Остановка нейросети Ollama (Порт 11434)...
echo ===================================================
taskkill /F /IM ollama.exe >nul 2>&1
taskkill /F /IM "ollama app.exe" >nul 2>&1
echo Ollama выгружена из памяти (видеопамять освобождена).

echo.
echo ===================================================
echo [3/3] Остановка Whisper Transcriber (Порт 8000)...
echo ===================================================
for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":8000" ^| findstr "LISTENING"') do (
    taskkill /F /PID %%P >nul 2>&1
)
echo Транскрибатор остановлен, порт 8000 свободен.

echo.
echo ===================================================
echo === ИТОГОВЫЙ СТАТУС ===
echo ===================================================
netstat -ano | findstr /C:":5678 " | findstr "LISTENING" >nul && (echo [!] n8n :5678 еще занят) || (echo [OK] n8n :5678 выключен)
netstat -ano | findstr /C:":11434 " | findstr "LISTENING" >nul && (echo [!] Ollama :11434 еще занята) || (echo [OK] Ollama :11434 выключена)
netstat -ano | findstr /C:":8000 " | findstr "LISTENING" >nul && (echo [!] Whisper :8000 еще занят) || (echo [OK] Whisper :8000 выключен)
echo ===================================================
echo Все ресурсы (CPU / RAM / VRAM) полностью освобождены.
echo.
timeout /t 3 /nobreak >nul
exit
