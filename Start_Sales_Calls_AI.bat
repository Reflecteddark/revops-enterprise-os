@echo off
title Sales Calls AI - Complete Starter
chcp 65001 >nul

echo ===================================================
echo [1/3] Проверка и запуск Ollama (DeepSeek-R1)...
echo ===================================================
netstat -ano | findstr /I "LISTENING" | findstr /C:":11434 " >nul
if not errorlevel 1 (
    echo Ollama уже работает на порту 11434.
) else (
    echo Запуск Ollama в фоновом режиме...
    start "" /b ollama serve >nul 2>&1
    timeout /t 2 /nobreak >nul
    echo Ollama запущена.
)

echo.
echo ===================================================
echo [2/3] Проверка и запуск Whisper Transcriber (Порт 8000)...
echo ===================================================
netstat -ano | findstr /I "LISTENING" | findstr /C:":8000 " >nul
if not errorlevel 1 (
    echo Транскрибатор уже работает на порту 8000.
) else (
    echo Запуск Faster-Whisper микросервиса...
    start "" /b "C:\Users\strel\.gemini\antigravity\scratch\game_vision_analytics\.venv\Scripts\python.exe" -m uvicorn server:app --app-dir "C:\Users\strel\.gemini\antigravity\scratch\transcriber_service" --host 127.0.0.1 --port 8000 >nul 2>&1
    timeout /t 3 /nobreak >nul
    echo Whisper запущен на порту 8000.
)

echo.
echo ===================================================
echo [3/3] Проверка и запуск n8n (Порт 5678)...
echo ===================================================
netstat -ano | findstr /I "LISTENING" | findstr /C:":5678 " >nul
if not errorlevel 1 (
    echo n8n уже работает на порту 5678.
    goto open_ui
)

echo Запуск службы n8n...
schtasks /Run /TN "JobHunter_N8N" >nul 2>&1

set attempts=0
:wait_n8n
timeout /t 2 /nobreak >nul
netstat -ano | findstr /I "LISTENING" | findstr /C:":5678 " >nul
if not errorlevel 1 (
    echo n8n успешно поднят на порту 5678!
    goto open_ui
)
set /a attempts+=1
if %attempts% geq 10 (
    echo [!] n8n запускается дольше обычного...
    goto open_ui
)
goto wait_n8n

:open_ui
echo.
echo ===================================================
echo  Все сервисы активны!
echo  Открытие воркфлоу в браузере...
echo ===================================================
if /i not "%1"=="--silent" (
    start http://localhost:5678/workflow/Sh4JkQtMKVdeRJn5
)

timeout /t 3 /nobreak >nul
exit
