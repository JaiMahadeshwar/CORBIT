@echo off
title CASEY V166 FINAL

echo ============================================
echo  CASEY V166 - FRESH START
echo ============================================
echo.

echo [1/4] Killing any running node processes...
taskkill /f /im node.exe 2>nul
timeout /t 2 /nobreak >nul

echo [2/4] Clearing ALL Vite caches in 897 folder...
cd /d "%~dp0frontend"
for /d %%i in (node_modules\.vite*) do (
    echo   Removing %%i
    rmdir /s /q "%%i"
)

echo [3/4] Starting Backend (port 8000)...
start "CASEY Backend" cmd /k "cd /d %~dp0backend && python -m uvicorn main:app --reload --port 8000"
timeout /t 3 >nul

echo [4/4] Starting Frontend (port 5173)...
start "CASEY Frontend" cmd /k "cd /d %~dp0frontend && npm run dev -- --port 5173"
timeout /t 8 >nul

start http://localhost:5173

echo.
echo ============================================
echo  CASEY IS RUNNING
echo  Frontend: http://localhost:5173
echo  Backend:  http://localhost:8000
echo ============================================
echo.
echo  IMPORTANT: When browser opens, press Ctrl+Shift+R
echo  to hard refresh and clear browser cache
echo.
pause
