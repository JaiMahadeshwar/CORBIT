@echo off
title CASEY TITAN X v39 EXECUTIVE CINEMATIC
cd /d "%~dp0"
echo Starting CASEY backend on http://127.0.0.1:8000 ...
start "CASEY Backend" cmd /k "cd /d %~dp0backend && python -m uvicorn main:app --reload --port 8000"
timeout /t 3 >nul
echo Starting CASEY frontend...
start "CASEY Frontend" cmd /k "cd /d %~dp0frontend && npm install && npm run dev -- --port 5173"
timeout /t 6 >nul
start http://localhost:5173
echo.
echo CASEY is starting. Keep the two terminal windows open.
echo If port 5173 is busy, use the URL printed in the frontend window.
pause
