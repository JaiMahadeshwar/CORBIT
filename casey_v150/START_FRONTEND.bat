@echo off
title CASEY Frontend

echo === CLEARING ALL VITE CACHES ===
cd /d "%~dp0frontend"

echo Killing node...
taskkill /f /im node.exe 2>nul
timeout /t 2 /nobreak >nul

echo Deleting ALL vite cache folders...
for /d %%i in (node_modules\.vite*) do rmdir /s /q "%%i"

echo Starting Vite dev server fresh...
npm run dev -- --port 5173
pause
