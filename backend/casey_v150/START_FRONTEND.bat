@echo off
cd /d "%~dp0frontend"
npm install
npm run dev -- --port 5173
pause
