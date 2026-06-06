@echo off
echo Stopping any running node processes...
taskkill /f /im node.exe 2>nul

echo Clearing Vite cache...
cd /d %~dp0frontend
rmdir /s /q node_modules\.vite 2>nul
rmdir /s /q .vite_nocache 2>nul

echo Starting fresh dev server...
npm run dev
