@echo off
echo === CASEY FRESH START ===
echo.
echo Step 1: Stopping any running node processes...
taskkill /f /im node.exe 2>nul
timeout /t 2 /nobreak >nul

echo Step 2: Clearing ALL Vite caches...
if exist "node_modules\.vite" rmdir /s /q "node_modules\.vite"
if exist "node_modules\.vite_fresh" rmdir /s /q "node_modules\.vite_fresh"
if exist ".vite_nocache" rmdir /s /q ".vite_nocache"

echo Step 3: Installing dependencies if needed...
if not exist "node_modules\vite" npm install

echo Step 4: Starting fresh dev server...
echo.
echo When you see "Local: http://localhost:5173" - open that URL
echo and press Ctrl+Shift+R to hard refresh the browser
echo.
npm run dev
