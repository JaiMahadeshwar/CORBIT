# Windows Install

1. Extract the ZIP.
2. Open the extracted `CASEY_TITAN_X_v30_CLEAN_PRO` folder.
3. Double-click `START_CASEY.bat`.
4. Open the URL shown by frontend, usually `http://localhost:5173`.

Manual commands:

Backend:
```powershell
cd backend
python -m uvicorn main:app --reload --port 8000
```

Frontend:
```powershell
cd frontend
npm install
npm run dev
```
