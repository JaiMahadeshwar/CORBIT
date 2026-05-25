# CASEY V102 Clean Rebuilt Validated

## Start backend
Open PowerShell:

```powershell
cd C:\Users\jaima\Downloads\CASEY_V102_CLEAN_REBUILT_VALIDATED\backend
python -m uvicorn main:app --reload
```

## Start frontend
Open a second PowerShell:

```powershell
cd C:\Users\jaima\Downloads\CASEY_V102_CLEAN_REBUILT_VALIDATED\frontend
npm config set registry https://registry.npmjs.org/
npm install
npm run dev
```

Then open the Vite localhost URL.

## Notes
- The previous bad zip had the project inside a `casey_v96_upgrade` folder. This one has the correct top-level folder.
- The internal OpenAI npm registry references were removed by deleting the old lockfile and adding a clean `.npmrc`.
- Demo repeat-run blocking is disabled in the backend for launch testing.
