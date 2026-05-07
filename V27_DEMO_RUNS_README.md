# CASEY TITAN v27 — Demo Runs Edition

This build fixes the front-screen demo so it actually runs a cinematic sequence.

## What changed
- New CASEY logo treatment.
- Clearer front-page description of what the tool does.
- Animated Earth, Moon and Mars scene.
- `Run cinematic demo` now opens a full-screen timed demo theatre.
- Demo theatre has pause/next controls and final buttons to run Earth or Space model.
- Earth demo generates a 2027 Riyadh AI hyperscale campus.
- Space demo generates a lunar/space frontier model.
- Added loading overlay so the user sees the app building the model.
- Kept the existing backend endpoints and exports.

## Run
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

Open the URL shown by Vite.
