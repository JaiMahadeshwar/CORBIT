# CASEY TITAN X v37 Full Check Report

Checked package areas:

- Frontend install/build: passed with `npm install` and `npm run build`.
- Frontend route: Vite serves at `http://localhost:5173/`.
- Backend import/compile: passed with `python -m py_compile backend/main.py`.
- Backend API: `/health`, `/demo/status`, `/demo/showcases`, `/projects`, `/generate` tested.
- Demo examples: Earth and Space demo prompts call `/generate` and return full project models.
- Output exports: Excel, CSV risk register, P6 XER, Word report, PDF report, PowerPoint deck, JSON model, and full ZIP all return valid files.
- Background: local CSS cinematic Atlas-style scene, not a missing YouTube/background video dependency.
- Music: copyright-safe in-browser procedural soundtrack with Atlas selected by default. No copyrighted audio files included.
- Windows launch scripts: included for one-click frontend/backend start.

Important run commands:

```powershell
START_CASEY.bat
```

Manual backend:

```powershell
cd backend
pip install -r requirements.txt
python -m uvicorn main:app --reload --port 8000
```

Manual frontend:

```powershell
cd frontend
npm install
npm run dev
```

Open:

```text
http://localhost:5173
```

Note: music starts only after a user click because browsers block autoplay audio. Click Launch Mission Control or Music On.
