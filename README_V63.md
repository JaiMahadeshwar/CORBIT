# CASEY TITAN X v63 True Universal Engine

This build fixes the export payload issue and installs the v63 universal output engine.

## What changed
- `/export/*` endpoints now accept safe JSON bodies and no longer return 422 for frontend export clicks.
- XER is dynamically generated from project type, Earth/Space mode, scenario and schedule level.
- No Mondrian / L01-10 content is copied into dynamic output.
- Industry-specific WBS templates added for data centres, airports, ports, rail, life sciences, fabs, energy and space programmes.
- L1-L5 schedules have different structure and activity counts.
- Scenario behaviour changes cost, risk, schedule durations and output narrative.
- Excel cost/risk outputs remain the primary source-of-truth pack.

## Run
Backend:
```powershell
cd backend
python main.py
```
Frontend:
```powershell
cd frontend
npm install
npm run dev
```
