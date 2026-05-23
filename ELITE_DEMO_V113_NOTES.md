# CASEY V113 Elite Demo Upgrade

Added final board-level interpretation layer for Earth and Space demos:

- Confidence now reads as board-defensibility, not a vague percentage.
- Overview includes What Confidence Means, Likely Board Questions and CASEY Final Position.
- Scenario pages now include Gained / Sacrificed / Exposed language.
- Selected scenario tiles use black text on cyan for readability.
- Methodology explains confidence as evidence maturity, procurement certainty, schedule logic, contingency adequacy and scenario posture.
- Build checked with `npm run build` and backend checked with `python -m py_compile backend/main.py`.

Run commands:

Backend:
```bash
cd backend
python -m uvicorn main:app --reload
```

Frontend:
```bash
cd frontend
npm install
npm run dev -- --host 0.0.0.0
```
