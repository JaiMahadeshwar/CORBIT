# CASEY V140 Real Final Demo Drop

This package was created from the uploaded V134 source tree and re-zipped as a real artifact.

## What is included
- V134 incumbent-pressure / assurance build preserved.
- Scenario rail, scenario propagation UI, assurance tab, board-pressure panels and export strip retained.
- Cost split reconciliation logic retained in frontend normalization.
- Production frontend build generated in `frontend/dist`.

## Verification performed in this environment
- `npm ci` completed in `frontend`.
- `npm run build` completed successfully in `frontend`.
- Python syntax compile passed for:
  - `backend/main.py`
  - `backend/v56_outputs.py`
  - `backend/v59_outputs.py`
  - `backend/v61_outputs.py`
  - `backend/v62_outputs.py`
  - `backend/v63_outputs.py`
  - `backend/v64_outputs.py`

## Demo instruction
Use this as the demo drop only after running one local smoke test:
1. Start backend from `backend`.
2. Start frontend from `frontend`.
3. Run Space demo.
4. Switch Base / Faster / Cheaper / Lower Risk / Premium.
5. Check Cost, Schedule, Risk, QCRA/QSRA, Assurance and exports.

## Honest note
This zip is a real packaged artifact from the files available in this conversation. It is not a claimed 10,000-project validation build.
