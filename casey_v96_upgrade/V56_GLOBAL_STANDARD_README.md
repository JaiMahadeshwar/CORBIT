# CASEY TITAN X v56 Global Standard Output Engine

This build focuses on the outputs only, exactly as requested:

- PDF, PPTX and weak Word outputs are removed from the primary source-of-truth pack.
- Excel is generated with XlsxWriter to avoid Microsoft Excel repair/corruption warnings.
- Every cost workbook includes selected scenario, scenario comparison, selected class estimate, all class levels, full risk register, QCRA curve, QSRA curve, QCRA tornado, QSRA tornado, all schedule levels and audit notes.
- Risk register is an XLSX workbook with no zero likelihood values and full cause/event/impact/owner/mitigation/trigger fields.
- Full output pack contains: cost workbook, risk register workbook, PRA-targeted XER, schedule CSV fallback and JSON audit.
- XER uses an ERMHDR 8.0-style legacy target. If local Primavera Risk Analysis still rejects it due to environment/version rules, use the schedule CSV fallback or re-export through P6 Professional.

## Run
Backend:
```powershell
cd backend
pip install -r requirements.txt
python main.py
```
Frontend:
```powershell
cd frontend
npm install
npm run dev
```
