# CASEY TITAN X v54 — Platinum Output Engine

This build keeps the existing full product flow and upgrades the export engine.

## What changed
- White-background, readable Excel cost model with executive dashboard
- Risk register XLSX with cause, event, impact, owner, mitigation, trigger, residual rating
- Separate QCRA curve and QSRA curve sheets
- Separate QCRA tornado and QSRA tornado sheets
- Scenario-linked recommendation and trade-off narrative
- Board PDF rebuilt for readability
- PPTX rebuilt as a cleaner board-decision deck
- DOCX rebuilt with decision snapshot, risk drivers and selected-class estimate
- XER now uses a legacy-targeted P6 style header and includes schedule logic
- Full ZIP keeps model audit and raw risk CSV

## Note on Primavera Risk Analysis
PRA import support varies by installed PRA/P6 version. v54 exports a legacy-targeted XER and also keeps a structured schedule table inside the Excel workbook. If PRA rejects XER, open/import the schedule in P6 Professional and re-export in the exact PRA-supported version.
