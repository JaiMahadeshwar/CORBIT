# CASEY TITAN X v52 Elite Output System

This release keeps the current frontend and model flow, and upgrades the export layer.

## What changed

- V52 export routes patch the existing `/export/*` endpoints, so the frontend buttons keep working.
- Excel outputs are now white-background, board-readable and include:
  - Executive Dashboard
  - Selected Class Cost Model
  - Scenario Comparison
  - Risk Register
  - QCRA/QSRA Analytics
  - Basis + Audit Trail
- Risk Register export is now XLSX, not CSV, with:
  - Cause
  - Risk Event
  - Impact Description
  - Mitigation
  - Owner
  - O/M/P cost and schedule impacts
  - Activity and CBS links
- PDF is a board decision pack, not a dense dark report.
- PPTX uses a light consulting-style board deck structure.
- XER is exported with a legacy-targeted `ERMHDR 20.12` header to avoid Primavera Risk Analysis rejecting 24.12 exports. Validate in your exact Primavera/PRA environment.
- Full ZIP pack includes XLSX, risk workbook, XER, DOCX, PDF, PPTX, JSON audit model and a readme.

## Still important

For production, validate the XER import in your exact Primavera Risk Analysis/P6 version because Oracle import support differs by installed version and patch level.
