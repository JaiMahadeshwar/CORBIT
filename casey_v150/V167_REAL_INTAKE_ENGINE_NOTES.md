# CASEY V167 — Real Intake Normalisation Engine

What changed:

- Replaced raw JSON upload output with a professional client-file challenge cockpit.
- Added backend intake normalisation for XLSX/XLSM, CSV/text and XER files.
- Upload now derives challenge numbers from the uploaded source file where possible:
  - client-file P50/headline signal
  - CASEY P80/P90 challenge range
  - board defensibility score
  - cost signals / risk rows / XER activity logic counts
- Added source intelligence objects for cost, risk and schedule parsing.
- Added real board attack kill-chain for uploaded files.
- Added professional UI metrics instead of code-looking JSON.
- Added explicit distinction between live project model and uploaded-file challenge model.

Important honest note:

This is a strong demo/prototype intake engine. It is not a full commercial-grade parser for every possible broken client file yet. The production version still needs a dedicated Python ingestion service with OCR/PDF extraction, robust schema learning, multi-file bundle reconciliation, WBS/CBS/activity graph matching and historical backtesting.

Demo path:

1. Start backend and frontend.
2. Open a generated project.
3. Go to Advisor.
4. Use one of the three challenge buttons or upload XLSX / CSV / XER.
5. The upload panel should show source-derived challenge metrics, findings, red flags, attack questions and next actions.
