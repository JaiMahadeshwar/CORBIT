# CASEY V106 Final Exports + Demo QA

Validated fixes included in this package:

- All five export buttons now download files instead of showing demo alert popups.
- Export Board Pack downloads the full ZIP output pack.
- Export Cost Workbook downloads a real workbook.
- Export Risk Register downloads a real risk-register workbook.
- Export XER downloads a real XER schedule file.
- Export QCRA/QSRA downloads a real QCRA/QSRA workbook with curves and tornado drivers.
- Added `/export/excel` alias so older frontend calls do not fail.
- Added selected-scenario CSS override so active Base/Faster/Cheaper/Lower Risk/Premium cards render dark text on cyan.
- Backend compiles successfully.
- API smoke test passed for generate plus `/export/workbook`, `/export/risk-register`, `/export/xer`, `/export/qcra-qsra`, `/export/all`, `/export/excel`.

Demo note: public exports are demo-stamped / demo-named and should be described as first-pass strategic intelligence, not certified estimates.
