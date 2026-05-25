# CASEY V148 Final Trust Core QA Report

## What changed in V148
- Added canonical selected-scenario state (`casey_state`) to every generated model.
- Added deterministic scenario signature (`scenario_signature`) used by UI payloads and export manifests.
- Added full export manifest inside `/export/all` and `/export/full-pack` as `00_CASEY_V148_EXPORT_MANIFEST.json`.
- Re-registered export endpoints so workbook, risk register, XER, DOCX, PDF, PPTX, JSON audit, QCRA/QSRA and full ZIP all re-canonicalize the selected scenario before file generation.
- Added `/qa/validate-model` and `/qa/readiness` backend endpoints.
- Preserved clickable global showcase library and Earth/Space project archetypes.

## Actual tests completed in this build session
- Backend Python compile: PASS.
- Frontend dependency install and production build: PASS.
- Manual model consistency checks across rail, AI data centre, lunar and Starship-style space prompts for all 5 scenarios: PASS.
- Export endpoint smoke checks on a California High-Speed Rail Faster payload:
  - `/export/workbook`: PASS
  - `/export/risk-register`: PASS
  - `/export/xer`: PASS
  - `/export/word`: PASS
  - `/export/pdf`: PASS
  - `/export/pptx`: PASS
  - `/export/json`: PASS
  - `/export/qcra-qsra`: PASS
  - `/export/all`: PASS, ZIP includes V148 manifest and all expected files.

## Important honesty note
A complete rendered export pass for 10,000 Earth + 5,000 Space examples was not completed inside this chat run window. The V148 package includes stronger validation and export-locking so you or Claude can run destructive QA locally/server-side. Do not represent the model as certified or production-validated until a full external QA run is complete.

## Recommended next QA command
Run automated destructive QA on a dedicated machine/server, not inside the live demo session. Start with 500–1,000 models, then scale to 15,000+.
