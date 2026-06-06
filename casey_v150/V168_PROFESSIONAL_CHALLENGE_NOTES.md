# CASEY V168 Professional Client-Side Challenge

What changed:
- Upload challenge no longer renders raw JSON.
- Uploaded files now appear as a professional assurance review: baseline remains separate from challenge delta.
- Removed aggressive wording such as kill-chain / red flags / EPC blame language from the visible challenge flow.
- Added board-grade sections: programme baseline, challenge exposure, benchmark comparison, source-file findings, commercial observations, board assurance questions and required next actions.
- Stress-test mutation now parses the existing P50/schedule correctly, so shocks do not collapse a rail megaprogramme into a $1B / 3-month project.
- Upload endpoint now returns V168 professionalised intake output while preserving V167 parser compatibility.

How to see it:
1. Run START_CASEY.bat or run backend + frontend separately.
2. Open the Advisor tab.
3. Use the three challenge buttons or upload an XLSX/XER/risk workbook.
4. The result should read like an independent consultant review, not raw code.

Validation run in this package:
- Frontend production build passed.
- Backend Python compile passed.
- quick_check.py passed across Earth / Space scenario samples.
- export_check.py generated workbook, risk workbook, XER, DOCX, PDF and PPTX bytes successfully.

Limit:
- This is still a demo-grade parser/normaliser. Real universal messy-file ingestion for every client format needs a dedicated backend pipeline with deeper sheet/schema/entity inference.
