# CASEY V166 Pro Challenge Build

What changed:
- Rebuilt production `frontend/dist` so the visible advisor/client challenge screen is updated.
- Replaced raw JSON-style challenge results with a professional Client Challenge Room.
- Added three no-file demo challenge buttons: messy cost workbook, XER schedule, and risk register.
- Added professional intake verdict, red flags, traceability chain, and next-action cards.
- Added runtime/scenario controls already available under Scenario Intel and Causal OS.

How to demo:
1. Run backend and frontend as before.
2. Open the app.
3. Click `Advisor`.
4. Use the right-hand `Live Client Challenge Room` buttons.
5. Use left-hand board attack buttons for adversarial board Q&A.
6. Use `Scenario Intel` to click scenario comparison and runtime controls.

Note: the upload parser in this zip is still a front-end/demo normalisation layer. For production-grade messy XLSX/PDF/XER parsing, connect the Python intake parser/backend pipeline.
