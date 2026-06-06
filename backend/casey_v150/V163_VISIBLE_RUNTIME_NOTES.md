# CASEY V163 Visible Runtime + Professional Intake Fix

What changed:
- Upload challenge no longer renders raw JSON to the user.
- Uploaded files now show a professional Intake Normalization Engine card with findings, red flags, lineage and next actions.
- Added a visible **Runtime OS** tab beside Advisor/Methodology.
- Runtime OS includes clickable Holy Grail controls:
  - signalling slips 4 months
  - procurement evidence weakens
  - reserve cut by 10%
  - operator acceptance delay
- Runtime controls mutate the active model state: P50, schedule, confidence, board briefing and CASEY thinking update together.
- Exports use the current model payload after runtime changes.

How to demo:
1. Generate or open a project.
2. Click **Advisor** to upload/run sample challenge and show the professional intake output.
3. Click **Runtime OS** to use the live what-if buttons.
4. Return to Overview/Scenario Intel/QCRA/QSRA to show state propagation.
5. Use Outputs to export the changed model.

Validation performed:
- Frontend production build passed.
- Backend Python compile passed.
