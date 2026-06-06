# CASEY V142 Final 10/10 Demo Credibility Drop

Real patch applied on top of V141 crash fix.

## What changed
- Increased reserve realism across all scenarios; reserve is now scenario-controlled and high enough to survive board challenge.
- Added final credibility polish layer in backend: irregular confidence outputs, asymmetric QCRA/QSRA tails, and non-neat schedule movement.
- Rebuilt scenario matrix so Base/Faster/Cheaper/Lower Risk/Premium show less synthetic numbers.
- Reconciled Direct + Indirect + Reserve to selected P50 after final scenario changes.
- Added audit spine fields: P50 reconciliation, P80 downside, reserve challenge, evidence gate.
- Reworded visible AI-style language into operational board language.
- Changed CASEY benchmark label from marketing-like top-decile to 91st pct.
- Preserved V141 crash fix for parseMoneyLocal.

## Validation performed
- Backend Python compile passed.
- Frontend npm build passed after npm install.
- Generated Base/Faster/Cheaper/Lower Risk/Premium via FastAPI TestClient.
- Verified cost rows reconcile to selected P50 for all five scenarios.

## Demo guidance
Use Base, then switch Faster and Cheaper to show ugly tail expansion. Then switch Lower Risk or Premium to show reserve and confidence buying down P80/P90 exposure.
