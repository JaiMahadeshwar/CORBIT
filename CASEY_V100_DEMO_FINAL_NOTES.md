# CASEY V100 Demo Final

Final demo hardening included:

- Public demo limiter disabled for demo-launch mode so repeat local/browser runs no longer return 403.
- Scenario buttons now re-run the backend engine instead of using a front-end-only sensitivity stub.
- Base, Faster, Cheaper, Lower Risk and Premium now keep dashboard KPIs, QCRA, QSRA, cost workbook, schedule logic, risk register and exports aligned.
- QCRA curve is now explicitly a cost probability distribution, not spend over time.
- QSRA curve is now explicitly a finish-date probability distribution, not related to cost.
- Curve tooltips and labels now say QCRA total outturn and QSRA finish date.
- P50 in the QCRA/QSRA curves reconciles to the headline P50 cost and headline schedule.
- P80/P90 show board downside exposure / finish-risk tail.
- Scenario comparison versus Base is retained in the model and exports.

Smoke test completed on AI data centre scenarios:

- Base: reference case.
- Faster: higher cost, shorter duration, lower confidence, wider tail.
- Cheaper: lower P50, slower duration, lower confidence, fat P80/P90 risk tail.
- Lower Risk: higher cost, slower duration, higher confidence, tighter tail.
- Premium: higher cost, similar/faster duration, stronger confidence and controlled downside.
