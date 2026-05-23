# CASEY V116 Elite Demo Completion Notes

This build is intended for the strongest public demo version before the full production tool.

## What was added

- Confidence is now framed as **board-defensibility, not generic optimism**.
- Overview now includes an evidence threshold map showing benchmark fit, evidence maturity, procurement certainty, schedule logic and resilience / reserve.
- Overview now includes a contradiction scan so CASEY exposes the trade-off instead of making every scenario look positive.
- Faster, Cheaper, Lower Risk, Premium and Base scenarios remain tied to cost, schedule, risk, confidence, QCRA/QSRA and export logic.
- Local/admin testing can reset one test email without weakening the public one-run lock.
- Health endpoint added at `/elite-demo/health`.
- Test reset endpoint added at `/elite-demo/reset-test-email`.

## Public lock rule

Public users still get one credible Earth or Space intelligence run per email.

Admin testers can use emails listed in `CASEY_ADMIN_EMAILS`, or reset a local test email from localhost using:

```powershell
Invoke-RestMethod -Method Post -Uri http://127.0.0.1:8000/elite-demo/reset-test-email -ContentType "application/json" -Body '{"email":"test@yahoo.com"}'
```

## Demo close line

Traditional project controls reports show numbers. CASEY shows the board what the numbers are trying to hide.
