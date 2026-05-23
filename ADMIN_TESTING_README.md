# CASEY V108 Admin Testing + Public Lock

This build keeps the public demo commercial lock but lets you test unlimited runs.

## Public rule
- One credible Earth OR Space run per email.
- One demo export download per export type/run unless paid access is enabled.

## Admin testing rule
Set this in backend `.env` or in your terminal before starting the backend:

```powershell
$env:CASEY_ADMIN_EMAILS="test@yahoo.com,jim@yahoo.com,jai@yahoo.com"
```

Any email in `CASEY_ADMIN_EMAILS` can run unlimited tests and repeat exports.
Public users still get one run only.

Optional API key bypass:

```powershell
$env:CASEY_ADMIN_KEY="your-secret-key"
```

Then call `/demo/status?email=someone@example.com&admin_key=your-secret-key`, or send header `x-casey-admin-key`.

## Start
Backend:
```powershell
cd backend
python -m uvicorn main:app --reload
```

Frontend:
```powershell
cd frontend
npm install
npm run dev
```

## Test admin status
Open:
`http://127.0.0.1:8000/demo/status?email=test@yahoo.com`

Expected:
`admin_bypass: true`, `remaining: unlimited`.
