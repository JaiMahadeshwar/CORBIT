# CASEY / CONTROLORBIT — Deploy Now

## Your URLs after deploy
- Frontend: https://your-app.vercel.app
- Backend: https://casey-backend.onrender.com

## Step 1: Backend on Render (free tier works)
1. Go to render.com → New Web Service
2. Connect your GitHub repo (push this folder first)
3. Root directory: `backend`
4. Build command: `pip install -r requirements.txt`
5. Start command: `uvicorn main:app --host 0.0.0.0 --port $PORT`
6. Set environment variables:
   - `CASEY_ADMIN_EMAILS` = your real email (gets unlimited runs)
   - `CASEY_DEMO_LIMIT_PER_IP` = 1
   - `CASEY_ADMIN_KEY` = pick a long secret key (e.g. casey_prod_2024_secret)

## Step 2: Frontend on Vercel (free)
1. Go to vercel.com → New Project
2. Connect GitHub repo
3. Root directory: `frontend`
4. Build command: `npm run build`
5. Output directory: `dist`
6. Set environment variable:
   - `VITE_BACKEND_URL` = your Render URL (e.g. https://casey-backend.onrender.com)

## Demo gate behaviour
- Public visitors: 1 model run, 1 export download, then locked
- Your admin access: add `?admin=YOUR_KEY` to the URL
  e.g. https://your-app.vercel.app?admin=casey_prod_2024_secret
- Locked users see the Pricing tab with contact details

## Admin email bypass
Set CASEY_ADMIN_EMAILS to your email on Render.
When you enter your email in the generate form, the gate lifts automatically.

## Demo script (60 seconds)
1. Open the URL. Type: "HS2 Phase 2b tunnelling stations systems migration"
2. Click Generate. Watch the model build in 3 seconds.
3. Go to Live Stress Test → click "What if procurement evidence is missing?"
4. Watch confidence drop, P80 rise, board language change.
5. Click Export Board Pack. Download the PDF.
6. Go to Advisor → click "If this programme fails, what will be blamed publicly?"
7. Show the answer. Ask T&T to match it in real time.
