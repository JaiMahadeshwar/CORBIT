# CORBIT deployment settings

## Frontend on Vercel

- Root Directory: `frontend`
- Framework: `Vite`
- Install Command: `npm install --legacy-peer-deps --no-audit --no-fund`
- Build Command: `npm run build`
- Output Directory: `dist`

## Backend on Render

- Root Directory: `backend`
- Build Command: `pip install -r requirements.txt`
- Start Command: `uvicorn main:app --host 0.0.0.0 --port 10000`
- Environment variable: `OPENAI_API_KEY`

## Video

The frontend points to Bunny CDN:

`https://corbit.b-cdn.net/casey_hero_film.mp4`

Do not put the MP4 back in GitHub.
