# CASEY Final Public Demo Host-Ready Build

Use this ZIP for the public one-free-run demo.

## What this build supports

- Turner & Townsend / Mace / Jacobs / AECOM-style Earth infrastructure projects
- all sizes from small to mega-mega
- estimate class 1 to 5
- schedule level 1 to 5
- risk register outputs
- XER-style schedule logic
- benchmark memory
- procurement heatmaps
- critical path narrative
- orbital compute / space data centre prompts
- lunar / Mars / LEO / cislunar infrastructure prompts

## Prompt protection

The public demo blocks:
- random nonsense
- vague prompts
- joke prompts
- prompt injection attempts
- contradictory Earth/Space prompts
- weak briefs without asset/location/size/stage/concern

## Deployment env vars

Backend Render:
CASEY_PUBLIC_DEMO_LIMIT=1
CASEY_DEMO_LIMIT_PER_IP=1
CASEY_DB=casey_titan_v26.sqlite3
CASEY_ADMIN_TOKEN=choose-a-long-secret

Frontend Vercel:
VITE_API_URL=https://YOUR-RENDER-BACKEND.onrender.com

## Public wording

"CASEY generates first-pass project intelligence for early-stage cost, schedule and risk thinking."

Do not call it a certified estimate.