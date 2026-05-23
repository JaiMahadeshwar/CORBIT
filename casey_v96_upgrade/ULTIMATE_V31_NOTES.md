# CASEY TITAN X — Ultimate V31

Built from the uploaded V30 package.

## What changed
- Chose **Option A** for the background video: `nzly71xHgwE`.
- Replaced the moving-planet hero treatment with a muted cinematic YouTube background layer.
- Added mission-control HUD rings, grid overlays, orbital signal pulse, and darker aerospace contrast.
- Updated the brand language to feel more memorable: **Control the orbit before it breaks budget.**
- Reworked buttons, cards, tabs, headers, KPIs, tables, and export buttons into a sharper SpaceX-style command console.
- Kept the existing FastAPI backend, generation endpoints, advisor chat, demo flow, and export factory.
- Rebuilt the production Vite frontend in `frontend/dist`.

## Run locally
1. Start backend: `START_BACKEND.bat` or `cd backend && uvicorn main:app --reload`
2. Start frontend: `START_FRONTEND.bat` or `cd frontend && npm install && npm run dev`
3. Open the frontend URL shown by Vite.

## Background video note
The app embeds the background from YouTube as a muted iframe. Autoplay depends on browser and network rules. If you want a fully offline version later, replace the iframe with a local `.mp4` in `frontend/public` and point `.heroVideo` to a `<video>` element.
