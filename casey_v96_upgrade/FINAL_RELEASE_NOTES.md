# CORBIT Ultimate Release

This package is deployment-ready for the current architecture:

- Frontend: Vercel, root directory `frontend`
- Backend: Render, root directory `backend`
- Video: Bunny CDN at `https://corbit.b-cdn.net/casey_hero_film.mp4`

Vercel settings:

- Framework: Vite
- Root Directory: frontend
- Install Command: npm install --legacy-peer-deps
- Build Command: npm run build
- Output Directory: dist

This release includes the final cinematic first-viewport CSS pass and keeps `react-is` in dependencies to avoid the Recharts/Vite build issue.
