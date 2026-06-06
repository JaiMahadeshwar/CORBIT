# CASEY TITAN X V147 Production Launch Checklist

## Demo theatre
1. Start backend and frontend.
2. Open the app.
3. Run **Earth demo** first.
4. Walk tabs in order: Overview, Scenarios, Cost, Schedule, Risk Map, Monte Carlo, Peers, Advisor, Uploads, Exports.
5. Run **Space demo** and repeat the same journey.
6. Close by downloading the ZIP output pack.

## Production setup
- Put `OPENAI_API_KEY` in backend `.env`; never expose it in frontend.
- Use hosted Postgres for `CASEY_DB` in production.
- Add Clerk/Auth0/Supabase Auth before real customers.
- Configure Stripe price IDs before enabling payments.
- Deploy frontend on Vercel/Netlify.
- Deploy backend on Render/Railway/Fly/Azure/AWS.
- Add domain + SSL + logging + backups.

## Demo promise
CASEY demonstrates how one prompt becomes a connected capital-project intelligence pack: estimate, schedule, mapped risks, QCRA/QSRA, scenarios, peer intelligence, uploads and board exports.
