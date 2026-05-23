# CASEY One-Shot Public Demo System

This build adds a public one-shot demo flow:

- Frontend modal: "Run one free project"
- Backend endpoint: `POST /public-demo/generate`
- One run per email/network/browser/device token
- Guided input gate to force better quality briefs
- Generates first-pass class estimate, schedule intelligence and risk register
- Saves leads and generated models to SQLite
- Feedback endpoint for a controlled improvement loop

## Important anti-abuse note
No system can perfectly stop VPNs, device randomizers or determined abuse. This build makes abuse harder using:

- email check
- IP hash
- browser/device fingerprint hash
- persistent client token
- backend database enforcement

For stronger production enforcement add:

- email verification / magic link
- Cloudflare Turnstile
- paid account login
- stricter rate limiting

## Learning loop
Do not train live on user data automatically. This build logs public demo requests and feedback so you can review weak outputs and improve prompts/examples safely.

Recommended loop:

1. User runs project
2. CASEY stores brief + output + optional feedback
3. You review low-rated runs weekly
4. Improve rules/prompts/templates
5. Later fine-tune only with consented, cleaned data

## Deploy
Push this ZIP to GitHub. Render redeploys backend. Vercel redeploys frontend. Ensure environment variables include:

- `OPENAI_API_KEY` when you later switch to LLM generation
- `CASEY_DB=casey_titan_v26.sqlite3`
- `CASEY_PUBLIC_DEMO_LIMIT=1`
- `CASEY_DEMO_LIMIT_PER_IP=1`