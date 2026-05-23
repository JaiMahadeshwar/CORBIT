# CASEY Public One-Try Demo Launch Notes

Use this build to open a controlled public demo on controlorbit.com.

Recommended CTA:
"Run One Free CASEY Intelligence Pack"

Recommended form intro:
"Describe one Earth or Space project. CASEY will generate a first-pass class estimate, schedule and risk view. One free run per user."

Recommended disclaimer:
"CASEY public demo outputs are first-pass strategic project-controls intelligence. They are not certified QS, engineering or tender advice."

Recommended production environment variables:
- CASEY_PUBLIC_DEMO_LIMIT=1
- CASEY_DEMO_LIMIT_PER_IP=1
- CASEY_DB=casey_titan_v26.sqlite3
- OPENAI_API_KEY=only when enabling LLM polish later

Recommended future hardening:
- work-email verification / magic link
- Cloudflare Turnstile
- admin dashboard for reviewing public demo runs
- feedback review workflow
- weekly prompt/rule improvement cycle