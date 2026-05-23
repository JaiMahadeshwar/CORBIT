# CASEY One-Try Demo — UI + Logging Ready

## What users see

The homepage now clearly says:

- Run One Free CASEY Intelligence Pack
- Enter your email
- Describe one Earth or Space project
- Receive a first-pass class estimate, schedule view and risk register
- One run per user

The demo modal asks for:

- Work email
- Project type: Earth or Space
- Location
- Size / capacity
- Stage
- Main concern
- Project brief

## Where user details are logged

The backend stores public demo runs in SQLite.

Database path is controlled by:

`CASEY_DB`

Default:

`casey_titan_v26.sqlite3`

Table:

`public_demo_uses`

Stored fields:

- run_id
- email_hash
- ip_hash
- fingerprint_hash
- client_token_hash
- project_type
- project_text
- model_json
- created_at

Important: emails are hashed in the current table for privacy. The raw project brief and generated output are stored.

## Where feedback is logged

Table:

`public_demo_feedback`

Stored fields:

- run_id
- rating
- comment
- created_at

## Where full generated outputs are logged

The generated CASEY model/output is stored in:

1. `public_demo_uses.model_json`
2. `projects.model_json`

The original project brief is stored in:

1. `public_demo_uses.project_text`
2. `projects.prompt`

## Admin log access

This build adds two admin endpoints:

`GET /public-demo/admin/summary`

`GET /public-demo/admin/runs?limit=100`

To use them, set this Render environment variable:

`CASEY_ADMIN_TOKEN=choose-a-long-secret-password`

Then call the endpoint with this header:

`x-casey-admin-token: your-secret-token`

These endpoints return recent demo runs, project briefs and generated outputs.

## Render environment variables

Use:

`CASEY_PUBLIC_DEMO_LIMIT=1`

`CASEY_DEMO_LIMIT_PER_IP=1`

`CASEY_DB=casey_titan_v26.sqlite3`

`CASEY_ADMIN_TOKEN=your-long-secret-admin-token`

Optional later:

`OPENAI_API_KEY=your-api-key`

## Public wording

Use this wording:

"CASEY generates first-pass strategic project-controls intelligence. Outputs should be reviewed by qualified professionals before commercial use."

Do not call outputs certified estimates.