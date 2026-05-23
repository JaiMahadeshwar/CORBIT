# CASEY Mission-Control One-Try Demo Build

## What changed

This build keeps the homepage cinematic and opens the free demo as a dark mission-control access modal.

The free run now feels like:

- CASEY Intelligence Access
- One public mission run
- Earth / Orbital infrastructure
- exclusive controlled access
- not a cheap calculator form

## Prompt forcing / quality gate

The frontend now guides users through:

1. Work email
2. Project domain: Earth or Space
3. Location / operating environment
4. Size / capacity
5. Stage
6. Main concern
7. Programme brief

The backend blocks weak/random prompts before generation.

Blocked examples include:

- asdf / qwerty / lorem ipsum
- hello world
- jokes / non-project requests
- prompt injection attempts
- vague one-line prompts
- contradictory Earth vs Space prompts
- LEO + Moon/Mars mixed location prompts
- repeated nonsense text

The free run is not consumed unless the backend accepts the brief.

## Output positioning

Use this wording publicly:

"CASEY generates a first-pass project intelligence pack for early-stage cost, schedule and risk thinking."

Do not call outputs certified estimates.

## Logging

Runs are stored in SQLite:

- public_demo_uses
- projects

Feedback is stored in:

- public_demo_feedback

Admin endpoints are protected by CASEY_ADMIN_TOKEN.