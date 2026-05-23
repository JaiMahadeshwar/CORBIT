# CASEY V133 No-Blank-Screen Hardening Report

## What was fixed

This build adds frontend render hardening so malformed or incomplete sector payloads cannot blank the full product UI.

### Added
- Global React error boundary around the entire application.
- Sector model normalisation before `setModel()`.
- Defensive QCRA/QSRA curve validation before Recharts render.
- Safe fallback causal graphs for airport, rail, hyperscale, pharma, semiconductor, oil & gas, nuclear, defence, space, water, ports, healthcare and energy.
- Safe default arrays for cost breakdown, schedule detail, risk register, scenario matrix, benchmark comparison, mission-control cards and board briefing.
- Sector vocabulary scrubber to reduce obvious ontology leakage in the rendered model.
- Graceful recovery screen instead of black screen if a React render exception still occurs.

## Specific blank-screen risk addressed

The rail example likely exposed an incomplete/invalid render object in one of these areas:
- causal graph nodes
- QCRA/QSRA curve points
- scenario waterfall data
- benchmark comparison data
- schedule-detail table rows

V133 now normalises these before rendering.

## Build validation

Frontend production build completed successfully with Vite.

Backend Python compile completed successfully.

Smoke-tested backend generation for:
- Rail / Transit
- Airport / Aviation
- Space / Mission Assurance

## Important note

This is a no-blank-screen hardening fix. It does not remove the need for final browser walkthrough on the live Vercel deployment after upload, because browser/CDN cache and production environment variables can still differ from local files.
