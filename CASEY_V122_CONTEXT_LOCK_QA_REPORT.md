# CASEY V122 Context Lock QA Report

## Fix applied
- Scenario selection now carries a canonical active project context from the current model.
- `/generate` accepts `active_model` and rebuilds scenarios from the locked prompt/archetype instead of falling back to the default demo seed.
- Backend guard preserves `mode`, `subsector`, `title`, `location`, and `scale` if a non-base scenario ever classifies away from the locked project universe.
- Added explicit Space Power Grid classifier override so LEO/orbital power infrastructure cannot be misrouted to Earth energy.

## Regression tests run
- 100,000 Earth synthetic projects
- 100,000 Space synthetic projects
- Rotating scenarios across Base, Faster, Cheaper, Lower Risk, Premium
- Total projects tested: 200,000
- Total scenario-context transitions tested: 200,000
- Result: PASS

## Build verification
- Backend Python compile: PASS
- Frontend Vite production build: PASS

## Demo-critical behavior
- Space -> Faster remains Space
- Space -> Cheaper remains Space
- Space -> Lower Risk remains Space
- Space -> Premium remains Space
- Earth examples remain Earth
- Non-default archetypes are protected by active context lock
