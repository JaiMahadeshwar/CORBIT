# CASEY V130.1 Executive Intelligence QA Report

## What was actually changed

- Added a final sector-first routing hotfix so old fallback subsectors cannot contaminate new runs.
- Added full sector-native causal graph overrides for Space, Defence, Energy, Oil & Gas, Healthcare, Water, Ports, Mining and Roads, alongside the existing Airport, Rail, Hyperscale/Data Centre, Semiconductor and Life Sciences hardening.
- Added governance challenge intelligence: owner evidence, board challenge questions, risk-transfer warnings and confidence-defensibility language.
- Added second-order contradiction intelligence: acceleration, cheaper, base and resilience trade-offs now challenge the hidden consequence, not only the headline number.
- Tightened export-facing fields: cost lines, schedule rows, risk register, benchmark memory, causal chain, confidence drivers and hidden JSON surfaces are rebuilt from the locked sector graph.
- Added front-end display support for model-generated contradiction intelligence.
- Added visual polish pass: lower text density in key sections, smoother transitions and more premium chart styling.

## Smoke-tested sectors

| Sector | Routing | Obvious ontology leak check |
|---|---:|---:|
| Airport / Aviation | Pass | Pass |
| Rail / Transit | Pass | Pass |
| Digital Infrastructure / Hyperscale Data Centre | Pass | Pass |
| Space / Mission Assurance | Pass | Pass |
| Defence / Secure Infrastructure | Pass | Pass |
| Energy / Power Infrastructure | Pass | Pass |
| Oil & Gas / Process Infrastructure | Pass | Pass |
| Healthcare / Hospital Infrastructure | Pass | Pass |
| Water / Environmental Infrastructure | Pass | Pass |
| Ports / Marine Infrastructure | Pass | Pass |
| Semiconductor / Advanced Manufacturing | Pass | Pass |
| Life Sciences / Biologics Manufacturing | Pass | Pass |

## Specific leak fixes verified

- Liquid cooling is allowed in hyperscale/data centre only.
- ORAT/baggage/airside language is allowed in airport only.
- Possessions/signalling/rolling stock are locked to rail.
- Mission assurance/payload/range language is locked to space.
- HAZOP/LNG/hydrocarbon language is locked to oil and gas.
- CQV/GMP/media-fill language is locked to life sciences.
- Lithography/yield/UPW language is locked to semiconductor.

## Build check

Frontend production build completed successfully with Vite.
Backend Python compile check completed successfully.

## Remaining honest note

This is a strong pre-demo hardening pass. It is not a mathematical proof across every possible natural-language prompt, but it directly addresses the leakage patterns seen in the screenshots and adds a final sector-first routing guard so explicit user project descriptions win over stale fallback model classifications.
