# CASEY V98 Executive Scenario + QCRA Fix Report

## What was fixed

- Added explicit Scenario vs Base comparison to the executive overview.
- Added selected-scenario delta fields: cost delta, schedule delta, confidence delta and plain-English board interpretation.
- Added cost bridge vs Base explaining why Faster/Cheaper/Lower Risk/Premium moved.
- Added schedule bridge vs Base explaining how months are saved or added.
- Reworked QCRA/QSRA display copy so users understand these are probability curves, not cashflow or monthly schedule charts.
- Anchored QCRA and QSRA curves so P50 exactly reconciles to the headline KPI.
- Added P50 and P80 reference markers to QCRA/QSRA charts.
- Added P50/P80/P90 metric cards above QCRA/QSRA curves.
- Stopped build_model from applying scenario multipliers before the scenario cascade; the cascade now compares each scenario against the true Base.
- Board briefing and output memo now include scenario-vs-base language.

## QA spot-check

Tested backend model generation for Base, Faster and Cheaper on a hyperscale AI data centre prompt.

- Base: P50 curve point equals headline P50.
- Faster: scenario comparison shows higher cost, faster schedule and lower confidence vs Base.
- Cheaper: scenario comparison shows lower cost, slower schedule and lower confidence vs Base.
- QCRA/QSRA P50 point equals headline KPI in the returned model.

## Remaining note

Frontend dependency install/build was not run in this container because node_modules/vite are not present in the unpacked archive. Backend Python compiles successfully.
