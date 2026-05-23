# CASEY V97 Executive Curve + Scenario Intelligence QA

Applied final demo-readiness fixes:

- QCRA/QSRA curves now reconcile exactly to the headline P50 values.
- Scenario curves now have distinct executive meaning:
  - Base: balanced board reference case.
  - Faster: shorter P50, higher P80/P90 tail from consumed float and concurrency.
  - Cheaper: lower P50, ugly/fat tail from stripped reserve and deferred resilience.
  - Lower Risk: higher/slower median, compressed downside tail.
  - Premium: higher approval value, controlled downside tail.
- Cost range now uses the same QCRA curve as the dashboard and exports instead of separate synthetic multipliers.
- QCRA/QSRA tooltips now carry percentile meaning and P50 anchoring.
- Added curve readout language so non-technical executives understand what the graph means.
- Added Scenario Intel tab to expose strategic delta, confidence movement, gained/lost trade-offs and top decisions required.
- Frontend build completed successfully.

Smoke test example used:
Global biologics and obesity therapeutics manufacturing campus in North Carolina with multi-product GMP facilities, aseptic fill-finish lines, high-throughput packaging, validated clean utilities, automated warehousing, cold-chain distribution and accelerated scale-up for commercial launch demand.

Scenario check:
- Base: P50 cost $6.9B, QCRA P50 6.9, schedule 73 months, QSRA P50 73.
- Faster: P50 cost $8.1B, QCRA P50 8.1, schedule 57 months, QSRA P50 57, wider P80/P90 tail.
- Cheaper: P50 cost $5.9B, QCRA P50 5.9, schedule 83 months, QSRA P50 83, fat risk tail.
- Lower Risk: P50 cost $7.7B, QCRA P50 7.7, schedule 85 months, QSRA P50 85, compressed tail.
- Premium: P50 cost $8.8B, QCRA P50 8.8, schedule 70 months, QSRA P50 70, controlled tail.

Known note:
The frontend build warning about chunk size is non-blocking for demo use.
