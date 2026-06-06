# CASEY V150 Enterprise Trust Runtime — QA Report

## What was actually changed

- Added V150 canonical scenario runtime on top of V148 trust core.
- Added governance state metrics:
  - Board defensibility
  - Governance stress
  - Tail exposure
  - Evidence volatility
  - Reserve pressure
  - Confidence drift
- Added weighted confidence breakdown:
  - benchmark fit
  - schedule density
  - procurement certainty
  - evidence maturity
  - reserve adequacy
  - systems integration
  - operational readiness
- Added deterministic causal graph / propagation checksum.
- Added narrative entropy watch to detect repeated synthetic language.
- Added V150 export manifest and scenario-state JSON into full export packs.
- Re-registered core export routes so payloads are re-canonicalized before files are built.
- Added QA endpoints:
  - `/qa/v150/validate-model`
  - `/qa/v150/readiness`
- Added frontend V150 Trust Runtime Bar.
- Added `tools/v150_mass_validation.py` for dedicated-machine 10,000 Earth / 5,000 Space validation.

## Build checks completed in this environment

- Backend Python compile: PASSED
- Frontend production build: PASSED
- V150 local mass validation sample: 100 Earth + 100 Space = 200 scenarios, 0 failures
- V150 export smoke sample: 10 full output packs, 0 failures

## Important honesty note

The full 10,000 Earth + 5,000 Space stress run did not complete inside this execution window. The V150 package includes the actual harness needed to run it locally:

```bash
python tools/v150_mass_validation.py --earth 10000 --space 5000 --exports 100
```

For a production-grade signoff, run that on a dedicated machine and archive the generated JSON report.

## Demo readiness

This version is stronger than V148 for executive demos because it exposes a persistent trust runtime, scenario signature, propagation checksum, governance metrics and export manifests.

It is still best described as a strategic intelligence simulation / executive demo platform, not a certified estimate engine.
