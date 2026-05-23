# CASEY v9.5 Demo Launch QA Report

## v9.5 additions

This build adds a deeper scenario cascade for the public demo:

- Scenario-specific executive shock insights
- Strategic Delta tab
- Confidence Breakdown panel
- Top Decisions Required panel
- Board Memo Snapshot
- Scenario-specific risk register mutation
- Scenario-specific cost line mutation
- Scenario-specific schedule logic mutation
- Scenario-specific QCRA/QSRA curve mutation
- Faster / Cheaper / Lower Risk / Premium logic now has clearer causal consequences
- Outputs now carry scenario explanation and decision posture
- Life Sciences and Space examples tested

## QA checks

| Example | Scenario | Mode | Subsector | P50 | Schedule | Risk | Confidence | Delta | Risks | Pass |
|---|---|---|---|---|---|---|---:|---:|---:|---|
| life sciences | base | Earth | Life Sciences / Biologics Manufacturing | $6.9B | 73 months | Medium-High | 61% | 4 | 8 | Yes |
| life sciences | faster | Earth | Life Sciences / Biologics Manufacturing | $7.9B | 60 months | High | 44% | 4 | 10 | Yes |
| life sciences | cheaper | Earth | Life Sciences / Biologics Manufacturing | $6.1B | 80 months | High | 43% | 4 | 10 | Yes |
| orbital | base | Space | Orbital Compute / Manufacturing | $53.5B | 159 months | Medium-High | 47% | 4 | 10 | Yes |
| orbital | faster | Space | Orbital Compute / Manufacturing | $61.0B | 112 months | High | 33% | 4 | 10 | Yes |
| orbital | cheaper | Space | Orbital Compute / Manufacturing | $47.1B | 175 months | High | 28% | 4 | 10 | Yes |

## Demo positioning

Use this as the final demo launch build. It is not a certified estimating system. It is a first-pass project intelligence and decision-support demo.

Recommended demo flow:
1. Run Life Sciences example.
2. Show Executive Shock Insight.
3. Click Faster, Cheaper, Lower Risk.
4. Open Strategic Delta.
5. Open Risk, Schedule, QCRA/QSRA.
6. Explain that the full enterprise version unlocks downloadable board packs, XER, cost workbook and risk register.
