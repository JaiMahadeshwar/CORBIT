# CASEY v9.5+ Holy-Shit Demo QA Report

## What was fixed

This build is focused on making scenario intelligence feel real, scary and board-grade.

### Scenario consequence engine
Every scenario now has a clear strategic trade:
- what you gained
- what you gave up
- board consequence
- confidence movement
- reserve philosophy
- scenario-controlled export metadata

### QCRA / QSRA curves
Curves now have different personalities:
- Base: balanced uncertainty
- Faster: compressed median but aggressive P80/P95 tail
- Cheaper: lower median but ugly long tail and fragile reserve
- Lower Risk: tighter distribution and lower tail exposure
- Premium: higher median but controlled downside

### Risk register
Scenario-specific risks are inserted at the top:
- Faster: concurrency, acceleration premium, recovery float exhaustion, CQV / mission-assurance compression
- Cheaper: deferred resilience, contingency underfunding, lifecycle transfer, clean utility / redundancy gaps
- Lower Risk: assurance gates and protected float
- Premium: scope discipline and value-for-money challenge

### Schedule / XER logic
Schedule rows now change basis and criticality by scenario:
- Faster: overlapped workfronts, reduced float, near-critical density
- Cheaper: deferred redundancy, slower phasing, longer readiness tail
- Lower Risk: evidence gates, protected commissioning float
- Premium: priority procurement and managed readiness gates

### Cost estimate logic
Cost lines no longer move as a flat percentage only:
- Faster adds expediting, premium suppliers and productivity drag
- Cheaper reduces capital scope but flags transferred operational risk
- Lower Risk adds assurance, procurement buffer and validation float
- Premium adds resilience and redundancy

### Semiconductor routing
Semiconductor / fab / EUV / wafer / UPW prompts now override accidental life-sciences classification.

## QA sample

| Example | Scenario | Mode | Subsector | P50 | Schedule | Risk / Conf | QCRA P50→P90 | QSRA P50→P90 | Top risk | Pass |
|---|---|---|---|---|---|---|---|---|---|---|
| life | base | Earth | Life Sciences / Biologics Manufacturing | $6.9B | 73 months | Medium-High / 61% | 6.9→8.0 | 73→82 | Market escalation | Yes |
| life | faster | Earth | Life Sciences / Biologics Manufacturing | $8.1B | 57 months | High / 41% | 8.1→10.5 | 57→66 | Media-fill readiness collision | Yes |
| life | cheaper | Earth | Life Sciences / Biologics Manufacturing | $5.9B | 83 months | High / 41% | 5.9→7.5 | 83→113 | Clean utility redundancy gap | Yes |
| life | lower_risk | Earth | Life Sciences / Biologics Manufacturing | $7.7B | 85 months | Medium / 86% | 7.7→8.2 | 85→89 | Assurance gate delay | Yes |
| life | premium | Earth | Life Sciences / Biologics Manufacturing | $8.8B | 70 months | Medium-Low / 93% | 8.8→9.4 | 70→72 | Premium scope creep | Yes |
| semi | base | Earth | Semiconductor / Advanced Manufacturing | $9.5B | 83 months | Medium-High / 58% | 9.5→11.0 | 83→93 | Market escalation | Yes |
| semi | faster | Earth | Semiconductor / Advanced Manufacturing | $11.2B | 65 months | High / 41% | 11.2→14.4 | 65→75 | Concurrent commissioning overload | Yes |
| semi | cheaper | Earth | Semiconductor / Advanced Manufacturing | $8.2B | 95 months | High / 37% | 8.2→10.4 | 95→129 | Deferred resilience failure | Yes |
| semi | lower_risk | Earth | Semiconductor / Advanced Manufacturing | $10.6B | 96 months | Medium / 82% | 10.6→11.3 | 96→101 | Assurance gate delay | Yes |
| semi | premium | Earth | Semiconductor / Advanced Manufacturing | $12.2B | 80 months | Medium-Low / 90% | 12.2→13.0 | 80→83 | Premium scope creep | Yes |
| space | base | Space | Orbital Compute / Manufacturing | $53.5B | 159 months | Medium-High / 47% | 53.5→62.0 | 159→179 | Launch manifest delay | Yes |
| space | faster | Space | Orbital Compute / Manufacturing | $63.1B | 122 months | High / 30% | 63.1→81.3 | 122→142 | Mission assurance compression | Yes |
| space | cheaper | Space | Orbital Compute / Manufacturing | $46.0B | 181 months | High / 27% | 46.0→58.5 | 181→246 | Redundancy deletion after deployment | Yes |
| space | lower_risk | Space | Orbital Compute / Manufacturing | $59.9B | 184 months | Medium / 74% | 59.9→63.5 | 184→193 | Assurance gate delay | Yes |
| space | premium | Space | Orbital Compute / Manufacturing | $68.5B | 153 months | Medium-Low / 81% | 68.5→73.0 | 153→158 | Premium scope creep | Yes |

## Demo guidance

For Turner & Townsend / Mace / Jacobs-type users, demo this sequence:

1. Run Life Sciences.
2. Show Base executive shock.
3. Click Cheaper.
4. Point to Strategic Delta: “what you gained / what you gave up.”
5. Show risk register top risks changed.
6. Show QCRA/QSRA curves changed: cheaper has ugly tail.
7. Click Faster.
8. Show acceleration risk and CQV compression.
9. Click Outputs and say full users get scenario-controlled cost workbook, XER, risk register and board pack.

## Caveat

This is still first-pass strategic intelligence, not a certified estimate or schedule. Use it for demo and positioning.
