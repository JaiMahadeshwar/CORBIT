# CASEY Public Demo Final Validator + Output QA Report

## What changed

This build fixes the issue where credible executive briefs were blocked by the public demo validator.

The validator now:
- allows credible infrastructure briefs even if they are not written in a rigid format
- correctly accepts the life sciences North Carolina example
- still blocks rubbish, prompt injection and repetitive spam
- uses asset + location/environment + concern signals rather than rigid word-count gates
- keeps one public run logic
- keeps scenarios usable inside the same generated intelligence pack
- keeps enterprise outputs locked

## Validator QA

| Test | Should pass | Passed | Score | Issues |
|---|---:|---:|---:|---|
| Life sciences full | Yes | Yes | 100 | - |
| Short credible data centre | Yes | Yes | 100 | - |
| Short credible space | Yes | Yes | 90 | - |
| Rubbish | No | No | 30 | Add a little more project detail: asset, location/environment and main concern.; State the asset type, e.g. data centre, GMP campus, rail, hospital, lunar base or orbital platform.; Add a location or operating environment. |
| Jailbreak | No | No | 44 | This does not look like a credible project brief. Please enter a real infrastructure or space programme. |
| Repetitive | No | No | 72 | The brief appears repetitive or low quality. Please write a real project description. |

## Output QA

| Example | Mode | Subsector | P50 | Schedule | Risk | Confidence | Pass |
|---|---|---|---|---|---|---:|---|
| life_sciences | Earth | Life Sciences / Biologics Manufacturing | $6.9B | 73 months | Medium-High | 61% | Yes |
| data_centre | Earth | Digital Infrastructure / Hyperscale Data Centre | $8.2B | 69 months | Medium | 65% | Yes |
| semiconductor | Earth | Semiconductor / Advanced Manufacturing | $21.2B | 92 months | Medium | 66% | Yes |
| airport | Earth | Airport / Aviation | $21.9B | 131 months | Medium-High | 61% | Yes |
| rail | Earth | Rail / Transit | $8.3B | 99 months | Medium | 66% | Yes |
| defence | Earth | Defence / Secure Mission Infrastructure | $5.5B | 69 months | Medium | 66% | Yes |
| orbital | Space | Orbital Compute / Manufacturing | $53.5B | 159 months | Extreme | 47% | Yes |
| mars | Space | Mars Surface Habitat/Base | $138.6B | 230 months | Extreme | 47% | Yes |

## Key expected behaviour

The following should now run:

Global biologics and obesity therapeutics manufacturing campus in North Carolina with multi-product GMP facilities, aseptic fill-finish lines, high-throughput packaging, validated clean utilities, automated warehousing, cold-chain distribution and accelerated scale-up for commercial launch demand. Programme includes process integration, CQV validation, FDA inspection readiness, operational staffing, utility resilience and phased production ramp-up while maintaining continuity of supply. Main concerns are qualification sequencing, specialist equipment procurement, commissioning throughput, regulatory readiness, utility reliability, schedule compression and maintaining production continuity during phased expansion.

Expected output:
- Earth
- Life Sciences / Biologics Manufacturing
- sector-specific CQV/GMP/FDA language
- no Space misclassification
- no launch-readiness leakage
- coherent cost/schedule/risk pages
- locked enterprise outputs
