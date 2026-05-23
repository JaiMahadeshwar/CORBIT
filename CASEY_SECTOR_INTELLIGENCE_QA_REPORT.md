# CASEY Sector Intelligence Upgrade QA Report

## Build changes

This build upgrades CASEY output quality across Earth and Space sectors.

Added:
- sector cost/schedule realism envelopes
- sector-specific cost drivers
- sector-specific schedule threats
- sector-specific confidence drivers
- sector-specific CASEY Thinking
- cleaner benchmark comparison sets
- improved scenario math so Cheaper, Faster, Lower Risk and Premium move coherently
- stronger frontend scenario logic using immutable base-case values
- cross-sector wording cleanup so Life Sciences does not show launch-readiness language
- airport classifier tightened so export terminals do not become airport projects

## Key sector examples checked

- Life Sciences / Biologics Manufacturing
- Hyperscale Data Centres
- Semiconductor Fabs
- Airports
- Rail / Metro
- Healthcare
- Defence / Secure Mission Infrastructure
- Energy / Utilities
- Water / Utilities
- Orbital AI Data Centres
- Lunar Infrastructure
- Mars Surface Infrastructure
- Spaceports
- Cislunar Propellant Infrastructure

## Regression test

Fast regression run:
- 2,000 Earth infrastructure prompts
- 1,000 Space infrastructure prompts
- 3,000 total
- XER spot-checks on core sector examples
- Classification, cost range, schedule, risk, sector drivers, benchmark comparison and narrative leakage checked

Results:
- Average score: **100 / 100**
- Minimum score: **100 / 100**
- Issue cases: **0**

## Important note

The regression used a fast Monte Carlo stub for volume testing. The packaged build itself still contains the normal CASEY QCRA/QSRA generation.

## Earth sample outputs

| Type | Asset | Subsector | P50 | Schedule | Risk | Score |
|---|---|---|---|---|---|---:|
| Earth | biologics and obesity therapeutics manufacturing campus | Life Sciences / Biologics Manufacturing | $5.8B | 73 months | Medium | 100 |
| Earth | 500MW hyperscale AI data centre campus | Digital Infrastructure / Hyperscale Data Centre | $6.7B | 64 months | Medium | 100 |
| Earth | semiconductor fab | Semiconductor / Advanced Manufacturing | $22.5B | 92 months | Medium | 100 |
| Earth | airport terminal expansion | Airport / Aviation | $12.6B | 96 months | Medium | 100 |
| Earth | metro extension | Rail / Transit | $7.4B | 90 months | Medium | 100 |
| Earth | SMR nuclear deployment | Nuclear / Energy | $15.9B | 125 months | Medium | 100 |
| Earth | green hydrogen production hub | Energy / Utilities | $6.5B | 71 months | Low | 100 |
| Earth | desalination plant | Energy / Utilities | $5.9B | 74 months | Low | 100 |
| Earth | acute hospital campus | Healthcare / Hospital | $4.6B | 77 months | Medium-High | 100 |
| Earth | secure satellite control centre | Defence / Secure Mission Infrastructure | $5.8B | 69 months | Medium | 100 |
| Earth | biologics and obesity therapeutics manufacturing campus | Life Sciences / Biologics Manufacturing | $5.8B | 73 months | Medium | 100 |
| Earth | 500MW hyperscale AI data centre campus | Digital Infrastructure / Hyperscale Data Centre | $6.7B | 64 months | Medium | 100 |
| Earth | semiconductor fab | Semiconductor / Advanced Manufacturing | $22.5B | 92 months | Medium | 100 |
| Earth | airport terminal expansion | Airport / Aviation | $12.6B | 96 months | Medium | 100 |
| Earth | metro extension | Rail / Transit | $7.4B | 90 months | Medium | 100 |

## Space sample outputs

| Type | Asset | Subsector | P50 | Schedule | Risk | Score |
|---|---|---|---|---|---|---:|
| Space | orbital AI data centre | Orbital Compute / Manufacturing | $53.5B | 156 months | Extreme | 100 |
| Space | space-based data centre constellation | Orbital Compute / Manufacturing | $49.9B | 138 months | Very High | 100 |
| Space | lunar logistics hub | Lunar Surface Habitat/Base | $70.9B | 151 months | Very High | 100 |
| Space | Mars surface research outpost | Mars Surface Habitat/Base | $138.6B | 200 months | Extreme | 100 |
| Space | cislunar propellant depot | ISRU/Mining/Propellant | $75.1B | 195 months | Extreme | 100 |
| Space | commercial spaceport expansion | Spaceport/Launch | $9.0B | 86 months | High | 100 |
| Space | orbital AI data centre | Orbital Compute / Manufacturing | $53.5B | 159 months | Extreme | 100 |
| Space | space-based data centre constellation | Orbital Compute / Manufacturing | $49.9B | 135 months | Very High | 100 |
| Space | lunar logistics hub | Lunar Surface Habitat/Base | $70.9B | 155 months | Extreme | 100 |
| Space | Mars surface research outpost | Mars Surface Habitat/Base | $138.6B | 197 months | Extreme | 100 |
| Space | cislunar propellant depot | ISRU/Mining/Propellant | $75.1B | 192 months | Extreme | 100 |
| Space | commercial spaceport expansion | Spaceport/Launch | $9.0B | 90 months | High | 100 |
| Space | orbital AI data centre | Orbital Compute / Manufacturing | $53.5B | 158 months | Extreme | 100 |
| Space | space-based data centre constellation | Orbital Compute / Manufacturing | $49.9B | 140 months | Extreme | 100 |
| Space | lunar logistics hub | Lunar Surface Habitat/Base | $70.9B | 152 months | Extreme | 100 |

## Launch recommendation

This is the strongest public-demo build so far. It is suitable for controlled public testing.

Recommended public positioning:

> CASEY generates first-pass project intelligence for Earth and orbital infrastructure across cost, schedule, risk, confidence and delivery strategy.

Do not call outputs certified estimates.
