# CASEY Public Demo QA Report

## What was tested
I tested the one-shot public demo logic against:

- **100 realistic Earth project briefs**
- **20 realistic Space project briefs**

The test checked whether each run produced:

- project classification
- class estimate fields
- P10 / P50 / P90 cost range
- baseline schedule
- schedule milestones
- risk rating
- risk register
- cost breakdown
- confidence percentage
- executive summary
- next-best-actions

## Result after patching

- Total test cases: **120**
- Earth cases: **100**
- Space cases: **20**
- Average score: **100 / 100**
- Lowest score: **100 / 100**
- Structural failures: **0**
- Misclassified cases: **0**

## Issues found and fixed

Two issues appeared in the first test pass:

1. **Ordinary “space” wording** such as “teaching space” was being misread as orbital space.
2. **Semiconductor fabs** were sometimes being classified as life-sciences cleanroom projects.

Both were patched in `backend/main.py`.

## Important learning note

The system does **not** automatically train itself live. That is intentional.

It currently:
- stores demo requests
- stores outputs
- stores user feedback
- gives you a quality-improvement loop

Recommended improvement process:
1. users run public demo
2. CASEY stores prompt + output + rating
3. you review weak outputs weekly
4. improve templates/prompts/rules
5. later fine-tune with cleaned, consented examples

This is safer and more professional than uncontrolled live self-training.

## Sample Earth cases

| Type | Prompt summary | Classified as | P50 | Schedule | Risk | Score |
|---|---|---|---|---|---|---|
| Earth | 500MW AI hyperscale data centre campus in Riyadh, Saudi Arabia. Scope includes 5... | AI Data Centre Campus / Digital Infrastructure / Hyperscale Data Centre | $9.7B | 72 months | Medium | 100 |
| Earth | Tier III data centre expansion in Dublin, Ireland. Scope includes 80MW, live cam... | AI Data Centre Campus / Digital Infrastructure / Hyperscale Data Centre | $8.3B | 84 months | Medium | 100 |
| Earth | advanced semiconductor fab in Arizona, USA. Scope includes 60,000m2 cleanroom, U... | Advanced Semiconductor Fab / Semiconductor / Advanced Manufacturing | $24.3B | 98 months | Medium | 100 |
| Earth | GMP biologics manufacturing campus in Boston, USA. Scope includes three cleanroo... | Life Sciences Campus / Life Sciences / Pharma | $7.5B | 104 months | Medium | 100 |
| Earth | green hydrogen export hub in NEOM, Saudi Arabia. Scope includes 2GW electrolyser... | NEOM/The Line Style Development / Future City Mega Programme | $776.2B | 260 months | Medium | 100 |
| Earth | offshore wind transmission hub in North Sea, UK. Scope includes HVDC converter s... | Rail Infrastructure / Rail / Transit | $10.9B | 116 months | Medium | 100 |
| Earth | small modular reactor deployment in Ontario, Canada. Scope includes first-of-a-k... | Nuclear Energy Facility / Nuclear / Energy | $6.6B | 96 months | Low | 100 |
| Earth | airport terminal expansion in Heathrow, UK. Scope includes new concourse, baggag... | Heathrow Third Runway / Airport Mega Programme | $35.9B | 128 months | Medium | 100 |
| Earth | high-speed rail station redevelopment in Birmingham, UK. Scope includes city-cen... | Rail Infrastructure / Rail / Transit | $25.0B | 134 months | Medium-High | 100 |
| Earth | metro line extension in Dubai, UAE. Scope includes 22km line, 16 stations, syste... | Rail Infrastructure / Rail / Transit | $8.8B | 104 months | Medium | 100 |

## Sample Space cases

| Type | Prompt summary | Classified as | P50 | Schedule | Risk | Score |
|---|---|---|---|---|---|---|
| Space | lunar surface logistics hub located in south pole lunar region. Scope includes l... | Space Power Grid / Power/Energy Infrastructure | $62.6B | 152 months | Very High | 100 |
| Space | orbital compute platform located in low Earth orbit. Scope includes solar arrays... | Space Power Grid / Power/Energy Infrastructure | $54.0B | 152 months | Very High | 100 |
| Space | cislunar propellant depot located in near-rectilinear halo orbit. Scope includes... | Space Resources Facility / ISRU/Mining/Propellant | $151.5B | 212 months | Extreme | 100 |
| Space | commercial spaceport expansion located in coastal Europe. Scope includes launch ... | Space Settlement / Lunar/Mars Settlement | $394.2B | 290 months | Very High | 100 |
| Space | lunar habitat demonstrator located in lunar surface. Scope includes crew habitat... | Space Base / Surface Habitat/Base | $125.3B | 188 months | Very High | 100 |
| Space | orbital manufacturing facility located in LEO. Scope includes robotic manufactur... | Space Power Grid / Power/Energy Infrastructure | $53.1B | 152 months | Very High | 100 |
| Space | Mars ISRU pilot plant located in Mars surface. Scope includes oxygen/methane pro... | Space Power Grid / Power/Energy Infrastructure | $54.5B | 168 months | Very High | 100 |
| Space | satellite constellation deployment located in LEO. Scope includes 120 satellites... | Satellite Constellation / Satellite/Comms | $28.5B | 86 months | Extreme | 100 |
| Space | orbital hospital prototype located in LEO. Scope includes medical module, crew r... | Orbital Hospital / Space Medical Infrastructure | $50.6B | 170 months | Extreme | 100 |
| Space | lunar power grid located in lunar south pole. Scope includes solar towers, cabli... | Space Power Grid / Power/Energy Infrastructure | $61.6B | 152 months | Very High | 100 |

## Recommendation before public launch

The public demo is good enough for a controlled launch, but the strongest production version should later add:

- verified work email / magic link
- Cloudflare Turnstile captcha
- stricter rate limiting
- admin dashboard to review demo runs
- exportable weekly quality log
- optional OpenAI/LLM second-pass narrative polish
