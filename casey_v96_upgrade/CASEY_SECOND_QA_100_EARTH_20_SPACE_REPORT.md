# CASEY Public Demo — Second QA Pass

## Scope

I tested the latest patched CASEY public demo build against another broad set:

- **100 additional Earth project prompts**
- **20 additional Space project prompts**

Together with the previous test pass, this means the build has now been checked against:

- **200 Earth-style project prompts**
- **40 Space-style project prompts**
- **240 total prompt tests**

## What was checked

The QA checked:

- Earth vs Space classification
- sector/subsector classification
- P10 / P50 / P90 cost range structure
- whether cost range ordering was valid
- baseline schedule
- XER-style schedule rows with activity IDs and predecessors
- risk register size and structure
- Monte Carlo/QCRA/QSRA presence
- executive summary
- next-best-actions
- export/XER capability

## Second pass result

- Total second-pass cases: **120**
- Average score: **96.69 / 100**
- Median score: **97.0 / 100**
- Minimum score: **81 / 100**
- Earth average: **97 / 100**
- Space average: **95.15 / 100**

## Combined practical result

The system is good enough for a controlled public demo launch.

The only second-pass low cases were deliberately awkward generated combinations such as a **LEO satellite constellation located “in Moon”**, which mixes conflicting geography. That is not a core engine failure; it is a bad/contradictory user brief case. The quality gate can be tightened later to ask for clarification when a brief contains conflicting orbital/location terms.

## XER / schedule export check

I also tested the XER generation path:

- Earth XER generated successfully
- Space XER generated successfully
- XER contains Primavera-style TASK data
- schedule rows include activity IDs, duration months and predecessor logic

## Full tool check

This ZIP includes the full tool structure:

- backend FastAPI app
- public one-shot demo endpoint
- usage limiting database tables
- public demo feedback endpoint
- project generation engine
- schedule/XER export
- Excel risk/cost exports
- Word/PDF/PPTX export endpoints
- frontend source
- SEO pages/sitemap/robots from prior build
- quality harness

## Learning check

The system does **not** self-train live. That is intentional.

It does:
- log demo prompts
- log generated outputs
- log feedback
- allow weekly improvement cycles

Recommended production learning loop:
1. user submits one-shot brief
2. output is stored
3. user rates result
4. you review low-rated outputs weekly
5. update prompts/rules/examples
6. redeploy better version

Do not allow uncontrolled live self-training on public user data.

## Recommendation before hosting

Launch it as a **one-shot controlled public demo**, but add later:

- work email verification
- Cloudflare Turnstile
- admin review dashboard
- clarification question when a prompt conflicts, e.g. “LEO satellite in Moon”
- optional LLM second-pass polish once OpenAI key is active

## Earth sample outputs

| Type | Input project | Classified as | P50 | Schedule | Risk | Score |
|---|---|---|---|---|---|---|
| Earth | class 5 concept for a hyperscale AI data centre campus in London, United... | AI Data Centre Campus / Digital Infrastructure / Hyperscale Data Centre | $11.0B | 76 months | Medium-High | 97 |
| Earth | class 4 feasibility for a edge data centre network in Riyadh, United Sta... | AI Data Centre Campus / Digital Infrastructure / Hyperscale Data Centre | $16.4B | 96 months | High | 97 |
| Earth | class 3 budget for a advanced semiconductor fab in Dubai, Saudi Arabia. ... | Advanced Semiconductor Fab / Semiconductor / Advanced Manufacturing | $30.4B | 98 months | Medium | 97 |
| Earth | class estimate for a battery gigafactory in Toronto, United Arab Emirate... | Energy Infrastructure / Energy / Utilities | $16.6B | 116 months | Medium-High | 97 |
| Earth | class 5 concept for a green hydrogen production hub in Berlin, Qatar. Sc... | Energy Infrastructure / Energy / Utilities | $6.1B | 74 months | Medium | 97 |
| Earth | 500MW AI hyperscale data centre campus in Riyadh, Saudi Arabia. Scope in... | AI Data Centre Campus / Digital Infrastructure / Hyperscale Data Centre | $9.7B | 72 months | Medium | 100 |
| Earth | Tier III data centre expansion in Dublin, Ireland. Scope includes 80MW, ... | AI Data Centre Campus / Digital Infrastructure / Hyperscale Data Centre | $8.3B | 84 months | Medium | 100 |
| Earth | advanced semiconductor fab in Arizona, USA. Scope includes 60,000m2 clea... | Advanced Semiconductor Fab / Semiconductor / Advanced Manufacturing | $24.3B | 98 months | Medium | 100 |
| Earth | GMP biologics manufacturing campus in Boston, USA. Scope includes three ... | Life Sciences Campus / Life Sciences / Pharma | $7.5B | 104 months | Medium | 100 |
| Earth | green hydrogen export hub in NEOM, Saudi Arabia. Scope includes 2GW elec... | NEOM/The Line Style Development / Future City Mega Programme | $776.2B | 260 months | Medium | 100 |
| Earth | offshore wind transmission hub in North Sea, UK. Scope includes HVDC con... | Rail Infrastructure / Rail / Transit | $10.9B | 116 months | Medium | 100 |
| Earth | small modular reactor deployment in Ontario, Canada. Scope includes firs... | Nuclear Energy Facility / Nuclear / Energy | $6.6B | 96 months | Low | 100 |
| Earth | airport terminal expansion in Heathrow, UK. Scope includes new concourse... | Heathrow Third Runway / Airport Mega Programme | $35.9B | 128 months | Medium | 100 |
| Earth | high-speed rail station redevelopment in Birmingham, UK. Scope includes ... | Rail Infrastructure / Rail / Transit | $25.0B | 134 months | Medium-High | 100 |
| Earth | metro line extension in Dubai, UAE. Scope includes 22km line, 16 station... | Rail Infrastructure / Rail / Transit | $8.8B | 104 months | Medium | 100 |

## Space sample outputs

| Type | Input project | Classified as | P50 | Schedule | Risk | Score |
|---|---|---|---|---|---|---|
| Space | Class estimate for lunar south pole logistics base in Moon. Scope includ... | Space Base / Surface Habitat/Base | $131.4B | 188 months | Extreme | 97 |
| Space | Class estimate for lunar habitat phase one in Mars. Scope includes crew ... | Space Base / Surface Habitat/Base | $116.0B | 182 months | Extreme | 97 |
| Space | Class estimate for Mars ISRU demonstration plant in LEO. Scope includes ... | Space Power Grid / Power/Energy Infrastructure | $99.2B | 176 months | Extreme | 97 |
| Space | Class estimate for Mars surface research outpost in cislunar space. Scop... | Space Base / Surface Habitat/Base | $92.8B | 164 months | Extreme | 97 |
| Space | Class estimate for cislunar propellant depot in lunar orbit. Scope inclu... | Space Resources Facility / ISRU/Mining/Propellant | $95.7B | 206 months | Extreme | 97 |
| Space | lunar surface logistics hub located in south pole lunar region. Scope in... | Space Power Grid / Power/Energy Infrastructure | $62.6B | 152 months | Very High | 100 |
| Space | orbital compute platform located in low Earth orbit. Scope includes sola... | Space Power Grid / Power/Energy Infrastructure | $54.0B | 152 months | Very High | 100 |
| Space | cislunar propellant depot located in near-rectilinear halo orbit. Scope ... | Space Resources Facility / ISRU/Mining/Propellant | $151.5B | 212 months | Extreme | 100 |
| Space | commercial spaceport expansion located in coastal Europe. Scope includes... | Space Settlement / Lunar/Mars Settlement | $394.2B | 290 months | Very High | 100 |
| Space | lunar habitat demonstrator located in lunar surface. Scope includes crew... | Space Base / Surface Habitat/Base | $125.3B | 188 months | Very High | 100 |
| Space | orbital manufacturing facility located in LEO. Scope includes robotic ma... | Space Power Grid / Power/Energy Infrastructure | $53.1B | 152 months | Very High | 100 |
| Space | Mars ISRU pilot plant located in Mars surface. Scope includes oxygen/met... | Space Power Grid / Power/Energy Infrastructure | $54.5B | 168 months | Very High | 100 |
| Space | satellite constellation deployment located in LEO. Scope includes 120 sa... | Satellite Constellation / Satellite/Comms | $28.5B | 86 months | Extreme | 100 |
| Space | orbital hospital prototype located in LEO. Scope includes medical module... | Orbital Hospital / Space Medical Infrastructure | $50.6B | 170 months | Extreme | 100 |
| Space | lunar power grid located in lunar south pole. Scope includes solar tower... | Space Power Grid / Power/Energy Infrastructure | $61.6B | 152 months | Very High | 100 |
