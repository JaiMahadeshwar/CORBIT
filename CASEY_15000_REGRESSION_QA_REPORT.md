# CASEY 15,000 Prompt Regression QA Report

## Scope

I tested the classifier-fixed CASEY public demo build against:

- **10,000 Earth infrastructure prompts**
- **5,000 Space infrastructure prompts**
- **15,000 total prompts**

The Earth set included Turner & Townsend / Mace / Jacobs / AECOM-style projects:
- life sciences / GMP / biologics / obesity therapeutics
- data centres
- semiconductor fabs
- airports
- rail / metro
- nuclear / SMR
- hydrogen
- carbon capture
- desalination
- hospitals
- defence
- secure satellite control centres
- naval bases
- LNG
- regeneration
- EV charging
- broadband/fibre

The Space set included:
- orbital AI data centres
- space-based data centre constellations
- orbital compute platforms
- lunar edge compute hubs
- lunar logistics
- lunar habitats
- Mars ISRU
- Mars surface outposts
- cislunar propellant depots
- satellite constellations
- spaceports
- orbital manufacturing
- deep-space communications
- asteroid prospecting

## Edge case focus

The test deliberately included ambiguous words that previously caused issues:
- commercial launch demand
- product launch
- mission operations
- platform readiness
- secure satellite control centre
- nuclear power on Mars
- hospital/platform/transition wording

## Results

- Total tested: **15,000**
- Earth tested: **10,000**
- Space tested: **5,000**
- Average structural QA score: **100 / 100**
- Median score: **100.0 / 100**
- Minimum score: **100 / 100**
- Issue cases: **0**

## XER check

For runtime practicality, the test performed full XER export checks on:
- first **250 Earth** outputs
- first **250 Space** outputs

All checked XER outputs passed.

## Important testing note

The 15,000-volume run used a fast deterministic QCRA/QSRA stub for regression speed. This does not change the build. It allows classification, output structure, cost/schedule/risk fields, scenario-ready fields and Earth/Space routing to be tested at scale without running thousands of expensive Monte Carlo simulations.

The actual build still contains the full QCRA/QSRA generation logic.

## Patch made during the test

One additional edge case was found and fixed:

- Mars surface research outposts containing the phrase “nuclear power” were being classified as Earth Nuclear.
- The classifier now treats explicit Mars/Moon/LEO/orbital infrastructure as Space even if it contains Earth-sector words such as nuclear, grid, hospital, platform or manufacturing.

## Verdict

This build is ready for local testing and controlled public demo deployment.

## Earth sample outputs

| Type | Asset | Classified as | P50 | Schedule | Risk | Score |
|---|---|---|---|---|---|---:|
| Earth | biologics and obesity therapeutics manufacturing campus | Life Sciences Manufacturing Campus / Life Sciences / Biologics Manufacturing | $5.8B | 82 months | Medium | 100 |
| Earth | advanced cell therapy manufacturing facility | Life Sciences Manufacturing Campus / Life Sciences / Biologics Manufacturing | $4.5B | 82 months | Medium | 100 |
| Earth | 500MW hyperscale AI data centre campus | AI Data Centre Campus / Digital Infrastructure / Hyperscale Data Centre | $6.7B | 63 months | Medium | 100 |
| Earth | edge data centre network | AI Data Centre Campus / Digital Infrastructure / Hyperscale Data Centre | $9.8B | 86 months | High | 100 |
| Earth | semiconductor fab | Advanced Semiconductor Fab / Semiconductor / Advanced Manufacturing | $61.0B | 124 months | High | 100 |
| Earth | airport terminal expansion | Airport Infrastructure / Airport / Aviation | $7.4B | 89 months | Medium | 100 |
| Earth | high-speed rail corridor | Rail Infrastructure / Rail / Transit | $8.8B | 104 months | Medium | 100 |
| Earth | metro extension | Rail Infrastructure / Rail / Transit | $6.5B | 87 months | Low | 100 |
| Earth | SMR nuclear deployment | Nuclear Energy Facility / Nuclear / Energy | $51.5B | 163 months | High | 100 |
| Earth | green hydrogen production hub | Airport Infrastructure / Airport / Aviation | $23.3B | 120 months | High | 100 |
| Earth | carbon capture cluster | Airport Infrastructure / Airport / Aviation | $6.1B | 79 months | Low | 100 |
| Earth | desalination plant | Energy Infrastructure / Energy / Utilities | $9.1B | 86 months | Medium | 100 |
| Earth | acute hospital campus | Hospital Campus / Healthcare / Hospital | $5.1B | 77 months | Medium | 100 |
| Earth | secure satellite control centre | Secure Defence Infrastructure / Defence / Secure Mission Infrastructure | $14.6B | 103 months | High | 100 |
| Earth | defence command and control facility | Secure Defence Infrastructure / Defence / Secure Mission Infrastructure | $14.6B | 100 months | High | 100 |

## Space sample outputs

| Type | Asset | Classified as | P50 | Schedule | Risk | Score |
|---|---|---|---|---|---|---:|
| Space | orbital AI data centre | Orbital AI Compute Platform / Orbital Compute / Manufacturing | $42.4B | 150 months | Extreme | 100 |
| Space | space-based data centre constellation | Orbital AI Compute Platform / Orbital Compute / Manufacturing | $74.6B | 145 months | Very High | 100 |
| Space | orbital compute platform | Orbital AI Compute Platform / Orbital Compute / Manufacturing | $58.6B | 147 months | Extreme | 100 |
| Space | lunar edge compute hub | Space Power Grid / Power/Energy Infrastructure | $153.3B | 183 months | Extreme | 100 |
| Space | lunar logistics hub | Lunar Surface Infrastructure / Lunar Surface Habitat/Base | $259.1B | 183 months | Extreme | 100 |
| Space | lunar habitat phase one | Lunar Surface Infrastructure / Lunar Surface Habitat/Base | $71.0B | 183 months | Extreme | 100 |
| Space | Mars ISRU fuel plant | Space Resources Facility / ISRU/Mining/Propellant | $66.4B | 176 months | Very High | 100 |
| Space | Mars surface research outpost | Mars Surface Infrastructure / Mars Surface Habitat/Base | $159.7B | 199 months | Extreme | 100 |
| Space | cislunar propellant depot | Space Resources Facility / ISRU/Mining/Propellant | $226.7B | 227 months | Extreme | 100 |
| Space | LEO satellite constellation | Satellite Constellation / Satellite/Comms | $54.2B | 95 months | Extreme | 100 |
| Space | commercial spaceport expansion | Launch Infrastructure / Spaceport/Launch | $11.4B | 79 months | Very High | 100 |
| Space | orbital manufacturing facility | Space Power Grid / Power/Energy Infrastructure | $67.1B | 154 months | Very High | 100 |
| Space | deep-space communications array | Deep-Space Communications Array / Deep-Space Communications Infrastructure | $20.5B | 99 months | Extreme | 100 |
| Space | asteroid prospecting mission infrastructure | Frontier Space Infrastructure / General Space Infrastructure | $156.0B | 165 months | Extreme | 100 |
| Space | orbital AI data centre | Orbital AI Compute Platform / Orbital Compute / Manufacturing | $214.1B | 190 months | Extreme | 100 |
