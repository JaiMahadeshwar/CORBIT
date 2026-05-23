# CASEY Final Sector Behaviour Build QA

## Added in this build

This build adds the final public-demo polish requested:

- Sector signature behaviours for Earth and Space sectors
- Executive shock insight per run
- CASEY Confidence Engine badge
- Sector-specific behavioural curves
- More human schedule basis language
- Less robotic QSRA/P80 display values
- Stronger contradiction logic inside scenario narratives
- Scenario views preserve executive insight
- Life sciences, data centres, semiconductor, airports, rail, healthcare, defence, energy, water, orbital, lunar, Mars, spaceport, satellite/comms and ISRU/resource behaviours

## Key behavioural examples

Life Sciences:
- mechanical completion is not the finish line
- CQV, GMP turnover, media fills and validated production readiness dominate

Data Centres:
- energisation and commissioning concurrency dominate shell construction

Semiconductor:
- process tool install and qualification dominate cleanroom shell delivery

Airport:
- ORAT and live operations dominate physical completion

Rail:
- possessions, utilities and systems assurance dominate visible civil progress

Defence:
- accreditation and mission-system acceptance become hidden critical paths

Space:
- launch cadence, TRL, qualification, thermal/power and autonomous recovery dominate

## QA sample

| Example | Mode | Subsector | P50 | Schedule | P80 | Risk | Shock insight | Pass |
|---|---|---|---|---|---:|---|---|---|
| life_sciences | Earth | Life Sciences / Biologics Manufacturing | $6.9B | 73 months | 84 | Medium-High | Mechanical completion is unlikely to be the true finish line; validated production readiness and deviation closure are t... | Yes |
| data_centre | Earth | Digital Infrastructure / Hyperscale Data Centre | $7.2B | 60 months | 73 | Medium | The dominant delivery constraint is likely energisation and commissioning concurrency, not shell construction productivi... | Yes |
| semiconductor | Earth | Semiconductor / Advanced Manufacturing | $21.2B | 92 months | 104 | Medium | The critical path is likely tool install and qualification cadence rather than the cleanroom shell alone.... | Yes |
| airport | Earth | Airport / Aviation | $21.9B | 131 months | 147 | Medium-High | The programme may finish construction before the airport is operationally ready to absorb the change.... | Yes |
| rail | Earth | Rail / Transit | $8.3B | 99 months | 111 | Medium | Possessions, utility diversions and systems assurance are likely to dominate the critical path more than visible civil p... | Yes |
| defence | Earth | Defence / Secure Mission Infrastructure | $5.5B | 69 months | 79 | Medium | Security accreditation and mission-system acceptance can become the hidden critical path even after the facility is buil... | Yes |
| energy | Earth | Energy / Utilities | $6.5B | 72 months | 83 | Low | The approval, grid/process interface and commissioning evidence may drive confidence more than equipment installation.... | Yes |
| orbital | Space | Orbital Compute / Manufacturing | $53.5B | 159 months | 179 | Extreme | The decisive constraint is not launch alone; it is thermal-power balance, autonomous servicing and recoverability after ... | Yes |
| mars | Space | Mars Surface Habitat/Base | $138.6B | 230 months | 255 | Extreme | The programme is governed by launch windows, autonomy and life-support reliability; recovery options are extremely limit... | Yes |

## Notes

The system cannot guarantee perfect results for every possible user input, but this build is significantly more robust for broad Earth and Space infrastructure examples and still blocks low-quality / nonsense prompts.
