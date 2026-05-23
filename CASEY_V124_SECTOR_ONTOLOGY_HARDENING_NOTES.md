# CASEY V124 Sector Ontology Hardening

This build adds a final sector-lock layer across Earth and Space before public demo.

## Fixed
- Removed cross-sector causal leakage such as data-centre cooling chains appearing in airport/rail examples.
- Added sector-locked causal graph nodes for: airport, rail, hyperscale data centre, semiconductor, life sciences, defence, oil & gas, energy/power, space, healthcare, water and ports/marine.
- Rewrites mission control signals, benchmark comparison, confidence drivers, cost drivers, schedule threats, uncertainty narrative, board briefing and CASEY thinking from the selected sector ontology.
- Updated frontend causal graph so it reads `model.causal_graph_nodes` instead of hardcoded data-centre defaults.
- Changed loading wording from launch-specific language to neutral delivery-tail language.

## Tested sample prompts
- Heathrow airport terminal expansion: airport-native ORAT/baggage/airside chain.
- California rail: possessions/signalling/systems migration chain.
- Microsoft AI campus: grid/cooling/IST chain retained where valid.
- Lunar infrastructure: launch/payload/mission assurance chain retained where valid.
- Defence secure facility: accreditation/mission-system/assurance chain.
- LNG/oil and gas: FEED/long-lead/module/tie-in/start-up chain.
- Energy/power: permitting/grid/equipment/energisation chain.
- Semiconductor fab: cleanroom/UPW/tool/yield-ramp chain.
- Healthcare: clinical/equipment/infection-control/phased occupancy chain.
- Water/desalination: permits/MEICA/process commissioning/compliance chain.

## Release note
This is still a first-pass strategic intelligence demo, not a certified estimate or live market-data product.
