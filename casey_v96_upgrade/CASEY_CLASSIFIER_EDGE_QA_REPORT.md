# CASEY Classifier Fix QA Report

Fixed the Earth/Space classifier so Earth place names and strong Earth-sector signals override ambiguous words like product launch, mission, platform, payload, and satellite control centre.

## Key fix

The life sciences prompt with North Carolina + GMP + FDA + CQV + commercial launch demand now classifies as:

- Earth
- Life Sciences / Biologics Manufacturing

## Edge cases tested

| Case | Expected | Actual | Title | Subsector | Pass |
|---|---|---|---|---|---|
| Life sciences commercial launch demand | Earth | Earth | Life Sciences Manufacturing Campus | Life Sciences / Biologics Manufacturing | Yes |
| Pharma product launch | Earth | Earth | Life Sciences Manufacturing Campus | Life Sciences / Biologics Manufacturing | Yes |
| Airport launch retail | Earth | Earth | Airport Infrastructure | Airport / Aviation | Yes |
| Defence mission operations | Earth | Earth | Secure Defence Infrastructure | Defence / Secure Mission Infrastructure | Yes |
| Data platform launch | Earth | Earth | AI Data Centre Campus | Digital Infrastructure / Hyperscale Data Centre | Yes |
| Semiconductor launch | Earth | Earth | Advanced Semiconductor Fab | Semiconductor / Advanced Manufacturing | Yes |
| Hospital digital platform | Earth | Earth | Hospital Campus | Healthcare / Hospital | Yes |
| Orbital AI data centre | Space | Space | Orbital AI Compute Platform | Orbital Compute / Manufacturing | Yes |
| Spaceport launch complex | Space | Space | Launch Infrastructure | Spaceport/Launch | Yes |
| Lunar base | Space | Space | Space Base | Surface Habitat/Base | Yes |
| Mars ISRU | Space | Space | Space Resources Facility | ISRU/Mining/Propellant | Yes |
| Satellite constellation | Space | Space | Satellite Constellation | Satellite/Comms | Yes |
| Cislunar depot | Space | Space | Space Resources Facility | ISRU/Mining/Propellant | Yes |

Result: 13/13 edge cases passed.
