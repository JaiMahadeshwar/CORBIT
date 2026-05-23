# CASEY V134 Full Sector Render / Export QA Report

## Purpose
Fix the rail and life-sciences render failures that produced the recovery screen, and harden the same failure path across every Earth and Space sector.

## Fixes applied
- Added text-safe render normalization for all arrays that can be rendered in React.
- Added object-as-child protection for board briefing, decisions, memo outputs, contradiction intelligence and calibration signals.
- Added emergency recovered dashboard that shows the generated model instead of a blank/recovery dead end if any component path fails.
- Added sector-complete backend normalisation for cost rows, schedule rows and risk rows.
- Added numeric-safe conversion for export rows where values arrive as strings like `$8.4B`.
- Added nuclear-specific ontology fallback, because nuclear was routing but falling through to a generic label.
- Re-smoke-tested exports across all selected Earth and Space sector families.

## Sectors smoke-tested
- Rail / Transit
- Life Sciences / Biologics Manufacturing
- Airport / Aviation
- Digital Infrastructure / Hyperscale Data Centre
- Space / Mission Assurance
- Oil & Gas / Process Infrastructure
- Healthcare / Hospital Infrastructure
- Water / Environmental Infrastructure
- Ports / Marine Infrastructure
- Nuclear / Regulated Generation
- Semiconductor / Advanced Manufacturing
- Defence / Secure Infrastructure
- Energy / Power Infrastructure
- Roads / Highways Infrastructure
- Mining / Metals Infrastructure

## Export smoke tests
For each sector above, the following endpoints returned successfully:
- Cost workbook
- Risk register
- XER
- QCRA/QSRA
- JSON audit model
- Full board pack ZIP

## Result
- Generation failures: 0
- Required model field failures: 0
- Export smoke failures: 0
- Checked leakage failures: 0
- Frontend production build: passed

## Important note
This build addresses the specific rail and Lilly/life-sciences failure class by preventing incomplete or mixed-shape payloads from reaching unsafe React render paths.
