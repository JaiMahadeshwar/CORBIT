# CASEY V123 Live Calibration Demo Layer

This build adds a subtle institutional live-calibration layer so users feel that CASEY is not operating on static demo assumptions.

## What appears in the app
- LIVE CALIBRATION SIGNALS ACTIVE strip under the confidence engine badge
- Delivery Environment Calibration panel in the Overview
- Sector-specific calibration signals in Mission Control / Uncertainty Narrative
- Export model/audit fields: `live_calibration_active`, `live_calibration_signals`, `live_calibration_summary`

## What the demo should say
CASEY continuously recalibrates confidence, contingency and delivery exposure against current sector conditions, benchmark behaviour and live operating signals.

## Important positioning
This is a demo calibration layer and strategic intelligence signal, not a certified market-data feed. For production, connect Bloomberg/Reuters/ENR/FAA/NASA/ERP/P6 feeds through governed data connectors.

## QA
See `CASEY_V123_LIVE_CALIBRATION_QA_REPORT.json`.
