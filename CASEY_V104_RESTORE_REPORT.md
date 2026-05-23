# CASEY V104 Executive Scenario Graphs Restored

Built from the V103 ZIP uploaded by the user.

Restored / fixed:
- Top Exposure Drivers / tornado chart field mapping (`driver` + `contribution`)
- QCRA/QSRA P50/P80/P90 curve readouts
- Scenario delta intelligence card values
- Scenario-vs-base comparison object
- Cost waterfall vs base
- Schedule waterfall vs base
- Confidence breakdown
- Board memo / gained / gave-up narrative
- Scenario selector / export strip retained from V103
- Local demo one-run limiter bypass retained

Backend QA:
- build_model runs for Base, Faster, Cheaper, Lower Risk and Premium.
- QCRA P50 equals headline cost.
- QSRA P50 equals headline schedule.
- Tornado chart data contains driver/contribution keys.
- Scenario delta card data contains label/value/meaning keys.
