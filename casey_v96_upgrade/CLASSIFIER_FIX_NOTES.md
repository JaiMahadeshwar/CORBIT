# CASEY Earth/Space Classifier Fix

This build fixes misclassification caused by ambiguous terms like:
- commercial launch demand
- product launch
- platform
- mission operations
- satellite control centre

New logic:
- Earth locations such as North Carolina, Arizona, Texas, UK, Riyadh, Dubai etc strongly support Earth classification.
- Strong Earth sectors such as GMP, FDA, CQV, biologics, pharma, data centre, semiconductor, airport, hospital, defence etc override ambiguous single space words.
- Space requires stronger signals such as orbital, LEO, lunar, Mars, cislunar, spaceport, rocket, launch vehicle, satellite constellation etc.
- "Commercial launch demand" is treated as product launch language unless rocket/spaceport/orbital context exists.

Use this ZIP for the next local test.