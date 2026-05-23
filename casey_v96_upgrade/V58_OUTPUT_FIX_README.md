# CASEY v58 Platinum Reliable Outputs

This build fixes the output issues identified in v57:

- PDF/PPTX/Word are removed from the primary pack.
- Cost model and risk register are rebuilt with conservative XLSX formatting to avoid Excel repair/style warnings.
- Scenario comparison has populated P50, range, schedule, risk, confidence, QCRA P80 and QSRA P80 values for every scenario.
- Cost model always shows selected scenario and selected class.
- All estimate classes are included.
- All schedule levels are included in the cost workbook and dynamic CSV.
- Full risk register is first in the risk workbook and contains all generated risks, not just a top 10 view.
- No zero likelihood risks; probability is clamped to at least 5%.
- QCRA and QSRA curves and tornado drivers are separate.
- XER export uses the PRA-compatible L01-10 template you provided. Dynamic schedule data is also exported as CSV because PRA rejects many generated/new-version XER files.

Primary outputs:

1. CASEY_v58_Platinum_Cost_Model.xlsx
2. CASEY_v58_Platinum_Risk_Register.xlsx
3. CASEY_v58_PRA_Compatible_Template_L01-10.xer
4. CASEY_v58_All_Schedule_Levels_DYNAMIC.csv
5. CASEY_v58_Model_Audit.json

If PRA still rejects schedule import, use the dynamic CSV and re-export through Primavera P6 Professional in the PRA-supported version.
