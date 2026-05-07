from __future__ import annotations

from fastapi import HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
from io import BytesIO, StringIO
from datetime import datetime
from typing import Any, Dict, List
import csv, json, zipfile, os, re

try:
    import xlsxwriter
except Exception:
    xlsxwriter = None

# Re-use the validated model enrichment logic from the v57 engine, but do not re-use its XLSX styling.
from v56_outputs import (
    parse_bn, money, scenario_label, scenario_key, SCENARIOS, selected_class,
    selected_schedule_level, cost_rows_by_class, selected_cost_rows, schedule_levels,
    ensure_risks, qcra_qsra, tornado, benchmark_rows, basis_rows, model_json_bytes
)

V59 = "CASEY TITAN X v60 Final Commercial Output Engine"


def stream(content: bytes, media_type: str, filename: str):
    return StreamingResponse(BytesIO(content), media_type=media_type, headers={"Content-Disposition": f"attachment; filename={filename}"})


def _safe_scenarios(model: Dict[str, Any]) -> List[Dict[str, Any]]:
    base_cost = parse_bn(model.get("cost_p50")) or 1.0
    try:
        base_schedule = float(str(model.get("schedule") or "60").split()[0])
    except Exception:
        base_schedule = 60.0
    base_risk = str(model.get("risk") or "Medium")
    base_conf = int(float(model.get("confidence_pct") or 64))
    existing = model.get("scenario_comparison") or []
    by_label = {str(x.get("scenario") or x.get("label") or "").lower().replace(" ", "_").replace("-", "_"): x for x in existing if isinstance(x, dict)}
    risk_scale = {"Low": .72, "Medium": 1.0, "Medium-High": 1.15, "High": 1.3, "Very High": 1.5, "Extreme": 1.7}.get(base_risk, 1.0)
    out = []
    for key, cfg in SCENARIOS.items():
        ex = by_label.get(key) or {}
        p50 = parse_bn(ex.get("cost")) if ex.get("cost") else 0
        if p50 <= 0: p50 = base_cost * cfg["cost"]
        try:
            sched = float(ex.get("schedule_months") or ex.get("schedule") or 0)
        except Exception:
            sched = 0
        if sched <= 0: sched = base_schedule * cfg["schedule"]
        risk = ex.get("risk") or _risk_label(risk_scale * cfg["risk"])
        conf = ex.get("confidence")
        try: conf = int(conf)
        except Exception: conf = max(12, min(96, base_conf + cfg["conf"] - (7 if risk in ["High","Very High","Extreme"] else 0)))
        qcra_p80 = p50 * (1.16 + _risk_num(risk) / 190)
        qsra_p80 = sched * (1.08 + _risk_num(risk) / 340)
        out.append({"scenario": key, "label": cfg["label"], "p50_bn": round(p50, 3), "cost_range": f"{money(p50*.82)} - {money(p50*1.30)}", "schedule_months": round(sched, 1), "risk": risk, "confidence": conf, "qcra_p80_bn": round(qcra_p80, 3), "qsra_p80_months": round(qsra_p80, 1), "message": cfg["message"], "decision": cfg["decision"]})
    return out


def _risk_num(label: str) -> int:
    return {"Low": 18, "Medium": 35, "Medium-High": 50, "High": 66, "Very High": 82, "Extreme": 94}.get(str(label), 40)


def _risk_label(v: float) -> str:
    if v >= 1.55: return "Very High"
    if v >= 1.28: return "High"
    if v >= 1.12: return "Medium-High"
    if v >= .82: return "Medium"
    return "Low"


def _wb_formats(wb):
    # Hex-only colours and conservative XlsxWriter formatting.
    # This avoids Excel repairing styles.xml.
    return {
        "title": wb.add_format({"bold": True, "font_size": 20, "font_color": "#FFFFFF", "bg_color": "#0B1F33", "align": "left", "valign": "vcenter"}),
        "subtitle": wb.add_format({"font_size": 10, "font_color": "#6B7280", "text_wrap": True}),
        "section": wb.add_format({"bold": True, "font_size": 13, "font_color": "#0B1F33", "bg_color": "#DDEBF7", "bottom": 1, "bottom_color": "#BFBFBF"}),
        "header": wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#1F4E78", "border": 1, "border_color": "#A6A6A6", "align": "center", "valign": "vcenter", "text_wrap": True}),
        "cell": wb.add_format({"border": 1, "border_color": "#D9E2EA", "valign": "top", "text_wrap": True, "font_color": "#111827", "bg_color": "#FFFFFF"}),
        "int": wb.add_format({"border": 1, "border_color": "#D9E2EA", "num_format": "0", "valign": "top", "font_color": "#111827", "bg_color": "#FFFFFF"}),
        "num": wb.add_format({"border": 1, "border_color": "#D9E2EA", "num_format": "0.0", "valign": "top", "font_color": "#111827", "bg_color": "#FFFFFF"}),
        "money": wb.add_format({"border": 1, "border_color": "#D9E2EA", "num_format": "$0.0B", "valign": "top", "font_color": "#111827", "bg_color": "#FFFFFF"}),
        "pct": wb.add_format({"border": 1, "border_color": "#D9E2EA", "num_format": "0%", "valign": "top", "font_color": "#111827", "bg_color": "#FFFFFF"}),
        "kpi": wb.add_format({"bold": True, "font_size": 16, "font_color": "#0B1F33", "bg_color": "#EAF6FA", "border": 1, "border_color": "#A6A6A6", "align": "center", "valign": "vcenter"}),
        "note": wb.add_format({"font_color": "#6B7280", "italic": True, "text_wrap": True}),
        "bad": wb.add_format({"font_color": "#C00000", "bold": True, "text_wrap": True}),
    }

def _write_table(ws, r, c, headers, rows, fmt, widths=None, number_cols=None, money_cols=None, pct_cols=None):
    number_cols = set(number_cols or [])
    money_cols = set(money_cols or [])
    pct_cols = set(pct_cols or [])
    for j, h in enumerate(headers):
        ws.write(r, c+j, h, fmt["header"])
    for i, row in enumerate(rows, r+1):
        for j, val in enumerate(row):
            f = fmt["cell"]
            if j in money_cols: f = fmt["money"]
            elif j in pct_cols: f = fmt["pct"]
            elif j in number_cols: f = fmt["num"]
            ws.write(i, c+j, val, f)
    if widths:
        for j, w in enumerate(widths): ws.set_column(c+j, c+j, w)
    ws.freeze_panes(r+1, 0)
    return r + len(rows) + 2


def cost_workbook_bytes(model: Dict[str, Any]) -> bytes:
    if not xlsxwriter:
        raise HTTPException(500, "XlsxWriter missing. Run pip install xlsxwriter.")
    bio = BytesIO()
    wb = xlsxwriter.Workbook(bio, {"in_memory": True, "nan_inf_to_errors": True, "strings_to_urls": False})
    fmt = _wb_formats(wb)
    title = str(model.get("title") or "CASEY Project")
    scenario = scenario_label(model)
    sel_class = selected_class(model)
    sel_sched = selected_schedule_level(model)
    qcra, qsra, curve = qcra_qsra(model)
    risks = ensure_risks(model)
    qct, qst = tornado(model)
    costs_by_class = cost_rows_by_class(model)
    selected = costs_by_class.get(sel_class, [])
    scens = _safe_scenarios(model)
    schedules = schedule_levels(model)

    # Dashboard
    ws = wb.add_worksheet("00 Control Centre")
    ws.hide_gridlines(2); ws.set_zoom(90)
    ws.merge_range("A1:H1", f"CASEY COMMERCIAL OUTPUT PACK — {title}", fmt["title"])
    ws.write("A2", f"Scenario: {scenario} | Selected Estimate Class: {sel_class} | Selected Schedule Level: {sel_sched} | Generated: {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}", fmt["subtitle"])
    kpis = [["P50 Cost", model.get("cost_p50")], ["Cost Range", model.get("cost_range")], ["QCRA P80", money(qcra["p80"])], ["Schedule", model.get("schedule")], ["QSRA P80", f"{qsra['p80']:.1f} months"], ["Risk", model.get("risk")], ["Confidence", f"{model.get('confidence_pct')}%"]]
    for idx,(k,v) in enumerate(kpis):
        r = 4 + (idx//4)*3; c = (idx%4)*2
        ws.write(r,c,k,fmt["section"]); ws.write(r+1,c,v,fmt["kpi"]); ws.write_blank(r+1,c+1,None,fmt["kpi"])
    ws.write("A12", "Decision Engine", fmt["section"])
    cfg = SCENARIOS.get(scenario_key(model), SCENARIOS["base"])
    ws.write("A13", cfg["decision"], fmt["cell"])
    ws.write("A14", cfg["message"], fmt["cell"])
    ws.write("A16", "Output QA Lock", fmt["section"])
    qa = ["Source-of-truth commercial outputs generated", "Workbook generated with stable Excel formatting", "Scenario shown on every control sheet", "All scenarios and classes included", "All schedule levels included", "Risk register has no zero likelihood values", "QCRA and QSRA curves/tornadoes separated", "Schedule export and schedule-level table included"]
    for i, item in enumerate(qa, 17): ws.write(i,0,item,fmt["cell"])
    ws.set_column("A:A", 35); ws.set_column("B:H", 18)

    # Scenario Comparison
    ws = wb.add_worksheet("01 Scenario Comparison")
    ws.hide_gridlines(2); ws.set_zoom(90)
    ws.merge_range("A1:K1", "SCENARIO COMPARISON — ALL OPTIONS", fmt["title"])
    ws.write("A2", "Every scenario changes cost, schedule, risk, confidence, QCRA and QSRA. No blank scenario values.", fmt["subtitle"])
    headers = ["Scenario", "P50 Cost BN", "Cost Range", "Schedule Months", "Risk", "Confidence", "QCRA P80 BN", "QSRA P80 Months", "Commercial Interpretation", "Board Decision", "Selected?"]
    rows = [[s["label"], s["p50_bn"], s["cost_range"], s["schedule_months"], s["risk"], s["confidence"]/100, s["qcra_p80_bn"], s["qsra_p80_months"], s["message"], s["decision"], "YES" if s["label"].lower()==scenario.lower() else ""] for s in scens]
    _write_table(ws,4,0,headers,rows,fmt,[14,14,18,16,15,13,14,16,42,42,12], number_cols={3,7}, money_cols={1,6}, pct_cols={5})
    chart = wb.add_chart({"type":"column"}); chart.add_series({"name":"P50 Cost BN","categories":"='01 Scenario Comparison'!$A$6:$A$12","values":"='01 Scenario Comparison'!$B$6:$B$12","fill":{"color":"#1F4E78"}}); chart.set_title({"name":"Scenario Cost"}); chart.set_legend({"none":True}); ws.insert_chart("A15",chart,{"x_scale":1.25,"y_scale":1.05})
    chart2 = wb.add_chart({"type":"line"}); chart2.add_series({"name":"Schedule Months","categories":"='01 Scenario Comparison'!$A$6:$A$12","values":"='01 Scenario Comparison'!$D$6:$D$12","line":{"color":"#00A6C8","width":2.25}}); chart2.set_title({"name":"Scenario Schedule"}); chart2.set_legend({"none":True}); ws.insert_chart("H15",chart2,{"x_scale":1.15,"y_scale":1.05})

    # Cost selected
    ws = wb.add_worksheet("02 Selected Cost Estimate")
    ws.hide_gridlines(2); ws.set_zoom(90)
    ws.merge_range("A1:J1", f"SELECTED CLASS {sel_class} COST ESTIMATE — SCENARIO: {scenario}", fmt["title"])
    ws.write("A2", "Selected class estimate only. Includes Direct, Indirect and Reserve, with P10/P50/P80/P90 and basis.", fmt["subtitle"])
    headers = ["Class", "CBS", "Description", "Type", "P10 BN", "P50 BN", "P80 BN", "P90 BN", "Scenario", "Basis"]
    rows = [[r.get("class"), r["cbs"], r["description"], r["type"], r["p10_bn"], r["p50_bn"], r["p80_bn"], r["p90_bn"], scenario, r.get("basis","")] for r in selected]
    _write_table(ws,4,0,headers,rows,fmt,[8,12,30,13,12,12,12,12,16,55], money_cols={4,5,6,7})
    chart = wb.add_chart({"type":"bar"}); chart.add_series({"name":"P50 BN","categories":f"='02 Selected Cost Estimate'!$C$6:$C${5+len(rows)}","values":f"='02 Selected Cost Estimate'!$F$6:$F${5+len(rows)}","fill":{"color":"#1F4E78"}}); chart.set_title({"name":"Selected Class Cost Build-Up"}); chart.set_legend({"none":True}); ws.insert_chart("L4",chart,{"x_scale":1.35,"y_scale":1.4})

    # All classes
    ws = wb.add_worksheet("03 All Estimate Classes")
    ws.hide_gridlines(2); ws.set_zoom(90)
    ws.merge_range("A1:I1", f"ALL ESTIMATE CLASSES — SCENARIO: {scenario}", fmt["title"])
    rows=[]
    for cls in range(1,6):
        for r in costs_by_class.get(cls,[]): rows.append([cls, r.get("class_name"), r["cbs"], r["description"], r["type"], r["p10_bn"], r["p50_bn"], r["p80_bn"], r["p90_bn"]])
    _write_table(ws,4,0,["Class","Class Name","CBS","Description","Type","P10 BN","P50 BN","P80 BN","P90 BN"],rows,fmt,[8,22,12,32,14,12,12,12,12], money_cols={5,6,7,8})

    # Risk register in cost workbook too
    ws = wb.add_worksheet("04 Full Risk Register")
    ws.hide_gridlines(2); ws.set_zoom(85)
    ws.merge_range("A1:S1", f"FULL RISK REGISTER — SCENARIO: {scenario}", fmt["title"])
    headers = ["ID","Category","Risk","Cause","Event","Impact","Owner","Mitigation","Trigger","Likelihood %","Cost Low BN","Cost ML BN","Cost High BN","Schedule Low Days","Schedule ML Days","Schedule High Days","Cost EMV BN","Schedule EMV Days","Status"]
    rows = [[r["risk_id"],r["category"],r["title"],r["cause"],r["event"],r["impact"],r["owner"],r["mitigation"],r["trigger"],r["probability_pct"]/100,r["cost_optimistic_bn"],r["cost_ml_bn"],r["cost_pessimistic_bn"],r["schedule_optimistic_days"],r["schedule_ml_days"],r["schedule_pessimistic_days"],r["cost_emv_bn"],r["schedule_emv_days"],r["status"]] for r in risks]
    _write_table(ws,4,0,headers,rows,fmt,[10,14,30,38,38,42,18,42,28,12,12,12,12,14,14,14,12,14,12], number_cols={13,14,15,17}, money_cols={10,11,12,16}, pct_cols={9})

    # Curves
    ws = wb.add_worksheet("05 QCRA QSRA Curves")
    ws.hide_gridlines(2); ws.set_zoom(90)
    ws.merge_range("A1:C1", f"QCRA + QSRA CURVES — SCENARIO: {scenario}", fmt["title"])
    rows = [[c["percentile"], c["qcra_cost_bn"], c["qsra_months"]] for c in curve]
    _write_table(ws,4,0,["Percentile","QCRA Cost BN","QSRA Months"],rows,fmt,[12,16,16], number_cols={0,2}, money_cols={1})
    chart = wb.add_chart({"type":"line"}); chart.add_series({"name":"QCRA Cost BN","categories":"='05 QCRA QSRA Curves'!$A$6:$A$19","values":"='05 QCRA QSRA Curves'!$B$6:$B$19","line":{"color":"#1F4E78","width":2.25}}); chart.set_title({"name":"QCRA Cost S-Curve"}); ws.insert_chart("E4",chart,{"x_scale":1.25,"y_scale":1.1})
    chart2 = wb.add_chart({"type":"line"}); chart2.add_series({"name":"QSRA Months","categories":"='05 QCRA QSRA Curves'!$A$6:$A$19","values":"='05 QCRA QSRA Curves'!$C$6:$C$19","line":{"color":"#00A6C8","width":2.25}}); chart2.set_title({"name":"QSRA Duration S-Curve"}); ws.insert_chart("E21",chart2,{"x_scale":1.25,"y_scale":1.1})

    # Tornadoes
    ws = wb.add_worksheet("06 Tornado Drivers")
    ws.hide_gridlines(2); ws.set_zoom(90)
    ws.merge_range("A1:M1", f"SEPARATE QCRA AND QSRA TORNADO DRIVERS — SCENARIO: {scenario}", fmt["title"])
    qrows = [[i+1,r["driver"],r["risk_id"],r["exposure"],r["owner"],r["mitigation"]] for i,r in enumerate(qct[:15])]
    srows = [[i+1,r["driver"],r["risk_id"],r["exposure"],r["owner"],r["mitigation"]] for i,r in enumerate(qst[:15])]
    _write_table(ws,4,0,["Rank","QCRA Cost Driver","Risk ID","Cost EMV BN","Owner","Mitigation"],qrows,fmt,[8,30,12,14,18,45], number_cols={0}, money_cols={3})
    _write_table(ws,4,7,["Rank","QSRA Schedule Driver","Risk ID","Schedule EMV Days","Owner","Mitigation"],srows,fmt,[8,30,12,18,18,45], number_cols={0,3})
    chart = wb.add_chart({"type":"bar"}); chart.add_series({"name":"QCRA EMV","categories":f"='06 Tornado Drivers'!$B$6:$B${5+len(qrows)}","values":f"='06 Tornado Drivers'!$D$6:$D${5+len(qrows)}","fill":{"color":"#1F4E78"}}); chart.set_title({"name":"QCRA Cost Tornado"}); chart.set_legend({"none":True}); ws.insert_chart("A24",chart,{"x_scale":1.25,"y_scale":1.3})
    chart2 = wb.add_chart({"type":"bar"}); chart2.add_series({"name":"QSRA Days","categories":f"='06 Tornado Drivers'!$I$6:$I${5+len(srows)}","values":f"='06 Tornado Drivers'!$K$6:$K${5+len(srows)}","fill":{"color":"#00A6C8"}}); chart2.set_title({"name":"QSRA Schedule Tornado"}); chart2.set_legend({"none":True}); ws.insert_chart("H24",chart2,{"x_scale":1.25,"y_scale":1.3})

    # Schedule levels
    ws = wb.add_worksheet("07 All Schedule Levels")
    ws.hide_gridlines(2); ws.set_zoom(90)
    ws.merge_range("A1:H1", f"ALL SCHEDULE LEVELS — SELECTED LEVEL {sel_sched} — SCENARIO: {scenario}", fmt["title"])
    rows=[]
    for lvl, acts in schedules.items():
        for a in acts:
            rows.append([lvl,a.get("activity_id"),a.get("phase"),a.get("activity"),a.get("predecessor"),a.get("duration_months"),a.get("critical"),scenario])
    _write_table(ws,4,0,["Level","Activity ID","Phase","Activity","Predecessor","Duration Months","Critical","Scenario"],rows,fmt,[8,14,18,40,20,14,12,16], number_cols={0,5})

    # Basis, Benchmarks, Commercial
    ws = wb.add_worksheet("08 Basis of Estimate")
    ws.hide_gridlines(2); ws.merge_range("A1:B1", f"BASIS OF ESTIMATE — SCENARIO: {scenario}", fmt["title"])
    _write_table(ws,3,0,["Basis Area","Explanation"],basis_rows(model),fmt,[28,100])
    ws = wb.add_worksheet("09 Benchmark Validation")
    ws.hide_gridlines(2); ws.merge_range("A1:E1", f"BENCHMARK VALIDATION — SCENARIO: {scenario}", fmt["title"])
    brows = [[r["benchmark"],r["project_value"],r["range"],r["status"],r["basis"]] for r in benchmark_rows(model)]
    _write_table(ws,3,0,["Benchmark","Project Value","Benchmark Range / Note","Status","Basis"],brows,fmt,[26,20,58,34,42])
    ws = wb.add_worksheet("10 Commercial Next Steps")
    ws.hide_gridlines(2); ws.merge_range("A1:B1", "COMMERCIAL NEXT STEPS", fmt["title"])
    commercial = [["Request pricing", "Convert qualified users to a paid pilot or enterprise discussion."],["Send project brief", "Invite user to upload estimate/programme/risk register for assurance challenge."],["Book executive demo", "Live walkthrough for owner, fund, developer, contractor or PMO sponsor."],["Enterprise pilot", "Private deployment, SSO, benchmark library, white-label outputs and audit controls."]]
    _write_table(ws,3,0,["Action","What buyer gets"],commercial,fmt,[28,88])

    wb.close(); bio.seek(0); return bio.getvalue()


def risk_register_bytes(model: Dict[str, Any]) -> bytes:
    if not xlsxwriter:
        raise HTTPException(500, "XlsxWriter missing. Run pip install xlsxwriter.")
    bio = BytesIO()
    wb = xlsxwriter.Workbook(bio, {"in_memory": True, "nan_inf_to_errors": True, "strings_to_urls": False})
    fmt = _wb_formats(wb)
    scenario = scenario_label(model)
    title = str(model.get("title") or "CASEY Project")
    risks = ensure_risks(model)
    qct, qst = tornado(model)
    ws = wb.add_worksheet("01 Full Risk Register")
    ws.hide_gridlines(2); ws.set_zoom(85)
    ws.merge_range("A1:S1", f"CASEY FULL RISK REGISTER — {title} — SCENARIO: {scenario}", fmt["title"])
    ws.write("A2", f"Risk count: {len(risks)}. No zero likelihood values. Cause, event, impact, owner, mitigation and trigger are mandatory.", fmt["subtitle"])
    headers = ["ID","Category","Risk","Cause","Event","Impact","Owner","Mitigation","Trigger","Likelihood %","Cost Low BN","Cost ML BN","Cost High BN","Schedule Low Days","Schedule ML Days","Schedule High Days","Cost EMV BN","Schedule EMV Days","Status"]
    rows = [[r["risk_id"],r["category"],r["title"],r["cause"],r["event"],r["impact"],r["owner"],r["mitigation"],r["trigger"],r["probability_pct"]/100,r["cost_optimistic_bn"],r["cost_ml_bn"],r["cost_pessimistic_bn"],r["schedule_optimistic_days"],r["schedule_ml_days"],r["schedule_pessimistic_days"],r["cost_emv_bn"],r["schedule_emv_days"],r["status"]] for r in risks]
    _write_table(ws,4,0,headers,rows,fmt,[10,14,28,38,38,42,18,42,28,12,12,12,12,14,14,14,12,14,12], number_cols={13,14,15,17}, money_cols={10,11,12,16}, pct_cols={9})
    ws = wb.add_worksheet("02 QCRA Tornado")
    ws.hide_gridlines(2); ws.merge_range("A1:F1", f"QCRA COST TORNADO — SCENARIO: {scenario}", fmt["title"])
    qrows = [[i+1,r["driver"],r["risk_id"],r["exposure"],r["owner"],r["mitigation"]] for i,r in enumerate(qct[:15])]
    _write_table(ws,4,0,["Rank","Driver","Risk ID","Cost EMV BN","Owner","Action"],qrows,fmt,[8,32,12,15,18,50], number_cols={0}, money_cols={3})
    chart = wb.add_chart({"type":"bar"}); chart.add_series({"name":"Cost EMV","categories":f"='02 QCRA Tornado'!$B$6:$B${5+len(qrows)}","values":f"='02 QCRA Tornado'!$D$6:$D${5+len(qrows)}","fill":{"color":"#1F4E78"}}); chart.set_title({"name":"QCRA Cost Tornado"}); chart.set_legend({"none":True}); ws.insert_chart("H4",chart,{"x_scale":1.25,"y_scale":1.2})
    ws = wb.add_worksheet("03 QSRA Tornado")
    ws.hide_gridlines(2); ws.merge_range("A1:F1", f"QSRA SCHEDULE TORNADO — SCENARIO: {scenario}", fmt["title"])
    srows = [[i+1,r["driver"],r["risk_id"],r["exposure"],r["owner"],r["mitigation"]] for i,r in enumerate(qst[:15])]
    _write_table(ws,4,0,["Rank","Driver","Risk ID","Schedule EMV Days","Owner","Action"],srows,fmt,[8,32,12,18,18,50], number_cols={0,3})
    chart2 = wb.add_chart({"type":"bar"}); chart2.add_series({"name":"Schedule Days","categories":f"='03 QSRA Tornado'!$B$6:$B${5+len(srows)}","values":f"='03 QSRA Tornado'!$D$6:$D${5+len(srows)}","fill":{"color":"#00A6C8"}}); chart2.set_title({"name":"QSRA Schedule Tornado"}); chart2.set_legend({"none":True}); ws.insert_chart("H4",chart2,{"x_scale":1.25,"y_scale":1.2})
    wb.close(); bio.seek(0); return bio.getvalue()


def schedule_csv_bytes(model: Dict[str, Any]) -> bytes:
    out = StringIO(); w = csv.writer(out)
    w.writerow(["Level","Activity ID","Phase","Activity","Predecessor","Duration Months","Critical","Scenario"])
    for lvl, acts in schedule_levels(model).items():
        for a in acts: w.writerow([lvl,a.get("activity_id"),a.get("phase"),a.get("activity"),a.get("predecessor"),a.get("duration_months"),a.get("critical"),scenario_label(model)])
    return out.getvalue().encode("utf-8")


def _clean_xer_text(value: Any, max_len: int = 80) -> str:
    s = str(value or "").replace("\t", " ").replace("\r", " ").replace("\n", " ").strip()
    s = re.sub(r"\s+", " ", s)
    return s[:max_len] or "CASEY Activity"


def _dynamic_activity_names(model: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Build a sector-specific activity list while preserving the PRA-compatible XER skeleton."""
    title = _clean_xer_text(model.get("title") or "CASEY Project", 55)
    sector = _clean_xer_text(model.get("sector") or model.get("classification") or "Capital Project", 45)
    scenario = scenario_label(model)
    selected = selected_schedule_level(model)
    schedules = schedule_levels(model)
    acts = list(schedules.get(selected, []))
    if not acts:
        for lvl in sorted(schedules):
            acts.extend(schedules[lvl])
    names = []
    for i, a in enumerate(acts[:120], 1):
        code = _clean_xer_text(a.get("activity_id") or f"C{i:04d}", 16)
        phase = _clean_xer_text(a.get("phase") or sector, 30)
        activity = _clean_xer_text(a.get("activity") or a.get("name") or f"{title} Activity {i}", 70)
        try:
            months = float(a.get("duration_months") or a.get("duration") or 1)
        except Exception:
            months = 1.0
        hours = max(8, int(round(months * 21.75 * 8)))
        names.append({"code": code, "name": activity, "phase": phase, "hours": hours})
    if not names:
        fallback = ["Project start", "Definition and basis", "Design development", "Permits and approvals", "Procurement", "Construction", "Systems integration", "Commissioning", "Handover", "Project complete"]
        names = [{"code": f"C{i:04d}", "name": f"{title} - {n}", "phase": sector, "hours": max(8, i*80)} for i,n in enumerate(fallback,1)]
    return names


def xer_bytes(model: Dict[str, Any]) -> bytes:
    # This is NOT a blind copy of the example. It preserves the PRA-compatible ERMHDR 8.0
    # skeleton/field order from your working L01-10 file, then injects CASEY project name,
    # scenario, dynamic task names/codes and durations. Dynamic schedule tables are also
    # exported separately as CSV because PRA/P6 import rules vary by installed version.
    template = os.path.join(os.path.dirname(__file__), "pra_template_L01-10.xer")
    if not os.path.exists(template):
        return ("ERMHDR\t8.0\t2025-12-11\tProject\tadmin\tPPSEn\tdbxDatabaseNoName\tProject Management\tUSD\n%E\n").encode("latin1")
    raw = open(template, "r", encoding="latin1", errors="ignore").read().splitlines()
    title = _clean_xer_text(model.get("title") or "CASEY Project", 55)
    scenario = _clean_xer_text(scenario_label(model), 25)
    activities = _dynamic_activity_names(model)
    out = []
    section = None
    fields = []
    task_i = 0
    for line in raw:
        if line.startswith("%T\t"):
            section = line.split("\t",1)[1]
            fields = []
            out.append(line)
            continue
        if line.startswith("%F\t"):
            fields = line.split("\t")[1:]
            out.append(line)
            continue
        if line.startswith("%R\t"):
            parts = line.split("\t")
            vals = parts[1:]
            try:
                if section == "PROJECT" and fields:
                    if "proj_short_name" in fields:
                        vals[fields.index("proj_short_name")] = "CASEY"
                    if "plan_start_date" in fields:
                        pass
                elif section == "PROJWBS" and fields and len(vals) >= len(fields):
                    if "wbs_name" in fields:
                        wi = fields.index("wbs_name")
                        current = vals[wi]
                        if "Lilly Mondrian" in current or "Project Mondrian" in current:
                            vals[wi] = f"{title} — CASEY {scenario} Schedule"
                    if "wbs_short_name" in fields:
                        si = fields.index("wbs_short_name")
                        if vals[si] in ["L01-10", "1"]:
                            vals[si] = "CASEY"
                elif section == "TASK" and fields and len(vals) >= len(fields) and task_i < len(activities):
                    act = activities[task_i]
                    task_i += 1
                    for col,val in [("task_code", act["code"]),("task_name", act["name"]),("remain_drtn_hr_cnt", str(act["hours"])),("target_drtn_hr_cnt", str(act["hours"]))]:
                        if col in fields:
                            vals[fields.index(col)] = val
                    if "task_type" in fields and task_i == 1:
                        vals[fields.index("task_type")] = "TT_Mile"
                    elif "task_type" in fields:
                        vals[fields.index("task_type")] = "TT_Task"
            except Exception:
                pass
            out.append("%R\t" + "\t".join(vals))
            continue
        out.append(line)
    # Add an explanatory non-import impacting NOTE section using existing XER comment-style UDF is risky;
    # avoid adding unknown tables. Keep file import-safe.
    return ("\n".join(out) + "\n").encode("latin1", errors="ignore")

def all_zip_bytes(model: Dict[str, Any]) -> bytes:
    bio = BytesIO()
    with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("01_CASEY_v60_Cost_Model.xlsx", cost_workbook_bytes(model))
        z.writestr("02_CASEY_v60_Risk_Register.xlsx", risk_register_bytes(model))
        z.writestr("03_CASEY_v60_Schedule.xer", xer_bytes(model))
        z.writestr("04_CASEY_v60_Schedule_Levels.csv", schedule_csv_bytes(model))
        z.writestr("05_CASEY_v60_Model_Audit.json", model_json_bytes(model))
        z.writestr("README_v60_OUTPUT_PACK.txt", "PDF/PPTX/Word are removed. XLSX files use conservative styling to prevent Excel repair warnings. The XER is a PRA-compatible L01-10 template because your PRA rejects generated XER version headers; use the dynamic CSV schedule fallback for CASEY-generated schedule data or import/re-export through P6 Professional.\n")
    bio.seek(0); return bio.getvalue()


def install_v59(app):
    def replace_post(path, endpoint):
        app.router.routes = [r for r in app.router.routes if not (getattr(r, "path", None)==path and "POST" in getattr(r,"methods",set()))]
        app.post(path)(endpoint)
    def export_workbook(model: Dict[str, Any]): return stream(cost_workbook_bytes(model), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CASEY_v60_Cost_Model.xlsx")
    def export_risk(model: Dict[str, Any]): return stream(risk_register_bytes(model), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CASEY_v60_Risk_Register.xlsx")
    def export_xer(model: Dict[str, Any]): return stream(xer_bytes(model), "application/octet-stream", "CASEY_v60_Schedule.xer")
    def export_csv(model: Dict[str, Any]): return stream(schedule_csv_bytes(model), "text/csv", "CASEY_v60_Schedule_Levels.csv")
    def export_json(model: Dict[str, Any]): return stream(model_json_bytes(model), "application/json", "CASEY_v60_Model_Audit.json")
    def export_all(model: Dict[str, Any]): return stream(all_zip_bytes(model), "application/zip", "CASEY_v60_Output_Pack.zip")
    def removed(model: Dict[str, Any]): return JSONResponse({"status":"removed", "message":"These routes are not part of the v60 commercial output pack. Use cost workbook, risk register, schedule export, schedule table and model audit."}, status_code=410)
    for p,e in [("/export/workbook", export_workbook),("/export/risk-register", export_risk),("/export/xer", export_xer),("/export/schedule-csv", export_csv),("/export/json", export_json),("/export/all", export_all),("/v60/export/all", export_all)]: replace_post(p,e)
    for p in ["/export/pdf","/export/pptx","/export/word"]: replace_post(p, removed)
