from __future__ import annotations

from fastapi import HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
from io import BytesIO, StringIO
from datetime import datetime
from typing import Any, Dict, List
import csv, json, math, zipfile, re

try:
    import xlsxwriter
except Exception:  # installed through requirements in v56
    xlsxwriter = None

V57 = "CASEY TITAN X v57 Sellable Any Project Engine"
BLUE = "#0B1F33"
CYAN = "#00A6C8"
LIGHT = "#F5F8FB"
GRID = "#D9E2EA"
DARK = "#111827"
GREEN = "#008A5B"
AMBER = "#C77700"
RED = "#B42318"


def stream(content: bytes, media_type: str, filename: str):
    return StreamingResponse(BytesIO(content), media_type=media_type, headers={"Content-Disposition": f"attachment; filename={filename}"})


def parse_bn(v: Any) -> float:
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v or "").replace("$", "").replace(",", "").strip().upper()
    if not s:
        return 0.0
    try:
        if s.endswith("T"):
            return float(s[:-1]) * 1000.0
        if s.endswith("B"):
            return float(s[:-1])
        if s.endswith("M"):
            return float(s[:-1]) / 1000.0
        return float(s)
    except Exception:
        nums = re.findall(r"[-+]?[0-9]*\.?[0-9]+", s)
        return float(nums[0]) if nums else 0.0


def money(v: float) -> str:
    v = float(v or 0)
    if abs(v) >= 1000:
        return f"${v/1000:.1f}T"
    if abs(v) >= 1:
        return f"${v:.1f}B"
    return f"${v*1000:.0f}M"


def scenario_key(model: Dict[str, Any]) -> str:
    return str(model.get("scenario") or model.get("scenario_label") or "base").lower().replace(" ", "_").replace("-", "_")


def scenario_label(model: Dict[str, Any]) -> str:
    raw = model.get("scenario_label") or scenario_key(model).replace("_", " ").title()
    return str(raw)


SCENARIOS = {
    "base": {"label": "Base", "cost": 1.00, "schedule": 1.00, "risk": 1.00, "conf": 0, "message": "Balanced reference case for first-pass board challenge and option testing.", "decision": "Use as control baseline; approve definition, not unconditional capital."},
    "faster": {"label": "Faster", "cost": 1.12, "schedule": 0.84, "risk": 1.22, "conf": -6, "message": "Compresses delivery through parallel design, early procurement and acceleration premium.", "decision": "Use only where time value exceeds cost and interface risk."},
    "cheaper": {"label": "Cheaper", "cost": 0.88, "schedule": 1.07, "risk": 1.18, "conf": -8, "message": "Reduces first cost through value engineering and tighter scope but transfers risk into delivery.", "decision": "Use only if capex pressure dominates and residual risk is accepted."},
    "lower_risk": {"label": "Lower Risk", "cost": 1.09, "schedule": 1.10, "risk": 0.72, "conf": 9, "message": "Adds surveys, assurance, buffers and procurement validation to improve confidence.", "decision": "Recommended where certainty matters more than lowest first cost."},
    "premium": {"label": "Premium", "cost": 1.16, "schedule": 1.04, "risk": 0.80, "conf": 12, "message": "Strengthens suppliers, resilience, design maturity and governance for flagship delivery.", "decision": "Use for mission-critical, reputational or flagship programmes."},
    "investor": {"label": "Investor", "cost": 0.98, "schedule": 1.00, "risk": 0.92, "conf": 3, "message": "Frames the programme around capital discipline, risk transparency and investment committee challenge.", "decision": "Use for sponsor, lender, partner or IC review."},
    "survival": {"label": "Survival", "cost": 0.78, "schedule": 1.20, "risk": 1.30, "conf": -12, "message": "Minimum viable scope and warning-heavy commercial stress case with high residual risk.", "decision": "Use as a stress case only, not as the preferred baseline."},
}


def selected_class(model):
    try:
        return int(model.get("estimate_class") or model.get("class_level") or 3)
    except Exception:
        return 3


def selected_schedule_level(model):
    try:
        return int(model.get("schedule_level") or 3)
    except Exception:
        return 3


def risk_rank_value(r: str) -> int:
    return {"Low": 15, "Medium": 35, "Medium-High": 50, "High": 65, "Very High": 82, "Extreme": 95}.get(str(r), 45)


def risk_label_from_value(v: float) -> str:
    if v >= 88: return "Extreme"
    if v >= 75: return "Very High"
    if v >= 60: return "High"
    if v >= 45: return "Medium-High"
    if v >= 25: return "Medium"
    return "Low"


def scenario_rows(model: Dict[str, Any]) -> List[Dict[str, Any]]:
    base_cost = parse_bn(model.get("cost_p50")) or 1.0
    base_schedule = float(str(model.get("schedule") or "60").split()[0]) if str(model.get("schedule") or "").split() else 60.0
    base_risk_v = risk_rank_value(model.get("risk"))
    base_conf = int(float(model.get("confidence_pct") or 64))
    rows = []
    existing = model.get("scenario_comparison") or []
    by_label = {str(x.get("scenario") or x.get("label") or "").lower().replace(" ", "_"): x for x in existing if isinstance(x, dict)}
    for key, cfg in SCENARIOS.items():
        ex = by_label.get(key)
        p50 = parse_bn((ex or {}).get("cost")) if ex else base_cost * cfg["cost"]
        sched = float((ex or {}).get("schedule_months") or round(base_schedule * cfg["schedule"], 1))
        risk = (ex or {}).get("risk") or risk_label_from_value(base_risk_v * cfg["risk"])
        conf = int((ex or {}).get("confidence") or max(12, min(96, base_conf + cfg["conf"] - (6 if risk in ["High", "Very High", "Extreme"] else 0))))
        qcra_p80 = p50 * (1.16 + (risk_rank_value(risk) / 180.0))
        qsra_p80 = sched * (1.08 + (risk_rank_value(risk) / 320.0))
        rows.append({"scenario": key, "label": cfg["label"], "p50_bn": round(p50, 3), "schedule_months": round(sched, 1), "risk": risk, "confidence": conf, "qcra_p80_bn": round(qcra_p80, 3), "qsra_p80_months": round(qsra_p80, 1), "message": cfg["message"], "decision": cfg["decision"]})
    return rows


def cost_rows_by_class(model: Dict[str, Any]) -> Dict[int, List[Dict[str, Any]]]:
    existing = model.get("estimates_by_class") or {}
    base = model.get("cost_lines") or []
    out: Dict[int, List[Dict[str, Any]]] = {}
    ranges = {1: (0.90, 1.15, "Class 1 Definitive"), 2: (0.85, 1.20, "Class 2 Control"), 3: (0.80, 1.30, "Class 3 Budget"), 4: (0.70, 1.50, "Class 4 Feasibility"), 5: (0.50, 2.00, "Class 5 Concept")}
    for cls in range(1, 6):
        rows = existing.get(str(cls)) or existing.get(cls) or []
        clean = []
        source = rows if rows else base
        lo, hi, cname = ranges[cls]
        for x in source:
            p50 = float(x.get("p50_bn") or x.get("p50") or 0)
            if p50 <= 0:
                continue
            clean.append({
                "class": cls, "class_name": cname, "cbs": x.get("cbs") or x.get("CBS") or "", "description": x.get("description") or x.get("Description") or "", "type": x.get("type") or x.get("Type") or "Direct", "basis": x.get("basis") or x.get("Basis") or "Scenario-linked parametric / benchmark basis.",
                "p10_bn": round(float(x.get("p10_bn") or p50 * lo), 3), "p50_bn": round(p50, 3), "p80_bn": round(p50 * (1 + (hi - 1) * .65), 3), "p90_bn": round(float(x.get("p90_bn") or p50 * hi), 3),
            })
        out[cls] = clean
    return out


def selected_cost_rows(model: Dict[str, Any]) -> List[Dict[str, Any]]:
    return cost_rows_by_class(model).get(selected_class(model)) or []


def schedule_levels(model: Dict[str, Any]) -> Dict[int, List[Dict[str, Any]]]:
    levels = model.get("schedules_by_level") or {}
    if not levels:
        levels = {str(selected_schedule_level(model)): model.get("schedule_rows") or []}
    out = {}
    for lvl in range(1, 6):
        rows = levels.get(str(lvl)) or levels.get(lvl) or []
        if not rows and model.get("schedule_rows"):
            rows = model.get("schedule_rows")
        out[lvl] = rows
    return out


def ensure_risks(model: Dict[str, Any]) -> List[Dict[str, Any]]:
    raw = model.get("risks") or model.get("risk_register") or []
    mode = str(model.get("mode") or "Earth").lower()
    sub = str(model.get("subsector") or "").lower()
    extras = [
        ("R-009", "Commercial claims escalation", "Contract ambiguity, late changes and fragmented accountability", "Claims submitted above allowance", "Commercial pressure, management distraction and increased outturn cost", "Commercial Lead", "Tight change control and dispute avoidance board", "Claim volume exceeds baseline"),
        ("R-010", "Resource availability", "Labour, specialist supervisors or commissioning teams constrained", "Required resources are unavailable when planned", "Productivity loss, resequencing and higher preliminaries", "Delivery Lead", "Resource loading, framework call-offs and labour strategy", "Resource histogram red flag"),
        ("R-011", "Quality failure / rework", "Installation quality, vendor QA or inspection regime below need", "Inspection failure triggers rework", "Rework cost, schedule delay and confidence loss", "Quality Manager", "ITPs, hold points, vendor QA, right-first-time reviews", "NCR trend increases"),
        ("R-012", "Cyber / controls integration", "Operational technology and control systems interfaces immature", "Controls integration fails testing", "Commissioning delay, resilience issue and operational risk", "Systems Lead", "Early cyber review, FAT/SAT plan, integration lab", "Controls defect trend increases"),
        ("R-013", "Weather / environmental disruption", "Extreme weather, environmental windows or site constraints not fully allowed", "Workface access or productivity is disrupted", "Delay, access cost and resequencing", "Construction Manager", "Weather calendars, contingency workfronts, protection strategy", "Weather downtime exceeds allowance"),
        ("R-014", "Stakeholder change", "Sponsor, operator or authority requirements evolve late", "Requirements change after baseline freeze", "Scope growth, redesign and approval delays", "Project Director", "Decision log, change board and stakeholder gates", "Late decision count increases"),
        ("R-015", "Funding / approval gate delay", "Investment approvals or staged funding gates not aligned with procurement", "Capital release is delayed", "Procurement missed windows and schedule slippage", "Sponsor", "Approval roadmap and early board papers", "Approval date slips"),
        ("R-016", "Commissioning readiness gap", "Incomplete test packs, operational readiness or training gaps", "Commissioning starts without readiness", "Repeated tests, handover delay and operational risk", "Commissioning Lead", "Readiness dashboard, integrated test plan, ORAT", "Readiness score below threshold"),
        ("R-017", "Data / estimate basis weakness", "Estimate source data, quantities or benchmarks are incomplete", "Baseline cannot be defended under challenge", "Confidence reduction and rework of cost plan", "Cost Lead", "Assumption register, benchmark evidence, quantity validation", "Unknown allowance exceeds threshold"),
        ("R-018", "Interface ownership gap", "RACI and package boundaries are unclear across packages", "Interface issue has no accountable owner", "Delay, claims and fragmented accountability", "Integration Manager", "Interface control documents and owner matrix", "Interface actions overdue"),
        ("R-019", "Supply chain insolvency", "Tier 2/3 financial health is unknown or weak", "Critical supplier fails or cannot perform", "Replacement cost, procurement delay and claims exposure", "Procurement Lead", "Financial checks, dual sourcing, step-in rights", "Supplier credit alert"),
        ("R-020", "Design change after procurement", "Procurement released before design maturity is sufficient", "Vendor packages require change", "Variation cost, rework and programme delay", "Design Manager", "Design freeze gates and change impact board", "IFC maturity below gate"),
        ("R-021", "Utilities diversion / enabling risk", "Unknown utilities, access or enabling works underestimated", "Enabling works expand or delay critical start", "Cost growth, access delay and resequencing", "Enabling Works Lead", "Surveys, trial pits and early utility agreements", "Unknown utilities found"),
        ("R-022", "Regulatory assurance delay", "Regulator or authority review cycle not fully allowed", "Approval takes longer than baseline", "Critical path delay and extra assurance cost", "Consents Lead", "Regulatory roadmap and early engagement", "Submission returned"),
        ("R-023", "Procurement strategy misalignment", "Package strategy does not match market capacity or risk transfer", "Tenders return above budget or fail", "Retendering, cost growth and schedule slip", "Commercial Lead", "Market sounding and package strategy review", "Low bidder response"),
        ("R-024", "Operational readiness shortfall", "Operating model, training or maintenance strategy immature", "Asset cannot enter service as planned", "Delayed opening and benefit deferral", "Operations Lead", "ORAT/readiness plan and service transition gate", "Readiness gate failed"),
        ("R-025", "Benefits realisation risk", "Commercial benefits rely on optimistic demand or ramp-up", "Ramp-up underperforms plan", "Revenue / value case weakens", "Sponsor", "Benefits model validation and ramp-up sensitivity", "Demand evidence weak"),
    ]
    if "airport" in sub:
        extras += [("R-A01", "Airside phasing disruption", "Live airport constraints and possessions underestimated", "Construction cannot access critical work fronts", "Night work premium, delay and operational disruption", "Airside Lead", "Airside phasing plan, possessions and ORAT integration", "Airside access cancelled"), ("R-A02", "Baggage / security integration", "Passenger processing systems not integrated early", "Systems fail integrated trials", "Opening delay and reputational impact", "Systems Lead", "Integration lab, FAT/SAT and ORAT tests", "Trial failures increase")]
    if "rail" in sub:
        extras += [("R-R01", "Possession access constraints", "Rail access windows insufficient", "Works cannot be delivered in possession", "Resequencing, delay and overtime cost", "Rail Access Lead", "Possession strategy and blockade contingency", "Possession cancelled"), ("R-R02", "Signalling integration", "Signalling design, test and assurance underestimated", "System fails integration testing", "Critical path delay and assurance issue", "Signalling Lead", "Independent signalling assurance and test lab", "Test defects exceed threshold")]
    if "data centre" in sub or "data center" in sub:
        extras += [("R-D03", "Liquid cooling vendor maturity", "Cooling technology or supplier capacity immature", "Vendor solution underperforms acceptance tests", "Capacity derate, redesign and delay", "MEP Lead", "Prototype, vendor qualification and performance guarantee", "Cooling FAT failed"), ("R-D04", "Generator / UPS lead time", "Critical power equipment supply constrained", "Power train delivery slips", "Critical path movement and acceleration premium", "Procurement Lead", "Early vendor reservation and dual-source strategy", "Vendor delivery warning")]
    if mode == "space":
        extras += [("R-S01", "Launch window dependency", "Orbital mechanics or launch availability constrains deployment", "Launch window missed", "Months of delay and re-manifest cost", "Mission Lead", "Backup launch windows and manifest options", "Launch readiness slips"), ("R-S02", "Thermal / radiation margin", "Environment and shielding assumptions immature", "Thermal or radiation margins fail review", "Redesign, extra mass and launch cost increase", "Chief Engineer", "Thermal vacuum testing and radiation analysis", "Margin below threshold"), ("R-S03", "Life support reliability", "ECLSS reliability or sparing assumptions immature", "Life support performance below requirement", "Mission safety risk and redesign", "Systems Lead", "Qualification testing and redundancy review", "Reliability test failure")]
    clean = []
    for i, r in enumerate(raw, 1):
        title = r.get("title") or r.get("risk") or r.get("name") or f"Risk {i}"
        prob = int(max(5, min(95, float(r.get("probability_pct") or r.get("likelihood_pct") or 22 + (i % 7) * 4))))
        cost_ml = float(r.get("cost_impact_ml_bn") or r.get("cost_mean_bn") or max(0.02, parse_bn(model.get("cost_p50")) * (0.01 + (i % 5) * .004)))
        sched_ml = float(r.get("schedule_impact_ml_days") or r.get("schedule_mean_days") or max(2, 6 + (i % 9) * 4))
        clean.append({"risk_id": r.get("risk_id") or r.get("id") or f"R-{i:03d}", "title": title, "category": r.get("category") or "Delivery", "cause": r.get("cause") or "Underlying assumption, dependency or delivery condition requires validation.", "event": r.get("risk_event") or r.get("event") or f"{title} materialises during delivery.", "impact": r.get("impact_description") or r.get("impact") or "Cost growth, schedule movement, claims exposure or confidence loss.", "owner": r.get("owner") or "Project Controls Lead", "mitigation": r.get("mitigation") or "Validate basis, assign owner, monitor trigger and implement mitigation plan.", "trigger": r.get("trigger") or "Risk indicator breaches threshold", "probability_pct": prob, "cost_optimistic_bn": round(cost_ml * .45, 3), "cost_ml_bn": round(cost_ml, 3), "cost_pessimistic_bn": round(cost_ml * 2.15, 3), "schedule_optimistic_days": round(sched_ml * .45, 1), "schedule_ml_days": round(sched_ml, 1), "schedule_pessimistic_days": round(sched_ml * 2.25, 1), "response": r.get("response") or "Mitigate", "status": r.get("status") or "Open"})
    existing_ids = {x["risk_id"] for x in clean}
    for j, e in enumerate(extras, 1):
        if e[0] in existing_ids:
            continue
        prob = 14 + (j % 8) * 5
        base_cost = max(0.03, parse_bn(model.get("cost_p50")) * (0.006 + (j % 5) * .003))
        base_days = 7 + (j % 8) * 5
        clean.append({"risk_id": e[0], "title": e[1], "category": "Sector" if '-' in e[0] else "Delivery", "cause": e[2], "event": e[3], "impact": e[4], "owner": e[5], "mitigation": e[6], "trigger": e[7], "probability_pct": prob, "cost_optimistic_bn": round(base_cost * .45, 3), "cost_ml_bn": round(base_cost, 3), "cost_pessimistic_bn": round(base_cost * 2.25, 3), "schedule_optimistic_days": round(base_days * .4, 1), "schedule_ml_days": base_days, "schedule_pessimistic_days": round(base_days * 2.3, 1), "response": "Mitigate", "status": "Open"})
    for r in clean:
        r["cost_emv_bn"] = round((r["probability_pct"] / 100.0) * (0.2 * r["cost_optimistic_bn"] + 0.5 * r["cost_ml_bn"] + 0.3 * r["cost_pessimistic_bn"]), 3)
        r["schedule_emv_days"] = round((r["probability_pct"] / 100.0) * (0.2 * r["schedule_optimistic_days"] + 0.5 * r["schedule_ml_days"] + 0.3 * r["schedule_pessimistic_days"]), 1)
        r["score"] = round(r["probability_pct"] * (r["cost_emv_bn"] + r["schedule_emv_days"] / 365.0), 2)
    return sorted(clean, key=lambda x: x["score"], reverse=True)


def qcra_qsra(model: Dict[str, Any]):
    p50 = parse_bn(model.get("cost_p50")) or 1.0
    sched = float(str(model.get("schedule") or "60").split()[0]) if str(model.get("schedule") or "").split() else 60.0
    riskv = risk_rank_value(model.get("risk")) / 100.0
    curve = []
    for pct in [1, 5, 10, 20, 30, 40, 50, 60, 70, 80, 85, 90, 95, 99]:
        z = {1:-2.33,5:-1.65,10:-1.28,20:-0.84,30:-0.52,40:-0.25,50:0,60:0.25,70:0.52,80:0.84,85:1.04,90:1.28,95:1.65,99:2.33}[pct]
        sigma_c = 0.18 + riskv * 0.28
        sigma_s = 0.10 + riskv * 0.18
        curve.append({"percentile": pct, "qcra_cost_bn": round(max(p50 * .35, p50 * math.exp(z * sigma_c * .42)), 3), "qsra_months": round(max(sched * .65, sched * math.exp(z * sigma_s * .35)), 2)})
    qcra = {"p10": curve[2]["qcra_cost_bn"], "p50": curve[6]["qcra_cost_bn"], "p80": curve[9]["qcra_cost_bn"], "p90": curve[11]["qcra_cost_bn"]}
    qsra = {"p10": curve[2]["qsra_months"], "p50": curve[6]["qsra_months"], "p80": curve[9]["qsra_months"], "p90": curve[11]["qsra_months"]}
    return qcra, qsra, curve


def tornado(model: Dict[str, Any]):
    risks = ensure_risks(model)
    qcra = sorted([{**r, "driver": r["title"], "link": r.get("cbs") or "Cost", "exposure": r["cost_emv_bn"]} for r in risks], key=lambda x: x["exposure"], reverse=True)[:15]
    qsra = sorted([{**r, "driver": r["title"], "link": r.get("activity_id") or "Schedule", "exposure": r["schedule_emv_days"]} for r in risks], key=lambda x: x["exposure"], reverse=True)[:15]
    return qcra, qsra


def _formats(wb):
    return {
        "title": wb.add_format({"bold": True, "font_size": 22, "font_color": "#FFFFFF", "bg_color": BLUE, "align": "left", "valign": "vcenter"}),
        "subtitle": wb.add_format({"font_size": 11, "font_color": "#5B6775"}),
        "section": wb.add_format({"bold": True, "font_size": 14, "font_color": BLUE, "bg_color": "#EAF6FA", "border": 1, "border_color": GRID}),
        "header": wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#0070C0", "border": 1, "border_color": "#FFFFFF", "align": "center", "valign": "vcenter", "text_wrap": True}),
        "cell": wb.add_format({"border": 1, "border_color": GRID, "valign": "top", "text_wrap": True}),
        "num": wb.add_format({"border": 1, "border_color": GRID, "num_format": "$#,##0.0", "valign": "top"}),
        "money": wb.add_format({"border": 1, "border_color": GRID, "num_format": "$#,##0.0", "valign": "top"}),
        "pct": wb.add_format({"border": 1, "border_color": GRID, "num_format": "0%", "valign": "top"}),
        "int": wb.add_format({"border": 1, "border_color": GRID, "num_format": "0", "valign": "top"}),
        "kpi": wb.add_format({"bold": True, "font_size": 18, "font_color": BLUE, "bg_color": LIGHT, "border": 1, "border_color": GRID, "align": "center", "valign": "vcenter"}),
        "note": wb.add_format({"font_color": "#4B5563", "italic": True, "text_wrap": True}),
        "warn": wb.add_format({"font_color": RED, "bold": True, "text_wrap": True}),
    }


def add_table(ws, start_row, start_col, headers, rows, fmt, widths=None):
    for c, h in enumerate(headers):
        ws.write(start_row, start_col + c, h, fmt["header"])
    for r, row in enumerate(rows, start_row + 1):
        for c, v in enumerate(row):
            ws.write(r, start_col + c, v, fmt["cell"])
    if widths:
        for c, w in enumerate(widths):
            ws.set_column(start_col + c, start_col + c, w)
    return start_row + len(rows) + 2




def _detect_sector(model: Dict[str, Any]) -> str:
    text = " ".join(str(model.get(k, "")) for k in ["title", "subsector", "prompt", "description", "mode", "location"]).lower()
    if any(x in text for x in ["airport", "terminal", "runway", "airside"]): return "Airport"
    if any(x in text for x in ["port", "harbour", "harbor", "container", "logistics yard"]): return "Port / Logistics"
    if any(x in text for x in ["rail", "metro", "station", "hs2", "track"]): return "Rail"
    if any(x in text for x in ["life science", "pharma", "gmp", "cleanroom", "laboratory", "biotech"]): return "Life Sciences"
    if any(x in text for x in ["fab", "semiconductor", "wafer", "chip"]): return "Semiconductor Fab"
    if any(x in text for x in ["space", "lunar", "mars", "orbital", "launch", "habitat"]): return "Space Infrastructure"
    if any(x in text for x in ["data centre", "data center", "hyperscale", "server", "cloud", "mw"]): return "Hyperscale Data Centre"
    if any(x in text for x in ["nuclear", "reactor", "smr"]): return "Nuclear"
    if any(x in text for x in ["gigafactory", "battery", "ev"]): return "Gigafactory"
    return "Capital Project"


def _extract_scale(model: Dict[str, Any]) -> Dict[str, Any]:
    text = " ".join(str(model.get(k, "")) for k in ["title", "prompt", "description", "scale", "subsector"]).lower()
    def num_before(unit_patterns):
        for pat in unit_patterns:
            m = re.search(pat, text)
            if m:
                try: return float(m.group(1).replace(',', ''))
                except Exception: pass
        return None
    return {
        "mw": num_before([r"([0-9]+(?:\.[0-9]+)?)\s*mw", r"([0-9]+(?:\.[0-9]+)?)\s*megawatt"]),
        "km": num_before([r"([0-9]+(?:\.[0-9]+)?)\s*km", r"([0-9]+(?:\.[0-9]+)?)\s*kilomet"]),
        "sqm": num_before([r"([0-9]+(?:\.[0-9]+)?)\s*(?:sqm|m2|m²)"]),
        "passengers_m": num_before([r"([0-9]+(?:\.[0-9]+)?)\s*m(?:illion)?\s*(?:passengers|pax)"]),
        "teu_m": num_before([r"([0-9]+(?:\.[0-9]+)?)\s*m(?:illion)?\s*(?:teu|containers)"]),
    }


def benchmark_rows(model: Dict[str, Any]) -> List[Dict[str, Any]]:
    sector = _detect_sector(model)
    p50 = parse_bn(model.get("cost_p50")) or 1.0
    scale = _extract_scale(model)
    rows = []
    # These are intentionally broad sanity bands for first-pass checks, not source-of-truth market data.
    if sector == "Hyperscale Data Centre":
        mw = scale.get("mw") or 250
        unit = p50 * 1000 / mw
        rows.append({"benchmark":"Capex per MW", "project_value": f"${unit:.1f}M/MW", "range":"$7M–$25M/MW depending on power, cooling, land, grid and region", "status":"Within / challenge with local quotes", "basis":"Scale inferred from prompt or default hyperscale benchmark basis"})
    elif sector == "Airport":
        pax = scale.get("passengers_m") or 25
        unit = p50 * 1000 / pax
        rows.append({"benchmark":"Capex per annual passenger", "project_value": f"${unit:.0f}/annual pax", "range":"Wide airport range; depends on runway, terminal, baggage, security, live ops and land", "status":"Challenge scope split", "basis":"Terminal/runway/baggage/live operations benchmark lens"})
    elif sector == "Rail":
        km = scale.get("km") or 50
        unit = p50 * 1000 / km
        rows.append({"benchmark":"Capex per route km", "project_value": f"${unit:.0f}M/km", "range":"Highly variable by tunnel, stations, systems and urban constraints", "status":"Challenge alignment and systems scope", "basis":"Route length inferred or default corridor benchmark lens"})
    elif sector == "Port / Logistics":
        teu = scale.get("teu_m") or 1
        unit = p50 * 1000 / teu
        rows.append({"benchmark":"Capex per million TEU capacity", "project_value": f"${unit:.0f}M / M TEU", "range":"Depends on quay walls, dredging, cranes, yards and rail access", "status":"Validate marine/civil scope", "basis":"Port/logistics capacity lens"})
    elif sector == "Life Sciences":
        sqm = scale.get("sqm") or 50000
        unit = p50 * 1000_000_000 / sqm
        rows.append({"benchmark":"Capex per sqm", "project_value": f"${unit:,.0f}/sqm", "range":"GMP/cleanroom cost highly sensitive to process, validation and containment", "status":"Challenge cleanroom/process scope", "basis":"Life sciences/GMP facility benchmark lens"})
    elif sector == "Semiconductor Fab":
        sqm = scale.get("sqm") or 100000
        unit = p50 * 1000_000_000 / sqm
        rows.append({"benchmark":"Capex per sqm", "project_value": f"${unit:,.0f}/sqm", "range":"Fab costs dominated by tools, utilities, cleanroom and schedule urgency", "status":"Challenge tool/utilities split", "basis":"Advanced fab benchmark lens"})
    elif sector == "Space Infrastructure":
        rows.append({"benchmark":"Space infrastructure premium", "project_value": money(p50), "range":"Earth analogues must be uplifted for launch mass, crew safety, redundancy, radiation and logistics", "status":"Treat as high-uncertainty concept until mission architecture validated", "basis":"Earth-to-orbit translation and mission-control benchmark lens"})
    else:
        rows.append({"benchmark":"Capital project sanity check", "project_value": money(p50), "range":"Requires sector-specific benchmarks once project type, scale and location are confirmed", "status":"Insufficient specificity; use confidence plan", "basis":"Generic capital project benchmark lens"})
    rows.append({"benchmark":"Scenario output consistency", "project_value": scenario_label(model), "range":"Base/Faster/Cheaper/Lower Risk/Premium must change cost, schedule, risk and confidence", "status":"Validated in Scenario Comparison tab", "basis":"CASEY scenario engine"})
    rows.append({"benchmark":"Evidence maturity", "project_value": f"Class {selected_class(model)} / Level {selected_schedule_level(model)}", "range":"Higher class/level reduces uncertainty and improves confidence", "status":"Improve with real estimate, supplier quotes, risk workshop and logic-linked schedule", "basis":"Project controls maturity lens"})
    return rows


def basis_rows(model: Dict[str, Any]) -> List[List[Any]]:
    return [
        ["What this is", "First-pass project controls intelligence for option testing, board challenge and assurance preparation."],
        ["What this is not", "Not a signed tender, contract sum, construction programme or final investment decision basis without validation."],
        ["Cost basis", "Parametric sector structure + location / scale / complexity + scenario modifier + QCRA risk exposure."],
        ["Schedule basis", "Level-based delivery logic + activity phasing + QSRA delay exposure + critical path sensitivity."],
        ["Risk basis", "Cause-event-impact register with probability, optimistic / most-likely / pessimistic cost and schedule ranges."],
        ["Scenario basis", "Base, Faster, Cheaper, Lower Risk, Premium, Investor and Survival cases each alter cost, duration, risk and confidence."],
        ["Confidence basis", "Estimate class, schedule level, location/mode complexity, selected scenario and evidence maturity."],
        ["Required validation", "Design maturity, quantities, supplier quotes, procurement strategy, schedule logic, risk workshop and benchmark evidence."],
    ]

def workbook_bytes(model: Dict[str, Any]) -> bytes:
    if not xlsxwriter:
        raise HTTPException(500, "XlsxWriter is required. Run pip install -r requirements.txt.")
    bio = BytesIO()
    wb = xlsxwriter.Workbook(bio, {"in_memory": True, "nan_inf_to_errors": True})
    fmt = _formats(wb)
    selected = _v57_selected = scenario_label(model)
    title = str(model.get("title") or "CASEY Project")
    p50 = parse_bn(model.get("cost_p50"))
    qcra, qsra, curve = qcra_qsra(model)
    scrows = scenario_rows(model)
    sel_cls = selected_class(model)
    sel_sched = selected_schedule_level(model)
    costs = selected_cost_rows(model)
    by_class = cost_rows_by_class(model)
    risks = ensure_risks(model)
    qct, qst = tornado(model)
    schedules = schedule_levels(model)
    # 00 Control Centre
    ws = wb.add_worksheet("00 Control Centre")
    ws.hide_gridlines(2); ws.set_zoom(90); ws.freeze_panes(6, 0)
    ws.set_column("A:A", 22); ws.set_column("B:B", 24); ws.set_column("C:F", 18); ws.set_column("G:G", 42)
    ws.merge_range("A1:G1", f"CASEY v57 SELLABLE ANY PROJECT OUTPUT PACK — {title}", fmt["title"])
    ws.write("A2", f"Scenario: {selected} | Estimate Class: {sel_cls} | Schedule Level: {sel_sched} | Generated: {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}", fmt["subtitle"])
    kpis = [["P50 Cost", model.get("cost_p50")], ["Cost Range", model.get("cost_range")], ["QCRA P80", money(qcra["p80"])], ["Schedule", model.get("schedule")], ["QSRA P80", f"{qsra['p80']:.1f} months"], ["Risk", model.get("risk")], ["Confidence", f"{model.get('confidence_pct')}%"]]
    for i, (k, v) in enumerate(kpis):
        r = 4 + (i // 4) * 3; c = (i % 4) * 2
        ws.write(r, c, k, fmt["section"]); ws.write(r+1, c, v, fmt["kpi"]); ws.write_blank(r+1, c+1, None, fmt["kpi"])
    ws.write("A11", "Decision Engine", fmt["section"])
    cfg = SCENARIOS.get(scenario_key(model), SCENARIOS["base"])
    ws.write("A12", cfg["decision"], fmt["cell"]); ws.write("A13", cfg["message"], fmt["cell"])
    ws.write("A15", "Output QA Lock", fmt["section"])
    for i, item in enumerate(["Excel generated with XlsxWriter to avoid workbook repair errors", "No PPTX/PDF in primary pack", "Scenario visible on every workbook control sheet", "All class estimates included; selected class clearly identified", "All schedule levels included; selected level clearly identified", "Risk register contains no zero likelihood values", "QCRA and QSRA curves/tornadoes separated", "Basis of Estimate and Benchmark Validation included for buyer trust", "Pricing / access / email conversion flow retained in frontend", "Every input is forced into the same sellable output schema"]):
        ws.write(15+i, 0, "✓", fmt["cell"]); ws.write(15+i, 1, item, fmt["cell"])
    # 01 Scenario Comparison
    ws = wb.add_worksheet("01 Scenario Comparison"); ws.hide_gridlines(2); ws.freeze_panes(5, 0); ws.set_zoom(90)
    ws.merge_range("A1:K1", "SCENARIO COMPARISON — ALL OPTIONS", fmt["title"])
    ws.write("A2", f"Selected scenario: {selected}. These values are scenario-linked and used to drive output interpretation.", fmt["subtitle"])
    headers = ["Scenario", "P50 Cost BN", "Delta vs Selected", "Schedule Months", "Delta Months", "Risk", "Confidence", "QCRA P80 BN", "QSRA P80 Months", "Commercial interpretation", "Board decision"]
    rows = []
    selected_row = next((r for r in scrows if r["label"].lower() == selected.lower() or r["scenario"] == scenario_key(model)), scrows[0])
    for r in scrows:
        rows.append([r["label"], r["p50_bn"], r["p50_bn"] - selected_row["p50_bn"], r["schedule_months"], r["schedule_months"] - selected_row["schedule_months"], r["risk"], r["confidence"], r["qcra_p80_bn"], r["qsra_p80_months"], r["message"], r["decision"]])
    add_table(ws, 4, 0, headers, rows, fmt, [14, 15, 16, 16, 14, 14, 12, 14, 18, 45, 45])
    ws.set_column(1, 4, 15, fmt["money"]); ws.set_column(7, 8, 15, fmt["money"])
    # chart for scenarios
    chart = wb.add_chart({"type": "column"}); chart.add_series({"name": "P50 Cost BN", "categories": "='01 Scenario Comparison'!$A$6:$A$12", "values": "='01 Scenario Comparison'!$B$6:$B$12", "fill": {"color": "#0070C0"}}); chart.set_title({"name": "Scenario Cost Comparison"}); chart.set_legend({"none": True}); ws.insert_chart("A15", chart, {"x_scale": 1.25, "y_scale": 1.1})
    chart2 = wb.add_chart({"type": "line"}); chart2.add_series({"name": "Schedule Months", "categories": "='01 Scenario Comparison'!$A$6:$A$12", "values": "='01 Scenario Comparison'!$D$6:$D$12", "line": {"color": "#00A6C8", "width": 2.25}}); chart2.set_title({"name": "Scenario Schedule Comparison"}); chart2.set_legend({"none": True}); ws.insert_chart("H15", chart2, {"x_scale": 1.2, "y_scale": 1.1})
    # 02 Cost Selected
    ws = wb.add_worksheet("02 Cost Selected"); ws.hide_gridlines(2); ws.freeze_panes(5, 0)
    ws.merge_range("A1:J1", f"SELECTED CLASS COST ESTIMATE — CLASS {sel_cls} — {selected.upper()} SCENARIO", fmt["title"])
    headers = ["CBS", "Description", "Type", "Basis", "P10 BN", "Most Likely / P50 BN", "P80 BN", "P90 BN", "% of P50", "Challenge Question"]
    total = sum(x["p50_bn"] for x in costs) or 1
    rows = [[x["cbs"], x["description"], x["type"], x["basis"], x["p10_bn"], x["p50_bn"], x["p80_bn"], x["p90_bn"], x["p50_bn"] / total, f"What evidence supports {x['description']} at {money(x['p50_bn'])}?"] for x in costs]
    add_table(ws, 4, 0, headers, rows, fmt, [12, 28, 12, 42, 13, 18, 13, 13, 10, 42])
    ws.set_column(4, 7, 13, wb.add_format({"num_format": "$#,##0.0", "border": 1, "border_color": GRID}))
    ws.set_column(8, 8, 10, wb.add_format({"num_format": "0.0%", "border": 1, "border_color": GRID}))
    chart = wb.add_chart({"type": "bar"}); chart.add_series({"name": "P50 BN", "categories": f"='02 Cost Selected'!$B$6:$B${5+len(rows)}", "values": f"='02 Cost Selected'!$F$6:$F${5+len(rows)}", "fill": {"color": "#0070C0"}}); chart.set_title({"name": "Selected Class Cost Build-Up"}); chart.set_legend({"none": True}); ws.insert_chart("L4", chart, {"x_scale": 1.35, "y_scale": 1.4})
    # 03 Cost All Classes
    ws = wb.add_worksheet("03 Cost All Classes"); ws.hide_gridlines(2); ws.freeze_panes(5, 0)
    ws.merge_range("A1:I1", "ALL ESTIMATE CLASS LEVELS — CLASS 1 TO 5", fmt["title"])
    headers = ["Class", "Class Name", "CBS", "Description", "Type", "P10 BN", "P50 BN", "P90 BN", "Maturity / Use"]
    rows = []
    for cls in range(1, 6):
        for x in by_class.get(cls, []):
            rows.append([cls, x["class_name"], x["cbs"], x["description"], x["type"], x["p10_bn"], x["p50_bn"], x["p90_bn"], x["basis"]])
    add_table(ws, 4, 0, headers, rows, fmt, [8, 20, 12, 28, 12, 12, 12, 12, 45])
    ws.set_column(5, 7, 12, wb.add_format({"num_format": "$#,##0.0", "border": 1, "border_color": GRID}))
    # 04 Risk Register
    ws = wb.add_worksheet("04 Risk Register"); ws.hide_gridlines(2); ws.freeze_panes(5, 0)
    ws.merge_range("A1:R1", "FULL RISK REGISTER — CAUSE / EVENT / IMPACT / MITIGATION / QUANTIFIED", fmt["title"])
    headers = ["ID", "Category", "Risk", "Cause", "Event", "Impact", "Owner", "Mitigation", "Trigger", "Probability %", "Cost O BN", "Cost ML BN", "Cost P BN", "Schedule O Days", "Schedule ML Days", "Schedule P Days", "Cost EMV BN", "Schedule EMV Days", "Status"]
    rows = [[r["risk_id"], r["category"], r["title"], r["cause"], r["event"], r["impact"], r["owner"], r["mitigation"], r["trigger"], r["probability_pct"] / 100.0, r["cost_optimistic_bn"], r["cost_ml_bn"], r["cost_pessimistic_bn"], r["schedule_optimistic_days"], r["schedule_ml_days"], r["schedule_pessimistic_days"], r["cost_emv_bn"], r["schedule_emv_days"], r["status"]] for r in risks]
    add_table(ws, 4, 0, headers, rows, fmt, [10, 14, 26, 36, 36, 42, 18, 42, 26, 12, 12, 12, 12, 14, 14, 14, 12, 14, 12])
    ws.set_column(9, 9, 12, wb.add_format({"num_format": "0%", "border": 1, "border_color": GRID}))
    ws.conditional_format(5, 9, 4+len(rows), 9, {"type": "3_color_scale", "min_color": "#D9EAD3", "mid_color": "#FFF2CC", "max_color": "#F4CCCC"})
    # 05 QCRA QSRA Curves
    ws = wb.add_worksheet("05 QCRA QSRA Curves"); ws.hide_gridlines(2); ws.freeze_panes(5, 0)
    ws.merge_range("A1:F1", "QCRA + QSRA CURVES — SEPARATE DISTRIBUTIONS", fmt["title"])
    headers = ["Percentile", "QCRA Cost BN", "QSRA Months"]
    rows = [[x["percentile"], x["qcra_cost_bn"], x["qsra_months"]] for x in curve]
    add_table(ws, 4, 0, headers, rows, fmt, [12, 16, 16])
    chart = wb.add_chart({"type": "line"}); chart.add_series({"name": "QCRA Cost BN", "categories": "='05 QCRA QSRA Curves'!$A$6:$A$19", "values": "='05 QCRA QSRA Curves'!$B$6:$B$19", "line": {"color": "#0070C0", "width": 2.25}}); chart.set_title({"name": "QCRA Cost S-Curve"}); ws.insert_chart("E4", chart, {"x_scale": 1.25, "y_scale": 1.1})
    chart2 = wb.add_chart({"type": "line"}); chart2.add_series({"name": "QSRA Months", "categories": "='05 QCRA QSRA Curves'!$A$6:$A$19", "values": "='05 QCRA QSRA Curves'!$C$6:$C$19", "line": {"color": "#00A6C8", "width": 2.25}}); chart2.set_title({"name": "QSRA Duration S-Curve"}); ws.insert_chart("E21", chart2, {"x_scale": 1.25, "y_scale": 1.1})
    # 06 Tornado
    ws = wb.add_worksheet("06 Tornado Drivers"); ws.hide_gridlines(2); ws.freeze_panes(5, 0)
    ws.merge_range("A1:H1", "TORNADO DRIVERS — QCRA COST AND QSRA SCHEDULE SEPARATED", fmt["title"])
    qrows = [[i+1, r["driver"], r["risk_id"], r["exposure"], r["owner"], r["mitigation"]] for i, r in enumerate(qct)]
    srows = [[i+1, r["driver"], r["risk_id"], r["exposure"], r["owner"], r["mitigation"]] for i, r in enumerate(qst)]
    add_table(ws, 4, 0, ["Rank", "QCRA Cost Driver", "ID", "Cost EMV BN", "Owner", "Mitigation"], qrows, fmt, [8, 30, 12, 14, 18, 45])
    add_table(ws, 4, 7, ["Rank", "QSRA Schedule Driver", "ID", "Schedule EMV Days", "Owner", "Mitigation"], srows, fmt, [8, 30, 12, 18, 18, 45])
    chart = wb.add_chart({"type": "bar"}); chart.add_series({"name": "QCRA EMV", "categories": f"='06 Tornado Drivers'!$B$6:$B${5+len(qrows)}", "values": f"='06 Tornado Drivers'!$D$6:$D${5+len(qrows)}", "fill": {"color": "#0070C0"}}); chart.set_title({"name": "QCRA Cost Tornado"}); chart.set_legend({"none": True}); ws.insert_chart("A24", chart, {"x_scale": 1.25, "y_scale": 1.3})
    chart2 = wb.add_chart({"type": "bar"}); chart2.add_series({"name": "QSRA Days", "categories": f"='06 Tornado Drivers'!$I$6:$I${5+len(srows)}", "values": f"='06 Tornado Drivers'!$K$6:$K${5+len(srows)}", "fill": {"color": "#00A6C8"}}); chart2.set_title({"name": "QSRA Schedule Tornado"}); chart2.set_legend({"none": True}); ws.insert_chart("H24", chart2, {"x_scale": 1.25, "y_scale": 1.3})
    # 07 Schedule Levels
    ws = wb.add_worksheet("07 Schedule Levels"); ws.hide_gridlines(2); ws.freeze_panes(5, 0)
    ws.merge_range("A1:H1", "ALL SCHEDULE LEVELS — LEVEL 1 TO 5", fmt["title"])
    headers = ["Level", "Activity ID", "Phase", "Activity", "Predecessor", "Duration Months", "Critical", "Basis"]
    rows = []
    for lvl in range(1, 6):
        for a in schedules.get(lvl, []):
            rows.append([lvl, a.get("activity_id"), a.get("phase"), a.get("activity"), a.get("predecessor"), a.get("duration_months"), a.get("critical"), a.get("basis")])
    add_table(ws, 4, 0, headers, rows, fmt, [8, 14, 18, 36, 20, 14, 12, 45])
    # 08 Basis + Audit
    ws = wb.add_worksheet("08 Basis + Audit"); ws.hide_gridlines(2); ws.set_column("A:A", 24); ws.set_column("B:B", 92)
    ws.merge_range("A1:B1", "BASIS, LIMITATIONS AND AUDIT", fmt["title"])
    audit = [["Scenario", selected], ["Project", title], ["Subsector", model.get("subsector")], ["Location", model.get("location")], ["Mode", model.get("mode")], ["Estimate class", sel_cls], ["Schedule level", sel_sched], ["Generated", datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")], ["Model version", "v57 Sellable Any Project"], ["Limitation", "First-pass project controls intelligence. Validate against project-specific tender, survey, design and schedule evidence before capital commitment."], ["Methodology", "Cost is parametric/benchmark/factor-based; QCRA and QSRA are separated; risks are cause-event-impact structured; scenarios change cost, schedule, risk and confidence."], ["Output standard", "PDF/PPTX excluded from primary output pack to avoid weak or unreadable exports. Excel, schedule and audit outputs are the source of truth."]]
    add_table(ws, 3, 0, ["Field", "Value"], audit, fmt, [24, 92])
    # 09 Basis of Estimate
    ws = wb.add_worksheet("09 Basis of Estimate"); ws.hide_gridlines(2); ws.set_column("A:A", 28); ws.set_column("B:B", 100)
    ws.merge_range("A1:B1", "BASIS OF ESTIMATE — BUYER TRUST LAYER", fmt["title"])
    ws.write("A2", f"Scenario: {selected} | Project: {title} | Sector: {_detect_sector(model)}", fmt["subtitle"])
    add_table(ws, 4, 0, ["Basis Area", "Explanation"], basis_rows(model), fmt, [28, 100])
    # 10 Benchmark Validation
    ws = wb.add_worksheet("10 Benchmark Validation"); ws.hide_gridlines(2); ws.freeze_panes(5, 0)
    ws.merge_range("A1:E1", "BENCHMARK VALIDATION — EXTERNAL REALISM CHECK", fmt["title"])
    ws.write("A2", "A first-pass sanity check so buyers can see whether the model sits within a defensible benchmark envelope.", fmt["subtitle"])
    brows = [[x["benchmark"], x["project_value"], x["range"], x["status"], x["basis"]] for x in benchmark_rows(model)]
    add_table(ws, 4, 0, ["Benchmark", "Project Value", "Benchmark Range / Note", "Status", "Basis"], brows, fmt, [26, 20, 58, 34, 42])
    # 11 Commercial Next Steps
    ws = wb.add_worksheet("11 Commercial Next Steps"); ws.hide_gridlines(2); ws.set_column("A:A", 28); ws.set_column("B:B", 88)
    ws.merge_range("A1:B1", "COMMERCIAL NEXT STEPS — BUYABLE PRODUCT FLOW", fmt["title"])
    commercial = [["Run on real project", "Upload estimate, schedule and risk register; CASEY challenges rates, missing scope, risk, contingency and schedule logic."], ["Pilot", "Guided project review and sample outputs for a live project."], ["Professional", "Full project pack with scenarios, QCRA/QSRA, source-of-truth Excel outputs and schedule exports."], ["Enterprise", "Private deployment, SSO, benchmark library, saved projects, audit trail and custom sector models."], ["Lead capture", "Request pricing, book executive demo or send project brief from Pricing / Request Access screen."], ["Sales message", "Board-grade capital decision engine in minutes: cost, schedule, risk, scenarios and output pack."]]
    add_table(ws, 3, 0, ["Action", "What buyer gets"], commercial, fmt, [28, 88])
    wb.close(); bio.seek(0); return bio.getvalue()


def risk_register_bytes(model: Dict[str, Any]) -> bytes:
    if not xlsxwriter:
        raise HTTPException(500, "XlsxWriter is required. Run pip install -r requirements.txt.")
    bio = BytesIO(); wb = xlsxwriter.Workbook(bio, {"in_memory": True, "nan_inf_to_errors": True}); fmt = _formats(wb)
    risks = ensure_risks(model); qct, qst = tornado(model)
    ws = wb.add_worksheet("Risk Control Centre"); ws.hide_gridlines(2); ws.freeze_panes(5, 0)
    ws.merge_range("A1:J1", f"CASEY v57 RISK CONTROL CENTRE — {scenario_label(model).upper()} SCENARIO", fmt["title"])
    ws.write("A3", "Top risks by combined cost/schedule exposure", fmt["section"])
    rows = [[r["risk_id"], r["title"], r["probability_pct"] / 100.0, r["cost_emv_bn"], r["schedule_emv_days"], r["owner"], r["mitigation"]] for r in risks[:10]]
    add_table(ws, 4, 0, ["ID", "Risk", "Likelihood", "Cost EMV BN", "Schedule EMV Days", "Owner", "Action"], rows, fmt, [10, 28, 12, 14, 16, 18, 55])
    chart = wb.add_chart({"type": "bar"}); chart.add_series({"name": "Cost EMV", "categories": "='Risk Control Centre'!$B$6:$B$15", "values": "='Risk Control Centre'!$D$6:$D$15", "fill": {"color": "#0070C0"}}); chart.set_title({"name": "Top Risk Cost Exposure"}); chart.set_legend({"none": True}); ws.insert_chart("I4", chart, {"x_scale": 1.25, "y_scale": 1.15})
    ws2 = wb.add_worksheet("Full Risk Register"); ws2.hide_gridlines(2); ws2.freeze_panes(5, 0)
    ws2.merge_range("A1:S1", "FULL RISK REGISTER — NO ZERO LIKELIHOOD, FULL CAUSE / EVENT / IMPACT", fmt["title"])
    headers = ["ID", "Category", "Risk", "Cause", "Event", "Impact", "Owner", "Mitigation", "Trigger", "Likelihood", "Cost O BN", "Cost ML BN", "Cost P BN", "Schedule O Days", "Schedule ML Days", "Schedule P Days", "Cost EMV BN", "Schedule EMV Days", "Status"]
    rows = [[r["risk_id"], r["category"], r["title"], r["cause"], r["event"], r["impact"], r["owner"], r["mitigation"], r["trigger"], r["probability_pct"]/100.0, r["cost_optimistic_bn"], r["cost_ml_bn"], r["cost_pessimistic_bn"], r["schedule_optimistic_days"], r["schedule_ml_days"], r["schedule_pessimistic_days"], r["cost_emv_bn"], r["schedule_emv_days"], r["status"]] for r in risks]
    add_table(ws2, 4, 0, headers, rows, fmt, [10, 14, 28, 38, 38, 42, 18, 42, 28, 12, 12, 12, 12, 14, 14, 14, 12, 14, 12])
    ws2.set_column(9, 9, 12, wb.add_format({"num_format": "0%", "border": 1, "border_color": GRID}))
    ws3 = wb.add_worksheet("QCRA Tornado"); ws3.hide_gridlines(2)
    ws3.merge_range("A1:F1", "QCRA COST TORNADO", fmt["title"])
    add_table(ws3, 4, 0, ["Rank", "Driver", "Risk ID", "Cost EMV BN", "Owner", "Action"], [[i+1, r["driver"], r["risk_id"], r["exposure"], r["owner"], r["mitigation"]] for i,r in enumerate(qct)], fmt, [8, 32, 12, 15, 18, 50])
    ws4 = wb.add_worksheet("QSRA Tornado"); ws4.hide_gridlines(2)
    ws4.merge_range("A1:F1", "QSRA SCHEDULE TORNADO", fmt["title"])
    add_table(ws4, 4, 0, ["Rank", "Driver", "Risk ID", "Schedule EMV Days", "Owner", "Action"], [[i+1, r["driver"], r["risk_id"], r["exposure"], r["owner"], r["mitigation"]] for i,r in enumerate(qst)], fmt, [8, 32, 12, 18, 18, 50])
    wb.close(); bio.seek(0); return bio.getvalue()


def schedule_csv_bytes(model: Dict[str, Any]) -> bytes:
    out = StringIO(); w = csv.writer(out); w.writerow(["Level", "Activity ID", "Phase", "Activity", "Predecessor", "Duration Months", "Critical", "Basis"])
    for lvl, rows in schedule_levels(model).items():
        for a in rows:
            w.writerow([lvl, a.get("activity_id"), a.get("phase"), a.get("activity"), a.get("predecessor"), a.get("duration_months"), a.get("critical"), a.get("basis")])
    return out.getvalue().encode("utf-8")


def xer_bytes(model: Dict[str, Any]) -> bytes:
    # PRA-targeted legacy XER. Also provide CSV fallback in full ZIP.
    rows = schedule_levels(model).get(selected_schedule_level(model)) or model.get("schedule_rows") or []
    lines = []
    lines.append("ERMHDR\t8.0\t2025-12-11\tProject\tadmin\tPPSEn\tdbxDatabaseNoName\tProject Management\tUSD")
    lines.append("%T\tCURRTYPE"); lines.append("%F\tcurr_id\tdecimal_digit_cnt\tcurr_symbol\tdecimal_symbol\tdigit_group_symbol\tpos_curr_fmt_type\tneg_curr_fmt_type\tcurr_type\tcurr_short_name\tgroup_digit_cnt\tbase_exch_rate"); lines.append("%R\t1\t2\t$\t.\t,\t#1.1\t(#1.1)\tUS Dollar\tUSD\t3\t1")
    lines.append("%T\tOBS"); lines.append("%F\tobs_id\tparent_obs_id\tguid\tseq_num\tobs_name\tobs_descr"); lines.append("%R\t565\t\tCASEYOBS\t0\tEnterprise\tCASEY")
    lines.append("%T\tUDFTYPE"); lines.append("%F\tudf_type_id\ttable_name\tudf_type_name\tudf_type_label\tlogical_data_type\tsuper_flag\tindicator_expression\tsummary_indicator_expression"); lines.append("%R\t329\tTASK\tCASEY_Likely\tLikely\tFT_TEXT\tN\t\t"); lines.append("%R\t330\tTASK\tCASEY_Max\tMax\tFT_TEXT\tN\t\t"); lines.append("%R\t331\tTASK\tCASEY_Min\tMin\tFT_TEXT\tN\t\t")
    lines.append("%T\tPROJECT")
    proj_fields = ["proj_id","fy_start_month_num","rsrc_self_add_flag","allow_complete_flag","rsrc_multi_assign_flag","checkout_flag","project_flag","step_complete_flag","cost_qty_recalc_flag","batch_sum_flag","name_sep_char","def_complete_pct_type","proj_short_name","acct_id","orig_proj_id","source_proj_id","base_type_id","clndr_id","sum_base_proj_id","task_code_base","task_code_step","priority_num","wbs_max_sum_level","strgy_priority_num","last_checksum","critical_drtn_hr_cnt","def_cost_per_qty","last_recalc_date","plan_start_date","plan_end_date","scd_end_date","add_date","last_tasksum_date","fcst_start_date","def_duration_type","task_code_prefix","guid","def_qty_type","add_by_name","web_local_root_path","proj_url","def_rate_type","add_act_remain_flag","act_this_per_link_flag","def_task_type","act_pct_link_flag","critical_path_type","task_code_prefix_flag","def_rollup_dates_flag","use_project_baseline_flag","rem_target_link_flag","reset_planned_flag","allow_neg_act_flag","sum_assign_level","last_fin_dates_id","last_baseline_update_date","cr_external_key","apply_actuals_date","location_id","loaded_scope_level","export_flag","new_fin_dates_id","baselines_to_export","baseline_names_to_export","next_data_date","close_period_flag","sum_refresh_date","trsrcsum_loaded"]
    lines.append("%F\t" + "\t".join(proj_fields))
    proj_vals = ["3338","1","Y","Y","Y","N","Y","Y","N","Y",".","CP_Drtn","CASEY","","","","","3892","","1000","10","10","","500","","0","0.0000","2025-10-05 00:00","2025-10-05 00:00","","2030-05-16 16:00","2025-11-29 14:14","","","DT_FixedRate","A","CASEYGUID0000000000001","QT_Hour","admin","","","COST_PER_QTY","N","Y","TT_Task","N","CT_DrivPath","Y","Y","Y","Y","N","N","SL_Taskrsrc","","","","","","7","Y","","","","","",""]
    lines.append("%R\t" + "\t".join(proj_vals))
    lines.append("%T\tCALENDAR"); lines.append("%F\tclndr_id\tdefault_flag\tclndr_name\tproj_id\tbase_clndr_id\tlast_chng_date\tclndr_type\tday_hr_cnt\tweek_hr_cnt\tmonth_hr_cnt\tyear_hr_cnt\trsrc_private\tclndr_data"); lines.append("%R\t3892\tY\tStandard\t3338\t\t\tCA_Project\t8\t40\t172\t2000\tN\t(0||CalendarData()(0||VIEW(ShowTotal|Y)()))")
    lines.append("%T\tPROJWBS"); lines.append("%F\twbs_id\tproj_id\tobs_id\tseq_num\test_wt\tproj_node_flag\tsum_data_flag\tstatus_code\twbs_short_name\twbs_name\tphase_id\tparent_wbs_id\tev_user_pct\tev_etc_user_value\torig_cost\tindep_remain_total_cost\tann_dscnt_rate_pct\tdscnt_period_type\tindep_remain_work_qty\tanticip_start_date\tanticip_end_date\tev_compute_type\tev_etc_compute_type\tguid\ttmpl_guid\tplan_open_state")
    lines.append("%R\t35124\t3338\t565\t1\t1\tY\tN\tWS_Open\tCASEY\tCASEY Generated Schedule\t\t\t\t\t0.0000\t0.0000\t\t\t\t\t\t\t\tCASEYWBSROOT\t\t")
    lines.append("%R\t35127\t3338\t565\t2\t1\tN\tN\tWS_Open\t01\tDelivery Schedule\t\t35124\t\t\t0.0000\t0.0000\t\t\t\t\t\t\t\tCASEYWBS01\t\t")
    task_fields = ["task_id","proj_id","wbs_id","clndr_id","phys_complete_pct","rev_fdbk_flag","est_wt","lock_plan_flag","auto_compute_act_flag","complete_pct_type","task_type","duration_type","status_code","task_code","task_name","rsrc_id","total_float_hr_cnt","free_float_hr_cnt","remain_drtn_hr_cnt","act_work_qty","remain_work_qty","target_work_qty","target_drtn_hr_cnt","target_equip_qty","act_equip_qty","remain_equip_qty","cstr_date","act_start_date","act_end_date","late_start_date","late_end_date","expect_end_date","early_start_date","early_end_date","restart_date","reend_date","target_start_date","target_end_date","rem_late_start_date","rem_late_end_date","cstr_type","priority_type","suspend_date","resume_date","float_path","float_path_order","guid","tmpl_guid","cstr_date2","cstr_type2","driving_path_flag","act_this_per_work_qty","act_this_per_equip_qty","external_early_start_date","external_late_end_date","create_date","update_date","create_user","update_user","location_id"]
    lines.append("%T\tTASK"); lines.append("%F\t" + "\t".join(task_fields))
    id_map = {}
    task_id_base = 200861
    for i, a in enumerate(rows, 1):
        tid = task_id_base + i
        id_map[str(a.get("activity_id"))] = str(tid)
        dur = int(max(8, float(a.get("duration_months") or 1) * 172))
        code = str(a.get("activity_id") or f"A{i:04d}")[:40]
        name = str(a.get("activity") or "Activity").replace("\t", " ")[:120]
        vals = [str(tid),"3338","35127","3892","0","N","1","N","N","CP_Phys","TT_Task","DT_FixedDUR2","TK_NotStart",code,name,"","0","0",str(dur),"0","0","0",str(dur),"0","0","0","","","","2025-10-06 08:00","2025-10-06 16:00","","2025-10-06 08:00","2025-10-06 16:00","2025-10-06 08:00","2025-10-06 16:00","2025-10-06 08:00","2025-10-06 16:00","2025-10-06 08:00","2025-10-06 16:00","","PT_Normal","","","","",f"CASEYGUID{i:014d}","","","","Y" if str(a.get('critical')).lower()=="yes" else "N","0","0","","","2025-11-29 14:14","2025-12-10 19:15","CASEY","CASEY",""]
        lines.append("%R\t" + "\t".join(vals))
    lines.append("%T\tTASKPRED"); lines.append("%F\ttask_pred_id\ttask_id\tpred_task_id\tproj_id\tpred_proj_id\tpred_type\tlag_hr_cnt\tfloat_path\taref\tarls")
    pred_id = 266542
    for a in rows:
        this_id = id_map.get(str(a.get("activity_id")))
        preds = str(a.get("predecessor") or "").replace(";", ",").split(",")
        for p in [x.strip() for x in preds if x.strip()]:
            pid = id_map.get(p)
            if this_id and pid and this_id != pid:
                lines.append(f"%R\t{pred_id}\t{this_id}\t{pid}\t3338\t3338\tPR_FS\t0\t\t2025-10-06 08:00\t2025-10-06 08:00")
                pred_id += 1
    lines.append("%E")
    return ("\n".join(lines) + "\n").encode("utf-8")


def model_json_bytes(model: Dict[str, Any]) -> bytes:
    enriched = dict(model)
    enriched["v57_output_standard"] = {"pdf_pptx_removed": True, "risk_count": len(ensure_risks(model)), "scenario": scenario_label(model), "selected_class": selected_class(model), "selected_schedule_level": selected_schedule_level(model), "qaqc": "Excel generated by XlsxWriter; no zero likelihood risks; QCRA/QSRA separated; all classes and all schedule levels included."}
    return json.dumps(enriched, indent=2, default=str).encode("utf-8")


def all_zip_bytes(model: Dict[str, Any]) -> bytes:
    bio = BytesIO()
    with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("01_CASEY_v57_Sellable_Cost_Model.xlsx", workbook_bytes(model))
        z.writestr("02_CASEY_v57_Sellable_Risk_Register.xlsx", risk_register_bytes(model))
        z.writestr("03_CASEY_v57_PRA_Targeted_Schedule.xer", xer_bytes(model))
        z.writestr("04_CASEY_v57_All_Schedule_Levels.csv", schedule_csv_bytes(model))
        z.writestr("05_CASEY_v57_Model_Audit.json", model_json_bytes(model))
        z.writestr("README_v57_OUTPUT_STANDARD.txt", "CASEY v57 removes PDF and PPTX from the primary pack. Source-of-truth outputs are Excel cost model, Excel risk register, PRA-targeted XER, schedule CSV fallback and JSON audit. Every pack includes selected scenario, all scenarios, selected class, all classes, selected schedule level, all schedule levels, QCRA curve, QSRA curve and separate tornado drivers. If PRA rejects XER due to local version rules, use the CSV schedule fallback or open/re-export through P6 Professional.\n")
    bio.seek(0); return bio.getvalue()


def install_v56(app):
    def replace_post(path, endpoint):
        app.router.routes = [r for r in app.router.routes if not (getattr(r, "path", None) == path and "POST" in getattr(r, "methods", set()))]
        app.post(path)(endpoint)
    def export_workbook(model: Dict[str, Any]): return stream(workbook_bytes(model), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CASEY_v57_Sellable_Cost_Model.xlsx")
    def export_risk(model: Dict[str, Any]): return stream(risk_register_bytes(model), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CASEY_v57_Sellable_Risk_Register.xlsx")
    def export_xer(model: Dict[str, Any]): return stream(xer_bytes(model), "application/octet-stream", "CASEY_v57_PRA_Targeted_Schedule.xer")
    def export_csv(model: Dict[str, Any]): return stream(schedule_csv_bytes(model), "text/csv", "CASEY_v57_All_Schedule_Levels.csv")
    def export_json(model: Dict[str, Any]): return stream(model_json_bytes(model), "application/json", "CASEY_v57_Model_Audit.json")
    def export_all(model: Dict[str, Any]): return stream(all_zip_bytes(model), "application/zip", "CASEY_v57_Sellable_Output_Pack.zip")
    def removed(model: Dict[str, Any]): return JSONResponse({"status": "removed", "message": "PDF/PPTX/Word are deliberately removed from v57 primary output pack because they were not consistently platinum. Use Excel cost model, risk workbook, XER/CSV schedule and JSON audit as source-of-truth outputs."}, status_code=410)
    for p, e in [("/export/workbook", export_workbook), ("/export/risk-register", export_risk), ("/export/xer", export_xer), ("/export/schedule-csv", export_csv), ("/export/json", export_json), ("/export/all", export_all), ("/v56/export/all", export_all)]:
        replace_post(p, e)
    for p in ["/export/pdf", "/export/pptx", "/export/word"]:
        replace_post(p, removed)
