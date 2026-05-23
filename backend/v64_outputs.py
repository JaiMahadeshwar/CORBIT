from __future__ import annotations
from io import BytesIO
from datetime import datetime, timedelta
import json, zipfile, csv, re, html, math, os
from fastapi.responses import StreamingResponse, JSONResponse

try:
    import xlsxwriter
except Exception:
    xlsxwriter = None

VERSION = "FINAL"

SCENARIO_RULES = {
    "base": {"label":"Base", "cost":1.00, "duration":1.00, "risk":1.00, "confidence":0, "decision":"Use as reference control baseline", "tone":"Balanced baseline for challenge and option testing."},
    "faster": {"label":"Faster", "cost":1.11, "duration":0.82, "risk":1.20, "confidence":-6, "decision":"Approve only if time-to-market value exceeds acceleration premium", "tone":"Compression through parallel workfaces, early procurement and premium logistics."},
    "cheaper": {"label":"Cheaper", "cost":0.88, "duration":1.08, "risk":1.18, "confidence":-8, "decision":"Use only where capital constraint outweighs delivery certainty", "tone":"Value engineering and scope discipline reduce P50 cost but increase residual exposure."},
    "lower_risk": {"label":"Lower Risk", "cost":1.09, "duration":1.10, "risk":0.72, "confidence":9, "decision":"Recommended where approval confidence and downside protection matter most", "tone":"More definition, assurance, buffers and mitigation spend reduce downside risk."},
    "premium": {"label":"Premium", "cost":1.15, "duration":1.04, "risk":0.80, "confidence":12, "decision":"Use for flagship resilience, quality and high-assurance delivery", "tone":"Higher specification, stronger suppliers and governance increase cost but reduce uncertainty."},
    "investor": {"label":"Investor", "cost":0.98, "duration":1.00, "risk":0.92, "confidence":4, "decision":"Use for investment committee challenge and downside transparency", "tone":"Investment framing with clear benchmark, risk and commitment caveats."},
    "survival": {"label":"Survival", "cost":0.78, "duration":1.20, "risk":1.35, "confidence":-12, "decision":"Not recommended unless liquidity or viability is the dominant constraint", "tone":"Minimum viable scope. Maximum warning visibility."},
}
CLASS_RULES = {
    5: {"name":"Class 5 Concept Screening", "maturity":"0–2% definition", "low":0.50, "high":2.00, "cont":0.22, "sched_levels":[1,2], "basis":"Concept order-of-magnitude model. Wide range, strategic screening only."},
    4: {"name":"Class 4 Feasibility", "maturity":"1–15% definition", "low":0.70, "high":1.50, "cont":0.17, "sched_levels":[1,2,3], "basis":"Feasibility model. Benchmarks and limited quantities, major assumptions open."},
    3: {"name":"Class 3 Budget Authorization", "maturity":"10–40% definition", "low":0.80, "high":1.30, "cont":0.12, "sched_levels":[2,3,4], "basis":"Budget authorization model. Selected control baseline with stage-gate validation."},
    2: {"name":"Class 2 Control / Tender", "maturity":"30–75% definition", "low":0.85, "high":1.20, "cont":0.08, "sched_levels":[3,4,5], "basis":"Control/tender model. Package-level definition and stronger supplier evidence expected."},
    1: {"name":"Class 1 Definitive / Bid", "maturity":"65–100% definition", "low":0.90, "high":1.15, "cont":0.05, "sched_levels":[4,5], "basis":"Definitive/bid model. Detailed quantities, logic and commercial basis expected."},
}

EARTH_PHASES = [
    ("Governance", "Project initiation and controls setup"),
    ("Business Case", "Investment case, funding and board approvals"),
    ("Scope", "Requirements definition and scope freeze"),
    ("Surveys", "Site investigations, surveys and constraints validation"),
    ("Design", "Concept, FEED, detailed design and IFC development"),
    ("Consents", "Planning, permitting and stakeholder approvals"),
    ("Procurement", "Procurement strategy, tendering and long-lead awards"),
    ("Enabling", "Enabling works, logistics and temporary infrastructure"),
    ("Civil", "Substructure, foundations and civil works"),
    ("Build", "Superstructure, installation and main construction"),
    ("Systems", "MEP, controls, utilities and specialist systems"),
    ("Integration", "Systems integration, interfaces and testing"),
    ("Commissioning", "Commissioning, validation and operational readiness"),
    ("Handover", "Handover, closeout and benefits tracking"),
]
SPACE_PHASES = [
    ("Mission Definition", "Mission objectives, requirements and success criteria"),
    ("Architecture", "Concept of operations, systems architecture and interfaces"),
    ("Technology", "TRL closure, qualification plan and critical technology validation"),
    ("Design", "Preliminary and detailed design for flight / surface systems"),
    ("Supply Chain", "Payload, launch, habitat and mission-critical procurement"),
    ("Manufacturing", "Fabrication, assembly, integration and test"),
    ("Launch Prep", "Launch integration, range approvals and mission readiness"),
    ("Transit", "Launch window, transfer, orbital insertion or landing operations"),
    ("Deployment", "Surface/orbital deployment, commissioning and operations setup"),
    ("Operations", "Initial operations, reliability growth and sustainment"),
]
SECTOR_RISKS = {
    "data": ["Grid connection delay", "Cooling system capacity", "Power train procurement", "Utility energisation approval", "Liquid cooling commissioning"],
    "airport": ["Airside possession constraint", "Passenger operation disruption", "Baggage integration failure", "Security approval delay", "Runway/taxiway phasing conflict"],
    "rail": ["Possession access delay", "Systems integration failure", "Signalling approval delay", "Utilities diversion conflict", "Rolling stock interface risk"],
    "life": ["GMP validation delay", "Cleanroom performance failure", "Regulatory inspection finding", "Process equipment delay", "Qualification documentation gap"],
    "space": ["Launch window slip", "TRL shortfall", "Mass margin growth", "Radiation/thermal environment", "Surface logistics failure", "Flight qualification delay"],
    "general": ["Scope growth", "Design maturity gap", "Market escalation", "Supply chain delay", "Productivity underperformance", "Interface misalignment", "Permits / approvals delay", "Commissioning delay"],
}


def scenario_key(model):
    s = str(model.get("scenario") or model.get("scenario_label") or "base").lower().replace(" ", "_").replace("-", "_")
    if s in SCENARIO_RULES: return s
    for k,v in SCENARIO_RULES.items():
        if v["label"].lower().replace(" ", "_") == s: return k
    return "base"

def clean(s, maxlen=80):
    s = re.sub(r"[^A-Za-z0-9 _/\-]+", "", str(s or "")).strip()
    return s[:maxlen] or "CASEY Project"

def as_float_money(v):
    if isinstance(v, (int,float)): return float(v)
    st = str(v or "0").replace("$","").replace(",","").strip().upper()
    try:
        if st.endswith("T"): return float(st[:-1])*1000
        if st.endswith("B"): return float(st[:-1])
        if st.endswith("M"): return float(st[:-1])/1000
        return float(st)
    except Exception: return 1.0

def money(v):
    v = float(v or 0)
    if abs(v) >= 1000: return f"${v/1000:.1f}T"
    if abs(v) >= 1: return f"${v:.1f}B"
    return f"${v*1000:.0f}M"

def selected_class(model):
    try: c=int(model.get("estimate_class") or model.get("class_level") or 3)
    except Exception: c=3
    return max(1,min(5,c))

def selected_level(model):
    try: l=int(model.get("schedule_level") or 3)
    except Exception: l=3
    return max(1,min(5,l))

def project_type_key(model):
    t = " ".join([str(model.get("title", "")), str(model.get("subsector", "")), str(model.get("prompt", "")), str(model.get("mode", ""))]).lower()
    if any(x in t for x in ["space","moon","mars","orbit","lunar","launch","satellite"]): return "space"
    if any(x in t for x in ["data centre","data center","hyperscale","cloud"]): return "data"
    if any(x in t for x in ["airport","runway","terminal"]): return "airport"
    if any(x in t for x in ["rail","metro","station","transit"]): return "rail"
    if any(x in t for x in ["pharma","gmp","life sciences","cleanroom","lab"]): return "life"
    return "general"

def schedule_phases(model):
    return SPACE_PHASES if project_type_key(model)=="space" else EARTH_PHASES

def level_activity_count(level):
    return {1:6, 2:14, 3:28, 4:55, 5:95}.get(int(level),28)

def build_schedule_rows(model, level=None, scenario=None):
    level = int(level or selected_level(model))
    scenario = scenario or scenario_key(model)
    rule = SCENARIO_RULES.get(scenario, SCENARIO_RULES["base"])
    total_months = max(6, int(re.findall(r"\d+", str(model.get("schedule") or "60"))[0]) if re.findall(r"\d+", str(model.get("schedule") or "60")) else 60)
    duration_mult = rule["duration"]
    phases = schedule_phases(model)
    n = level_activity_count(level)
    rows=[]; pred=""
    for i in range(n):
        phase, base = phases[i % len(phases)]
        detail_suffix = ""
        if level >= 3: detail_suffix = f" — package {1 + i//len(phases)}"
        if level >= 4: detail_suffix += " / controls activity"
        if level >= 5: detail_suffix += f" / workface {1 + (i%4)}"
        act_id = f"L{level}-{(i+1):03d}"
        dur = max(1, round((total_months * duration_mult / max(1,n)) * (0.75 + (i%5)*0.10), 1))
        if i==0: name = "Project start / mobilisation"
        elif i==n-1: name = "Project complete / operational handover"
        else: name = base + detail_suffix
        critical = "Yes" if i in list(range(max(0,n-6), n)) or phase in ["Procurement","Systems","Integration","Commissioning","Launch Prep","Deployment"] else "No"
        rows.append({"level":level,"activity_id":act_id,"phase":phase,"activity":name,"predecessor":pred,"duration_months":dur,"duration_days":int(round(dur*21.75)),"critical":critical,"scenario":SCENARIO_RULES[scenario]["label"],"basis":f"Level {level} schedule generated for {clean(model.get('title'))}; scenario {SCENARIO_RULES[scenario]['label']} applies {duration_mult:.2f} duration factor."})
        pred=act_id
    return rows

def all_schedule_levels(model):
    return {lvl: build_schedule_rows(model,lvl,scenario_key(model)) for lvl in range(1,6)}

def normalize_cost_lines(model, scenario=None, cls=None):
    scenario = scenario or scenario_key(model); cls = int(cls or selected_class(model))
    base = list(model.get("cost_lines") or model.get("cost_breakdown") or [])
    if not base:
        base = [{"cbs":"01.01","description":"Project works","type":"Direct","p50_bn":as_float_money(model.get("cost_p50"))*0.7,"basis":"Generated project works allowance"}, {"cbs":"02.01","description":"Indirects and preliminaries","type":"Indirect","p50_bn":as_float_money(model.get("cost_p50"))*0.18,"basis":"Generated indirects allowance"}, {"cbs":"03.01","description":"Risk reserve","type":"Reserve","p50_bn":as_float_money(model.get("cost_p50"))*0.12,"basis":"Generated reserve allowance"}]
    rule = CLASS_RULES.get(cls, CLASS_RULES[3])
    scen = SCENARIO_RULES.get(scenario, SCENARIO_RULES["base"])
    out=[]
    for r in base:
        p50 = float(r.get("p50_bn") or as_float_money(r.get("p50") or 0)) * scen["cost"] / SCENARIO_RULES.get(scenario_key(model), SCENARIO_RULES["base"])["cost"] if False else float(r.get("p50_bn") or as_float_money(r.get("p50") or 0))
        # model already scenario-adjusted; class ranges adjust P10/P90 only
        p10 = p50 * rule["low"]
        p90 = p50 * rule["high"]
        p80 = p50 + (p90-p50)*0.62
        out.append({"class":cls,"class_name":rule["name"],"maturity":rule["maturity"],"cbs":r.get("cbs",""),"description":r.get("description",""),"type":r.get("type","Direct"),"p10_bn":round(p10,3),"p50_bn":round(p50,3),"p80_bn":round(p80,3),"p90_bn":round(p90,3),"scenario":SCENARIO_RULES[scenario]["label"],"basis":r.get("impact_basis") or r.get("basis") or rule["basis"]})
    return out

def all_class_costs(model):
    return {cls: normalize_cost_lines(model, scenario_key(model), cls) for cls in range(1,6)}

def build_risks(model):
    base = list(model.get("risks") or model.get("risk_register") or [])
    pkey = project_type_key(model)
    titles=[]
    for r in base: titles.append(r.get("title") or r.get("risk") or "Risk")
    titles += SECTOR_RISKS.get(pkey, []) + SECTOR_RISKS["general"]
    # expand to at least 30 risks with meaningful variants
    extra_templates = [
        "Commercial claim exposure", "Contractor insolvency", "Design interface gap", "Late client decision", "Authority condition change", "Logistics route constraint", "Material availability constraint", "Labour productivity variance", "Weather / environment disruption", "Digital systems integration", "Testing failure", "Operational readiness gap", "Handover documentation delay", "Cyber/security approval", "Stakeholder objection", "Quality defect rework", "Site access restriction", "Utilities diversion conflict", "Supplier quality failure", "Change control backlog", "Escalation beyond allowance", "Insurance / bonding constraint", "Commissioning resources shortage", "Data handover failure", "Training / mobilisation gap"
    ]
    titles += extra_templates
    seen=[]
    for t in titles:
        if t and t not in seen: seen.append(t)
    scen = SCENARIO_RULES.get(scenario_key(model), SCENARIO_RULES["base"])
    out=[]
    for i,t in enumerate(seen[:max(30, len(seen))]):
        rb = base[i] if i < len(base) else {}
        prob = rb.get("probability_pct") or int(18 + (i*7)%55)
        prob = max(5, min(85, int(round(prob * scen["risk"]))))
        cost_ml = rb.get("cost_m_bn") or rb.get("cost_ml_bn") or round(as_float_money(model.get("cost_p50"))*0.012*(1+(i%6)*0.22)*scen["risk"],3)
        cost_o = rb.get("cost_o_bn") or rb.get("cost_optimistic_bn") or round(cost_ml*0.35,3)
        cost_p = rb.get("cost_p_bn") or rb.get("cost_pessimistic_bn") or round(cost_ml*2.25,3)
        sched_ml = rb.get("schedule_m_days") or rb.get("schedule_ml_days") or int(10 + (i%9)*7*scen["risk"])
        sched_o = rb.get("schedule_o_days") or rb.get("schedule_optimistic_days") or max(1,int(sched_ml*0.35))
        sched_p = rb.get("schedule_p_days") or rb.get("schedule_pessimistic_days") or max(sched_ml+1,int(sched_ml*2.35))
        owner = rb.get("owner") or ["Project Director","Commercial Lead","Controls Lead","Design Manager","Procurement Lead","Delivery Lead","Commissioning Lead","Risk Manager"][i%8]
        cause = rb.get("cause") or f"{t} driven by project maturity, supply chain, interface or approval uncertainty under the {scen['label']} scenario."
        event = rb.get("risk_event") or rb.get("event") or f"{t} materialises and affects the active delivery baseline."
        impact = rb.get("impact_description") or rb.get("impact") or f"Potential cost growth, schedule movement, rework, contingency drawdown or delayed benefit realisation."
        mitigation = rb.get("mitigation") or f"Assign owner, validate basis, secure evidence, monitor trigger weekly and implement mitigation before next gate."
        out.append({"risk_id": rb.get("risk_id") or f"R-{i+1:03d}", "category": rb.get("category") or ("Space" if pkey=="space" else "Delivery"), "title": t, "cause": cause, "event": event, "impact": impact, "owner": owner, "mitigation": mitigation, "trigger": rb.get("trigger") or f"Trigger threshold for {t.lower()} breached", "probability_pct": prob, "cost_optimistic_bn": round(float(cost_o),3), "cost_ml_bn": round(float(cost_ml),3), "cost_pessimistic_bn": round(float(cost_p),3), "schedule_optimistic_days": int(sched_o), "schedule_ml_days": int(sched_ml), "schedule_pessimistic_days": int(sched_p), "cost_emv_bn": round(float(cost_ml)*prob/100,3), "schedule_emv_days": round(float(sched_ml)*prob/100,1), "status":"Open", "board_visibility":"Yes" if i<12 else "No", "scenario":scen["label"]})
    out.sort(key=lambda r: (r["cost_emv_bn"], r["schedule_emv_days"]), reverse=True)
    return out

def curves(model):
    p50=as_float_money(model.get("cost_p50")); sched=int(re.findall(r"\d+",str(model.get("schedule") or "60"))[0]) if re.findall(r"\d+",str(model.get("schedule") or "60")) else 60
    cls=selected_class(model); rule=CLASS_RULES[cls]; scen=SCENARIO_RULES[scenario_key(model)]
    rows=[]
    for pct in [1,5,10,20,30,40,50,60,70,80,85,90,95,99]:
        x=(pct-50)/50
        cost=p50*(1 + (rule["high"]-1)*max(0,x) + (1-rule["low"])*min(0,x))
        months=sched*(1 + (0.30*scen["risk"])*max(0,x) + (0.15)*min(0,x))
        rows.append({"percentile":pct,"qcra_cost_bn":round(cost,3),"qsra_months":round(months,2)})
    return rows

def benchmark_rows(model):
    pkey=project_type_key(model); p50=as_float_money(model.get("cost_p50")); sched=str(model.get("schedule"))
    if pkey=="data":
        return [["Cost per MW equivalent", "Model implies sector-scale hyperscale capital intensity", "$8M–$18M/MW broad benchmark", "Requires power/cooling validation", "Benchmark changes materially with grid scope, cooling and land."], ["Schedule benchmark", sched, "30–60+ months", "Grid connection likely critical", "Faster cases require explicit acceleration premium."]]
    if pkey=="airport": return [["Cost per passenger capacity", money(p50), "Varies by terminal/runway/baggage scope", "Validate airside phasing", "Airport estimates are dominated by live operations and systems integration."], ["Schedule benchmark", sched, "60–120+ months", "Possessions and approvals likely critical", "Runway and live terminal interfaces drive QSRA."]]
    if pkey=="space": return [["Mission complexity", money(p50), "Frontier programmes have extreme range", "Validate TRL and launch basis", "Mass margin, launch window and qualification dominate."], ["Schedule benchmark", sched, "60–300+ months", "Scenario-dependent", "Qualification and launch cadence drive QSRA."]]
    return [["Capital intensity", money(p50), "Sector benchmark required", "Use as first-pass challenge", "Benchmark basis must be validated with market evidence."], ["Schedule benchmark", sched, "Sector-specific", "Validate critical path", "Schedule logic should be confirmed before commitment."]]

def basis_rows(model):
    cls=CLASS_RULES[selected_class(model)]; scen=SCENARIO_RULES[scenario_key(model)]
    return [["Scenario", f"{scen['label']} — {scen['tone']}"], ["Estimate class", f"{cls['name']} ({cls['maturity']}). {cls['basis']}"], ["Schedule level", f"Level {selected_level(model)} schedule generated with {level_activity_count(selected_level(model))} activities and scenario duration factor {scen['duration']:.2f}."], ["Cost method", "Parametric CBS model using direct, indirect and reserve categories, adjusted for sector, location, scale, class and selected scenario."], ["Risk method", "Risk register generated with probability, O/M/P cost and schedule impacts, EMV, trigger, owner and mitigation. Zero likelihood is not permitted."], ["QCRA/QSRA", "Cost and schedule uncertainty are separated. Curves and tornadoes use distinct cost and duration exposure drivers."], ["Limitations", "First-pass intelligence model for challenge and decision support. Validate with design, supplier, programme and project-specific evidence before capital commitment."]]

def commercial_rows(model):
    scen=SCENARIO_RULES[scenario_key(model)]
    return [["Recommended option", scen["decision"]], ["Why", scen["tone"]], ["Board ask", "Approve next definition gate, not unconditional full capital commitment."], ["Buyer next step", "Run the model on the buyer’s real estimate, schedule and risk register to expose gaps and benchmark variance."], ["Commercial hook", "CASEY creates a first-pass board-grade project controls pack in minutes, reducing dependency on weeks of manual advisory preparation."]]

def xlsx_formats(wb):
    navy="#06111F"; blue="#0070C0"; cyan="#DDFBFF"; grey="#F5F7FA"; white="#FFFFFF"; black="#111827"; green="#E2F0D9"; amber="#FFF2CC"; red="#FCE4D6"
    return {
        "title": wb.add_format({"bold":True,"font_size":20,"font_color":white,"bg_color":navy,"align":"left","valign":"vcenter"}),
        "subtitle": wb.add_format({"font_size":10,"font_color":"#475569","bg_color":white,"text_wrap":True}),
        "section": wb.add_format({"bold":True,"font_size":12,"font_color":white,"bg_color":blue,"border":1,"border_color":"#D9E2F3"}),
        "header": wb.add_format({"bold":True,"font_color":white,"bg_color":blue,"border":1,"border_color":"#D9E2F3","text_wrap":True,"align":"center","valign":"vcenter"}),
        "cell": wb.add_format({"font_color":black,"bg_color":white,"border":1,"border_color":"#E2E8F0","text_wrap":True,"valign":"top"}),
        "alt": wb.add_format({"font_color":black,"bg_color":grey,"border":1,"border_color":"#E2E8F0","text_wrap":True,"valign":"top"}),
        "money": wb.add_format({"font_color":black,"bg_color":white,"border":1,"border_color":"#E2E8F0","num_format":"$0.0B","valign":"top"}),
        "num": wb.add_format({"font_color":black,"bg_color":white,"border":1,"border_color":"#E2E8F0","num_format":"0.0","valign":"top"}),
        "pct": wb.add_format({"font_color":black,"bg_color":white,"border":1,"border_color":"#E2E8F0","num_format":"0%","valign":"top"}),
        "kpi": wb.add_format({"bold":True,"font_size":18,"font_color":black,"bg_color":cyan,"border":1,"border_color":"#9DECF5","align":"center","valign":"vcenter"}),
        "good": wb.add_format({"font_color":black,"bg_color":green,"border":1,"border_color":"#E2E8F0","text_wrap":True}),
        "warn": wb.add_format({"font_color":black,"bg_color":amber,"border":1,"border_color":"#E2E8F0","text_wrap":True}),
        "bad": wb.add_format({"font_color":black,"bg_color":red,"border":1,"border_color":"#E2E8F0","text_wrap":True}),
    }

def write_table(ws, row, col, headers, rows, fmt, widths=None, money_cols=None, num_cols=None, pct_cols=None):
    money_cols=money_cols or set(); num_cols=num_cols or set(); pct_cols=pct_cols or set()
    for j,h in enumerate(headers): ws.write(row,col+j,h,fmt["header"])
    for i,r in enumerate(rows, row+1):
        for j,v in enumerate(r):
            f = fmt["cell"] if (i-row)%2 else fmt["alt"]
            if j in money_cols: f=fmt["money"]
            elif j in pct_cols: f=fmt["pct"]
            elif j in num_cols: f=fmt["num"]
            ws.write(i,col+j,v,f)
    if widths:
        for j,w in enumerate(widths): ws.set_column(col+j,col+j,w)
    ws.autofilter(row,col,row+len(rows),col+len(headers)-1)

def cost_workbook_bytes(model):
    if xlsxwriter is None: raise RuntimeError("xlsxwriter not installed")
    bio=BytesIO(); wb=xlsxwriter.Workbook(bio, {"in_memory": True, "constant_memory": False, "strings_to_urls": False})
    fmt=xlsx_formats(wb); scenario=scenario_key(model); scen=SCENARIO_RULES[scenario]; cls=selected_class(model); lvl=selected_level(model)
    risks=build_risks(model); qcurve=curves(model); cost_sel=normalize_cost_lines(model, scenario, cls); allcosts=all_class_costs(model); schedules=all_schedule_levels(model)
    qcra=sorted(risks, key=lambda r:r["cost_emv_bn"], reverse=True); qsra=sorted(risks, key=lambda r:r["schedule_emv_days"], reverse=True)
    title=clean(model.get("title"),80)
    # control centre
    ws=wb.add_worksheet("00 Control Centre"); ws.hide_gridlines(2); ws.set_zoom(90); ws.set_row(0,28)
    ws.merge_range("A1:H1", f"CASEY {VERSION.upper()} COMMERCIAL CONTROL PACK — {title}", fmt["title"])
    ws.write("A2", f"Scenario: {scen['label']} | Estimate Class: {cls} | Schedule Level: L{lvl} | Project Type: {model.get('mode')} / {model.get('subsector')} | Generated {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}", fmt["subtitle"])
    kpis=[("P50 Cost",model.get("cost_p50")),("Cost Range",model.get("cost_range")),("Schedule",model.get("schedule")),("Risk",model.get("risk")),("Confidence",f"{model.get('confidence_pct')}%"),("Recommended Option",scen["label"]),("Selected Class",CLASS_RULES[cls]["name"]),("Selected Level",f"L{lvl} / {level_activity_count(lvl)} activities")]
    for i,(k,v) in enumerate(kpis):
        r=4+(i//4)*3; c=(i%4)*2
        ws.write(r,c,k,fmt["section"]); ws.write(r+1,c,v,fmt["kpi"]); ws.write_blank(r+1,c+1,None,fmt["kpi"])
    write_table(ws,12,0,["Decision Area","CASEY Position"],commercial_rows(model),fmt,[28,110])
    ws.set_column("A:H",18)
    # scenario compare
    ws=wb.add_worksheet("01 Scenario Comparison"); ws.hide_gridlines(2); ws.merge_range("A1:K1","SCENARIO COMPARISON — COST / SCHEDULE / RISK TRADE-OFFS",fmt["title"])
    rows=[]
    p50=as_float_money(model.get("cost_p50")); sched=int(re.findall(r"\d+", str(model.get("schedule") or "60"))[0]) if re.findall(r"\d+", str(model.get("schedule") or "60")) else 60
    base_key=scenario
    for k,r in SCENARIO_RULES.items():
        # show relative to selected run baseline for buyer clarity
        rows.append([r["label"], round(p50*r["cost"]/scen["cost"],2), f"{money(p50*r['cost']/scen['cost']*CLASS_RULES[cls]['low'])} - {money(p50*r['cost']/scen['cost']*CLASS_RULES[cls]['high'])}", round(sched*r["duration"]/scen["duration"],1), "Higher" if r["risk"]>1.15 else ("Lower" if r["risk"]<0.85 else "Medium"), max(0.05,min(0.95,(float(model.get("confidence_pct") or 60)+r["confidence"]-scen["confidence"])/100)), r["tone"], r["decision"], "YES" if k==base_key else "", f"Cost x{r['cost']:.2f}; Time x{r['duration']:.2f}; Risk x{r['risk']:.2f}"])
    write_table(ws,4,0,["Scenario","P50 Cost BN","Range","Schedule Months","Risk Direction","Confidence","What Changes","Board Decision","Selected","Rule"],rows,fmt,[14,14,22,16,18,13,44,44,12,26],money_cols={1},num_cols={3},pct_cols={5})
    chart=wb.add_chart({"type":"column"}); chart.add_series({"name":"P50 Cost BN","categories":"='01 Scenario Comparison'!$A$6:$A$12","values":"='01 Scenario Comparison'!$B$6:$B$12","fill":{"color":"#0070C0"}}); chart.set_title({"name":"Scenario Cost"}); chart.set_legend({"none":True}); ws.insert_chart("A15",chart,{"x_scale":1.15,"y_scale":1.05})
    chart2=wb.add_chart({"type":"line"}); chart2.add_series({"name":"Schedule Months","categories":"='01 Scenario Comparison'!$A$6:$A$12","values":"='01 Scenario Comparison'!$D$6:$D$12","line":{"color":"#00A6C8","width":2.25}}); chart2.set_title({"name":"Scenario Schedule"}); chart2.set_legend({"none":True}); ws.insert_chart("H15",chart2,{"x_scale":1.15,"y_scale":1.05})
    # selected cost
    ws=wb.add_worksheet("02 Selected Class Cost"); ws.hide_gridlines(2); ws.merge_range("A1:J1",f"SELECTED CLASS {cls} COST ESTIMATE — {scen['label']} SCENARIO",fmt["title"])
    rows=[[r["class"],r["cbs"],r["description"],r["type"],r["p10_bn"],r["p50_bn"],r["p80_bn"],r["p90_bn"],r["scenario"],r["basis"]] for r in cost_sel]
    write_table(ws,4,0,["Class","CBS","Description","Type","P10 BN","P50 BN","P80 BN","P90 BN","Scenario","Basis"],rows,fmt,[8,12,32,14,12,12,12,12,16,70],money_cols={4,5,6,7})
    chart=wb.add_chart({"type":"bar"}); chart.add_series({"name":"P50 BN","categories":f"='02 Selected Class Cost'!$C$6:$C${5+len(rows)}","values":f"='02 Selected Class Cost'!$F$6:$F${5+len(rows)}","fill":{"color":"#0070C0"}}); chart.set_title({"name":"Selected Class Cost Build-Up"}); chart.set_legend({"none":True}); ws.insert_chart("L4",chart,{"x_scale":1.2,"y_scale":1.35})
    # all classes
    ws=wb.add_worksheet("03 All Estimate Classes"); ws.hide_gridlines(2); ws.merge_range("A1:J1",f"ALL ESTIMATE CLASSES — {scen['label']} SCENARIO",fmt["title"])
    rows=[]
    for c in range(1,6):
        for r in allcosts[c]: rows.append([c,CLASS_RULES[c]["name"],r["cbs"],r["description"],r["type"],r["p10_bn"],r["p50_bn"],r["p80_bn"],r["p90_bn"],"SELECTED" if c==cls else ""])
    write_table(ws,4,0,["Class","Class Name","CBS","Description","Type","P10 BN","P50 BN","P80 BN","P90 BN","Selected"],rows,fmt,[8,26,12,32,14,12,12,12,12,12],money_cols={5,6,7,8})
    # schedule levels
    ws=wb.add_worksheet("04 Schedule Levels"); ws.hide_gridlines(2); ws.merge_range("A1:I1",f"ALL SCHEDULE LEVELS — SELECTED L{lvl} — {scen['label']} SCENARIO",fmt["title"])
    rows=[]
    for lv,acts in schedules.items():
        for a in acts: rows.append([lv,a["activity_id"],a["phase"],a["activity"],a["predecessor"],a["duration_months"],a["duration_days"],a["critical"],"SELECTED" if lv==lvl else ""])
    write_table(ws,4,0,["Level","Activity ID","Phase","Activity","Predecessor","Duration Months","Duration Days","Critical","Selected"],rows,fmt,[8,14,22,48,18,16,14,12,12],num_cols={0,5,6})
    # risk register
    ws=wb.add_worksheet("05 Full Risk Register"); ws.hide_gridlines(2); ws.merge_range("A1:S1",f"FULL RISK REGISTER — {scen['label']} SCENARIO",fmt["title"])
    rows=[[r["risk_id"],r["category"],r["title"],r["cause"],r["event"],r["impact"],r["owner"],r["mitigation"],r["trigger"],r["probability_pct"]/100,r["cost_optimistic_bn"],r["cost_ml_bn"],r["cost_pessimistic_bn"],r["schedule_optimistic_days"],r["schedule_ml_days"],r["schedule_pessimistic_days"],r["cost_emv_bn"],r["schedule_emv_days"],r["board_visibility"]] for r in risks]
    write_table(ws,4,0,["ID","Category","Risk","Cause","Event","Impact","Owner","Mitigation","Trigger","Likelihood","Cost Low BN","Cost ML BN","Cost High BN","Sched Low Days","Sched ML Days","Sched High Days","Cost EMV BN","Sched EMV Days","Board"],rows,fmt,[10,14,28,40,40,48,18,48,30,12,12,12,12,14,14,14,12,14,10],money_cols={10,11,12,16},num_cols={13,14,15,17},pct_cols={9})
    # curves and tornado
    ws=wb.add_worksheet("06 QCRA QSRA Curves"); ws.hide_gridlines(2); ws.merge_range("A1:C1",f"QCRA AND QSRA CURVES — {scen['label']} SCENARIO",fmt["title"])
    rows=[[r["percentile"],r["qcra_cost_bn"],r["qsra_months"]] for r in qcurve]
    write_table(ws,4,0,["Percentile","QCRA Cost BN","QSRA Months"],rows,fmt,[12,16,16],money_cols={1},num_cols={0,2})
    chart=wb.add_chart({"type":"line"}); chart.add_series({"name":"QCRA Cost BN","categories":"='06 QCRA QSRA Curves'!$A$6:$A$19","values":"='06 QCRA QSRA Curves'!$B$6:$B$19","line":{"color":"#0070C0","width":2.25}}); chart.set_title({"name":"QCRA Cost S-Curve"}); ws.insert_chart("E4",chart,{"x_scale":1.2,"y_scale":1.1})
    chart2=wb.add_chart({"type":"line"}); chart2.add_series({"name":"QSRA Months","categories":"='06 QCRA QSRA Curves'!$A$6:$A$19","values":"='06 QCRA QSRA Curves'!$C$6:$C$19","line":{"color":"#00A6C8","width":2.25}}); chart2.set_title({"name":"QSRA Schedule S-Curve"}); ws.insert_chart("E21",chart2,{"x_scale":1.2,"y_scale":1.1})
    ws=wb.add_worksheet("07 Tornado Drivers"); ws.hide_gridlines(2); ws.merge_range("A1:M1",f"SEPARATE QCRA + QSRA TORNADO DRIVERS — {scen['label']} SCENARIO",fmt["title"])
    qrows=[[i+1,r["title"],r["risk_id"],r["cost_emv_bn"],r["owner"],r["mitigation"]] for i,r in enumerate(qcra[:15])]
    srows=[[i+1,r["title"],r["risk_id"],r["schedule_emv_days"],r["owner"],r["mitigation"]] for i,r in enumerate(qsra[:15])]
    write_table(ws,4,0,["Rank","QCRA Cost Driver","Risk ID","Cost EMV BN","Owner","Mitigation"],qrows,fmt,[8,30,12,14,18,45],money_cols={3},num_cols={0})
    write_table(ws,4,7,["Rank","QSRA Schedule Driver","Risk ID","Schedule EMV Days","Owner","Mitigation"],srows,fmt,[8,30,12,18,18,45],num_cols={0,3})
    chart=wb.add_chart({"type":"bar"}); chart.add_series({"name":"QCRA EMV","categories":f"='07 Tornado Drivers'!$B$6:$B${5+len(qrows)}","values":f"='07 Tornado Drivers'!$D$6:$D${5+len(qrows)}","fill":{"color":"#0070C0"}}); chart.set_title({"name":"QCRA Cost Tornado"}); chart.set_legend({"none":True}); ws.insert_chart("A24",chart,{"x_scale":1.2,"y_scale":1.3})
    chart2=wb.add_chart({"type":"bar"}); chart2.add_series({"name":"QSRA Days","categories":f"='07 Tornado Drivers'!$I$6:$I${5+len(srows)}","values":f"='07 Tornado Drivers'!$K$6:$K${5+len(srows)}","fill":{"color":"#00A6C8"}}); chart2.set_title({"name":"QSRA Schedule Tornado"}); chart2.set_legend({"none":True}); ws.insert_chart("H24",chart2,{"x_scale":1.2,"y_scale":1.3})
    # basis/benchmark/commercial
    ws=wb.add_worksheet("08 Basis of Estimate"); ws.hide_gridlines(2); ws.merge_range("A1:B1",f"BASIS OF ESTIMATE — {scen['label']} SCENARIO",fmt["title"]); write_table(ws,3,0,["Basis Area","Explanation"],basis_rows(model),fmt,[28,110])
    ws=wb.add_worksheet("09 Benchmark Validation"); ws.hide_gridlines(2); ws.merge_range("A1:E1",f"BENCHMARK VALIDATION — {scen['label']} SCENARIO",fmt["title"]); write_table(ws,3,0,["Benchmark","Project Value","Benchmark Range / Note","Status","Basis"],benchmark_rows(model),fmt,[28,22,55,36,55])
    ws=wb.add_worksheet("10 Commercial Next Steps"); ws.hide_gridlines(2); ws.merge_range("A1:B1",f"COMMERCIAL NEXT STEPS — {scen['label']} SCENARIO",fmt["title"]); write_table(ws,3,0,["Action","Detail"],commercial_rows(model),fmt,[28,110])
    wb.close(); bio.seek(0); return bio.getvalue()

def risk_workbook_bytes(model):
    if xlsxwriter is None: raise RuntimeError("xlsxwriter not installed")
    bio=BytesIO(); wb=xlsxwriter.Workbook(bio,{"in_memory":True,"strings_to_urls":False}); fmt=xlsx_formats(wb); scen=SCENARIO_RULES[scenario_key(model)]
    risks=build_risks(model); qcra=sorted(risks,key=lambda r:r["cost_emv_bn"],reverse=True); qsra=sorted(risks,key=lambda r:r["schedule_emv_days"],reverse=True)
    ws=wb.add_worksheet("00 Risk Control Centre"); ws.hide_gridlines(2); ws.merge_range("A1:H1",f"CASEY {VERSION.upper()} RISK REGISTER — {scen['label']} SCENARIO",fmt["title"])
    ws.write("A3","This workbook is the full risk register, not a top-10 extract. Every risk includes cause, event, impact, owner, mitigation, trigger, likelihood and quantified cost/schedule exposure.",fmt["subtitle"])
    write_table(ws,5,0,["Metric","Value"],[["Risk Count",len(risks)],["Top QCRA Driver",qcra[0]["title"]],["Top QSRA Driver",qsra[0]["title"]],["No Zero Likelihood","YES"],["Scenario",scen["label"]]],fmt,[28,60])
    ws=wb.add_worksheet("01 Full Risk Register"); ws.hide_gridlines(2); ws.merge_range("A1:S1",f"FULL RISK REGISTER — {scen['label']}",fmt["title"])
    rows=[[r["risk_id"],r["category"],r["title"],r["cause"],r["event"],r["impact"],r["owner"],r["mitigation"],r["trigger"],r["probability_pct"]/100,r["cost_optimistic_bn"],r["cost_ml_bn"],r["cost_pessimistic_bn"],r["schedule_optimistic_days"],r["schedule_ml_days"],r["schedule_pessimistic_days"],r["cost_emv_bn"],r["schedule_emv_days"],r["board_visibility"]] for r in risks]
    write_table(ws,4,0,["ID","Category","Risk","Cause","Event","Impact","Owner","Mitigation","Trigger","Likelihood","Cost Low BN","Cost ML BN","Cost High BN","Sched Low Days","Sched ML Days","Sched High Days","Cost EMV BN","Sched EMV Days","Board"],rows,fmt,[10,14,28,40,40,48,18,48,30,12,12,12,12,14,14,14,12,14,10],money_cols={10,11,12,16},num_cols={13,14,15,17},pct_cols={9})
    ws=wb.add_worksheet("02 QCRA Cost Tornado"); ws.hide_gridlines(2); ws.merge_range("A1:F1","QCRA COST TORNADO",fmt["title"]); write_table(ws,4,0,["Rank","Risk","ID","Cost EMV BN","Owner","Mitigation"],[[i+1,r["title"],r["risk_id"],r["cost_emv_bn"],r["owner"],r["mitigation"]] for i,r in enumerate(qcra[:20])],fmt,[8,32,12,14,20,55],money_cols={3},num_cols={0})
    ws=wb.add_worksheet("03 QSRA Schedule Tornado"); ws.hide_gridlines(2); ws.merge_range("A1:F1","QSRA SCHEDULE TORNADO",fmt["title"]); write_table(ws,4,0,["Rank","Risk","ID","Schedule EMV Days","Owner","Mitigation"],[[i+1,r["title"],r["risk_id"],r["schedule_emv_days"],r["owner"],r["mitigation"]] for i,r in enumerate(qsra[:20])],fmt,[8,32,12,18,20,55],num_cols={0,3})
    wb.close(); bio.seek(0); return bio.getvalue()

def schedule_csv_bytes(model):
    output=[]; import io
    buf=io.StringIO(); w=csv.writer(buf)
    w.writerow(["Level","Activity ID","Phase","Activity","Predecessor","Duration Months","Duration Days","Critical","Scenario","Basis"])
    for lvl,acts in all_schedule_levels(model).items():
        for a in acts: w.writerow([lvl,a["activity_id"],a["phase"],a["activity"],a["predecessor"],a["duration_months"],a["duration_days"],a["critical"],a["scenario"],a["basis"]])
    return buf.getvalue().encode("utf-8")

def model_json_bytes(model):
    enriched=dict(model); enriched["casey_version"]=VERSION; enriched["v62_schedule_levels"]=all_schedule_levels(model); enriched["v62_risks"]=build_risks(model); enriched["v62_basis"]=basis_rows(model); enriched["v62_benchmarks"]=benchmark_rows(model)
    return json.dumps(enriched, indent=2, default=str).encode("utf-8")


def _template_tables():
    template_path=os.path.join(os.path.dirname(__file__), "pra_template_L01-10.xer")
    if not os.path.exists(template_path): template_path="/mnt/data/L01-10.xer"
    raw=open(template_path,"rb").read().decode("latin1", errors="ignore")
    tables={}; order=[]; current=None
    for line in raw.splitlines():
        if line.startswith("%T\t"):
            current=line.split("\t",1)[1]; order.append(current); tables[current]={"header":line,"fields":[],"rows":[]}
        elif current and line.startswith("%F\t"):
            tables[current]["fields"]=line.split("\t")[1:]
        elif current and line.startswith("%R\t"):
            tables[current]["rows"].append(line.split("\t")[1:])
    return tables, order

def _safe_xer_text(v, maxlen=120):
    return clean(v, maxlen).replace("\t"," ").replace("\n"," ").replace("\r"," ")

def _row(fields, values, base=None):
    vals=list(base[:] if base else [""]*len(fields))
    if len(vals)<len(fields): vals += [""]*(len(fields)-len(vals))
    if len(vals)>len(fields): vals=vals[:len(fields)]
    idx={f:i for i,f in enumerate(fields)}
    for k,v in values.items():
        if k in idx: vals[idx[k]]=str(v)
    return "%R\t"+"\t".join(vals)

def patch_template_xer(model):
    """Build a real dynamic PRA-compatible XER.

    This no longer copies the Mondrian L01-10 activities. It reuses only the proven
    ERMHDR/version/table-field structure, then writes a new CASEY project, WBS,
    activity list, relationships and PRA O/M/P duration UDF values based on the
    selected project type, scenario and schedule level.
    """
    tables, order = _template_tables()
    scen_key=scenario_key(model); scen=SCENARIO_RULES[scen_key]
    lvl=selected_level(model); cls=selected_class(model); ptype=project_type_key(model)
    mode="Space" if ptype=="space" else "Earth"
    title=_safe_xer_text(model.get("title") or model.get("project_name") or ("Lunar Base" if mode=="Space" else "Capital Project"),80)
    proj_id=3338
    proj_short=_safe_xer_text(f"CASEY_{mode}_{scen['label']}_L{lvl}",32).replace(" ","_")[:32]
    proj_name=_safe_xer_text(f"CASEY {title} — {mode} — {scen['label']} Scenario — Level {lvl} Schedule",120)
    start_dt=datetime(2026,1,5,8,0)
    acts=build_schedule_rows(model,lvl,scen_key)
    # Map phases to WBS ids, plus root project WBS.
    phases=[]
    for a in acts:
        if a["phase"] not in phases: phases.append(a["phase"])
    root_wbs=35000
    wbs_ids={phase:35001+i for i,phase in enumerate(phases)}

    out=[]
    out.append("ERMHDR\t8.0\t2026-01-05\tProject\tadmin\tPPSEn\tdbxDatabaseNoName\tProject Management\tUSD")

    def emit_table(name, rows=None):
        if name not in tables: return
        out.append(tables[name]["header"])
        out.append("%F\t"+"\t".join(tables[name]["fields"]))
        if rows is None:
            for r in tables[name]["rows"]:
                out.append("%R\t"+"\t".join(r))
        else:
            out.extend(rows)

    # Safe static metadata tables that do not point to old activities.
    for name in ["CURRTYPE","OBS","UDFTYPE"]:
        emit_table(name)

    # PROJECT: one dynamic CASEY project row.
    pf=tables["PROJECT"]["fields"]; base_project=tables["PROJECT"]["rows"][0]
    emit_table("PROJECT", [_row(pf, {
        "proj_id":proj_id,
        "proj_short_name":proj_short,
        "clndr_id":3892,
        "wbs_max_sum_level":lvl,
        "plan_start_date":start_dt.strftime("%Y-%m-%d %H:%M"),
        "last_recalc_date":start_dt.strftime("%Y-%m-%d %H:%M"),
        "scd_end_date":(start_dt+timedelta(days=sum(a["duration_days"] for a in acts))).strftime("%Y-%m-%d %H:%M"),
        "guid":f"CASEY{mode}{scen['label']}L{lvl}"[:22],
        "add_by_name":"CASEY",
        "export_flag":"Y"
    }, base_project)])

    # CALENDAR and SCHEDOPTIONS: retain proven import-compatible records, but project ids remain stable.
    emit_table("CALENDAR")
    emit_table("SCHEDOPTIONS")

    # PROJWBS: dynamic root and phase WBS only.
    wf=tables["PROJWBS"]["fields"]; base_wbs=tables["PROJWBS"]["rows"][0]
    wbs_rows=[]
    wbs_rows.append(_row(wf, {"wbs_id":root_wbs,"proj_id":proj_id,"obs_id":565,"seq_num":0,"proj_node_flag":"Y","wbs_short_name":proj_short,"wbs_name":proj_name,"parent_wbs_id":"","guid":f"CASEY_ROOT_{proj_short}"[:22]}, base_wbs))
    for i,phase in enumerate(phases):
        wbs_rows.append(_row(wf, {"wbs_id":wbs_ids[phase],"proj_id":proj_id,"obs_id":565,"seq_num":i+1,"proj_node_flag":"N","wbs_short_name":f"{i+1:02d}","wbs_name":_safe_xer_text(phase,60),"parent_wbs_id":root_wbs,"guid":f"CASEY_WBS_{i+1:02d}_{phase}"[:22]}, base_wbs))
    emit_table("PROJWBS", wbs_rows)

    # TASK: dynamic activity list. Count changes by level and names change by Earth/Space/scenario/project.
    tf=tables["TASK"]["fields"]; base_task=tables["TASK"]["rows"][0]
    task_rows=[]; id_map={}; current=start_dt
    for i,a in enumerate(acts):
        tid=200000+i+1; id_map[a["activity_id"]]=tid
        dur_days=max(1,int(a["duration_days"])); dur_hr=dur_days*8
        s=current; e=current+timedelta(days=dur_days)
        current=e+timedelta(days=1)
        status="TK_NotStart"
        task_rows.append(_row(tf, {
            "task_id":tid,"proj_id":proj_id,"wbs_id":wbs_ids[a["phase"]],"clndr_id":3892,
            "phys_complete_pct":"0","est_wt":"1","complete_pct_type":"CP_Drtn","task_type":"TT_Task","duration_type":"DT_FixedRate","status_code":status,
            "task_code":a["activity_id"],"task_name":_safe_xer_text(a["activity"],120),
            "total_float_hr_cnt":"0" if a["critical"]=="Yes" else str(80+(i%5)*24),"free_float_hr_cnt":"0" if a["critical"]=="Yes" else "40",
            "remain_drtn_hr_cnt":dur_hr,"target_drtn_hr_cnt":dur_hr,
            "target_start_date":s.strftime("%Y-%m-%d %H:%M"),"target_end_date":e.strftime("%Y-%m-%d %H:%M"),
            "early_start_date":s.strftime("%Y-%m-%d %H:%M"),"early_end_date":e.strftime("%Y-%m-%d %H:%M"),
            "late_start_date":s.strftime("%Y-%m-%d %H:%M"),"late_end_date":e.strftime("%Y-%m-%d %H:%M"),
            "restart_date":s.strftime("%Y-%m-%d %H:%M"),"reend_date":e.strftime("%Y-%m-%d %H:%M"),
            "driving_path_flag":"Y" if a["critical"]=="Yes" else "N",
            "guid":f"CASEY_TASK_{lvl}_{i+1:03d}"[:22],
            "create_user":"CASEY","update_user":"CASEY","create_date":start_dt.strftime("%Y-%m-%d %H:%M"),"update_date":start_dt.strftime("%Y-%m-%d %H:%M")
        }, base_task))
    emit_table("TASK", task_rows)

    # TASKPRED: explicit FS chain so PRA/P6 sees logic, not a loose activity list.
    if "TASKPRED" in tables:
        pfld=tables["TASKPRED"]["fields"]; base_pred=tables["TASKPRED"]["rows"][0] if tables["TASKPRED"]["rows"] else None
        pred_rows=[]
        for i in range(1,len(acts)):
            pred_rows.append(_row(pfld, {"task_pred_id":260000+i,"task_id":200000+i+1,"pred_task_id":200000+i,"proj_id":proj_id,"pred_proj_id":proj_id,"pred_type":"PR_FS","lag_hr_cnt":0,"float_path":"","aref":"","arls":""}, base_pred))
        emit_table("TASKPRED", pred_rows)

    # UDFVALUE: PRA uncertainty duration fields Likely/Max/Min against every dynamic task.
    if "UDFVALUE" in tables:
        uf=tables["UDFVALUE"]["fields"]; udf_rows=[]
        for i,a in enumerate(acts):
            tid=200000+i+1; d=max(1,int(a["duration_days"]))
            values=[(329,d),(330,max(d+1,int(d*1.35))),(331,max(1,int(d*0.75)))]
            for udf_id,val in values:
                udf_rows.append(_row(uf,{"udf_type_id":udf_id,"fk_id":tid,"proj_id":proj_id,"udf_text":val}, None))
        emit_table("UDFVALUE", udf_rows)

    text="\n".join(out)+"\n"
    # Hard guard: no Mondrian/L01-10 remnants in dynamic schedule content.
    for bad in ["Mondrian","Lilly Mondrian","L01-10","Shire"]:
        text=text.replace(bad, proj_short if bad=="L01-10" else "CASEY")
    return text.encode("latin1", errors="ignore")

def stream(data, media, filename):
    return StreamingResponse(BytesIO(data), media_type=media, headers={"Content-Disposition": f'attachment; filename="{filename}"'})

def all_zip_bytes(model):
    bio=BytesIO()
    with zipfile.ZipFile(bio,"w",zipfile.ZIP_DEFLATED) as z:
        z.writestr("CASEY_v62_Cost_Model.xlsx", cost_workbook_bytes(model))
        z.writestr("CASEY_v62_Risk_Register.xlsx", risk_workbook_bytes(model))
        z.writestr("CASEY_v62_Schedule.xer", patch_template_xer(model))
        z.writestr("CASEY_v62_Schedule_Levels.csv", schedule_csv_bytes(model))
        z.writestr("CASEY_v62_Model_Audit.json", model_json_bytes(model))
        z.writestr("README_v62.txt", "CASEY v62 outputs: dynamic selected scenario, estimate class and schedule level. XER preserves PRA-compatible ERMHDR 8.0 structure while replacing project/WBS/task content with CASEY schedule data. CSV schedule levels are included as audit/import fallback.\n")
    bio.seek(0); return bio.getvalue()

def install_v62(app):
    def replace_post(path, endpoint):
        app.router.routes=[r for r in app.router.routes if not (getattr(r,"path",None)==path and "POST" in getattr(r,"methods",set()))]
        app.post(path)(endpoint)
    def workbook(model): return stream(cost_workbook_bytes(model),"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","CASEY_v62_Cost_Model.xlsx")
    def risks(model): return stream(risk_workbook_bytes(model),"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","CASEY_v62_Risk_Register.xlsx")
    def xer(model): return stream(patch_template_xer(model),"application/octet-stream","CASEY_v62_Schedule.xer")
    def csvexp(model): return stream(schedule_csv_bytes(model),"text/csv","CASEY_v62_Schedule_Levels.csv")
    def jsonexp(model): return stream(model_json_bytes(model),"application/json","CASEY_v62_Model_Audit.json")
    def allpack(model): return stream(all_zip_bytes(model),"application/zip","CASEY_v62_Output_Pack.zip")
    def removed(model): return JSONResponse({"status":"not_in_primary_pack","message":"Primary source-of-truth pack is Excel cost model, Excel risk register, XER schedule, schedule CSV and JSON audit."},status_code=410)
    for p,e in [("/export/workbook",workbook),("/export/risk-register",risks),("/export/xer",xer),("/export/schedule-csv",csvexp),("/export/json",jsonexp),("/export/all",allpack),("/v62/export/all",allpack)]: replace_post(p,e)
    for p in ["/export/pdf","/export/pptx","/export/word"]: replace_post(p, removed)

# =========================
# V63 TRUE UNIVERSAL ENGINE OVERRIDES
# =========================
from fastapi import Body
VERSION = "FINAL"

UNIVERSAL_WBS = {
    "data": [
        ("Governance", ["Business case", "Funding gate", "Controls mobilisation"]),
        ("Site + Utilities", ["Land readiness", "Enabling works", "Grid connection", "Water and drainage", "Temporary power"]),
        ("Design", ["Basis of design", "White space design", "Power architecture", "Cooling architecture", "IFC release"]),
        ("Procurement", ["Long lead transformers", "Generators and UPS", "Cooling plant", "Racks and containment", "Controls systems"]),
        ("Construction", ["Civils", "Shell and core", "MEP install", "Data halls", "External works"]),
        ("Integration", ["BMS/EPMS integration", "Power testing", "Cooling testing", "IST", "Client systems"]),
        ("Commissioning", ["Level 1-5 commissioning", "Reliability run", "Handover", "Operational readiness"]),
    ],
    "airport": [
        ("Governance", ["Business case", "Aviation approvals", "Stakeholder strategy"]),
        ("Airside", ["Runway possessions", "Taxiway works", "Apron works", "Airfield ground lighting"]),
        ("Terminal", ["Terminal shell", "Passenger processing", "Retail/interface areas", "Security systems"]),
        ("Baggage", ["BHS design", "Conveyors", "Controls", "Integrated testing"]),
        ("Landside", ["Roads", "Car parks", "Rail/bus interfaces", "Utilities diversions"]),
        ("Systems", ["ICT", "Security", "Fire life safety", "Operational systems"]),
        ("ORAT", ["Trial operations", "Airline readiness", "Migration", "Opening"]),
    ],
    "port": [
        ("Governance", ["Port authority approvals", "Business case", "Stakeholders"]),
        ("Marine Works", ["Dredging", "Quay walls", "Berths", "Breakwater/sea defence"]),
        ("Yard", ["Pavement", "Drainage", "Stacking yards", "Utilities"]),
        ("Equipment", ["Cranes", "Automation", "Power supply", "Maintenance facilities"]),
        ("Rail/Road", ["Gate complex", "Customs", "Rail link", "Road access"]),
        ("Systems", ["Terminal operating system", "Security", "Customs integration", "Testing"]),
        ("Commissioning", ["Operational trials", "Throughput test", "Handover"]),
    ],
    "rail": [
        ("Governance", ["Business case", "Route approvals", "Possession strategy"]),
        ("Civil Route", ["Earthworks", "Structures", "Tunnels", "Drainage"]),
        ("Track", ["Track formation", "Slab/ballast", "Turnouts", "Welding"]),
        ("Stations", ["Station civils", "MEP", "Architectural fit-out", "Access systems"]),
        ("Systems", ["Signalling", "Telecoms", "Power", "SCADA"]),
        ("Testing", ["Static testing", "Dynamic testing", "Trial running", "Safety case"]),
        ("Entry Into Service", ["Operator readiness", "Handover", "Benefits realisation"]),
    ],
    "life": [
        ("Governance", ["User requirements", "Validation master plan", "Regulatory strategy"]),
        ("Design", ["Process basis", "Cleanroom classification", "Utilities design", "GMP review"]),
        ("Procurement", ["Process equipment", "Cleanroom systems", "Automation", "Long-lead utilities"]),
        ("Construction", ["Shell", "Cleanroom envelope", "MEP install", "Process install"]),
        ("CQV", ["Commissioning", "Qualification", "Validation", "Regulatory inspection"]),
        ("Operational Readiness", ["SOPs", "Training", "Tech transfer", "Start-up"]),
    ],
    "fab": [
        ("Governance", ["Investment gate", "Technology basis", "Supplier strategy"]),
        ("Site", ["Enabling", "Clean utilities", "Power", "Water/waste"]),
        ("Design", ["Cleanroom basis", "Process utilities", "Tool hook-up design", "IFC"]),
        ("Procurement", ["Long-lead tools", "Bulk gas", "UPW", "Electrical gear"]),
        ("Construction", ["Civil/structure", "Cleanroom", "MEP", "Tool install"]),
        ("Ramp", ["Commissioning", "Tool qualification", "Yield ramp", "Operations"]),
    ],
    "energy": [
        ("Governance", ["Business case", "Grid/utility agreements", "Permits"]),
        ("Development", ["Land", "Surveys", "Environmental", "Stakeholders"]),
        ("Design", ["FEED", "Interconnection", "Balance of plant", "IFC"]),
        ("Procurement", ["Major equipment", "EPC tender", "Grid equipment", "Spares"]),
        ("Construction", ["Civil works", "Equipment install", "Electrical", "Grid tie-in"]),
        ("Commissioning", ["Cold commissioning", "Hot commissioning", "Performance test", "COD"]),
    ],
    "space": [
        ("Mission", ["Mission objectives", "CONOPS", "Requirements", "Success criteria"]),
        ("Architecture", ["Systems architecture", "Interfaces", "Mass budget", "Power/thermal basis"]),
        ("Technology", ["TRL closure", "Prototype", "Qualification plan", "Critical tests"]),
        ("Design", ["PDR", "CDR", "Flight/surface design", "Safety review"]),
        ("Supply Chain", ["Payload procurement", "Launch procurement", "Habitat systems", "Spares/logistics"]),
        ("AIT", ["Manufacturing", "Assembly", "Integration", "Environmental test"]),
        ("Launch + Transit", ["Launch campaign", "Range approval", "Transfer", "Insertion/landing"]),
        ("Deployment", ["Site activation", "Systems deployment", "Commissioning", "Initial operations"]),
        ("Sustainment", ["Crew/robotic ops", "Maintenance", "Resupply", "Reliability growth"]),
    ],
    "general": [
        ("Governance", ["Business case", "Funding", "Controls mobilisation"]),
        ("Definition", ["Scope", "Surveys", "Design basis", "Approvals"]),
        ("Procurement", ["Strategy", "Tender", "Long leads", "Awards"]),
        ("Delivery", ["Enabling", "Civil works", "Systems", "Integration"]),
        ("Commissioning", ["Testing", "Operational readiness", "Handover"]),
    ],
}

def _payload_model(payload):
    """Accept either the raw CASEY model or frontend wrapper payload."""
    if payload is None:
        payload = {}
    if not isinstance(payload, dict):
        payload = {}
    base = {}
    if isinstance(payload.get("project"), dict):
        base.update(payload.get("project") or {})
    if isinstance(payload.get("model"), dict):
        base.update(payload.get("model") or {})
    base.update({k:v for k,v in payload.items() if k not in ["project","model"]})
    # normalise common frontend names
    if "level" in base and "schedule_level" not in base:
        base["schedule_level"] = base["level"]
    if "estimate_class" in base and "class_level" not in base:
        base["class_level"] = base["estimate_class"]
    if "sector" in base and "mode" not in base:
        base["mode"] = base["sector"]
    if not base.get("title"):
        prompt = str(base.get("prompt") or base.get("project_command") or "Capital Project")
        # simple title extraction
        base["title"] = clean(prompt[:70], 70)
    if not base.get("scenario"):
        base["scenario"] = "base"
    if not base.get("schedule_level"):
        base["schedule_level"] = 3
    if not base.get("class_level"):
        base["class_level"] = 3
    if not base.get("cost_p50"):
        key = _universal_project_key(base)
        defaults = {"data":10.7,"airport":8.5,"port":5.2,"rail":6.8,"life":2.4,"fab":7.8,"energy":4.6,"space":32.0,"general":1.8}
        base["cost_p50"] = defaults.get(key,1.8)
    if not base.get("schedule"):
        key = _universal_project_key(base)
        defaults = {"data":"64 months","airport":"84 months","port":"66 months","rail":"90 months","life":"54 months","fab":"72 months","energy":"60 months","space":"156 months","general":"48 months"}
        base["schedule"] = defaults.get(key,"48 months")
    return base

def _all_text(model):
    return " ".join(str(model.get(k, "")) for k in ["prompt","title","project_name","sector","mode","subsector","description"]).lower()

def _universal_project_key(model):
    t=_all_text(model)
    if any(w in t for w in ["space","moon","lunar","mars","orbit","orbital","satellite","launch","habitat","isru","asteroid"]): return "space"
    if any(w in t for w in ["airport","runway","terminal","airside","baggage"]): return "airport"
    if any(w in t for w in ["port","harbour","harbor","quay","berth","container terminal","marine"]): return "port"
    if any(w in t for w in ["rail","metro","station","transit","signalling","signaling","hs2"]): return "rail"
    if any(w in t for w in ["pharma","gmp","biotech","cleanroom","life sciences","laboratory","lab"]): return "life"
    if any(w in t for w in ["semiconductor","fab","gigafactory","factory","manufacturing","battery"]): return "fab"
    if any(w in t for w in ["energy","hydrogen","wind","solar","nuclear","grid","power plant","substation"]): return "energy"
    if any(w in t for w in ["data centre","data center","datacenter","hyperscale","ai campus","cloud"]): return "data"
    return "general"

# Override project type detection used by risk/benchmark/schedule functions.
def project_type_key(model):
    return _universal_project_key(_payload_model(model))

def schedule_phases(model):
    m=_payload_model(model)
    key=_universal_project_key(m)
    return UNIVERSAL_WBS.get(key, UNIVERSAL_WBS["general"])

def _phase_packages_for_level(model, level):
    """Return (phase, package/task) list. L2 phases, L3 packages, L4 controls tasks, L5 workfaces."""
    wbs = schedule_phases(model)
    pairs=[]
    if level <= 1:
        for phase, pkgs in wbs[:max(4,min(6,len(wbs)))]:
            pairs.append((phase, f"{phase} milestone"))
    elif level == 2:
        for phase, pkgs in wbs:
            pairs.append((phase, f"{phase} phase delivery"))
            # add a second key gate for complex sectors
            if len(wbs) <= 6 or phase in ["Procurement","Construction","Commissioning","AIT","Launch + Transit","Systems","Terminal","Airside"]:
                pairs.append((phase, f"{phase} control gate"))
    elif level == 3:
        for phase, pkgs in wbs:
            for p in pkgs[:3]: pairs.append((phase,p))
    elif level == 4:
        for phase, pkgs in wbs:
            for p in pkgs:
                pairs.append((phase,f"{p} — plan")); pairs.append((phase,f"{p} — execute"))
    else:
        for phase, pkgs in wbs:
            for p in pkgs:
                pairs.extend([(phase,f"{p} — engineering"),(phase,f"{p} — procurement"),(phase,f"{p} — installation"),(phase,f"{p} — test/accept")])
    # cap/ensure sensible counts
    target={1:6,2:14,3:28,4:55,5:95}.get(level,28)
    if len(pairs) < target:
        idx=0
        while len(pairs)<target:
            phase, pkgs=wbs[idx % len(wbs)]
            pairs.append((phase, f"{phase} additional control activity {len(pairs)+1}"))
            idx+=1
    return pairs[:target]

def build_schedule_rows(model, level=None, scenario=None):
    m=_payload_model(model)
    level=int(level or selected_level(m)); scenario=scenario or scenario_key(m)
    scen=SCENARIO_RULES.get(scenario,SCENARIO_RULES["base"])
    total_months=max(6, int(re.findall(r"\d+", str(m.get("schedule") or "60"))[0]) if re.findall(r"\d+", str(m.get("schedule") or "60")) else 60)
    # Strict scenario behaviour: faster overlaps/compresses; lower risk extends; cheaper modestly extends; premium has assurance duration.
    pairs=_phase_packages_for_level(m, level)
    n=len(pairs)
    rows=[]
    prev=""
    overlap = scenario in ["faster"] and level >= 3
    for i,(phase,task) in enumerate(pairs):
        act_id=f"L{level}-{i+1:03d}"
        # phase weighting: testing/commissioning/procurement tend to be longer/riskier
        weight=1.0 + (0.22 if phase.lower() in ["procurement","commissioning","integration","systems","ait","launch + transit","deployment"] else 0) + ((i%5)*0.04)
        dur=max(1, round((total_months*scen["duration"] / max(1,n))*weight, 1))
        if i==0: name="Project start / mobilisation"
        elif i==n-1: name="Project complete / operational handover"
        else: name=task
        if overlap and i>2 and i%4==0:
            pred=f"{prev};{f'L{level}-{max(1,i-2):03d}'}"
        else:
            pred=prev
        critical="Yes" if phase.lower() in ["procurement","commissioning","integration","systems","launch + transit","deployment","testing","orat","cqv"] or i>=n-8 else "No"
        rows.append({"level":level,"activity_id":act_id,"phase":phase,"activity":clean(name,120),"predecessor":pred,"duration_months":dur,"duration_days":int(round(dur*21.75)),"critical":critical,"scenario":scen["label"],"basis":f"{_universal_project_key(m).title()} WBS, Level {level}, {scen['label']} scenario. {scen['tone']}"})
        prev=act_id
    return rows

def all_schedule_levels(model):
    m=_payload_model(model)
    return {lvl: build_schedule_rows(m,lvl,scenario_key(m)) for lvl in range(1,6)}

# Override costs so selected scenario clearly affects all class estimates.
def normalize_cost_lines(model, scenario=None, cls=None):
    m=_payload_model(model); scenario=scenario or scenario_key(m); cls=int(cls or selected_class(m))
    key=_universal_project_key(m); scen=SCENARIO_RULES.get(scenario,SCENARIO_RULES["base"]); rule=CLASS_RULES.get(cls,CLASS_RULES[3])
    p50_total=as_float_money(m.get("cost_p50"))*scen["cost"]
    templates={
        "data":[("01.01","Site / enabling","Direct",.06),("01.02","Substructure / foundations","Direct",.08),("01.03","Shell / civil works","Direct",.12),("01.04","Power train / substations","Direct",.18),("01.05","Cooling systems","Direct",.14),("01.06","Data halls / white space","Direct",.13),("01.07","MEP / controls","Direct",.12),("01.08","IT/security/specialist systems","Direct",.06),("02.01","Preliminaries / general conditions","Indirect",.07),("02.02","Design / professional fees","Indirect",.05),("03.01","QCRA contingency","Reserve",.06),("03.02","Management reserve","Reserve",.03)],
        "airport":[("01.01","Airside/runway works","Direct",.16),("01.02","Terminal building","Direct",.22),("01.03","Baggage handling systems","Direct",.10),("01.04","Security/ICT/systems","Direct",.08),("01.05","Landside access/utilities","Direct",.14),("01.06","MEP/fire/life safety","Direct",.12),("02.01","Live operations/prelims","Indirect",.08),("02.02","Design/assurance/ORAT","Indirect",.06),("03.01","QCRA contingency","Reserve",.04)],
        "space":[("01.01","Mission architecture","Direct",.07),("01.02","Habitat/vehicle structures","Direct",.15),("01.03","Power/thermal/life support","Direct",.16),("01.04","Payload/science/operations systems","Direct",.10),("01.05","Launch/transit/logistics","Direct",.18),("01.06","Surface/orbital deployment","Direct",.13),("01.07","Qualification/testing","Direct",.10),("02.01","Programme assurance/safety","Indirect",.06),("02.02","Mission management/integration","Indirect",.05),("03.01","Frontier risk reserve","Reserve",.10)],
        "general":[("01.01","Enabling works","Direct",.10),("01.02","Civil/structural works","Direct",.22),("01.03","MEP/systems","Direct",.20),("01.04","Specialist equipment","Direct",.12),("01.05","Fit-out/finishes","Direct",.08),("02.01","Prelims/general conditions","Indirect",.09),("02.02","Design/professional fees","Indirect",.07),("03.01","QCRA contingency","Reserve",.08),("03.02","Management reserve","Reserve",.04)]
    }
    lines=templates.get(key, templates["general"])
    # normalize if template not 1.0
    s=sum(x[3] for x in lines)
    out=[]
    for cbs,desc,typ,p in lines:
        p50=p50_total*(p/s)
        p10=p50*rule["low"]; p90=p50*rule["high"]; p80=p50+(p90-p50)*0.62
        out.append({"class":cls,"class_name":rule["name"],"maturity":rule["maturity"],"cbs":cbs,"description":desc,"type":typ,"p10_bn":round(p10,3),"p50_bn":round(p50,3),"p80_bn":round(p80,3),"p90_bn":round(p90,3),"scenario":scen["label"],"basis":f"{key.title()} sector CBS, {scen['label']} scenario factor {scen['cost']:.2f}, {rule['name']} range."})
    return out

def all_class_costs(model):
    m=_payload_model(model)
    return {cls: normalize_cost_lines(m, scenario_key(m), cls) for cls in range(1,6)}

# V64 final commercial wrappers with Body(default={}) and hard model/export alignment.
def all_zip_bytes(model):
    m=_payload_model(model)
    bio=BytesIO()
    with zipfile.ZipFile(bio,"w",zipfile.ZIP_DEFLATED) as z:
        z.writestr("CASEY_FINAL_Cost_Model.xlsx", cost_workbook_bytes(m))
        z.writestr("CASEY_FINAL_Risk_Register.xlsx", risk_workbook_bytes(m))
        z.writestr("CASEY_FINAL_Schedule.xer", patch_template_xer(m))
        z.writestr("CASEY_FINAL_Schedule_Levels.csv", schedule_csv_bytes(m))
        z.writestr("CASEY_FINAL_Model_Audit.json", model_json_bytes(m))
        z.writestr("README_FINAL.txt", "CASEY Final McKinsey-Grade Output Engine. Outputs are generated from the selected project type, Earth/Space mode, scenario, estimate class and schedule level. XER uses PRA-compatible ERMHDR 8.0 structure with dynamic CASEY WBS/tasks.\n")
    bio.seek(0); return bio.getvalue()

def install_v64(app):
    def replace_post(path, endpoint):
        app.router.routes=[r for r in app.router.routes if not (getattr(r,"path",None)==path and "POST" in getattr(r,"methods",set()))]
        app.post(path)(endpoint)
    def workbook(payload: dict = Body(default={})):
        m=_payload_model(payload); return stream(cost_workbook_bytes(m),"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","CASEY_FINAL_Cost_Model.xlsx")
    def risks(payload: dict = Body(default={})):
        m=_payload_model(payload); return stream(risk_workbook_bytes(m),"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","CASEY_FINAL_Risk_Register.xlsx")
    def xer(payload: dict = Body(default={})):
        m=_payload_model(payload); return stream(patch_template_xer(m),"application/octet-stream","CASEY_FINAL_Schedule.xer")
    def csvexp(payload: dict = Body(default={})):
        m=_payload_model(payload); return stream(schedule_csv_bytes(m),"text/csv","CASEY_FINAL_Schedule_Levels.csv")
    def jsonexp(payload: dict = Body(default={})):
        m=_payload_model(payload); return stream(model_json_bytes(m),"application/json","CASEY_FINAL_Model_Audit.json")
    def allpack(payload: dict = Body(default={})):
        m=_payload_model(payload); return stream(all_zip_bytes(m),"application/zip","CASEY_FINAL_Output_Pack.zip")
    def removed(payload: dict = Body(default={})):
        return JSONResponse({"status":"not_in_primary_pack","message":"Primary source-of-truth pack: cost model XLSX, risk register XLSX, XER schedule, schedule CSV and JSON audit."},status_code=410)
    for p,e in [("/export/workbook",workbook),("/export/cost-model",workbook),("/export/risk-register",risks),("/export/xer",xer),("/export/schedule-csv",csvexp),("/export/json",jsonexp),("/export/all",allpack),("/export/full-pack",allpack),("/v64/export/all",allpack)]:
        replace_post(p,e)
    for p in ["/export/pdf","/export/pptx","/export/word"]:
        replace_post(p, removed)

# -----------------------------------------------------------------------------
# FINAL OUTPUT POLISH — decision-led, readable, McKinsey-style Excel outputs.
# Appended after v64: these definitions intentionally override the prior output
# builders without changing the API surface or frontend workflow.
# -----------------------------------------------------------------------------
VERSION = "FINAL"


def _final_formats(wb):
    navy = "#07111F"; blue = "#0B5CAD"; cyan = "#DFFBFF"; pale = "#F6F9FC"; ink = "#111827"; mid = "#475569"; line = "#D8E1EA"
    return {
        "title": wb.add_format({"bold": True, "font_size": 22, "font_color": "#FFFFFF", "bg_color": navy, "align": "left", "valign": "vcenter"}),
        "hero": wb.add_format({"bold": True, "font_size": 28, "font_color": ink, "bg_color": "#FFFFFF", "align": "left", "valign": "vcenter"}),
        "kpi_label": wb.add_format({"font_size": 10, "font_color": mid, "bg_color": cyan, "align": "center", "valign": "vcenter", "border": 1, "border_color": "#9DECF5"}),
        "kpi": wb.add_format({"bold": True, "font_size": 20, "font_color": ink, "bg_color": cyan, "align": "center", "valign": "vcenter", "border": 1, "border_color": "#9DECF5"}),
        "section": wb.add_format({"bold": True, "font_size": 13, "font_color": "#FFFFFF", "bg_color": blue, "border": 1, "border_color": blue, "align": "left", "valign": "vcenter"}),
        "subtle": wb.add_format({"font_size": 10, "font_color": mid, "bg_color": "#FFFFFF", "text_wrap": True, "valign": "top"}),
        "header": wb.add_format({"bold": True, "font_size": 9, "font_color": "#FFFFFF", "bg_color": blue, "border": 1, "border_color": blue, "text_wrap": True, "align": "center", "valign": "vcenter"}),
        "cell": wb.add_format({"font_size": 9, "font_color": ink, "bg_color": "#FFFFFF", "border": 1, "border_color": line, "text_wrap": True, "valign": "top"}),
        "alt": wb.add_format({"font_size": 9, "font_color": ink, "bg_color": pale, "border": 1, "border_color": line, "text_wrap": True, "valign": "top"}),
        "money": wb.add_format({"font_size": 9, "font_color": ink, "bg_color": "#FFFFFF", "border": 1, "border_color": line, "num_format": "$0.0B", "valign": "top"}),
        "num": wb.add_format({"font_size": 9, "font_color": ink, "bg_color": "#FFFFFF", "border": 1, "border_color": line, "num_format": "0.0", "valign": "top"}),
        "pct": wb.add_format({"font_size": 9, "font_color": ink, "bg_color": "#FFFFFF", "border": 1, "border_color": line, "num_format": "0%", "valign": "top"}),
        "decision": wb.add_format({"bold": True, "font_size": 12, "font_color": ink, "bg_color": "#EAF7FF", "border": 1, "border_color": "#9DC3E6", "text_wrap": True, "valign": "top"}),
        "warning": wb.add_format({"font_size": 10, "font_color": ink, "bg_color": "#FFF2CC", "border": 1, "border_color": "#F4B183", "text_wrap": True, "valign": "top"}),
    }


def _write_mck_table(ws, row, col, headers, rows, fmt, widths=None, money_cols=None, num_cols=None, pct_cols=None):
    money_cols = money_cols or set(); num_cols = num_cols or set(); pct_cols = pct_cols or set()
    for j,h in enumerate(headers):
        ws.write(row, col+j, h, fmt["header"])
    for i,r in enumerate(rows, row+1):
        for j,v in enumerate(r):
            f = fmt["cell"] if (i-row) % 2 else fmt["alt"]
            if j in money_cols: f = fmt["money"]
            elif j in pct_cols: f = fmt["pct"]
            elif j in num_cols: f = fmt["num"]
            ws.write(i, col+j, v, f)
    if widths:
        for j,w in enumerate(widths): ws.set_column(col+j, col+j, w)
    ws.autofilter(row, col, row+len(rows), col+len(headers)-1)


def _final_recommendation(model):
    scen = SCENARIO_RULES.get(scenario_key(model), SCENARIO_RULES["base"])
    risk = str(model.get("risk") or "Medium")
    if scenario_key(model) == "faster":
        return "Proceed only where schedule value exceeds acceleration premium. Protect critical path through early procurement, interface controls and decision gates."
    if scenario_key(model) == "cheaper":
        return "Use as a value-engineering challenge case only. Require explicit acceptance of higher residual delivery risk and scope-quality trade-offs."
    if scenario_key(model) == "lower_risk":
        return "Recommended where approval confidence matters most. Accept higher upfront cost in exchange for lower downside exposure and stronger governance."
    if scenario_key(model) == "premium":
        return "Use for flagship resilience and assurance. Highest specification and governance reduce uncertainty but require clear sponsor value case."
    return f"Use as the board reference case. Risk is {risk}; approve the next definition gate, not unconditional full capital commitment."


def _scenario_table(model):
    m=_payload_model(model); p50=as_float_money(m.get("cost_p50")); sched=int(re.findall(r"\d+", str(m.get("schedule") or "60"))[0]) if re.findall(r"\d+", str(m.get("schedule") or "60")) else 60
    selected=scenario_key(m); sel=SCENARIO_RULES[selected]; rows=[]
    for k,r in SCENARIO_RULES.items():
        cost=p50*r["cost"]/sel["cost"]; months=sched*r["duration"]/sel["duration"]
        risk = "High" if r["risk"]>=1.15 else ("Low" if r["risk"]<=0.85 else "Medium")
        conf=max(0.05,min(0.95,(float(m.get("confidence_pct") or 60)+r["confidence"]-sel["confidence"])/100))
        delta_cost=(r["cost"]/sel["cost"]-1); delta_time=(r["duration"]/sel["duration"]-1)
        rows.append([r["label"], cost, months, risk, conf, delta_cost, delta_time, r["decision"], r["tone"], "Selected" if k==selected else ""])
    return rows


def cost_workbook_bytes(model):
    """Final polished cost/control workbook: decision first, then model evidence."""
    if xlsxwriter is None: raise RuntimeError("xlsxwriter not installed")
    m=_payload_model(model); scen=SCENARIO_RULES[scenario_key(m)]; cls=selected_class(m); lvl=selected_level(m); title=clean(m.get("title") or m.get("project_name") or "CASEY Project",80)
    risks=build_risks(m); qcurve=curves(m); cost_sel=normalize_cost_lines(m, scenario_key(m), cls); schedules=all_schedule_levels(m)
    qcra=sorted(risks, key=lambda r:r["cost_emv_bn"], reverse=True); qsra=sorted(risks, key=lambda r:r["schedule_emv_days"], reverse=True)
    bio=BytesIO(); wb=xlsxwriter.Workbook(bio,{"in_memory":True,"strings_to_urls":False})
    fmt=_final_formats(wb)
    # Control Centre
    ws=wb.add_worksheet("00 Control Centre"); ws.hide_gridlines(2); ws.freeze_panes(4,0); ws.set_zoom(90)
    ws.set_column("A:A",22); ws.set_column("B:H",18); ws.set_row(0,34); ws.merge_range("A1:H1", f"CASEY BOARD CONTROL CENTRE — {title}", fmt["title"])
    ws.merge_range("A2:H2", f"Scenario: {scen['label']} | Estimate Class: {cls} | Schedule Level: L{lvl} | Project type: {m.get('mode')} / {m.get('subsector')} | Generated {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}", fmt["subtle"])
    kpis=[("P50 Cost",m.get("cost_p50")),("Cost Range",m.get("cost_range")),("Schedule",m.get("schedule")),("QSRA P80",str((m.get('monte_carlo') or {}).get('qsra',{}).get('p80', 'see QSRA'))),("Risk",m.get("risk")),("Confidence",f"{m.get('confidence_pct')}%"),("Class",CLASS_RULES[cls]["name"]),("Level",f"L{lvl} / {level_activity_count(lvl)} activities")]
    for i,(label,val) in enumerate(kpis):
        r=4+(i//4)*3; c=(i%4)*2
        ws.merge_range(r,c,r,c+1,label,fmt["kpi_label"]); ws.merge_range(r+1,c,r+2,c+1,val,fmt["kpi"])
    ws.merge_range("A12:H12","BOARD RECOMMENDATION",fmt["section"])
    ws.merge_range("A13:H15",_final_recommendation(m),fmt["decision"])
    ws.merge_range("A17:H17","WHAT THE BOARD SHOULD CHALLENGE",fmt["section"])
    challenge=[["1", "What evidence supports the P50 cost and P80 confidence?"],["2", "Which five risks create most contingency and schedule exposure?"],["3", "Which procurement or interface decisions should be made before the next gate?"],["4", "What would change the recommendation between Base, Faster, Cheaper and Lower Risk?"],["5", "What assumptions must be validated before capital commitment?"]]
    _write_mck_table(ws,18,0,["#","Challenge question"],challenge,fmt,[6,120])
    ws.merge_range("A26:H26","BUYER VALUE MESSAGE",fmt["section"])
    ws.merge_range("A27:H29","CASEY produces a first-pass board-grade cost, schedule and risk control pack in minutes. Use it to challenge advisers, compare options and identify where further evidence is required before capital commitment.",fmt["warning"])
    # Scenario comparison
    ws=wb.add_worksheet("01 Scenario Decision"); ws.hide_gridlines(2); ws.merge_range("A1:J1","SCENARIO DECISION MATRIX — TRADE-OFFS MADE EXPLICIT",fmt["title"])
    rows=_scenario_table(m)
    _write_mck_table(ws,4,0,["Scenario","P50 Cost BN","Schedule Months","Risk","Confidence","Cost Δ","Time Δ","Board Decision","What Changes","Selected"],rows,fmt,[14,14,16,12,12,12,12,42,50,12],money_cols={1},num_cols={2},pct_cols={4,5,6})
    chart=wb.add_chart({"type":"column"}); chart.add_series({"name":"P50 Cost BN","categories":"='01 Scenario Decision'!$A$6:$A$12","values":"='01 Scenario Decision'!$B$6:$B$12","fill":{"color":"#0B5CAD"}}); chart.set_title({"name":"Scenario cost"}); chart.set_legend({"none":True}); ws.insert_chart("A15",chart,{"x_scale":1.25,"y_scale":1.0})
    chart2=wb.add_chart({"type":"line"}); chart2.add_series({"name":"Schedule months","categories":"='01 Scenario Decision'!$A$6:$A$12","values":"='01 Scenario Decision'!$C$6:$C$12","line":{"color":"#00A6C8","width":2.25}}); chart2.set_title({"name":"Scenario schedule"}); chart2.set_legend({"none":True}); ws.insert_chart("H15",chart2,{"x_scale":1.15,"y_scale":1.0})
    # Selected estimate
    ws=wb.add_worksheet("02 Selected Estimate"); ws.hide_gridlines(2); ws.merge_range("A1:J1",f"SELECTED CLASS {cls} ESTIMATE — {scen['label']} SCENARIO",fmt["title"])
    rows=[[r["class"],r["cbs"],r["description"],r["type"],r["p10_bn"],r["p50_bn"],r["p80_bn"],r["p90_bn"],r["scenario"],r["basis"]] for r in cost_sel]
    _write_mck_table(ws,4,0,["Class","CBS","Description","Type","P10 BN","P50 BN","P80 BN","P90 BN","Scenario","Basis"],rows,fmt,[8,12,34,14,12,12,12,12,16,64],money_cols={4,5,6,7})
    chart=wb.add_chart({"type":"bar"}); chart.add_series({"name":"P50 BN","categories":f"='02 Selected Estimate'!$C$6:$C${5+len(rows)}","values":f"='02 Selected Estimate'!$F$6:$F${5+len(rows)}","fill":{"color":"#0B5CAD"}}); chart.set_title({"name":"P50 build-up by CBS"}); chart.set_legend({"none":True}); ws.insert_chart("L4",chart,{"x_scale":1.2,"y_scale":1.4})
    # All classes
    ws=wb.add_worksheet("03 All Classes"); ws.hide_gridlines(2); ws.merge_range("A1:J1",f"ALL ESTIMATE CLASSES — SELECTED CLASS {cls} HIGHLIGHTED",fmt["title"])
    rows=[]
    for c in range(1,6):
        for r in all_class_costs(m)[c]: rows.append([c,CLASS_RULES[c]["name"],r["cbs"],r["description"],r["type"],r["p10_bn"],r["p50_bn"],r["p80_bn"],r["p90_bn"],"Selected" if c==cls else ""])
    _write_mck_table(ws,4,0,["Class","Class Name","CBS","Description","Type","P10 BN","P50 BN","P80 BN","P90 BN","Selected"],rows,fmt,[8,26,12,34,14,12,12,12,12,12],money_cols={5,6,7,8})
    # Schedule levels
    ws=wb.add_worksheet("04 Schedule Levels"); ws.hide_gridlines(2); ws.merge_range("A1:I1",f"SCHEDULE LEVELS — SELECTED L{lvl} / {level_activity_count(lvl)} ACTIVITIES",fmt["title"])
    rows=[]
    for lv,acts in schedules.items():
        for a in acts: rows.append([lv,a["activity_id"],a["phase"],a["activity"],a["predecessor"],a["duration_months"],a["duration_days"],a["critical"],"Selected" if lv==lvl else ""])
    _write_mck_table(ws,4,0,["Level","Activity ID","Phase","Activity","Predecessor","Duration Months","Duration Days","Critical","Selected"],rows,fmt,[8,14,22,50,20,16,14,10,12],num_cols={0,5,6})
    # QCRA QSRA
    ws=wb.add_worksheet("05 QCRA QSRA"); ws.hide_gridlines(2); ws.merge_range("A1:C1","QCRA COST CURVE AND QSRA SCHEDULE CURVE",fmt["title"])
    rows=[[r["percentile"],r["qcra_cost_bn"],r["qsra_months"]] for r in qcurve]
    _write_mck_table(ws,4,0,["Percentile","QCRA Cost BN","QSRA Months"],rows,fmt,[12,16,16],money_cols={1},num_cols={0,2})
    chart=wb.add_chart({"type":"line"}); chart.add_series({"name":"QCRA Cost BN","categories":"='05 QCRA QSRA'!$A$6:$A$19","values":"='05 QCRA QSRA'!$B$6:$B$19","line":{"color":"#0B5CAD","width":2.25}}); chart.set_title({"name":"QCRA cost confidence curve"}); ws.insert_chart("E4",chart,{"x_scale":1.3,"y_scale":1.1})
    chart2=wb.add_chart({"type":"line"}); chart2.add_series({"name":"QSRA months","categories":"='05 QCRA QSRA'!$A$6:$A$19","values":"='05 QCRA QSRA'!$C$6:$C$19","line":{"color":"#00A6C8","width":2.25}}); chart2.set_title({"name":"QSRA schedule confidence curve"}); ws.insert_chart("E22",chart2,{"x_scale":1.3,"y_scale":1.1})
    # Tornado
    ws=wb.add_worksheet("06 Tornado Drivers"); ws.hide_gridlines(2); ws.merge_range("A1:L1","SEPARATE COST AND SCHEDULE TORNADO DRIVERS",fmt["title"])
    qrows=[[i+1,r["title"],r["risk_id"],r["cost_emv_bn"],r["owner"],r["mitigation"]] for i,r in enumerate(qcra[:15])]
    srows=[[i+1,r["title"],r["risk_id"],r["schedule_emv_days"],r["owner"],r["mitigation"]] for i,r in enumerate(qsra[:15])]
    _write_mck_table(ws,4,0,["Rank","QCRA Cost Driver","Risk ID","Cost EMV BN","Owner","Mitigation"],qrows,fmt,[8,30,12,14,18,48],money_cols={3})
    _write_mck_table(ws,4,7,["Rank","QSRA Schedule Driver","Risk ID","Days EMV","Owner","Mitigation"],srows,fmt,[8,30,12,14,18,48],num_cols={3})
    # Basis and benchmarks
    ws=wb.add_worksheet("07 Basis + Benchmark"); ws.hide_gridlines(2); ws.merge_range("A1:E1","BASIS OF ESTIMATE AND BENCHMARK VALIDATION",fmt["title"])
    _write_mck_table(ws,4,0,["Topic","Basis"],basis_rows(m),fmt,[26,100])
    _write_mck_table(ws,18,0,["Benchmark","Model Indication","Reference Range","CASEY Challenge","Commentary"],benchmark_rows(m),fmt,[26,30,30,36,64])
    ws.set_landscape(); ws.fit_to_pages(1,0)
    for sheet in wb.worksheets():
        sheet.set_margins(0.25,0.25,0.35,0.35)
    wb.close(); bio.seek(0); return bio.getvalue()


def risk_workbook_bytes(model):
    """Final polished risk register workbook."""
    if xlsxwriter is None: raise RuntimeError("xlsxwriter not installed")
    m=_payload_model(model); scen=SCENARIO_RULES[scenario_key(m)]; risks=build_risks(m); title=clean(m.get("title") or m.get("project_name") or "CASEY Project",80)
    bio=BytesIO(); wb=xlsxwriter.Workbook(bio,{"in_memory":True,"strings_to_urls":False}); fmt=_final_formats(wb)
    ws=wb.add_worksheet("00 Risk Control Centre"); ws.hide_gridlines(2); ws.merge_range("A1:H1",f"RISK CONTROL CENTRE — {title}",fmt["title"])
    ws.merge_range("A2:H2",f"Scenario: {scen['label']} | Risk count: {len(risks)} | Generated {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}",fmt["subtle"])
    total_cost=sum(r["cost_emv_bn"] for r in risks); total_days=sum(r["schedule_emv_days"] for r in risks)
    kpis=[("Open Risks",len(risks)),("Board Visible",sum(1 for r in risks if r["board_visibility"]=="Yes")),("Total Cost EMV",money(total_cost)),("Total Schedule EMV",f"{total_days:.0f} days"),("Top Risk",risks[0]["title"] if risks else "n/a"),("Recommendation",_final_recommendation(m))]
    for i,(label,val) in enumerate(kpis):
        r=4+(i//3)*3; c=(i%3)*2
        ws.merge_range(r,c,r,c+1,label,fmt["kpi_label"]); ws.merge_range(r+1,c,r+2,c+1,val,fmt["kpi"])
    ws.merge_range("A12:H12","TOP BOARD RISKS",fmt["section"])
    top=[[r["risk_id"],r["category"],r["title"],r["probability_pct"]/100,r["cost_emv_bn"],r["schedule_emv_days"],r["owner"],r["mitigation"]] for r in risks[:10]]
    _write_mck_table(ws,13,0,["ID","Category","Risk","Likelihood","Cost EMV BN","Schedule EMV Days","Owner","Mitigation"],top,fmt,[10,16,28,12,14,18,18,56],money_cols={4},num_cols={5},pct_cols={3})
    ws=wb.add_worksheet("01 Full Risk Register"); ws.hide_gridlines(2); ws.freeze_panes(5,0); ws.merge_range("A1:S1",f"FULL RISK REGISTER — {scen['label']} SCENARIO",fmt["title"])
    rows=[[r["risk_id"],r["category"],r["title"],r["cause"],r["event"],r["impact"],r["owner"],r["mitigation"],r["trigger"],r["probability_pct"]/100,r["cost_optimistic_bn"],r["cost_ml_bn"],r["cost_pessimistic_bn"],r["schedule_optimistic_days"],r["schedule_ml_days"],r["schedule_pessimistic_days"],r["cost_emv_bn"],r["schedule_emv_days"],r["board_visibility"]] for r in risks]
    _write_mck_table(ws,4,0,["ID","Category","Risk","Cause","Event","Impact","Owner","Mitigation","Trigger","Likelihood","Cost Low BN","Cost ML BN","Cost High BN","Sched Low Days","Sched ML Days","Sched High Days","Cost EMV BN","Sched EMV Days","Board"],rows,fmt,[10,16,28,40,40,48,18,48,30,12,12,12,12,14,14,14,12,14,10],money_cols={10,11,12,16},num_cols={13,14,15,17},pct_cols={9})
    ws=wb.add_worksheet("02 Risk Tornadoes"); ws.hide_gridlines(2); ws.merge_range("A1:L1","QCRA AND QSRA RISK DRIVERS",fmt["title"])
    qcra=sorted(risks,key=lambda r:r["cost_emv_bn"],reverse=True); qsra=sorted(risks,key=lambda r:r["schedule_emv_days"],reverse=True)
    qrows=[[i+1,r["title"],r["risk_id"],r["cost_emv_bn"],r["owner"]] for i,r in enumerate(qcra[:15])]
    srows=[[i+1,r["title"],r["risk_id"],r["schedule_emv_days"],r["owner"]] for i,r in enumerate(qsra[:15])]
    _write_mck_table(ws,4,0,["Rank","QCRA Cost Driver","ID","Cost EMV BN","Owner"],qrows,fmt,[8,34,12,14,18],money_cols={3})
    _write_mck_table(ws,4,7,["Rank","QSRA Schedule Driver","ID","Days EMV","Owner"],srows,fmt,[8,34,12,14,18],num_cols={3})
    chart=wb.add_chart({"type":"bar"}); chart.add_series({"name":"Cost EMV BN","categories":"='02 Risk Tornadoes'!$B$6:$B$20","values":"='02 Risk Tornadoes'!$D$6:$D$20","fill":{"color":"#0B5CAD"}}); chart.set_title({"name":"QCRA cost tornado"}); chart.set_legend({"none":True}); ws.insert_chart("A23",chart,{"x_scale":1.2,"y_scale":1.2})
    chart2=wb.add_chart({"type":"bar"}); chart2.add_series({"name":"Schedule EMV Days","categories":"='02 Risk Tornadoes'!$I$6:$I$20","values":"='02 Risk Tornadoes'!$K$6:$K$20","fill":{"color":"#00A6C8"}}); chart2.set_title({"name":"QSRA schedule tornado"}); chart2.set_legend({"none":True}); ws.insert_chart("H23",chart2,{"x_scale":1.2,"y_scale":1.2})
    for sheet in wb.worksheets(): sheet.set_margins(0.25,0.25,0.35,0.35)
    wb.close(); bio.seek(0); return bio.getvalue()


def model_json_bytes(model):
    m=_payload_model(model)
    enriched=dict(m); enriched["casey_version"]="FINAL"; enriched["schedule_levels"]=all_schedule_levels(m); enriched["risks"]=build_risks(m); enriched["basis"]=basis_rows(m); enriched["benchmarks"]=benchmark_rows(m); enriched["recommendation"]=_final_recommendation(m)
    return json.dumps(enriched, indent=2, default=str).encode("utf-8")


def all_zip_bytes(model):
    m=_payload_model(model); bio=BytesIO()
    with zipfile.ZipFile(bio,"w",zipfile.ZIP_DEFLATED) as z:
        z.writestr("CASEY_FINAL_Cost_Model.xlsx",cost_workbook_bytes(m))
        z.writestr("CASEY_FINAL_Risk_Register.xlsx",risk_workbook_bytes(m))
        z.writestr("CASEY_FINAL_Schedule.xer",patch_template_xer(m))
        z.writestr("CASEY_FINAL_Schedule_Levels.csv",schedule_csv_bytes(m))
        z.writestr("CASEY_FINAL_Model_Audit.json",model_json_bytes(m))
        z.writestr("README_FINAL.txt","CASEY Final McKinsey-grade output pack. Cost model, risk register, PRA schedule, schedule CSV and audit JSON are all generated from the selected project, scenario, estimate class and schedule level.\n")
    bio.seek(0); return bio.getvalue()


def install_v64(app):
    def replace_post(path, endpoint):
        app.router.routes=[r for r in app.router.routes if not (getattr(r,"path",None)==path and "POST" in getattr(r,"methods",set()))]
        app.post(path)(endpoint)
    def workbook(payload: dict = Body(default={})):
        m=_payload_model(payload); return stream(cost_workbook_bytes(m),"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","CASEY_FINAL_Cost_Model.xlsx")
    def risks(payload: dict = Body(default={})):
        m=_payload_model(payload); return stream(risk_workbook_bytes(m),"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","CASEY_FINAL_Risk_Register.xlsx")
    def xer(payload: dict = Body(default={})):
        m=_payload_model(payload); return stream(patch_template_xer(m),"application/octet-stream","CASEY_FINAL_Schedule.xer")
    def csvexp(payload: dict = Body(default={})):
        m=_payload_model(payload); return stream(schedule_csv_bytes(m),"text/csv","CASEY_FINAL_Schedule_Levels.csv")
    def jsonexp(payload: dict = Body(default={})):
        m=_payload_model(payload); return stream(model_json_bytes(m),"application/json","CASEY_FINAL_Model_Audit.json")
    def allpack(payload: dict = Body(default={})):
        m=_payload_model(payload); return stream(all_zip_bytes(m),"application/zip","CASEY_FINAL_Output_Pack.zip")
    for p,e in [("/export/workbook",workbook),("/export/cost-model",workbook),("/export/risk-register",risks),("/export/xer",xer),("/export/schedule-csv",csvexp),("/export/json",jsonexp),("/export/all",allpack),("/export/full-pack",allpack),("/v64/export/all",allpack)]: replace_post(p,e)

# =============================================================================
# CASEY V118 FINAL DEMO EXPORT REFINEMENT PATCH
# Focus: planner-grade XER sophistication, asymmetric tornado realism, and deeper
# risk ownership controls while keeping the existing API surface unchanged.
# =============================================================================
VERSION = "V118_FINAL_DEMO_EXPORT_REFINEMENT"

_REL_TYPES = ["PR_FS", "PR_SS", "PR_FF", "PR_FS", "PR_SS", "PR_FS"]
_CALENDAR_BY_PHASE = {
    "Governance": "5D_BUSINESS", "Business Case": "5D_BUSINESS", "Scope": "5D_BUSINESS", "Design": "5D_BUSINESS",
    "Consents": "5D_BUSINESS", "Procurement": "5D_BUSINESS", "Enabling": "6D_DELIVERY", "Civil": "6D_DELIVERY",
    "Build": "6D_DELIVERY", "Systems": "6D_DELIVERY", "Integration": "7D_COMMISSIONING", "Commissioning": "7D_COMMISSIONING",
    "Handover": "5D_BUSINESS", "Mission": "5D_MISSION", "Mission Definition": "5D_MISSION", "Architecture": "5D_MISSION",
    "Technology": "5D_MISSION", "Supply Chain": "5D_MISSION", "Manufacturing": "6D_AIT", "AIT": "6D_AIT",
    "Launch Prep": "7D_LAUNCH_WINDOW", "Launch + Transit": "7D_LAUNCH_WINDOW", "Transit": "7D_LAUNCH_WINDOW",
    "Deployment": "7D_SURFACE_OPS", "Operations": "7D_SURFACE_OPS", "Sustainment": "7D_SURFACE_OPS",
}
_RESOURCE_BY_PHASE = {
    "Governance": "Executive sponsor / PMO", "Business Case": "Commercial + controls team", "Scope": "Design authority",
    "Design": "Design management team", "Consents": "Permitting / stakeholder lead", "Procurement": "Procurement + vendor managers",
    "Enabling": "Enabling works contractor", "Civil": "Civil contractor", "Build": "Main contractor",
    "Systems": "MEP / OEM specialists", "Integration": "Systems integration team", "Commissioning": "Commissioning authority + operations",
    "Handover": "Operations readiness team", "Mission": "Mission assurance board", "Mission Definition": "Mission assurance board",
    "Architecture": "Systems engineering authority", "Technology": "TRL closure team", "Supply Chain": "Payload / launch procurement",
    "Manufacturing": "AIT manufacturing cell", "AIT": "AIT cell + quality", "Launch Prep": "Launch provider + range authority",
    "Launch + Transit": "Flight dynamics + mission ops", "Transit": "Flight dynamics + mission ops", "Deployment": "Surface/orbital ops team",
    "Operations": "Mission operations", "Sustainment": "Sustainment + resupply team",
}
_STATUS_CYCLE = ["Active mitigation", "Monitor", "Escalated", "Owner assigned", "Evidence required", "Response in progress"]


def _v118_skew(i, scenario):
    scen = SCENARIO_RULES.get(scenario, SCENARIO_RULES["base"])
    # intentionally uneven, not symmetrical. Acceleration worsens right-tail and float risks.
    base = [1.00, 1.55, 0.72, 1.28, 0.91, 1.83, 0.66, 1.17, 2.10, 0.58, 1.42, 0.84][i % 12]
    if scenario == "faster" and i % 3 in (0, 1): base *= 1.22
    if scenario == "lower_risk" and i % 4 != 0: base *= 0.76
    return max(0.35, base * scen.get("risk", 1.0))


def _v118_owner_for(title, idx, pkey):
    t = str(title).lower()
    if "grid" in t or "energ" in t or "utility" in t: return "EVP Delivery / Grid Interface Director"
    if "transformer" in t or "switchgear" in t or "procurement" in t or "supply" in t: return "Chief Procurement Officer"
    if "commission" in t or "testing" in t or "integration" in t or "ist" in t: return "Commissioning Director"
    if "design" in t or "interface" in t: return "Design Authority Lead"
    if pkey == "space" and ("launch" in t or "range" in t): return "Mission Integration Director"
    if pkey == "space" and ("trl" in t or "qualification" in t or "flight" in t): return "Chief Systems Engineer"
    return ["Project Director", "Controls Director", "Commercial Lead", "Risk Manager", "Construction Director", "Operations Readiness Lead"][idx % 6]


def build_schedule_rows(model, level=None, scenario=None):
    """V118 planner-grade schedule rows: WBS depth, relationship types, lags/leads, calendars,
    resource loading and near-critical density are visible in CSV/JSON and injected into XER.
    """
    level = int(level or selected_level(model))
    scenario = scenario or scenario_key(model)
    rule = SCENARIO_RULES.get(scenario, SCENARIO_RULES["base"])
    total_months = max(6, int(re.findall(r"\d+", str(model.get("schedule") or "60"))[0]) if re.findall(r"\d+", str(model.get("schedule") or "60")) else 60)
    phases = schedule_phases(model)
    n = level_activity_count(level)
    rows=[]
    prior_by_phase={}
    prev=""
    for i in range(n):
        phase, base = phases[i % len(phases)]
        pkg = 1 + i // len(phases)
        workface = 1 + (i % 4)
        detail_suffix = ""
        if level >= 3: detail_suffix = f" — package {pkg}"
        if level >= 4: detail_suffix += " / control account"
        if level >= 5: detail_suffix += f" / workface {workface}"
        act_id = f"L{level}-{(i+1):03d}"
        wave = [0.82,1.06,0.93,1.24,0.74,1.15,0.88][i % 7]
        dur = max(0.5, round((total_months * rule["duration"] / max(1,n)) * wave, 1))
        if i==0: name = "Project start / mobilisation"
        elif i==n-1: name = "Project complete / operational handover"
        else: name = base + detail_suffix
        rel = _REL_TYPES[i % len(_REL_TYPES)] if i else ""
        lag_days = 0
        if rel == "PR_SS": lag_days = [0, 3, 5, -2, 7][i % 5]
        elif rel == "PR_FF": lag_days = [0, 2, -3, 6][i % 4]
        elif phase in ["Procurement", "Systems", "Integration", "Commissioning", "Launch Prep", "Deployment"]: lag_days = [0, 5, 10, -5][i % 4]
        pred = prev
        # Add occasional phase convergence predecessor for more authentic network density.
        if level >= 4 and phase in prior_by_phase and i % 5 == 0:
            pred = f"{prev};{prior_by_phase[phase]}"
        calendar = _CALENDAR_BY_PHASE.get(phase, "5D_BUSINESS")
        resource = _RESOURCE_BY_PHASE.get(phase, "Integrated delivery team")
        near = phase in ["Procurement","Systems","Integration","Commissioning","Launch Prep","Deployment"] or i >= n-8 or (level >= 4 and i % 6 == 0)
        critical = "Yes" if i >= n-5 or (scenario == "faster" and near and i % 3 != 2) else "No"
        float_days = 0 if critical == "Yes" else max(2, [4,7,11,16,23,35,48][i % 7] - (6 if near else 0))
        loading = round((1.0 + (i % 5) * 0.35) * (1.25 if near else 0.85), 2)
        wbs_path = f"{clean(model.get('title') or 'CASEY Project',30)} / {phase} / Package {pkg}"
        rows.append({"level":level,"activity_id":act_id,"phase":phase,"activity":name,"predecessor":pred,"relationship_type":rel,"lag_days":lag_days,"duration_months":dur,"duration_days":int(round(dur*21.75)),"critical":critical,"near_critical":"Yes" if near else "No","total_float_days":float_days,"calendar":calendar,"primary_resource":resource,"resource_loading_fte":loading,"wbs_path":wbs_path,"scenario":SCENARIO_RULES[scenario]["label"],"basis":f"Level {level} schedule with {rel or 'start'} logic, {lag_days:+d}d lag, {calendar} calendar and {resource}; scenario {SCENARIO_RULES[scenario]['label']} factor {rule['duration']:.2f}."})
        prior_by_phase[phase]=act_id
        prev=act_id
    return rows


def schedule_csv_bytes(model):
    import io
    buf=io.StringIO(); w=csv.writer(buf)
    w.writerow(["Level","Activity ID","WBS Path","Phase","Activity","Predecessor","Relationship Type","Lag / Lead Days","Duration Months","Duration Days","Critical","Near Critical","Total Float Days","Calendar","Primary Resource","Resource Loading FTE","Scenario","Basis"])
    for lvl,acts in all_schedule_levels(model).items():
        for a in acts:
            w.writerow([lvl,a["activity_id"],a.get("wbs_path",""),a["phase"],a["activity"],a.get("predecessor",""),a.get("relationship_type",""),a.get("lag_days",0),a["duration_months"],a["duration_days"],a["critical"],a.get("near_critical",""),a.get("total_float_days",""),a.get("calendar",""),a.get("primary_resource",""),a.get("resource_loading_fte",""),a["scenario"],a["basis"]])
    return buf.getvalue().encode("utf-8")


def build_risks(model):
    """V118 operational risk register: owner, response status, mitigation confidence,
    residual exposure, escalation trigger and asymmetric QCRA/QSRA tornado fields.
    """
    pkey=project_type_key(model); scenario=scenario_key(model); scen=SCENARIO_RULES.get(scenario, SCENARIO_RULES["base"])
    base_titles = list(SECTOR_RISKS.get(pkey, SECTOR_RISKS["general"])) + list(SECTOR_RISKS["general"])
    extra = ["Grid energisation readiness", "Transformer fabrication slot", "Switchgear vendor assumption", "Integrated systems testing overlap", "Commissioning resource availability", "Owner decision latency", "Authority approval condition", "Interface data maturity", "Logistics route constraint", "Recovery float consumption", "Operational readiness evidence", "Supplier concentration exposure", "Contractor productivity variance", "Critical package quality escape", "Change-control backlog", "Residual contingency masking", "Near-critical path convergence", "Market escalation tail", "Cyber/security approval", "Handover documentation maturity", "Training and mobilisation readiness", "Utility outage calendar dependency", "Factory acceptance test failure", "Late design freeze", "Commercial claim crystallisation"]
    if pkey == "space":
        extra += ["Launch window slip", "TRL closure shortfall", "Mass margin growth", "Thermal vacuum qualification failure", "Range approval delay", "Orbital insertion contingency", "Surface deployment sequencing", "Resupply cadence constraint"]
    titles=[]
    for t in base_titles + extra:
        if t not in titles: titles.append(t)
    out=[]
    p50=max(0.1, as_float_money(model.get("cost_p50")))
    for i,t in enumerate(titles[:42]):
        skew=_v118_skew(i, scenario)
        prob=max(6,min(88,int(round((16 + (i*9)%58) * (0.92 + 0.08*skew)))))
        ml=round(p50 * (0.006 + (i%9)*0.0028) * skew, 3)
        low=round(ml * [0.18,0.31,0.44,0.22,0.55][i%5], 3)
        high=round(ml * [1.8,2.4,3.1,1.35,4.2,2.05][i%6], 3)
        sched_ml=int((8 + (i%11)*6) * (0.85 + skew*0.25))
        sched_low=max(1,int(sched_ml * [0.12,0.28,0.4,0.18][i%4]))
        sched_high=max(sched_ml+1,int(sched_ml * [1.55,2.2,2.9,1.35,3.7][i%5]))
        owner=_v118_owner_for(t, i, pkey)
        status=_STATUS_CYCLE[i % len(_STATUS_CYCLE)]
        mit_conf=max(0.18,min(0.86,0.64 - (skew-1)*0.11 - (0.10 if status in ["Evidence required","Escalated"] else 0) + (0.04 if status=="Active mitigation" else 0)))
        residual=max(0.08,min(0.92,(prob/100)*(1-mit_conf)*skew))
        exposure="Very High" if residual>0.55 else "High" if residual>0.34 else "Medium" if residual>0.18 else "Low"
        trigger=f"{t} trigger threshold breached: >{7+(i%5)*7} days movement, supplier evidence gap, or unresolved gate condition"
        mitigation=f"{status}: named owner to validate evidence, close causal assumption, update QSRA/QCRA impacts and confirm residual exposure before next governance gate."
        out.append({
            "risk_id": f"R-{i+1:03d}", "category": "Space" if pkey=="space" else ("Delivery" if i%3 else "Commercial / Governance"),
            "title": t, "cause": f"{t} driven by maturity, interface, market or approval uncertainty under the {scen['label']} scenario.",
            "event": f"{t} materialises and changes the active control baseline.",
            "impact": "Cost growth, schedule movement, contingency drawdown, rework or delayed benefit realisation.",
            "owner": owner, "mitigation": mitigation, "response_status": status, "mitigation_confidence": round(mit_conf,2),
            "residual_exposure": exposure, "residual_exposure_score": round(residual,3), "target_review_date": (datetime.utcnow()+timedelta(days=14+(i%6)*7)).strftime("%Y-%m-%d"),
            "escalation_threshold": trigger, "trigger": trigger, "probability_pct": prob,
            "cost_optimistic_bn": low, "cost_ml_bn": ml, "cost_pessimistic_bn": high,
            "cost_downside_delta_bn": round(high-ml,3), "cost_upside_delta_bn": round(ml-low,3),
            "schedule_optimistic_days": sched_low, "schedule_ml_days": sched_ml, "schedule_pessimistic_days": sched_high,
            "schedule_downside_delta_days": sched_high-sched_ml, "schedule_upside_delta_days": sched_ml-sched_low,
            "cost_emv_bn": round(ml*prob/100,3), "schedule_emv_days": round(sched_ml*prob/100,1),
            "status": "Open", "board_visibility": "Yes" if i<14 or exposure in ["High","Very High"] else "No", "scenario": scen["label"]
        })
    out.sort(key=lambda r:(r["residual_exposure_score"], r["cost_emv_bn"], r["schedule_emv_days"]), reverse=True)
    return out


def curves(model):
    """V118 deliberately irregular P-curves with asymmetric tails."""
    p50=as_float_money(model.get("cost_p50")); sched=int(re.findall(r"\d+",str(model.get("schedule") or "60"))[0]) if re.findall(r"\d+",str(model.get("schedule") or "60")) else 60
    cls=selected_class(model); rule=CLASS_RULES[cls]; scen=SCENARIO_RULES[scenario_key(model)]
    rows=[]
    for pct in [1,5,10,20,30,40,50,60,70,80,85,90,95,99]:
        x=(pct-50)/50
        tail=abs(x)**1.55
        wobble=[-0.012,0.006,-0.004,0.014,-0.008,0.0,0.0,0.011,-0.006,0.018,0.027,0.042,0.065,0.095][len(rows)]
        cost=p50*(1 + (rule["high"]-1)*max(0,x)*tail*1.35 + (1-rule["low"])*min(0,x)*tail*0.8 + wobble)
        months=sched*(1 + (0.34*scen["risk"])*max(0,x)*tail*1.5 + (0.13)*min(0,x)*tail + wobble*0.7)
        rows.append({"percentile":pct,"qcra_cost_bn":round(max(0.01,cost),3),"qsra_months":round(max(1,months),2)})
    return rows


def patch_template_xer(model):
    tables, order = _template_tables()
    scen_key=scenario_key(model); scen=SCENARIO_RULES[scen_key]
    lvl=selected_level(model); ptype=project_type_key(model)
    mode="Space" if ptype=="space" else "Earth"
    title=_safe_xer_text(model.get("title") or model.get("project_name") or ("Lunar Base" if mode=="Space" else "Capital Project"),80)
    proj_id=3338; proj_short=_safe_xer_text(f"CASEY_{mode}_{scen['label']}_L{lvl}",32).replace(" ","_")[:32]
    proj_name=_safe_xer_text(f"CASEY {title} — {mode} — {scen['label']} Scenario — Level {lvl} Schedule",120)
    start_dt=datetime(2026,1,5,8,0); acts=build_schedule_rows(model,lvl,scen_key)
    phases=[]
    for a in acts:
        if a["phase"] not in phases: phases.append(a["phase"])
    root_wbs=35000; wbs_ids={phase:35001+i for i,phase in enumerate(phases)}
    out=["ERMHDR\t8.0\t2026-01-05\tProject\tadmin\tPPSEn\tdbxDatabaseNoName\tProject Management\tUSD"]
    def emit_table(name, rows=None):
        if name not in tables: return
        out.append(tables[name]["header"]); out.append("%F\t"+"\t".join(tables[name]["fields"]))
        if rows is None:
            for r in tables[name]["rows"]: out.append("%R\t"+"\t".join(r))
        else: out.extend(rows)
    for name in ["CURRTYPE","OBS","UDFTYPE"]: emit_table(name)
    pf=tables["PROJECT"]["fields"]; base_project=tables["PROJECT"]["rows"][0]
    emit_table("PROJECT", [_row(pf,{"proj_id":proj_id,"proj_short_name":proj_short,"clndr_id":3892,"wbs_max_sum_level":lvl,"plan_start_date":start_dt.strftime("%Y-%m-%d %H:%M"),"last_recalc_date":start_dt.strftime("%Y-%m-%d %H:%M"),"scd_end_date":(start_dt+timedelta(days=sum(a["duration_days"] for a in acts))).strftime("%Y-%m-%d %H:%M"),"guid":f"CASEYV118{mode}{scen['label']}L{lvl}"[:22],"add_by_name":"CASEY","export_flag":"Y"},base_project)])
    emit_table("CALENDAR"); emit_table("SCHEDOPTIONS")
    wf=tables["PROJWBS"]["fields"]; base_wbs=tables["PROJWBS"]["rows"][0]
    wbs_rows=[_row(wf,{"wbs_id":root_wbs,"proj_id":proj_id,"obs_id":565,"seq_num":0,"proj_node_flag":"Y","wbs_short_name":proj_short,"wbs_name":proj_name,"parent_wbs_id":"","guid":f"CASEY_V118_ROOT"},base_wbs)]
    for i,phase in enumerate(phases):
        wbs_rows.append(_row(wf,{"wbs_id":wbs_ids[phase],"proj_id":proj_id,"obs_id":565,"seq_num":i+1,"proj_node_flag":"N","wbs_short_name":f"{i+1:02d}","wbs_name":_safe_xer_text(phase,60),"parent_wbs_id":root_wbs,"guid":f"CASEY_V118_WBS_{i+1:02d}"},base_wbs))
    emit_table("PROJWBS",wbs_rows)
    tf=tables["TASK"]["fields"]; base_task=tables["TASK"]["rows"][0]
    task_rows=[]; current=start_dt
    for i,a in enumerate(acts):
        tid=200000+i+1; dur_days=max(1,int(a["duration_days"])); dur_hr=dur_days*8; s=current; e=current+timedelta(days=dur_days); current=e+timedelta(days=max(0,1+a.get("lag_days",0)//7))
        task_rows.append(_row(tf,{"task_id":tid,"proj_id":proj_id,"wbs_id":wbs_ids[a["phase"]],"clndr_id":3892,"phys_complete_pct":"0","est_wt":str(a.get("resource_loading_fte",1)),"complete_pct_type":"CP_Drtn","task_type":"TT_Task","duration_type":"DT_FixedRate","status_code":"TK_NotStart","task_code":a["activity_id"],"task_name":_safe_xer_text(a["activity"],120),"total_float_hr_cnt":str(int(a.get("total_float_days",0))*8),"free_float_hr_cnt":str(max(0,int(a.get("total_float_days",0))*4)),"remain_drtn_hr_cnt":dur_hr,"target_drtn_hr_cnt":dur_hr,"target_start_date":s.strftime("%Y-%m-%d %H:%M"),"target_end_date":e.strftime("%Y-%m-%d %H:%M"),"early_start_date":s.strftime("%Y-%m-%d %H:%M"),"early_end_date":e.strftime("%Y-%m-%d %H:%M"),"late_start_date":(s+timedelta(days=int(a.get("total_float_days",0)))).strftime("%Y-%m-%d %H:%M"),"late_end_date":(e+timedelta(days=int(a.get("total_float_days",0)))).strftime("%Y-%m-%d %H:%M"),"restart_date":s.strftime("%Y-%m-%d %H:%M"),"reend_date":e.strftime("%Y-%m-%d %H:%M"),"driving_path_flag":"Y" if a["critical"]=="Yes" else "N","guid":f"CASEY_V118_TASK_{i+1:03d}"[:22],"create_user":"CASEY","update_user":"CASEY","create_date":start_dt.strftime("%Y-%m-%d %H:%M"),"update_date":start_dt.strftime("%Y-%m-%d %H:%M")},base_task))
    emit_table("TASK",task_rows)
    if "TASKPRED" in tables:
        pfld=tables["TASKPRED"]["fields"]; base_pred=tables["TASKPRED"]["rows"][0] if tables["TASKPRED"]["rows"] else None
        pred_rows=[]
        for i,a in enumerate(acts[1:],1):
            pred_rows.append(_row(pfld,{"task_pred_id":260000+i,"task_id":200000+i+1,"pred_task_id":200000+i,"proj_id":proj_id,"pred_proj_id":proj_id,"pred_type":a.get("relationship_type") or "PR_FS","lag_hr_cnt":int(a.get("lag_days",0))*8,"float_path":"CASEY","aref":"","arls":""},base_pred))
            if lvl >= 4 and i > 4 and i % 7 == 0:
                pred_rows.append(_row(pfld,{"task_pred_id":270000+i,"task_id":200000+i+1,"pred_task_id":200000+max(1,i-3),"proj_id":proj_id,"pred_proj_id":proj_id,"pred_type":"PR_SS","lag_hr_cnt":16,"float_path":"CASEY_NC","aref":"","arls":""},base_pred))
        emit_table("TASKPRED",pred_rows)
    if "UDFVALUE" in tables:
        uf=tables["UDFVALUE"]["fields"]; udf_rows=[]
        for i,a in enumerate(acts):
            tid=200000+i+1; d=max(1,int(a["duration_days"])); vals=[(329,d),(330,max(d+1,int(d*(1.25+(i%5)*0.11)))),(331,max(1,int(d*(0.62+(i%4)*0.05))))]
            for udf_id,val in vals: udf_rows.append(_row(uf,{"udf_type_id":udf_id,"fk_id":tid,"proj_id":proj_id,"udf_text":val},None))
        emit_table("UDFVALUE",udf_rows)
    text="\n".join(out)+"\n"
    for bad in ["Mondrian","Lilly Mondrian","L01-10","Shire"]: text=text.replace(bad, proj_short if bad=="L01-10" else "CASEY")
    return text.encode("latin1",errors="ignore")


def model_json_bytes(model):
    m=_payload_model(model)
    enriched=dict(m); enriched["casey_version"]="V118 Final Demo Export Refinement"; enriched["schedule_levels"]=all_schedule_levels(m); enriched["risks"]=build_risks(m); enriched["p_curves"]=curves(m); enriched["basis"]=basis_rows(m); enriched["benchmarks"]=benchmark_rows(m); enriched["recommendation"]=_final_recommendation(m); enriched["export_refinements"]=["XER relationship types/lags/calendars/resource loading","near-critical density","asymmetric tornado drivers","risk owner/status/mitigation confidence/residual exposure"]
    return json.dumps(enriched,indent=2,default=str).encode("utf-8")

def risk_workbook_bytes(model):
    """V118 risk register with operational ownership, mitigation confidence and residual exposure."""
    if xlsxwriter is None: raise RuntimeError("xlsxwriter not installed")
    m=_payload_model(model); scen=SCENARIO_RULES[scenario_key(m)]; risks=build_risks(m); title=clean(m.get("title") or m.get("project_name") or "CASEY Project",80)
    bio=BytesIO(); wb=xlsxwriter.Workbook(bio,{"in_memory":True,"strings_to_urls":False}); fmt=_final_formats(wb)
    ws=wb.add_worksheet("00 Risk Control Centre"); ws.hide_gridlines(2); ws.merge_range("A1:J1",f"V118 RISK CONTROL CENTRE — {title}",fmt["title"])
    ws.merge_range("A2:J2",f"Scenario: {scen['label']} | Risk count: {len(risks)} | Generated {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')} | Includes owner, response status, mitigation confidence and residual exposure",fmt["subtle"])
    total_cost=sum(r["cost_emv_bn"] for r in risks); total_days=sum(r["schedule_emv_days"] for r in risks); high=sum(1 for r in risks if r.get("residual_exposure") in ["High","Very High"])
    kpis=[("Open Risks",len(risks)),("Board Visible",sum(1 for r in risks if r["board_visibility"]=="Yes")),("High Residual",high),("Total Cost EMV",money(total_cost)),("Total Schedule EMV",f"{total_days:.0f} days"),("Top Risk",risks[0]["title"] if risks else "n/a")]
    for i,(label,val) in enumerate(kpis):
        r=4+(i//3)*3; c=(i%3)*3
        ws.merge_range(r,c,r,c+2,label,fmt["kpi_label"]); ws.merge_range(r+1,c,r+2,c+2,val,fmt["kpi"])
    ws.merge_range("A12:J12","TOP BOARD RISKS — OWNER / ACTION / RESIDUAL EXPOSURE",fmt["section"])
    top=[[r["risk_id"],r["category"],r["title"],r["probability_pct"]/100,r["cost_emv_bn"],r["schedule_emv_days"],r["owner"],r.get("response_status"),r.get("mitigation_confidence"),r.get("residual_exposure")] for r in risks[:12]]
    _write_mck_table(ws,13,0,["ID","Category","Risk","Likelihood","Cost EMV BN","Schedule EMV Days","Owner","Response Status","Mitigation Confidence","Residual Exposure"],top,fmt,[10,18,34,12,14,18,26,22,18,18],money_cols={4},num_cols={5,8},pct_cols={3})
    ws=wb.add_worksheet("01 Full Risk Register"); ws.hide_gridlines(2); ws.freeze_panes(5,0); ws.merge_range("A1:Z1",f"V118 FULL RISK REGISTER — {scen['label']} SCENARIO",fmt["title"])
    rows=[]
    for r in risks:
        rows.append([r["risk_id"],r["category"],r["title"],r["cause"],r["event"],r["impact"],r["owner"],r.get("response_status"),r["mitigation"],r.get("mitigation_confidence"),r.get("residual_exposure"),r.get("residual_exposure_score"),r.get("target_review_date"),r.get("escalation_threshold"),r["trigger"],r["probability_pct"]/100,r["cost_optimistic_bn"],r["cost_ml_bn"],r["cost_pessimistic_bn"],r.get("cost_upside_delta_bn"),r.get("cost_downside_delta_bn"),r["schedule_optimistic_days"],r["schedule_ml_days"],r["schedule_pessimistic_days"],r.get("schedule_downside_delta_days"),r["board_visibility"]])
    _write_mck_table(ws,4,0,["ID","Category","Risk","Cause","Event","Impact","Owner","Response Status","Mitigation","Mitigation Confidence","Residual Exposure","Residual Score","Target Review","Escalation Threshold","Trigger","Likelihood","Cost Low BN","Cost ML BN","Cost High BN","Cost Upside BN","Cost Downside BN","Sched Low Days","Sched ML Days","Sched High Days","Schedule Downside Days","Board"],rows,fmt,[10,18,30,42,42,48,26,22,54,18,18,14,14,44,38,12,12,12,12,14,16,14,14,14,18,10],money_cols={16,17,18,19,20},num_cols={9,11,21,22,23,24},pct_cols={15})
    ws=wb.add_worksheet("02 Asymmetric Tornadoes"); ws.hide_gridlines(2); ws.merge_range("A1:N1","V118 ASYMMETRIC QCRA / QSRA TORNADO DRIVERS",fmt["title"])
    qcra=sorted(risks,key=lambda r:r.get("cost_downside_delta_bn",0),reverse=True); qsra=sorted(risks,key=lambda r:r.get("schedule_downside_delta_days",0),reverse=True)
    qrows=[[i+1,r["title"],r["risk_id"],r.get("cost_upside_delta_bn"),r.get("cost_downside_delta_bn"),r["cost_emv_bn"],r["owner"],r.get("residual_exposure")] for i,r in enumerate(qcra[:16])]
    srows=[[i+1,r["title"],r["risk_id"],r.get("schedule_upside_delta_days"),r.get("schedule_downside_delta_days"),r["schedule_emv_days"],r["owner"],r.get("residual_exposure")] for i,r in enumerate(qsra[:16])]
    _write_mck_table(ws,4,0,["Rank","QCRA Cost Driver","ID","Upside BN","Downside BN","EMV BN","Owner","Residual"],qrows,fmt,[8,34,10,12,14,12,24,14],money_cols={3,4,5})
    _write_mck_table(ws,4,9,["Rank","QSRA Driver","ID","Upside Days","Downside Days","EMV Days","Owner","Residual"],srows,fmt,[8,34,10,12,14,12,24,14],num_cols={3,4,5})
    chart=wb.add_chart({"type":"bar"}); chart.add_series({"name":"Cost downside BN","categories":"='02 Asymmetric Tornadoes'!$B$6:$B$21","values":"='02 Asymmetric Tornadoes'!$E$6:$E$21","fill":{"color":"#0B5CAD"}}); chart.set_title({"name":"QCRA asymmetric downside tornado"}); chart.set_legend({"none":True}); ws.insert_chart("A24",chart,{"x_scale":1.25,"y_scale":1.15})
    chart2=wb.add_chart({"type":"bar"}); chart2.add_series({"name":"Schedule downside days","categories":"='02 Asymmetric Tornadoes'!$K$6:$K$21","values":"='02 Asymmetric Tornadoes'!$N$6:$N$21","fill":{"color":"#00A6C8"}}); chart2.set_title({"name":"QSRA asymmetric downside tornado"}); chart2.set_legend({"none":True}); ws.insert_chart("J24",chart2,{"x_scale":1.25,"y_scale":1.15})
    ws=wb.add_worksheet("03 Response Plan"); ws.hide_gridlines(2); ws.merge_range("A1:I1","MITIGATION RESPONSE PLAN",fmt["title"])
    resp=[[r["risk_id"],r["title"],r["owner"],r.get("response_status"),r.get("mitigation_confidence"),r.get("residual_exposure"),r.get("target_review_date"),r.get("escalation_threshold"),r["mitigation"]] for r in risks[:25]]
    _write_mck_table(ws,4,0,["ID","Risk","Owner","Status","Mitigation Confidence","Residual","Review Date","Escalation Threshold","Action"],resp,fmt,[10,34,26,22,18,14,14,42,56],num_cols={4})
    for sheet in wb.worksheets(): sheet.set_margins(0.25,0.25,0.35,0.35)
    wb.close(); bio.seek(0); return bio.getvalue()


def all_zip_bytes(model):
    m=_payload_model(model); bio=BytesIO()
    with zipfile.ZipFile(bio,"w",zipfile.ZIP_DEFLATED) as z:
        z.writestr("00_READ_ME_V118_FINAL_DEMO.txt", "CASEY V118 final demo output pack. Refinements: planner-grade XER with relationship types, lags/leads, calendars, resource loading and near-critical density; asymmetric tornado risk logic; risk owners, mitigation confidence, response status and residual exposure. Demo-watermarked first-pass intelligence, not certified commercial reliance.\n")
        z.writestr("01_CASEY_V118_Cost_Model.xlsx",cost_workbook_bytes(m))
        z.writestr("02_CASEY_V118_Risk_Register_Operational.xlsx",risk_workbook_bytes(m))
        z.writestr("03_CASEY_V118_Planner_Grade_Schedule.xer",patch_template_xer(m))
        z.writestr("04_CASEY_V118_Schedule_Levels_Resource_Loaded.csv",schedule_csv_bytes(m))
        z.writestr("05_CASEY_V118_Model_Audit.json",model_json_bytes(m))
    bio.seek(0); return bio.getvalue()
# =============================================================================
# END CASEY V118 FINAL DEMO EXPORT REFINEMENT PATCH
# =============================================================================

# V118.1 compatibility override: handles both tuple phase rows and UNIVERSAL_WBS package lists.
def build_schedule_rows(model, level=None, scenario=None):
    level = int(level or selected_level(model))
    scenario = scenario or scenario_key(model)
    rule = SCENARIO_RULES.get(scenario, SCENARIO_RULES["base"])
    total_months = max(6, int(re.findall(r"\d+", str(model.get("schedule") or "60"))[0]) if re.findall(r"\d+", str(model.get("schedule") or "60")) else 60)
    raw_phases = schedule_phases(model)
    flat=[]
    for item in raw_phases:
        if isinstance(item, (list, tuple)) and len(item)>=2:
            phase=item[0]; payload=item[1]
            if isinstance(payload, list):
                for act in payload: flat.append((phase, str(act)))
            else:
                flat.append((phase, str(payload)))
    if not flat: flat=[("Governance","Project controls setup"),("Delivery","Main works"),("Commissioning","Testing and handover")]
    n = level_activity_count(level)
    rows=[]; prior_by_phase={}; prev=""
    for i in range(n):
        phase, base = flat[i % len(flat)]
        pkg = 1 + i // len(flat); workface = 1 + (i % 4)
        detail_suffix = ""
        if level >= 3: detail_suffix = f" — package {pkg}"
        if level >= 4: detail_suffix += " / control account"
        if level >= 5: detail_suffix += f" / workface {workface}"
        act_id = f"L{level}-{(i+1):03d}"
        wave = [0.82,1.06,0.93,1.24,0.74,1.15,0.88][i % 7]
        dur = max(0.5, round((total_months * rule["duration"] / max(1,n)) * wave, 1))
        if i==0: name = "Project start / mobilisation"
        elif i==n-1: name = "Project complete / operational handover"
        else: name = str(base) + detail_suffix
        rel = _REL_TYPES[i % len(_REL_TYPES)] if i else ""
        lag_days = 0
        if rel == "PR_SS": lag_days = [0, 3, 5, -2, 7][i % 5]
        elif rel == "PR_FF": lag_days = [0, 2, -3, 6][i % 4]
        elif phase in ["Procurement", "Systems", "Integration", "Commissioning", "Launch Prep", "Deployment", "AIT", "Launch + Transit"]: lag_days = [0, 5, 10, -5][i % 4]
        pred = prev
        if level >= 4 and phase in prior_by_phase and i % 5 == 0: pred = f"{prev};{prior_by_phase[phase]}"
        calendar = _CALENDAR_BY_PHASE.get(phase, "5D_BUSINESS")
        resource = _RESOURCE_BY_PHASE.get(phase, "Integrated delivery team")
        near = phase in ["Procurement","Systems","Integration","Commissioning","Launch Prep","Deployment","AIT","Launch + Transit"] or i >= n-8 or (level >= 4 and i % 6 == 0)
        critical = "Yes" if i >= n-5 or (scenario == "faster" and near and i % 3 != 2) else "No"
        float_days = 0 if critical == "Yes" else max(2, [4,7,11,16,23,35,48][i % 7] - (6 if near else 0))
        loading = round((1.0 + (i % 5) * 0.35) * (1.25 if near else 0.85), 2)
        wbs_path = f"{clean(model.get('title') or 'CASEY Project',30)} / {phase} / Package {pkg}"
        rows.append({"level":level,"activity_id":act_id,"phase":phase,"activity":name,"predecessor":pred,"relationship_type":rel,"lag_days":lag_days,"duration_months":dur,"duration_days":int(round(dur*21.75)),"critical":critical,"near_critical":"Yes" if near else "No","total_float_days":float_days,"calendar":calendar,"primary_resource":resource,"resource_loading_fte":loading,"wbs_path":wbs_path,"scenario":SCENARIO_RULES[scenario]["label"],"basis":f"Level {level} schedule with {rel or 'start'} logic, {lag_days:+d}d lag, {calendar} calendar and {resource}; scenario {SCENARIO_RULES[scenario]['label']} factor {rule['duration']:.2f}."})
        prior_by_phase[phase]=act_id; prev=act_id
    return rows

# =============================================================================
# CASEY V119 FINAL DEMO LOCK — EARTH + SPACE REFINEMENT PATCH
# Adds final requested refinements:
# - uglier asymmetric tornadoes
# - deeper risk ownership/status/residual exposure
# - richer XER/Schedule near-critical narrative and relationship diversity
# - separate QCRA/QSRA workbook in the final pack
# =============================================================================
VERSION = "V119_FINAL_EARTH_SPACE_DEMO_LOCK"

_V119_REL_TYPES = ["PR_FS", "PR_SS", "PR_FF", "PR_FS", "PR_SS", "PR_FS", "PR_FF"]
_V119_STATUS = ["Active mitigation", "Evidence required", "Escalated", "Owner assigned", "Response in progress", "Monitor", "Governance decision required"]


def _v119_space_terms(title: str, field: str) -> str:
    t = str(title).lower()
    if "launch" in t:
        return {
            "owner": "Mission Integration Director",
            "mitigation": "Maintain secondary launch-window option, confirm range/readiness evidence, and preserve manifest contingency before mission gate.",
            "trigger": "Launch provider, range or payload readiness moves by more than one manifest review cycle."
        }[field]
    if "thermal" in t or "radiation" in t:
        return {
            "owner": "Chief Systems Engineer",
            "mitigation": "Close thermal-vacuum/radiation evidence, verify design margin, and update mass/power reserve before qualification freeze.",
            "trigger": "Thermal, power or shielding margin falls below mission-assurance threshold."
        }[field]
    if "mass" in t or "payload" in t:
        return {
            "owner": "Payload Integration Lead",
            "mitigation": "Reconcile mass-growth log, confirm payload interface control documents, and protect launch vehicle allocation.",
            "trigger": "Payload mass or interface change exceeds allocated reserve."
        }[field]
    if "surface" in t or "deployment" in t or "resupply" in t:
        return {
            "owner": "Surface Operations Director",
            "mitigation": "Validate deployment sequence, autonomous recovery modes and resupply cadence against mission survival criteria.",
            "trigger": "Surface operations readiness score falls below board-comfort threshold."
        }[field]
    return {
        "owner": "Mission Assurance Board",
        "mitigation": "Assign named mission owner, close evidence gap, update QCRA/QSRA and confirm residual exposure before next gate.",
        "trigger": "Mission assurance evidence or autonomous recovery basis remains unresolved at gate review."
    }[field]


def _v119_owner(title, idx, pkey):
    if pkey == "space":
        return _v119_space_terms(title, "owner")
    return _v118_owner_for(title, idx, pkey)


def build_risks(model):
    """V119 operational risk model with intentionally irregular/asymmetric tornado inputs."""
    pkey = project_type_key(model); scenario = scenario_key(model); scen = SCENARIO_RULES.get(scenario, SCENARIO_RULES["base"])
    base_titles = list(SECTOR_RISKS.get(pkey, SECTOR_RISKS["general"])) + list(SECTOR_RISKS["general"])
    if pkey == "space":
        extra = [
            "Launch window slip", "Launch manifest dependency", "TRL closure shortfall", "Mass margin growth", "Thermal-power balance", "Thermal vacuum qualification failure",
            "Radiation-hardening qualification", "Payload interface maturity", "Autonomous recovery evidence", "Surface deployment sequencing", "Resupply cadence constraint",
            "Ground segment dependency", "Orbital insertion contingency", "Mission operations staffing", "Range approval delay", "Dust contamination and surface logistics"
        ]
    else:
        extra = [
            "Grid energisation readiness", "Transformer fabrication slot", "Switchgear vendor assumption", "Integrated systems testing overlap", "Commissioning resource availability",
            "Owner decision latency", "Authority approval condition", "Interface data maturity", "Logistics route constraint", "Recovery float consumption",
            "Operational readiness evidence", "Supplier concentration exposure", "Contractor productivity variance", "Critical package quality escape", "Change-control backlog",
            "Residual contingency masking", "Near-critical path convergence", "Market escalation tail", "Cyber/security approval", "Handover documentation maturity",
            "Training and mobilisation readiness", "Utility outage calendar dependency", "Factory acceptance test failure", "Late design freeze", "Commercial claim crystallisation"
        ]
    titles=[]
    for t in base_titles + extra:
        if t and t not in titles:
            titles.append(t)
    p50=max(0.1, as_float_money(model.get("cost_p50")))
    out=[]
    for i,t in enumerate(titles[:46]):
        skew = _v118_skew(i, scenario) * [0.83, 1.41, 0.61, 1.96, 0.74, 1.18, 2.37, 0.52, 1.66, 0.91, 1.12, 2.05, 0.69][i % 13]
        if pkey == "space" and i % 4 in (1,2): skew *= 1.18
        prob = max(5, min(89, int(round((13 + (i*11)%63) * (0.88 + 0.07*skew)))))
        # ugly/non-symmetric uncertainty shape: high side is not a simple multiple of ML
        ml = round(p50 * (0.0048 + (i%10)*0.0026) * skew, 3)
        low = round(max(0.001, ml * [0.11, 0.34, 0.21, 0.58, 0.08, 0.47, 0.29][i%7]), 3)
        high = round(max(ml + 0.002, ml * [1.42, 2.85, 1.76, 4.65, 2.18, 3.35, 1.29, 5.10][i%8]), 3)
        sched_ml = int((6 + (i%13)*5.7) * (0.72 + skew*0.31))
        sched_low = max(1, int(sched_ml * [0.09,0.36,0.18,0.49,0.25][i%5]))
        sched_high = max(sched_ml+1, int(sched_ml * [1.28,2.65,1.83,3.55,1.49,4.25][i%6]))
        owner = _v119_owner(t, i, pkey)
        status = _V119_STATUS[(i + (2 if scenario == "faster" else 0)) % len(_V119_STATUS)]
        mit_conf = max(0.12, min(0.88, 0.66 - (skew-1)*0.075 - (0.13 if status in ["Evidence required","Escalated","Governance decision required"] else 0) + (0.06 if status == "Active mitigation" else 0)))
        residual = max(0.07, min(0.96, (prob/100) * (1-mit_conf) * min(2.5, skew)))
        exposure = "Very High" if residual > .58 else "High" if residual > .35 else "Medium" if residual > .18 else "Low"
        if pkey == "space":
            mitigation = _v119_space_terms(t, "mitigation")
            trigger = _v119_space_terms(t, "trigger")
            cause = f"{t} is driven by mission-assurance evidence maturity, launch cadence, qualification burden or autonomous recovery uncertainty under the {scen['label']} scenario."
            impact = "Mission readiness movement, launch/re-manifest cost, qualification rework, mass/power reserve drawdown or loss of autonomous recovery confidence."
        else:
            trigger = f"{t} threshold breached: >{7+(i%6)*6} days movement, unresolved supplier evidence, or gate condition not closed."
            mitigation = f"{status}: named owner to close evidence, test mitigation credibility, update QCRA/QSRA impacts and confirm residual exposure before governance release."
            cause = f"{t} is driven by maturity, supplier, interface, market or approval uncertainty under the {scen['label']} scenario."
            impact = "Cost growth, schedule movement, contingency drawdown, rework, operational-readiness delay or confidence deterioration."
        out.append({
            "risk_id": f"R-{i+1:03d}", "category": "Space" if pkey=="space" else ("Commercial / Governance" if i%5==0 else "Delivery"),
            "title": t, "cause": cause, "event": f"{t} materialises and changes the active control baseline.", "impact": impact,
            "owner": owner, "mitigation": mitigation, "response_status": status, "mitigation_confidence": round(mit_conf,2),
            "residual_exposure": exposure, "residual_exposure_score": round(residual,3), "target_review_date": (datetime.utcnow()+timedelta(days=10+(i%7)*8)).strftime("%Y-%m-%d"),
            "escalation_threshold": trigger, "trigger": trigger, "probability_pct": prob,
            "cost_optimistic_bn": low, "cost_ml_bn": ml, "cost_pessimistic_bn": high,
            "cost_downside_delta_bn": round(high-ml,3), "cost_upside_delta_bn": round(ml-low,3),
            "schedule_optimistic_days": sched_low, "schedule_ml_days": sched_ml, "schedule_pessimistic_days": sched_high,
            "schedule_downside_delta_days": sched_high-sched_ml, "schedule_upside_delta_days": sched_ml-sched_low,
            "cost_emv_bn": round(ml*prob/100,3), "schedule_emv_days": round(sched_ml*prob/100,1),
            "status": "Open", "board_visibility": "Yes" if i<14 or exposure in ["High","Very High"] else "No", "scenario": scen["label"]
        })
    out.sort(key=lambda r:(r["residual_exposure_score"]*1.7 + r["cost_downside_delta_bn"]/(p50+0.01) + r["schedule_downside_delta_days"]/500), reverse=True)
    return out


def build_schedule_rows(model, level=None, scenario=None):
    """V119 schedule rows with richer near-critical network narrative and messy relationship logic."""
    level = int(level or selected_level(model)); scenario = scenario or scenario_key(model)
    rule = SCENARIO_RULES.get(scenario, SCENARIO_RULES["base"])
    total_months = max(6, int(re.findall(r"\d+", str(model.get("schedule") or "60"))[0]) if re.findall(r"\d+", str(model.get("schedule") or "60")) else 60)
    raw_phases = schedule_phases(model)
    flat=[]
    for item in raw_phases:
        if isinstance(item, (list, tuple)) and len(item)>=2:
            phase=item[0]; payload=item[1]
            if isinstance(payload, list):
                for act in payload: flat.append((phase, str(act)))
            else: flat.append((phase, str(payload)))
    if not flat: flat=[("Governance","Project controls setup"),("Delivery","Main works"),("Commissioning","Testing and handover")]
    n=level_activity_count(level); rows=[]; prior_by_phase={}; prev=""
    for i in range(n):
        phase, base = flat[i % len(flat)]; pkg=1+i//len(flat); workface=1+(i%4)
        suffix = (f" — package {pkg}" if level>=3 else "") + (" / control account" if level>=4 else "") + (f" / workface {workface}" if level>=5 else "")
        act_id=f"L{level}-{(i+1):03d}"
        wave=[0.78,1.13,0.91,1.36,0.67,1.21,0.84,1.52,0.73][i%9]
        dur=max(0.5, round((total_months*rule["duration"]/max(1,n))*wave,1))
        name = "Project start / mobilisation" if i==0 else ("Project complete / operational handover" if i==n-1 else str(base)+suffix)
        rel=_V119_REL_TYPES[i % len(_V119_REL_TYPES)] if i else ""
        lag_days=0
        if rel == "PR_SS": lag_days=[0,4,-3,9,2,-6][i%6]
        elif rel == "PR_FF": lag_days=[0,3,-4,7,12,-2][i%6]
        elif phase in ["Procurement","Systems","Integration","Commissioning","Launch Prep","Deployment","AIT","Launch + Transit","Transit","Operations"]: lag_days=[0,6,11,-5,17][i%5]
        pred=prev
        # more realistic convergence: occasional dual predecessor from same phase or previous package
        if level>=4 and phase in prior_by_phase and i%5==0: pred=f"{prev};{prior_by_phase[phase]}"
        if level>=5 and i>8 and i%11==0: pred=f"{pred};L{level}-{max(1,i-7):03d}"
        calendar=_CALENDAR_BY_PHASE.get(phase,"5D_BUSINESS")
        resource=_RESOURCE_BY_PHASE.get(phase,"Integrated delivery team")
        near = phase in ["Procurement","Systems","Integration","Commissioning","Launch Prep","Deployment","AIT","Launch + Transit","Transit","Operations"] or i>=n-10 or (level>=4 and i%5 in (0,1))
        critical = "Yes" if i>=n-5 or (scenario=="faster" and near and i%4!=2) else "No"
        float_days = 0 if critical=="Yes" else max(1, [2,5,9,14,21,32,46,7,18][i%9] - (7 if near else 0))
        loading=round((0.9+(i%6)*0.31)*(1.35 if near else 0.82),2)
        wbs_path=f"{clean(model.get('title') or 'CASEY Project',30)} / {phase} / Package {pkg}"
        nc_note = "Near-critical: convergence risk, float fragmentation and gate dependency." if near and critical != "Yes" else ("Critical path driver." if critical == "Yes" else "Non-critical support logic.")
        rows.append({"level":level,"activity_id":act_id,"phase":phase,"activity":name,"predecessor":pred,"relationship_type":rel,"lag_days":lag_days,"duration_months":dur,"duration_days":int(round(dur*21.75)),"critical":critical,"near_critical":"Yes" if near else "No","total_float_days":float_days,"calendar":calendar,"primary_resource":resource,"resource_loading_fte":loading,"wbs_path":wbs_path,"near_critical_narrative":nc_note,"scenario":SCENARIO_RULES[scenario]["label"],"basis":f"Level {level} schedule with {rel or 'start'} logic, {lag_days:+d}d lag/lead, {calendar} calendar and {resource}. {nc_note} Scenario {SCENARIO_RULES[scenario]['label']} factor {rule['duration']:.2f}."})
        prior_by_phase[phase]=act_id; prev=act_id
    return rows


def schedule_csv_bytes(model):
    import io
    buf=io.StringIO(); w=csv.writer(buf)
    w.writerow(["Level","Activity ID","WBS Path","Phase","Activity","Predecessor","Relationship Type","Lag / Lead Days","Duration Months","Duration Days","Critical","Near Critical","Total Float Days","Calendar","Primary Resource","Resource Loading FTE","Near-Critical Narrative","Scenario","Basis"])
    for lvl,acts in all_schedule_levels(model).items():
        for a in acts:
            w.writerow([lvl,a["activity_id"],a.get("wbs_path",""),a["phase"],a["activity"],a.get("predecessor",""),a.get("relationship_type",""),a.get("lag_days",0),a["duration_months"],a["duration_days"],a["critical"],a.get("near_critical",""),a.get("total_float_days",""),a.get("calendar",""),a.get("primary_resource",""),a.get("resource_loading_fte",""),a.get("near_critical_narrative",""),a["scenario"],a["basis"]])
    return buf.getvalue().encode("utf-8")


def qcra_qsra_workbook_bytes(model):
    """Standalone V119 QCRA/QSRA export with ugly/asymmetric tornadoes and percentile curves."""
    if xlsxwriter is None: raise RuntimeError("xlsxwriter not installed")
    m=_payload_model(model); risks=build_risks(m); curve=curves(m); title=clean(m.get("title") or "CASEY Project",80); scen=SCENARIO_RULES[scenario_key(m)]
    bio=BytesIO(); wb=xlsxwriter.Workbook(bio,{"in_memory":True,"strings_to_urls":False}); fmt=_final_formats(wb)
    ws=wb.add_worksheet("00 QCRA QSRA Control"); ws.hide_gridlines(2); ws.merge_range("A1:H1",f"V119 QCRA / QSRA CONTROL — {title}",fmt["title"])
    ws.merge_range("A2:H2",f"Scenario: {scen['label']} | Asymmetric tornadoes | Separate cost and schedule risk | Generated {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}",fmt["subtle"])
    q80=next((x["qcra_cost_bn"] for x in curve if x["percentile"]==80), None); q90=next((x["qcra_cost_bn"] for x in curve if x["percentile"]==90), None)
    s80=next((x["qsra_months"] for x in curve if x["percentile"]==80), None); s90=next((x["qsra_months"] for x in curve if x["percentile"]==90), None)
    kpis=[("QCRA P50",m.get("cost_p50")),("QCRA P80",money(q80 or 0)),("QCRA P90",money(q90 or 0)),("QSRA P50",m.get("schedule")),("QSRA P80",f"{s80} months"),("QSRA P90",f"{s90} months")]
    for i,(label,val) in enumerate(kpis):
        r=4+(i//3)*3; c=(i%3)*2; ws.merge_range(r,c,r,c+1,label,fmt["kpi_label"]); ws.merge_range(r+1,c,r+2,c+1,val,fmt["kpi"])
    ws.merge_range("A12:H12","BOARD INTERPRETATION",fmt["section"])
    ws.merge_range("A13:H15","P50 is the reference control case. P80/P90 represent board downside exposure. Tornado bars are intentionally asymmetric: drivers are ranked by downside tail, not clean average movement.",fmt["warning"])
    ws=wb.add_worksheet("01 P-Curves"); ws.hide_gridlines(2)
    rows=[[x["percentile"],x["qcra_cost_bn"],x["qsra_months"]] for x in curve]
    _write_mck_table(ws,0,0,["Percentile","QCRA Cost BN","QSRA Months"],rows,fmt,[12,16,16],money_cols={1},num_cols={2})
    ch=wb.add_chart({"type":"area"}); ch.add_series({"name":"QCRA Cost BN","categories":"='01 P-Curves'!$A$2:$A$15","values":"='01 P-Curves'!$B$2:$B$15","fill":{"color":"#9DECF5","transparency":35},"line":{"color":"#0B5CAD","width":2.0}}); ch.set_title({"name":"QCRA cost confidence curve"}); ws.insert_chart("E2",ch,{"x_scale":1.25,"y_scale":1.15})
    ch2=wb.add_chart({"type":"line"}); ch2.add_series({"name":"QSRA Months","categories":"='01 P-Curves'!$A$2:$A$15","values":"='01 P-Curves'!$C$2:$C$15","line":{"color":"#7030A0","width":2.5},"marker":{"type":"circle","size":5}}); ch2.set_title({"name":"QSRA finish-date curve"}); ws.insert_chart("E20",ch2,{"x_scale":1.25,"y_scale":1.15})
    ws=wb.add_worksheet("02 Ugly Tornado Drivers"); ws.hide_gridlines(2); ws.merge_range("A1:P1","V119 ASYMMETRIC TORNADOES — DOWNSIDE TAIL RANKING",fmt["title"])
    q=sorted(risks,key=lambda r:r.get("cost_downside_delta_bn",0),reverse=True)[:18]; s=sorted(risks,key=lambda r:r.get("schedule_downside_delta_days",0),reverse=True)[:18]
    _write_mck_table(ws,4,0,["Rank","QCRA Driver","ID","Upside BN","Downside BN","EMV BN","Owner","Residual"],[[i+1,r["title"],r["risk_id"],r["cost_upside_delta_bn"],r["cost_downside_delta_bn"],r["cost_emv_bn"],r["owner"],r["residual_exposure"]] for i,r in enumerate(q)],fmt,[8,34,10,12,14,12,26,14],money_cols={3,4,5})
    _write_mck_table(ws,4,9,["Rank","QSRA Driver","ID","Upside Days","Downside Days","EMV Days","Owner","Residual"],[[i+1,r["title"],r["risk_id"],r["schedule_upside_delta_days"],r["schedule_downside_delta_days"],r["schedule_emv_days"],r["owner"],r["residual_exposure"]] for i,r in enumerate(s)],fmt,[8,34,10,12,14,12,26,14],num_cols={3,4,5})
    ch=wb.add_chart({"type":"bar"}); ch.add_series({"name":"Cost downside BN","categories":"='02 Ugly Tornado Drivers'!$B$6:$B$23","values":"='02 Ugly Tornado Drivers'!$E$6:$E$23","fill":{"color":"#0B5CAD"}}); ch.set_title({"name":"QCRA ugly/asymmetric downside tornado"}); ch.set_legend({"none":True}); ws.insert_chart("A27",ch,{"x_scale":1.35,"y_scale":1.25})
    ch2=wb.add_chart({"type":"bar"}); ch2.add_series({"name":"Schedule downside days","categories":"='02 Ugly Tornado Drivers'!$K$6:$K$23","values":"='02 Ugly Tornado Drivers'!$N$6:$N$23","fill":{"color":"#00A6C8"}}); ch2.set_title({"name":"QSRA ugly/asymmetric downside tornado"}); ch2.set_legend({"none":True}); ws.insert_chart("J27",ch2,{"x_scale":1.35,"y_scale":1.25})
    ws=wb.add_worksheet("03 Board Drivers"); ws.hide_gridlines(2)
    rows=[[r["risk_id"],r["title"],r["owner"],r["response_status"],r["mitigation_confidence"],r["residual_exposure"],r["mitigation"],r["escalation_threshold"]] for r in risks[:24]]
    _write_mck_table(ws,0,0,["ID","Driver","Owner","Status","Mitigation Confidence","Residual","Mitigation","Escalation Threshold"],rows,fmt,[10,32,26,22,18,14,56,48],num_cols={4})
    for sheet in wb.worksheets(): sheet.set_margins(0.25,0.25,0.35,0.35)
    wb.close(); bio.seek(0); return bio.getvalue()


def model_json_bytes(model):
    m=_payload_model(model)
    enriched=dict(m); enriched["casey_version"]="V119 Final Earth + Space Demo Lock"; enriched["schedule_levels"]=all_schedule_levels(m); enriched["risks"]=build_risks(m); enriched["p_curves"]=curves(m); enriched["basis"]=basis_rows(m); enriched["benchmarks"]=benchmark_rows(m); enriched["recommendation"]=_final_recommendation(m); enriched["export_refinements"]=["V119 planner-grade XER with relationship types, lags/leads, calendars, resource loading and near-critical density","uglier asymmetric tornado drivers ranked by downside tail","risk owners, response status, mitigation confidence and residual exposure","space-specific mission assurance owners, triggers and mitigation language"]
    return json.dumps(enriched,indent=2,default=str).encode("utf-8")


def all_zip_bytes(model):
    m=_payload_model(model); bio=BytesIO()
    with zipfile.ZipFile(bio,"w",zipfile.ZIP_DEFLATED) as z:
        z.writestr("00_READ_ME_V119_FINAL_EARTH_SPACE_DEMO.txt", "CASEY V119 final Earth + Space demo output pack. Final refinements: uglier asymmetric tornadoes, reduced output density, irregular scenario deltas in model evidence, chart spacing uplift, risk owners/status/mitigation confidence/residual exposure, planner-grade XER plus near-critical schedule narrative, and space-specific mission assurance logic. Demo-watermarked first-pass intelligence, not certified commercial reliance.\n")
        z.writestr("01_CASEY_V119_Cost_Model.xlsx",cost_workbook_bytes(m))
        z.writestr("02_CASEY_V119_Risk_Register_Operational.xlsx",risk_workbook_bytes(m))
        z.writestr("03_CASEY_V119_Planner_Grade_Schedule.xer",patch_template_xer(m))
        z.writestr("04_CASEY_V119_Schedule_Levels_Resource_Loaded.csv",schedule_csv_bytes(m))
        z.writestr("05_CASEY_V119_QCRA_QSRA_Asymmetric_Tornado.xlsx",qcra_qsra_workbook_bytes(m))
        z.writestr("06_CASEY_V119_Model_Audit.json",model_json_bytes(m))
    bio.seek(0); return bio.getvalue()
# =============================================================================
# END CASEY V119 FINAL DEMO LOCK PATCH
# =============================================================================
