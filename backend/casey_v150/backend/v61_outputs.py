from __future__ import annotations
from io import BytesIO
from datetime import datetime, timedelta
import json, zipfile, csv, re, html, math, os
from fastapi.responses import StreamingResponse, JSONResponse

try:
    import xlsxwriter
except Exception:
    xlsxwriter = None

VERSION = "v62"

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
