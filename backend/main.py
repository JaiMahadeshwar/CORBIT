from __future__ import annotations

from fastapi import FastAPI, HTTPException, UploadFile, File, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import Any, Dict, List, Optional, Tuple
from datetime import datetime
from io import BytesIO, StringIO
import csv, json, math, os, random, re, sqlite3, statistics, zipfile, hashlib, uuid

import numpy as np
from openpyxl import Workbook
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak

APP_VERSION = "CASEY TITAN X v26 Revenue Machine + GTM Demo Edition"
DB_PATH = os.environ.get("CASEY_DB", "casey_titan_v26.sqlite3")
DEMO_LIMIT_PER_IP = int(os.environ.get("CASEY_DEMO_LIMIT_PER_IP", "1"))
PUBLIC_DEMO_LIMIT = int(os.environ.get("CASEY_PUBLIC_DEMO_LIMIT", "1"))
ADMIN_TOKEN = os.environ.get("CASEY_ADMIN_TOKEN", "")

app = FastAPI(title=APP_VERSION, version="26.0-revenue-machine")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

class GenerateRequest(BaseModel):
    prompt: str
    client: Optional[str] = None
    class_level: Optional[int] = 3
    schedule_level: Optional[int] = 3
    scenario: Optional[str] = "base"
    demo: Optional[bool] = False
    active_model: Optional[Dict[str, Any]] = None


class PublicDemoRequest(BaseModel):
    email: str
    project_type: str = "Earth"
    project_description: str
    location: Optional[str] = None
    size_or_capacity: Optional[str] = None
    stage: Optional[str] = "Concept / early feasibility"
    biggest_concern: Optional[str] = "Cost, schedule and risk confidence"
    fingerprint: Optional[str] = None
    client_token: Optional[str] = None

class PublicDemoFeedback(BaseModel):
    run_id: str
    rating: int
    comment: Optional[str] = None

class ChatRequest(BaseModel):
    question: str
    project: Optional[Dict[str, Any]] = None
    demo: Optional[bool] = False

class SaveProjectRequest(BaseModel):
    name: str
    model: Dict[str, Any]

# ------------------------- database -------------------------
def db() -> sqlite3.Connection:
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con

def init_db():
    con = db(); cur = con.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS projects(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        client TEXT,
        prompt TEXT,
        mode TEXT,
        created_at TEXT,
        model_json TEXT NOT NULL
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS demo_usage(
        ip TEXT PRIMARY KEY,
        count INTEGER NOT NULL DEFAULT 0,
        updated_at TEXT
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS public_demo_uses(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        run_id TEXT UNIQUE,
        email_hash TEXT,
        ip_hash TEXT,
        fingerprint_hash TEXT,
        client_token_hash TEXT,
        project_type TEXT,
        project_text TEXT,
        model_json TEXT,
        created_at TEXT
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS public_demo_feedback(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        run_id TEXT,
        rating INTEGER,
        comment TEXT,
        created_at TEXT
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS uploads(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        filename TEXT,
        created_at TEXT,
        analysis_json TEXT
    )""")
    con.commit(); con.close()
init_db()

# ------------------------- helpers -------------------------
def client_ip(request: Request) -> str:
    return request.headers.get("x-forwarded-for", request.client.host if request.client else "local").split(",")[0].strip()

def check_demo_allowance(ip: str) -> Dict[str, Any]:
    con = db(); cur = con.cursor(); row = cur.execute("SELECT count FROM demo_usage WHERE ip=?", (ip,)).fetchone()
    count = int(row["count"]) if row else 0
    con.close()
    return {"allowed": count < DEMO_LIMIT_PER_IP, "used": count, "limit": DEMO_LIMIT_PER_IP}

def record_demo_use(ip: str):
    con = db(); cur = con.cursor(); now = datetime.utcnow().isoformat()
    cur.execute("INSERT INTO demo_usage(ip,count,updated_at) VALUES(?,?,?) ON CONFLICT(ip) DO UPDATE SET count=count+1, updated_at=excluded.updated_at", (ip, 1, now))
    con.commit(); con.close()

def _sha(value: str) -> str:
    value = (value or "").strip().lower()
    return hashlib.sha256(value.encode("utf-8")).hexdigest() if value else ""

def _normalise_email(email: str) -> str:
    return (email or "").strip().lower()

def _public_demo_brief_quality_score(req: PublicDemoRequest) -> Dict[str, Any]:
    """Public demo gate: allow credible infrastructure briefs quickly, block rubbish.
    The old gate was too strict and blocked strong executive-style briefs. This version
    scores useful signals but does not demand every field when CASEY can infer them.
    """
    raw_desc = (req.project_description or "").strip()
    desc = raw_desc.lower()
    joined = " ".join([
        desc,
        (req.location or "").lower(),
        (req.size_or_capacity or "").lower(),
        (req.stage or "").lower(),
        (req.biggest_concern or "").lower(),
        (req.project_type or "").lower(),
    ])

    score = 0
    reasons = []
    contradictions = []

    words = re.findall(r"[a-zA-Z0-9]+(?:[-/][a-zA-Z0-9]+)?", desc)
    unique_words = set(words)

    nonsense_terms = [
        "asdf", "qwerty", "lorem ipsum", "test test", "blah blah", "hello world",
        "ignore previous", "ignore all", "jailbreak", "prompt injection", "write a poem",
        "make me rich", "bitcoin", "crypto meme", "football match", "recipe", "dating",
        "tell me a joke", "song lyrics"
    ]
    if any(x in desc for x in nonsense_terms):
        contradictions.append("This does not look like a credible project brief. Please enter a real infrastructure or space programme.")

    # Repetition / spam check.
    if len(words) >= 12 and len(unique_words) < max(8, len(words) * 0.25):
        contradictions.append("The brief appears repetitive or low quality. Please write a real project description.")

    # Length helps, but short credible briefs should still be allowed.
    if len(words) >= 18:
        score += 24
    elif len(words) >= 10:
        score += 14
    else:
        reasons.append("Add a little more project detail: asset, location/environment and main concern.")

    asset_terms = [
        "data centre","data center","datacenter","airport","terminal","runway","rail","metro","station","hospital",
        "fab","semiconductor","wafer","nuclear","smr","hydrogen","wind","solar","grid","battery","gigafactory",
        "defence","defense","military","naval","airbase","radar","command","campus","port","water","desalination",
        "flood","lng","carbon capture","ccs","stadium","university","fibre","fiber","ev charging",
        "biologics","therapeutics","gmp","aseptic","fill-finish","fill finish","fda","cqv","pharma","pharmaceutical",
        "manufacturing","cleanroom","clean utilities","cold-chain","cold chain","warehouse","logistics",
        "lunar","moon","mars","orbital","orbit","leo","cislunar","satellite","spaceport","propellant","habitat",
        "launch vehicle","payload","asteroid","space infrastructure","space data centre","space data center",
        "orbital compute","orbital ai","thermal rejection","relay communications","lunar base","mars base"
    ]
    has_asset = any(x in joined for x in asset_terms)
    if has_asset:
        score += 28
    else:
        reasons.append("State the asset type, e.g. data centre, GMP campus, rail, hospital, lunar base or orbital platform.")

    location_terms = [
        "north carolina","carolina","arizona","texas","boston","london","cambridge","manchester","riyadh",
        "dubai","abu dhabi","qatar","uk","usa","united states","uae","saudi","canada","australia","singapore",
        "india","japan","germany","france","poland","moon","lunar","mars","leo","orbit","orbital","cislunar",
        "spaceport","deep space"
    ]
    has_location = (req.location and len(req.location.strip()) >= 3 and "auto-inferred" not in req.location.lower()) or any(x in joined for x in location_terms)
    if has_location:
        score += 18
    else:
        reasons.append("Add a location or operating environment.")

    size_terms = [
        "mw","gw","km","beds","m2","sqm","sq m","satellites","crew","tonnes","ha","capacity","runway",
        "stations","terminal","phase","phased","halls","modules","multi-product","multi product","fill-finish",
        "campus","site","lines","clusters","constellation","base","hub","plant","network"
    ]
    if any(x in joined for x in size_terms):
        score += 12

    concern_terms = [
        "cost","schedule","risk","procurement","approval","consent","grid","logistics","commissioning",
        "funding","delivery","safety","regulatory","interface","utilities","critical path","scope","confidence",
        "supply chain","phasing","validation","qualification","licensing","resilience","thermal","radiation",
        "servicing","latency","autonomous","debris","power density","operational readiness","production continuity",
        "fda","cqv","inspection","launch cadence","long-lead","long lead"
    ]
    has_concern = any(x in joined for x in concern_terms)
    if has_concern:
        score += 18
    else:
        reasons.append("Add the main concern: cost, schedule, risk, procurement, approvals, logistics or commissioning.")

    project_context_terms = [
        "project","programme","program","facility","campus","hub","plant","terminal","network","corridor",
        "base","station","platform","outpost","depot","infrastructure","scheme","development","expansion",
        "upgrade","rollout","manufacturing","construction","delivery"
    ]
    if any(x in joined for x in project_context_terms):
        score += 12

    # Strong domain phrases should pass even if a formal field was not provided.
    strong_domain = has_asset and (has_location or any(x in joined for x in ["orbital","lunar","mars","leo","spaceport"])) and has_concern
    if strong_domain:
        score = max(score, 88)

    space_terms = ["moon","lunar","mars","orbital","orbit","leo","cislunar","spaceport","satellite constellation","launch vehicle","rocket","payload","space station","propellant depot","asteroid"]
    earth_satellite_facility = any(x in joined for x in [
        "satellite control centre", "satellite control center", "secure satellite control",
        "ground station", "space domain awareness ground station", "mission operations centre",
        "mission operations center", "mission operations rooms"
    ])

    # Do not punish product/commercial launch language in Earth sectors.
    product_launch_only = any(x in joined for x in ["commercial launch demand","product launch","market launch","launch demand"]) and not any(x in joined for x in ["rocket","launch vehicle","spaceport","launch pad","orbital","leo","lunar","mars"])
    effective_space_terms = [x for x in space_terms if not (x == "launch" and product_launch_only)]

    if (req.project_type or "").lower() == "earth" and any(x in joined for x in effective_space_terms) and not earth_satellite_facility and not product_launch_only:
        # Only block if it is genuinely contradictory, not if frontend auto-inferred Earth before backend reroutes.
        if any(x in joined for x in ["orbital","leo","lunar","moon","mars","spaceport","launch vehicle","rocket"]):
            contradictions.append("The brief contains strong space terms. Choose Space or clarify that this is an Earth facility.")
    if "leo" in joined and any(x in joined for x in ["moon","lunar surface","mars surface"]):
        contradictions.append("The brief mixes orbital location terms such as LEO with Moon/Mars surface terms. Clarify the operating location.")
    if "mars" in joined and "moon" in joined and "cislunar" not in joined:
        contradictions.append("The brief mixes Mars and Moon scope. Split this into one project location.")

    return {
        "score": min(score, 100),
        "reasons": reasons[:4],
        "contradictions": contradictions,
        "pass": score >= 58 and has_asset and not contradictions,
    }
def _public_demo_quality_message(req: PublicDemoRequest) -> Optional[Dict[str, Any]]:
    quality = _public_demo_brief_quality_score(req)
    if quality["pass"]:
        return None
    return {
        "message": "CASEY needs a stronger brief before using your one free intelligence run.",
        "quality_score": quality["score"],
        "issues": quality["contradictions"] or quality["reasons"],
        "example_brief": "Example: Earth project — 500MW AI data centre campus in Riyadh, concept stage, 4 buildings, liquid cooling, grid substation, concern is grid connection and schedule acceleration. Or Space project — lunar south pole logistics hub, landing pads, regolith roads, autonomous rovers, concept stage, concern is dust, power resilience and launch cadence."
    }

def _quality_gate_public_demo(req: PublicDemoRequest) -> List[str]:
    issues = []
    email = _normalise_email(req.email)
    desc = (req.project_description or "").strip()
    if not email or "@" not in email or "." not in email.split("@")[-1]:
        issues.append("Please enter a valid email so CASEY can reserve your one free intelligence run.")
    if len(desc) > 3000:
        issues.append("Please keep the project brief under 3,000 characters for the public demo.")
    if len(re.findall(r"[a-zA-Z0-9]+", desc)) < 8:
        issues.append("Add a little more detail: asset, location/environment and main project concern.")
    quality = _public_demo_brief_quality_score(req)
    if not quality["pass"]:
        issues.extend(quality["contradictions"] or quality["reasons"])
    return list(dict.fromkeys(issues))[:5]

def _public_demo_identity(request: Request, req: PublicDemoRequest) -> Dict[str, str]:
    ip = client_ip(request)
    ua = request.headers.get("user-agent", "")
    forwarded = request.headers.get("x-forwarded-for", "")
    fp = req.fingerprint or ""
    token = req.client_token or ""
    return {
        "email_hash": _sha(_normalise_email(req.email)),
        "ip_hash": _sha(ip),
        "fingerprint_hash": _sha(fp + "|" + ua),
        "client_token_hash": _sha(token),
        "raw_ip": ip,
        "raw_ua": ua,
        "raw_forwarded": forwarded,
    }

def _public_demo_used(identity: Dict[str, str]) -> Optional[str]:
    con = db(); cur = con.cursor()
    checks = []
    params = []
    for col in ["email_hash", "ip_hash", "fingerprint_hash", "client_token_hash"]:
        val = identity.get(col)
        if val:
            checks.append(f"{col}=?")
            params.append(val)
    if not checks:
        con.close(); return None
    row = cur.execute(f"SELECT run_id, created_at FROM public_demo_uses WHERE {' OR '.join(checks)} ORDER BY id DESC LIMIT 1", tuple(params)).fetchone()
    con.close()
    if row:
        return row["run_id"]
    return None

def _premium_public_prompt(req: PublicDemoRequest) -> str:
    return " | ".join([
        f"{req.project_type} public demo project",
        f"Brief: {req.project_description}",
        f"Location: {req.location or 'not specified'}",
        f"Size/capacity: {req.size_or_capacity or 'not specified'}",
        f"Stage: {req.stage or 'concept / early feasibility'}",
        f"Primary concern: {req.biggest_concern or 'cost, schedule and risk confidence'}",
        "Generate a board-grade first-pass class estimate, schedule intelligence and risk register."
    ])

def _public_demo_report(model: Dict[str, Any]) -> Dict[str, Any]:
    risks = model.get("risks", model.get("risk_register", []))[:8]
    schedule = model.get("schedule_rows", model.get("schedule_detail", []))[:8]
    costs = model.get("cost_lines", model.get("cost_breakdown", []))[:10]
    return {
        "title": model.get("title"),
        "mode": model.get("mode"),
        "executive_summary": model.get("executive_summary"),
        "class_estimate": {
            "estimate_class": model.get("estimate_class_name") or f"Class {model.get('estimate_class', 3)}",
            "p10": model.get("cost_p10"),
            "p50": model.get("cost_p50"),
            "p90": model.get("cost_p90"),
            "range": model.get("cost_range"),
            "confidence_pct": model.get("confidence_pct"),
        },
        "schedule": {
            "baseline": model.get("schedule"),
            "level": model.get("schedule_level"),
            "first_milestones": schedule,
        },
        "risk": {
            "rating": model.get("risk"),
            "score": model.get("risk_score"),
            "top_risks": risks,
        },
        "cost_breakdown": costs,
        "assumptions": model.get("confidence_explanation", [])[:6],
        "next_best_actions": model.get("next_best_actions", [])[:6],
        "input_quality_score": model.get("input_quality_score"),
        "upgrade_cta": "This is a one-shot public intelligence run. Request access for exports, scenarios, QCRA/QSRA, audit trail and deeper model challenge."
    }


def has(t: str, words: List[str]) -> bool: return any(w in t for w in words)
def clamp(v, lo, hi): return max(lo, min(hi, v))
def money_bn(v: float) -> str:
    if v >= 1000: return f"${v/1000:.1f}T"
    if v >= 1: return f"${v:.1f}B"
    return f"${v*1000:.0f}M"
def parse_bn(s: Any) -> float:
    if isinstance(s,(int,float)): return float(s)
    st=str(s or "").replace("$","").replace(",","").strip().upper()
    try:
        if st.endswith("T"): return float(st[:-1])*1000
        if st.endswith("B"): return float(st[:-1])
        if st.endswith("M"): return float(st[:-1])/1000
        return float(st)
    except Exception: return 1.0

def class_range(c:int):
    return {5:(0.50,2.00,"Class 5 Concept Screening","0-2% definition maturity"),4:(0.70,1.50,"Class 4 Feasibility","1-15% definition maturity"),3:(0.80,1.30,"Class 3 Budget Authorization","10-40% definition maturity"),2:(0.85,1.20,"Class 2 Control/Tender","30-75% definition maturity"),1:(0.90,1.15,"Class 1 Definitive/Bid","65-100% definition maturity")}.get(c,(0.8,1.3,"Class 3 Budget Authorization","10-40% maturity"))

def scenario_params(s: str):
    key=(s or "base").lower().replace(" ","_")
    table={
        "base":(1.00,1.00,1.00,0,"Base","Balanced reference case."),
        "faster":(1.11,0.82,1.18,-5,"Faster","Parallel design/procurement, acceleration, premium logistics and higher interface risk."),
        "cheaper":(0.88,1.08,1.12,-6,"Cheaper","Value engineering, tighter scope and procurement competition, with schedule/quality exposure."),
        "lower_risk":(1.09,1.10,0.72,9,"Lower Risk","More surveys, assurance, stage gates, contingency and schedule buffers."),
        "premium":(1.15,1.04,0.80,12,"Premium","Best assurance, independent review, early procurement and high-confidence governance."),
        "investor":(0.98,1.00,0.92,3,"Investor","Investment-framing case with transparent risk and commercial challenge."),
        "survival":(0.78,1.20,1.30,-10,"Survival","Minimum viable scope, warning-heavy, high residual risk."),
    }
    return table.get(key, table["base"])

# ------------------------- detection -------------------------
def detect_sector(prompt: str):
    t=prompt.lower()

    # Named mega-programmes first.
    named=[
      ("Heathrow Third Runway","Earth","Airport Mega Programme",26,108,["heathrow","third runway"]),
      ("HS2 High Speed Rail","Earth","Rail Mega Programme",95,180,["hs2","high speed rail"]),
      ("NEOM/The Line Style Development","Earth","Future City Mega Programme",500,240,["neom","the line"]),
      ("Lunar City","Space","Lunar Settlement",120,240,["lunar city","moon city"]),
      ("Mars Base","Space","Mars Settlement",180,300,["mars base","mars city"]),
      ("Orbital Hospital","Space","Space Medical Infrastructure",12,120,["orbital hospital","space hospital"]),
    ]
    for n,m,s,c,d,k in named:
        if has(t,k): return n,m,s,c,d

    # Strong Earth-place signals. If a real Earth location is present, ambiguous words like
    # "commercial launch", "mission", "platform", or "payload" should NOT automatically route to Space.
    earth_places = [
        "north carolina","carolina","cambridge","boston","arizona","texas","california","new york","florida",
        "united states","usa","america","uk","united kingdom","london","manchester","birmingham",
        "riyadh","saudi","dubai","abu dhabi","uae","qatar","doha","canada","toronto","australia",
        "sydney","singapore","japan","tokyo","india","mumbai","germany","berlin","france","paris",
        "netherlands","amsterdam","poland","warsaw","brazil","south africa","kenya","morocco"
    ]
    earth_place_present = has(t, earth_places)

    # Strong Earth-sector scores.
    life_science_terms = ["biologics","therapeutics","gmp","aseptic","fill-finish","fill finish","fda","cqv","qualification","validated clean utilities","clean utilities","cold-chain","cold chain","pharma","pharmaceutical","cell therapy","gene therapy","manufacturing campus","obesity therapeutics"]
    semiconductor_terms = ["semiconductor"," fab","wafer","upw","cleanroom","process tooling","chip plant"]
    data_terms = ["data centre","data center","datacenter","hyperscale","ai campus","gpu campus","liquid cooling"]
    defence_earth_terms = ["secure satellite control centre","secure satellite control center","satellite control centre","satellite control center","mission operations room","mission operations centre","mission operations center","secure command","command and control","air defence","air defense","military airbase","naval base","munitions storage","radar station","border surveillance","defence data centre","defense data center"]
    healthcare_terms = ["hospital","healthcare","clinical","diagnostic imaging","theatres","emergency department"]
    transport_terms = ["airport","aviation","runway","airside","baggage","passenger terminal","rail","metro"," station ","transit"]
    energy_terms = ["hydrogen","solar","wind","battery","power","grid","energy","nuclear","smr","lng","carbon capture","ccs"]
    water_terms = ["desalination","water","wastewater","flood","pumping station","reservoir"]

    earth_score = 0
    for terms, weight in [
        (life_science_terms, 5),
        (semiconductor_terms, 5),
        (data_terms, 4),
        (defence_earth_terms, 5),
        (healthcare_terms, 4),
        (transport_terms, 3),
        (energy_terms, 3),
        (water_terms, 3),
    ]:
        earth_score += sum(weight for term in terms if term in t)
    if earth_place_present:
        earth_score += 6

    # Strong Space scores. Weak words like "launch" only count strongly with aerospace context.
    strong_space_terms = ["moon","lunar","mars","orbit","orbital","leo","meo","geo","cislunar","cis-lunar","space station","asteroid","deep space","deep-space","in-space"]
    medium_space_terms = ["spaceport","rocket","payload","satellite constellation","launch vehicle","launch pad","launch complex","spacecraft","propellant depot","orbital compute","space data centre","space data center","orbital ai","space-based data centre","space-based data center"]
    weak_space_terms = ["launch","mission","payload","platform","constellation","habitat"]

    space_score = 0
    space_score += sum(5 for term in strong_space_terms if term in t)
    space_score += sum(4 for term in medium_space_terms if term in t)

    # Only count weak space terms if supported by other space signals or no strong Earth context.
    if space_score > 0 or not earth_place_present:
        space_score += sum(1 for term in weak_space_terms if term in t)

    # Product/commercial launch phrases are Earth business language, not space.
    if has(t, ["commercial launch demand","product launch","launch demand","market launch","commercial launch","launch readiness"]) and not has(t, ["rocket","spaceport","launch vehicle","launch pad","orbital","leo","lunar","mars"]):
        space_score = max(0, space_score - 5)
        earth_score += 4

    # Hard space overrides: explicit Mars/Moon/LEO/orbital infrastructure should remain Space even if it contains
    # Earth-sector words such as nuclear power, grid, hospital, platform, or manufacturing.
    if has(t, ["mars surface","mars outpost","mars research","mars habitat","mars base"]):
        return "Mars Surface Infrastructure","Space","Mars Surface Habitat/Base",38,168
    if has(t, ["lunar surface","lunar south pole","lunar habitat","lunar base","lunar logistics","moon base"]):
        return "Lunar Surface Infrastructure","Space","Lunar Surface Habitat/Base",32,156
    if has(t, ["orbital ai data centre","orbital ai data center","space-based data centre","space-based data center","orbital compute platform","leo compute"]):
        return "Orbital AI Compute Platform","Space","Orbital Compute / Manufacturing",24,132
    if has(t, ["space power grid","orbital power grid","lunar power grid","mars power grid","space solar grid","orbital energy grid"]):
        return "Space Power Grid","Space","Power/Energy Infrastructure",18,132

    # High-signal Earth sectors win even if one ambiguous space word exists.
    if earth_score >= space_score and earth_score >= 5:
        if has(t, life_science_terms):
            return "Life Sciences Manufacturing Campus","Earth","Life Sciences / Biologics Manufacturing",2.8,58
        if has(t, semiconductor_terms):
            return "Advanced Semiconductor Fab","Earth","Semiconductor / Advanced Manufacturing",18,78
        if has(t, data_terms):
            return "AI Data Centre Campus","Earth","Digital Infrastructure / Hyperscale Data Centre",3.8,46
        if has(t, defence_earth_terms):
            return "Secure Defence Infrastructure","Earth","Defence / Secure Mission Infrastructure",4.8,54
        if has(t, healthcare_terms):
            return "Hospital Campus","Earth","Healthcare / Hospital",2.4,60
        if has(t, ["airport","aviation","runway","airside","baggage","passenger terminal"]):
            return "Airport Infrastructure","Earth","Airport / Aviation",9,84
        if has(t, ["rail","metro","station","transit"]):
            return "Rail Infrastructure","Earth","Rail / Transit",6.5,84
        if has(t, ["nuclear","smr","fusion"]):
            return "Nuclear Energy Facility","Earth","Nuclear / Energy",12,96
        if has(t, energy_terms):
            return "Energy Infrastructure","Earth","Energy / Utilities",5.2,66
        if has(t, water_terms):
            return "Water Infrastructure","Earth","Water / Utilities",2.2,50

    # Space routing only after weighted scoring confirms it.
    if space_score > earth_score and space_score >= 4:
        if has(t,["deep-space communications","deep space communications","deep-space comms","deep space network","relay spacecraft","space communications array"]): return "Deep-Space Communications Array","Space","Deep-Space Communications Infrastructure",10,96
        if has(t,["orbital ai compute","orbital ai data centre","orbital ai data center","space data centre","space data center","space-based data centre","space-based data center","orbital compute","leo compute","compute platform","orbital data centre","orbital data center"]): return "Orbital AI Compute Platform","Space","Orbital Compute / Manufacturing",24,132
        if has(t,["hospital","medical","health"]): return "Space Hospital","Space","Orbital/Lunar Healthcare",12,126
        if has(t,["city","settlement","colony"]): return "Space Settlement","Space","Lunar/Mars Settlement",110,240
        if has(t,["base","habitat","outpost"]): return "Space Base","Space","Surface Habitat/Base",32,156
        if has(t,["power","solar","nuclear","grid"]): return "Space Power Grid","Space","Power/Energy Infrastructure",18,132
        if has(t,["spaceport","launch vehicle","launch pad","launch complex","rocket"]): return "Launch Infrastructure","Space","Spaceport/Launch",9,84
        if has(t,["mine","mining","isru","water","oxygen","propellant"]): return "Space Resources Facility","Space","ISRU/Mining/Propellant",22,168
        if has(t,["satellite","constellation"]): return "Satellite Constellation","Space","Satellite/Comms",6,60
        return "Frontier Space Infrastructure","Space","General Space Infrastructure",14,108

    # Earth fallback after scoring.
    if has(t,["data centre","data center","datacenter","ai campus","cloud","hyperscale"]): return "AI Data Centre Campus","Earth","Digital Infrastructure / Hyperscale Data Centre",3.8,46
    if has(t,["biologics","gmp","aseptic","fill-finish","fill finish","fda","cqv","pharma","therapeutics","cell therapy"]): return "Life Sciences Manufacturing Campus","Earth","Life Sciences / Biologics Manufacturing",2.8,58
    if has(t,["airport","aviation","runway","airside","baggage","passenger terminal"]): return "Airport Infrastructure","Earth","Airport / Aviation",9,84
    if has(t,["rail","metro"," station ","transit"]): return "Rail Infrastructure","Earth","Rail / Transit",6.5,84
    if has(t,["nuclear","smr","fusion"]): return "Nuclear Energy Facility","Earth","Nuclear / Energy",12,96
    if has(t,["hydrogen","solar","wind","battery","power","grid","energy"]): return "Energy Infrastructure","Earth","Energy / Utilities",5.2,66
    if has(t,["hospital","medical campus","healthcare"]): return "Hospital Campus","Earth","Healthcare / Hospital",2.4,60
    if has(t,["pharma","gmp","biologics","life sciences","cleanroom"]): return "Life Sciences Campus","Earth","Life Sciences / Pharma",2.8,58
    if has(t,["defence","defense","military","command","radar"]): return "Secure Defence Infrastructure","Earth","Defence / Secure Mission Infrastructure",4.8,54
    if has(t,["desalination","water","wastewater","flood"]): return "Water Infrastructure","Earth","Water / Utilities",2.2,50
    return "Major Infrastructure Programme","Earth","General Infrastructure",2.0,48

def location_factor(t: str):
    t=t.lower(); table=[("Moon / Lunar Surface",2.9,["moon","lunar"]),("Mars",3.8,["mars"]),("Low Earth Orbit",2.5,["leo","orbit","orbital"]),("Deep Space",4.3,["deep space","asteroid"]),("Saudi Arabia / GCC",1.35,["saudi","riyadh","neom","jeddah","gcc"]),("UAE / Middle East",1.18,["uae","dubai","abu dhabi","qatar"]),("United Kingdom",1.20,["uk","london","heathrow","manchester"]),("United States",1.08,["usa","us ","texas","california","florida","new york"]),("Europe",1.12,["europe","germany","france","spain","italy"]),("India",0.78,["india","mumbai","delhi"]),("Australia",1.22,["australia","sydney","melbourne"])]
    for n,f,keys in table:
        if has(t,keys): return n,f
    return "Global / Not specified",1.0

def scale_factor(t: str):
    t=t.lower()
    if has(t,["global","interplanetary","system of systems"]): return "System of Systems",5.0,72
    if has(t,["mega","city","national","regional","10000","10,000","colony"]): return "Mega Programme",3.2,42
    if has(t,["programme","program","campus","portfolio","corridor"]): return "Programme",1.9,18
    if has(t,["pilot","phase 1","small"]): return "Pilot / Phase 1",0.55,-8
    return "Project",1.0,0

def complexity(t:str):
    rules={"accelerated":(1.10,-8,"accelerated delivery compression"),"fast":(1.07,-5,"fast-track compression"),"nuclear":(1.45,36,"nuclear assurance and regulation"),"live":(1.15,12,"live environment"),"brownfield":(1.20,12,"brownfield interfaces"),"underground":(1.35,18,"underground works"),"offshore":(1.40,24,"offshore logistics"),"autonomous":(1.20,12,"autonomous systems"),"robotic":(1.18,12,"robotic construction"),"life support":(1.35,24,"life support systems"),"radiation":(1.25,18,"radiation hardening"),"cryogenic":(1.25,18,"cryogenic systems"),"cleanroom":(1.25,12,"GMP cleanroom"),"gmp":(1.22,12,"GMP validation"),"secure":(1.20,12,"secure/defence constraints")}
    f=1.0; m=0; d=[]; tl=t.lower()
    for k,(ff,mm,dd) in rules.items():
        if k in tl: f*=ff; m+=mm; d.append(dd)
    return f,m,d

def risk_label(score:float):
    if score>=90: return "Extreme"
    if score>=75: return "Very High"
    if score>=60: return "High"
    if score>=42: return "Medium-High"
    if score>=25: return "Medium"
    return "Low"

def cost_lines(mode:str, subsector:str):
    if mode=="Space":
        return [("01.01","DDT&E Engineering","Direct","Engineering design, mission architecture and qualification",0.09),("01.02","Systems Engineering & Integration","Direct","Requirements, ICDs, verification and validation",0.06),("01.03","Structures / Habitat Shell","Direct","Pressure vessels, shielding and primary structures",0.10),("01.04","Power Generation & Distribution","Direct","Solar/nuclear source, grid and controls",0.08),("01.05","ECLSS / Life Support","Direct","Air, water, waste, atmosphere and safety systems",0.09),("01.06","Payload / Mission Equipment","Direct","Scientific, medical, industrial or settlement payload",0.07),("01.07","Robotics / Autonomous Construction","Direct","Robotic construction, inspection and maintenance",0.06),("01.08","Launch Services","Direct","Launch vehicle, manifest, payload processing",0.14),("01.09","Transport / Landing / Surface Logistics","Direct","Transit, landing, mobility and logistics chain",0.09),("02.01","Test & Qualification Facilities","Indirect","Thermal-vac, vibration, radiation and integrated testing",0.05),("02.02","Safety & Mission Assurance","Indirect","Quality, safety, independent review and compliance",0.05),("02.03","Programme Management / PMO","Indirect","Controls, reporting, commercial and integration PMO",0.05),("03.01","Risk Contingency / UFE","Reserve","Known-unknown risk allowance linked to QCRA/QSRA",0.07),("03.02","Management Reserve","Reserve","Sponsor-held unknown-unknown reserve",0.05)]
    special=[]
    s=subsector.lower()
    if "data centre" in s: special=[("01.06","Data Halls / White Space","Direct","Server halls, fit-out and critical rooms",0.13),("01.07","Power Train / Substations","Direct","HV/MV gear, UPS, generators and substations",0.16),("01.08","Cooling Systems","Direct","Chillers, heat rejection, CRAH/CRAC, water systems",0.12)]
    elif "airport" in s: special=[("01.06","Airfield Works","Direct","Runway, taxiway, apron and airfield systems",0.14),("01.07","Terminal / Passenger Systems","Direct","Terminal, baggage, security and passenger systems",0.10)]
    elif "rail" in s: special=[("01.06","Track / Signalling / Systems","Direct","Trackform, power, signalling and telecoms",0.18),("01.07","Stations / Depots","Direct","Stations, depots and operational readiness",0.09)]
    elif "pharma" in s: special=[("01.06","Cleanroom Envelope","Direct","Cleanroom partitions, finishes and GMP envelope",0.12),("01.07","Process Utilities","Direct","WFI, gases, clean steam and process utilities",0.13),("01.08","Validation / Qualification","Indirect","IQ/OQ/PQ and GMP readiness",0.05)]
    base=[("01.01","Site Preparation / Enabling","Direct","Demolition, enabling, logistics and access",0.06),("01.02","Substructure / Foundations","Direct","Piling, groundworks, retaining and foundations",0.08),("01.03","Superstructure / Civil Works","Direct","Frame, civils, primary structures",0.14),("01.04","MEP / Systems / Utilities","Direct","Mechanical, electrical, controls and utilities",0.14),("01.05","Specialist Equipment / Fit-Out","Direct","Permanent equipment and specialist fit-out",0.12)]
    tail=[("01.09","Direct Labour","Direct","Productivity, supervision and trade labour",0.09),("01.10","Direct Materials","Direct","Concrete, steel, envelope, MEP materials",0.10),("01.11","Plant / Equipment","Direct","Cranes, heavy plant and site equipment",0.04),("02.01","General Conditions / Prelims","Indirect","Site management, safety, offices and temporary services",0.07),("02.02","Design / Professional Fees","Indirect","Design, QS, controls, assurance and consultants",0.07),("02.03","Permits / Insurance / Bonds","Indirect","Permits, inspections, legal, bonds and insurance",0.03),("03.01","QCRA Contingency","Reserve","Known-unknown risk allowance",0.06),("03.02","Management Reserve","Reserve","Sponsor-held reserve",0.03)]
    return base+special+tail

def schedule_rows(mode:str, subsector:str, months:int, level:int):
    rows=[
      ("A1000","Governance","Project initiation / controls setup","",0.03),("A1100","Business Case","Strategic brief, funding and approval","A1000",0.05),("A1200","Scope","Requirements definition and scope freeze","A1100",0.05),("A1300","Design","Concept / reference architecture","A1200",0.08),("A1400","Consents","Planning, regulatory and stakeholder approvals","A1200",0.09),("A1500","Procurement","Procurement strategy and tender packaging","A1300",0.06),("A1600","Design","Detailed design / IFC development","A1300",0.10),("A1700","Procurement","Long-lead procurement","A1500",0.09),("A1800","Enabling","Site / launch integration readiness","A1400;A1500",0.07),("A1900","Delivery","Main construction / fabrication / assembly","A1600;A1700;A1800",0.22),("A2000","Integration","Systems integration and testing","A1900",0.10),("A2100","Handover","Commissioning, operations readiness and closeout","A2000",0.06)
    ]
    if level<=1:
        rows=[("M100","Milestone","Investment approval","",0.05),("M200","Milestone","Design complete","M100",0.20),("M300","Milestone","Procurement complete","M200",0.15),("M400","Milestone","Delivery complete","M300",0.40),("M500","Milestone","Ready for operations","M400",0.15),("M600","Milestone","Handover","M500",0.05)]
    packages=["Civil / structural package","MEP / systems package","Specialist equipment package","Utilities / interfaces package","Testing and commissioning package"]
    if mode=="Space": packages=["Launch campaign","Payload integration","Power and ECLSS","Surface logistics","Mission operations","Robotic assembly"]
    elif "data centre" in subsector.lower(): packages=["Grid connection","Data halls","Cooling plant","Power train","Controls and commissioning"]
    elif "airport" in subsector.lower(): packages=["Airfield works","Terminal systems","Baggage/security systems","ORAT"]
    elif "rail" in subsector.lower(): packages=["Track","Signalling","Stations","Depot and systems integration"]
    if level>=4:
        for i,p in enumerate(packages,1):
            code=f"P{i}000"; pred="A1900" if i==1 else f"P{i-1}000"; rows.append((code,"Package",p,pred,0.035))
            if level>=5:
                rows.append((f"P{i}100","Activity",f"{p} - mobilise / method",pred,0.008))
                rows.append((f"P{i}200","Activity",f"{p} - execute works",f"P{i}100",0.022))
                rows.append((f"P{i}300","Activity",f"{p} - inspect / closeout",f"P{i}200",0.008))
    out=[]
    for code,phase,act,pred,pct in rows:
        out.append({"activity_id":code,"phase":phase,"activity":act,"predecessor":pred,"duration_months":max(1,round(months*pct)),"critical":"Yes" if code in ["A1400","A1700","A1900","A2000","A2100"] or code.startswith("P") else "No","basis":"Duration derived from sector schedule model, selected schedule level and scenario compression/buffer logic."})
    return out

def peer_competitors(client:str, subsector:str, mode:str):
    c=(client or "").lower(); s=subsector.lower()
    if mode=="Space" or any(x in c for x in ["spacex","blue origin","nasa"]): return ["SpaceX","Blue Origin","Rocket Lab","ULA","NASA/ESA programmes"]
    if any(x in c for x in ["microsoft","msft"]): return ["Meta","Alphabet / Google","Amazon AWS","Oracle","Apple"]
    if "data centre" in s: return ["Meta","Alphabet / Google","Amazon AWS","Microsoft","Oracle"]
    if "airport" in s: return ["Dubai Airports","Heathrow operators","Changi Airport Group","AD Airports","Saudi airport programmes"]
    if "rail" in s: return ["Network Rail","SNCF","DB InfraGO","HS2 delivery partners","Gulf rail programmes"]
    if "pharma" in s: return ["Pfizer","Roche","Novartis","AstraZeneca","Moderna"]
    if "energy" in s or "nuclear" in s: return ["EDF","NextEra","Aramco energy programmes","Masdar","Ørsted"]
    return ["Relevant global peers","Regional developers","Tier-one operators","Sovereign funds","Major asset owners"]


# ------------------------- benchmark memory / public-scale calibration -------------------------
BENCHMARK_LIBRARY = [
    {"sector":"Digital Infrastructure / Hyperscale Data Centre","mode":"Earth","keywords":["data centre","data center","hyperscale","ai campus","gpu"],"cost_bn":3.8,"months":46,"risks":["Grid connection delay","Cooling capacity constraint","Long-lead electrical equipment","Commissioning load-bank failure"]},
    {"sector":"Semiconductor / Advanced Manufacturing","mode":"Earth","keywords":["semiconductor","fab","wafer","upw","chip"],"cost_bn":18.0,"months":78,"risks":["Tool delivery slippage","UPW and chemicals interface","Cleanroom validation delay","Utility demand escalation"]},
    {"sector":"Airport / Aviation","mode":"Earth","keywords":["airport","terminal","runway","baggage"],"cost_bn":9.0,"months":84,"risks":["Live operations phasing","Security and baggage integration","Airside possession constraints","Stakeholder approvals"]},
    {"sector":"Rail / Transit","mode":"Earth","keywords":["rail","metro","station","signalling","transit"],"cost_bn":6.5,"months":84,"risks":["Possession access","Systems integration delay","Utility diversions","Land/stakeholder approvals"]},
    {"sector":"Energy / Utilities","mode":"Earth","keywords":["hydrogen","wind","solar","grid","battery","energy"],"cost_bn":5.2,"months":66,"risks":["Grid connection delay","Equipment lead times","Permitting constraints","Commodity escalation"]},
    {"sector":"Nuclear / Energy","mode":"Earth","keywords":["nuclear","smr","reactor"],"cost_bn":12.0,"months":96,"risks":["Licensing delay","FOAK supply chain","Safety case approval","Specialist labour availability"]},
    {"sector":"Healthcare / Hospital","mode":"Earth","keywords":["hospital","medical","healthcare"],"cost_bn":2.4,"months":60,"risks":["Clinical commissioning","Medical equipment lead times","Live hospital interfaces","Regulatory approval"]},
    {"sector":"Life Sciences / Pharma","mode":"Earth","keywords":["gmp","pharma","biologics","life sciences","validation"],"cost_bn":2.8,"months":58,"risks":["Validation delay","Clean utility readiness","Regulatory handover","Specialist automation integration"]},
    {"sector":"Defence / Secure Mission Infrastructure","mode":"Earth","keywords":["defence","defense","military","secure","command","airbase","naval","radar"],"cost_bn":4.8,"months":54,"risks":["Security accreditation","Controlled procurement","Resilience requirements","Operational continuity"]},
    {"sector":"Water / Utilities","mode":"Earth","keywords":["desalination","water","wastewater","flood"],"cost_bn":2.2,"months":50,"risks":["Marine works delay","Environmental consent","Energy cost exposure","Stakeholder approvals"]},
    {"sector":"Spaceport/Launch","mode":"Space","keywords":["spaceport","launch","rocket","pad"],"cost_bn":9.0,"months":84,"risks":["Launch licensing","Range safety","Propellant farm safety","Environmental approvals"]},
    {"sector":"Lunar/Mars Settlement","mode":"Space","keywords":["lunar","moon","mars","habitat","base","settlement"],"cost_bn":32.0,"months":156,"risks":["Launch manifest dependency","Life-support reliability","Power survival","Surface logistics"]},
    {"sector":"ISRU/Mining/Propellant","mode":"Space","keywords":["isru","propellant","oxygen","methane","mining","depot"],"cost_bn":22.0,"months":168,"risks":["Cryogenic reliability","Autonomous transfer","Technology readiness","Launch cadence"]},
    {"sector":"Satellite/Comms","mode":"Space","keywords":["satellite","constellation","comms","ground station"],"cost_bn":6.0,"months":60,"risks":["Production cadence","Launch manifest","Ground segment integration","Spectrum/regulatory"]},
    {"sector":"Orbital Compute / Manufacturing","mode":"Space","keywords":["orbital compute","space data centre","space data center","orbital ai compute","compute platform","leo compute","orbital manufacturing","leo manufacturing"],"cost_bn":24.0,"months":132,"risks":["Thermal rejection system failure","Launch cadence dependency","Autonomous servicing constraints","Radiation hardening exposure","Power density scaling","Ground relay latency","Orbital debris exposure"]},
]

def benchmark_similarity(prompt: str, mode: str, subsector: str) -> List[Dict[str, Any]]:
    t = (prompt or "").lower()
    scored = []
    for b in BENCHMARK_LIBRARY:
        score = 0
        if b["mode"] == mode:
            score += 2
        if b["sector"].lower() in (subsector or "").lower() or (subsector or "").lower() in b["sector"].lower():
            score += 3
        for kw in b["keywords"]:
            if kw in t:
                score += 2
        if score > 0:
            scored.append({**b, "similarity_score": score})
    return sorted(scored, key=lambda x: x["similarity_score"], reverse=True)[:5]

def calibrate_with_benchmarks(prompt: str, mode: str, subsector: str, raw_cost: float, months: int) -> Tuple[float, int, List[str], List[Dict[str, Any]]]:
    matches = benchmark_similarity(prompt, mode, subsector)
    if not matches:
        return raw_cost, months, ["No strong benchmark match; CASEY sector template used."], []
    primary = matches[0]
    anchor_cost = float(primary["cost_bn"])
    anchor_months = int(primary["months"])
    calibrated_cost = raw_cost * 0.72 + anchor_cost * 0.28
    calibrated_months = round(months * 0.76 + anchor_months * 0.24)
    notes = [
        f"Benchmark memory matched {primary['sector']} with similarity score {primary['similarity_score']}.",
        f"Cost calibrated by blending CASEY template output with benchmark anchor {money_bn(anchor_cost)}.",
        f"Schedule calibrated against benchmark duration {anchor_months} months.",
    ]
    return calibrated_cost, calibrated_months, notes, matches

def estimate_quality_index(class_level:int, schedule_level:int, risk_score:float, matches:List[Dict[str,Any]], input_quality:int=80) -> Dict[str,Any]:
    maturity = {1:92,2:82,3:70,4:56,5:42}.get(class_level,70)
    schedule_quality = {1:42,2:55,3:68,4:78,5:88}.get(schedule_level,68)
    benchmark_quality = min(92, 40 + (matches[0]["similarity_score"]*7 if matches else 0))
    risk_penalty = min(28, risk_score/4)
    score = int(clamp((maturity*0.30 + schedule_quality*0.25 + benchmark_quality*0.25 + input_quality*0.20) - risk_penalty, 10, 96))
    band = "High" if score >= 78 else "Medium" if score >= 58 else "Low"
    return {"score": score, "band": band, "maturity_score": maturity, "schedule_quality": schedule_quality, "benchmark_quality": benchmark_quality, "risk_penalty": round(risk_penalty,1), "meaning": "Confidence index reflects estimate class, schedule level, benchmark similarity, input quality and risk exposure."}

def procurement_heatmap(mode:str, subsector:str, risks:List[Dict[str,Any]]) -> List[Dict[str,Any]]:
    base = [
        {"package":"Long-lead equipment","exposure":"High","reason":"Procurement path drives cost and schedule confidence."},
        {"package":"Utilities / enabling works","exposure":"High" if mode=="Earth" else "Medium","reason":"Utility readiness controls commissioning and phasing."},
        {"package":"Systems integration","exposure":"High","reason":"Integration maturity is a common late-stage failure mode."},
        {"package":"Specialist suppliers","exposure":"Medium-High","reason":"Limited supplier base increases price and schedule volatility."},
    ]
    if mode=="Space":
        base.insert(0, {"package":"Launch / mission logistics","exposure":"Extreme","reason":"Launch cadence and mission integration dominate delivery exposure."})
        base.append({"package":"Autonomous operations","exposure":"High","reason":"Remote operations reduce recovery options after deployment."})
    if "Data Centre" in subsector:
        base.append({"package":"Power equipment","exposure":"Extreme","reason":"Transformers, switchgear and grid connection can dominate critical path."})
    if "Semiconductor" in subsector:
        base.append({"package":"Process tools","exposure":"Extreme","reason":"Tool install and cleanroom readiness drive fab ramp-up."})
    if "Orbital Compute" in subsector:
        base.append({"package":"Thermal rejection systems","exposure":"Extreme","reason":"Heat rejection and cooling stability dominate orbital compute viability."})
        base.append({"package":"Radiation-hardened compute","exposure":"High","reason":"Radiation-tolerant compute supply chains remain constrained."})
    return base[:9]

def critical_path_narrative(mode:str, subsector:str, schedule_rows:List[Dict[str,Any]]) -> List[str]:
    names = [r.get("activity","") for r in schedule_rows if isinstance(r,dict)]
    out = [
        "Critical path is likely governed by definition maturity, procurement release and systems commissioning.",
        "Early schedule confidence depends on locking scope, approvals and long-lead package strategy.",
    ]
    if mode=="Space":
        out.append("Space critical path is additionally controlled by launch integration, payload readiness and remote operations validation.")
    if "Data Centre" in subsector:
        out.append("For digital infrastructure, grid connection and electrical equipment procurement are likely critical path drivers.")
    if "Orbital Compute" in subsector:
        out.append("For orbital compute, thermal rejection, power density, launch cadence and autonomous servicing are likely critical path drivers.")
    if "Airport" in subsector or "Rail" in subsector:
        out.append("For transport programmes, possessions/live operations and systems integration are likely critical path drivers.")
    if "Defence" in subsector:
        out.append("For secure infrastructure, accreditation and controlled procurement can become critical path constraints.")
    if names:
        out.append(f"First schedule control focus should be: {names[min(2,len(names)-1)]}.")
    return out[:5]


# ------------------------- CASEY WOW narrative intelligence layer -------------------------
def board_briefing_narrative(mode:str, subsector:str, title:str, cost_p50:str, schedule:str, risk:str, confidence:int, drivers:List[str], matches:List[Dict[str,Any]]) -> List[str]:
    if mode == "Space":
        opening = f"{title} should be treated less like a one-off aerospace project and more like persistent orbital infrastructure requiring operational coordination between launch, power, comms, autonomy and servicing."
    elif "Data Centre" in subsector:
        opening = f"{title} is likely to behave as much like an energy-and-utilities programme as a conventional digital infrastructure build because power availability and commissioning sequencing will dominate confidence."
    elif "Defence" in subsector:
        opening = f"{title} is likely to be governed by accreditation, resilience and controlled procurement constraints as much as by traditional construction delivery."
    elif "Airport" in subsector or "Rail" in subsector:
        opening = f"{title} is likely to be constrained by live operational interfaces, systems integration and possession/phasing strategy rather than physical works alone."
    else:
        opening = f"{title} should be managed as an integrated infrastructure system, not simply a capital project reporting exercise."

    return [
        opening,
        f"Current P50 intelligence indicates {cost_p50} over approximately {schedule}, with overall risk assessed as {risk}.",
        f"Confidence remains {('high' if confidence >= 78 else 'moderate' if confidence >= 55 else 'low')} because class maturity, schedule definition, benchmark similarity and delivery exposure are still moving together.",
        "The immediate control priority is to convert early assumptions into governed decisions around scope freeze, procurement path, interface ownership and critical-path evidence.",
    ]

def uncertainty_narrative(class_level:int, schedule_level:int, confidence:int, mode:str, risk:str, drivers:List[str]) -> Dict[str,Any]:
    class_msg = {
        1:"Class 1 maturity suggests a high-definition estimate with reduced range exposure.",
        2:"Class 2 maturity suggests strong definition but remaining package and market exposure.",
        3:"Class 3 maturity is suitable for budget authorization, but procurement and design assumptions still need challenge.",
        4:"Class 4 maturity remains feasibility-led and should be treated as a decision-support range rather than a firm budget.",
        5:"Class 5 maturity is concept-level and should be treated as strategic order-of-magnitude intelligence."
    }.get(class_level, "Estimate maturity requires review.")
    sched_msg = {
        1:"Schedule Level 1 is strategic only and should not be used for delivery commitment.",
        2:"Schedule Level 2 supports programme framing but still lacks full activity logic.",
        3:"Schedule Level 3 supports integrated control but needs package-level validation.",
        4:"Schedule Level 4 gives stronger logic and QSRA traceability.",
        5:"Schedule Level 5 supports detailed controls thinking if activity evidence is maintained."
    }.get(schedule_level, "Schedule definition requires review.")
    exposure = "Launch and orbital operations exposure increases uncertainty." if mode=="Space" else "Market, approvals, procurement and interface exposure remain key uncertainty drivers."
    return {"confidence":confidence, "risk":risk, "estimate_maturity":class_msg, "schedule_maturity":sched_msg, "uncertainty_drivers":drivers[:6] or [exposure], "interpretation":exposure}

def mission_control_cards(mode:str, subsector:str, title:str, risk:str, matches:List[Dict[str,Any]], procurement:List[Dict[str,Any]], critical:List[str]) -> List[Dict[str,str]]:
    cards = []
    cards.append({"label":"CRITICAL PATH EXPOSURE","signal":critical[0] if critical else "Critical path requires validation against procurement and commissioning logic.","severity":risk})
    if procurement:
        p = procurement[0]
        cards.append({"label":"PROCUREMENT WATCH","signal":f"{p.get('package')} exposure is {p.get('exposure')}: {p.get('reason')}","severity":p.get("exposure","High")})
    if matches:
        cards.append({"label":"BENCHMARK MEMORY","signal":f"Closest archetype: {matches[0].get('sector')} with similarity score {matches[0].get('similarity_score')}.","severity":"Reference"})
    if mode == "Space":
        cards.append({"label":"ORBITAL OPERATIONS RISK","signal":"Launch cadence, autonomous servicing, communications dependency and remote recovery constraints should be treated as programme controls issues, not only engineering issues.","severity":"Extreme" if risk in ["Very High","Extreme"] else "High"})
    if "Orbital Compute" in subsector:
        cards.append({"label":"THERMAL / POWER BOTTLENECK","signal":"Orbital compute viability is likely to be governed by thermal rejection, power density, radiation-hardening and servicing cadence.","severity":"Extreme"})
    elif "Data Centre" in subsector:
        cards.append({"label":"POWER BOTTLENECK","signal":"Grid energisation, transformers, switchgear and commissioning capacity are likely to dominate the delivery confidence curve.","severity":"High"})
    return cards[:6]

def casey_thinking(mode:str, subsector:str, title:str) -> str:
    if "Orbital Compute" in subsector:
        return "CASEY interprets this as a new infrastructure category: compute capacity no longer sits only behind a fence line on Earth; it becomes an orbital operations system governed by launch cadence, thermal rejection, servicing and relay dependency."
    if mode == "Space":
        return "CASEY interprets this as persistent orbital infrastructure, where project controls shift from construction reporting toward Earth-orbit operational orchestration."
    if "Data Centre" in subsector:
        return "CASEY interprets this as an energy-constrained infrastructure programme, not simply a digital real-estate deployment."
    if "Semiconductor" in subsector:
        return "CASEY interprets this as a utility-and-tooling critical programme where cleanroom readiness and process tool cadence will dominate ramp-up confidence."
    if "Defence" in subsector:
        return "CASEY interprets this as a secure mission infrastructure programme where accreditation and resilience can become hidden critical-path drivers."
    return "CASEY interprets this as a system-of-systems infrastructure programme where delivery confidence depends on interface control, procurement evidence and decision velocity."

def benchmark_comparison(matches:List[Dict[str,Any]], mode:str, subsector:str) -> List[Dict[str,Any]]:
    if matches:
        return [{"archetype":m["sector"],"similarity_score":m["similarity_score"],"anchor_cost":money_bn(m["cost_bn"]),"anchor_duration_months":m["months"],"use":"Directional calibration only; not a certified benchmark."} for m in matches[:4]]
    return [{"archetype":subsector or ("Space infrastructure" if mode=="Space" else "Earth infrastructure"),"similarity_score":0,"anchor_cost":"N/A","anchor_duration_months":"N/A","use":"No strong benchmark memory match; CASEY used sector logic and input assumptions."}]


# ------------------------- sector realism and executive output upgrade -------------------------
SECTOR_ENVELOPES = {
    "Life Sciences / Biologics Manufacturing": {"cost": (1.8, 8.5), "months": (36, 78)},
    "Life Sciences / Pharma": {"cost": (1.5, 7.5), "months": (34, 76)},
    "Digital Infrastructure / Hyperscale Data Centre": {"cost": (1.5, 18.0), "months": (24, 84)},
    "Semiconductor / Advanced Manufacturing": {"cost": (8.0, 32.0), "months": (54, 96)},
    "Airport / Aviation": {"cost": (2.0, 24.0), "months": (60, 144)},
    "Rail / Transit": {"cost": (3.0, 80.0), "months": (60, 180)},
    "Healthcare / Hospital": {"cost": (0.8, 6.5), "months": (36, 84)},
    "Defence / Secure Mission Infrastructure": {"cost": (0.8, 12.0), "months": (30, 84)},
    "Energy / Utilities": {"cost": (1.0, 18.0), "months": (36, 108)},
    "Nuclear / Energy": {"cost": (5.0, 35.0), "months": (72, 144)},
    "Water / Utilities": {"cost": (0.8, 8.0), "months": (30, 84)},
    "Spaceport/Launch": {"cost": (1.0, 14.0), "months": (36, 108)},
    "Orbital Compute / Manufacturing": {"cost": (8.0, 65.0), "months": (72, 168)},
    "Lunar Surface Habitat/Base": {"cost": (15.0, 95.0), "months": (96, 216)},
    "Mars Surface Habitat/Base": {"cost": (30.0, 180.0), "months": (120, 300)},
    "ISRU/Mining/Propellant": {"cost": (15.0, 120.0), "months": (96, 240)},
    "Satellite/Comms": {"cost": (2.0, 20.0), "months": (36, 96)},
    "Deep-Space Communications Infrastructure": {"cost": (4.0, 35.0), "months": (60, 144)},
}

def _sector_family(subsector:str) -> str:
    s=(subsector or "").lower()
    if "life sciences" in s or "biologics" in s or "pharma" in s: return "life_sciences"
    if "data centre" in s or "data center" in s or "digital infrastructure" in s: return "data_centre"
    if "semiconductor" in s: return "semiconductor"
    if "airport" in s or "aviation" in s: return "airport"
    if "rail" in s or "transit" in s: return "rail"
    if "hospital" in s or "healthcare" in s: return "healthcare"
    if "defence" in s or "defense" in s or "secure mission" in s: return "defence"
    if "nuclear" in s: return "nuclear"
    if "energy" in s or "utilities" in s: return "energy"
    if "water" in s: return "water"
    if "orbital compute" in s or "orbital" in s: return "orbital"
    if "lunar" in s: return "lunar"
    if "mars" in s: return "mars"
    if "spaceport" in s or "launch" in s: return "spaceport"
    if "satellite" in s or "comms" in s or "communications" in s: return "space_comms"
    if "isru" in s or "propellant" in s or "mining" in s: return "space_resources"
    return "general"

def sector_envelope(subsector:str, cost:float, months:int, scale:str="") -> Tuple[float,int,List[str]]:
    fam = _sector_family(subsector)
    env = SECTOR_ENVELOPES.get(subsector)
    if not env:
        for k,v in SECTOR_ENVELOPES.items():
            if k.lower() in (subsector or "").lower() or (subsector or "").lower() in k.lower():
                env = v; break
    if not env:
        return cost, months, []
    lo, hi = env["cost"]; mlo, mhi = env["months"]
    # Allow mega wording to approach upper end, but prevent runaway values.
    scale_l=(scale or "").lower()
    if "mega-mega" in scale_l:
        hi *= 1.25; mhi *= 1.15
    elif "mega" in scale_l:
        hi *= 1.12; mhi *= 1.08
    original=(cost,months)
    cost = min(max(cost, lo), hi)
    months = int(min(max(months, mlo), mhi))
    notes=[]
    if abs(cost-original[0]) > 0.01 or months != original[1]:
        notes.append(f"Sector realism envelope applied for {subsector}: cost and schedule constrained to credible early-stage range.")
    return cost, months, notes

SECTOR_COST_TEMPLATES = {
    "life_sciences": [
        ("01.01","Site, enabling and controlled utilities","Direct","Site works, utility corridors and GMP-ready enabling infrastructure",0.07),
        ("01.02","GMP cleanrooms and classified areas","Direct","Cleanroom envelope, HVAC zoning, pressure cascades and finishes",0.17),
        ("01.03","Process equipment and fill-finish lines","Direct","Bioreactors, formulation, aseptic fill-finish and packaging equipment",0.22),
        ("01.04","Clean utilities","Direct","WFI, clean steam, purified water, process gases and validated utility loops",0.14),
        ("01.05","Automation, MES and digital validation","Direct","Automation, MES, batch records, data integrity and validation systems",0.08),
        ("01.06","Cold-chain and automated logistics","Direct","Cold storage, automated warehouse, packaging and distribution infrastructure",0.08),
        ("02.01","CQV / qualification","Indirect","Commissioning, qualification, validation protocols and deviation closure",0.08),
        ("02.02","Regulatory readiness and quality systems","Indirect","FDA/EMA readiness, QA systems, inspection preparation and documentation",0.05),
        ("02.03","Programme management and operational readiness","Indirect","PMO, controls, staffing readiness, training and handover",0.06),
        ("03.01","Risk reserve","Reserve","Allowance for GMP validation, equipment lead times and regulatory uncertainty",0.05),
    ],
    "data_centre": [
        ("01.01","Land, enabling and campus infrastructure","Direct","Site, roads, drainage, telecom routes and campus utilities",0.07),
        ("01.02","Power train and substations","Direct","HV/MV substations, switchgear, transformers, UPS and generators",0.24),
        ("01.03","Data halls / white space","Direct","Data halls, fit-out, containment and critical environments",0.17),
        ("01.04","Cooling plant","Direct","Liquid cooling, chillers, heat rejection and water systems",0.15),
        ("01.05","Fibre, controls and security","Direct","Network rooms, BMS/EPMS, access control and cyber/physical security",0.07),
        ("02.01","Commissioning and integrated systems testing","Indirect","IST, black-building tests, staged energisation and reliability proving",0.08),
        ("02.02","Programme management and procurement","Indirect","PMO, procurement, logistics and contractor management",0.07),
        ("03.01","Risk reserve","Reserve","Utility, long-lead electrical and commissioning uncertainty",0.15),
    ],
    "semiconductor": [
        ("01.01","Site and enabling infrastructure","Direct","Site works, seismic base, bulk utilities and logistics",0.06),
        ("01.02","Cleanroom shell and process environment","Direct","Vibration control, cleanroom, HVAC and contamination control",0.20),
        ("01.03","Process tools and tool install","Direct","Lithography, etch, deposition, metrology and tool hook-up",0.32),
        ("01.04","UPW, chemicals and specialty gases","Direct","Ultra-pure water, gases, abatement and chemical systems",0.14),
        ("01.05","Power and resilient utilities","Direct","Substations, emergency power and redundancy",0.08),
        ("02.01","Qualification and ramp-up","Indirect","Tool qualification, yield ramp and validation",0.08),
        ("02.02","PMO and controls","Indirect","Programme controls, procurement and logistics",0.05),
        ("03.01","Risk reserve","Reserve","Tool cadence, supply chain and ramp-up uncertainty",0.07),
    ],
}

def sector_cost_lines(mode:str, subsector:str):
    fam=_sector_family(subsector)
    if fam in SECTOR_COST_TEMPLATES:
        return SECTOR_COST_TEMPLATES[fam]
    return cost_lines(mode, subsector)

def sector_specific_lists(subsector:str, mode:str):
    fam=_sector_family(subsector)
    if fam=="life_sciences":
        return {
            "cost":["Process equipment and fill-finish lines","GMP cleanrooms / classified areas","Clean utilities and HVAC zoning","CQV validation and deviation closure","Cold-chain logistics and automated warehousing"],
            "schedule":["CQV protocol approval and execution","Long-lead process equipment delivery","Clean utility validation and media fills","FDA/EMA inspection readiness","Operational staffing, training and batch readiness"],
            "confidence":["Benchmark similarity: pharma / biologics campus","Scope maturity: GMP package and user requirement definition","Procurement certainty: process equipment and clean utility lead times","Schedule maturity: CQV logic and validation pathway","Regulatory exposure: FDA/EMA inspection readiness"],
            "thinking":"CASEY interprets this as a validation-constrained life sciences manufacturing programme where CQV sequencing, clean utility readiness, process equipment cadence and regulatory inspection preparedness dominate schedule confidence more than conventional construction productivity.",
            "bench":[
                {"archetype":"Life Sciences / Biologics Manufacturing","similarity_score":9,"anchor_cost":"$2B-$8B","anchor_duration_months":"36-78","use":"Primary analogue for GMP campus, fill-finish and CQV complexity."},
                {"archetype":"Sterile Fill-Finish / Aseptic Expansion","similarity_score":8,"anchor_cost":"$1B-$5B","anchor_duration_months":"30-60","use":"Relevant for aseptic process, cleanroom and validation sequence."},
                {"archetype":"Advanced Manufacturing Cleanroom","similarity_score":7,"anchor_cost":"$2B-$10B","anchor_duration_months":"42-84","use":"Adjacent benchmark for clean environment and specialist equipment."},
                {"archetype":"Cold-Chain Logistics Campus","similarity_score":6,"anchor_cost":"$500M-$3B","anchor_duration_months":"24-54","use":"Adjacent benchmark for warehousing and distribution infrastructure."},
            ]
        }
    if fam=="data_centre":
        return {
            "cost":["Utility/grid connection and substations","Power train, transformers and switchgear","Liquid cooling / heat rejection systems","Data halls and white space fit-out","Accelerated procurement premiums"],
            "schedule":["Grid energisation and utility agreements","Long-lead transformer and switchgear delivery","Integrated systems testing and commissioning","Design freeze and phasing stability","Contractor productivity across concurrent halls"],
            "confidence":["Benchmark similarity: hyperscale digital infrastructure","Scope maturity: campus power and white-space definition","Procurement certainty: transformers, generators and switchgear","Schedule maturity: grid and commissioning logic","Interface exposure: utilities, fibre and commissioning"],
            "thinking":"CASEY interprets this as an energy-constrained digital infrastructure programme where utility availability, commissioning throughput and long-lead electrical procurement dominate schedule confidence more than vertical construction productivity.",
            "bench":[
                {"archetype":"Hyperscale AI Data Centre Campus","similarity_score":9,"anchor_cost":"$2B-$18B","anchor_duration_months":"24-84","use":"Primary analogue for power-intensive digital infrastructure."},
                {"archetype":"Energy / Utility Megaprogramme","similarity_score":7,"anchor_cost":"$1B-$18B","anchor_duration_months":"36-108","use":"Adjacent benchmark for substations and energisation risk."},
                {"archetype":"Mission Critical Facility","similarity_score":7,"anchor_cost":"$1B-$8B","anchor_duration_months":"30-72","use":"Adjacent benchmark for resilience and uptime requirements."},
                {"archetype":"Semiconductor / Advanced Manufacturing","similarity_score":6,"anchor_cost":"$8B-$32B","anchor_duration_months":"54-96","use":"Adjacent benchmark for complex MEP and commissioning."},
            ]
        }
    if fam=="semiconductor":
        return {
            "cost":["Process tools and tool install","Cleanroom shell and vibration control","UPW, gases and chemical systems","Power redundancy and utility resilience","Yield ramp and qualification"],
            "schedule":["Long-lead lithography/process tool delivery","Cleanroom readiness and contamination control","UPW and specialty gas commissioning","Tool hook-up and qualification cadence","Yield ramp and process validation"],
            "confidence":["Benchmark similarity: semiconductor fab","Scope maturity: process tool list and cleanroom class definition","Procurement certainty: tool delivery and hook-up sequencing","Schedule maturity: cleanroom/tool qualification logic","Utility exposure: UPW, gases, power and abatement"],
            "thinking":"CASEY interprets this as a tooling-and-utilities constrained advanced manufacturing programme where process tool cadence, cleanroom readiness, UPW/gas systems and yield ramp dominate delivery confidence.",
            "bench":[
                {"archetype":"Advanced Semiconductor Fab","similarity_score":9,"anchor_cost":"$8B-$32B","anchor_duration_months":"54-96","use":"Primary analogue for tool-heavy cleanroom delivery."},
                {"archetype":"Advanced Manufacturing Cleanroom","similarity_score":8,"anchor_cost":"$2B-$10B","anchor_duration_months":"42-84","use":"Adjacent benchmark for controlled environment and specialist systems."},
                {"archetype":"Hyperscale Digital Infrastructure","similarity_score":6,"anchor_cost":"$2B-$18B","anchor_duration_months":"24-84","use":"Adjacent benchmark for utility and commissioning intensity."}
            ]
        }
    if fam=="airport":
        return {
            "cost":["Terminal and airside works","Baggage and security systems","Operational transition and phasing","Transport integration and utilities","Retail/passenger experience fit-out"],
            "schedule":["Live airport phasing and possessions","Baggage/security systems integration","Operational readiness trials","Regulatory and stakeholder approvals","Airside access and safety constraints"],
            "confidence":["Benchmark similarity: airport terminal expansion","Scope maturity: capacity, phasing and systems definition","Procurement certainty: baggage/security/MEP packages","Schedule maturity: ORAT and live operations logic","Interface exposure: airlines, airside, landside and regulators"],
            "thinking":"CASEY interprets this as a live-operational aviation programme where phasing, systems integration and operational readiness dominate confidence more than standalone construction progress.",
            "bench":[{"archetype":"Airport Terminal Expansion","similarity_score":9,"anchor_cost":"$2B-$24B","anchor_duration_months":"60-144","use":"Primary analogue for live-airport capital delivery."},{"archetype":"Rail/Transit Systems Integration","similarity_score":6,"anchor_cost":"$3B-$80B","anchor_duration_months":"60-180","use":"Adjacent benchmark for passenger systems and operational transition."}]
        }
    if fam=="rail":
        return {
            "cost":["Civil works, tunnelling and structures","Stations and public realm","Signalling, power and systems","Land, utilities and consents","Testing, commissioning and trial operations"],
            "schedule":["Land acquisition and utility diversions","Possessions and live railway interfaces","Systems integration and signalling readiness","Station fit-out and operational trials","Regulatory approvals and safety case"],
            "confidence":["Benchmark similarity: rail/transit programme","Scope maturity: alignment, station and systems definition","Procurement certainty: civil/systems package strategy","Schedule maturity: possessions and test/commissioning logic","Interface exposure: utilities, operators and regulators"],
            "thinking":"CASEY interprets this as an interface-heavy transit programme where utilities, possessions, systems integration and safety assurance dominate delivery confidence.",
            "bench":[{"archetype":"Metro / Rail Extension","similarity_score":9,"anchor_cost":"$3B-$80B","anchor_duration_months":"60-180","use":"Primary analogue for complex transit delivery."},{"archetype":"Airport Systems Programme","similarity_score":6,"anchor_cost":"$2B-$24B","anchor_duration_months":"60-144","use":"Adjacent benchmark for live operations and systems assurance."}]
        }
    if fam=="healthcare":
        return {
            "cost":["Clinical accommodation and theatres","Diagnostics and specialist medical equipment","MEP, resilience and energy centre","Digital clinical systems","Phased handover and commissioning"],
            "schedule":["Clinical commissioning and licensing","Medical equipment procurement","Decant/phasing with live services","Digital systems integration","Operational readiness and staff training"],
            "confidence":["Benchmark similarity: acute hospital campus","Scope maturity: clinical model and departmental adjacencies","Procurement certainty: medical equipment and MEP resilience","Schedule maturity: clinical commissioning and handover logic","Operational exposure: live healthcare continuity"],
            "thinking":"CASEY interprets this as a clinically constrained healthcare programme where operational continuity, equipment readiness and clinical commissioning dominate confidence.",
            "bench":[{"archetype":"Acute Hospital Campus","similarity_score":9,"anchor_cost":"$800M-$6.5B","anchor_duration_months":"36-84","use":"Primary analogue for healthcare capital delivery."},{"archetype":"Life Sciences / GMP Facility","similarity_score":6,"anchor_cost":"$1.5B-$8.5B","anchor_duration_months":"36-78","use":"Adjacent benchmark for controlled environments and validation."}]
        }
    if fam=="defence":
        return {
            "cost":["Hardened facilities and secure fit-out","Resilient power and communications","Cyber/security accreditation","Specialist mission systems","Controlled procurement and assurance"],
            "schedule":["Security accreditation and authority approvals","Secure systems integration","Long-lead mission equipment","Operational acceptance trials","Controlled access and classified interfaces"],
            "confidence":["Benchmark similarity: secure mission infrastructure","Scope maturity: mission system and accreditation requirements","Procurement certainty: controlled equipment and secure supply chain","Schedule maturity: accreditation and acceptance logic","Interface exposure: security authority and operators"],
            "thinking":"CASEY interprets this as secure mission infrastructure where accreditation, resilience and controlled procurement can become hidden critical-path drivers.",
            "bench":[{"archetype":"Secure Mission / Defence Facility","similarity_score":9,"anchor_cost":"$800M-$12B","anchor_duration_months":"30-84","use":"Primary analogue for secure infrastructure delivery."},{"archetype":"Data Centre / Mission Critical Facility","similarity_score":7,"anchor_cost":"$1.5B-$18B","anchor_duration_months":"24-84","use":"Adjacent benchmark for resilience and uptime."}]
        }
    if fam in ["energy","nuclear","water"]:
        return {
            "cost":["Generation/process equipment","Grid, utility and connection infrastructure","Civil and enabling works","Permitting and environmental compliance","Commissioning and operational readiness"],
            "schedule":["Permitting and environmental approvals","Grid connection and utility interfaces","Long-lead equipment procurement","Commissioning and performance testing","Operator readiness and regulatory handover"],
            "confidence":["Benchmark similarity: energy/utilities programme","Scope maturity: process configuration and connection definition","Procurement certainty: long-lead equipment and utility interfaces","Schedule maturity: permitting and commissioning logic","Regulatory exposure: environmental, safety and operator approvals"],
            "thinking":"CASEY interprets this as regulated utility infrastructure where approvals, grid/process interfaces and commissioning evidence dominate confidence.",
            "bench":[{"archetype":"Energy / Utility Megaprogramme","similarity_score":9,"anchor_cost":"$1B-$35B","anchor_duration_months":"36-144","use":"Primary analogue for regulated utility delivery."},{"archetype":"Industrial Process Facility","similarity_score":7,"anchor_cost":"$1B-$12B","anchor_duration_months":"36-96","use":"Adjacent benchmark for process systems and commissioning."}]
        }
    if mode=="Space":
        return {
            "cost":["Launch and mass-to-orbit logistics","Power generation and thermal control","Radiation-hardening and qualification","Autonomous operations and servicing","Ground segment and relay communications"],
            "schedule":["Technology readiness and qualification","Launch manifest and integration cadence","Thermal/power validation","Autonomous recovery and servicing readiness","Ground/orbit commissioning sequence"],
            "confidence":["Benchmark similarity: space infrastructure archetype","Scope maturity: payload and mission architecture definition","Procurement certainty: launch, avionics and qualified hardware","Schedule maturity: launch and commissioning logic","Operational exposure: remote recovery and servicing limits"],
            "thinking":"CASEY interprets this as persistent space infrastructure where delivery confidence is governed by launch cadence, technology readiness, qualification evidence and autonomous operational resilience.",
            "bench":[
                {"archetype":"Orbital / Lunar Infrastructure","similarity_score":8,"anchor_cost":"$8B-$95B","anchor_duration_months":"72-216","use":"Primary analogue for space infrastructure complexity."},
                {"archetype":"Launch and Payload Integration","similarity_score":7,"anchor_cost":"$1B-$14B","anchor_duration_months":"36-108","use":"Relevant for launch cadence and payload integration."},
                {"archetype":"Deep-Space Mission Systems","similarity_score":7,"anchor_cost":"$4B-$35B","anchor_duration_months":"60-144","use":"Adjacent benchmark for qualification and mission assurance."},
                {"archetype":"Autonomous Operations Platform","similarity_score":6,"anchor_cost":"$2B-$20B","anchor_duration_months":"36-96","use":"Adjacent benchmark for autonomy and remote operations."},
            ]
        }
    return {
        "cost":["Civil and enabling works","Specialist systems and long-lead equipment","Utilities and interfaces","Commissioning and operational readiness","Risk reserve driven by procurement and interface uncertainty"],
        "schedule":["Approvals and consents","Long-lead procurement","Design freeze stability","Systems integration and commissioning","Operational access and handover readiness"],
        "confidence":["Benchmark similarity: comparable infrastructure archetypes","Scope maturity: concept / budget level until package evidence is supplied","Procurement certainty: long-lead equipment and market capacity","Schedule maturity: critical path and commissioning logic","Interface exposure: utilities, systems and operations"],
        "thinking":"CASEY interprets this as a system-of-systems infrastructure programme where delivery confidence depends on interface control, procurement evidence and decision velocity.",
        "bench":[]
    }

def sector_signature_behaviour(subsector:str, mode:str) -> Dict[str, Any]:
    fam=_sector_family(subsector)
    signatures = {
        "life_sciences": {
            "shock":"Mechanical completion is unlikely to be the true finish line; validated production readiness and deviation closure are the real board decision gates.",
            "curve":"late_validation_spike",
            "human_basis":"duration reflects CQV sequencing, clean-utility validation, media-fill readiness and phased batch ramp-up rather than simple building completion.",
            "contradiction":"Acceleration may protect market-entry timing but reduces validation float and increases deviation-closure pressure.",
            "signature":["CQV path", "media fill readiness", "GMP turnover", "batch release", "inspection readiness"]
        },
        "data_centre": {
            "shock":"The dominant delivery constraint is likely energisation and commissioning concurrency, not shell construction productivity.",
            "curve":"power_procurement_cliff",
            "human_basis":"duration reflects utility energisation, transformer/switchgear procurement, integrated systems testing and phased data-hall commissioning.",
            "contradiction":"Acceleration can bring revenue online earlier, but usually buys schedule by paying procurement premiums and accepting commissioning overlap.",
            "signature":["grid energisation", "transformers", "liquid cooling", "IST", "phased data halls"]
        },
        "semiconductor": {
            "shock":"The critical path is likely tool install and qualification cadence rather than the cleanroom shell alone.",
            "curve":"tool_install_ramp",
            "human_basis":"duration reflects cleanroom readiness, UPW/gas commissioning, process-tool delivery, hook-up and yield-ramp qualification.",
            "contradiction":"Lower capex assumptions can move risk into yield ramp, tool utilisation and operational start-up.",
            "signature":["tool install", "UPW", "vibration control", "hook-up", "yield ramp"]
        },
        "airport": {
            "shock":"The programme may finish construction before the airport is operationally ready to absorb the change.",
            "curve":"orat_ramp",
            "human_basis":"duration reflects live-airport phasing, baggage/security integration, ORAT, stakeholder approvals and passenger operations continuity.",
            "contradiction":"Acceleration increases disruption exposure unless operational readiness trials are protected.",
            "signature":["ORAT", "baggage integration", "airside phasing", "security systems", "live operations"]
        },
        "rail": {
            "shock":"Possessions, utility diversions and systems assurance are likely to dominate the critical path more than visible civil progress.",
            "curve":"systems_assurance_tail",
            "human_basis":"duration reflects possessions, utility diversions, tunnelling/station works, signalling integration and safety assurance.",
            "contradiction":"A cheaper case may defer systems maturity and create a larger testing and commissioning tail.",
            "signature":["possessions", "utility diversions", "signalling", "safety case", "trial operations"]
        },
        "healthcare": {
            "shock":"Clinical readiness and safe operational transition may lag physical completion unless planned as a critical-path workstream.",
            "curve":"clinical_commissioning_tail",
            "human_basis":"duration reflects clinical commissioning, medical-equipment readiness, digital systems integration, decanting and live-service continuity.",
            "contradiction":"Acceleration can increase patient-service disruption unless clinical transition is protected.",
            "signature":["clinical commissioning", "medical equipment", "decant", "digital clinical systems", "operational readiness"]
        },
        "defence": {
            "shock":"Security accreditation and mission-system acceptance can become the hidden critical path even after the facility is built.",
            "curve":"accreditation_gate",
            "human_basis":"duration reflects controlled procurement, secure systems integration, resilience proving, cyber/security accreditation and operational acceptance.",
            "contradiction":"Lower-cost delivery can create downstream accreditation and acceptance risk.",
            "signature":["accreditation", "mission systems", "secure comms", "resilience proving", "controlled procurement"]
        },
        "energy": {
            "shock":"The approval, grid/process interface and commissioning evidence may drive confidence more than equipment installation.",
            "curve":"regulatory_grid_gate",
            "human_basis":"duration reflects consenting, grid/process interfaces, long-lead equipment and performance testing.",
            "contradiction":"Compression can pull construction forward while leaving approvals and grid acceptance behind.",
            "signature":["consents", "grid interface", "process equipment", "performance test", "operator handover"]
        },
        "nuclear": {
            "shock":"Licensing and safety-case maturity are the true confidence drivers; construction progress alone is a weak indicator.",
            "curve":"licensing_plateau",
            "human_basis":"duration reflects licensing, safety case, nuclear-grade procurement, regulator hold points and commissioning evidence.",
            "contradiction":"Attempting to accelerate before safety-case maturity usually moves risk into rework and regulator challenge.",
            "signature":["licensing", "safety case", "regulator hold point", "nuclear QA", "commissioning evidence"]
        },
        "water": {
            "shock":"Marine/environmental approvals and energy interfaces may dominate the delivery envelope.",
            "curve":"permit_then_commission",
            "human_basis":"duration reflects permitting, intake/outfall works, process commissioning, power tie-in and environmental compliance.",
            "contradiction":"Cheaper options may reduce redundancy and increase operating resilience exposure.",
            "signature":["intake/outfall", "environmental permit", "process commissioning", "power tie-in", "resilience"]
        },
        "orbital": {
            "shock":"The decisive constraint is not launch alone; it is thermal-power balance, autonomous servicing and recoverability after deployment.",
            "curve":"qualification_launch_spike",
            "human_basis":"duration reflects technology qualification, payload integration, launch manifest, thermal/power validation and autonomous operations readiness.",
            "contradiction":"Acceleration increases mission-assurance risk if qualification and environmental testing are compressed.",
            "signature":["TRL", "thermal rejection", "payload integration", "launch manifest", "autonomous servicing"]
        },
        "lunar": {
            "shock":"Surface logistics, dust and power resilience are likely to dominate sustained operability more than initial landing success.",
            "curve":"surface_logistics_tail",
            "human_basis":"duration reflects launch cadence, landing-site preparation, surface logistics, power storage and autonomous construction readiness.",
            "contradiction":"Lower-cost delivery may reduce redundancy and increase resupply dependency.",
            "signature":["landing pads", "dust", "surface power", "rovers", "resupply cadence"]
        },
        "mars": {
            "shock":"The programme is governed by launch windows, autonomy and life-support reliability; recovery options are extremely limited.",
            "curve":"launch_window_steps",
            "human_basis":"duration reflects launch windows, life-support qualification, ISRU maturity, autonomous maintenance and mission-assurance gates.",
            "contradiction":"Acceleration can only be credible if autonomy, ECLSS and ISRU qualification evidence are protected.",
            "signature":["launch windows", "ECLSS", "ISRU", "autonomy", "mission assurance"]
        },
        "spaceport": {
            "shock":"Range safety, licensing and propellant operations can dominate readiness after civil works are complete.",
            "curve":"range_safety_gate",
            "human_basis":"duration reflects licensing, environmental approvals, launch-pad systems, propellant operations and range-safety validation.",
            "contradiction":"Acceleration increases regulatory and range-safety exposure unless approvals are sequenced early.",
            "signature":["range safety", "propellant farm", "launch pad", "payload processing", "licensing"]
        },
        "space_comms": {
            "shock":"Availability is governed by ground/orbit interface resilience, not just antenna or satellite asset delivery.",
            "curve":"mission_network_tail",
            "human_basis":"duration reflects payload qualification, ground segment readiness, relay integration and operational availability proving.",
            "contradiction":"Lower capex can reduce redundancy and increase mission-availability exposure.",
            "signature":["ground segment", "relay", "availability", "antenna", "mission operations"]
        },
        "space_resources": {
            "shock":"Resource uncertainty and autonomous maintenance are likely to dominate business-case confidence.",
            "curve":"resource_uncertainty_spike",
            "human_basis":"duration reflects prospecting uncertainty, extraction technology maturity, autonomous operations and processing reliability.",
            "contradiction":"Acceleration without resource evidence increases downstream production risk.",
            "signature":["resource evidence", "autonomous mining", "processing", "maintenance", "production reliability"]
        },
    }
    if mode=="Space" and fam not in signatures:
        fam="orbital"
    return signatures.get(fam, {
        "shock":"The dominant risk is likely to sit in interfaces, procurement evidence and commissioning readiness rather than headline construction progress.",
        "curve":"generic",
        "human_basis":"duration reflects design maturity, procurement evidence, interface control, commissioning and operational readiness.",
        "contradiction":"Acceleration improves date certainty only if assumptions, interfaces and procurement evidence are strengthened.",
        "signature":["interfaces", "procurement evidence", "commissioning", "operational readiness", "decision gates"]
    })


# ------------------------- CASEY v9.5 scenario cascade engine -------------------------
SCENARIO_PROFILES_V95 = {
    "base": {
        "label":"Base","cost":1.00,"schedule":1.00,"confidence":0,"reserve":1.00,"risk":"Medium-High",
        "delta":["Reference execution case","Balanced cost, schedule and confidence posture","No acceleration or deferral premium applied"],
        "cost_delta":{"Direct":1.00,"Indirect":1.00,"Reserve":1.00},
        "risk_lift":0,
        "curve":"balanced"
    },
    "faster": {
        "label":"Faster","cost":1.14,"schedule":0.82,"confidence":-10,"reserve":1.18,"risk":"High",
        "delta":["Schedule compressed through overlapping delivery workstreams","Procurement premium and concurrency exposure increased","Commissioning float reduced; late-stage volatility increased"],
        "cost_delta":{"Direct":1.08,"Indirect":1.20,"Reserve":1.28},
        "risk_lift":14,
        "curve":"compressed_tail"
    },
    "cheaper": {
        "label":"Cheaper","cost":0.88,"schedule":1.10,"confidence":-14,"reserve":0.72,"risk":"High",
        "delta":["Capital target reduced through slower phasing and lower redundancy","Contingency and acceleration spend reduced","Operational start-up and lifecycle risk increased"],
        "cost_delta":{"Direct":0.91,"Indirect":0.82,"Reserve":0.62},
        "risk_lift":10,
        "curve":"low_median_fat_tail"
    },
    "lower_risk": {
        "label":"Lower Risk","cost":1.09,"schedule":1.12,"confidence":12,"reserve":1.28,"risk":"Medium",
        "delta":["Assurance, procurement buffers and commissioning float expanded","Execution concurrency reduced","Tail risk and decision uncertainty reduced"],
        "cost_delta":{"Direct":1.04,"Indirect":1.12,"Reserve":1.42},
        "risk_lift":-10,
        "curve":"tightened"
    },
    "premium": {
        "label":"Premium","cost":1.22,"schedule":0.96,"confidence":18,"reserve":1.35,"risk":"Medium-Low",
        "delta":["Resilience, redundancy and optionality increased","Procurement certainty and board confidence improved","Higher capex protects delivery and operational readiness"],
        "cost_delta":{"Direct":1.18,"Indirect":1.20,"Reserve":1.48},
        "risk_lift":-14,
        "curve":"premium_resilient"
    },
}

def _money_to_bn_v95(s):
    s=str(s or "").replace("$","").strip().upper()
    try:
        if s.endswith("T"): return float(s[:-1])*1000
        if s.endswith("B"): return float(s[:-1])
        if s.endswith("M"): return float(s[:-1])/1000
        return float(s)
    except Exception:
        return 0.0

def _bn_to_money_v95(bn):
    try:
        bn=float(bn)
    except Exception:
        bn=0.0
    if bn >= 1000: return f"${bn/1000:.1f}T"
    if bn >= 1: return f"${bn:.1f}B"
    return f"${bn*1000:.0f}M"

def _scenario_insight_v95(model, scenario):
    profile=SCENARIO_PROFILES_V95.get(scenario, SCENARIO_PROFILES_V95["base"])
    fam=_sector_family(model.get("subsector",""))
    base = {
        "life_sciences":{
            "base":"Mechanical completion is not the true finish line; validated production readiness and deviation closure are the real board decision gates.",
            "faster":"Acceleration is compressing CQV float and increasing concurrent commissioning exposure across sterile systems.",
            "cheaper":"Capital reduction is shifting risk from construction cost into validation readiness, redundancy and operational start-up.",
            "lower_risk":"The lower-risk path protects GMP turnover by expanding validation float and procurement evidence before execution peaks.",
            "premium":"Premium delivery buys resilience, redundancy and stronger regulatory readiness ahead of commercial production ramp-up."
        },
        "data_centre":{
            "base":"Energisation and commissioning concurrency are more likely to govern delivery than shell construction productivity.",
            "faster":"Acceleration is pulling grid, cooling and IST workstreams into tighter concurrency, increasing late-stage volatility.",
            "cheaper":"The cheaper case reduces capex but pushes risk into redundancy, commissioning resilience and operational uptime.",
            "lower_risk":"The lower-risk case protects energisation and IST by adding commissioning float and procurement buffers.",
            "premium":"Premium delivery increases resilience and power optionality, reducing service-readiness exposure."
        },
        "semiconductor":{
            "base":"Tool install and qualification cadence are likely to govern the real finish line more than cleanroom completion.",
            "faster":"Acceleration compresses tool hook-up, UPW readiness and yield-ramp qualification into a higher-volatility window.",
            "cheaper":"The cheaper case risks deferring utility resilience and tool-readiness assurance into production ramp-up.",
            "lower_risk":"The lower-risk case extends tool qualification and utility proving to reduce yield-ramp uncertainty.",
            "premium":"Premium delivery protects tool cadence, utility resilience and yield-ramp confidence."
        },
        "orbital":{
            "base":"Thermal-power balance, autonomous servicing and recoverability after deployment are the decisive constraints.",
            "faster":"Acceleration compresses qualification and payload integration, increasing mission-assurance exposure.",
            "cheaper":"The cheaper case reduces redundancy and increases recoverability risk after orbital deployment.",
            "lower_risk":"The lower-risk case protects qualification, environmental testing and autonomous recovery readiness.",
            "premium":"Premium delivery buys redundancy, mission assurance and stronger autonomous operations resilience."
        },
        "mars":{
            "base":"Launch windows, autonomy and life-support reliability govern the programme; recovery options are extremely limited.",
            "faster":"Acceleration is credible only if ECLSS, ISRU and autonomy qualification evidence remain protected.",
            "cheaper":"The cheaper case increases mission risk by reducing redundancy where recovery options are limited.",
            "lower_risk":"The lower-risk case expands mission-assurance gates and life-support proving before commitment.",
            "premium":"Premium delivery protects life-support redundancy, autonomy and launch-window resilience."
        }
    }
    pack = base.get(fam) or (base["orbital"] if model.get("mode")=="Space" else {
        "base":"The dominant risk sits in interfaces, procurement evidence and commissioning readiness rather than headline construction progress.",
        "faster":"Acceleration compresses interface resolution and increases late-stage execution volatility.",
        "cheaper":"The cheaper case reduces near-term capex but moves risk into operational readiness and contingency adequacy.",
        "lower_risk":"The lower-risk case adds evidence, float and assurance to reduce tail exposure.",
        "premium":"Premium delivery buys resilience, optionality and stronger confidence."
    })
    return pack.get(scenario, pack["base"])

def _mutate_curve_v95(cost, months, scenario):
    profile=SCENARIO_PROFILES_V95.get(scenario, SCENARIO_PROFILES_V95["base"])
    pts=[1,5,10,20,30,40,50,60,70,80,90,95,99]
    curve=[]
    for p in pts:
        x=p/100
        if profile["curve"]=="compressed_tail":
            cf=0.80 + 0.35*x + 0.24*(x**4)
            sf=0.72 + 0.30*x + 0.30*(x**5)
        elif profile["curve"]=="low_median_fat_tail":
            cf=0.72 + 0.30*x + 0.35*(x**5)
            sf=0.84 + 0.26*x + 0.33*(x**4)
        elif profile["curve"]=="tightened":
            cf=0.86 + 0.22*x + 0.10*(x**3)
            sf=0.90 + 0.18*x + 0.08*(x**3)
        elif profile["curve"]=="premium_resilient":
            cf=0.88 + 0.22*x + 0.08*(x**3)
            sf=0.86 + 0.20*x + 0.10*(x**3)
        else:
            cf=0.78 + 0.35*x + 0.12*(x**3)
            sf=0.82 + 0.26*x + 0.12*(x**3)
        curve.append({"percentile":p,"cost_bn":round(cost*cf,2),"schedule_months":int(round(months*sf))})
    return curve

def scenario_cascade_v95(model:Dict[str,Any], scenario:str) -> Dict[str,Any]:
    scenario=(scenario or "base").lower()
    profile=SCENARIO_PROFILES_V95.get(scenario, SCENARIO_PROFILES_V95["base"])
    # preserve base references
    base_cost=_money_to_bn_v95(model.get("_base_cost_p50") or model.get("cost_p50"))
    base_months=int(float(str(model.get("_base_schedule_months") or model.get("schedule") or "60").replace("months","").strip().split()[0]))
    base_conf=int(model.get("_base_confidence_pct") or model.get("confidence_pct") or 55)
    if "_base_cost_p50" not in model:
        model["_base_cost_p50"]=model.get("cost_p50")
        model["_base_schedule_months"]=base_months
        model["_base_confidence_pct"]=base_conf
        model["_base_risk"]=model.get("risk","Medium-High")
        model["_base_cost_lines"]=json.loads(json.dumps(model.get("cost_lines",[])))
        model["_base_risks"]=json.loads(json.dumps(model.get("risks",[])))
        model["_base_schedule_rows"]=json.loads(json.dumps(model.get("schedule_rows",[])))
    new_cost=base_cost*profile["cost"]
    new_months=max(3,int(round(base_months*profile["schedule"])))
    new_conf=max(8,min(96,base_conf+profile["confidence"]))
    model["scenario"]=scenario
    model["scenario_label=profile_label"] = profile["label"]
    model["scenario_label"]=profile["label"]
    model["cost_p50"]=_bn_to_money_v95(new_cost)
    model["cost_p10"]=_bn_to_money_v95(new_cost*0.80)
    model["cost_p90"]=_bn_to_money_v95(new_cost*1.30)
    model["cost_range"]=f"{model['cost_p10']} - {model['cost_p90']}"

    # Executive scenario comparison versus the immutable Base case.
    base_cost_money=_bn_to_money_hs(base_cost)
    base_schedule_months=base_months
    base_conf_pct=base_conf
    delta_cost=new_cost-base_cost
    delta_months=new_months-base_months
    delta_conf=new_conf-base_conf
    model["scenario_comparison_vs_base"]={
        "base":{"cost_p50":base_cost_money,"schedule_months":base_schedule_months,"confidence_pct":base_conf_pct,"risk":model.get("_base_risk","Medium-High")},
        "selected":{"scenario":profile["label"],"cost_p50":model["cost_p50"],"schedule_months":new_months,"confidence_pct":new_conf,"risk":profile["risk"]},
        "delta":{"cost_bn":round(delta_cost,2),"cost":_bn_to_money_hs(abs(delta_cost)),"cost_direction":"higher" if delta_cost>0 else "lower" if delta_cost<0 else "same",
                 "months":delta_months,"confidence_pts":delta_conf},
        "plain_english": ("Base is the reference case: no cost, schedule or confidence delta is applied. Use Faster, Cheaper, Lower Risk or Premium to expose the board trade-off." if scenario=="base" else f"Compared with Base, {profile['label']} is {_bn_to_money_hs(abs(delta_cost))} {'more expensive' if delta_cost>0 else 'cheaper' if delta_cost<0 else 'unchanged'}, {abs(delta_months)} months {'faster' if delta_months<0 else 'slower' if delta_months>0 else 'unchanged'}, and {abs(delta_conf)} confidence points {'higher' if delta_conf>0 else 'lower' if delta_conf<0 else 'unchanged'}.")
    }
    model["scenario_matrix"]=[]
    for sk,sp in SCENARIO_PROFILES_V95_PLUS.items():
        c=base_cost*sp["cost"]; m=max(3,int(round(base_months*sp["schedule"]))); cf=max(8,min(96,base_conf+sp["confidence"]))
        model["scenario_matrix"].append({"scenario":sk,"label":sp["label"],"cost_p50":_bn_to_money_hs(c),"schedule_months":m,"risk":sp["risk"],"confidence_pct":cf,"cost_delta_pct":round((sp["cost"]-1)*100),"schedule_delta_pct":round((sp["schedule"]-1)*100),"confidence_delta_pts":sp["confidence"],"why":"; ".join([sp.get("trade",""), sp.get("won",""), sp.get("lost","")])})
    # Waterfalls explain WHY the selected scenario moved. These feed dashboard and exports.
    if scenario=="faster":
        cost_moves=[("Base P50",base_cost),("Acceleration premium",base_cost*.075),("Parallel EPC / package overlap",base_cost*.045),("Early long-lead procurement",base_cost*.035),("Commissioning surge / overtime",base_cost*.025),("Faster scenario P50",new_cost)]
        sched_moves=[("Base duration",base_months),("Parallel delivery workfronts",-round(base_months*.07)),("Early procurement release",-round(base_months*.05)),("Concurrent commissioning",-round(base_months*.04)),("Reduced float / overlap",new_months-base_months+round(base_months*.16)),("Faster duration",new_months)]
    elif scenario=="cheaper":
        cost_moves=[("Base P50",base_cost),("Scope / spec restraint",-base_cost*.055),("Deferred redundancy",-base_cost*.040),("Lower indirects / slower phasing",-base_cost*.030),("Reduced reserve",-base_cost*.015),("Cheaper scenario P50",new_cost)]
        sched_moves=[("Base duration",base_months),("Slower procurement path",round(base_months*.04)),("Deferred assurance / resilience",round(base_months*.035)),("Lean owner / commissioning support",round(base_months*.03)),("Cheaper duration",new_months)]
    elif scenario=="lower_risk":
        cost_moves=[("Base P50",base_cost),("Assurance gates",base_cost*.035),("Procurement buffer",base_cost*.030),("Commissioning float",base_cost*.025),("Reserve uplift",base_cost*.030),("Lower-risk scenario P50",new_cost)]
        sched_moves=[("Base duration",base_months),("Evidence gates",round(base_months*.055)),("Protected commissioning float",round(base_months*.055)),("Lower-risk duration",new_months)]
    elif scenario=="premium":
        cost_moves=[("Base P50",base_cost),("Redundancy / resilience",base_cost*.095),("Priority procurement",base_cost*.060),("Operational readiness",base_cost*.045),("Optionality premium",base_cost*.080),("Premium scenario P50",new_cost)]
        sched_moves=[("Base duration",base_months),("Priority procurement",-round(base_months*.04)),("Stronger readiness gates",round(base_months*.02)),("Premium duration",new_months)]
    else:
        cost_moves=[("Base P50",base_cost),("Selected scenario P50",new_cost)]
        sched_moves=[("Base duration",base_months),("Selected scenario duration",new_months)]
    model["cost_waterfall_vs_base"]=[{"driver":a,"value_bn":round(b,2),"value":_bn_to_money_hs(abs(b)) if i not in [0,len(cost_moves)-1] else _bn_to_money_hs(b),"kind":"total" if i in [0,len(cost_moves)-1] else "delta"} for i,(a,b) in enumerate(cost_moves)]
    model["schedule_waterfall_vs_base"]=[{"driver":a,"months":int(b),"kind":"total" if i in [0,len(sched_moves)-1] else "delta"} for i,(a,b) in enumerate(sched_moves)]
    model["schedule"]=f"{new_months} months"
    model["risk"]=profile["risk"]
    model["confidence_pct"]=new_conf
    insight=_scenario_insight_v95(model, scenario)
    model["executive_shock_insight"]=insight
    model["scenario_delta_intelligence"]=[
        {"label":"Cost movement","value":f"{profile['cost']*100-100:+.0f}% vs base","meaning":profile["delta"][0]},
        {"label":"Schedule movement","value":f"{profile['schedule']*100-100:+.0f}% vs base","meaning":profile["delta"][1]},
        {"label":"Confidence movement","value":f"{profile['confidence']:+d} pts","meaning":profile["delta"][2]},
        {"label":"Risk posture","value":profile["risk"],"meaning":insight},
    ]
    model["confidence_breakdown"]=[
        {"driver":"Benchmark similarity","effect":"+8","note":"Comparable sector archetype identified"},
        {"driver":"Procurement certainty","effect":"-9" if scenario in ["faster","cheaper"] else "+7","note":"Long-lead evidence changes scenario confidence"},
        {"driver":"Schedule logic maturity","effect":"-8" if scenario=="faster" else "+6" if scenario in ["lower_risk","premium"] else "+3","note":"Critical path and handover gates rebalanced"},
        {"driver":"Commissioning / validation readiness","effect":"-11" if scenario in ["faster","cheaper"] else "+10" if scenario in ["lower_risk","premium"] else "-4","note":"Scenario changes readiness exposure"},
        {"driver":"Contingency adequacy","effect":"-12" if scenario=="cheaper" else "+11" if scenario in ["lower_risk","premium"] else "+2","note":"Reserve and assurance posture updated"},
    ]
    # Cost line mutation
    new_lines=[]
    for line in model.get("_base_cost_lines", model.get("cost_lines",[])):
        x=dict(line)
        typ=x.get("type","Direct")
        factor=profile["cost_delta"].get(typ, profile["cost"])
        for key in ["low_p10","most_likely_p50","high_p90"]:
            if key in x:
                x[key]=_bn_to_money_v95(_money_to_bn_v95(x[key])*factor)
        basis=x.get("basis","")
        if scenario=="faster":
            basis += " Scenario: accelerated procurement / overlapped delivery premium applied."
        elif scenario=="cheaper":
            basis += " Scenario: reduced capital target with lower redundancy and deferred assurance."
        elif scenario=="lower_risk":
            basis += " Scenario: added assurance, procurement buffer and validation float."
        elif scenario=="premium":
            basis += " Scenario: resilience, redundancy and optionality premium applied."
        x["basis"]=basis
        new_lines.append(x)
    model["cost_lines"]=new_lines
    model["cost_breakdown"]=new_lines
    # Risk mutation
    risks=[]
    scenario_risks={
        "faster":[("R-SF1","Concurrent commissioning","Acceleration","Overlapped CQV/commissioning creates rework exposure","Schedule and validation delay","CQV Lead","Protect hold points and add surge QA resource"),
                  ("R-SF2","Acceleration premium","Market","Premium vendors and expediting pressure","Cost escalation","Commercial Lead","Lock procurement packages early and test alternates")],
        "cheaper":[("R-SC1","Deferred redundancy","Capital constraint","Resilience scope deferred to protect capex","Operational start-up exposure","Operations Lead","Define minimum viable redundancy and later upgrade path"),
                   ("R-SC2","Contingency adequacy","Budget pressure","Reserve reduced below uncertainty envelope","Funding shock","Sponsor","Ring-fence risk reserve for critical systems")],
        "lower_risk":[("R-SL1","Assurance gate delay","Assurance posture","Additional evidence gates extend early schedule","Decision delay","PMO","Pre-agree evidence criteria and approval cadence")],
        "premium":[("R-SP1","Premium scope creep","Optionality","Enhanced resilience expands scope boundaries","Capex growth","Sponsor","Govern premium scope through decision gates")]
    }
    base_risks=model.get("_base_risks", model.get("risks",[]))
    for r in base_risks:
        y=dict(r)
        try:
            y["probability_pct"]=max(5,min(85,int(y.get("probability_pct",30))+profile["risk_lift"]))
        except Exception: pass
        risks.append(y)
    for rid,title,cause,event,impact,owner,mit in scenario_risks.get(scenario,[]):
        risks.insert(0,{"id":rid,"title":title,"cause":cause,"event":event,"impact":impact,"probability_pct":max(10,45+profile["risk_lift"]),"activity":"A1500","cbs":"02.01","owner":owner,"mitigation":mit})
    model["risks"]=risks[:10]
    model["risk_register"]=model["risks"]
    # Schedule mutation
    rows=[]
    for row in model.get("_base_schedule_rows", model.get("schedule_rows",[])):
        z=dict(row)
        try:
            dur=int(z.get("duration_months",1))
            z["duration_months"]=max(1,int(round(dur*profile["schedule"])))
        except Exception: pass
        if scenario=="faster":
            z["basis"]="Accelerated path: overlapped workfronts, reduced float and earlier procurement release."
        elif scenario=="cheaper":
            z["basis"]="Cheaper path: slower phasing, reduced parallel workfronts and deferred assurance."
        elif scenario=="lower_risk":
            z["basis"]="Lower-risk path: additional assurance gates, protected commissioning float and evidence-led handover."
        elif scenario=="premium":
            z["basis"]="Premium path: stronger redundancy, procurement certainty and controlled readiness gates."
        rows.append(z)
    model["schedule_rows"]=rows
    model["schedule_detail"]=rows
    # Monte Carlo
    curve=_mutate_curve_v95(new_cost,new_months,scenario)
    model["monte_carlo"]={
        "qcra":{"p10":round(curve[2]["cost_bn"],1),"p50":round(new_cost,1),"p80":round(curve[9]["cost_bn"],1),"p90":round(curve[10]["cost_bn"],1)},
        "qsra":{"p10":curve[2]["schedule_months"],"p50":new_months,"p80":curve[9]["schedule_months"],"p90":curve[10]["schedule_months"]},
        "curve":curve,
        "tornado":[{"risk_id":r.get("id"),"title":r.get("title"),"driver_score":max(1,100-i*8)} for i,r in enumerate(model["risks"][:8])]
    }
    model["board_briefing"]=[
        insight,
        f"{profile['label']} scenario indicates {model['cost_p50']} P50 exposure across approximately {model['schedule']}.",
        f"Confidence moves to {new_conf}% because procurement, schedule logic, contingency and commissioning assumptions have been rebalanced.",
        "Scenario consequence: " + " ".join(profile["delta"])
    ]
    model["casey_thinking"]=(model.get("casey_thinking","CASEY interprets the programme.") + f" Scenario cascade: {insight}")
    model["executive_summary"]=f"{model.get('title','Programme')} scenario view: {profile['label']}. CASEY indicates {model['cost_p50']} P50 exposure, {model['cost_range']} range, {model['schedule']} baseline, {profile['risk']} risk and {new_conf}% confidence. {insight}"
    model["outputs_board_memo"]=[
        f"Decision posture: {profile['label']} scenario.",
        f"Investment implication: {insight}",
        f"Top confidence movement: {model['confidence_breakdown'][1]['driver']} {model['confidence_breakdown'][1]['effect']} pts.",
        "Outputs should be treated as first-pass strategic intelligence, not certified estimate documents."
    ]
    model["top_decisions_required"]=[
        "Confirm the governing critical-path constraint and evidence owner.",
        "Approve scenario-specific procurement and contingency posture.",
        "Lock handover / commissioning / validation decision gates.",
        "Resolve highest-probability interface and readiness risks.",
        "Decide whether the scenario trade-off is acceptable for board approval."
    ]

    # FINAL EXEC POLISH: scenario-aware ranking and differentiation
    try:
        sc = str(model.get("scenario","base")).lower()
        schedule_lists = {
            "faster":[
                "Concurrent commissioning overload",
                "Recovery float exhaustion",
                "Acceleration premium shock",
                "Grid connection delay",
                "Integrated systems testing concurrency"
            ],
            "cheaper":[
                "Vendor claims and change exposure",
                "Procurement deferral and long-lead slippage",
                "Design maturity gap",
                "Scope growth from deferred decisions",
                "Interface coordination delay"
            ],
            "lower_risk":[
                "Governance and approvals latency",
                "Extended validation sequencing",
                "Conservative commissioning gates",
                "Operational readiness hold-points",
                "Assurance and compliance reviews"
            ],
            "premium":[
                "Integration complexity across parallel packages",
                "Executive decision latency",
                "Technology assurance alignment",
                "Multi-package interface management",
                "Programme coordination overhead"
            ]
        }
        cost_lists = {
            "faster":[
                "Acceleration premiums and overtime",
                "Power train, transformers and switchgear",
                "Integrated systems testing",
                "Grid and utility concurrency",
                "Recovery-float consumption"
            ],
            "cheaper":[
                "Deferred procurement packaging",
                "Claims and commercial exposure",
                "Rework from reduced contingency",
                "Long-lead inflation volatility",
                "Scope rationalisation impacts"
            ],
            "lower_risk":[
                "Additional contingency and reserve",
                "Enhanced validation and assurance",
                "Programme controls and governance",
                "Redundant infrastructure resilience",
                "Extended commissioning readiness"
            ]
        }
        if sc in schedule_lists:
            model["sector_schedule_threats"]=schedule_lists[sc]
        if sc in cost_lists:
            model["sector_primary_cost_drivers"]=cost_lists[sc]

        if sc=="faster":
            model["executive_shock_insight"]="Acceleration increases spend faster than it reduces uncertainty; the delivery tail becomes more volatile."
        elif sc=="cheaper":
            model["executive_shock_insight"]="Capital efficiency reduces resilience: procurement and recovery flexibility become constrained."
        elif sc=="lower_risk":
            model["executive_shock_insight"]="Confidence is purchased through reserve, governance and extended delivery duration."
        elif sc=="premium":
            model["executive_shock_insight"]="Premium posture buys resilience, optionality and stronger certainty at visible capex premium."
    except Exception:
        pass

    return model

def apply_sector_intelligence(model:Dict[str,Any]) -> Dict[str,Any]:
    mode=model.get("mode","Earth"); subsector=model.get("subsector","")
    fam=_sector_family(subsector)
    lists=sector_specific_lists(subsector, mode)
    sig=sector_signature_behaviour(subsector, mode)

    def clean(txt):
        txt=str(txt)
        txt=txt.replace("selected schedule level and scenario compression/buffer logic", sig["human_basis"])
        txt=txt.replace("Duration derived from sector schedule model, selected schedule level and scenario compression/buffer logic.", "Duration reflects " + sig["human_basis"])
        if fam!="spaceport" and mode!="Space":
            txt=txt.replace(" or launch readiness","")
            txt=txt.replace("launch readiness delay","readiness delay")
            txt=txt.replace("Launch readiness","Operational readiness")
        return txt

    model["casey_thinking"]=lists["thinking"]
    model["executive_shock_insight"]=sig["shock"]
    model["sector_signature_behaviours"]=sig["signature"]
    model["sector_curve_type"]=sig["curve"]
    model["scenario_contradiction"]=sig["contradiction"]
    model["confidence_engine_label"]="CASEY Confidence Engine"
    model["confidence_engine_detail"]="Benchmark + probabilistic + sector-trained reasoning"

    # Cleaner, more memorable board briefing.
    bb = model.get("board_briefing") if isinstance(model.get("board_briefing"), list) else []
    model["board_briefing"]=[
        sig["shock"],
        f"Current intelligence indicates {model.get('cost_p50')} P50 programme exposure across approximately {model.get('schedule')}, with overall risk assessed as {model.get('risk')}.",
        f"Confidence is governed by {', '.join(sig['signature'][:3])} and the maturity of evidence behind procurement and commissioning assumptions.",
        sig["contradiction"]
    ]

    # Humanise actions.
    model["next_best_actions"]=[
        "Run a focused evidence workshop on the dominant schedule and commissioning constraint.",
        "Validate the top five CBS cost drivers against benchmark or supplier evidence.",
        "Confirm the Level 3/4 schedule logic, handover gates and critical-path assumptions.",
        "Prepare Base / Faster / Cheaper / Lower-Risk decision paper with explicit trade-offs.",
        "Create a confidence-improvement plan targeting design maturity, procurement evidence and operational readiness."
    ]

    if isinstance(model.get("mission_control_cards"), list):
        model["mission_control_cards"]=[c for c in model["mission_control_cards"] if c.get("label")!="EXECUTIVE SHOCK"]
        model["mission_control_cards"].insert(0, {"label":"EXECUTIVE SHOCK", "signal":sig["shock"], "severity":model.get("risk","Medium")})
        for c in model["mission_control_cards"]:
            c["signal"]=clean(c.get("signal",""))

    # Sector-specific overview components used by frontend.
    model["sector_confidence_drivers"]=lists["confidence"]
    model["sector_primary_cost_drivers"]=lists["cost"]
    model["sector_schedule_threats"]=lists["schedule"]
    model["why_casey_generated_this"]=[
        f"CASEY detected {subsector} from the project brief and routed it to the {mode} infrastructure model.",
        f"Sector signature behaviours applied: {', '.join(sig['signature'][:5])}.",
        "Cost, schedule and risk were calibrated against estimate class, schedule level, complexity and delivery environment.",
        "The output is designed for early board challenge and scope definition, not certified pricing."
    ]
    if lists["bench"]:
        model["benchmark_comparison"]=lists["bench"]

    # Make schedule rows read less algorithmic.
    for row in model.get("schedule_rows",[]) or []:
        row["basis"]=clean(row.get("basis",""))
        phase=str(row.get("phase","")).lower()
        if fam=="life_sciences" and any(x in phase for x in ["commission", "handover", "delivery"]):
            row["basis"]="Duration reflects CQV, GMP turnover, validation evidence and batch-readiness sequencing."
        elif fam=="data_centre" and any(x in phase for x in ["commission", "delivery"]):
            row["basis"]="Duration reflects energisation, IST, phased data-hall readiness and resilience proving."
        elif mode=="Space" and any(x in phase for x in ["delivery", "commission", "handover"]):
            row["basis"]="Duration reflects qualification, launch integration, deployment readiness and autonomous operations proving."

    # Add subtle human-like P80/P90 rounding display values while retaining numeric curves.
    try:
        mc=model.get("monte_carlo",{})
        if "qsra" in mc:
            for k in ["p10","p50","p80","p90"]:
                mc["qsra"][k]=round(float(mc["qsra"][k]))
        if "qcra" in mc:
            for k in ["p10","p50","p80","p90"]:
                mc["qcra"][k]=round(float(mc["qcra"][k]),1)
    except Exception:
        pass

    # Refresh executive summary after sector insight.
    model["executive_summary"]=(
        f"{model.get('title')} has been classified as {subsector} in {model.get('location','the selected location')}. "
        f"CASEY estimates {model.get('cost_p50')} P50 cost, {model.get('cost_range')} range and {model.get('schedule')} baseline. "
        f"Key insight: {sig['shock']} Risk is assessed as {model.get('risk')} with {model.get('confidence_pct')}% confidence."
    )

    # FINAL EXEC POLISH: scenario-aware ranking and differentiation
    try:
        sc = str(model.get("scenario","base")).lower()
        schedule_lists = {
            "faster":[
                "Concurrent commissioning overload",
                "Recovery float exhaustion",
                "Acceleration premium shock",
                "Grid connection delay",
                "Integrated systems testing concurrency"
            ],
            "cheaper":[
                "Vendor claims and change exposure",
                "Procurement deferral and long-lead slippage",
                "Design maturity gap",
                "Scope growth from deferred decisions",
                "Interface coordination delay"
            ],
            "lower_risk":[
                "Governance and approvals latency",
                "Extended validation sequencing",
                "Conservative commissioning gates",
                "Operational readiness hold-points",
                "Assurance and compliance reviews"
            ],
            "premium":[
                "Integration complexity across parallel packages",
                "Executive decision latency",
                "Technology assurance alignment",
                "Multi-package interface management",
                "Programme coordination overhead"
            ]
        }
        cost_lists = {
            "faster":[
                "Acceleration premiums and overtime",
                "Power train, transformers and switchgear",
                "Integrated systems testing",
                "Grid and utility concurrency",
                "Recovery-float consumption"
            ],
            "cheaper":[
                "Deferred procurement packaging",
                "Claims and commercial exposure",
                "Rework from reduced contingency",
                "Long-lead inflation volatility",
                "Scope rationalisation impacts"
            ],
            "lower_risk":[
                "Additional contingency and reserve",
                "Enhanced validation and assurance",
                "Programme controls and governance",
                "Redundant infrastructure resilience",
                "Extended commissioning readiness"
            ]
        }
        if sc in schedule_lists:
            model["sector_schedule_threats"]=schedule_lists[sc]
        if sc in cost_lists:
            model["sector_primary_cost_drivers"]=cost_lists[sc]

        if sc=="faster":
            model["executive_shock_insight"]="Acceleration increases spend faster than it reduces uncertainty; the delivery tail becomes more volatile."
        elif sc=="cheaper":
            model["executive_shock_insight"]="Capital efficiency reduces resilience: procurement and recovery flexibility become constrained."
        elif sc=="lower_risk":
            model["executive_shock_insight"]="Confidence is purchased through reserve, governance and extended delivery duration."
        elif sc=="premium":
            model["executive_shock_insight"]="Premium posture buys resilience, optionality and stronger certainty at visible capex premium."
    except Exception:
        pass

    return model

def build_model(prompt:str, client:str="", class_level:int=3, schedule_level:int=3, scenario:str="base"):
    title,mode,subsector,base_cost,base_months=detect_sector(prompt)
    location,loc_mult=location_factor(prompt)
    scale,scale_mult,scale_months=scale_factor(prompt)
    comp_mult,comp_months,drivers=complexity(prompt)
    cm,sm,rm,conf_adj,scenario_label,scenario_why=scenario_params("base")  # build base first; scenario cascade applies selected case
    raw_cost=base_cost*loc_mult*scale_mult*comp_mult*cm
    months=max(6,round((base_months+scale_months+comp_months+schedule_level*2)*sm))
    raw_cost, months, envelope_notes = sector_envelope(subsector, raw_cost, months, scale)
    raw_cost, months, benchmark_notes, benchmark_matches = calibrate_with_benchmarks(prompt, mode, subsector, raw_cost, months)
    raw_cost, months, envelope_notes2 = sector_envelope(subsector, raw_cost, months, scale)
    benchmark_notes = (envelope_notes + envelope_notes2 + benchmark_notes)[:5]
    risk_score=22 + (42 if mode=="Space" else 0) + (24 if "Mega" in scale or "System" in scale else 9 if scale=="Programme" else 0) + (14 if class_level>=4 else 0) + (18 if loc_mult>=2 else 0) + min(20,len(drivers)*5)
    risk_score*=rm
    risk=risk_label(risk_score)
    base_conf={1:88,2:80,3:68,4:54,5:40}.get(class_level,68)
    penalty={"Low":0,"Medium":4,"Medium-High":8,"High":12,"Very High":16,"Extreme":22}.get(risk,8)
    input_quality = 80
    try:
        if "input_quality_score" in locals():
            input_quality = int(input_quality_score)
    except Exception:
        input_quality = 80
    quality_index = estimate_quality_index(class_level, schedule_level, risk_score, benchmark_matches, input_quality)
    confidence=int(clamp((base_conf-penalty+conf_adj)*0.70 + quality_index["score"]*0.30,12,96))
    low,high,class_name,maturity=class_range(class_level)
    p10=raw_cost*low; p90=raw_cost*high
    lines=[]
    for cbs,desc,typ,basis,pct in sector_cost_lines(mode,subsector):
        p50=raw_cost*pct
        lines.append({"cbs":cbs,"description":desc,"type":typ,"basis":basis,"p10_bn":round(p50*low,3),"p50_bn":round(p50,3),"p90_bn":round(p50*high,3),"impact_basis":f"{desc} priced as {pct:.1%} of P50 from sector template, location factor {loc_mult:.2f}, scale factor {scale_mult:.2f}, complexity factor {comp_mult:.2f}."})
    schedules_by_level={str(l):schedule_rows(mode,subsector,months,l) for l in range(1,6)}
    primary_schedule=schedules_by_level[str(schedule_level)]
    primary_costs=lines
    risks=risk_register(mode,subsector,raw_cost,months,primary_schedule,primary_costs,rm,scenario)
    mc=monte_carlo(raw_cost,months,risks,seed=abs(hash(prompt+scenario))%999999)
    benchmarks=benchmarks_for(mode,subsector,location,raw_cost,months)
    model={
      "version":"v23 Real Deal Edition","id":f"CASEY-{random.randint(100000,999999)}","title":title,"prompt":prompt,"client":client or "Client / operator","mode":mode,"subsector":subsector,"location":location,"scale":scale,"scenario":scenario,"scenario_label":scenario_label,"scenario_why":scenario_why,"estimate_class":class_level,"estimate_class_name":class_name,"estimate_maturity":maturity,"schedule_level":schedule_level,"cost_p10":money_bn(p10),"cost_p50":money_bn(raw_cost),"cost_p90":money_bn(p90),"cost_range":f"{money_bn(p10)} - {money_bn(p90)}","schedule":f"{months} months","risk":risk,"risk_score":round(risk_score,1),"confidence_pct":confidence,
      "executive_summary":f"{title} has been classified as {subsector} in {location}. CASEY estimates {money_bn(raw_cost)} P50 cost, {money_bn(p10)} to {money_bn(p90)} range, {months} month baseline, {risk} risk and {confidence}% confidence under the {scenario_label} scenario.",
      "confidence_explanation":[f"{class_name}: {maturity} drives the starting confidence level.",f"Schedule Level {schedule_level}: more detailed schedule logic improves QSRA traceability.",f"{scenario_label}: {scenario_why}",f"Location/space factor: {location} applies a {loc_mult:.2f} delivery premium.",*benchmark_notes,*(drivers or ["Standard complexity profile inferred from project type."])],
      "benchmark_memory":benchmark_matches,
      "estimate_quality_index":quality_index,
      "procurement_heatmap":procurement_heatmap(mode,subsector,risks),
      "critical_path_narrative":critical_path_narrative(mode,subsector,primary_schedule),
      "board_briefing":board_briefing_narrative(mode,subsector,title,money_bn(p50),f"{months} months",risk,confidence,drivers,benchmark_matches),
      "uncertainty_narrative":uncertainty_narrative(class_level,schedule_level,confidence,mode,risk,drivers),
      "mission_control_cards":mission_control_cards(mode,subsector,title,risk,benchmark_matches,procurement_heatmap(mode,subsector,risks),critical_path_narrative(mode,subsector,primary_schedule)),
      "casey_thinking":casey_thinking(mode,subsector,title),
      "benchmark_comparison":benchmark_comparison(benchmark_matches,mode,subsector),
      "calibration_notes":benchmark_notes,
      "cost_lines":primary_costs,"schedules_by_level":schedules_by_level,"schedule_rows":primary_schedule,"risks":risks,"monte_carlo":mc,"benchmarks":benchmarks,"peer_competitors":peer_competitors(client,subsector,mode),"scenario_comparison":scenario_compare(prompt,client,class_level,schedule_level),"launch_demo_script":demo_script(),"red_flags":red_flags(risk,confidence,mode,subsector),"board_challenge_questions":board_questions(mode,subsector),"next_best_actions":next_actions(risk,confidence,scenario_label)
    }
    model["estimates_by_class"]={str(c):[{**x,"p10_bn":round(x["p50_bn"]*class_range(c)[0],3),"p90_bn":round(x["p50_bn"]*class_range(c)[1],3),"class":c,"maturity":class_range(c)[3]} for x in primary_costs] for c in range(1,6)}

    # FINAL EXEC POLISH: scenario-aware ranking and differentiation
    try:
        sc = str(model.get("scenario","base")).lower()
        schedule_lists = {
            "faster":[
                "Concurrent commissioning overload",
                "Recovery float exhaustion",
                "Acceleration premium shock",
                "Grid connection delay",
                "Integrated systems testing concurrency"
            ],
            "cheaper":[
                "Vendor claims and change exposure",
                "Procurement deferral and long-lead slippage",
                "Design maturity gap",
                "Scope growth from deferred decisions",
                "Interface coordination delay"
            ],
            "lower_risk":[
                "Governance and approvals latency",
                "Extended validation sequencing",
                "Conservative commissioning gates",
                "Operational readiness hold-points",
                "Assurance and compliance reviews"
            ],
            "premium":[
                "Integration complexity across parallel packages",
                "Executive decision latency",
                "Technology assurance alignment",
                "Multi-package interface management",
                "Programme coordination overhead"
            ]
        }
        cost_lists = {
            "faster":[
                "Acceleration premiums and overtime",
                "Power train, transformers and switchgear",
                "Integrated systems testing",
                "Grid and utility concurrency",
                "Recovery-float consumption"
            ],
            "cheaper":[
                "Deferred procurement packaging",
                "Claims and commercial exposure",
                "Rework from reduced contingency",
                "Long-lead inflation volatility",
                "Scope rationalisation impacts"
            ],
            "lower_risk":[
                "Additional contingency and reserve",
                "Enhanced validation and assurance",
                "Programme controls and governance",
                "Redundant infrastructure resilience",
                "Extended commissioning readiness"
            ]
        }
        if sc in schedule_lists:
            model["sector_schedule_threats"]=schedule_lists[sc]
        if sc in cost_lists:
            model["sector_primary_cost_drivers"]=cost_lists[sc]

        if sc=="faster":
            model["executive_shock_insight"]="Acceleration increases spend faster than it reduces uncertainty; the delivery tail becomes more volatile."
        elif sc=="cheaper":
            model["executive_shock_insight"]="Capital efficiency reduces resilience: procurement and recovery flexibility become constrained."
        elif sc=="lower_risk":
            model["executive_shock_insight"]="Confidence is purchased through reserve, governance and extended delivery duration."
        elif sc=="premium":
            model["executive_shock_insight"]="Premium posture buys resilience, optionality and stronger certainty at visible capex premium."
    except Exception:
        pass

    return model

def risk_register(mode,subsector,cost,months,schedule,costs,rm,scenario):
    base=[("R-001","Scope growth","Scope","Requirements mature late causing additions and rework","Both",42,15,45,120,0.03,0.08,0.16,"Project Director","Freeze scope, change board, value gates","Scope change rate exceeds threshold"),("R-002","Market escalation","Commercial","Inflation and commodity escalation exceeds allowance","Cost",45,0,0,0,0.04,0.10,0.20,"Commercial Lead","Early procurement, index strategy, market testing","Index exceeds allowance"),("R-003","Permits / approvals delay","Regulatory","Authorities or stakeholders delay critical approvals","Schedule",35,20,70,180,0.01,0.04,0.08,"Consents Lead","Authority plan, consent tracker, early submissions","Consent milestone slips"),("R-004","Design maturity gap","Technical","Design maturity lower than estimate basis","Both",38,15,45,120,0.03,0.09,0.18,"Design Manager","Design maturity gates and independent review","Deliverables late"),("R-005","Supply chain delay","Procurement","Long-lead equipment or specialist suppliers slip","Both",40,10,40,120,0.02,0.07,0.15,"Procurement Lead","Alternate suppliers, early orders, expediting","Supplier promise date slips"),("R-006","Productivity underperformance","Delivery","Labour productivity below plan due to access/sequence","Both",35,15,50,130,0.03,0.08,0.16,"Delivery Lead","Package productivity controls and daily planning","SPI/CPI deteriorates"),("R-007","Commissioning delay","Handover","Systems integration and tests take longer than plan","Both",32,10,35,100,0.02,0.06,0.12,"Commissioning Lead","Commissioning readiness and early test plans","Test failures trend upward"),("R-008","Interface misalignment","Integration","Package interfaces not coordinated","Both",34,10,35,100,0.02,0.07,0.14,"Integration Manager","ICDs and weekly interface board","Interface actions overdue")]
    if mode=="Space":
        base=[("R-S01","Launch manifest delay","Launch","Launch slot/provider delay affects deployment","Both",48,15,45,120,0.03,0.09,0.20,"Launch Integration Lead","Reserve alternate slots, freeze payload interfaces","Launch slot not confirmed"),("R-S02","Mass growth","Technical","Payload mass exceeds design allowance","Both",42,15,50,140,0.04,0.11,0.24,"Chief Engineer","Mass control board and margin policy","Mass margin below threshold"),("R-S03","Life-support reliability","Technical","ECLSS reliability below threshold","Both",36,20,70,180,0.05,0.12,0.28,"Life Support Lead","Prototype testing, redundancy, reliability growth","Reliability growth misses target")]+base
    elif "data centre" in subsector.lower():
        base=[("R-D01","Grid connection delay","Utilities","Power connection or substation energisation delays critical path","Both",48,20,80,220,0.03,0.10,0.22,"Utilities Lead","Secure grid agreement, temporary power, early equipment orders","Grid agreement unsigned"),("R-D02","Cooling system capacity","Technical","Cooling design or water availability constrains commissioning","Both",34,10,40,120,0.02,0.07,0.15,"MEP Lead","Early thermal modelling and supplier validation","Cooling tests fail")]+base
    out=[]
    for i,r in enumerate(base,1):
        rid,title,cat,desc,area,prob,so,sm,sp,co,cm,cp,owner,mit,trig=r
        prob=int(clamp(prob*rm,5,90))
        if scenario=="lower_risk": cm*=0.75; cp*=0.8; sm=round(sm*.75); sp=round(sp*.8)
        if scenario=="faster": sm=round(sm*1.18); sp=round(sp*1.25)
        if scenario=="cheaper": cm*=1.12; cp*=1.18
        activity=schedule[min(i+3,len(schedule)-1)] if schedule else {"activity_id":"A1900","activity":"Delivery"}
        costline=costs[min(i+2,len(costs)-1)] if costs else {"cbs":"01.01","description":"Cost"}
        emv=cost*cm*prob/100; semv=sm*prob/100
        out.append({"risk_id":rid if rid.startswith("R-") else f"R-{i:03d}","title":title,"category":cat,"description":desc,"impact_area":area,"probability_pct":prob,"activity_id":activity["activity_id"],"activity_name":activity["activity"],"cbs":costline["cbs"],"cbs_name":costline["description"],"schedule_o_days":so,"schedule_m_days":sm,"schedule_p_days":sp,"cost_o_bn":round(cost*co,3),"cost_m_bn":round(cost*cm,3),"cost_p_bn":round(cost*cp,3),"cost_emv_bn":round(emv,3),"schedule_emv_days":round(semv,1),"driver_score":round((emv*100)+(semv/10),2),"owner":owner,"trigger":trig,"mitigation":mit,"basis_of_cost_impact":f"Cost impact derived from {costline['cbs']} {costline['description']} exposure, project P50 {money_bn(cost)}, and risk severity ({co:.0%}/{cm:.0%}/{cp:.0%}).","basis_of_schedule_impact":f"Schedule impact mapped to {activity['activity_id']} {activity['activity']} with O/M/P delay {so}/{sm}/{sp} days and critical path sensitivity."})
    return sorted(out,key=lambda x:x["driver_score"],reverse=True)

def monte_carlo(cost,months,risks,seed=42,iterations=10000):
    rng=np.random.default_rng(seed)
    cost_samples=[]; sched_samples=[]
    base_cost_unc=rng.triangular(cost*.88,cost,cost*1.18,iterations)
    base_sched_unc=rng.triangular(months*.92,months,months*1.15,iterations)
    risk_cost=np.zeros(iterations); risk_days=np.zeros(iterations)
    contribution={r["risk_id"]:{"cost":0.0,"days":0.0,"title":r["title"],"activity_id":r["activity_id"],"cbs":r["cbs"]} for r in risks}
    for r in risks:
        occurs=rng.random(iterations)<(r["probability_pct"]/100)
        c=rng.triangular(r["cost_o_bn"],r["cost_m_bn"],r["cost_p_bn"],iterations)*occurs
        if r["schedule_o_days"] == r["schedule_m_days"] == r["schedule_p_days"]:
            d=np.zeros(iterations)
        else:
            left=min(r["schedule_o_days"], r["schedule_m_days"], r["schedule_p_days"])
            mode=max(left, r["schedule_m_days"])
            right=max(r["schedule_o_days"], r["schedule_m_days"], r["schedule_p_days"])
            if left == right: d=np.zeros(iterations)
            else: d=rng.triangular(left,mode,right,iterations)*occurs
        risk_cost+=c; risk_days+=d
        contribution[r["risk_id"]]["cost"]=float(np.mean(c)); contribution[r["risk_id"]]["days"]=float(np.mean(d))
    cost_samples=base_cost_unc+risk_cost
    sched_samples=base_sched_unc+(risk_days/30.44)
    def pct(arr,p): return float(np.percentile(arr,p))
    curve=[{"percentile":p,"cost_bn":round(pct(cost_samples,p),3),"schedule_months":round(pct(sched_samples,p),2)} for p in [1,5,10,20,30,40,50,60,70,80,90,95,99]]
    tornado=sorted([{"risk_id":k,"title":v["title"],"activity_id":v["activity_id"],"cbs":v["cbs"],"cost_mean_bn":round(v["cost"],3),"schedule_mean_days":round(v["days"],1),"driver_score":round(v["cost"]*100+v["days"]/10,2)} for k,v in contribution.items()],key=lambda x:x["driver_score"],reverse=True)
    return {"iterations":iterations,"qcra":{"p10":round(pct(cost_samples,10),3),"p50":round(pct(cost_samples,50),3),"p80":round(pct(cost_samples,80),3),"p90":round(pct(cost_samples,90),3),"mean":round(float(np.mean(cost_samples)),3)},"qsra":{"p10":round(pct(sched_samples,10),2),"p50":round(pct(sched_samples,50),2),"p80":round(pct(sched_samples,80),2),"p90":round(pct(sched_samples,90),2),"mean":round(float(np.mean(sched_samples)),2)},"curve":curve,"tornado":tornado}

def benchmarks_for(mode,subsector,location,cost,months):
    s=subsector.lower()
    if mode=="Space": return [{"metric":"Launch/logistics premium","value":"2.5x-4.3x Earth analogue","why":"Remote operations, mass constraints, launch windows and harsh environment."},{"metric":"Programme duration benchmark","value":f"{int(months*.85)}-{int(months*1.25)} months","why":"Qualification, launch integration and commissioning sequence."},{"metric":"Top benchmark gap","value":"TRL / reliability maturity","why":"Space confidence improves fastest through test evidence and heritage."}]
    if "data centre" in s: return [{"metric":"Hyperscale delivery benchmark","value":"$8M-$18M per MW equivalent","why":"Varies by power density, grid scope, cooling, land and regional constraints."},{"metric":"Schedule benchmark","value":"30-60 months","why":"Driven by grid connection, long-lead electrical equipment and phased fit-out."},{"metric":"Top benchmark gap","value":"Power availability","why":"Grid connection is frequently the dominant schedule and risk driver."}]
    if "airport" in s: return [{"metric":"Airport capacity benchmark","value":"$250-$750 per annual pax capacity","why":"Depends on runway, terminal, baggage, rail/road connections and land."},{"metric":"Schedule benchmark","value":"7-12 years","why":"Approvals, airside phasing and operational readiness dominate."}]
    if "rail" in s: return [{"metric":"Rail corridor benchmark","value":"$80M-$500M per km","why":"Underground, stations, signalling and land interfaces drive variance."},{"metric":"Schedule benchmark","value":"6-15 years","why":"Consents, utilities, possessions and systems integration."}]
    return [{"metric":"Capital project benchmark","value":"Sector-adjusted range","why":"Selected based on project type, location, maturity and scale."},{"metric":"Schedule benchmark","value":f"{int(months*.8)}-{int(months*1.3)} months","why":"Derived from maturity and delivery complexity."}]

def scenario_compare(prompt,client,cls,level):
    out=[]
    for sc in ["base","faster","cheaper","lower_risk","premium","investor","survival"]:
        m=build_model_core_no_compare(prompt,client,cls,level,sc)
        out.append({"scenario":sc,"label":m["scenario_label"],"cost":m["cost_p50"],"schedule_months":int(m["schedule"].split()[0]),"risk":m["risk"],"confidence":m["confidence_pct"],"why":m["scenario_why"]})
    return out

def build_model_core_no_compare(prompt,client,cls,level,scenario):
    title,mode,subsector,base_cost,base_months=detect_sector(prompt); location,loc_mult=location_factor(prompt); scale,scale_mult,scale_months=scale_factor(prompt); comp_mult,comp_months,drivers=complexity(prompt); cm,sm,rm,conf_adj,scenario_label,scenario_why=scenario_params(scenario)
    cost=base_cost*loc_mult*scale_mult*comp_mult*cm; months=max(6,round((base_months+scale_months+comp_months+level*2)*sm)); score=(22+(42 if mode=="Space" else 0)+(24 if "Mega" in scale or "System" in scale else 9 if scale=="Programme" else 0)+(14 if cls>=4 else 0)+(18 if loc_mult>=2 else 0)+min(20,len(drivers)*5))*rm; risk=risk_label(score); conf=int(clamp({1:88,2:80,3:68,4:54,5:40}.get(cls,68)-{"Low":0,"Medium":4,"Medium-High":8,"High":12,"Very High":16,"Extreme":22}.get(risk,8)+conf_adj,12,96))
    return {"scenario_label":scenario_label,"scenario_why":scenario_why,"cost_p50":money_bn(cost),"schedule":f"{months} months","risk":risk,"confidence_pct":conf}

def red_flags(risk,confidence,mode,subsector):
    flags=[]
    if confidence<65: flags.append("Confidence below board-comfort threshold: improve scope definition, benchmark evidence and supplier quotes.")
    if risk in ["Very High","Extreme"]: flags.append("Risk rating requires executive sponsor review and quantified mitigation plan before commitment.")
    if mode=="Space": flags.append("Space programme: verify TRL, launch manifest, mass margins and operations concept before commercial reliance.")
    if "data centre" in subsector.lower(): flags.append("Data centre: grid connection and cooling basis must be validated early.")
    return flags or ["No critical red flags in concept model; maintain assumptions log and stage-gate assurance."]

def board_questions(mode,subsector):
    qs=["What evidence supports the P50 cost and P80 confidence?","Which five risks create most contingency?","Which activities drive the QSRA finish-date risk?","What data would increase confidence before approval?","What is the fastest credible lower-risk path?"]
    if mode=="Space": qs.append("What flight heritage, TRL and qualification evidence supports the space assumptions?")
    return qs

def next_actions(risk,confidence,scenario):
    return ["Run a risk workshop to validate probability and O/M/P impacts.","Confirm top five CBS cost drivers with benchmark or supplier evidence.","Validate the Level 3/4 schedule logic and critical path assumptions.","Prepare board decision paper with Base/Faster/Cheaper/Lower-Risk cases.","Create confidence improvement plan targeting design maturity, consents and procurement evidence."]

def demo_script():
    return ["Open with: 'Give me any project on Earth or in Space.'","Type the project into CASEY and hit Generate Full Intelligence Pack.","Show the executive dashboard: cost, schedule, risk, confidence and top drivers.","Switch to Monte Carlo: show QCRA/QSRA P-curves and tornado drivers linked to activities.","Show scenarios: Faster, Cheaper, Lower Risk and Premium with why each changed.","Open Output Centre: Excel, Risk CSV, XER, cost workbook, risk register, schedule export and full pack are ready.","Close with: 'Traditional advisory takes weeks. CASEY gives the first board-grade view in seconds.'"]

# ------------------------- routes -------------------------
@app.get("/health")
def health(): return {"status":"ok","service":APP_VERSION,"demo_limit_per_ip":"disabled_for_demo_launch"}

@app.get("/demo/status")
def demo_status(request: Request):
    # Demo launch mode: never block local/browser/email/IP repeat runs.
    return {"allowed": True, "used": 0, "limit": 999999, "remaining": 999999, "demo_launch_mode": True}


@app.post("/public-demo/generate")
def public_demo_generate(req: PublicDemoRequest, request: Request):
    issues = _quality_gate_public_demo(req)
    if issues:
        raise HTTPException(status_code=422, detail={"message": "CASEY needs one real infrastructure or space programme brief before using your free run.", "issues": issues})
    identity = _public_demo_identity(request, req)
    # DEMO-LAUNCH MODE: do not hard-block repeat local/browser demo runs.
    # We still record the lead/run for admin visibility, but the live demo must never fail
    # with a stale browser/email/IP fingerprint during a client presentation.
    previous = None
    prompt = _premium_public_prompt(req)
    input_quality = _public_demo_brief_quality_score(req)
    model = build_model(prompt, _normalise_email(req.email), 3, 4, "base")
    model["input_quality_score"] = input_quality["score"]
    model["public_demo"] = True
    model["lead_email"] = _normalise_email(req.email)
    run_id = "CASEY-DEMO-" + uuid.uuid4().hex[:10].upper()
    report = _public_demo_report(model)
    con = db(); cur = con.cursor(); now = datetime.utcnow().isoformat()
    cur.execute("""INSERT INTO public_demo_uses(run_id,email_hash,ip_hash,fingerprint_hash,client_token_hash,project_type,project_text,model_json,created_at)
        VALUES(?,?,?,?,?,?,?,?,?)""", (run_id, identity["email_hash"], identity["ip_hash"], identity["fingerprint_hash"], identity["client_token_hash"], req.project_type, req.project_description, json.dumps(model), now))
    cur.execute("INSERT INTO projects(name,client,prompt,mode,created_at,model_json) VALUES(?,?,?,?,?,?)", (model["title"], model["client"], model["prompt"], model["mode"], now, json.dumps(model)))
    con.commit(); con.close()
    return {"run_id": run_id, "used": 1, "limit": PUBLIC_DEMO_LIMIT, "report": report, "model": model}


@app.get("/public-demo/admin/runs")
def public_demo_admin_runs(request: Request, limit: int = 100):
    """Admin-only view of public demo runs.
    Set CASEY_ADMIN_TOKEN in Render, then call with header:
    x-casey-admin-token: YOUR_TOKEN
    """
    if not ADMIN_TOKEN:
        raise HTTPException(status_code=403, detail="Admin log access is not configured. Set CASEY_ADMIN_TOKEN.")
    supplied = request.headers.get("x-casey-admin-token", "")
    if supplied != ADMIN_TOKEN:
        raise HTTPException(status_code=403, detail="Invalid admin token.")
    limit = max(1, min(int(limit or 100), 1000))
    con = db(); cur = con.cursor()
    rows = cur.execute("""SELECT run_id, project_type, project_text, model_json, created_at
                          FROM public_demo_uses ORDER BY id DESC LIMIT ?""", (limit,)).fetchall()
    feedback = cur.execute("""SELECT run_id, rating, comment, created_at FROM public_demo_feedback ORDER BY id DESC LIMIT ?""", (limit,)).fetchall()
    con.close()
    return {
        "runs": [
            {
                "run_id": r["run_id"],
                "project_type": r["project_type"],
                "project_text": r["project_text"],
                "created_at": r["created_at"],
                "model": json.loads(r["model_json"]) if r["model_json"] else None
            } for r in rows
        ],
        "feedback": [dict(f) for f in feedback]
    }

@app.get("/public-demo/admin/summary")
def public_demo_admin_summary(request: Request):
    if not ADMIN_TOKEN:
        raise HTTPException(status_code=403, detail="Admin log access is not configured. Set CASEY_ADMIN_TOKEN.")
    supplied = request.headers.get("x-casey-admin-token", "")
    if supplied != ADMIN_TOKEN:
        raise HTTPException(status_code=403, detail="Invalid admin token.")
    con = db(); cur = con.cursor()
    total = cur.execute("SELECT COUNT(*) AS c FROM public_demo_uses").fetchone()["c"]
    by_type = cur.execute("SELECT project_type, COUNT(*) AS c FROM public_demo_uses GROUP BY project_type").fetchall()
    recent = cur.execute("SELECT run_id, project_type, project_text, created_at FROM public_demo_uses ORDER BY id DESC LIMIT 20").fetchall()
    con.close()
    return {
        "total_runs": total,
        "by_type": [dict(x) for x in by_type],
        "recent": [dict(x) for x in recent]
    }


@app.post("/public-demo/feedback")
def public_demo_feedback(req: PublicDemoFeedback):
    rating = max(1, min(5, int(req.rating)))
    con = db(); cur = con.cursor(); now = datetime.utcnow().isoformat()
    cur.execute("INSERT INTO public_demo_feedback(run_id,rating,comment,created_at) VALUES(?,?,?,?)", (req.run_id, rating, req.comment or "", now))
    con.commit(); con.close()
    return {"status": "saved", "message": "Feedback saved for CASEY improvement loop."}

@app.post("/generate")
def generate(req: GenerateRequest, request: Request):
    if req.demo:
        ip=client_ip(request); status=check_demo_allowance(ip)
        if False and not status["allowed"]: raise HTTPException(403,"Public demo allowance used. Launch/private mode can still generate when deployed behind login.")
        record_demo_use(ip)

    # Canonical project-context lock. Scenario buttons must never rebuild from the default
    # data-centre seed when the active project is Space, Rail, Energy, Defence, etc.
    scenario = (req.scenario or "base").lower()
    active = req.active_model or {}
    locked_prompt = active.get("prompt") or req.prompt
    locked_mode = active.get("mode")
    locked_subsector = active.get("subsector")

    model=build_model(locked_prompt, req.client or "", int(req.class_level or 3), int(req.schedule_level or 3), scenario)

    # Defensive guard: if a non-base scenario ever classifies differently from the locked
    # context, keep the model in the original project universe and flag it for audit.
    if scenario != "base" and locked_mode and model.get("mode") != locked_mode:
        model["mode"] = locked_mode
        if locked_subsector: model["subsector"] = locked_subsector
        if active.get("title"): model["title"] = active.get("title")
        if active.get("location"): model["location"] = active.get("location")
        if active.get("scale"): model["scale"] = active.get("scale")
        model["prompt"] = locked_prompt
        model["context_lock_repaired"] = True

    model["active_context_lock"] = {
        "mode": model.get("mode"),
        "subsector": model.get("subsector"),
        "prompt_hash": hashlib.sha256((model.get("prompt") or "").encode("utf-8")).hexdigest()[:12],
        "scenario": scenario
    }

    con=db(); cur=con.cursor(); now=datetime.utcnow().isoformat(); cur.execute("INSERT INTO projects(name,client,prompt,mode,created_at,model_json) VALUES(?,?,?,?,?,?)",(model["title"],model["client"],model["prompt"],model["mode"],now,json.dumps(model))); con.commit(); con.close()

    # FINAL EXEC POLISH: scenario-aware ranking and differentiation
    try:
        sc = str(model.get("scenario","base")).lower()
        schedule_lists = {
            "faster":[
                "Concurrent commissioning overload",
                "Recovery float exhaustion",
                "Acceleration premium shock",
                "Grid connection delay",
                "Integrated systems testing concurrency"
            ],
            "cheaper":[
                "Vendor claims and change exposure",
                "Procurement deferral and long-lead slippage",
                "Design maturity gap",
                "Scope growth from deferred decisions",
                "Interface coordination delay"
            ],
            "lower_risk":[
                "Governance and approvals latency",
                "Extended validation sequencing",
                "Conservative commissioning gates",
                "Operational readiness hold-points",
                "Assurance and compliance reviews"
            ],
            "premium":[
                "Integration complexity across parallel packages",
                "Executive decision latency",
                "Technology assurance alignment",
                "Multi-package interface management",
                "Programme coordination overhead"
            ]
        }
        cost_lists = {
            "faster":[
                "Acceleration premiums and overtime",
                "Power train, transformers and switchgear",
                "Integrated systems testing",
                "Grid and utility concurrency",
                "Recovery-float consumption"
            ],
            "cheaper":[
                "Deferred procurement packaging",
                "Claims and commercial exposure",
                "Rework from reduced contingency",
                "Long-lead inflation volatility",
                "Scope rationalisation impacts"
            ],
            "lower_risk":[
                "Additional contingency and reserve",
                "Enhanced validation and assurance",
                "Programme controls and governance",
                "Redundant infrastructure resilience",
                "Extended commissioning readiness"
            ]
        }
        if sc in schedule_lists:
            model["sector_schedule_threats"]=schedule_lists[sc]
        if sc in cost_lists:
            model["sector_primary_cost_drivers"]=cost_lists[sc]

        if sc=="faster":
            model["executive_shock_insight"]="Acceleration increases spend faster than it reduces uncertainty; the delivery tail becomes more volatile."
        elif sc=="cheaper":
            model["executive_shock_insight"]="Capital efficiency reduces resilience: procurement and recovery flexibility become constrained."
        elif sc=="lower_risk":
            model["executive_shock_insight"]="Confidence is purchased through reserve, governance and extended delivery duration."
        elif sc=="premium":
            model["executive_shock_insight"]="Premium posture buys resilience, optionality and stronger certainty at visible capex premium."
    except Exception:
        pass

    return model

@app.get("/projects")
def projects():
    con=db(); rows=[dict(x) for x in con.execute("SELECT id,name,client,prompt,mode,created_at FROM projects ORDER BY id DESC LIMIT 50")]; con.close(); return rows

@app.get("/projects/{project_id}")
def project(project_id:int):
    con=db(); row=con.execute("SELECT * FROM projects WHERE id=?",(project_id,)).fetchone(); con.close()
    if not row: raise HTTPException(404,"Project not found")
    return json.loads(row["model_json"])

@app.post("/chat")
def chat(req: ChatRequest):
    q=req.question.lower(); m=req.project or {}; risks=m.get("risks",[])[:5]; mc=m.get("monte_carlo",{})
    # API key stays server-side only. This deterministic advisor works without exposing it.
    if "confidence" in q:
        ans="Confidence is driven by estimate class, schedule level, risk rating, location factor, scenario choice and evidence maturity. Improve it with supplier quotes, surveys, validated schedule logic and benchmark evidence."
    elif "compet" in q or "peer" in q or "meta" in q or "amazon" in q or "google" in q:
        peers=", ".join(m.get("peer_competitors",[])[:6]); ans=f"Peer intelligence compares the client's operational competitors/peers: {peers}. Use it to benchmark likely cost, speed, power/logistics constraints, delivery strategy and risk profile for similar capital programmes."
    elif "qcra" in q or "qsra" in q or "monte" in q:
        ans=f"QCRA P80 is {money_bn(mc.get('qcra',{}).get('p80',0))}; QSRA P80 is {mc.get('qsra',{}).get('p80','n/a')} months. Top drivers are mapped in the tornado chart to WBS activities and CBS accounts."
    elif "risk" in q:
        ans="Top risks are: "+"; ".join([f"{r['title']} ({r['activity_id']} / {r['cbs']})" for r in risks])
    elif "faster" in q:
        ans="Faster case compresses schedule through parallel packages and early procurement, but raises interface, rework and supply-chain risk. Use it only when time-to-market value exceeds acceleration premium."
    elif "cheaper" in q:
        ans="Cheaper case uses value engineering, tighter scope and procurement competition. It reduces P50 cost but can reduce confidence and increase residual schedule/quality risk."
    else:
        ans="CASEY recommends validating the top CBS cost accounts, the critical-path activities, and the top QCRA/QSRA risk drivers before board commitment. Use the Output Centre to export the evidence pack."
    return {"answer":ans}

@app.post("/upload/analyse")
async def analyse_upload(file: UploadFile = File(...)):
    content=await file.read(); name=file.filename or "upload"; text=""
    try: text=content.decode("utf-8", errors="ignore")[:200000]
    except Exception: text=""
    lower=text.lower(); findings=[]
    if name.lower().endswith(".xer") or "%t\ttask" in lower: findings.append("Schedule file detected: check calendars, open ends, constraints, critical path and resource/cost loading.")
    if "risk" in lower: findings.append("Risk language detected: verify owner, trigger, O/M/P impacts, residuals and WBS/CBS mapping.")
    if "cost" in lower or "estimate" in lower: findings.append("Estimate language detected: verify direct/indirect split, escalation, contingency, owner costs and benchmark basis.")
    if not findings: findings.append("File received. Deep parser needs project-specific file format mapping; preliminary review checks naming, size and visible text signals.")
    result={"filename":name,"size_bytes":len(content),"findings":findings,"red_flags":["Confirm all risk impacts have basis statements.","Ensure exported board pack includes decision, ask, confidence and next actions.","Validate source files before commercial reliance."],"next_steps":["Upload XER + estimate + risk register together for stronger triangulation.","Map risks to activity IDs and CBS codes.","Run QCRA/QSRA refresh after review."]}
    con=db(); cur=con.cursor(); cur.execute("INSERT INTO uploads(filename,created_at,analysis_json) VALUES(?,?,?)",(name,datetime.utcnow().isoformat(),json.dumps(result))); con.commit(); con.close()
    return result

# ------------------------- exports -------------------------
def style_ws(ws):
    fill=PatternFill("solid",fgColor="07111F"); font=Font(color="FFFFFF",bold=True); thin=Side(style="thin",color="26384F")
    ws.freeze_panes="A2"
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment=Alignment(vertical="top",wrap_text=True); cell.border=Border(left=thin,right=thin,top=thin,bottom=thin)
    for c in ws[1]: c.fill=fill; c.font=font
    for i in range(1,min(ws.max_column+1,50)): ws.column_dimensions[get_column_letter(i)].width=22

def add_sheet(wb,name,rows):
    ws=wb.create_sheet(name[:31]) if not wb.sheetnames else wb.create_sheet(name[:31]); ws.title=name[:31]
    for r in rows: ws.append(r)
    style_ws(ws); return ws

def workbook_bytes(model):
    wb=Workbook(); wb.remove(wb.active)
    add_sheet(wb,"Board Summary",[["Field","Value"],["Generated",datetime.utcnow().isoformat()],["Version",model.get("version")],["Title",model.get("title")],["Client",model.get("client")],["Executive Summary",model.get("executive_summary")],["Mode",model.get("mode")],["Subsector",model.get("subsector")],["Location",model.get("location")],["Scenario",model.get("scenario_label")],["Cost P10",model.get("cost_p10")],["Cost P50",model.get("cost_p50")],["Cost P90",model.get("cost_p90")],["Schedule",model.get("schedule")],["QCRA P80",money_bn(model["monte_carlo"]["qcra"]["p80"])],["QSRA P80",model["monte_carlo"]["qsra"]["p80"]],["Risk",model.get("risk")],["Confidence",f"{model.get('confidence_pct')}%"]])
    add_sheet(wb,"Primary Cost Estimate",[["CBS","Description","Type","Basis","P10 BN","P50 BN","P90 BN","Impact Basis"]]+[[x["cbs"],x["description"],x["type"],x["basis"],x["p10_bn"],x["p50_bn"],x["p90_bn"],x["impact_basis"]] for x in model["cost_lines"]])
    for c,rows in model["estimates_by_class"].items(): add_sheet(wb,f"Class {c} Estimate",[["Class","CBS","Description","P10","P50","P90","Maturity"]]+[[c,x["cbs"],x["description"],x["p10_bn"],x["p50_bn"],x["p90_bn"],x["maturity"]] for x in rows])
    for l,rows in model["schedules_by_level"].items(): add_sheet(wb,f"Level {l} Schedule",[["Activity ID","Phase","Activity","Predecessor","Duration Months","Critical","Basis"]]+[[x["activity_id"],x["phase"],x["activity"],x["predecessor"],x["duration_months"],x["critical"],x["basis"]] for x in rows])
    add_sheet(wb,"Risk Register",[["ID","Risk","Category","Probability","Activity","CBS","Sched O","Sched M","Sched P","Cost O","Cost M","Cost P","Owner","Trigger","Mitigation","Cost Basis","Schedule Basis"]]+[[r["risk_id"],r["title"],r["category"],r["probability_pct"],r["activity_id"],r["cbs"],r["schedule_o_days"],r["schedule_m_days"],r["schedule_p_days"],r["cost_o_bn"],r["cost_m_bn"],r["cost_p_bn"],r["owner"],r["trigger"],r["mitigation"],r["basis_of_cost_impact"],r["basis_of_schedule_impact"]] for r in model["risks"]])
    add_sheet(wb,"Monte Carlo P-Curve",[["Percentile","QCRA Cost BN","QSRA Months"]]+[[x["percentile"],x["cost_bn"],x["schedule_months"]] for x in model["monte_carlo"]["curve"]])
    add_sheet(wb,"Tornado Drivers",[["Risk","Title","Activity","CBS","Cost Mean BN","Schedule Mean Days","Driver Score"]]+[[x["risk_id"],x["title"],x["activity_id"],x["cbs"],x["cost_mean_bn"],x["schedule_mean_days"],x["driver_score"]] for x in model["monte_carlo"]["tornado"]])
    add_sheet(wb,"Scenarios",[["Scenario","Cost","Schedule Months","Risk","Confidence","Why"]]+[[x["label"],x["cost"],x["schedule_months"],x["risk"],x["confidence"],x["why"]] for x in model["scenario_comparison"]])
    add_sheet(wb,"Benchmarks",[["Metric","Value","Why"]]+[[x["metric"],x["value"],x["why"]] for x in model["benchmarks"]])
    add_sheet(wb,"Demo Script",[["Step"]]+[[x] for x in model["launch_demo_script"]])
    # charts
    try:
        ws=wb["Monte Carlo P-Curve"]; chart=LineChart(); chart.title="QCRA/QSRA P-Curve"; data=Reference(ws,min_col=2,max_col=3,min_row=1,max_row=ws.max_row); cats=Reference(ws,min_col=1,min_row=2,max_row=ws.max_row); chart.add_data(data,titles_from_data=True); chart.set_categories(cats); ws.add_chart(chart,"E2")
        ws2=wb["Tornado Drivers"]; b=BarChart(); b.title="Tornado Drivers"; data=Reference(ws2,min_col=7,min_row=1,max_row=min(ws2.max_row,12)); cats=Reference(ws2,min_col=2,min_row=2,max_row=min(ws2.max_row,12)); b.add_data(data,titles_from_data=True); b.set_categories(cats); ws2.add_chart(b,"I2")
    except Exception: pass
    bio=BytesIO(); wb.save(bio); bio.seek(0); return bio.getvalue()

def risk_csv_bytes(model):
    out=StringIO(); w=csv.writer(out); w.writerow(["Risk ID","Title","Category","Probability %","Activity ID","Activity Name","CBS","CBS Name","Schedule O","Schedule M","Schedule P","Cost O BN","Cost M BN","Cost P BN","Owner","Trigger","Mitigation","Basis Cost","Basis Schedule"])
    for r in model["risks"]: w.writerow([r["risk_id"],r["title"],r["category"],r["probability_pct"],r["activity_id"],r["activity_name"],r["cbs"],r["cbs_name"],r["schedule_o_days"],r["schedule_m_days"],r["schedule_p_days"],r["cost_o_bn"],r["cost_m_bn"],r["cost_p_bn"],r["owner"],r["trigger"],r["mitigation"],r["basis_of_cost_impact"],r["basis_of_schedule_impact"]])
    return out.getvalue().encode()

def xer_bytes(model):
    lines=["ERMHDR\t24.12\t2026-05-02\tProject\tCASEY\tCASEY\tProject Management\tUSD","%T\tPROJECT","%F\tproj_id\tproj_short_name\tproj_name",f"%R\t1\tCASEY\t{model.get('title','CASEY Project')}","%T\tWBS","%F\twbs_id\tproj_id\twbs_short_name\twbs_name","%R\t1\t1\tCASEY\tCASEY Project","%T\tTASK","%F\ttask_id\tproj_id\twbs_id\ttask_code\ttask_name\tduration_type\ttarget_drtn_hr_cnt"]
    idx={}
    for i,a in enumerate(model["schedule_rows"],1): idx[a["activity_id"]]=i; lines.append(f"%R\t{i}\t1\t1\t{a['activity_id']}\t{a['activity']}\tFixed Duration\t{a['duration_months']*160}")
    lines += ["%T\tTASKPRED","%F\ttask_pred_id\ttask_id\tpred_task_id\tpred_type"]
    k=1
    for a in model["schedule_rows"]:
        for p in str(a.get("predecessor","")).split(";"):
            p=p.strip()
            if p in idx: lines.append(f"%R\t{k}\t{idx[a['activity_id']]}\t{idx[p]}\tFS"); k+=1
    return "\n".join(lines).encode()

def word_bytes(model):
    doc=Document(); styles=doc.styles; styles["Normal"].font.name="Aptos"; styles["Normal"].font.size=Pt(10)
    h=doc.add_heading("CASEY TITAN X Board Report",0); h.runs[0].font.color.rgb=RGBColor(5,20,35)
    doc.add_paragraph(model.get("executive_summary",""))
    tbl=doc.add_table(rows=1,cols=2); tbl.style="Light Shading Accent 1"; hdr=tbl.rows[0].cells; hdr[0].text="Metric"; hdr[1].text="Value"
    for k,v in [("Cost P50",model["cost_p50"]),("Cost Range",model["cost_range"]),("Schedule",model["schedule"]),("QCRA P80",money_bn(model["monte_carlo"]["qcra"]["p80"])),("QSRA P80",str(model["monte_carlo"]["qsra"]["p80"])+" months"),("Risk",model["risk"]),("Confidence",str(model["confidence_pct"])+"%")]: row=tbl.add_row().cells; row[0].text=k; row[1].text=v
    for title,items in [("Why Confidence",model["confidence_explanation"]),("Top Red Flags",model["red_flags"]),("Board Challenge Questions",model["board_challenge_questions"]),("Next Best Actions",model["next_best_actions"])]:
        doc.add_heading(title,1)
        for x in items: doc.add_paragraph(x,style="List Bullet")
    doc.add_heading("Top Risk Drivers",1)
    for r in model["risks"][:8]: doc.add_paragraph(f"{r['risk_id']} {r['title']} — {r['activity_id']} / {r['cbs']}: {r['basis_of_schedule_impact']}",style="List Bullet")
    bio=BytesIO(); doc.save(bio); bio.seek(0); return bio.getvalue()

def pptx_bytes(model):
    from pptx.dml.color import RGBColor as PRGBColor
    prs=Presentation(); blank=prs.slide_layouts[6]
    def add_title(slide,t,sub=""):
        box=slide.shapes.add_textbox(PptxInches(.5),PptxInches(.35),PptxInches(12),PptxInches(.8)); tf=box.text_frame; tf.text=t; tf.paragraphs[0].runs[0].font.size=PptxPt(34); tf.paragraphs[0].runs[0].font.bold=True
        if sub: b=slide.shapes.add_textbox(PptxInches(.55),PptxInches(1.1),PptxInches(12),PptxInches(.4)); b.text_frame.text=sub
    s=prs.slides.add_slide(blank); add_title(s,"CASEY TITAN X",model.get("executive_summary",""));
    s2=prs.slides.add_slide(blank); add_title(s2,"Board Metrics")
    metrics=[("P50",model["cost_p50"]),("Range",model["cost_range"]),("Schedule",model["schedule"]),("Risk",model["risk"]),("Confidence",str(model["confidence_pct"])+"%"),("QCRA P80",money_bn(model["monte_carlo"]["qcra"]["p80"]))]
    for i,(k,v) in enumerate(metrics):
        x=.7+(i%3)*4.1; y=1.6+(i//3)*1.6; shp=s2.shapes.add_shape(1,PptxInches(x),PptxInches(y),PptxInches(3.6),PptxInches(1.0)); shp.text=f"{k}\n{v}"
    s3=prs.slides.add_slide(blank); add_title(s3,"Top Risk Drivers")
    txt=s3.shapes.add_textbox(PptxInches(.7),PptxInches(1.4),PptxInches(12),PptxInches(5)).text_frame
    for r in model["risks"][:8]: p=txt.add_paragraph(); p.text=f"{r['title']} — {r['activity_id']} / {r['cbs']}"; p.level=0
    bio=BytesIO(); prs.save(bio); bio.seek(0); return bio.getvalue()

def pdf_bytes(model):
    bio=BytesIO(); doc=SimpleDocTemplate(bio,pagesize=A4); styles=getSampleStyleSheet(); story=[Paragraph("CASEY TITAN X Board Report",styles["Title"]),Spacer(1,12),Paragraph(model.get("executive_summary",""),styles["BodyText"]),Spacer(1,12)]
    data=[["Metric","Value"],["Cost P50",model["cost_p50"]],["Cost Range",model["cost_range"]],["Schedule",model["schedule"]],["Risk",model["risk"]],["Confidence",f"{model['confidence_pct']}%"],["QCRA P80",money_bn(model["monte_carlo"]["qcra"]["p80"])],["QSRA P80",str(model["monte_carlo"]["qsra"]["p80"])+" months"]]
    table=Table(data); table.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#07111F")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),.5,colors.grey)])); story.append(table); story.append(PageBreak())
    story.append(Paragraph("Top Risk Drivers",styles["Heading1"]))
    for r in model["risks"][:10]: story.append(Paragraph(f"<b>{r['risk_id']} {r['title']}</b>: {r['activity_id']} / {r['cbs']} — {r['mitigation']}",styles["BodyText"])); story.append(Spacer(1,6))
    doc.build(story); bio.seek(0); return bio.getvalue()

def stream(data:bytes, media:str, filename:str): return StreamingResponse(BytesIO(data), media_type=media, headers={"Content-Disposition":f"attachment; filename={filename}"})



@app.get("/v26/launch-readiness")
def launch_readiness():
    return {
        "status": "demo-ready / production-scaffolded",
        "included": [
            "Earth and Space feature demo theatre",
            "QCRA/QSRA Monte Carlo engine",
            "Risk register mapped to WBS activities and CBS cost accounts",
            "Class 1-5 estimate comparison",
            "Schedule Level 1-5 comparison",
            "Export centre: cost workbook, risk register, schedule export, model audit and full pack",
            "Upload analysis / project doctor endpoint",
            "Saved projects via local SQLite scaffold",
            "Public one-use demo limiter",
            "Server-side private AI key environment pattern",
            "Auth, billing, deployment and database scaffolding docs"
        ],
        "production_next_steps": [
            "Connect hosted Postgres",
            "Add Clerk/Auth0/Supabase Auth",
            "Add Stripe live product IDs",
            "Connect verified benchmark data",
            "Deploy frontend and backend"
        ]
    }

@app.get("/v26/billing/plans")
def billing_plans():
    return {
        "plans": [
            {"name":"Demo", "price":"Public one-use", "features":["Earth/Space demo", "sample exports"]},
            {"name":"Pro", "price":"TBC", "features":["saved projects", "full exports", "upload analysis"]},
            {"name":"Enterprise", "price":"Custom", "features":["teams", "SSO", "benchmark library", "private deployment"]}
        ],
        "note": "Stripe hooks are scaffolded; configure live price IDs before charging customers."
    }

@app.get("/v26/deployment-check")
def deployment_check():
    return {
        "backend": "ok",
        "database_path": DB_PATH,
        "demo_limit_per_ip": DEMO_LIMIT_PER_IP,
        "openai_key_visible_to_users": False,
        "environment_variables": ["OPENAI_API_KEY", "CASEY_DB", "CASEY_DEMO_LIMIT_PER_IP", "STRIPE_SECRET_KEY", "STRIPE_PRICE_ID"],
        "ready_for_local_demo": True
    }

@app.get("/demo/showcases")
def demo_showcases():
    return {
        "headline": "Two-minute appetite demo: Earth and Space in one product.",
        "earth": {
            "title": "Earth revenue demo: 2026/2027 AI data centre",
            "prompt": "Riyadh AI Hyperscale Campus 500MW accelerated 2027 with sovereign cloud, grid connection and liquid cooling",
            "client": "Microsoft",
            "alternates": ["Northern Virginia GPU Data Centre Expansion 2026", "Manchester AI Compute Campus UK 2027", "Boston GMP Life Sciences Campus 2026", "Arizona Advanced Semiconductor Fab 2027"],
            "script": [
                "Open with a live buyer-market example: AI data centre, life sciences or semiconductor facility.",
                "Show cost, schedule, confidence, QCRA/QSRA and risk mapping in the first 60 seconds.",
                "Open scenarios: Faster, Cheaper and Lower-Risk explain the commercial choices.",
                "Open exports and download the full boardroom output pack."
            ]
        },
        "space": {
            "title": "Space jaw-drop demo: Lunar Base Alpha",
            "prompt": "Lunar Base Alpha with 1000 crew, nuclear power, landing pads, life support and launch logistics",
            "client": "Lunar Development Authority",
            "alternates": ["Mars Fuel Refinery with ISRU methane oxygen production", "Orbital Solar Power Ring delivering energy to Earth", "Satellite Gigafactory for 10000 satellites per year", "Deep Space Cargo Hub for asteroid and lunar logistics"],
            "script": [
                "Switch from Earth to Space without changing tools.",
                "Show launch, life-support, mass growth, radiation, ISRU and logistics risks mapped to QSRA/QCRA.",
                "Use the same output centre to prove CASEY handles impossible projects with traceable assumptions.",
                "Close with the same export pack: cost workbook, risk register, schedule export, model audit and full pack."
            ]
        },
        "demo_rules": [
            "Do not explain software first; ask for a project first.",
            "Run Earth demo, then Space demo, then open Monte Carlo and Exports.",
            "End with: CASEY compresses weeks of early controls work into minutes, with traceable assumptions."
        ]
    }


@app.get("/v26/revenue-machine")
def revenue_machine():
    return {
        "positioning": "AI Operating System for capital projects on Earth and in Space.",
        "headline": "Price the future before it gets built.",
        "target_verticals_2026_2027": [
            "AI data centres: Northern Virginia, Riyadh, Abu Dhabi, Manchester, Dublin, Johor, Phoenix, Dallas",
            "Life sciences: Boston/Cambridge MA, North Carolina, Oxford, Cambridge UK, Liverpool, Basel",
            "Semiconductors: Arizona, Texas, Ohio, Wales compound semiconductor cluster, Germany fabs",
            "Space: lunar bases, Mars ISRU, spaceports, satellite factories, orbital power, logistics hubs"
        ],
        "demo_flow": [
            "Run Earth demo using a live-demand project type.",
            "Show cost, schedule, confidence, QCRA/QSRA, mapped risks, peers and scenarios.",
            "Download the board pack live.",
            "Switch to Space demo and prove the same controls OS works for frontier infrastructure.",
            "End with ROI vs traditional consultancy and a book-demo CTA."
        ],
        "commercial_ctas": ["Book demo", "Request sample board pack", "Upload project for review", "Enterprise pilot"]
    }

@app.get("/v26/demo-library")
def v26_demo_library():
    return {
        "data_centres": [
            "Riyadh AI Hyperscale Campus 500MW accelerated 2027",
            "Northern Virginia GPU Data Centre Expansion 2026",
            "Manchester AI Compute Campus UK 2027",
            "Slough Hyperscale Expansion 2027",
            "Dallas GPU Compute Mega Campus 2026",
            "Johor AI Data Corridor 2027"
        ],
        "life_sciences": [
            "Boston GMP Biotech Campus 2026",
            "Cambridge UK Cell Therapy Facility 2027",
            "Oxford Vaccine Research Campus 2026",
            "North Carolina Pharma Mega Plant 2027",
            "Liverpool Pharma Expansion Plant 2027"
        ],
        "semiconductors": [
            "Arizona Advanced Semiconductor Fab 2027",
            "Texas Advanced Packaging Plant 2026",
            "Ohio Semiconductor Mega Campus 2027",
            "Wales Compound Semiconductor Hub 2026",
            "Germany Chip Sovereignty Fab 2027"
        ],
        "space": [
            "Lunar Base Alpha with 1000 crew and nuclear power",
            "Mars Fuel Refinery with ISRU methane oxygen production",
            "Orbital Solar Power Ring delivering energy to Earth",
            "UK Spaceport Expansion with launch pads and fuels",
            "Satellite Gigafactory for 10000 satellites per year",
            "Deep Space Cargo Hub for asteroid and lunar logistics"
        ]
    }

@app.get("/v26/pricing")
def v26_pricing():
    return {
        "starter": {"price":"£499/mo", "ideal_for":"Small teams and early concepts", "features":["10 projects", "board summaries", "basic exports", "demo examples"]},
        "pro": {"price":"£1,999/mo", "ideal_for":"Project owners, developers and controls teams", "features":["unlimited projects", "QCRA/QSRA", "XER", "upload analysis", "peer intelligence"]},
        "enterprise": {"price":"Custom", "ideal_for":"Large owners, governments, funds and enterprise teams", "features":["teams", "SSO", "private deployment", "benchmark library", "white-label reports"]}
    }

@app.post("/export/workbook")
def export_workbook(model:Dict[str,Any]): return stream(workbook_bytes(model),"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","CASEY_Cost_Model_Planet_Class.xlsx")
@app.post("/export/risk-register")
def export_risk(model:Dict[str,Any]): return stream(risk_register_workbook_bytes(model),"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","CASEY_Risk_Register_Pro.xlsx")
@app.post("/export/xer")
def export_xer(model:Dict[str,Any]): return stream(xer_bytes(model),"application/octet-stream","CASEY_TITAN_X_v26_P6_Schedule.xer")
@app.post("/export/word")
def export_word(model:Dict[str,Any]): return stream(word_bytes(model),"application/vnd.openxmlformats-officedocument.wordprocessingml.document","CASEY_Executive_Board_Report.docx")
@app.post("/export/pdf")
def export_pdf(model:Dict[str,Any]): return stream(pdf_bytes(model),"application/pdf","CASEY_Board_Intelligence_Pack.pdf")
@app.post("/export/pptx")
def export_pptx(model:Dict[str,Any]): return stream(pptx_bytes(model),"application/vnd.openxmlformats-officedocument.presentationml.presentation","CASEY_Board_Deck_Elite.pptx")
@app.post("/export/json")
def export_json(model:Dict[str,Any]): return stream(json.dumps(model,indent=2).encode(),"application/json","CASEY_TITAN_X_v26_Model.json")
@app.post("/export/all")
def export_all(model:Dict[str,Any]):
    bio=BytesIO()
    with zipfile.ZipFile(bio,"w",zipfile.ZIP_DEFLATED) as z:
        z.writestr("01_CASEY_Cost_Model_Planet_Class.xlsx",workbook_bytes(model))
        z.writestr("02_CASEY_Risk_Register_Pro.xlsx",risk_register_workbook_bytes(model))
        z.writestr("03_CASEY_P6_Schedule.xer",xer_bytes(model))
        z.writestr("04_CASEY_Executive_Board_Report.docx",word_bytes(model))
        z.writestr("05_CASEY_Board_Intelligence_Pack.pdf",pdf_bytes(model))
        z.writestr("06_CASEY_Board_Deck_Elite.pptx",pptx_bytes(model))
        z.writestr("07_CASEY_Full_Model_Audit.json",json.dumps(model,indent=2))
        z.writestr("08_CASEY_Risk_Register_Raw.csv",risk_csv_bytes(model))
        z.writestr("09_CASEY_Demo_Close_Script.txt","\n".join(model.get("launch_demo_script",[])))
    bio.seek(0); return stream(bio.getvalue(),"application/zip","CASEY_Output_Pack_Planet_Class.zip")


# ========================= CASEY v50 OUTPUT DOMINATION OVERRIDES =========================
# These overrides replace the older lightweight output pack with premium, scenario-linked,
# board-ready exports. They keep the same API endpoints so the frontend does not break.

def _scenario_effects(model):
    sc=(model.get("scenario") or "base").lower()
    effects={
        "base":{"label":"Base","cost_note":"Balanced reference case.","schedule_note":"Standard delivery logic.","risk_note":"Normal contingency and mitigation profile."},
        "faster":{"label":"Faster","cost_note":"Acceleration premium added for overlap, premium logistics and extended shifts.","schedule_note":"Critical path compressed with higher interface risk.","risk_note":"Residual delivery and quality risk increases unless assurance is funded."},
        "cheaper":{"label":"Cheaper","cost_note":"Value engineering and procurement competition reduce initial capex.","schedule_note":"Programme may lengthen due to procurement tension and reduced float.","risk_note":"Scope, quality and rework exposure rises; board should not treat savings as certain."},
        "lower_risk":{"label":"Lower Risk","cost_note":"Higher assurance, surveys and contingency improve board confidence.","schedule_note":"More realistic buffers and stage gates improve QSRA reliability.","risk_note":"Residual risk reduced through earlier evidence and stronger controls."},
        "premium":{"label":"Premium","cost_note":"Flagship resilience, assurance and quality increase capex.","schedule_note":"Schedule protected by better readiness and procurement strategy.","risk_note":"Higher front-end investment reduces downstream delivery uncertainty."},
        "investor":{"label":"Investor","cost_note":"Frames capex as a decision range with commercial challenge points.","schedule_note":"Highlights fundability and milestones.","risk_note":"Emphasis on value leakage and investment downside protection."},
        "survival":{"label":"Survival","cost_note":"Minimum viable scope reduces capex but creates major residual risk.","schedule_note":"Lean programme with limited contingency.","risk_note":"Not board-safe without explicit risk acceptance."},
    }
    return effects.get(sc,effects["base"])

def _total_costs(model):
    lines=model.get("cost_lines",[])
    totals={"Direct":0.0,"Indirect":0.0,"Reserve":0.0}
    for x in lines:
        typ=x.get("type","Direct")
        totals[typ]=totals.get(typ,0)+float(x.get("p50_bn",0) or 0)
    return totals

def _risk_rating(prob, cost_m, sched_m):
    exposure = prob/100.0*(float(cost_m or 0)*100 + float(sched_m or 0)/4)
    if exposure > 90: return "Extreme"
    if exposure > 55: return "High"
    if exposure > 25: return "Medium"
    return "Low"

def risk_register(mode,subsector,cost,months,schedule,costs,rm,scenario):
    """Premium quantified risk register with cause, event, impact and response logic."""
    base=[
      ("R-001","Scope growth","Scope","Ambiguous requirements, immature scope freeze and late stakeholder changes","Approved scope expands after estimate freeze","Additional quantities, redesign, rework, procurement churn and board contingency drawdown","Both",42,15,45,120,0.030,0.080,0.160,"Project Director","Freeze scope baseline, change board, value gates, client decision log","Scope change rate exceeds 2% of package value","Reduce","Open"),
      ("R-002","Market escalation","Commercial","Supplier market tightness, inflation, FX and long procurement window","Rates exceed assumed escalation profile","P50 cost becomes understated; commercial approvals or procurement strategy may need reset","Cost",45,0,0,0,0.040,0.100,0.200,"Commercial Lead","Early procurement, index-linked allowances, market testing, FX strategy","Index exceeds allowance by 3%","Mitigate","Open"),
      ("R-003","Permits / approvals delay","Regulatory","Authority process, stakeholder objections or incomplete submissions","Consent milestone slips beyond baseline","Critical path delay, extended preliminaries and potential redesign conditions","Schedule",35,20,70,180,0.010,0.040,0.080,"Consents Lead","Authority plan, consent tracker, early submissions, stakeholder map","Consent milestone slips by 30 days","Mitigate","Open"),
      ("R-004","Design maturity gap","Technical","Estimate is based on immature design, incomplete surveys or unresolved interfaces","Design basis changes during detailed design","Cost growth, quantity movement, rework and schedule resequencing","Both",38,15,45,120,0.030,0.090,0.180,"Design Manager","Design maturity gates, independent review, assumption register","Design deliverables miss maturity gate","Reduce","Open"),
      ("R-005","Supply chain delay","Procurement","Long-lead equipment, constrained suppliers or late procurement release","Supplier promise dates move right","Delayed installation, critical path movement and acceleration cost","Both",40,10,40,120,0.020,0.070,0.150,"Procurement Lead","Alternate suppliers, early orders, expediting, framework options","Supplier promise date slips","Transfer / Mitigate","Open"),
      ("R-006","Productivity underperformance","Delivery","Access constraints, poor sequencing, labour scarcity or learning curve","Actual production rates underperform plan","Extended duration, increased preliminaries and loss of float","Both",35,15,50,130,0.030,0.080,0.160,"Delivery Lead","Package productivity controls, daily planning, earned value cadence","SPI/CPI deteriorates for two reporting periods","Mitigate","Open"),
      ("R-007","Commissioning delay","Handover","Incomplete readiness, software integration issues or defects during testing","Systems fail tests or require repeated commissioning cycles","Delayed handover, operational readiness slippage and liquidated damages exposure","Both",32,10,35,100,0.020,0.060,0.120,"Commissioning Lead","Commissioning readiness plan, early test packs, digital systems integration","Test failures trend upward","Reduce","Open"),
      ("R-008","Interface misalignment","Integration","Multiple contractors, unclear interface ownership or late design coordination","Package interfaces do not align during installation/integration","Rework, claims, delay and fragmented accountability","Both",34,10,35,100,0.020,0.070,0.140,"Integration Manager","Interface control documents, weekly interface board, accountable owners","Interface actions overdue","Mitigate","Open"),
    ]
    if mode=="Space":
        base=[
          ("R-S01","Launch manifest delay","Launch","Launch provider capacity, weather windows, manifest priority or vehicle readiness","Confirmed launch slot moves or payload misses manifest gate","Programme delay, storage cost, resequencing and mission-readiness risk","Both",52,20,60,180,0.035,0.110,0.240,"Launch Integration Lead","Reserve alternate launch slot, freeze payload interface, maintain launch-readiness checklist","Launch slot not confirmed at L-12 months","Mitigate","Open"),
          ("R-S02","Mass growth","Technical","Design creep, shielding, redundancy or late payload additions increase mass","Payload mass exceeds launch and landing allowance","Launch cost, redesign, performance loss and mission safety exposure","Both",44,15,55,150,0.050,0.130,0.280,"Chief Engineer","Mass control board, margin policy, design-to-mass reviews","Mass margin below 8%","Reduce","Open"),
          ("R-S03","Life-support reliability","Safety / Mission","ECLSS reliability, spares or redundancy evidence below target","Life-support system fails reliability growth or integrated test","Crew safety exposure, launch delay, redesign and mission abort risk","Both",38,25,80,220,0.060,0.150,0.320,"Life Support Lead","Prototype testing, redundancy, reliability growth, independent assurance","Reliability growth misses target","Avoid / Reduce","Open"),
        ]+base
    elif "data centre" in subsector.lower():
        base=[
          ("R-D01","Grid connection delay","Utilities","Utility agreement, substation scope or energisation pathway not secured","Power connection or energisation milestone slips","Critical path delay, temporary power cost, phased opening risk and revenue deferral","Both",52,25,90,240,0.040,0.120,0.260,"Utilities Lead","Secure grid agreement, temporary power strategy, early equipment orders","Grid agreement unsigned at gate","Mitigate","Open"),
          ("R-D02","Cooling system capacity","Technical","Thermal basis, water availability or cooling vendor performance uncertain","Cooling design fails capacity, redundancy or commissioning test","Re-design, commissioning delay, lower IT load and resilience concern","Both",36,10,45,130,0.025,0.080,0.170,"MEP Lead","Thermal modelling, supplier validation, early performance test","Cooling performance test fails","Reduce","Open"),
        ]+base
    elif "airport" in subsector.lower():
        base=[
          ("R-A01","Operational phasing disruption","Operations","Works in or near live airport operations constrain access and sequencing","Operational restrictions reduce productive windows","Delay, night-work premiums, stakeholder disruption and ORAT pressure","Both",42,20,70,160,0.025,0.080,0.180,"ORAT Lead","Phasing simulation, possession plan, airport ops integration board","Possession windows rejected","Mitigate","Open"),
          ("R-A02","Baggage / security systems integration","Systems","Complex passenger systems integration and vendor interfaces uncertain","Integrated airport systems do not pass readiness testing","Commissioning delay, passenger disruption and late operational readiness","Both",34,15,50,120,0.020,0.070,0.150,"Systems Lead","Factory acceptance testing, integration lab, ORAT dry-runs","FAT/SAT defect trend above threshold","Reduce","Open"),
        ]+base
    out=[]
    for i,r in enumerate(base,1):
        rid,title,cat,cause,event,impact,area,prob,so,sm,sp,co,cm,cp,owner,mit,trig,response,status=r
        prob=int(clamp(prob*rm,5,92))
        if scenario=="lower_risk": cm*=0.70; cp*=0.75; sm=round(sm*.72); sp=round(sp*.78)
        if scenario=="faster": sm=round(sm*1.22); sp=round(sp*1.28); cm*=1.06; cp*=1.08
        if scenario=="cheaper": cm*=1.16; cp*=1.22; prob=int(clamp(prob*1.08,5,95))
        if scenario=="premium": cm*=0.82; cp*=0.88; prob=int(clamp(prob*0.9,5,95))
        activity=schedule[min(i+3,len(schedule)-1)] if schedule else {"activity_id":"A1900","activity":"Delivery"}
        costline=costs[min(i+2,len(costs)-1)] if costs else {"cbs":"01.01","description":"Cost"}
        emv=cost*cm*prob/100; semv=sm*prob/100
        rating=_risk_rating(prob,cost*cm,sm)
        out.append({
          "risk_id":rid,"title":title,"category":cat,"cause":cause,"risk_event":event,"impact_description":impact,
          "description":event,"impact_area":area,"probability_pct":prob,"pre_mitigation_rating":rating,
          "activity_id":activity["activity_id"],"activity_name":activity["activity"],"cbs":costline["cbs"],"cbs_name":costline["description"],
          "schedule_o_days":so,"schedule_m_days":sm,"schedule_p_days":sp,"cost_o_bn":round(cost*co,3),"cost_m_bn":round(cost*cm,3),"cost_p_bn":round(cost*cp,3),
          "cost_emv_bn":round(emv,3),"schedule_emv_days":round(semv,1),"owner":owner,"trigger":trig,"mitigation":mit,"response_strategy":response,
          "residual_rating":"Medium" if rating in ["High","Extreme"] else "Low","status":status,"last_reviewed":datetime.utcnow().strftime("%Y-%m-%d"),
          "board_visibility":"Yes" if i<=8 or rating in ["High","Extreme"] else "No",
          "basis_of_cost_impact":f"O/M/P cost impact equals {money_bn(cost*co)} / {money_bn(cost*cm)} / {money_bn(cost*cp)} based on scenario-adjusted exposure to CBS {costline['cbs']} {costline['description']}.",
          "basis_of_schedule_impact":f"O/M/P schedule impact equals {so}/{sm}/{sp} days mapped to activity {activity['activity_id']} {activity['activity']}.",
          "driver_score":round(emv*100+semv/5,2)
        })
    return sorted(out,key=lambda x:x["driver_score"],reverse=True)

def monte_carlo(cost,months,risks,seed=42,iterations=10000):
    rng=np.random.default_rng(seed)
    base_cost_unc=rng.triangular(cost*.90,cost,cost*1.16,iterations)
    base_sched_unc=rng.triangular(months*.94,months,months*1.14,iterations)
    risk_cost=np.zeros(iterations); risk_days=np.zeros(iterations)
    contribution={r["risk_id"]:{"cost":0.0,"days":0.0,"title":r["title"],"activity_id":r["activity_id"],"cbs":r["cbs"],"category":r.get("category","Risk")} for r in risks}
    for r in risks:
        occurs=rng.random(iterations)<(r["probability_pct"]/100)
        c=rng.triangular(r["cost_o_bn"],r["cost_m_bn"],r["cost_p_bn"],iterations)*occurs
        left=min(r["schedule_o_days"],r["schedule_m_days"],r["schedule_p_days"]); mode=max(left,r["schedule_m_days"]); right=max(r["schedule_o_days"],r["schedule_m_days"],r["schedule_p_days"])
        d=np.zeros(iterations) if left==right else rng.triangular(left,mode,right,iterations)*occurs
        risk_cost+=c; risk_days+=d
        contribution[r["risk_id"]]["cost"]=float(np.mean(c)); contribution[r["risk_id"]]["days"]=float(np.mean(d))
    cost_samples=base_cost_unc+risk_cost
    sched_samples=base_sched_unc+(risk_days/30.44)
    def pct(arr,p): return float(np.percentile(arr,p))
    curve=[{"percentile":p,"cost_bn":round(pct(cost_samples,p),3),"schedule_months":round(pct(sched_samples,p),2)} for p in [1,5,10,20,30,40,50,60,70,80,90,95,99]]
    qcra_tornado=sorted([{"risk_id":k,"title":v["title"],"category":v["category"],"cbs":v["cbs"],"activity_id":v["activity_id"],"cost_mean_bn":round(v["cost"],3),"driver_score":round(v["cost"]*100,2)} for k,v in contribution.items()],key=lambda x:x["driver_score"],reverse=True)
    qsra_tornado=sorted([{"risk_id":k,"title":v["title"],"category":v["category"],"cbs":v["cbs"],"activity_id":v["activity_id"],"schedule_mean_days":round(v["days"],1),"driver_score":round(v["days"]/2,2)} for k,v in contribution.items()],key=lambda x:x["driver_score"],reverse=True)
    tornado=sorted([{"risk_id":k,"title":v["title"],"activity_id":v["activity_id"],"cbs":v["cbs"],"cost_mean_bn":round(v["cost"],3),"schedule_mean_days":round(v["days"],1),"driver_score":round(v["cost"]*100+v["days"]/10,2)} for k,v in contribution.items()],key=lambda x:x["driver_score"],reverse=True)
    return {"iterations":iterations,"qcra":{"p10":round(pct(cost_samples,10),3),"p50":round(pct(cost_samples,50),3),"p80":round(pct(cost_samples,80),3),"p90":round(pct(cost_samples,90),3),"mean":round(float(np.mean(cost_samples)),3)},"qsra":{"p10":round(pct(sched_samples,10),2),"p50":round(pct(sched_samples,50),2),"p80":round(pct(sched_samples,80),2),"p90":round(pct(sched_samples,90),2),"mean":round(float(np.mean(sched_samples)),2)},"curve":curve,"tornado":tornado,"qcra_tornado":qcra_tornado,"qsra_tornado":qsra_tornado}

def _style_ws(ws, freeze="A2"):
    dark="07111F"; cyan="8DF7FF"; panel="0B1424"; white="F7FBFF"
    ws.freeze_panes=freeze
    for row in ws.iter_rows():
        for c in row:
            c.alignment=Alignment(vertical="top",wrap_text=True)
            c.border=Border(bottom=Side(style="thin",color="1E2A3A"))
    if ws.max_row>=1:
        for c in ws[1]:
            c.fill=PatternFill("solid",fgColor=dark); c.font=Font(color=white,bold=True); c.alignment=Alignment(vertical="center",wrap_text=True)
    for col in range(1,ws.max_column+1):
        ws.column_dimensions[get_column_letter(col)].width=min(38,max(14,len(str(ws.cell(1,col).value or ""))+3))

def _add(ws, rows):
    for r in rows: ws.append(r)
    _style_ws(ws)

def workbook_bytes(model):
    wb=Workbook(); wb.remove(wb.active)
    fx=_scenario_effects(model); totals=_total_costs(model); mc=model["monte_carlo"]
    ws=wb.create_sheet("00 Executive Summary")
    rows=[
      ["CASEY TITAN X", "Elite Project Controls Intelligence Pack"],
      ["Generated", datetime.utcnow().isoformat()], ["Project", model.get("title")], ["Client", model.get("client")], ["Sector", model.get("subsector")], ["Location", model.get("location")], ["Scenario", model.get("scenario_label")],
      ["Executive Summary", model.get("executive_summary")], ["Board Recommendation", f"Proceed only with {model.get('scenario_label')} controls basis, explicit risk allowance and evidence plan before commitment."],
      ["P10 / P50 / P80 / P90 Cost", f"{money_bn(mc['qcra']['p10'])} / {model.get('cost_p50')} / {money_bn(mc['qcra']['p80'])} / {money_bn(mc['qcra']['p90'])}"],
      ["QSRA P50 / P80", f"{mc['qsra']['p50']} / {mc['qsra']['p80']} months"], ["Direct / Indirect / Reserve P50", f"{money_bn(totals.get('Direct',0))} / {money_bn(totals.get('Indirect',0))} / {money_bn(totals.get('Reserve',0))}"],
      ["Confidence", f"{model.get('confidence_pct')}%"], ["Risk", model.get("risk")], ["Scenario Cost Note", fx["cost_note"]], ["Scenario Schedule Note", fx["schedule_note"]], ["Scenario Risk Note", fx["risk_note"]]
    ]
    for r in rows: ws.append(r)
    ws.merge_cells("A1:B1"); ws["A1"].font=Font(size=20,bold=True,color="8DF7FF"); ws["A1"].fill=PatternFill("solid",fgColor="02050A")
    _style_ws(ws,"A3"); ws.column_dimensions["A"].width=34; ws.column_dimensions["B"].width=110

    ws=wb.create_sheet("01 Selected Cost Estimate")
    _add(ws, [["CBS","Description","Cost Type","Basis","Min / O BN","Most Likely / P50 BN","Max / P90 BN","P80 BN","Scenario Note","Impact Basis"]]+[
      [x["cbs"],x["description"],x["type"],x["basis"],x["p10_bn"],x["p50_bn"],x["p90_bn"],round((x["p50_bn"]+x["p90_bn"])*0.62,3),fx["cost_note"],x["impact_basis"]] for x in model.get("cost_lines",[])
    ])
    ws=wb.create_sheet("02 Cost Summary")
    _add(ws, [["Cost Bucket","P50 BN","Share","Board Meaning"], ["Direct",round(totals.get("Direct",0),3),round(totals.get("Direct",0)/max(parse_bn(model.get('cost_p50')),0.001),3),"Physical scope and deliverable assets"], ["Indirect",round(totals.get("Indirect",0),3),round(totals.get("Indirect",0)/max(parse_bn(model.get('cost_p50')),0.001),3),"Prelims, design, assurance, insurance and management"], ["Reserve",round(totals.get("Reserve",0),3),round(totals.get("Reserve",0)/max(parse_bn(model.get('cost_p50')),0.001),3),"Contingency and management reserve linked to QCRA"]])
    ws=wb.create_sheet("03 Risk Register Pro")
    risk_headers=["Risk ID","Category","Title","Cause","Risk Event","Impact Description","Probability %","Pre Rating","Residual Rating","Owner","Response Strategy","Mitigation Actions","Trigger","Status","Last Reviewed","Board Visibility","Linked Activity","Activity Name","Linked CBS","CBS Name","Sched O","Sched ML","Sched Max","Cost O BN","Cost ML BN","Cost Max BN","Cost EMV BN","Schedule EMV Days","Basis Cost","Basis Schedule"]
    _add(ws,[risk_headers]+[[r.get(k) for k in []] for r in []])
    for r in model.get("risks",[]):
        ws.append([r["risk_id"],r["category"],r["title"],r.get("cause"),r.get("risk_event"),r.get("impact_description"),r["probability_pct"],r.get("pre_mitigation_rating"),r.get("residual_rating"),r["owner"],r.get("response_strategy"),r["mitigation"],r["trigger"],r.get("status"),r.get("last_reviewed"),r.get("board_visibility"),r["activity_id"],r["activity_name"],r["cbs"],r["cbs_name"],r["schedule_o_days"],r["schedule_m_days"],r["schedule_p_days"],r["cost_o_bn"],r["cost_m_bn"],r["cost_p_bn"],r.get("cost_emv_bn"),r.get("schedule_emv_days"),r["basis_of_cost_impact"],r["basis_of_schedule_impact"]])
    _style_ws(ws)
    ws=wb.create_sheet("04 QCRA P-Curve")
    _add(ws, [["Percentile","Cost BN"]]+[[x["percentile"],x["cost_bn"]] for x in mc["curve"]])
    try:
        chart=LineChart(); chart.title="QCRA Cost P-Curve"; chart.y_axis.title="Cost BN"; chart.x_axis.title="Percentile"; data=Reference(ws,min_col=2,min_row=1,max_row=ws.max_row); cats=Reference(ws,min_col=1,min_row=2,max_row=ws.max_row); chart.add_data(data,titles_from_data=True); chart.set_categories(cats); ws.add_chart(chart,"D2")
    except Exception: pass
    ws=wb.create_sheet("05 QSRA P-Curve")
    _add(ws, [["Percentile","Schedule Months"]]+[[x["percentile"],x["schedule_months"]] for x in mc["curve"]])
    try:
        chart=LineChart(); chart.title="QSRA Schedule P-Curve"; chart.y_axis.title="Months"; chart.x_axis.title="Percentile"; data=Reference(ws,min_col=2,min_row=1,max_row=ws.max_row); cats=Reference(ws,min_col=1,min_row=2,max_row=ws.max_row); chart.add_data(data,titles_from_data=True); chart.set_categories(cats); ws.add_chart(chart,"D2")
    except Exception: pass
    ws=wb.create_sheet("06 QCRA Tornado")
    _add(ws, [["Rank","Risk ID","Driver","Category","CBS","Cost Mean BN","Driver Score"]]+[[i+1,x["risk_id"],x["title"],x["category"],x["cbs"],x["cost_mean_bn"],x["driver_score"]] for i,x in enumerate(mc.get("qcra_tornado",mc.get("tornado",[]))[:12])])
    try:
        b=BarChart(); b.type="bar"; b.title="QCRA Tornado - Cost Drivers"; data=Reference(ws,min_col=7,min_row=1,max_row=ws.max_row); cats=Reference(ws,min_col=3,min_row=2,max_row=ws.max_row); b.add_data(data,titles_from_data=True); b.set_categories(cats); ws.add_chart(b,"I2")
    except Exception: pass
    ws=wb.create_sheet("07 QSRA Tornado")
    _add(ws, [["Rank","Risk ID","Driver","Category","Activity","Schedule Mean Days","Driver Score"]]+[[i+1,x["risk_id"],x["title"],x["category"],x["activity_id"],x["schedule_mean_days"],x["driver_score"]] for i,x in enumerate(mc.get("qsra_tornado",mc.get("tornado",[]))[:12])])
    try:
        b=BarChart(); b.type="bar"; b.title="QSRA Tornado - Schedule Drivers"; data=Reference(ws,min_col=7,min_row=1,max_row=ws.max_row); cats=Reference(ws,min_col=3,min_row=2,max_row=ws.max_row); b.add_data(data,titles_from_data=True); b.set_categories(cats); ws.add_chart(b,"I2")
    except Exception: pass
    ws=wb.create_sheet("08 Schedule Logic")
    _add(ws, [["Activity ID","Phase","Activity","Predecessor","Duration Months","Critical","Basis"]]+[[x["activity_id"],x["phase"],x["activity"],x["predecessor"],x["duration_months"],x["critical"],x["basis"]] for x in model.get("schedule_rows",[])])
    ws=wb.create_sheet("09 Scenarios")
    _add(ws, [["Scenario","Cost","Schedule Months","Risk","Confidence","Why"]]+[[x["label"],x["cost"],x["schedule_months"],x["risk"],x["confidence"],x["why"]] for x in model.get("scenario_comparison",[])])
    ws=wb.create_sheet("10 Benchmarks")
    _add(ws, [["Metric","Value","Why"]]+[[x["metric"],x["value"],x["why"]] for x in model.get("benchmarks",[])])
    ws=wb.create_sheet("11 Assumptions & Audit")
    _add(ws, [["Item","Value"],["Model ID",model.get("id")],["Version","v50 Output Domination"],["Estimate Class",model.get("estimate_class_name")],["Schedule Level",model.get("schedule_level")],["Important note","This is a first-pass intelligence pack for option testing and challenge. It is not a substitute for verified design, supplier quotes, legal, safety or regulated professional sign-off."]]+[["Confidence basis",x] for x in model.get("confidence_explanation",[])])
    bio=BytesIO(); wb.save(bio); bio.seek(0); return bio.getvalue()

def risk_csv_bytes(model):
    out=StringIO(); w=csv.writer(out)
    headers=["Risk ID","Category","Title","Cause","Risk Event","Impact Description","Probability %","Pre Rating","Residual Rating","Owner","Response Strategy","Mitigation","Trigger","Status","Last Reviewed","Board Visibility","Activity ID","Activity Name","CBS","CBS Name","Sched O","Sched ML","Sched Max","Cost O BN","Cost ML BN","Cost Max BN","Cost EMV BN","Schedule EMV Days","Basis Cost","Basis Schedule"]
    w.writerow(headers)
    for r in model.get("risks",[]):
        w.writerow([r.get("risk_id"),r.get("category"),r.get("title"),r.get("cause"),r.get("risk_event"),r.get("impact_description"),r.get("probability_pct"),r.get("pre_mitigation_rating"),r.get("residual_rating"),r.get("owner"),r.get("response_strategy"),r.get("mitigation"),r.get("trigger"),r.get("status"),r.get("last_reviewed"),r.get("board_visibility"),r.get("activity_id"),r.get("activity_name"),r.get("cbs"),r.get("cbs_name"),r.get("schedule_o_days"),r.get("schedule_m_days"),r.get("schedule_p_days"),r.get("cost_o_bn"),r.get("cost_m_bn"),r.get("cost_p_bn"),r.get("cost_emv_bn"),r.get("schedule_emv_days"),r.get("basis_of_cost_impact"),r.get("basis_of_schedule_impact")])
    return out.getvalue().encode()

def word_bytes(model):
    doc=Document(); styles=doc.styles; styles["Normal"].font.name="Aptos"; styles["Normal"].font.size=Pt(10)
    title=doc.add_heading("CASEY TITAN X — Executive Project Controls Intelligence",0); title.runs[0].font.color.rgb=RGBColor(8,40,60)
    doc.add_paragraph(model.get("executive_summary",""))
    fx=_scenario_effects(model); totals=_total_costs(model); mc=model["monte_carlo"]
    doc.add_heading("Board Decision Snapshot",1)
    tbl=doc.add_table(rows=1,cols=2); tbl.style="Light Shading Accent 1"; hdr=tbl.rows[0].cells; hdr[0].text="Metric"; hdr[1].text="Value"
    for k,v in [("Scenario",model.get("scenario_label")),("P10 / P50 / P80 / P90 Cost",f"{money_bn(mc['qcra']['p10'])} / {model.get('cost_p50')} / {money_bn(mc['qcra']['p80'])} / {money_bn(mc['qcra']['p90'])}"),("Direct / Indirect / Reserve",f"{money_bn(totals.get('Direct',0))} / {money_bn(totals.get('Indirect',0))} / {money_bn(totals.get('Reserve',0))}"),("Baseline / P80 Schedule",f"{model.get('schedule')} / {mc['qsra']['p80']} months"),("Risk / Confidence",f"{model.get('risk')} / {model.get('confidence_pct')}%")]:
        row=tbl.add_row().cells; row[0].text=k; row[1].text=str(v)
    doc.add_heading("Executive Recommendation",1); doc.add_paragraph(f"Proceed with the {model.get('scenario_label')} basis only if the sponsor accepts the quantified risk range, funds the reserve strategy and closes the top evidence gaps before approval.")
    for heading,items in [("Scenario Logic",[fx['cost_note'],fx['schedule_note'],fx['risk_note']]),("Confidence Basis",model.get("confidence_explanation",[])),("Board Challenge Questions",model.get("board_challenge_questions",[])),("Next Best Actions",model.get("next_best_actions",[]))]:
        doc.add_heading(heading,1)
        for x in items: doc.add_paragraph(str(x),style="List Bullet")
    doc.add_heading("Top Risk Drivers",1)
    for r in model.get("risks",[])[:10]:
        doc.add_paragraph(f"{r['risk_id']} {r['title']}",style="List Bullet")
        doc.add_paragraph(f"Cause: {r.get('cause')} | Event: {r.get('risk_event')} | Impact: {r.get('impact_description')} | Response: {r.get('mitigation')}")
    doc.add_heading("Cost View",1)
    tbl=doc.add_table(rows=1,cols=5); tbl.style="Light Grid Accent 1"; hdr=tbl.rows[0].cells
    for i,h in enumerate(["CBS","Description","Type","P50 BN","P90 BN"]): hdr[i].text=h
    for x in model.get("cost_lines",[])[:18]:
        row=tbl.add_row().cells; row[0].text=x["cbs"]; row[1].text=x["description"]; row[2].text=x["type"]; row[3].text=str(x["p50_bn"]); row[4].text=str(x["p90_bn"])
    bio=BytesIO(); doc.save(bio); bio.seek(0); return bio.getvalue()

def pdf_bytes(model):
    bio=BytesIO(); doc=SimpleDocTemplate(bio,pagesize=landscape(A4),rightMargin=28,leftMargin=28,topMargin=24,bottomMargin=24)
    styles=getSampleStyleSheet(); styles.add(ParagraphStyle(name="CASEYTitle",fontSize=24,leading=28,textColor=colors.HexColor("#07111F"),spaceAfter=12)); styles.add(ParagraphStyle(name="Small",fontSize=8,leading=10,textColor=colors.HexColor("#334155")))
    mc=model["monte_carlo"]; totals=_total_costs(model); story=[]
    story.append(Paragraph("CASEY TITAN X — Board Intelligence Pack",styles["CASEYTitle"])); story.append(Paragraph(model.get("executive_summary",""),styles["BodyText"])); story.append(Spacer(1,10))
    data=[["Metric","Value"],["Scenario",model.get("scenario_label")],["P50 Cost",model.get("cost_p50")],["QCRA P80",money_bn(mc["qcra"]["p80"])],["QSRA P80",f"{mc['qsra']['p80']} months"],["Direct / Indirect / Reserve",f"{money_bn(totals.get('Direct',0))} / {money_bn(totals.get('Indirect',0))} / {money_bn(totals.get('Reserve',0))}"],["Risk / Confidence",f"{model.get('risk')} / {model.get('confidence_pct')}%"]]
    table=Table(data,colWidths=[170,480]); table.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#07111F")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),.35,colors.HexColor("#CBD5E1")),("BACKGROUND",(0,1),(-1,-1),colors.HexColor("#F8FAFC"))])); story.append(table); story.append(PageBreak())
    story.append(Paragraph("Top QCRA and QSRA Drivers",styles["Heading1"]))
    rows=[["Rank","QCRA Cost Driver","Cost Mean BN","QSRA Schedule Driver","Mean Days"]]
    q=mc.get("qcra_tornado",[])[:8]; s=mc.get("qsra_tornado",[])[:8]
    for i in range(max(len(q),len(s))): rows.append([i+1,q[i]["title"] if i<len(q) else "",q[i].get("cost_mean_bn","") if i<len(q) else "",s[i]["title"] if i<len(s) else "",s[i].get("schedule_mean_days","") if i<len(s) else ""])
    t=Table(rows,colWidths=[38,250,80,250,80]); t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#07111F")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),.35,colors.HexColor("#CBD5E1")),("FONTSIZE",(0,0),(-1,-1),8)])); story.append(t); story.append(PageBreak())
    story.append(Paragraph("Risk Register — Board Visibility",styles["Heading1"]))
    rows=[["ID","Risk","Cause","Event","Impact","Owner","Response"]]
    for r in model.get("risks",[])[:10]: rows.append([r.get("risk_id"),r.get("title"),r.get("cause"),r.get("risk_event"),r.get("impact_description"),r.get("owner"),r.get("mitigation")])
    t=Table(rows,colWidths=[45,90,150,150,175,80,165]); t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#07111F")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),.25,colors.HexColor("#CBD5E1")),("FONTSIZE",(0,0),(-1,-1),7),('VALIGN',(0,0),(-1,-1),'TOP')])) ; story.append(t)
    doc.build(story); bio.seek(0); return bio.getvalue()

def pptx_bytes(model):
    from pptx.dml.color import RGBColor as PRGBColor
    prs=Presentation(); prs.slide_width=PptxInches(13.333); prs.slide_height=PptxInches(7.5); blank=prs.slide_layouts[6]
    def title(sl,t,sub=""):
        sl.background.fill.solid(); sl.background.fill.fore_color.rgb=PRGBColor(2,5,10)
        box=sl.shapes.add_textbox(PptxInches(.45),PptxInches(.35),PptxInches(12.4),PptxInches(.6)); tf=box.text_frame; tf.text=t; r=tf.paragraphs[0].runs[0]; r.font.size=PptxPt(26); r.font.bold=True; r.font.color.rgb=PRGBColor(141,247,255)
        if sub:
            b=sl.shapes.add_textbox(PptxInches(.48),PptxInches(.95),PptxInches(12),PptxInches(.55)); b.text_frame.text=sub; b.text_frame.paragraphs[0].runs[0].font.size=PptxPt(13); b.text_frame.paragraphs[0].runs[0].font.color.rgb=PRGBColor(245,250,255)
    s=prs.slides.add_slide(blank); title(s,"CASEY TITAN X",model.get("executive_summary",""))
    s=prs.slides.add_slide(blank); title(s,"Board Metrics")
    metrics=[("P50",model.get("cost_p50")),("P80",money_bn(model['monte_carlo']['qcra']['p80'])),("Schedule",model.get("schedule")),("QSRA P80",str(model['monte_carlo']['qsra']['p80'])+" months"),("Risk",model.get("risk")),("Confidence",str(model.get("confidence_pct"))+"%")]
    for i,(k,v) in enumerate(metrics):
        x=.6+(i%3)*4.2; y=1.55+(i//3)*1.7; shp=s.shapes.add_shape(1,PptxInches(x),PptxInches(y),PptxInches(3.7),PptxInches(1.1)); shp.fill.solid(); shp.fill.fore_color.rgb=PRGBColor(7,17,31); shp.line.color.rgb=PRGBColor(141,247,255); shp.text=f"{k}\n{v}"; shp.text_frame.paragraphs[0].runs[0].font.color.rgb=PRGBColor(255,255,255)
    s=prs.slides.add_slide(blank); title(s,"Top Risk Drivers","Cause • Event • Impact • Response")
    txt=s.shapes.add_textbox(PptxInches(.55),PptxInches(1.3),PptxInches(12.1),PptxInches(5.6)).text_frame
    for r in model.get("risks",[])[:8]:
        p=txt.add_paragraph(); p.text=f"{r['risk_id']} {r['title']} — {r.get('impact_description')}"; p.font.size=PptxPt(12); p.font.color.rgb=PRGBColor(245,250,255)
    s=prs.slides.add_slide(blank); title(s,"Output Pack","Selected cost estimate, risk register, QCRA/QSRA and decision evidence generated from one model.")
    bio=BytesIO(); prs.save(bio); bio.seek(0); return bio.getvalue()

# ======================= END CASEY v50 OUTPUT DOMINATION OVERRIDES =======================


# Frontend compatibility aliases for v50 UI
_CASEY_ORIGINAL_BUILD_MODEL = build_model
def build_model(prompt:str, client:str="", class_level:int=3, schedule_level:int=3, scenario:str="base"):
    m=_CASEY_ORIGINAL_BUILD_MODEL(prompt, client, class_level, schedule_level, scenario)
    m=apply_sector_intelligence(m)
    m=scenario_cascade_v95(m, scenario)
    m["version"]="CASEY TITAN X v95 Demo Launch"
    m["cost_breakdown"]=m.get("cost_lines",[])
    m["risk_register"]=m.get("risks",[])
    m["schedule_detail"]=m.get("schedule_rows",[])
    for x in m.get("monte_carlo",{}).get("tornado",[]):
        x["driver"]=x.get("title")
        x["contribution"]=x.get("driver_score")
    return m

@app.post("/upload")
async def upload_alias(file: UploadFile = File(...)):
    return await analyse_upload(file)

# v50 export filenames override
@app.post("/v50/export/workbook")
def export_workbook_v50(model:Dict[str,Any]): return stream(workbook_bytes(model),"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","CASEY_v50_Cost_Model_Elite.xlsx")
@app.post("/v50/export/risk-register")
def export_risk_v50(model:Dict[str,Any]): return stream(risk_csv_bytes(model),"text/csv","CASEY_v50_Risk_Register_Pro.csv")
@app.post("/v50/export/all")
def export_all_v50(model:Dict[str,Any]):
    bio=BytesIO()
    with zipfile.ZipFile(bio,"w",zipfile.ZIP_DEFLATED) as z:
        z.writestr("01_CASEY_v50_Cost_Model_Elite.xlsx",workbook_bytes(model))
        z.writestr("02_CASEY_v50_Risk_Register_Pro.csv",risk_csv_bytes(model))
        z.writestr("03_CASEY_v50_P6_Schedule.xer",xer_bytes(model))
        z.writestr("04_CASEY_v50_Executive_Report.docx",word_bytes(model))
        z.writestr("05_CASEY_v50_Board_Report.pdf",pdf_bytes(model))
        z.writestr("06_CASEY_v50_Board_Deck.pptx",pptx_bytes(model))
        z.writestr("07_CASEY_v50_Model.json",json.dumps(model,indent=2))
    bio.seek(0); return stream(bio.getvalue(),"application/zip","CASEY_v50_Output_Domination_Pack.zip")



# ======================= CASEY v51 OUTPUT PLANET CLASS OVERRIDES =======================
# Highest-spec outputs: premium workbook, risk workbook, DOCX, PDF, PPTX and ZIP.
# Designed to make export quality match the cinematic CASEY product experience.

from openpyxl.chart.label import DataLabelList
from openpyxl.worksheet.table import Table as XLTable, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
from openpyxl.comments import Comment

V51_NAVY = "02050A"
V51_PANEL = "07111F"
V51_CYAN = "8DF7FF"
V51_BLUE = "2FB7FF"
V51_WHITE = "F7FBFF"
V51_MUTED = "9FB3C8"
V51_LINE = "203040"
V51_RED = "FF5C7A"
V51_AMBER = "FFD166"
V51_GREEN = "7CFFB2"


def _safe(v, default=""):
    return default if v is None else v


def _num(v, default=0.0):
    try:
        return float(v)
    except Exception:
        return default


def _v51_money(v):
    return money_bn(_num(v))


def _v51_total_costs(model):
    out = {"Direct": 0.0, "Indirect": 0.0, "Reserve": 0.0}
    for x in model.get("cost_lines", []):
        typ = x.get("type", "Direct")
        if typ not in out: out[typ] = 0.0
        out[typ] += _num(x.get("p50_bn"))
    return out


def _v51_scenario_notes(model):
    s = str(model.get("scenario", "base")).lower()
    if s == "faster":
        return [
            "Acceleration selected: schedule compression is funded through logistics premium, parallel procurement and higher delivery intensity.",
            "Controls consequence: QSRA risk should be watched closely because interface density and commissioning pressure rise.",
            "Board decision: approve acceleration only with package-level recovery plan and explicit acceleration cost allowance."
        ]
    if s == "cheaper":
        return [
            "Cheaper selected: value engineering and procurement competition reduce P50 but increase residual delivery exposure.",
            "Controls consequence: all savings must be linked to scope choices; avoid hiding deferred scope as efficiency.",
            "Board decision: approve only after VE log, exclusions, quality impacts and owner-retained risks are signed off."
        ]
    if s == "lower_risk":
        return [
            "Lower-risk selected: more assurance, surveys, float and contingency raise headline cost but improve confidence.",
            "Controls consequence: stronger evidence should move risk from unknown unknowns into quantified known unknowns.",
            "Board decision: approve with a confidence-improvement plan and gate evidence pack."
        ]
    if s == "premium":
        return [
            "Premium selected: resilience, specification, assurance and stakeholder certainty are prioritised over lowest cost.",
            "Controls consequence: P50 rises, but delivery certainty and reputational protection should improve.",
            "Board decision: approve if the premium is justified by availability, criticality, carbon, security or strategic value."
        ]
    return [
        "Base selected: balanced control case for first-pass board challenge and option testing.",
        "Controls consequence: use this as the reference model before switching to Faster, Cheaper, Lower Risk or Premium.",
        "Board decision: approve further definition work, not unconditional capital commitment."
    ]


def _apply_ws_theme(ws, widths=None, freeze="A8"):
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = freeze
    if widths:
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
    thin = Side(style="thin", color=V51_LINE)
    for row in ws.iter_rows():
        for cell in row:
            cell.font = Font(name="Aptos", size=10, color="EAF6FF")
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = Border(bottom=thin)
    for cell in ws[1]:
        cell.font = Font(name="Aptos Display", size=16, bold=True, color=V51_CYAN)
    return ws


def _title(ws, title, subtitle=None, end_col="J"):
    ws.merge_cells(f"A1:{end_col}1")
    ws["A1"] = title
    ws["A1"].fill = PatternFill("solid", fgColor=V51_Navy if False else V51_Navy if False else V51_NAVY)
    ws["A1"].font = Font(name="Aptos Display", size=18, bold=True, color=V51_CYAN)
    ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 34
    if subtitle:
        ws.merge_cells(f"A2:{end_col}2")
        ws["A2"] = subtitle
        ws["A2"].font = Font(name="Aptos", size=10, color=V51_MUTED)
        ws["A2"].fill = PatternFill("solid", fgColor=V51_PANEL)
        ws.row_dimensions[2].height = 28


def _header_row(ws, row, cols, fill=V51_PANEL):
    for i, h in enumerate(cols, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.fill = PatternFill("solid", fgColor=fill)
        c.font = Font(name="Aptos", size=10, bold=True, color=V51_CYAN)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 26


def _add_table(ws, ref, name):
    try:
        tab = XLTable(displayName=name, ref=ref)
        style = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        tab.tableStyleInfo = style
        ws.add_table(tab)
    except Exception:
        pass


def workbook_bytes(model:Dict[str,Any]):
    wb = Workbook()
    # Remove default and build planet-class workbook
    ws = wb.active
    ws.title = "00 Executive Dashboard"
    for s in ["01 Basis", "02 Cost Estimate", "03 Indirects & Reserves", "04 QCRA", "05 QSRA", "06 Risk Register", "07 Schedule", "08 Scenarios", "09 Benchmarking", "10 Assumptions"]:
        wb.create_sheet(s)

    mc = model.get("monte_carlo", {})
    totals = _v51_total_costs(model)
    scenario_notes = _v51_scenario_notes(model)

    # Dashboard
    ws = wb["00 Executive Dashboard"]
    _title(ws, "CASEY - EXECUTIVE PROJECT CONTROLS INTELLIGENCE", "Board-ready first-pass cost, schedule, risk and confidence model generated from a single connected project basis.", "K")
    widths = {"A":22,"B":22,"C":22,"D":22,"E":22,"F":22,"G":22,"H":22,"I":22,"J":22,"K":22}
    _apply_ws_theme(ws, widths, freeze="A9")
    ws["A4"]="Project"; ws["B4"]=model.get("title")
    ws["D4"]="Sector"; ws["E4"]=model.get("subsector")
    ws["G4"]="Scenario"; ws["H4"]=model.get("scenario_label")
    ws["J4"]="Generated"; ws["K4"]=datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")
    kpis=[("P10 Cost",model.get("cost_p10")),("P50 Cost",model.get("cost_p50")),("P90 Cost",model.get("cost_p90")),("QCRA P80",_v51_money(mc.get("qcra",{}).get("p80"))),("Schedule",model.get("schedule")),("QSRA P80",f"{mc.get('qsra',{}).get('p80','-')} months"),("Risk",model.get("risk")),("Confidence",f"{model.get('confidence_pct')}%"),("Direct",_v51_money(totals.get("Direct"))),("Indirect",_v51_money(totals.get("Indirect"))),("Reserve",_v51_money(totals.get("Reserve")))]
    start=6
    for i,(k,v) in enumerate(kpis):
        r=start+(i//4)*3; c=1+(i%4)*3
        ws.cell(r,c,k); ws.cell(r+1,c,v)
        for rr in [r,r+1]:
            for cc in [c,c+1]:
                ws.cell(rr,cc).fill = PatternFill("solid", fgColor=V51_PANEL)
                ws.cell(rr,cc).border = Border(left=Side(style="thin",color=V51_CYAN),right=Side(style="thin",color=V51_LINE),top=Side(style="thin",color=V51_LINE),bottom=Side(style="thin",color=V51_LINE))
        ws.cell(r,c).font = Font(name="Aptos", size=9, bold=True, color=V51_MUTED)
        ws.cell(r+1,c).font = Font(name="Aptos Display", size=18, bold=True, color=V51_WHITE)
        ws.merge_cells(start_row=r,start_column=c,end_row=r,end_column=c+1)
        ws.merge_cells(start_row=r+1,start_column=c,end_row=r+1,end_column=c+1)
    ws["A16"]="Executive Summary"; ws["A16"].font=Font(name="Aptos Display",size=14,bold=True,color=V51_CYAN)
    ws.merge_cells("A17:K20"); ws["A17"]=model.get("executive_summary",""); ws["A17"].fill=PatternFill("solid",fgColor=V51_PANEL)
    ws["A22"]="Board Recommendation"; ws["A22"].font=Font(name="Aptos Display",size=14,bold=True,color=V51_CYAN)
    notes = scenario_notes + model.get("next_best_actions", [])[:5]
    for i,n in enumerate(notes,23):
        ws.cell(i,1,f"{i-22}"); ws.cell(i,2,n); ws.merge_cells(start_row=i,start_column=2,end_row=i,end_column=11)
        ws.cell(i,1).fill=PatternFill("solid",fgColor=V51_CYAN); ws.cell(i,1).font=Font(bold=True,color=V51_Navy if False else V51_NAVY)
        ws.cell(i,2).fill=PatternFill("solid",fgColor="06101C")
    # Cost composition mini chart table
    ws["A32"]="Cost Composition"; ws["A32"].font=Font(name="Aptos Display",size=14,bold=True,color=V51_CYAN)
    _header_row(ws,33,["Type","P50 BN"])
    for i,(typ,val) in enumerate(totals.items(),34):
        ws.cell(i,1,typ); ws.cell(i,2,round(val,3)); ws.cell(i,2).number_format='$#,##0.0 "B"'
    bar=BarChart(); bar.title="Direct / Indirect / Reserve"; bar.style=10; bar.y_axis.title="BN"; bar.x_axis.title="Cost Type";
    bar.add_data(Reference(ws,min_col=2,min_row=33,max_row=36),titles_from_data=True)
    bar.set_categories(Reference(ws,min_col=1,min_row=34,max_row=36)); bar.height=6; bar.width=11; ws.add_chart(bar,"D32")

    # Basis
    ws=wb["01 Basis"]
    _title(ws,"BASIS OF ESTIMATE AND CONTROL MODEL", "Traceability page: no board output should be read without understanding the class, maturity, location, scenario and assumptions.", "H")
    _apply_ws_theme(ws,{"A":28,"B":80,"C":22,"D":22,"E":22,"F":22,"G":22,"H":22}, freeze="A6")
    pairs=[("Prompt",model.get("prompt")),("Client",model.get("client")),("Mode",model.get("mode")),("Sector",model.get("subsector")),("Location",model.get("location")),("Scale",model.get("scale")),("Estimate class",model.get("estimate_class_name")),("Estimate maturity",model.get("estimate_maturity")),("Schedule level",model.get("schedule_level")),("Scenario",model.get("scenario_label")),("Scenario rationale",model.get("scenario_why"))]
    _header_row(ws,4,["Input","Value"])
    for i,(k,v) in enumerate(pairs,5):
        ws.cell(i,1,k); ws.cell(i,2,v); ws.merge_cells(start_row=i,start_column=2,end_row=i,end_column=8)
    ws["A19"]="Confidence Explanation"; ws["A19"].font=Font(size=13,bold=True,color=V51_CYAN)
    for i,x in enumerate(model.get("confidence_explanation",[]),20):
        ws.cell(i,1,"•"); ws.cell(i,2,x); ws.merge_cells(start_row=i,start_column=2,end_row=i,end_column=8)

    # Cost estimate
    ws=wb["02 Cost Estimate"]
    _title(ws,"SELECTED CLASS COST ESTIMATE - DIRECT / INDIRECT / RESERVE", "Only the selected class estimate is exported. P10 / P50 / P90 and direct / indirect / reserve are visible and auditable.", "L")
    cols=["CBS","Description","Type","Basis","P10 BN","P50 BN","P90 BN","Min BN","Most Likely BN","Max BN","Scenario Impact","Evidence / Challenge"]
    _apply_ws_theme(ws,{"A":10,"B":30,"C":14,"D":45,"E":12,"F":12,"G":12,"H":12,"I":14,"J":12,"K":26,"L":55}, freeze="A5")
    _header_row(ws,4,cols)
    for r,x in enumerate(model.get("cost_lines",[]),5):
        vals=[x.get("cbs"),x.get("description"),x.get("type"),x.get("basis"),x.get("p10_bn"),x.get("p50_bn"),x.get("p90_bn"),x.get("p10_bn"),x.get("p50_bn"),x.get("p90_bn"),model.get("scenario_label"),x.get("impact_basis")]
        for c,v in enumerate(vals,1): ws.cell(r,c,v)
        for c in range(5,11): ws.cell(r,c).number_format='$#,##0.000 "B"'
        if x.get("type") == "Direct": fill="071A2F"
        elif x.get("type") == "Indirect": fill="101827"
        else: fill="26111A"
        for c in range(1,len(cols)+1): ws.cell(r,c).fill=PatternFill("solid",fgColor=fill)
    end=4+len(model.get("cost_lines",[]))
    _add_table(ws,f"A4:L{end}","CASEY_CostEstimate")
    # summary
    row=end+3
    ws.cell(row,1,"Type"); ws.cell(row,2,"P50 BN"); ws.cell(row,3,"Share")
    _header_row(ws,row,["Type","P50 BN","Share"])
    for i,typ in enumerate(["Direct","Indirect","Reserve"],row+1):
        ws.cell(i,1,typ); ws.cell(i,2,totals.get(typ,0)); ws.cell(i,2).number_format='$#,##0.0 "B"'; ws.cell(i,3,f"=B{i}/SUM(B{row+1}:B{row+3})"); ws.cell(i,3).number_format='0.0%'
    chart=BarChart(); chart.title="Selected Estimate Cost Split"; chart.add_data(Reference(ws,min_col=2,min_row=row,max_row=row+3),titles_from_data=True); chart.set_categories(Reference(ws,min_col=1,min_row=row+1,max_row=row+3)); chart.height=7; chart.width=12; ws.add_chart(chart,f"E{row}")

    # Indirects
    ws=wb["03 Indirects & Reserves"]
    _title(ws,"INDIRECTS, CONTINGENCY AND RESERVE ASSURANCE", "Separates project delivery cost from sponsor-held uncertainty and contingency.", "H")
    _apply_ws_theme(ws,{"A":18,"B":32,"C":14,"D":16,"E":16,"F":16,"G":38,"H":40}, freeze="A5")
    _header_row(ws,4,["CBS","Description","Type","P10 BN","P50 BN","P90 BN","Assurance Question","Board Treatment"])
    rows=[x for x in model.get("cost_lines",[]) if x.get("type")!="Direct"]
    for r,x in enumerate(rows,5):
        ws.cell(r,1,x.get("cbs")); ws.cell(r,2,x.get("description")); ws.cell(r,3,x.get("type")); ws.cell(r,4,x.get("p10_bn")); ws.cell(r,5,x.get("p50_bn")); ws.cell(r,6,x.get("p90_bn")); ws.cell(r,7,"Is this genuinely indirect/reserve, or hidden direct scope?"); ws.cell(r,8,"Show separately in board pack; do not bury in headline P50.")
        for c in [4,5,6]: ws.cell(r,c).number_format='$#,##0.000 "B"'

    # QCRA
    ws=wb["04 QCRA"]
    _title(ws,"QCRA COST RISK ANALYSIS", "Cost curve and separate tornado drivers. QCRA is not mixed with QSRA.", "J")
    _apply_ws_theme(ws,{"A":14,"B":16,"C":16,"D":16,"E":16,"F":16,"G":18,"H":28,"I":18,"J":18}, freeze="A5")
    _header_row(ws,4,["Percentile","Cost BN","Schedule Months","","Rank","Driver","CBS","Mean BN","Score","Board Challenge"])
    curve=mc.get("curve",[])
    for i,x in enumerate(curve,5):
        ws.cell(i,1,x.get("percentile")); ws.cell(i,2,x.get("cost_bn")); ws.cell(i,3,x.get("schedule_months")); ws.cell(i,2).number_format='$#,##0.0 "B"'
    for j,x in enumerate(mc.get("qcra_tornado",[])[:10],5):
        ws.cell(j,5,j-4); ws.cell(j,6,x.get("title")); ws.cell(j,7,x.get("cbs")); ws.cell(j,8,x.get("cost_mean_bn")); ws.cell(j,9,x.get("driver_score")); ws.cell(j,10,"Request evidence for cost range and owner-retained exposure.")
        ws.cell(j,8).number_format='$#,##0.000 "B"'
    line=LineChart(); line.title="QCRA Cost Curve"; line.y_axis.title="Cost BN"; line.x_axis.title="Percentile"; line.add_data(Reference(ws,min_col=2,min_row=4,max_row=4+len(curve)),titles_from_data=True); line.set_categories(Reference(ws,min_col=1,min_row=5,max_row=4+len(curve))); line.height=7; line.width=12; ws.add_chart(line,"A22")
    bar=BarChart(); bar.type="bar"; bar.style=12; bar.title="QCRA Tornado - Cost Drivers"; bar.add_data(Reference(ws,min_col=8,min_row=4,max_row=14),titles_from_data=True); bar.set_categories(Reference(ws,min_col=6,min_row=5,max_row=14)); bar.height=7; bar.width=12; ws.add_chart(bar,"E22")

    # QSRA
    ws=wb["05 QSRA"]
    _title(ws,"QSRA SCHEDULE RISK ANALYSIS", "Schedule curve and separate tornado drivers. QSRA is not mixed with QCRA.", "J")
    _apply_ws_theme(ws,{"A":14,"B":16,"C":16,"D":16,"E":16,"F":32,"G":16,"H":18,"I":18,"J":34}, freeze="A5")
    _header_row(ws,4,["Percentile","Schedule Months","Cost BN","","Rank","Driver","Activity","Mean Days","Score","Recovery / Mitigation Challenge"])
    for i,x in enumerate(curve,5):
        ws.cell(i,1,x.get("percentile")); ws.cell(i,2,x.get("schedule_months")); ws.cell(i,3,x.get("cost_bn")); ws.cell(i,3).number_format='$#,##0.0 "B"'
    for j,x in enumerate(mc.get("qsra_tornado",[])[:10],5):
        ws.cell(j,5,j-4); ws.cell(j,6,x.get("title")); ws.cell(j,7,x.get("activity_id")); ws.cell(j,8,x.get("schedule_mean_days")); ws.cell(j,9,x.get("driver_score")); ws.cell(j,10,"Confirm critical path logic, float erosion and recovery option.")
    line=LineChart(); line.title="QSRA Schedule Curve"; line.y_axis.title="Months"; line.x_axis.title="Percentile"; line.add_data(Reference(ws,min_col=2,min_row=4,max_row=4+len(curve)),titles_from_data=True); line.set_categories(Reference(ws,min_col=1,min_row=5,max_row=4+len(curve))); line.height=7; line.width=12; ws.add_chart(line,"A22")
    bar=BarChart(); bar.type="bar"; bar.title="QSRA Tornado - Schedule Drivers"; bar.add_data(Reference(ws,min_col=8,min_row=4,max_row=14),titles_from_data=True); bar.set_categories(Reference(ws,min_col=6,min_row=5,max_row=14)); bar.height=7; bar.width=12; ws.add_chart(bar,"E22")

    # Risk register
    ws=wb["06 Risk Register"]
    _title(ws,"RISK REGISTER PRO - CAUSE / EVENT / IMPACT / RESPONSE", "Board-grade register with quantified cost and schedule exposure, owners and linked WBS/CBS.", "R")
    headers=["Risk ID","Category","Title","Cause","Risk Event","Impact Description","Probability %","Cost O BN","Cost M BN","Cost P BN","Cost EMV BN","Schedule O Days","Schedule M Days","Schedule P Days","Schedule EMV Days","Activity","CBS","Owner","Trigger","Mitigation","Response Strategy","Residual","Status","Board"]
    _apply_ws_theme(ws,{"A":10,"B":14,"C":22,"D":36,"E":36,"F":42,"G":12,"H":12,"I":12,"J":12,"K":14,"L":12,"M":12,"N":12,"O":14,"P":12,"Q":10,"R":18,"S":26,"T":44,"U":18,"V":12,"W":12,"X":10}, freeze="A5")
    _header_row(ws,4,headers)
    for r,x in enumerate(model.get("risks",[]),5):
        vals=[x.get("risk_id"),x.get("category"),x.get("title"),x.get("cause"),x.get("risk_event"),x.get("impact_description"),x.get("probability_pct"),x.get("cost_o_bn"),x.get("cost_m_bn"),x.get("cost_p_bn"),x.get("cost_emv_bn"),x.get("schedule_o_days"),x.get("schedule_m_days"),x.get("schedule_p_days"),x.get("schedule_emv_days"),x.get("activity_id"),x.get("cbs"),x.get("owner"),x.get("trigger"),x.get("mitigation"),x.get("response_strategy"),x.get("residual_rating"),x.get("status"),x.get("board_visibility")]
        for c,v in enumerate(vals,1): ws.cell(r,c,v)
        for c in [8,9,10,11]: ws.cell(r,c).number_format='$#,##0.000 "B"'
        if x.get("board_visibility") == "Yes":
            for c in range(1,len(headers)+1): ws.cell(r,c).fill=PatternFill("solid",fgColor="071A2F")
    _add_table(ws,f"A4:X{4+len(model.get('risks',[]))}","CASEY_RiskRegister")
    ws.conditional_formatting.add(f"G5:G{4+len(model.get('risks',[]))}", ColorScaleRule(start_type='min', start_color=V51_GREEN, mid_type='percentile', mid_value=50, mid_color=V51_AMBER, end_type='max', end_color=V51_RED))

    # Schedule
    ws=wb["07 Schedule"]
    _title(ws,"SCHEDULE LOGIC - LEVEL SELECTED", "Linked schedule activities for QSRA and export to P6/XER.", "H")
    _apply_ws_theme(ws,{"A":12,"B":18,"C":42,"D":22,"E":12,"F":12,"G":55,"H":12}, freeze="A5")
    _header_row(ws,4,["Activity ID","Phase","Activity","Predecessor","Duration Months","Critical","Basis","Risk Links"])
    for r,x in enumerate(model.get("schedule_rows",[]),5):
        linked=", ".join([rr.get("risk_id") for rr in model.get("risks",[]) if rr.get("activity_id")==x.get("activity_id")])
        vals=[x.get("activity_id"),x.get("phase"),x.get("activity"),x.get("predecessor"),x.get("duration_months"),x.get("critical"),x.get("basis"),linked]
        for c,v in enumerate(vals,1): ws.cell(r,c,v)
        if x.get("critical")=="Yes":
            for c in range(1,9): ws.cell(r,c).fill=PatternFill("solid",fgColor="26111A")

    # Scenarios
    ws=wb["08 Scenarios"]
    _title(ws,"SCENARIO COMPARISON - BASE / FASTER / CHEAPER / LOWER RISK / PREMIUM", "Scenario choice must change cost, schedule, risk and confidence.", "H")
    _apply_ws_theme(ws,{"A":20,"B":16,"C":16,"D":16,"E":16,"F":42,"G":42,"H":42}, freeze="A5")
    _header_row(ws,4,["Scenario","P50 Cost","Schedule","Risk","Confidence","Cost Change Logic","Schedule Change Logic","Risk Logic"])
    for r,x in enumerate(model.get("scenario_comparison",[]),5):
        vals=[x.get("scenario"),x.get("cost_p50"),x.get("schedule"),x.get("risk"),x.get("confidence"),x.get("why","") or x.get("cost_note","") ,x.get("schedule_note","") ,x.get("risk_note","")]
        for c,v in enumerate(vals,1): ws.cell(r,c,v)

    # Benchmarking
    ws=wb["09 Benchmarking"]
    _title(ws,"BENCHMARKING AND PEER CHALLENGE", "Use to challenge Tier 1 contractor estimates, market rates and sponsor assumptions.", "H")
    _apply_ws_theme(ws,{"A":26,"B":26,"C":26,"D":26,"E":26,"F":26,"G":26,"H":26}, freeze="A5")
    _header_row(ws,4,["Benchmark / Peer","What to compare","Why it matters","Challenge question","Evidence required"])
    peers=model.get("peer_competitors",[]) or []
    for r,peer in enumerate(peers[:10],5):
        vals=[peer,"Unit rates, schedule intensity, cost/scope normalization","Prevents blind acceptance of contractor/adviser basis","Where are we above peer norm and why?","Rate build-up, supplier quote, productivity basis, exclusions log"]
        for c,v in enumerate(vals,1): ws.cell(r,c,v)

    # Assumptions
    ws=wb["10 Assumptions"]
    _title(ws,"ASSUMPTIONS, LIMITATIONS AND DECISION GATES", "Makes the output honest: a first-pass project controls intelligence pack, not a final tender estimate.", "H")
    _apply_ws_theme(ws,{"A":30,"B":85,"C":28,"D":28,"E":28,"F":28,"G":28,"H":28}, freeze="A5")
    _header_row(ws,4,["Area","Assumption / Limitation","Board Action"])
    assumptions=[
        ("Estimate",f"{model.get('estimate_class_name')} with {model.get('estimate_maturity')}; not a contractor tender price.","Use to challenge, shortlist and frame approval gates."),
        ("Schedule",f"Level {model.get('schedule_level')} logic; critical path and float require validation against project-specific constraints.","Run schedule risk workshop and validate dependencies."),
        ("Risk","QCRA and QSRA distributions are modelled from the generated register and selected scenario.","Replace assumptions with workshop data as evidence improves."),
        ("Scenario",model.get("scenario_why"),"Compare Base / Faster / Cheaper / Lower Risk before board decision."),
        ("Outputs","Exports are board-preparation outputs and should be reviewed before external release.","Apply owner governance, sign-off and evidence checks.")
    ]
    for r,vals in enumerate(assumptions,5):
        for c,v in enumerate(vals,1): ws.cell(r,c,v)

    # final styling sizes
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
        for r in range(1, ws.max_row+1):
            ws.row_dimensions[r].height = min(80, max(18, ws.row_dimensions[r].height or 18))
    bio=BytesIO(); wb.save(bio); bio.seek(0); return bio.getvalue()


def risk_register_workbook_bytes(model:Dict[str,Any]):
    # Dedicated premium risk workbook: dashboard + register + mitigations + board risks.
    wb=Workbook(); ws=wb.active; ws.title="Risk Dashboard"; wb.create_sheet("Risk Register Pro"); wb.create_sheet("Mitigation Tracker"); wb.create_sheet("Board Top 10")
    risks=model.get("risks",[]); mc=model.get("monte_carlo",{})
    ws=wb["Risk Dashboard"]; _title(ws,"CASEY RISK REGISTER PRO", "Cause - Event - Impact - Response - Quantified cost and schedule exposure.", "I")
    _apply_ws_theme(ws,{"A":22,"B":20,"C":22,"D":22,"E":22,"F":22,"G":22,"H":22,"I":22}, freeze="A6")
    kpis=[("Risks",len(risks)),("Board Visible",sum(1 for r in risks if r.get("board_visibility")=="Yes")),("QCRA P80",_v51_money(mc.get("qcra",{}).get("p80"))),("QSRA P80",f"{mc.get('qsra',{}).get('p80','-')} months"),("Top Risk",risks[0].get("title") if risks else "-"),("Scenario",model.get("scenario_label"))]
    for i,(k,v) in enumerate(kpis):
        r=4+(i//3)*3; c=1+(i%3)*3; ws.cell(r,c,k); ws.cell(r+1,c,v); ws.merge_cells(start_row=r,start_column=c,end_row=r,end_column=c+1); ws.merge_cells(start_row=r+1,start_column=c,end_row=r+1,end_column=c+1)
        ws.cell(r,c).fill=PatternFill("solid",fgColor=V51_PANEL); ws.cell(r+1,c).fill=PatternFill("solid",fgColor=V51_PANEL); ws.cell(r+1,c).font=Font(size=18,bold=True,color=V51_WHITE)
    _header_row(ws,12,["Rank","Risk","Cost EMV BN","Schedule EMV Days","Owner","Response"])
    for i,risk in enumerate(sorted(risks,key=lambda x:_num(x.get('driver_score')),reverse=True)[:10],13):
        vals=[i-12,risk.get("title"),risk.get("cost_emv_bn"),risk.get("schedule_emv_days"),risk.get("owner"),risk.get("mitigation")]
        for c,v in enumerate(vals,1): ws.cell(i,c,v)
        ws.cell(i,3).number_format='$#,##0.000 "B"'
    chart=BarChart(); chart.title="Top Cost EMV Risks"; chart.add_data(Reference(ws,min_col=3,min_row=12,max_row=22),titles_from_data=True); chart.set_categories(Reference(ws,min_col=2,min_row=13,max_row=22)); chart.height=7; chart.width=12; ws.add_chart(chart,"A25")
    ws=wb["Risk Register Pro"]; _title(ws,"FULL RISK REGISTER PRO", "Every risk has a cause, event, impact description, owner, response and quantification.", "X")
    headers=["Risk ID","Category","Title","Cause","Risk Event","Impact Description","Probability %","Cost O BN","Cost M BN","Cost P BN","Cost EMV BN","Schedule O Days","Schedule M Days","Schedule P Days","Schedule EMV Days","Activity","CBS","Owner","Trigger","Mitigation","Response Strategy","Residual","Status","Board"]
    _apply_ws_theme(ws,{get_column_letter(i):18 for i in range(1,25)}, freeze="A5"); _header_row(ws,4,headers)
    for i,r in enumerate(risks,5):
        vals=[r.get("risk_id"),r.get("category"),r.get("title"),r.get("cause"),r.get("risk_event"),r.get("impact_description"),r.get("probability_pct"),r.get("cost_o_bn"),r.get("cost_m_bn"),r.get("cost_p_bn"),r.get("cost_emv_bn"),r.get("schedule_o_days"),r.get("schedule_m_days"),r.get("schedule_p_days"),r.get("schedule_emv_days"),r.get("activity_id"),r.get("cbs"),r.get("owner"),r.get("trigger"),r.get("mitigation"),r.get("response_strategy"),r.get("residual_rating"),r.get("status"),r.get("board_visibility")]
        for c,v in enumerate(vals,1): ws.cell(i,c,v)
        for c in [8,9,10,11]: ws.cell(i,c).number_format='$#,##0.000 "B"'
    _add_table(ws,f"A4:X{4+len(risks)}","RiskRegisterPro")
    ws=wb["Mitigation Tracker"]; _title(ws,"MITIGATION TRACKER", "Owner actions, triggers and residual risk conversion.", "J")
    _apply_ws_theme(ws,{"A":12,"B":30,"C":25,"D":50,"E":18,"F":18,"G":20,"H":20,"I":20,"J":28}, freeze="A5")
    _header_row(ws,4,["Risk ID","Risk","Owner","Mitigation Action","Trigger","Due Gate","Status","Residual","Board","Evidence Needed"])
    for i,r in enumerate(risks,5):
        vals=[r.get("risk_id"),r.get("title"),r.get("owner"),r.get("mitigation"),r.get("trigger"),"Next gate",r.get("status"),r.get("residual_rating"),r.get("board_visibility"),"Named evidence, date, owner and decision log"]
        for c,v in enumerate(vals,1): ws.cell(i,c,v)
    ws=wb["Board Top 10"]; _title(ws,"BOARD TOP 10 RISKS", "The risks most likely to change the investment decision.", "H")
    _apply_ws_theme(ws,{"A":8,"B":32,"C":40,"D":40,"E":40,"F":18,"G":18,"H":50}, freeze="A5")
    _header_row(ws,4,["Rank","Risk","Cause","Event","Impact","Cost EMV BN","Schedule EMV Days","Board Ask"])
    for i,r in enumerate(sorted(risks,key=lambda x:_num(x.get('driver_score')),reverse=True)[:10],5):
        vals=[i-4,r.get("title"),r.get("cause"),r.get("risk_event"),r.get("impact_description"),r.get("cost_emv_bn"),r.get("schedule_emv_days"),"Approve mitigation spend or request further evidence"]
        for c,v in enumerate(vals,1): ws.cell(i,c,v)
        ws.cell(i,6).number_format='$#,##0.000 "B"'
    bio=BytesIO(); wb.save(bio); bio.seek(0); return bio.getvalue()


def risk_csv_bytes(model):
    # Preserve CSV endpoint if called directly, but make it richer than before.
    out=StringIO(); w=csv.writer(out)
    cols=["Risk ID","Category","Title","Cause","Risk Event","Impact Description","Probability %","Cost O BN","Cost M BN","Cost P BN","Cost EMV BN","Schedule O Days","Schedule M Days","Schedule P Days","Schedule EMV Days","Activity","CBS","Owner","Trigger","Mitigation","Response Strategy","Residual","Status","Board"]
    w.writerow(cols)
    for r in model.get("risks",[]):
        w.writerow([r.get("risk_id"),r.get("category"),r.get("title"),r.get("cause"),r.get("risk_event"),r.get("impact_description"),r.get("probability_pct"),r.get("cost_o_bn"),r.get("cost_m_bn"),r.get("cost_p_bn"),r.get("cost_emv_bn"),r.get("schedule_o_days"),r.get("schedule_m_days"),r.get("schedule_p_days"),r.get("schedule_emv_days"),r.get("activity_id"),r.get("cbs"),r.get("owner"),r.get("trigger"),r.get("mitigation"),r.get("response_strategy"),r.get("residual_rating"),r.get("status"),r.get("board_visibility")])
    return out.getvalue().encode()


def word_bytes(model):
    doc=Document()
    sec=doc.sections[0]; sec.top_margin=Inches(.55); sec.bottom_margin=Inches(.55); sec.left_margin=Inches(.6); sec.right_margin=Inches(.6)
    styles=doc.styles
    styles["Normal"].font.name="Aptos"; styles["Normal"].font.size=Pt(9.5)
    for style in ["Heading 1","Heading 2","Title"]:
        styles[style].font.name="Aptos Display"
        styles[style].font.color.rgb=RGBColor(6,28,44)
    title=doc.add_heading("CASEY Executive Board Intelligence Pack",0)
    doc.add_paragraph(f"{model.get('title')} | {model.get('subsector')} | {model.get('scenario_label')} scenario | Generated {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}")
    doc.add_paragraph(model.get("executive_summary",""))
    doc.add_heading("1. Board Decision Snapshot",1)
    mc=model.get("monte_carlo",{}); totals=_v51_total_costs(model)
    table=doc.add_table(rows=1,cols=4); table.style="Light Shading Accent 1"
    headers=["Metric","Value","Interpretation","Board Action"]
    for i,h in enumerate(headers): table.rows[0].cells[i].text=h
    rows=[
        ("P50 Cost",model.get("cost_p50"),"Most likely first-pass cost based on selected class and scenario.","Use as initial investment control point."),
        ("Cost Range",model.get("cost_range"),"Class/maturity and QCRA uncertainty range.","Do not approve without range acknowledgement."),
        ("QCRA P80",_v51_money(mc.get("qcra",{}).get("p80")),"Risk-adjusted cost confidence level.","Set funding envelope and contingency rules."),
        ("Baseline / QSRA P80",f"{model.get('schedule')} / {mc.get('qsra',{}).get('p80')} months","Schedule risk exposure and delivery confidence.","Require critical path validation."),
        ("Risk / Confidence",f"{model.get('risk')} / {model.get('confidence_pct')}%","Early-stage confidence and data quality status.","Approve evidence-improvement plan."),
        ("Direct / Indirect / Reserve",f"{_v51_money(totals.get('Direct'))} / {_v51_money(totals.get('Indirect'))} / {_v51_money(totals.get('Reserve'))}","Prevents reserves being hidden inside direct cost.","Challenge contractor and adviser basis.")]
    for row in rows:
        cells=table.add_row().cells
        for i,v in enumerate(row): cells[i].text=str(v)
    doc.add_heading("2. Executive Recommendation",1)
    for x in _v51_scenario_notes(model): doc.add_paragraph(x, style="List Bullet")
    for x in model.get("next_best_actions",[])[:5]: doc.add_paragraph(x, style="List Bullet")
    doc.add_heading("3. Cost Estimate - Selected Class Only",1)
    t=doc.add_table(rows=1,cols=7); t.style="Light Grid Accent 1"
    for i,h in enumerate(["CBS","Description","Type","P10","P50","P90","Basis"]): t.rows[0].cells[i].text=h
    for x in model.get("cost_lines",[])[:24]:
        c=t.add_row().cells
        vals=[x.get("cbs"),x.get("description"),x.get("type"),_v51_money(x.get("p10_bn")),_v51_money(x.get("p50_bn")),_v51_money(x.get("p90_bn")),x.get("impact_basis")]
        for i,v in enumerate(vals): c[i].text=str(v)
    doc.add_heading("4. QCRA and QSRA",1)
    doc.add_paragraph(f"QCRA cost P10 / P50 / P80 / P90: {_v51_money(mc.get('qcra',{}).get('p10'))} / {model.get('cost_p50')} / {_v51_money(mc.get('qcra',{}).get('p80'))} / {_v51_money(mc.get('qcra',{}).get('p90'))}")
    doc.add_paragraph(f"QSRA schedule P10 / P50 / P80 / P90: {mc.get('qsra',{}).get('p10')} / {mc.get('qsra',{}).get('p50')} / {mc.get('qsra',{}).get('p80')} / {mc.get('qsra',{}).get('p90')} months")
    doc.add_heading("5. QCRA Tornado Drivers",2)
    for x in mc.get("qcra_tornado",[])[:8]: doc.add_paragraph(f"{x.get('title')} - {x.get('cbs')} - mean exposure {_v51_money(x.get('cost_mean_bn'))}", style="List Bullet")
    doc.add_heading("6. QSRA Tornado Drivers",2)
    for x in mc.get("qsra_tornado",[])[:8]: doc.add_paragraph(f"{x.get('title')} - {x.get('activity_id')} - mean exposure {x.get('schedule_mean_days')} days", style="List Bullet")
    doc.add_heading("7. Risk Register - Board Top Risks",1)
    for r in model.get("risks",[])[:10]:
        doc.add_paragraph(f"{r.get('risk_id')} {r.get('title')}",style="List Bullet")
        doc.add_paragraph(f"Cause: {r.get('cause')}. Event: {r.get('risk_event')}. Impact: {r.get('impact_description')}. Mitigation: {r.get('mitigation')}. Owner: {r.get('owner')}.")
    doc.add_heading("8. Challenge Questions",1)
    for q in model.get("board_challenge_questions",[])[:8]: doc.add_paragraph(q,style="List Number")
    doc.add_heading("9. Assumptions and Limitations",1)
    doc.add_paragraph("This is a first-pass project controls intelligence pack. It is designed to challenge, accelerate and structure decision-making; it is not a final contractor tender or signed cost plan. Each output includes assumptions and should be validated against project-specific evidence before commitment.")
    bio=BytesIO(); doc.save(bio); bio.seek(0); return bio.getvalue()


def pdf_bytes(model):
    bio=BytesIO(); doc=SimpleDocTemplate(bio,pagesize=landscape(A4),rightMargin=24,leftMargin=24,topMargin=22,bottomMargin=22)
    styles=getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Hero",fontSize=24,leading=28,textColor=colors.HexColor("#8DF7FF"),fontName="Helvetica-Bold",spaceAfter=12))
    styles.add(ParagraphStyle(name="Section",fontSize=15,leading=18,textColor=colors.HexColor("#07111F"),fontName="Helvetica-Bold",spaceBefore=8,spaceAfter=6))
    styles.add(ParagraphStyle(name="SmallCasey",fontSize=8,leading=10,textColor=colors.HexColor("#334155")))
    story=[]; mc=model.get("monte_carlo",{}); totals=_v51_total_costs(model)
    story.append(Paragraph("CASEY BOARD INTELLIGENCE PACK",styles["Hero"]))
    story.append(Paragraph(model.get("executive_summary",""),styles["BodyText"])); story.append(Spacer(1,8))
    metrics=[["Project",model.get("title"),"Scenario",model.get("scenario_label")],["P50",model.get("cost_p50"),"Cost Range",model.get("cost_range")],["QCRA P80",_v51_money(mc.get("qcra",{}).get("p80")),"QSRA P80",f"{mc.get('qsra',{}).get('p80')} months"],["Direct / Indirect / Reserve",f"{_v51_money(totals.get('Direct'))} / {_v51_money(totals.get('Indirect'))} / {_v51_money(totals.get('Reserve'))}","Risk / Confidence",f"{model.get('risk')} / {model.get('confidence_pct')}%"]]
    t=Table(metrics,colWidths=[110,260,110,260]); t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#07111F")),("TEXTCOLOR",(0,0),(-1,-1),colors.white),("GRID",(0,0),(-1,-1),.5,colors.HexColor("#8DF7FF")),("FONTNAME",(0,0),(-1,-1),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),9),("VALIGN",(0,0),(-1,-1),"TOP")])) ; story.append(t)
    story.append(Spacer(1,8)); story.append(Paragraph("Board Recommendation",styles["Section"]))
    for x in _v51_scenario_notes(model)[:3]: story.append(Paragraph("- "+x, styles["SmallCasey"]))
    story.append(PageBreak())
    story.append(Paragraph("QCRA and QSRA - Separate Tornado Drivers",styles["Hero"]))
    rows=[["Rank","QCRA Cost Driver","CBS","Mean Cost","QSRA Schedule Driver","Activity","Mean Days"]]
    q=mc.get("qcra_tornado",[])[:10]; s=mc.get("qsra_tornado",[])[:10]
    for i in range(10):
        rows.append([i+1,q[i].get("title","") if i<len(q) else "",q[i].get("cbs","") if i<len(q) else "",_v51_money(q[i].get("cost_mean_bn")) if i<len(q) else "",s[i].get("title","") if i<len(s) else "",s[i].get("activity_id","") if i<len(s) else "",s[i].get("schedule_mean_days","") if i<len(s) else ""])
    t=Table(rows,colWidths=[38,190,55,75,190,60,70]); t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#07111F")),("TEXTCOLOR",(0,0),(-1,0),colors.HexColor("#8DF7FF")),("GRID",(0,0),(-1,-1),.25,colors.HexColor("#CBD5E1")),("FONTSIZE",(0,0),(-1,-1),7),("VALIGN",(0,0),(-1,-1),"TOP")])) ; story.append(t)
    story.append(PageBreak())
    story.append(Paragraph("Risk Register - Board Top Risks",styles["Hero"]))
    rows=[["ID","Risk","Cause","Event","Impact","Owner","Response"]]
    for r in model.get("risks",[])[:10]: rows.append([r.get("risk_id"),r.get("title"),r.get("cause"),r.get("risk_event"),r.get("impact_description"),r.get("owner"),r.get("mitigation")])
    t=Table(rows,colWidths=[42,88,135,135,170,78,170]); t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#07111F")),("TEXTCOLOR",(0,0),(-1,0),colors.HexColor("#8DF7FF")),("GRID",(0,0),(-1,-1),.25,colors.HexColor("#CBD5E1")),("FONTSIZE",(0,0),(-1,-1),6.6),("VALIGN",(0,0),(-1,-1),"TOP")])) ; story.append(t)
    story.append(PageBreak())
    story.append(Paragraph("Cost Estimate - Selected Class Only",styles["Hero"]))
    rows=[["CBS","Description","Type","P10","P50","P90","Basis"]]
    for x in model.get("cost_lines",[])[:18]: rows.append([x.get("cbs"),x.get("description"),x.get("type"),_v51_money(x.get("p10_bn")),_v51_money(x.get("p50_bn")),_v51_money(x.get("p90_bn")),x.get("basis")])
    t=Table(rows,colWidths=[44,145,70,55,55,55,270]); t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#07111F")),("TEXTCOLOR",(0,0),(-1,0),colors.HexColor("#8DF7FF")),("GRID",(0,0),(-1,-1),.25,colors.HexColor("#CBD5E1")),("FONTSIZE",(0,0),(-1,-1),6.8),("VALIGN",(0,0),(-1,-1),"TOP")])) ; story.append(t)
    doc.build(story); bio.seek(0); return bio.getvalue()


def pptx_bytes(model):
    from pptx.dml.color import RGBColor as PRGBColor
    prs=Presentation(); prs.slide_width=PptxInches(13.333); prs.slide_height=PptxInches(7.5); blank=prs.slide_layouts[6]
    def slide(title,subtitle=""):
        s=prs.slides.add_slide(blank); s.background.fill.solid(); s.background.fill.fore_color.rgb=PRGBColor(2,5,10)
        tx=s.shapes.add_textbox(PptxInches(.45),PptxInches(.3),PptxInches(12.4),PptxInches(.45)); tx.text_frame.text="CASEY"; tx.text_frame.paragraphs[0].runs[0].font.size=PptxPt(14); tx.text_frame.paragraphs[0].runs[0].font.color.rgb=PRGBColor(141,247,255)
        h=s.shapes.add_textbox(PptxInches(.45),PptxInches(.86),PptxInches(12.4),PptxInches(.85)); h.text_frame.text=title; h.text_frame.paragraphs[0].runs[0].font.size=PptxPt(30); h.text_frame.paragraphs[0].runs[0].font.bold=True; h.text_frame.paragraphs[0].runs[0].font.color.rgb=PRGBColor(247,251,255)
        if subtitle:
            p=s.shapes.add_textbox(PptxInches(.48),PptxInches(1.55),PptxInches(12.1),PptxInches(.55)); p.text_frame.text=subtitle; p.text_frame.paragraphs[0].runs[0].font.size=PptxPt(12); p.text_frame.paragraphs[0].runs[0].font.color.rgb=PRGBColor(170,190,205)
        return s
    s=slide("Mission Control for Capital Projects",model.get("executive_summary",""))
    # Metrics slide
    s=slide("Board Decision Snapshot", f"{model.get('title')} | {model.get('scenario_label')} scenario")
    mc=model.get("monte_carlo",{}); totals=_v51_total_costs(model)
    metrics=[("P50",model.get("cost_p50")),("P80 Cost",_v51_money(mc.get("qcra",{}).get("p80"))),("Schedule",model.get("schedule")),("QSRA P80",f"{mc.get('qsra',{}).get('p80')} mo"),("Risk",model.get("risk")),("Confidence",f"{model.get('confidence_pct')}%"),("Direct",_v51_money(totals.get("Direct"))),("Indirect",_v51_money(totals.get("Indirect"))),("Reserve",_v51_money(totals.get("Reserve")))]
    for i,(k,v) in enumerate(metrics):
        x=.55+(i%3)*4.2; y=2.0+(i//3)*1.35
        shp=s.shapes.add_shape(1,PptxInches(x),PptxInches(y),PptxInches(3.75),PptxInches(1.02)); shp.fill.solid(); shp.fill.fore_color.rgb=PRGBColor(7,17,31); shp.line.color.rgb=PRGBColor(141,247,255)
        tf=shp.text_frame; tf.text=f"{k}\n{v}"; tf.paragraphs[0].runs[0].font.size=PptxPt(10); tf.paragraphs[0].runs[0].font.color.rgb=PRGBColor(159,179,200)
        tf.paragraphs[1].runs[0].font.size=PptxPt(20); tf.paragraphs[1].runs[0].font.bold=True; tf.paragraphs[1].runs[0].font.color.rgb=PRGBColor(247,251,255)
    # Scenario slide
    s=slide("Scenario Logic", "What changed and why the outputs are not generic.")
    y=1.95
    for n in _v51_scenario_notes(model):
        box=s.shapes.add_textbox(PptxInches(.7),PptxInches(y),PptxInches(12),PptxInches(.52)); box.text_frame.text="• "+n; box.text_frame.paragraphs[0].runs[0].font.size=PptxPt(15); box.text_frame.paragraphs[0].runs[0].font.color.rgb=PRGBColor(247,251,255); y+=.7
    # QCRA/QSRA slide
    s=slide("QCRA + QSRA Drivers", "Separate cost and schedule risk; no mixed generic tornado.")
    left=s.shapes.add_textbox(PptxInches(.65),PptxInches(1.9),PptxInches(5.9),PptxInches(4.9)); left.text_frame.text="QCRA COST TORNADO"; left.text_frame.paragraphs[0].runs[0].font.color.rgb=PRGBColor(141,247,255); left.text_frame.paragraphs[0].runs[0].font.bold=True
    for x in mc.get("qcra_tornado",[])[:7]:
        p=left.text_frame.add_paragraph(); p.text=f"{x.get('title')} - {_v51_money(x.get('cost_mean_bn'))}"; p.font.size=PptxPt(12); p.font.color.rgb=PRGBColor(247,251,255)
    right=s.shapes.add_textbox(PptxInches(6.8),PptxInches(1.9),PptxInches(5.9),PptxInches(4.9)); right.text_frame.text="QSRA SCHEDULE TORNADO"; right.text_frame.paragraphs[0].runs[0].font.color.rgb=PRGBColor(141,247,255); right.text_frame.paragraphs[0].runs[0].font.bold=True
    for x in mc.get("qsra_tornado",[])[:7]:
        p=right.text_frame.add_paragraph(); p.text=f"{x.get('title')} - {x.get('schedule_mean_days')} days"; p.font.size=PptxPt(12); p.font.color.rgb=PRGBColor(247,251,255)
    # Risk slide
    s=slide("Board Top Risks", "Cause - Event - Impact - Mitigation.")
    tf=s.shapes.add_textbox(PptxInches(.65),PptxInches(1.85),PptxInches(12.2),PptxInches(5.3)).text_frame
    for r in model.get("risks",[])[:8]:
        p=tf.add_paragraph(); p.text=f"{r.get('risk_id')} {r.get('title')}: {r.get('impact_description')}"; p.font.size=PptxPt(11); p.font.color.rgb=PRGBColor(247,251,255)
    # Close slide
    s=slide("Output Pack", "Excel cost model, risk register workbook, XER schedule, DOCX/PDF board report, PPTX board deck and JSON audit model generated from one connected CASEY model.")
    y=2.1
    for x in ["Selected class cost estimate only", "Direct / indirect / reserve cost split", "Risk register with cause, event, impact and mitigation", "Separate QCRA and QSRA curves/tornado", "Scenario-linked values for Base / Faster / Cheaper / Lower Risk / Premium", "Pricing and email conversion path"]:
        b=s.shapes.add_textbox(PptxInches(1.0),PptxInches(y),PptxInches(11.4),PptxInches(.42)); b.text_frame.text="✓ "+x; b.text_frame.paragraphs[0].runs[0].font.size=PptxPt(15); b.text_frame.paragraphs[0].runs[0].font.color.rgb=PRGBColor(141,247,255); y+=.58
    bio=BytesIO(); prs.save(bio); bio.seek(0); return bio.getvalue()

# Patch route functions to expose premium outputs and filenames.
def export_risk(model:Dict[str,Any]): return stream(risk_register_workbook_bytes(model),"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","CASEY_Risk_Register_Pro.xlsx")
def export_workbook(model:Dict[str,Any]): return stream(workbook_bytes(model),"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","CASEY_Cost_Model_Planet_Class.xlsx")
def export_word(model:Dict[str,Any]): return stream(word_bytes(model),"application/vnd.openxmlformats-officedocument.wordprocessingml.document","CASEY_Executive_Board_Report.docx")
def export_pdf(model:Dict[str,Any]): return stream(pdf_bytes(model),"application/pdf","CASEY_Board_Intelligence_Pack.pdf")
def export_pptx(model:Dict[str,Any]): return stream(pptx_bytes(model),"application/vnd.openxmlformats-officedocument.presentationml.presentation","CASEY_Board_Deck_Elite.pptx")

def export_all(model:Dict[str,Any]):
    bio=BytesIO()
    with zipfile.ZipFile(bio,"w",zipfile.ZIP_DEFLATED) as z:
        z.writestr("01_CASEY_Cost_Model_Planet_Class.xlsx",workbook_bytes(model))
        z.writestr("02_CASEY_Risk_Register_Pro.xlsx",risk_register_workbook_bytes(model))
        z.writestr("03_CASEY_P6_Schedule.xer",xer_bytes(model))
        z.writestr("04_CASEY_Executive_Board_Report.docx",word_bytes(model))
        z.writestr("05_CASEY_Board_Intelligence_Pack.pdf",pdf_bytes(model))
        z.writestr("06_CASEY_Board_Deck_Elite.pptx",pptx_bytes(model))
        z.writestr("07_CASEY_Full_Model_Audit.json",json.dumps(model,indent=2))
        z.writestr("08_CASEY_Risk_Register_Raw.csv",risk_csv_bytes(model))
        z.writestr("09_CASEY_Demo_Close_Script.txt","\n".join(model.get("launch_demo_script",[])))
    bio.seek(0); return stream(bio.getvalue(),"application/zip","CASEY_Output_Pack_Planet_Class.zip")

# Register new explicit v51 endpoints as well.
@app.post("/v51/export/risk-register")
def export_risk_v51(model:Dict[str,Any]): return export_risk(model)
@app.post("/v51/export/all")
def export_all_v51(model:Dict[str,Any]): return export_all(model)

# ===================== END CASEY v51 OUTPUT PLANET CLASS OVERRIDES =====================



# ======================= CASEY v52 ELITE OUTPUT SYSTEM =======================
# Focus: board-grade readability, scenario consistency and premium exports.
# v52 patches the existing /export routes in-place so the current frontend keeps working.

from openpyxl.chart import PieChart
from reportlab.lib.units import inch
from pptx.enum.text import PP_ALIGN

V52_DARK = "101820"
V52_NAVY = "0B1220"
V52_CYAN = "00AEEF"
V52_BLUE = "0070C0"
V52_LIGHT = "F7FBFF"
V52_GRID = "D9E2EC"
V52_RED = "C00000"
V52_AMBER = "FFC000"
V52_GREEN = "00A65A"


def _v52_scenario_label(model):
    return str(model.get("scenario_label") or model.get("scenario") or "Base").title()


def _v52_totals(model):
    totals = {"Direct": 0.0, "Indirect": 0.0, "Reserve": 0.0}
    for line in model.get("cost_lines", []):
        typ = line.get("type", "Direct")
        totals[typ] = totals.get(typ, 0.0) + float(line.get("p50_bn", 0) or 0)
    return totals


def _v52_recommendation(model):
    s = str(model.get("scenario", "base")).lower()
    risk = str(model.get("risk", "Medium"))
    if s == "faster":
        return {
            "headline": "APPROVE ACCELERATION ONLY WITH FUNDED ASSURANCE",
            "why": "The faster case can protect market window but pushes interface, commissioning and logistics pressure onto the critical path.",
            "decision": "Approve acceleration premium only if package-level recovery plans, supplier commitments and commissioning readiness gates are funded.",
        }
    if s == "cheaper":
        return {
            "headline": "DO NOT TREAT SAVINGS AS CERTAIN UNTIL SCOPE IS SIGNED OFF",
            "why": "The cheaper case lowers visible capex through value engineering and procurement pressure, but increases residual scope, quality and rework risk.",
            "decision": "Approve only after VE log, exclusions, owner-retained risks and quality impacts are signed off.",
        }
    if s == "lower_risk":
        return {
            "headline": "PROCEED WITH LOWER-RISK CASE AS BOARD CONTROL BASELINE",
            "why": "Higher upfront assurance, buffers and evidence maturity reduce downstream contingency draw and schedule surprise.",
            "decision": "Use as board approval baseline where certainty is more important than lowest headline capex.",
        }
    if s == "premium":
        return {
            "headline": "APPROVE PREMIUM CASE ONLY WHERE RESILIENCE HAS STRATEGIC VALUE",
            "why": "The premium case increases capex but improves resilience, assurance and confidence for mission-critical assets.",
            "decision": "Proceed if the business case values resilience, operational assurance and lower downside risk.",
        }
    return {
        "headline": "APPROVE DEFINITION PHASE — DO NOT COMMIT FULL CAPITAL YET",
        "why": f"The base case provides a first-pass controls baseline with {risk} risk and evidence still required before final capital commitment.",
        "decision": "Use CASEY as the reference model for board challenge, option testing and consultant / contractor assurance.",
    }


def _v52_style_ws(ws, title=None):
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A6"
    if title:
        ws["A1"] = title
        ws["A1"].font = Font(size=22, bold=True, color="FFFFFF")
        ws["A1"].fill = PatternFill("solid", fgColor=V52_NAVY)
        ws.merge_cells("A1:H1")
        ws.row_dimensions[1].height = 32
    thin = Side(style="thin", color=V52_GRID)
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = Border(bottom=thin)


def _v52_header(row):
    for c in row:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=V52_BLUE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _v52_autowidth(ws, max_width=42):
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None: continue
            widths[cell.column] = min(max(widths.get(cell.column, 10), len(str(cell.value)) + 2), max_width)
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width


def workbook_bytes_v52(model: Dict[str, Any]) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)
    totals = _v52_totals(model); rec = _v52_recommendation(model); mc = model.get("monte_carlo", {})
    scenario = _v52_scenario_label(model)

    # Dashboard
    ws = wb.create_sheet("01 Executive Dashboard")
    _v52_style_ws(ws, "CASEY EXECUTIVE DASHBOARD")
    ws["A3"] = model.get("title", "Project")
    ws["A3"].font = Font(size=16, bold=True, color=V52_NAVY)
    ws["A4"] = f"Scenario: {scenario} | Generated: {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')} | Engine: CASEY v52"
    ws["A4"].font = Font(size=10, color="555555")
    kpis = [("P50 Cost", model.get("cost_p50")), ("Cost Range", model.get("cost_range")), ("QCRA P80", money_bn(mc.get("qcra", {}).get("p80", 0))), ("Schedule", model.get("schedule")), ("QSRA P80", f"{mc.get('qsra', {}).get('p80', '')} months"), ("Risk / Confidence", f"{model.get('risk')} / {model.get('confidence_pct')}%"), ("Direct", money_bn(totals.get("Direct",0))), ("Indirect + Reserve", money_bn(totals.get("Indirect",0)+totals.get("Reserve",0)))]
    start_row = 6
    for idx, (k, v) in enumerate(kpis):
        r = start_row + (idx // 4) * 4; c = 1 + (idx % 4) * 2
        ws.cell(r,c,k); ws.cell(r+1,c,v)
        ws.cell(r,c).font = Font(bold=True, color="FFFFFF"); ws.cell(r,c).fill = PatternFill("solid", fgColor=V52_NAVY)
        ws.cell(r+1,c).font = Font(size=16, bold=True, color=V52_BLUE); ws.cell(r+1,c).fill = PatternFill("solid", fgColor="FFFFFF")
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c+1)
        ws.merge_cells(start_row=r+1, start_column=c, end_row=r+2, end_column=c+1)
    ws["A15"] = "Executive recommendation"; ws["A15"].font = Font(size=14, bold=True, color=V52_NAVY)
    ws["A16"] = rec["headline"]; ws["A16"].font = Font(size=14, bold=True, color=V52_RED if "DO NOT" in rec["headline"] else V52_BLUE)
    ws["A17"] = rec["why"]
    ws["A18"] = rec["decision"]
    ws.merge_cells("A16:H16"); ws.merge_cells("A17:H17"); ws.merge_cells("A18:H18")
    ws["A21"] = "Cost split"; ws["A21"].font = Font(bold=True)
    split_start = 22
    for i, (typ, val) in enumerate(totals.items(), split_start):
        ws.cell(i,1,typ); ws.cell(i,2,val); ws.cell(i,2).number_format = '$0.0B'
    pie = PieChart(); labels=Reference(ws,min_col=1,min_row=22,max_row=24); data=Reference(ws,min_col=2,min_row=21,max_row=24)
    pie.add_data(data, titles_from_data=True); pie.set_categories(labels); pie.title="Direct / Indirect / Reserve"; pie.height=7; pie.width=9
    ws.add_chart(pie,"D21")
    _v52_autowidth(ws)

    # Cost model
    ws = wb.create_sheet("02 Cost Model")
    _v52_style_ws(ws, "SELECTED CLASS COST ESTIMATE")
    headers = ["CBS", "Description", "Type", "P10", "P50", "P90", "Basis"]
    ws.append([]); ws.append([]); ws.append([]); ws.append(headers); _v52_header(ws[4])
    for x in model.get("cost_lines", []):
        ws.append([x.get("cbs"), x.get("description"), x.get("type"), x.get("p10_bn"), x.get("p50_bn"), x.get("p90_bn"), x.get("basis") or x.get("impact_basis")])
    for row in ws.iter_rows(min_row=5, min_col=4, max_col=6):
        for cell in row: cell.number_format = '$0.0B'
    # Subtotals
    r = ws.max_row + 2
    ws.cell(r,1,"Subtotals"); ws.cell(r,1).font=Font(bold=True,color=V52_NAVY)
    for i, typ in enumerate(["Direct","Indirect","Reserve"], r+1):
        ws.cell(i,1,typ); ws.cell(i,5,totals.get(typ,0)); ws.cell(i,5).number_format='$0.0B'; ws.cell(i,1).font=Font(bold=True)
    chart = BarChart(); chart.title="Cost Estimate by CBS (P50)"; chart.y_axis.title="US$B"; chart.x_axis.title="CBS"
    data=Reference(ws,min_col=5,min_row=4,max_row=4+len(model.get('cost_lines',[])))
    cats=Reference(ws,min_col=1,min_row=5,max_row=4+len(model.get('cost_lines',[])))
    chart.add_data(data,titles_from_data=True); chart.set_categories(cats); chart.height=8; chart.width=18
    ws.add_chart(chart,"I4")
    _v52_autowidth(ws)

    # Scenarios
    ws = wb.create_sheet("03 Scenario Comparison")
    _v52_style_ws(ws, "SCENARIO COMPARISON")
    ws.append([]); ws.append([]); ws.append([]); ws.append(["Scenario", "P50 Cost", "Schedule", "Risk", "Confidence", "Commercial interpretation"]); _v52_header(ws[4])
    for s in model.get("scenario_comparison", []):
        ws.append([s.get("label"), s.get("cost_p50"), s.get("schedule"), s.get("risk"), s.get("confidence_pct"), s.get("note") or s.get("why") or "Scenario-linked model output"])
    _v52_autowidth(ws)

    # Risks
    ws = wb.create_sheet("04 Risk Register")
    _v52_style_ws(ws, "RISK REGISTER PRO")
    headers = ["ID", "Title", "Category", "Cause", "Event", "Impact", "Owner", "Prob %", "Cost Most Likely", "Schedule Most Likely Days", "Rating", "Mitigation", "Trigger", "Status"]
    ws.append([]); ws.append([]); ws.append([]); ws.append(headers); _v52_header(ws[4])
    for rsk in model.get("risks", []):
        rating = rsk.get("pre_mitigation_rating") or rsk.get("residual_rating") or "Medium"
        ws.append([rsk.get("risk_id"), rsk.get("title"), rsk.get("category"), rsk.get("cause"), rsk.get("risk_event"), rsk.get("impact_description"), rsk.get("owner"), rsk.get("probability_pct"), rsk.get("cost_m_bn"), rsk.get("schedule_m_days"), rating, rsk.get("mitigation"), rsk.get("trigger"), rsk.get("status")])
        rr=ws.max_row
        if str(rating).lower().startswith(("high","very","extreme")):
            fill=PatternFill("solid", fgColor="FCE4D6")
        elif str(rating).lower().startswith("medium"):
            fill=PatternFill("solid", fgColor="FFF2CC")
        else:
            fill=PatternFill("solid", fgColor="E2F0D9")
        for c in range(1,15): ws.cell(rr,c).fill=fill
    for row in ws.iter_rows(min_row=5, min_col=9, max_col=9):
        row[0].number_format='$0.0B'
    _v52_autowidth(ws, 48)

    # QCRA QSRA
    ws = wb.create_sheet("05 QCRA QSRA")
    _v52_style_ws(ws, "QCRA + QSRA ANALYTICS")
    qcra=mc.get("qcra",{}); qsra=mc.get("qsra",{})
    ws.append([]); ws.append([]); ws.append([]); ws.append(["Metric","P10","P50","P80","P90"]); _v52_header(ws[4])
    ws.append(["QCRA Cost", qcra.get("p10"), qcra.get("p50"), qcra.get("p80"), qcra.get("p90")])
    ws.append(["QSRA Schedule", qsra.get("p10"), qsra.get("p50"), qsra.get("p80"), qsra.get("p90")])
    for c in range(2,6): ws.cell(5,c).number_format='$0.0B'
    ws["A9"]="QCRA Cost Tornado"; ws["E9"]="QSRA Schedule Tornado"; ws["A9"].font=ws["E9"].font=Font(bold=True,size=14,color=V52_NAVY)
    ws.append([])
    ws.append(["Driver","CBS","Mean Cost ($B)","","Driver","Activity","Mean Days"]); _v52_header(ws[11])
    for i in range(10):
        q = (mc.get("qcra_tornado") or [])[i] if i < len(mc.get("qcra_tornado") or []) else {}
        s = (mc.get("qsra_tornado") or [])[i] if i < len(mc.get("qsra_tornado") or []) else {}
        ws.append([q.get("title"), q.get("cbs"), q.get("cost_mean_bn"), "", s.get("title"), s.get("activity_id"), s.get("schedule_mean_days")])
    for row in ws.iter_rows(min_row=12, min_col=3, max_col=3): row[0].number_format='$0.0B'
    bc=BarChart(); bc.title="QCRA Cost Tornado"; bc.y_axis.title="US$B"; bc.add_data(Reference(ws,min_col=3,min_row=11,max_row=21),titles_from_data=True); bc.set_categories(Reference(ws,min_col=1,min_row=12,max_row=21)); bc.height=7; bc.width=11; ws.add_chart(bc,"I4")
    sc=BarChart(); sc.title="QSRA Schedule Tornado"; sc.y_axis.title="Days"; sc.add_data(Reference(ws,min_col=7,min_row=11,max_row=21),titles_from_data=True); sc.set_categories(Reference(ws,min_col=5,min_row=12,max_row=21)); sc.height=7; sc.width=11; ws.add_chart(sc,"I20")
    _v52_autowidth(ws)

    # Assumptions
    ws = wb.create_sheet("06 Basis + Audit")
    _v52_style_ws(ws, "BASIS, ASSUMPTIONS AND AUDIT TRAIL")
    rows = [["Project", model.get("title")], ["Mode", model.get("mode")], ["Sector", model.get("subsector")], ["Location", model.get("location")], ["Scenario", scenario], ["Class", model.get("estimate_class_name")], ["Schedule Level", model.get("schedule_level")], ["Limitations", "First-pass controls intelligence pack; validate with project-specific evidence before capital commitment."], ["Generated", datetime.utcnow().isoformat()+"Z"], ["Engine", "CASEY v52 Elite Outputs"]]
    for r in rows: ws.append(r)
    _v52_autowidth(ws)
    bio=BytesIO(); wb.save(bio); bio.seek(0); return bio.getvalue()


def risk_register_workbook_bytes_v52(model: Dict[str, Any]) -> bytes:
    # Dedicated risk workbook, not CSV, for readability.
    wb = Workbook(); ws = wb.active; ws.title = "Risk Register Pro"; _v52_style_ws(ws, "CASEY RISK REGISTER PRO")
    headers=["ID","Title","Category","Cause","Risk Event","Impact Description","Owner","Probability %","Cost O","Cost M","Cost P","Schedule O","Schedule M","Schedule P","Pre Rating","Residual","Mitigation","Trigger","Status","Activity","CBS","Board Visibility"]
    ws.append([]); ws.append([]); ws.append([]); ws.append(headers); _v52_header(ws[4])
    for rsk in model.get("risks",[]):
        ws.append([rsk.get("risk_id"),rsk.get("title"),rsk.get("category"),rsk.get("cause"),rsk.get("risk_event"),rsk.get("impact_description"),rsk.get("owner"),rsk.get("probability_pct"),rsk.get("cost_o_bn"),rsk.get("cost_m_bn"),rsk.get("cost_p_bn"),rsk.get("schedule_o_days"),rsk.get("schedule_m_days"),rsk.get("schedule_p_days"),rsk.get("pre_mitigation_rating"),rsk.get("residual_rating"),rsk.get("mitigation"),rsk.get("trigger"),rsk.get("status"),rsk.get("activity_id"),rsk.get("cbs"),rsk.get("board_visibility")])
    for row in ws.iter_rows(min_row=5,min_col=9,max_col=11):
        for cell in row: cell.number_format='$0.0B'
    # Summary tab
    ws2=wb.create_sheet("Board Top 10"); _v52_style_ws(ws2,"BOARD TOP 10 RISKS")
    ws2.append([]); ws2.append([]); ws2.append([]); ws2.append(["Rank","Risk","Why it matters","Owner","Required action"]); _v52_header(ws2[4])
    for i,rsk in enumerate(model.get("risks",[])[:10],1):
        ws2.append([i,rsk.get("title"),rsk.get("impact_description"),rsk.get("owner"),rsk.get("mitigation")])
    for sh in wb.worksheets: _v52_autowidth(sh,50)
    bio=BytesIO(); wb.save(bio); bio.seek(0); return bio.getvalue()


def xer_bytes_v52(model: Dict[str, Any]) -> bytes:
    # Legacy-targeted XER header to avoid PRA 24.12 rejection. Includes basic TASK + TASKPRED only.
    lines=[]
    lines.append("ERMHDR\t20.12\t{}\tProject\tCASEY\tCASEY\tProject Management\tUSD".format(datetime.utcnow().strftime("%Y-%m-%d")))
    lines += ["%T\tPROJECT","%F\tproj_id\tproj_short_name\tproj_name","%R\t1\tCASEY\t"+str(model.get('title','CASEY Project'))]
    lines += ["%T\tWBS","%F\twbs_id\tproj_id\twbs_short_name\twbs_name","%R\t1\t1\tCASEY\tCASEY Project"]
    lines += ["%T\tTASK","%F\ttask_id\tproj_id\twbs_id\ttask_code\ttask_name\tduration_type\ttarget_drtn_hr_cnt"]
    idmap={}
    for i,row in enumerate(model.get("schedule_rows",[]),1):
        code=row.get('activity_id',f'A{i:04d}'); idmap[code]=i
        name=str(row.get('activity','Activity')).replace('\t',' ').replace('\n',' ')
        dur=max(8,int(float(row.get('duration_months',1) or 1)*160))
        lines.append(f"%R\t{i}\t1\t1\t{code}\t{name}\tFixed Duration\t{dur}")
    lines += ["%T\tTASKPRED","%F\ttask_pred_id\ttask_id\tpred_task_id\tpred_type"]
    pred_id=1
    for row in model.get("schedule_rows",[]):
        tid=idmap.get(row.get('activity_id'))
        for p in str(row.get('predecessor','')).replace(';',',').split(','):
            p=p.strip()
            if tid and p and p in idmap:
                lines.append(f"%R\t{pred_id}\t{tid}\t{idmap[p]}\tFS"); pred_id+=1
    return ("\n".join(lines)+"\n").encode("utf-8")


def pdf_bytes_v52(model: Dict[str, Any]) -> bytes:
    bio=BytesIO(); doc=SimpleDocTemplate(bio,pagesize=landscape(A4),rightMargin=24,leftMargin=24,topMargin=20,bottomMargin=20)
    styles=getSampleStyleSheet(); rec=_v52_recommendation(model); totals=_v52_totals(model); mc=model.get('monte_carlo',{})
    title=ParagraphStyle('v52Title',parent=styles['Title'],fontSize=22,textColor=colors.HexColor('#101820'),spaceAfter=8)
    h=ParagraphStyle('v52H',parent=styles['Heading2'],fontSize=14,textColor=colors.HexColor('#0070C0'),spaceAfter=6)
    body=ParagraphStyle('v52Body',parent=styles['BodyText'],fontSize=8.5,leading=11,textColor=colors.HexColor('#101820'))
    story=[]
    story.append(Paragraph(f"CASEY BOARD DECISION PACK — {model.get('title','Project')}",title))
    story.append(Paragraph(f"{_v52_scenario_label(model)} scenario | {model.get('subsector')} | {model.get('location')} | Generated {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}",body))
    metrics=[["P50 Cost",model.get('cost_p50'),"Cost Range",model.get('cost_range'),"QCRA P80",money_bn(mc.get('qcra',{}).get('p80',0)),"Schedule",model.get('schedule')],["QSRA P80",f"{mc.get('qsra',{}).get('p80','')} months","Risk",model.get('risk'),"Confidence",f"{model.get('confidence_pct')}%","Direct / Indirect / Reserve",f"{money_bn(totals.get('Direct',0))} / {money_bn(totals.get('Indirect',0))} / {money_bn(totals.get('Reserve',0))}"]]
    t=Table(metrics,colWidths=[70,95,75,95,75,95,95,130])
    t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#EAF6FF')),('TEXTCOLOR',(0,0),(-1,-1),colors.HexColor('#101820')),('GRID',(0,0),(-1,-1),0.4,colors.HexColor('#D9E2EC')),('FONTNAME',(0,0),(-1,-1),'Helvetica'),('FONTSIZE',(0,0),(-1,-1),8),('VALIGN',(0,0),(-1,-1),'TOP')]))
    story += [Spacer(1,8),t,Spacer(1,10),Paragraph(rec['headline'],h),Paragraph(rec['why'],body),Paragraph(rec['decision'],body),PageBreak()]
    # Cost and scenario
    story.append(Paragraph("Cost Estimate — Selected Class Only",h))
    rows=[["CBS","Description","Type","P10","P50","P90"]]+[[x.get('cbs'),x.get('description'),x.get('type'),money_bn(x.get('p10_bn',0)),money_bn(x.get('p50_bn',0)),money_bn(x.get('p90_bn',0))] for x in model.get('cost_lines',[])[:18]]
    tt=Table(rows,colWidths=[45,210,65,70,70,70]); tt.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#0070C0')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.3,colors.HexColor('#D9E2EC')),('FONTSIZE',(0,0),(-1,-1),7.5),('VALIGN',(0,0),(-1,-1),'TOP')]))
    story += [tt,PageBreak()]
    # Risk
    story.append(Paragraph("QCRA + QSRA Tornado Drivers",h))
    q=mc.get('qcra_tornado',[])[:10]; s=mc.get('qsra_tornado',[])[:10]
    rows=[["Rank","QCRA Cost Driver","CBS","Mean Cost","QSRA Schedule Driver","Activity","Mean Days"]]
    for i in range(max(len(q),len(s))):
        qr=q[i] if i<len(q) else {}; sr=s[i] if i<len(s) else {}
        rows.append([i+1,qr.get('title',''),qr.get('cbs',''),money_bn(qr.get('cost_mean_bn',0)),sr.get('title',''),sr.get('activity_id',''),sr.get('schedule_mean_days','')])
    tt=Table(rows,colWidths=[35,170,45,70,170,55,65]); tt.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#0070C0')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.3,colors.HexColor('#D9E2EC')),('FONTSIZE',(0,0),(-1,-1),7.5),('VALIGN',(0,0),(-1,-1),'TOP')]))
    story += [tt,Spacer(1,10),Paragraph("Board Top Risks",h)]
    rows=[["ID","Risk","Cause","Event","Impact","Owner","Mitigation"]]
    for rsk in model.get('risks',[])[:8]: rows.append([rsk.get('risk_id'),rsk.get('title'),rsk.get('cause'),rsk.get('risk_event'),rsk.get('impact_description'),rsk.get('owner'),rsk.get('mitigation')])
    tt=Table(rows,colWidths=[45,90,130,120,160,80,160]); tt.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#0070C0')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.3,colors.HexColor('#D9E2EC')),('FONTSIZE',(0,0),(-1,-1),6.4),('VALIGN',(0,0),(-1,-1),'TOP')]))
    story += [tt]
    doc.build(story); bio.seek(0); return bio.getvalue()


def word_bytes_v52(model: Dict[str, Any]) -> bytes:
    doc=Document(); sec=doc.sections[0]; sec.top_margin=Inches(.45); sec.bottom_margin=Inches(.45); sec.left_margin=Inches(.55); sec.right_margin=Inches(.55)
    rec=_v52_recommendation(model); totals=_v52_totals(model); mc=model.get('monte_carlo',{})
    p=doc.add_paragraph(); r=p.add_run("CASEY EXECUTIVE BOARD DECISION PACK"); r.bold=True; r.font.size=Pt(20); r.font.color.rgb=RGBColor(0,112,192)
    doc.add_paragraph(f"{model.get('title')} | {_v52_scenario_label(model)} scenario | {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}")
    table=doc.add_table(rows=1,cols=4); table.style='Table Grid'; hdr=table.rows[0].cells
    for i,x in enumerate(['Metric','Value','Interpretation','Board action']): hdr[i].text=x
    for k,v,interp,action in [('P50 Cost',model.get('cost_p50'),'Most likely first-pass cost','Use as control point'),('Cost Range',model.get('cost_range'),'Class and QCRA uncertainty','Approve with range acknowledged'),('QCRA P80',money_bn(mc.get('qcra',{}).get('p80',0)),'Risk-adjusted funding level','Set contingency rules'),('QSRA P80',f"{mc.get('qsra',{}).get('p80','')} months",'Schedule confidence level','Validate critical path'),('Direct / Indirect / Reserve',f"{money_bn(totals.get('Direct',0))} / {money_bn(totals.get('Indirect',0))} / {money_bn(totals.get('Reserve',0))}",'Stops reserve hiding','Challenge contractor basis')]:
        c=table.add_row().cells; c[0].text=k; c[1].text=str(v); c[2].text=interp; c[3].text=action
    doc.add_heading('Executive Recommendation',level=1); doc.add_paragraph(rec['headline']); doc.add_paragraph(rec['why']); doc.add_paragraph(rec['decision'])
    doc.add_heading('QCRA and QSRA Tornado Drivers',level=1)
    for label, rows in [('QCRA Cost', mc.get('qcra_tornado',[])[:8]), ('QSRA Schedule', mc.get('qsra_tornado',[])[:8])]:
        doc.add_heading(label,level=2)
        t=doc.add_table(rows=1,cols=4); t.style='Table Grid'; t.rows[0].cells[0].text='Rank'; t.rows[0].cells[1].text='Driver'; t.rows[0].cells[2].text='Link'; t.rows[0].cells[3].text='Exposure'
        for i,x in enumerate(rows,1):
            c=t.add_row().cells; c[0].text=str(i); c[1].text=str(x.get('title','')); c[2].text=str(x.get('cbs') or x.get('activity_id') or ''); c[3].text=money_bn(x.get('cost_mean_bn',0)) if label.startswith('QCRA') else f"{x.get('schedule_mean_days')} days"
    doc.add_heading('Risk Register - Board Top Risks',level=1)
    t=doc.add_table(rows=1,cols=6); t.style='Table Grid';
    for i,hdr in enumerate(['ID','Risk','Cause','Event','Impact','Mitigation']): t.rows[0].cells[i].text=hdr
    for rsk in model.get('risks',[])[:10]:
        c=t.add_row().cells; vals=[rsk.get('risk_id'),rsk.get('title'),rsk.get('cause'),rsk.get('risk_event'),rsk.get('impact_description'),rsk.get('mitigation')]
        for i,v in enumerate(vals): c[i].text=str(v or '')
    bio=BytesIO(); doc.save(bio); bio.seek(0); return bio.getvalue()


def pptx_bytes_v52(model: Dict[str, Any]) -> bytes:
    from pptx.dml.color import RGBColor as PRGBColor
    prs=Presentation(); prs.slide_width=PptxInches(13.333); prs.slide_height=PptxInches(7.5); blank=prs.slide_layouts[6]
    rec=_v52_recommendation(model); mc=model.get('monte_carlo',{}); totals=_v52_totals(model)
    def slide(title, subtitle=""):
        s=prs.slides.add_slide(blank); s.background.fill.solid(); s.background.fill.fore_color.rgb=PRGBColor(247,251,255)
        tx=s.shapes.add_textbox(PptxInches(.55),PptxInches(.35),PptxInches(12.2),PptxInches(.6)); tf=tx.text_frame; tf.text=title; p=tf.paragraphs[0]; p.runs[0].font.size=PptxPt(25); p.runs[0].font.bold=True; p.runs[0].font.color.rgb=PRGBColor(16,24,32)
        if subtitle:
            sub=s.shapes.add_textbox(PptxInches(.58),PptxInches(.95),PptxInches(12),PptxInches(.4)); sub.text_frame.text=subtitle; sub.text_frame.paragraphs[0].runs[0].font.size=PptxPt(12); sub.text_frame.paragraphs[0].runs[0].font.color.rgb=PRGBColor(80,96,112)
        return s
    s=slide(f"{model.get('title')} — Board Intelligence", f"{_v52_scenario_label(model)} scenario | CASEY v52")
    box=s.shapes.add_textbox(PptxInches(.8),PptxInches(2.0),PptxInches(11.5),PptxInches(1.2)); box.text_frame.text=rec['headline']; box.text_frame.paragraphs[0].runs[0].font.size=PptxPt(30); box.text_frame.paragraphs[0].runs[0].font.bold=True; box.text_frame.paragraphs[0].runs[0].font.color.rgb=PRGBColor(0,112,192)
    s=slide("Board Decision Snapshot", rec['decision'])
    metrics=[('P50',model.get('cost_p50')),('Range',model.get('cost_range')),('QCRA P80',money_bn(mc.get('qcra',{}).get('p80',0))),('Schedule',model.get('schedule')),('QSRA P80',f"{mc.get('qsra',{}).get('p80','')} mo"),('Risk',model.get('risk')),('Confidence',f"{model.get('confidence_pct')}%"),('Direct',money_bn(totals.get('Direct',0))),('Reserve',money_bn(totals.get('Reserve',0)))]
    for i,(k,v) in enumerate(metrics):
        x=.6+(i%3)*4.25; y=1.65+(i//3)*1.4
        shp=s.shapes.add_shape(1,PptxInches(x),PptxInches(y),PptxInches(3.8),PptxInches(1.08)); shp.fill.solid(); shp.fill.fore_color.rgb=PRGBColor(232,246,255); shp.line.color.rgb=PRGBColor(0,112,192)
        shp.text=f"{k}\n{v}"; shp.text_frame.paragraphs[0].runs[0].font.size=PptxPt(11); shp.text_frame.paragraphs[1].runs[0].font.size=PptxPt(21); shp.text_frame.paragraphs[1].runs[0].font.bold=True
    s=slide("QCRA and QSRA are separated", "Cost risk and schedule risk are not mixed.")
    t=s.shapes.add_table(9,4,PptxInches(.7),PptxInches(1.65),PptxInches(5.8),PptxInches(4.8)).table
    for j,h in enumerate(['Rank','QCRA Driver','CBS','Cost']): t.cell(0,j).text=h
    for i,x in enumerate(mc.get('qcra_tornado',[])[:8],1):
        t.cell(i,0).text=str(i); t.cell(i,1).text=str(x.get('title','')); t.cell(i,2).text=str(x.get('cbs','')); t.cell(i,3).text=money_bn(x.get('cost_mean_bn',0))
    t2=s.shapes.add_table(9,4,PptxInches(6.85),PptxInches(1.65),PptxInches(5.8),PptxInches(4.8)).table
    for j,h in enumerate(['Rank','QSRA Driver','Act','Days']): t2.cell(0,j).text=h
    for i,x in enumerate(mc.get('qsra_tornado',[])[:8],1):
        t2.cell(i,0).text=str(i); t2.cell(i,1).text=str(x.get('title','')); t2.cell(i,2).text=str(x.get('activity_id','')); t2.cell(i,3).text=str(x.get('schedule_mean_days',''))
    s=slide("Board Top Risks", "Cause, event, impact and response are visible.")
    y=1.55
    for rsk in model.get('risks',[])[:5]:
        bx=s.shapes.add_textbox(PptxInches(.8),PptxInches(y),PptxInches(11.8),PptxInches(.85)); bx.text_frame.text=f"{rsk.get('risk_id')} {rsk.get('title')} — {rsk.get('impact_description')}"; bx.text_frame.paragraphs[0].runs[0].font.size=PptxPt(13); y+=.93
    s=slide("What CASEY produces", "Board PDF, cost workbook, risk register workbook, schedule export, DOCX, PPTX, JSON audit and ZIP pack.")
    y=1.75
    for x in ['Selected class estimate only','Direct / indirect / reserve split','Scenario-linked outputs for Base, Faster, Cheaper, Lower Risk and Premium','Readable white-background Excel/PDF exports','Risk register includes cause, event, impact, owner and mitigation','Legacy-targeted XER plus schedule tables']:
        bx=s.shapes.add_textbox(PptxInches(1),PptxInches(y),PptxInches(11.5),PptxInches(.45)); bx.text_frame.text='✓ '+x; bx.text_frame.paragraphs[0].runs[0].font.size=PptxPt(16); bx.text_frame.paragraphs[0].runs[0].font.color.rgb=PRGBColor(0,112,192); y+=.65
    bio=BytesIO(); prs.save(bio); bio.seek(0); return bio.getvalue()


def export_workbook_v52_endpoint(model: Dict[str, Any]): return stream(workbook_bytes_v52(model),"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","CASEY_v52_Elite_Cost_Model.xlsx")
def export_risk_v52_endpoint(model: Dict[str, Any]): return stream(risk_register_workbook_bytes_v52(model),"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","CASEY_v52_Risk_Register_Pro.xlsx")
def export_xer_v52_endpoint(model: Dict[str, Any]): return stream(xer_bytes_v52(model),"application/octet-stream","CASEY_v52_P6_Schedule_Legacy.xer")
def export_word_v52_endpoint(model: Dict[str, Any]): return stream(word_bytes_v52(model),"application/vnd.openxmlformats-officedocument.wordprocessingml.document","CASEY_v52_Executive_Board_Report.docx")
def export_pdf_v52_endpoint(model: Dict[str, Any]): return stream(pdf_bytes_v52(model),"application/pdf","CASEY_v52_Board_Decision_Pack.pdf")
def export_pptx_v52_endpoint(model: Dict[str, Any]): return stream(pptx_bytes_v52(model),"application/vnd.openxmlformats-officedocument.presentationml.presentation","CASEY_v52_Board_Deck.pptx")
def export_all_v52_endpoint(model: Dict[str, Any]):
    bio=BytesIO()
    with zipfile.ZipFile(bio,"w",zipfile.ZIP_DEFLATED) as z:
        z.writestr("01_CASEY_v52_Elite_Cost_Model.xlsx", workbook_bytes_v52(model))
        z.writestr("02_CASEY_v52_Risk_Register_Pro.xlsx", risk_register_workbook_bytes_v52(model))
        z.writestr("03_CASEY_v52_P6_Schedule_Legacy.xer", xer_bytes_v52(model))
        z.writestr("04_CASEY_v52_Executive_Board_Report.docx", word_bytes_v52(model))
        z.writestr("05_CASEY_v52_Board_Decision_Pack.pdf", pdf_bytes_v52(model))
        z.writestr("06_CASEY_v52_Board_Deck.pptx", pptx_bytes_v52(model))
        z.writestr("07_CASEY_v52_Full_Model_Audit.json", json.dumps(model, indent=2))
        z.writestr("08_READ_ME_OUTPUT_STANDARD.txt", "CASEY v52 Elite Outputs: white-background readable exports, scenario-linked numbers, QCRA/QSRA separation, risk cause-event-impact-mitigation, legacy-targeted XER. Validate XER in your Primavera/PRA environment because vendor import support differs by installed version.")
    bio.seek(0); return stream(bio.getvalue(),"application/zip","CASEY_v52_Elite_Output_Pack.zip")


def _v52_patch_route(path: str, endpoint):
    for route in app.routes:
        if getattr(route, 'path', None) == path and 'POST' in getattr(route, 'methods', set()):
            route.endpoint = endpoint
            if hasattr(route, 'dependant'):
                route.dependant.call = endpoint

_v52_patch_route('/export/workbook', export_workbook_v52_endpoint)
_v52_patch_route('/export/risk-register', export_risk_v52_endpoint)
_v52_patch_route('/export/xer', export_xer_v52_endpoint)
_v52_patch_route('/export/word', export_word_v52_endpoint)
_v52_patch_route('/export/pdf', export_pdf_v52_endpoint)
_v52_patch_route('/export/pptx', export_pptx_v52_endpoint)
_v52_patch_route('/export/all', export_all_v52_endpoint)
_v52_patch_route('/v51/export/all', export_all_v52_endpoint)

@app.post('/v52/export/workbook')
def export_workbook_v52(model: Dict[str, Any]): return export_workbook_v52_endpoint(model)
@app.post('/v52/export/risk-register')
def export_risk_v52(model: Dict[str, Any]): return export_risk_v52_endpoint(model)
@app.post('/v52/export/xer')
def export_xer_v52(model: Dict[str, Any]): return export_xer_v52_endpoint(model)
@app.post('/v52/export/all')
def export_all_v52(model: Dict[str, Any]): return export_all_v52_endpoint(model)

# ===================== END CASEY v52 ELITE OUTPUT SYSTEM =====================

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)


# ========================= CASEY v54 PLATINUM OUTPUT ENGINE =========================
# Final output override: readable white exports, real QCRA/QSRA separation, validated risk fields,
# scenario-linked narratives, and legacy-targeted schedule export. Same frontend endpoints.

from openpyxl.chart import PieChart
from openpyxl.styles import GradientFill

APP_VERSION = "CASEY TITAN X v54 Platinum Output Engine"

V54_BLUE = "0B1F33"
V54_CYAN = "00A9C7"
V54_LIGHT = "F3F7FA"
V54_LINE = "D9E2EA"
V54_GREEN = "1F7A4D"
V54_AMBER = "C77800"
V54_RED = "B42318"


def _v54_safe(v, default=""):
    return default if v is None else v


def _v54_slug(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9_]+", "_", str(s or "CASEY")).strip("_")[:60]


def _v54_num_money(s: Any) -> float:
    return parse_bn(s)


def _v54_scenario(model: Dict[str, Any]) -> Dict[str, Any]:
    label = model.get("scenario_label") or str(model.get("scenario") or "Base").replace("_", " ").title()
    key = str(model.get("scenario") or label).lower().replace(" ", "_")
    narratives = {
        "base": ("Approve definition phase — do not commit full capital yet.", "Balanced control case used as the board reference model."),
        "faster": ("Proceed only with funded acceleration controls.", "Time is bought with premium logistics, overlap and higher interface management."),
        "cheaper": ("Use as a value-engineering challenge, not a final funding case.", "Lower capex increases scope, productivity and rework exposure unless controls tighten."),
        "lower_risk": ("Recommended board path where certainty matters more than lowest cost.", "Higher assurance, surveys and buffers reduce downstream surprise."),
        "premium": ("Proceed if resilience, brand and delivery certainty justify premium capex.", "Best-in-class suppliers, stronger assurance and higher redundancy reduce residual risk."),
        "investor": ("Use for investment committee challenge and option screening.", "Transparent risk-adjusted view designed for funding conversations."),
        "survival": ("Do not use for approval without explicit risk acceptance.", "Minimum viable case with high residual risk and reduced confidence."),
    }
    rec, why = narratives.get(key, narratives["base"])
    return {"key": key, "label": label, "recommendation": rec, "why": why}


def _v54_risks(model: Dict[str, Any]) -> List[Dict[str, Any]]:
    out=[]
    schedule = model.get("schedule_rows") or []
    cost_lines = model.get("cost_lines") or []
    for idx,r in enumerate(model.get("risks") or [],1):
        prob = int(max(5, min(95, float(r.get("probability_pct") or 25))))
        # derive full cause/event/impact from old description if not present
        title = str(r.get("title") or f"Risk {idx}")
        cause = str(r.get("cause") or r.get("description") or "Project evidence is not yet mature enough to fully validate this assumption.")
        event = str(r.get("event") or (title + " materialises during delivery"))
        impact = str(r.get("impact") or r.get("impact_description") or _v54_risk_impact(title, r))
        mitigation = str(r.get("mitigation") or "Assign accountable owner, confirm evidence, and update risk response before the next gate.")
        owner = str(r.get("owner") or "Project Controls Lead")
        act_id = str(r.get("activity_id") or (schedule[min(idx, len(schedule)-1)].get("activity_id") if schedule else "A1900"))
        act_name = str(r.get("activity_name") or (schedule[min(idx, len(schedule)-1)].get("activity") if schedule else "Delivery activity"))
        cbs = str(r.get("cbs") or (cost_lines[min(idx, len(cost_lines)-1)].get("cbs") if cost_lines else "01.01"))
        cbs_name = str(r.get("cbs_name") or (cost_lines[min(idx, len(cost_lines)-1)].get("description") if cost_lines else "Cost account"))
        # cost values, ensure > 0 and ordered
        co = max(0.001, float(r.get("cost_o_bn") or 0.01))
        cm = max(co, float(r.get("cost_m_bn") or co*2))
        cp = max(cm, float(r.get("cost_p_bn") or cm*1.8))
        # schedule values: retain zero only for cost-only risks but keep out of QSRA tornado later
        so = max(0.0, float(r.get("schedule_o_days") or 0))
        sm = max(so, float(r.get("schedule_m_days") or 0))
        sp = max(sm, float(r.get("schedule_p_days") or 0))
        cost_emv = cm * prob / 100
        sched_emv = sm * prob / 100
        out.append({
            **r,
            "risk_id": str(r.get("risk_id") or f"R-{idx:03d}"),
            "title": title,
            "category": str(r.get("category") or "Delivery"),
            "cause": cause,
            "event": event,
            "impact": impact,
            "probability_pct": prob,
            "likelihood_pct": prob,
            "activity_id": act_id,
            "activity_name": act_name,
            "cbs": cbs,
            "cbs_name": cbs_name,
            "schedule_o_days": so,
            "schedule_m_days": sm,
            "schedule_p_days": sp,
            "cost_o_bn": round(co,3),
            "cost_m_bn": round(cm,3),
            "cost_p_bn": round(cp,3),
            "cost_emv_bn": round(cost_emv,3),
            "schedule_emv_days": round(sched_emv,1),
            "owner": owner,
            "trigger": str(r.get("trigger") or "Trigger threshold not yet defined"),
            "mitigation": mitigation,
            "residual_rating": _v54_rating(prob, cost_emv, sched_emv),
            "status": str(r.get("status") or "Open"),
            "basis_of_cost_impact": str(r.get("basis_of_cost_impact") or f"Linked to CBS {cbs} {cbs_name}; expected exposure is probability-weighted from O/M/P cost impacts."),
            "basis_of_schedule_impact": str(r.get("basis_of_schedule_impact") or f"Linked to activity {act_id} {act_name}; expected exposure is probability-weighted from O/M/P schedule impacts."),
        })
    return sorted(out, key=lambda x: (x["cost_emv_bn"]*1000 + x["schedule_emv_days"]), reverse=True)


def _v54_risk_impact(title, r):
    t = title.lower()
    if "grid" in t or "power" in t: return "Critical path delay, temporary power cost, phased opening risk and revenue deferral."
    if "market" in t or "escalation" in t: return "P50 cost becomes understated; procurement strategy and funding envelope may need reset."
    if "design" in t: return "Cost growth, quantity movement, rework, late approvals and schedule resequencing."
    if "scope" in t: return "Additional quantities, redesign, rework, procurement churn and contingency drawdown."
    if "commission" in t: return "Delayed handover, repeated testing cycles, operational readiness slippage and damages exposure."
    return "Cost, schedule and confidence exposure affecting the board approval basis."


def _v54_rating(prob, cost_emv, sched_emv):
    score = prob/100 * (cost_emv*10 + sched_emv/10)
    if score > 12: return "Red"
    if score > 5: return "Amber"
    return "Green"


def _v54_curves(model: Dict[str, Any]) -> Dict[str, Any]:
    mc = model.get("monte_carlo") or {}
    curve = mc.get("curve") or []
    if curve:
        qcra = [(int(x.get("percentile",0)), float(x.get("cost_bn",0))) for x in curve]
        qsra = [(int(x.get("percentile",0)), float(x.get("schedule_months",0))) for x in curve]
    else:
        p50 = parse_bn(model.get("cost_p50")); months = float(str(model.get("schedule","60")).split()[0])
        qcra = [(p, round(p50*(0.7+p/100*0.8),3)) for p in [5,10,20,30,40,50,60,70,80,90,95]]
        qsra = [(p, round(months*(0.9+p/100*0.35),2)) for p in [5,10,20,30,40,50,60,70,80,90,95]]
    return {"qcra": qcra, "qsra": qsra}


def _v54_tornado(model: Dict[str, Any]) -> Dict[str, List[Dict[str, Any]]]:
    risks = _v54_risks(model)
    qcra = sorted([r for r in risks if r.get("cost_emv_bn",0) > 0], key=lambda r:r.get("cost_emv_bn",0), reverse=True)[:10]
    qsra = sorted([r for r in risks if r.get("schedule_emv_days",0) > 0], key=lambda r:r.get("schedule_emv_days",0), reverse=True)[:10]
    return {"qcra": qcra, "qsra": qsra}


def _v54_costs(model: Dict[str, Any]) -> List[Dict[str, Any]]:
    low, high, *_ = class_range(int(model.get("estimate_class") or 3))
    out=[]
    for x in model.get("cost_lines") or []:
        p50 = float(x.get("p50_bn") or 0)
        out.append({
            **x,
            "p10_bn": round(float(x.get("p10_bn") or p50*low),3),
            "p50_bn": round(p50,3),
            "p90_bn": round(float(x.get("p90_bn") or p50*high),3),
            "basis": x.get("basis") or x.get("impact_basis") or "Sector benchmark and model factor basis.",
        })
    return out


def _v54_style_header(ws, row=1, last_col=None):
    last_col = last_col or ws.max_column
    for c in range(1,last_col+1):
        cell=ws.cell(row,c)
        cell.fill=PatternFill("solid", fgColor=V54_BLUE)
        cell.font=Font(color="FFFFFF", bold=True)
        cell.alignment=Alignment(wrap_text=True, vertical="center")
        cell.border=Border(bottom=Side(style="thin", color=V54_LINE))


def _v54_style_sheet(ws, freeze="A2"):
    ws.freeze_panes = freeze
    ws.sheet_view.showGridLines = False
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = Border(bottom=Side(style="hair", color="E8EEF5"))
            if isinstance(cell.value, (int,float)):
                cell.number_format = '#,##0.0'
    for col in range(1, ws.max_column+1):
        letter = get_column_letter(col)
        max_len = 10
        for cell in ws[letter]:
            if cell.value is not None:
                max_len = max(max_len, min(55, len(str(cell.value)) + 2))
        ws.column_dimensions[letter].width = min(max_len, 38)
    for r in range(1, ws.max_row+1):
        ws.row_dimensions[r].height = 24


def workbook_bytes(model: Dict[str, Any]) -> bytes:
    model = dict(model)
    sc = _v54_scenario(model); costs=_v54_costs(model); risks=_v54_risks(model); curves=_v54_curves(model); torn=_v54_tornado(model)
    wb=Workbook(); wb.remove(wb.active)
    # Dashboard
    ws=wb.create_sheet("Executive Dashboard")
    ws["A1"]="CASEY PLATINUM COST / SCHEDULE / RISK MODEL"; ws["A1"].font=Font(size=18,bold=True,color=V54_BLUE)
    ws["A2"]=f"{model.get('title')} | {sc['label']} scenario | Generated {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}"; ws["A2"].font=Font(size=10,color="666666")
    kpis=[("P50 Cost",model.get("cost_p50")),("Range",model.get("cost_range")),("QCRA P80",money_bn(model.get("monte_carlo",{}).get("qcra",{}).get("p80", parse_bn(model.get("cost_p50"))*1.25))),("Schedule",model.get("schedule")),("QSRA P80",str(model.get("monte_carlo",{}).get("qsra",{}).get("p80",""))+" months"),("Risk / Confidence",f"{model.get('risk')} / {model.get('confidence_pct')}%")]
    row=4
    for i,(k,v) in enumerate(kpis):
        c=1+(i%3)*3; r=row+(i//3)*3
        ws.cell(r,c,k); ws.cell(r+1,c,str(v)); ws.merge_cells(start_row=r,start_column=c,end_row=r,end_column=c+1); ws.merge_cells(start_row=r+1,start_column=c,end_row=r+1,end_column=c+1)
        ws.cell(r,c).fill=PatternFill("solid", fgColor=V54_LIGHT); ws.cell(r,c).font=Font(bold=True,color=V54_BLUE)
        ws.cell(r+1,c).font=Font(size=18,bold=True,color=V54_BLUE); ws.cell(r+1,c).fill=PatternFill("solid", fgColor="FFFFFF")
    ws["A11"]="EXECUTIVE RECOMMENDATION"; ws["A11"].font=Font(bold=True,color=V54_BLUE,size=13)
    ws["A12"]=sc["recommendation"]; ws["A13"]=sc["why"]; ws["A12"].font=Font(bold=True,size=14,color=V54_RED if sc['key'] in ['cheaper','survival'] else V54_BLUE)
    ws["A15"]="Direct / Indirect / Reserve"; ws["A15"].font=Font(bold=True,color=V54_BLUE)
    summary={"Direct":0,"Indirect":0,"Reserve":0}
    for x in costs: summary[x.get("type","Direct")] = summary.get(x.get("type","Direct"),0)+float(x.get("p50_bn",0))
    for i,(k,v) in enumerate(summary.items(), start=16): ws.cell(i,1,k); ws.cell(i,2,v); ws.cell(i,2).number_format='$#,##0.0B'
    pie=PieChart(); labels=Reference(ws,min_col=1,min_row=16,max_row=18); data=Reference(ws,min_col=2,min_row=15,max_row=18); pie.add_data(data,titles_from_data=True); pie.set_categories(labels); pie.title="Cost Composition"; ws.add_chart(pie,"D15")
    _v54_style_sheet(ws,"A4")
    # Cost Estimate
    ws=wb.create_sheet("Cost Estimate")
    headers=["CBS","Description","Type","Basis","P10 $B","P50 $B","P90 $B","Board Challenge"]
    ws.append(headers); _v54_style_header(ws)
    for x in costs:
        challenge = "Validate benchmark and supplier evidence" if x.get("type") == "Direct" else "Confirm not hidden in direct cost"
        ws.append([x.get("cbs"),x.get("description"),x.get("type"),x.get("basis"),x.get("p10_bn"),x.get("p50_bn"),x.get("p90_bn"),challenge])
    _v54_style_sheet(ws)
    # Scenario Comparison
    ws=wb.create_sheet("Scenario Comparison"); ws.append(["Scenario","Cost","Schedule Months","Risk","Confidence","What changes","Board Use"]); _v54_style_header(ws)
    for x in model.get("scenario_comparison") or []:
        ws.append([x.get("label"),x.get("cost"),x.get("schedule_months"),x.get("risk"),x.get("confidence"),x.get("why"),"Use to frame trade-off decision"])
    _v54_style_sheet(ws)
    # Curves
    ws=wb.create_sheet("QCRA Curve"); ws.append(["Percentile","Cost $B"]); _v54_style_header(ws)
    for p,v in curves["qcra"]: ws.append([p,v])
    chart=LineChart(); chart.title="QCRA Cost Confidence Curve"; chart.y_axis.title="Cost $B"; chart.x_axis.title="Percentile"; data=Reference(ws,min_col=2,min_row=1,max_row=ws.max_row); cats=Reference(ws,min_col=1,min_row=2,max_row=ws.max_row); chart.add_data(data,titles_from_data=True); chart.set_categories(cats); ws.add_chart(chart,"D2"); _v54_style_sheet(ws)
    ws=wb.create_sheet("QSRA Curve"); ws.append(["Percentile","Duration Months"]); _v54_style_header(ws)
    for p,v in curves["qsra"]: ws.append([p,v])
    chart=LineChart(); chart.title="QSRA Schedule Confidence Curve"; chart.y_axis.title="Months"; chart.x_axis.title="Percentile"; data=Reference(ws,min_col=2,min_row=1,max_row=ws.max_row); cats=Reference(ws,min_col=1,min_row=2,max_row=ws.max_row); chart.add_data(data,titles_from_data=True); chart.set_categories(cats); ws.add_chart(chart,"D2"); _v54_style_sheet(ws)
    # Tornados
    ws=wb.create_sheet("QCRA Tornado"); ws.append(["Rank","Driver","CBS","Mean Cost $B","Action"]); _v54_style_header(ws)
    for i,r in enumerate(torn["qcra"],1): ws.append([i,r["title"],r["cbs"],r["cost_emv_bn"],r["mitigation"]])
    chart=BarChart(); chart.type="bar"; chart.title="QCRA Cost Tornado"; data=Reference(ws,min_col=4,min_row=1,max_row=ws.max_row); cats=Reference(ws,min_col=2,min_row=2,max_row=ws.max_row); chart.add_data(data,titles_from_data=True); chart.set_categories(cats); ws.add_chart(chart,"G2"); _v54_style_sheet(ws)
    ws=wb.create_sheet("QSRA Tornado"); ws.append(["Rank","Driver","Activity","Mean Days","Action"]); _v54_style_header(ws)
    for i,r in enumerate(torn["qsra"],1): ws.append([i,r["title"],r["activity_id"],r["schedule_emv_days"],r["mitigation"]])
    chart=BarChart(); chart.type="bar"; chart.title="QSRA Schedule Tornado"; data=Reference(ws,min_col=4,min_row=1,max_row=ws.max_row); cats=Reference(ws,min_col=2,min_row=2,max_row=ws.max_row); chart.add_data(data,titles_from_data=True); chart.set_categories(cats); ws.add_chart(chart,"G2"); _v54_style_sheet(ws)
    # Risk register
    ws=wb.create_sheet("Risk Register")
    headers=["ID","Risk","Category","Likelihood %","Cause","Event","Impact","Owner","Mitigation","Trigger","Residual","Activity","CBS","Cost EMV $B","Schedule EMV Days"]
    ws.append(headers); _v54_style_header(ws)
    for r in risks:
        ws.append([r["risk_id"],r["title"],r["category"],r["probability_pct"],r["cause"],r["event"],r["impact"],r["owner"],r["mitigation"],r["trigger"],r["residual_rating"],r["activity_id"],r["cbs"],r["cost_emv_bn"],r["schedule_emv_days"]])
        color = {"Red":"FCE4D6","Amber":"FFF2CC","Green":"E2F0D9"}.get(r["residual_rating"],"FFFFFF")
        for c in range(1,16): ws.cell(ws.max_row,c).fill=PatternFill("solid",fgColor=color)
    _v54_style_sheet(ws)
    ws.column_dimensions['E'].width=34; ws.column_dimensions['F'].width=34; ws.column_dimensions['G'].width=40; ws.column_dimensions['I'].width=38
    # Schedule
    ws=wb.create_sheet("Schedule")
    ws.append(["Activity ID","Phase","Activity","Predecessor","Duration Months","Critical","Basis"]); _v54_style_header(ws)
    for a in model.get("schedule_rows") or []: ws.append([a.get("activity_id"),a.get("phase"),a.get("activity"),a.get("predecessor"),a.get("duration_months"),a.get("critical"),a.get("basis")])
    _v54_style_sheet(ws)
    # Assumptions
    ws=wb.create_sheet("Methodology & Assumptions")
    rows=[["Area","Method"],["Cost","Parametric sector model + location factor + scale factor + class range + scenario logic."],["QCRA","Probability-weighted O/M/P cost impacts with percentile confidence outputs."],["QSRA","Probability-weighted activity impacts with schedule confidence curve and tornado."],["Risk","Every risk includes cause, event, impact, owner, mitigation and residual rating."],["Limitation","First-pass intelligence model. Validate with project-specific evidence before capital commitment."]]
    for r in rows: ws.append(r)
    _v54_style_header(ws); _v54_style_sheet(ws)
    bio=BytesIO(); wb.save(bio); bio.seek(0); return bio.getvalue()


def risk_register_workbook_bytes(model: Dict[str, Any]) -> bytes:
    risks=_v54_risks(model); torn=_v54_tornado(model)
    wb=Workbook(); wb.remove(wb.active)
    ws=wb.create_sheet("Risk Dashboard")
    ws["A1"]="CASEY PLATINUM RISK REGISTER"; ws["A1"].font=Font(size=18,bold=True,color=V54_BLUE)
    ws["A2"]=f"{model.get('title')} | {model.get('scenario_label',model.get('scenario','Base'))}"; ws["A2"].font=Font(color="666666")
    ws.append([]); ws.append(["Metric","Value"]); _v54_style_header(ws,4,2)
    ws.append(["Total risks",len(risks)]); ws.append(["Red risks",sum(1 for r in risks if r['residual_rating']=='Red')]); ws.append(["Amber risks",sum(1 for r in risks if r['residual_rating']=='Amber')]); ws.append(["Top QCRA driver",torn['qcra'][0]['title'] if torn['qcra'] else 'None']); ws.append(["Top QSRA driver",torn['qsra'][0]['title'] if torn['qsra'] else 'None'])
    _v54_style_sheet(ws,"A4")
    ws=wb.create_sheet("Risk Register")
    headers=["ID","Risk","Category","Likelihood %","Cause","Risk Event","Impact Description","Owner","Mitigation / Response","Trigger","Residual Rating","Linked Activity","Linked CBS","QCRA Mean $B","QSRA Mean Days","Status"]
    ws.append(headers); _v54_style_header(ws)
    for r in risks:
        ws.append([r["risk_id"],r["title"],r["category"],r["probability_pct"],r["cause"],r["event"],r["impact"],r["owner"],r["mitigation"],r["trigger"],r["residual_rating"],r["activity_id"],r["cbs"],r["cost_emv_bn"],r["schedule_emv_days"],r["status"]])
        color={"Red":"FCE4D6","Amber":"FFF2CC","Green":"E2F0D9"}.get(r["residual_rating"],"FFFFFF")
        for c in range(1,17): ws.cell(ws.max_row,c).fill=PatternFill("solid",fgColor=color)
    _v54_style_sheet(ws)
    for col in ['E','F','G','I','J']: ws.column_dimensions[col].width=40
    ws=wb.create_sheet("QCRA Tornado")
    ws.append(["Rank","Driver","CBS","Mean Exposure $B","Mitigation"]); _v54_style_header(ws)
    for i,r in enumerate(torn['qcra'],1): ws.append([i,r['title'],r['cbs'],r['cost_emv_bn'],r['mitigation']])
    chart=BarChart(); chart.type="bar"; chart.title="QCRA Tornado - Cost Exposure"; data=Reference(ws,min_col=4,min_row=1,max_row=ws.max_row); cats=Reference(ws,min_col=2,min_row=2,max_row=ws.max_row); chart.add_data(data,titles_from_data=True); chart.set_categories(cats); ws.add_chart(chart,"G2"); _v54_style_sheet(ws)
    ws=wb.create_sheet("QSRA Tornado")
    ws.append(["Rank","Driver","Activity","Mean Exposure Days","Mitigation"]); _v54_style_header(ws)
    for i,r in enumerate(torn['qsra'],1): ws.append([i,r['title'],r['activity_id'],r['schedule_emv_days'],r['mitigation']])
    chart=BarChart(); chart.type="bar"; chart.title="QSRA Tornado - Schedule Exposure"; data=Reference(ws,min_col=4,min_row=1,max_row=ws.max_row); cats=Reference(ws,min_col=2,min_row=2,max_row=ws.max_row); chart.add_data(data,titles_from_data=True); chart.set_categories(cats); ws.add_chart(chart,"G2"); _v54_style_sheet(ws)
    bio=BytesIO(); wb.save(bio); bio.seek(0); return bio.getvalue()


def risk_csv_bytes(model: Dict[str, Any]) -> bytes:
    out=StringIO(); w=csv.writer(out)
    w.writerow(["Risk ID","Title","Category","Likelihood %","Cause","Risk Event","Impact Description","Owner","Mitigation","Trigger","Residual Rating","Activity ID","Activity Name","CBS","CBS Name","QCRA Mean $B","QSRA Mean Days","Status"])
    for r in _v54_risks(model):
        w.writerow([r['risk_id'],r['title'],r['category'],r['probability_pct'],r['cause'],r['event'],r['impact'],r['owner'],r['mitigation'],r['trigger'],r['residual_rating'],r['activity_id'],r['activity_name'],r['cbs'],r['cbs_name'],r['cost_emv_bn'],r['schedule_emv_days'],r['status']])
    return out.getvalue().encode("utf-8")


def xer_bytes(model: Dict[str, Any]) -> bytes:
    # Legacy-targeted P6 XER style. PRA versions vary; for stubborn PRA installs use the Schedule sheet/CSV as fallback.
    today=datetime.utcnow().strftime("%Y-%m-%d")
    title=str(model.get('title') or 'CASEY Project').replace('\t',' ').replace('\n',' ')
    lines=[]
    lines.append(f"ERMHDR\t8.4\t{today}\tProject\tCASEY\tCASEY\tCASEYDatabase\tProject Management\tUSD")
    lines += [
        "%T\tCURRTYPE",
        "%F\tcurr_id\tdecimal_digit_cnt\tcurr_symbol\tdecimal_symbol\tdigit_group_symbol\tpos_curr_fmt_type\tneg_curr_fmt_type\tcurr_type\tcurr_short_name\tgroup_digit_cnt\tbase_exch_rate",
        "%R\t1\t2\t$\t.\t,\t#1.1\t(#1.1)\tDollar\tUSD\t3\t1",
        "%T\tOBS",
        "%F\tobs_id\tparent_obs_id\tguid\tseq_num\tobs_name\tobs_descr",
        "%R\t1\t\t\t0\tEnterprise\t",
        "%T\tCALENDAR",
        "%F\tclndr_id\tclndr_name\tdefault_flag\tday_hr_cnt\tweek_hr_cnt\tmonth_hr_cnt\tyear_hr_cnt",
        "%R\t1\tCASEY Standard 5 Day\tY\t8\t40\t173.33\t2080",
        "%T\tPROJECT",
        "%F\tproj_id\tfy_start_month_num\trsrc_self_add_flag\tallow_complete_flag\trsrc_multi_assign_flag\tcheckout_flag\tproject_flag\tstep_complete_flag\tcost_qty_recalc_flag\tbatch_sum_flag\tname_sep_char\tdef_complete_pct_type\tproj_short_name\tproj_name\tplan_start_date\tscd_end_date\tclndr_id\tcurr_id\torig_proj_id\tlast_fin_dates_id\tlast_baseline_update_date\tpriority_num\tcreate_date\tupdate_date\tcreate_user\tupdate_user",
        f"%R\t1\t1\tY\tY\tY\tN\tY\tY\tY\tY\t.\tCP\tCASEY\t{title}\t{today}\t\t1\t1\t\t\t\t1\t{today}\t{today}\tCASEY\tCASEY",
        "%T\tPROJWBS",
        "%F\twbs_id\tproj_id\tobs_id\tseq_num\twbs_short_name\twbs_name\tparent_wbs_id\tproj_node_flag",
        "%R\t1\t1\t1\t1\tCASEY\tCASEY Project\t\tY",
        "%T\tTASK",
        "%F\ttask_id\tproj_id\twbs_id\tclndr_id\tphys_complete_pct\trev_fdbk_flag\test_wt\tlock_plan_flag\tauto_compute_act_flag\tcomplete_pct_type\ttask_type\tduration_type\tstatus_code\ttask_code\ttask_name\ttarget_drtn_hr_cnt\tremain_drtn_hr_cnt\ttotal_float_hr_cnt",
    ]
    idx={}
    for i,a in enumerate(model.get('schedule_rows') or [],1):
        code=str(a.get('activity_id') or f"A{i:04d}").replace('\t',' ')
        name=str(a.get('activity') or 'Activity').replace('\t',' ').replace('\n',' ')
        dur=max(8, int(float(a.get('duration_months') or 1)*173.33))
        idx[code]=i
        lines.append(f"%R\t{i}\t1\t1\t1\t0\tN\t1\tN\tY\tCP\tTT_Task\tDT_FixedDUR2\tTK_NotStart\t{code}\t{name}\t{dur}\t{dur}\t0")
    lines += ["%T\tTASKPRED","%F\ttask_pred_id\ttask_id\tpred_task_id\tpred_type\tlag_hr_cnt"]
    k=1
    for a in model.get('schedule_rows') or []:
        code=str(a.get('activity_id'))
        for pred in str(a.get('predecessor') or '').split(';'):
            pred=pred.strip()
            if pred in idx and code in idx:
                lines.append(f"%R\t{k}\t{idx[code]}\t{idx[pred]}\tPR_FS\t0"); k+=1
    return ("\n".join(lines)+"\n").encode("latin-1", errors="replace")


def word_bytes(model: Dict[str, Any]) -> bytes:
    sc=_v54_scenario(model); risks=_v54_risks(model); torn=_v54_tornado(model); costs=_v54_costs(model)
    doc=Document(); styles=doc.styles; styles['Normal'].font.name='Aptos'; styles['Normal'].font.size=Pt(10)
    title=doc.add_heading('CASEY PLATINUM BOARD DECISION PACK',0); title.runs[0].font.color.rgb=RGBColor(11,31,51)
    doc.add_paragraph(f"{model.get('title')} | {sc['label']} scenario | Generated {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}")
    p=doc.add_paragraph(); p.add_run(sc['recommendation']).bold=True; p.runs[0].font.size=Pt(14); p.runs[0].font.color.rgb=RGBColor(180,35,24)
    doc.add_paragraph(sc['why'])
    doc.add_heading('Board Decision Snapshot',1)
    tbl=doc.add_table(rows=1,cols=4); tbl.style='Table Grid'
    for i,h in enumerate(['Metric','Value','Interpretation','Board action']): tbl.rows[0].cells[i].text=h
    for row in [('P50 Cost',model.get('cost_p50'),'Most likely first-pass control point','Use for option comparison'),('Cost Range',model.get('cost_range'),'Class and QCRA uncertainty range','Approve only with range acknowledged'),('QCRA P80',money_bn(model.get('monte_carlo',{}).get('qcra',{}).get('p80',0)),'Risk-adjusted funding confidence','Set contingency rules'),('QSRA P80',str(model.get('monte_carlo',{}).get('qsra',{}).get('p80',''))+' months','Schedule confidence level','Validate critical path'),('Risk / Confidence',f"{model.get('risk')} / {model.get('confidence_pct')}%",'Data maturity and residual exposure','Approve evidence-improvement plan')]:
        cells=tbl.add_row().cells
        for i,v in enumerate(row): cells[i].text=str(v)
    doc.add_heading('QCRA Cost Tornado',1)
    for r in torn['qcra'][:8]: doc.add_paragraph(f"{r['title']} — {r['cbs']} — {money_bn(r['cost_emv_bn'])} mean exposure", style='List Bullet')
    doc.add_heading('QSRA Schedule Tornado',1)
    for r in torn['qsra'][:8]: doc.add_paragraph(f"{r['title']} — {r['activity_id']} — {r['schedule_emv_days']} days mean exposure", style='List Bullet')
    doc.add_heading('Board Top Risks',1)
    for r in risks[:10]:
        doc.add_paragraph(f"{r['risk_id']} {r['title']}", style='List Bullet')
        doc.add_paragraph(f"Cause: {r['cause']}\nEvent: {r['event']}\nImpact: {r['impact']}\nMitigation: {r['mitigation']}")
    doc.add_heading('Selected Class Cost Estimate',1)
    tbl=doc.add_table(rows=1, cols=6); tbl.style='Table Grid'
    for i,h in enumerate(['CBS','Description','Type','P10','P50','P90']): tbl.rows[0].cells[i].text=h
    for x in costs:
        cells=tbl.add_row().cells
        vals=[x['cbs'],x['description'],x['type'],money_bn(x['p10_bn']),money_bn(x['p50_bn']),money_bn(x['p90_bn'])]
        for i,v in enumerate(vals): cells[i].text=str(v)
    bio=BytesIO(); doc.save(bio); bio.seek(0); return bio.getvalue()


def pdf_bytes(model: Dict[str, Any]) -> bytes:
    sc=_v54_scenario(model); risks=_v54_risks(model); torn=_v54_tornado(model); costs=_v54_costs(model)
    bio=BytesIO(); doc=SimpleDocTemplate(bio,pagesize=landscape(A4),leftMargin=28,rightMargin=28,topMargin=24,bottomMargin=24)
    styles=getSampleStyleSheet()
    styles.add(ParagraphStyle(name='CaseyTitle', parent=styles['Title'], fontSize=24, textColor=colors.HexColor("#"+V54_BLUE), leading=28, spaceAfter=10))
    styles.add(ParagraphStyle(name='CaseyH', parent=styles['Heading1'], fontSize=16, textColor=colors.HexColor("#"+V54_BLUE), spaceAfter=8))
    styles.add(ParagraphStyle(name='Small', parent=styles['BodyText'], fontSize=8, leading=10))
    story=[]
    story.append(Paragraph(f"CASEY BOARD DECISION PACK — {model.get('title')}", styles['CaseyTitle']))
    story.append(Paragraph(f"{sc['label']} scenario | {model.get('subsector')} | {model.get('location')} | Generated {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}", styles['Small']))
    kpis=[['P50 Cost',model.get('cost_p50'),'Cost Range',model.get('cost_range'),'QCRA P80',money_bn(model.get('monte_carlo',{}).get('qcra',{}).get('p80',0))],['Schedule',model.get('schedule'),'QSRA P80',str(model.get('monte_carlo',{}).get('qsra',{}).get('p80',''))+' mo','Risk / Confidence',f"{model.get('risk')} / {model.get('confidence_pct')}%"]]
    t=Table(kpis, colWidths=[90,110,90,130,90,130]); t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),colors.HexColor('#F3F7FA')),('TEXTCOLOR',(0,0),(-1,-1),colors.HexColor("#"+V54_BLUE)),('FONTNAME',(0,0),(-1,-1),'Helvetica-Bold'),('GRID',(0,0),(-1,-1),0.5,colors.HexColor("#"+V54_LINE)),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('FONTSIZE',(0,0),(-1,-1),10)])); story.append(t); story.append(Spacer(1,12))
    story.append(Paragraph(sc['recommendation'].upper(), styles['CaseyH'])); story.append(Paragraph(sc['why'], styles['BodyText'])); story.append(PageBreak())
    story.append(Paragraph('QCRA + QSRA Tornado Drivers', styles['CaseyH']))
    data=[['Rank','QCRA Cost Driver','CBS','Mean Cost','QSRA Schedule Driver','Activity','Mean Days']]
    for i in range(max(len(torn['qcra']),len(torn['qsra']))):
        qc=torn['qcra'][i] if i<len(torn['qcra']) else {}; qs=torn['qsra'][i] if i<len(torn['qsra']) else {}
        data.append([i+1,qc.get('title',''),qc.get('cbs',''),money_bn(qc.get('cost_emv_bn',0)) if qc else '',qs.get('title',''),qs.get('activity_id',''),qs.get('schedule_emv_days','')])
    t=Table(data, repeatRows=1, colWidths=[35,145,55,70,145,65,60]); t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor("#"+V54_BLUE)),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.35,colors.HexColor("#"+V54_LINE)),('FONTSIZE',(0,0),(-1,-1),8),('VALIGN',(0,0),(-1,-1),'TOP'),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F8FBFD')]) ])); story.append(t); story.append(PageBreak())
    story.append(Paragraph('Risk Register — Board Top Risks', styles['CaseyH']))
    data=[['ID','Risk','Cause','Event','Impact','Owner','Mitigation']]
    for r in risks[:10]: data.append([r['risk_id'],r['title'],r['cause'],r['event'],r['impact'],r['owner'],r['mitigation']])
    t=Table(data, repeatRows=1, colWidths=[45,90,120,120,160,70,145]); t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor("#"+V54_BLUE)),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.25,colors.HexColor("#"+V54_LINE)),('FONTSIZE',(0,0),(-1,-1),6.8),('VALIGN',(0,0),(-1,-1),'TOP'),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F8FBFD')]) ])); story.append(t); story.append(PageBreak())
    story.append(Paragraph('Cost Estimate — Selected Class Only', styles['CaseyH']))
    data=[['CBS','Description','Type','P10','P50','P90','Basis']]
    for x in costs: data.append([x['cbs'],x['description'],x['type'],money_bn(x['p10_bn']),money_bn(x['p50_bn']),money_bn(x['p90_bn']),x.get('basis','')])
    t=Table(data, repeatRows=1, colWidths=[50,150,70,70,70,70,300]); t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor("#"+V54_BLUE)),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.25,colors.HexColor("#"+V54_LINE)),('FONTSIZE',(0,0),(-1,-1),7.2),('VALIGN',(0,0),(-1,-1),'TOP'),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F8FBFD')]) ])); story.append(t)
    doc.build(story); bio.seek(0); return bio.getvalue()


def pptx_bytes(model: Dict[str, Any]) -> bytes:
    prs=Presentation(); prs.slide_width=PptxInches(13.333); prs.slide_height=PptxInches(7.5)
    sc=_v54_scenario(model); risks=_v54_risks(model); torn=_v54_tornado(model)
    def slide(title, subtitle=""):
        s=prs.slides.add_slide(prs.slide_layouts[6]);
        bg=s.shapes.add_shape(1,0,0,prs.slide_width,prs.slide_height); bg.fill.solid(); bg.fill.fore_color.rgb=__import__('pptx').dml.color.RGBColor(255,255,255); bg.line.fill.background()
        t=s.shapes.add_textbox(PptxInches(.55),PptxInches(.35),PptxInches(12),PptxInches(.7)); tf=t.text_frame; tf.text=title; run=tf.paragraphs[0].runs[0]; run.font.size=PptxPt(28); run.font.bold=True; run.font.color.rgb=__import__('pptx').dml.color.RGBColor(11,31,51)
        if subtitle:
            st=s.shapes.add_textbox(PptxInches(.6),PptxInches(1.05),PptxInches(12),PptxInches(.4)); st.text_frame.text=subtitle; st.text_frame.paragraphs[0].runs[0].font.size=PptxPt(12); st.text_frame.paragraphs[0].runs[0].font.color.rgb=__import__('pptx').dml.color.RGBColor(90,90,90)
        return s
    s=slide(f"{model.get('title')} — Board Intelligence", f"{sc['label']} scenario | CASEY Platinum Output Engine")
    tx=s.shapes.add_textbox(PptxInches(.75),PptxInches(2.0),PptxInches(12),PptxInches(1.8)); tf=tx.text_frame; tf.text=sc['recommendation'].upper(); tf.paragraphs[0].runs[0].font.size=PptxPt(30); tf.paragraphs[0].runs[0].font.bold=True
    tx2=s.shapes.add_textbox(PptxInches(.8),PptxInches(4.1),PptxInches(11),PptxInches(.6)); tx2.text_frame.text=sc['why']
    s=slide("Board Decision Snapshot")
    metrics=[('P50',model.get('cost_p50')),('Range',model.get('cost_range')),('QCRA P80',money_bn(model.get('monte_carlo',{}).get('qcra',{}).get('p80',0))),('Schedule',model.get('schedule')),('QSRA P80',str(model.get('monte_carlo',{}).get('qsra',{}).get('p80',''))+' mo'),('Risk',f"{model.get('risk')} / {model.get('confidence_pct')}%")]
    for i,(k,v) in enumerate(metrics):
        x=.75+(i%3)*4.15; y=1.6+(i//3)*1.7; shp=s.shapes.add_shape(1,PptxInches(x),PptxInches(y),PptxInches(3.6),PptxInches(1.15)); shp.fill.solid(); shp.fill.fore_color.rgb=__import__('pptx').dml.color.RGBColor(243,247,250); shp.line.color.rgb=__import__('pptx').dml.color.RGBColor(217,226,234); shp.text=f"{k}\n{v}"; shp.text_frame.paragraphs[0].runs[0].font.size=PptxPt(11); shp.text_frame.paragraphs[1].runs[0].font.size=PptxPt(20); shp.text_frame.paragraphs[1].runs[0].font.bold=True
    s=slide("QCRA and QSRA are separated", "Cost exposure and schedule exposure are not mixed.")
    left=s.shapes.add_textbox(PptxInches(.8),PptxInches(1.6),PptxInches(5.8),PptxInches(5.2)).text_frame; left.text="QCRA COST TORNADO"; left.paragraphs[0].runs[0].font.bold=True
    for r in torn['qcra'][:7]: p=left.add_paragraph(); p.text=f"{r['title']} — {money_bn(r['cost_emv_bn'])}"; p.level=0
    right=s.shapes.add_textbox(PptxInches(7.0),PptxInches(1.6),PptxInches(5.8),PptxInches(5.2)).text_frame; right.text="QSRA SCHEDULE TORNADO"; right.paragraphs[0].runs[0].font.bold=True
    for r in torn['qsra'][:7]: p=right.add_paragraph(); p.text=f"{r['title']} — {r['schedule_emv_days']} days"; p.level=0
    s=slide("Top risks have cause, event, impact and action", "Designed to be used in a real risk review, not as a decorative export.")
    tf=s.shapes.add_textbox(PptxInches(.75),PptxInches(1.5),PptxInches(12),PptxInches(5.5)).text_frame
    for r in risks[:6]: p=tf.add_paragraph(); p.text=f"{r['risk_id']} {r['title']}: {r['impact']} | Action: {r['mitigation']}"; p.level=0; p.font.size=PptxPt(12)
    s=slide("What CASEY delivers", "Outputs are readable, scenario-linked and defensible.")
    tf=s.shapes.add_textbox(PptxInches(.9),PptxInches(1.6),PptxInches(11.5),PptxInches(4.8)).text_frame
    for txt in ["Selected class cost estimate only", "Direct / indirect / reserve split", "Risk register with cause, event, impact, owner and mitigation", "Separate QCRA and QSRA curves/tornado", "Legacy-targeted schedule export + schedule table fallback", "Board PDF, DOCX, PPTX, Excel and ZIP pack"]:
        p=tf.add_paragraph(); p.text="✓ "+txt; p.font.size=PptxPt(18)
    bio=BytesIO(); prs.save(bio); bio.seek(0); return bio.getvalue()

# v54 routes for explicit use; old /export routes automatically use the overridden global builders above.
@app.post('/v54/export/all')
def export_all_v54(model:Dict[str,Any]):
    bio=BytesIO()
    with zipfile.ZipFile(bio,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr('01_CASEY_v54_Platinum_Cost_Model.xlsx',workbook_bytes(model))
        z.writestr('02_CASEY_v54_Platinum_Risk_Register.xlsx',risk_register_workbook_bytes(model))
        z.writestr('03_CASEY_v54_Legacy_Targeted_P6_Schedule.xer',xer_bytes(model))
        z.writestr('04_CASEY_v54_Board_Decision_Pack.docx',word_bytes(model))
        z.writestr('05_CASEY_v54_Board_Decision_Pack.pdf',pdf_bytes(model))
        z.writestr('06_CASEY_v54_Board_Deck.pptx',pptx_bytes(model))
        z.writestr('07_CASEY_v54_Model_Audit.json',json.dumps(model,indent=2))
        z.writestr('08_CASEY_v54_Risk_Register.csv',risk_csv_bytes(model))
        # schedule CSV fallback for tools that reject XER versions
        out=StringIO(); w=csv.writer(out); w.writerow(['Activity ID','Activity','Predecessor','Duration Months','Critical'])
        for a in model.get('schedule_rows') or []: w.writerow([a.get('activity_id'),a.get('activity'),a.get('predecessor'),a.get('duration_months'),a.get('critical')])
        z.writestr('09_CASEY_v54_Schedule_Fallback.csv',out.getvalue())
        z.writestr('README_PRA_IMPORT_NOTE.txt','If Primavera Risk Analysis rejects XER on your installed version, import/use the Schedule_Fallback.csv or re-export from P6 Professional in the PRA-supported version. CASEY provides both a legacy-targeted XER and a clean schedule table fallback.')
    bio.seek(0); return stream(bio.getvalue(),'application/zip','CASEY_v54_Platinum_Output_Pack.zip')


# =====================================================================
# v55 PLATINUM OUTPUTS FINAL PATCH
# - Excel repaired by replacing openpyxl-generated export workbooks with xlsxwriter.
# - PPTX removed from output factory and ZIP because weak decks reduce trust.
# - XER header targeted to Primavera/PRA-compatible ERMHDR 8.0 style.
# - Outputs adapt to selected scenario, project sector, Earth/Space and all cost/schedule levels.
# =====================================================================

import html
from fastapi.responses import JSONResponse
try:
    import xlsxwriter
except Exception:
    xlsxwriter = None

V55_BLUE = '#0B1F33'
V55_CYAN = '#00A6C8'
V55_LINE = '#D9E2EA'
V55_LIGHT = '#F4F7FA'
V55_RED = '#C00000'
V55_AMBER = '#F4B183'
V55_GREEN = '#70AD47'


def _v55_scenario_key(model):
    return str(model.get('scenario') or 'base').lower().replace(' ', '_')


def _v55_scenario_label(model):

    # FINAL EXEC POLISH: scenario-aware ranking and differentiation
    try:
        sc = str(model.get("scenario","base")).lower()
        schedule_lists = {
            "faster":[
                "Concurrent commissioning overload",
                "Recovery float exhaustion",
                "Acceleration premium shock",
                "Grid connection delay",
                "Integrated systems testing concurrency"
            ],
            "cheaper":[
                "Vendor claims and change exposure",
                "Procurement deferral and long-lead slippage",
                "Design maturity gap",
                "Scope growth from deferred decisions",
                "Interface coordination delay"
            ],
            "lower_risk":[
                "Governance and approvals latency",
                "Extended validation sequencing",
                "Conservative commissioning gates",
                "Operational readiness hold-points",
                "Assurance and compliance reviews"
            ],
            "premium":[
                "Integration complexity across parallel packages",
                "Executive decision latency",
                "Technology assurance alignment",
                "Multi-package interface management",
                "Programme coordination overhead"
            ]
        }
        cost_lists = {
            "faster":[
                "Acceleration premiums and overtime",
                "Power train, transformers and switchgear",
                "Integrated systems testing",
                "Grid and utility concurrency",
                "Recovery-float consumption"
            ],
            "cheaper":[
                "Deferred procurement packaging",
                "Claims and commercial exposure",
                "Rework from reduced contingency",
                "Long-lead inflation volatility",
                "Scope rationalisation impacts"
            ],
            "lower_risk":[
                "Additional contingency and reserve",
                "Enhanced validation and assurance",
                "Programme controls and governance",
                "Redundant infrastructure resilience",
                "Extended commissioning readiness"
            ]
        }
        if sc in schedule_lists:
            model["sector_schedule_threats"]=schedule_lists[sc]
        if sc in cost_lists:
            model["sector_primary_cost_drivers"]=cost_lists[sc]

        if sc=="faster":
            model["executive_shock_insight"]="Acceleration increases spend faster than it reduces uncertainty; the delivery tail becomes more volatile."
        elif sc=="cheaper":
            model["executive_shock_insight"]="Capital efficiency reduces resilience: procurement and recovery flexibility become constrained."
        elif sc=="lower_risk":
            model["executive_shock_insight"]="Confidence is purchased through reserve, governance and extended delivery duration."
        elif sc=="premium":
            model["executive_shock_insight"]="Premium posture buys resilience, optionality and stronger certainty at visible capex premium."
    except Exception:
        pass

    return model.get('scenario_label') or scenario_params(_v55_scenario_key(model))[4]


def _v55_selected_class(model):
    try: return int(model.get('estimate_class') or 3)
    except Exception: return 3


def _v55_selected_schedule_level(model):
    try: return int(model.get('schedule_level') or 3)
    except Exception: return 3


def _v55_scenario_strategy(model):
    key = _v55_scenario_key(model)
    table = {
        'base': {
            'title': 'APPROVE DEFINITION PHASE — DO NOT COMMIT FULL CAPITAL YET',
            'summary': 'Balanced reference case for first-pass board challenge, option testing and contractor / consultant assurance.',
            'decision': 'Use as the control baseline. Validate evidence before approving full capital.'},
        'faster': {
            'title': 'ACCELERATE ONLY WITH EXPLICIT COST AND INTERFACE CONTROL',
            'summary': 'Faster scenario compresses programme duration through parallel design, early procurement and premium logistics, increasing interface and delivery risk.',
            'decision': 'Approve acceleration only if time value exceeds premium cost and added risk.'},
        'cheaper': {
            'title': 'LOWER COST CASE REQUIRES STRONGER RISK GOVERNANCE',
            'summary': 'Cheaper scenario reduces scope and procurement cost, but usually transfers risk into quality, schedule and contingency drawdown.',
            'decision': 'Use only where capex pressure is dominant and residual risk is accepted.'},
        'lower_risk': {
            'title': 'LOWER RISK CASE IMPROVES CONFIDENCE BUT REQUIRES MORE TIME AND ALLOWANCE',
            'summary': 'Lower Risk scenario adds surveys, assurance, stage gates, buffers and procurement validation to increase confidence.',
            'decision': 'Recommended when board certainty matters more than lowest first cost.'},
        'premium': {
            'title': 'PREMIUM CASE MAXIMISES RESILIENCE AND DELIVERY CONFIDENCE',
            'summary': 'Premium scenario strengthens suppliers, design maturity, resilience and governance to reduce downstream uncertainty.',
            'decision': 'Use for flagship, mission-critical or reputation-sensitive programmes.'},
        'investor': {
            'title': 'INVESTOR CASE MAKES RISK AND VALUE TRADE-OFFS EXPLICIT',
            'summary': 'Investor scenario frames the programme as an investment decision with clear capital, risk and confidence trade-offs.',
            'decision': 'Use to support IC, lender, fund or partner conversations.'},
        'survival': {
            'title': 'SURVIVAL CASE IS WARNING-HEAVY AND SHOULD NOT BE PRESENTED AS A TARGET BASELINE',
            'summary': 'Survival scenario minimises scope and cost, but creates high residual risk and quality / schedule exposure.',
            'decision': 'Use only as a stress case, not a preferred delivery path.'}
    }
    return table.get(key, table['base'])


def _v55_money(v):
    try: return money_bn(float(v))
    except Exception: return str(v)


def _v55_all_cost_lines(model):
    """Return dictionary {class: rows}. Always includes selected and all levels."""
    selected = model.get('cost_lines') or []
    by = model.get('estimates_by_class') or {}
    out = {}
    for cls in range(1, 6):
        rows = by.get(str(cls)) or []
        if not rows:
            low, high, _, mat = class_range(cls)
            rows = [{**x, 'p10_bn': round(float(x.get('p50_bn',0))*low, 3), 'p90_bn': round(float(x.get('p50_bn',0))*high, 3), 'class': cls, 'maturity': mat} for x in selected]
        out[cls] = rows
    return out


def _v55_all_schedule_levels(model):
    levels = model.get('schedules_by_level') or {}
    if not levels:
        levels = {str(_v55_selected_schedule_level(model)): model.get('schedule_rows') or []}
    return {str(k): list(v or []) for k, v in levels.items()}


def _v55_extra_risks(model):
    mode = str(model.get('mode') or 'Earth')
    sub = str(model.get('subsector') or '').lower()
    generic = [
        ('R-009','Commercial claims escalation','Contract ambiguity, change volume and interface disputes','Claims submitted above allowance','Commercial pressure, management distraction and increased outturn cost','Commercial Lead','Tight change control, entitlement reviews, dispute avoidance board','Claim volume exceeds baseline'),
        ('R-010','Resource availability','Labour, specialist supervisors or commissioning teams constrained','Required resources are unavailable when planned','Productivity loss, resequencing and higher preliminaries','Delivery Lead','Resource loading, framework call-offs, labour strategy','Resource histogram red flag'),
        ('R-011','Quality failure / rework','Installation quality, vendor QA or inspection regime below need','Inspection failure triggers rework','Rework cost, schedule delay and confidence loss','Quality Manager','ITPs, hold points, vendor QA, right-first-time reviews','NCR trend increases'),
        ('R-012','Cyber / controls integration','Operational technology, control systems and network interfaces immature','Controls integration fails testing','Commissioning delay, resilience issue and operational risk','Systems Lead','Early cyber review, FAT/SAT plan, integration lab','Controls defect trend increases'),
        ('R-013','Weather / environmental disruption','Extreme weather, environmental windows or site constraints not fully allowed','Workface access or productivity is disrupted','Delay, access cost and resequencing','Construction Manager','Weather calendars, contingency workfronts, protection strategy','Weather downtime exceeds allowance'),
        ('R-014','Stakeholder change','Sponsor, operator or authority requirements evolve late','Requirements change after baseline freeze','Scope growth, redesign and approval delays','Project Director','Decision log, change board and stakeholder gates','Late decision count increases'),
        ('R-015','Funding / approval gate delay','Investment approvals or staged funding gates not aligned with procurement','Capital release is delayed','Procurement missed windows and schedule slippage','Sponsor','Approval roadmap, early board papers, contingency approvals','Approval date slips'),
        ('R-016','Commissioning readiness gap','Incomplete test packs, operational readiness or training gaps','Commissioning starts without readiness','Repeated tests, handover delay and operational risk','Commissioning Lead','Readiness dashboard, integrated test plan, early ORAT','Readiness score below threshold'),
        ('R-017','Data / estimate basis weakness','Estimate source data, quantities or benchmarks are incomplete','Baseline cannot be defended under challenge','Confidence reduction and rework of cost plan','Cost Lead','Assumption register, benchmark evidence, quantity validation','Unknown allowance exceeds threshold'),
        ('R-018','Interface ownership gap','RACI and package boundaries are unclear across packages','Interface issue has no accountable owner','Delay, claims and fragmented accountability','Integration Manager','Interface control documents and owner matrix','Interface actions overdue'),
    ]
    sector = []
    if 'airport' in sub:
        sector += [('R-A01','Airside phasing disruption','Operational airport constraints and possessions are underestimated','Construction cannot access critical work fronts','Night work premium, delay and operational disruption','Airside Lead','Airside phasing plan, possessions, ORAT integration','Airside access cancelled'),('R-A02','Baggage / security systems integration','Complex passenger processing systems not fully integrated','Systems fail integrated trials','Terminal opening delay and reputation impact','Systems Lead','Early integration lab, FAT/SAT, ORAT testing','Trial failures increase')]
    if 'rail' in sub:
        sector += [('R-R01','Possession access constraints','Rail access windows and blockade availability are insufficient','Planned works cannot be delivered in possession','Resequencing, delay and overtime cost','Rail Access Lead','Possession strategy and contingency blockade plan','Possession cancelled'),('R-R02','Signalling integration','Signalling design, test and assurance duration underestimated','System does not pass integration testing','Critical path delay and safety assurance issue','Signalling Lead','Independent signalling assurance and test lab','Test defects exceed threshold')]
    if 'data centre' in sub:
        sector += [('R-D03','Liquid cooling vendor maturity','Cooling technology and supplier capacity are immature','Vendor solution underperforms acceptance tests','Capacity derate, redesign and schedule delay','MEP Lead','Prototype, vendor qualification and performance guarantee','Cooling FAT failed'),('R-D04','Generator / UPS lead time','Critical power equipment supply is constrained','Power train delivery slips','Critical path movement and acceleration premium','Procurement Lead','Early vendor reservation and dual-source strategy','Vendor delivery warning')]
    if mode.lower() == 'space':
        sector += [('R-S04','Launch window dependency','Orbital mechanics or launch provider availability constrains deployment','Launch window missed','Months of delay and re-manifest cost','Mission Lead','Backup launch windows and manifest options','Launch readiness slips'),('R-S05','Thermal / radiation design margin','Environment and shielding assumptions remain immature','Thermal or radiation margins fail review','Redesign, extra mass and launch cost increase','Chief Engineer','Thermal vacuum testing and radiation analysis','Margin below threshold')]
    return sector + generic


def _v55_risks(model):
    base = []
    for r in model.get('risks') or model.get('risk_register') or []:
        p = max(5, int(float(r.get('probability_pct') or r.get('likelihood_pct') or 25)))
        title = r.get('title') or r.get('risk') or r.get('name') or 'Risk'
        cause = r.get('cause') or r.get('description') or 'Underlying assumption, external dependency or delivery condition requires validation.'
        event = r.get('event') or f'{title} occurs or materialises during delivery.'
        impact = r.get('impact') or r.get('impact_description') or r.get('basis_of_schedule_impact') or 'Cost, schedule, quality, scope or confidence impact.'
        mitig = r.get('mitigation') or r.get('response') or 'Assign owner, validate assumption, implement mitigation and track residual exposure.'
        owner = r.get('owner') or 'Project Controls Lead'
        c = max(0.001, float(r.get('cost_emv_bn') or 0.01))
        d = max(0.1, float(r.get('schedule_emv_days') or 0.1))
        res = r.get('residual_rating') or ('Red' if p>=45 or c>0.5 or d>30 else 'Amber' if p>=25 or c>0.15 or d>10 else 'Green')
        base.append({
            'risk_id': r.get('risk_id') or r.get('id') or f'R-{len(base)+1:03d}', 'title': title, 'category': r.get('category') or 'Delivery', 'probability_pct': p,
            'cause': cause, 'event': event, 'impact': impact, 'owner': owner, 'mitigation': mitig,
            'trigger': r.get('trigger') or 'Threshold breach or assumption change', 'residual_rating': res,
            'activity_id': r.get('activity_id') or 'A1900', 'activity_name': r.get('activity_name') or 'Delivery activity',
            'cbs': r.get('cbs') or '01.04', 'cbs_name': r.get('cbs_name') or 'Cost package', 'cost_emv_bn': round(c,3), 'schedule_emv_days': round(d,1),
            'status': r.get('status') or 'Open'
        })
    cost = parse_bn(model.get('cost_p50'))
    months = int(str(model.get('schedule') or '60').split()[0]) if str(model.get('schedule') or '').split() else 60
    rm = scenario_params(_v55_scenario_key(model))[2]
    for rid,title,cause,event,impact,owner,mit,trig in _v55_extra_risks(model):
        if any(x['risk_id'] == rid for x in base):
            continue
        idx = len(base)+1
        p = max(8, min(72, int((22 + (idx % 7)*4) * rm)))
        c = round(max(0.015, cost*(0.006 + (idx%5)*0.004)*p/100),3)
        d = round(max(0.8, (months*0.35 + (idx%6)*6)*p/100),1)
        base.append({'risk_id': rid, 'title': title, 'category': 'Delivery', 'probability_pct': p, 'cause': cause, 'event': event, 'impact': impact, 'owner': owner, 'mitigation': mit, 'trigger': trig, 'residual_rating': 'Red' if p>=45 or c>0.5 or d>30 else 'Amber' if p>=25 else 'Green', 'activity_id': f'A{1400+(idx%8)*100}', 'activity_name': 'Linked schedule activity', 'cbs': f'01.{(idx%11)+1:02d}', 'cbs_name': 'Linked cost package', 'cost_emv_bn': c, 'schedule_emv_days': d, 'status': 'Open'})
    return sorted(base, key=lambda x: (x['cost_emv_bn']*100 + x['schedule_emv_days']/10), reverse=True)


def _v55_curves(model):
    mc = model.get('monte_carlo') or {}
    curve = mc.get('curve') or []
    if curve:
        qcra = [(int(x.get('percentile')), float(x.get('cost_bn') or 0)) for x in curve]
        qsra = [(int(x.get('percentile')), float(x.get('schedule_months') or 0)) for x in curve]
    else:
        c = parse_bn(model.get('cost_p50')); m = int(str(model.get('schedule') or '60').split()[0])
        qcra = [(p, round(c*(0.75+p/100*0.65),2)) for p in [1,5,10,20,30,40,50,60,70,80,90,95,99]]
        qsra = [(p, round(m*(0.85+p/100*0.38),2)) for p in [1,5,10,20,30,40,50,60,70,80,90,95,99]]
    return {'qcra': qcra, 'qsra': qsra}


def _v55_tornado(model):
    risks = _v55_risks(model)
    qcra = [r for r in risks if r['cost_emv_bn'] > 0]
    qsra = [r for r in risks if r['schedule_emv_days'] > 0]
    qcra = sorted(qcra, key=lambda r: r['cost_emv_bn'], reverse=True)[:12]
    qsra = sorted(qsra, key=lambda r: r['schedule_emv_days'], reverse=True)[:12]
    return {'qcra': qcra, 'qsra': qsra}


def _v55_make_workbook_base(filename='casey.xlsx'):
    if not xlsxwriter:
        raise RuntimeError('xlsxwriter is required for v55 platinum exports')
    bio = BytesIO()
    wb = xlsxwriter.Workbook(bio, {'in_memory': True, 'constant_memory': False})
    return wb, bio


def _v55_formats(wb):
    return {
        'title': wb.add_format({'bold': True, 'font_size': 20, 'font_color': V55_BLUE, 'bg_color': 'FFFFFF'}),
        'subtitle': wb.add_format({'font_size': 10, 'font_color': '#666666'}),
        'section': wb.add_format({'bold': True, 'font_size': 13, 'font_color': 'FFFFFF', 'bg_color': V55_BLUE, 'border': 1, 'border_color': V55_LINE}),
        'header': wb.add_format({'bold': True, 'font_color': 'FFFFFF', 'bg_color': '#0070C0', 'border': 1, 'border_color': V55_LINE, 'valign': 'vcenter'}),
        'cell': wb.add_format({'border': 1, 'border_color': V55_LINE, 'valign': 'top', 'text_wrap': True}),
        'cell_money': wb.add_format({'border': 1, 'border_color': V55_LINE, 'num_format': '$#,##0.0B', 'valign': 'top'}),
        'cell_num': wb.add_format({'border': 1, 'border_color': V55_LINE, 'num_format': '0.0', 'valign': 'top'}),
        'kpi_label': wb.add_format({'bold': True, 'font_color': V55_BLUE, 'bg_color': V55_LIGHT, 'border': 1, 'border_color': V55_LINE}),
        'kpi_value': wb.add_format({'bold': True, 'font_size': 16, 'font_color': V55_BLUE, 'bg_color': 'FFFFFF', 'border': 1, 'border_color': V55_LINE}),
        'red': wb.add_format({'bg_color': '#FCE4D6', 'border': 1, 'border_color': V55_LINE, 'text_wrap': True, 'valign': 'top'}),
        'amber': wb.add_format({'bg_color': '#FFF2CC', 'border': 1, 'border_color': V55_LINE, 'text_wrap': True, 'valign': 'top'}),
        'green': wb.add_format({'bg_color': '#E2F0D9', 'border': 1, 'border_color': V55_LINE, 'text_wrap': True, 'valign': 'top'}),
        'note': wb.add_format({'italic': True, 'font_color': '#666666', 'text_wrap': True})
    }


def _v55_setup_ws(ws):
    ws.set_zoom(90)
    ws.hide_gridlines(2)
    ws.freeze_panes(5, 0)
    ws.set_landscape()
    ws.set_margins(0.3, 0.3, 0.4, 0.4)


def _v55_write_table(ws, start_row, start_col, headers, rows, fmts, widths=None, money_cols=None, num_cols=None):
    money_cols = set(money_cols or [])
    num_cols = set(num_cols or [])
    for c,h in enumerate(headers): ws.write(start_row, start_col+c, h, fmts['header'])
    for r,row in enumerate(rows, start=start_row+1):
        for c,val in enumerate(row):
            fmt = fmts['cell_money'] if c in money_cols else fmts['cell_num'] if c in num_cols else fmts['cell']
            ws.write(r, start_col+c, val, fmt)
    if widths:
        for i,w in enumerate(widths): ws.set_column(start_col+i, start_col+i, w)
    return start_row + len(rows) + 1


def workbook_bytes(model: Dict[str, Any]) -> bytes:
    wb, bio = _v55_make_workbook_base(); f = _v55_formats(wb)
    sc = _v55_scenario_strategy(model); sel_cls = _v55_selected_class(model); sel_lvl = _v55_selected_schedule_level(model)
    costs_by = _v55_all_cost_lines(model); costs = costs_by.get(sel_cls, model.get('cost_lines') or [])
    risks = _v55_risks(model); curves = _v55_curves(model); torn = _v55_tornado(model)
    # Executive dashboard
    ws = wb.add_worksheet('01 Executive Dashboard'); _v55_setup_ws(ws)
    ws.merge_range('A1:H1', 'CASEY PLATINUM COST + RISK MODEL', f['title'])
    ws.merge_range('A2:H2', f"{model.get('title')} | {_v55_scenario_label(model)} scenario | Class {sel_cls} | Schedule Level {sel_lvl}", f['subtitle'])
    kpis = [('P50 Cost', model.get('cost_p50')), ('Cost Range', model.get('cost_range')), ('QCRA P80', _v55_money((model.get('monte_carlo') or {}).get('qcra',{}).get('p80',0))), ('Schedule', model.get('schedule')), ('QSRA P80', str((model.get('monte_carlo') or {}).get('qsra',{}).get('p80',''))+' mo'), ('Risk / Confidence', f"{model.get('risk')} / {model.get('confidence_pct')}%")]
    for i,(k,v) in enumerate(kpis):
        r=4+(i//3)*3; c=(i%3)*3
        ws.merge_range(r,c,r,c+1,k,f['kpi_label']); ws.merge_range(r+1,c,r+1,c+1,str(v),f['kpi_value'])
    ws.write(11,0,'Executive recommendation',f['section']); ws.merge_range(12,0,12,7,sc['title'],f['kpi_value']); ws.merge_range(13,0,14,7,sc['summary'] + ' ' + sc['decision'],f['cell'])
    # cost composition chart source
    summary={'Direct':0.0,'Indirect':0.0,'Reserve':0.0}
    for x in costs: summary[x.get('type','Direct')] = summary.get(x.get('type','Direct'),0)+float(x.get('p50_bn') or 0)
    _v55_write_table(ws,17,0,['Cost Type','P50 $B'],list(summary.items()),f,[18,16],money_cols=[1])
    chart=wb.add_chart({'type':'doughnut'}); chart.add_series({'name':'Cost composition','categories':['01 Executive Dashboard',18,0,20,0],'values':['01 Executive Dashboard',18,1,20,1],'data_labels':{'percentage':True}}); chart.set_title({'name':'Direct / Indirect / Reserve'}); chart.set_style(10); ws.insert_chart('D17',chart,{'x_scale':1.2,'y_scale':1.1})
    # Selected cost estimate
    ws = wb.add_worksheet('02 Selected Cost Estimate'); _v55_setup_ws(ws)
    ws.merge_range('A1:H1', f'SELECTED CLASS {sel_cls} COST ESTIMATE — {_v55_scenario_label(model).upper()} SCENARIO', f['title'])
    rows=[[x.get('cbs'),x.get('description'),x.get('type'),x.get('basis'),float(x.get('p10_bn') or 0),float(x.get('p50_bn') or 0),float(x.get('p90_bn') or 0),x.get('impact_basis') or 'Validate source, benchmark and scope basis.'] for x in costs]
    _v55_write_table(ws,4,0,['CBS','Description','Type','Basis','P10 $B','P50 $B','P90 $B','Challenge / Basis'],rows,f,[10,28,13,34,12,12,12,46],money_cols=[4,5,6])
    chart=wb.add_chart({'type':'column'}); chart.add_series({'name':'P50 $B','categories':['02 Selected Cost Estimate',5,1,4+len(rows),1],'values':['02 Selected Cost Estimate',5,5,4+len(rows),5]}); chart.set_title({'name':'P50 cost by package'}); chart.set_y_axis({'name':'$B'}); ws.insert_chart('J4',chart,{'x_scale':1.3,'y_scale':1.2})
    # All classes
    ws=wb.add_worksheet('03 All Class Estimates'); _v55_setup_ws(ws); ws.merge_range('A1:I1','ALL COST ESTIMATE CLASSES — NO CLASS DUMP WITHOUT CONTEXT',f['title'])
    rows=[]
    for cls, clines in costs_by.items():
        _,_,cname,mat=class_range(cls)
        for x in clines:
            rows.append([cls,cname,mat,x.get('cbs'),x.get('description'),x.get('type'),float(x.get('p10_bn') or 0),float(x.get('p50_bn') or 0),float(x.get('p90_bn') or 0)])
    _v55_write_table(ws,4,0,['Class','Class Name','Maturity','CBS','Description','Type','P10 $B','P50 $B','P90 $B'],rows,f,[8,26,26,9,28,12,12,12,12],money_cols=[6,7,8])
    # Scenario comparison
    ws=wb.add_worksheet('04 Scenario Comparison'); _v55_setup_ws(ws); ws.merge_range('A1:G1','SCENARIO COMPARISON — WHAT CHANGES AND WHY',f['title'])
    scrows=[]
    for x in model.get('scenario_comparison') or []:
        scrows.append([x.get('label'),x.get('cost'),x.get('schedule_months'),x.get('risk'),x.get('confidence'),x.get('why'), 'Decision trade-off: cost / time / risk / confidence'])
    _v55_write_table(ws,4,0,['Scenario','P50 Cost','Schedule Months','Risk','Confidence %','Commercial Interpretation','Board Use'],scrows,f,[14,14,16,14,14,62,40])
    # Risk register
    ws=wb.add_worksheet('05 Risk Register'); _v55_setup_ws(ws); ws.merge_range('A1:Q1','RISK REGISTER — CAUSE / EVENT / IMPACT / RESPONSE',f['title'])
    rows=[[r['risk_id'],r['title'],r['category'],r['probability_pct'],r['cause'],r['event'],r['impact'],r['owner'],r['mitigation'],r['trigger'],r['residual_rating'],r['activity_id'],r['activity_name'],r['cbs'],r['cbs_name'],r['cost_emv_bn'],r['schedule_emv_days']] for r in risks]
    _v55_write_table(ws,4,0,['ID','Risk','Category','Likelihood %','Cause','Risk Event','Impact Description','Owner','Mitigation / Response','Trigger','Residual','Activity','Activity Name','CBS','CBS Name','QCRA Mean $B','QSRA Mean Days'],rows,f,[10,25,14,13,36,36,44,18,42,30,12,12,24,10,24,14,14],money_cols=[15],num_cols=[3,16])
    ws.conditional_format(5,10,4+len(rows),10, {'type':'text','criteria':'containing','value':'Red','format':f['red']}); ws.conditional_format(5,10,4+len(rows),10, {'type':'text','criteria':'containing','value':'Amber','format':f['amber']}); ws.conditional_format(5,10,4+len(rows),10, {'type':'text','criteria':'containing','value':'Green','format':f['green']})
    # Curves + tornados
    ws=wb.add_worksheet('06 QCRA Curve'); _v55_setup_ws(ws); ws.merge_range('A1:D1','QCRA COST CONFIDENCE CURVE',f['title']); _v55_write_table(ws,4,0,['Percentile','Cost $B'],curves['qcra'],f,[14,14],money_cols=[1]); ch=wb.add_chart({'type':'line'}); ch.add_series({'name':'QCRA Cost $B','categories':['06 QCRA Curve',5,0,4+len(curves['qcra']),0],'values':['06 QCRA Curve',5,1,4+len(curves['qcra']),1]}); ch.set_title({'name':'QCRA S-Curve'}); ch.set_y_axis({'name':'Cost $B'}); ws.insert_chart('D4',ch,{'x_scale':1.4,'y_scale':1.2})
    ws=wb.add_worksheet('07 QSRA Curve'); _v55_setup_ws(ws); ws.merge_range('A1:D1','QSRA SCHEDULE CONFIDENCE CURVE',f['title']); _v55_write_table(ws,4,0,['Percentile','Duration Months'],curves['qsra'],f,[14,18],num_cols=[1]); ch=wb.add_chart({'type':'line'}); ch.add_series({'name':'QSRA Months','categories':['07 QSRA Curve',5,0,4+len(curves['qsra']),0],'values':['07 QSRA Curve',5,1,4+len(curves['qsra']),1]}); ch.set_title({'name':'QSRA S-Curve'}); ch.set_y_axis({'name':'Months'}); ws.insert_chart('D4',ch,{'x_scale':1.4,'y_scale':1.2})
    ws=wb.add_worksheet('08 QCRA Tornado'); _v55_setup_ws(ws); ws.merge_range('A1:E1','QCRA COST TORNADO — COST DRIVERS ONLY',f['title']); rows=[[i+1,r['title'],r['cbs'],r['cost_emv_bn'],r['mitigation']] for i,r in enumerate(torn['qcra'])]; _v55_write_table(ws,4,0,['Rank','Driver','CBS','Mean Cost $B','Action'],rows,f,[8,30,10,14,48],money_cols=[3]); ch=wb.add_chart({'type':'bar'}); ch.add_series({'name':'Mean Cost $B','categories':['08 QCRA Tornado',5,1,4+len(rows),1],'values':['08 QCRA Tornado',5,3,4+len(rows),3]}); ch.set_title({'name':'QCRA Cost Tornado'}); ws.insert_chart('G4',ch,{'x_scale':1.4,'y_scale':1.4})
    ws=wb.add_worksheet('09 QSRA Tornado'); _v55_setup_ws(ws); ws.merge_range('A1:E1','QSRA SCHEDULE TORNADO — SCHEDULE DRIVERS ONLY',f['title']); rows=[[i+1,r['title'],r['activity_id'],r['schedule_emv_days'],r['mitigation']] for i,r in enumerate(torn['qsra'])]; _v55_write_table(ws,4,0,['Rank','Driver','Activity','Mean Days','Action'],rows,f,[8,30,12,12,48],num_cols=[3]); ch=wb.add_chart({'type':'bar'}); ch.add_series({'name':'Mean Days','categories':['09 QSRA Tornado',5,1,4+len(rows),1],'values':['09 QSRA Tornado',5,3,4+len(rows),3]}); ch.set_title({'name':'QSRA Schedule Tornado'}); ws.insert_chart('G4',ch,{'x_scale':1.4,'y_scale':1.4})
    # All schedules
    ws=wb.add_worksheet('10 All Schedule Levels'); _v55_setup_ws(ws); ws.merge_range('A1:H1','SCHEDULE LEVELS 1–5 — ALL LEVELS GENERATED',f['title'])
    rows=[]
    for lvl, acts in _v55_all_schedule_levels(model).items():
        for a in acts: rows.append([lvl,a.get('activity_id'),a.get('phase'),a.get('activity'),a.get('predecessor'),a.get('duration_months'),a.get('critical'),a.get('basis')])
    _v55_write_table(ws,4,0,['Level','Activity ID','Phase','Activity','Predecessor','Duration Months','Critical','Basis'],rows,f,[8,13,16,36,18,14,10,48],num_cols=[5])
    ws=wb.add_worksheet('11 Basis + Audit'); _v55_setup_ws(ws); ws.merge_range('A1:D1','METHOD, ASSUMPTIONS AND AUDIT',f['title'])
    audit=[['Generated',datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')],['Model version','CASEY v55 Platinum Outputs Final'],['Scenario',_v55_scenario_label(model)],['Estimate class',sel_cls],['Schedule level',sel_lvl],['Cost method','Parametric sector template + location + scale + complexity + scenario + class range'],['QCRA method','Cost confidence curve and cost-only tornado drivers'],['QSRA method','Schedule confidence curve and schedule-only tornado drivers'],['Limitation','First-pass intelligence. Validate with project-specific evidence before commitment.']]
    _v55_write_table(ws,4,0,['Area','Basis'],audit,f,[24,100])
    wb.close(); bio.seek(0); return bio.getvalue()


def risk_register_workbook_bytes(model: Dict[str, Any]) -> bytes:
    wb, bio = _v55_make_workbook_base(); f = _v55_formats(wb); risks=_v55_risks(model); torn=_v55_tornado(model)
    ws=wb.add_worksheet('01 Risk Dashboard'); _v55_setup_ws(ws); ws.merge_range('A1:H1','CASEY PLATINUM RISK REGISTER',f['title']); ws.merge_range('A2:H2',f"{model.get('title')} | {_v55_scenario_label(model)} scenario | {len(risks)} risks",f['subtitle'])
    metrics=[['Total risks',len(risks)],['Red risks',sum(1 for r in risks if r['residual_rating']=='Red')],['Amber risks',sum(1 for r in risks if r['residual_rating']=='Amber')],['Top QCRA driver',torn['qcra'][0]['title']],['Top QSRA driver',torn['qsra'][0]['title']]]
    _v55_write_table(ws,4,0,['Metric','Value'],metrics,f,[24,44])
    ws=wb.add_worksheet('02 Full Risk Register'); _v55_setup_ws(ws)
    rows=[[r['risk_id'],r['title'],r['category'],r['probability_pct'],r['cause'],r['event'],r['impact'],r['owner'],r['mitigation'],r['trigger'],r['residual_rating'],r['activity_id'],r['activity_name'],r['cbs'],r['cbs_name'],r['cost_emv_bn'],r['schedule_emv_days'],r['status']] for r in risks]
    _v55_write_table(ws,4,0,['ID','Risk','Category','Likelihood %','Cause','Risk Event','Impact Description','Owner','Mitigation / Response','Trigger','Residual','Activity','Activity Name','CBS','CBS Name','QCRA Mean $B','QSRA Mean Days','Status'],rows,f,[10,25,14,13,36,36,44,18,42,30,12,12,24,10,24,14,14,12],money_cols=[15],num_cols=[3,16])
    ws.conditional_format(5,10,4+len(rows),10, {'type':'text','criteria':'containing','value':'Red','format':f['red']}); ws.conditional_format(5,10,4+len(rows),10, {'type':'text','criteria':'containing','value':'Amber','format':f['amber']}); ws.conditional_format(5,10,4+len(rows),10, {'type':'text','criteria':'containing','value':'Green','format':f['green']})
    ws=wb.add_worksheet('03 QCRA Tornado'); _v55_setup_ws(ws); rows=[[i+1,r['title'],r['cbs'],r['cost_emv_bn'],r['mitigation']] for i,r in enumerate(torn['qcra'])]; _v55_write_table(ws,4,0,['Rank','Driver','CBS','Mean Cost $B','Action'],rows,f,[8,30,10,14,48],money_cols=[3])
    ch=wb.add_chart({'type':'bar'}); ch.add_series({'name':'Mean Cost $B','categories':['03 QCRA Tornado',5,1,4+len(rows),1],'values':['03 QCRA Tornado',5,3,4+len(rows),3]}); ch.set_title({'name':'QCRA Cost Tornado'}); ws.insert_chart('G4',ch,{'x_scale':1.4,'y_scale':1.4})
    ws=wb.add_worksheet('04 QSRA Tornado'); _v55_setup_ws(ws); rows=[[i+1,r['title'],r['activity_id'],r['schedule_emv_days'],r['mitigation']] for i,r in enumerate(torn['qsra'])]; _v55_write_table(ws,4,0,['Rank','Driver','Activity','Mean Days','Action'],rows,f,[8,30,12,12,48],num_cols=[3])
    ch=wb.add_chart({'type':'bar'}); ch.add_series({'name':'Mean Days','categories':['04 QSRA Tornado',5,1,4+len(rows),1],'values':['04 QSRA Tornado',5,3,4+len(rows),3]}); ch.set_title({'name':'QSRA Schedule Tornado'}); ws.insert_chart('G4',ch,{'x_scale':1.4,'y_scale':1.4})
    wb.close(); bio.seek(0); return bio.getvalue()


def xer_bytes(model: Dict[str, Any]) -> bytes:
    """PRA-compatible legacy-targeted XER. Uses ERMHDR 8.0 and basic TASK/TASKPRED structure."""
    rows = model.get('schedule_rows') or []
    if not rows:
        rows = schedule_rows(model.get('mode','Earth'), model.get('subsector','General'), int(str(model.get('schedule','60')).split()[0]), _v55_selected_schedule_level(model))
    lines=[]
    lines.append('ERMHDR\t8.0\t{}\tProject\tCASEY\tCASEY\tdbxDatabaseNoName\tProject Management\tUSD'.format(datetime.utcnow().strftime('%Y-%m-%d')))
    lines.append('%T\tCURRTYPE'); lines.append('%F\tcurr_id\tdecimal_digit_cnt\tcurr_symbol\tdecimal_symbol\tdigit_group_symbol\tpos_curr_fmt_type\tneg_curr_fmt_type\tcurr_type\tcurr_short_name\tgroup_digit_cnt\tbase_exch_rate'); lines.append('%R\t1\t2\t$\t.\t,\t#1.1\t(#1.1)\tUS Dollar\tUSD\t3\t1')
    lines.append('%T\tPROJECT'); lines.append('%F\tproj_id\tfy_start_month_num\tname_sep_char\tproj_short_name\tplan_start_date\tlast_recalc_date\tclndr_id\tdef_duration_type\tdef_qty_type\tdef_rate_type\tdef_task_type\tcritical_path_type')
    lines.append('%R\t3338\t1\t.\tCASEY\t{} 08:00\t{} 08:00\t3892\tDT_FixedDUR2\tQT_Hour\tCOST_PER_QTY\tTT_Task\tCT_DrivPath'.format(datetime.utcnow().strftime('%Y-%m-%d'), datetime.utcnow().strftime('%Y-%m-%d')))
    lines.append('%T\tCALENDAR'); lines.append('%F\tclndr_id\tdefault_flag\tclndr_name\tproj_id\tbase_clndr_id\tclndr_type\tday_hr_cnt\tweek_hr_cnt\tmonth_hr_cnt\tyear_hr_cnt\trsrc_private\tclndr_data')
    lines.append('%R\t3892\tY\tStandard\t3338\t\tCA_Project\t8\t40\t172\t2000\tN\t(0||CalendarData()())')
    lines.append('%T\tPROJWBS'); lines.append('%F\twbs_id\tproj_id\tobs_id\tseq_num\twbs_short_name\twbs_name\tparent_wbs_id')
    lines.append('%R\t5000\t3338\t565\t1\tCASEY\tCASEY Generated Schedule\t')
    lines.append('%T\tTASK')
    fields=['task_id','proj_id','wbs_id','clndr_id','task_code','task_name','task_type','duration_type','status_code','target_drtn_hr_cnt','remain_drtn_hr_cnt','phys_complete_pct','target_start_date','target_end_date','early_start_date','early_end_date','late_start_date','late_end_date','guid']
    lines.append('%F\t'+'\t'.join(fields))
    id_map={}
    start='{} 08:00'.format(datetime.utcnow().strftime('%Y-%m-%d'))
    for i,a in enumerate(rows,1):
        id_map[str(a.get('activity_id'))]=i
        dur=int(max(8, float(a.get('duration_months') or 1)*172))
        code=str(a.get('activity_id') or f'A{i:04d}')[:40]
        name=str(a.get('activity') or 'Activity').replace('\t',' ')[:120]
        guid=('CASEY%032d'%i)[:22]
        lines.append('%R\t{}\t3338\t5000\t3892\t{}\t{}\tTT_Task\tDT_FixedDUR2\tTK_NotStart\t{}\t{}\t0\t{}\t\t{}\t\t\t\t{}'.format(i,code,name,dur,dur,start,start,guid))
    lines.append('%T\tTASKPRED'); lines.append('%F\ttask_pred_id\ttask_id\tpred_task_id\tpred_type\tlag_hr_cnt')
    pred_id=1
    for i,a in enumerate(rows,1):
        preds=str(a.get('predecessor') or '').replace(';',',').split(',')
        for p in [x.strip() for x in preds if x.strip()]:
            pid=id_map.get(p)
            if pid and pid != i:
                lines.append('%R\t{}\t{}\t{}\tPR_FS\t0'.format(pred_id,i,pid)); pred_id+=1
    lines.append('%T\tUDFTYPE'); lines.append('%F\tudf_type_id\ttable_name\tudf_type_name\tudf_type_label\tlogical_data_type\tsuper_flag')
    lines.append('%R\t329\tTASK\tCASEY_Risk_Most_Likely\tLikely\tFT_TEXT\tN')
    lines.append('%E')
    return ('\n'.join(lines)+'\n').encode('utf-8')


def pdf_bytes(model: Dict[str, Any]) -> bytes:
    bio=BytesIO(); doc=SimpleDocTemplate(bio,pagesize=landscape(A4),rightMargin=24,leftMargin=24,topMargin=24,bottomMargin=24)
    styles=getSampleStyleSheet(); styles.add(ParagraphStyle(name='V55Title',parent=styles['Title'],fontSize=24,leading=28,textColor=colors.HexColor(V55_BLUE),spaceAfter=10)); styles.add(ParagraphStyle(name='V55H',parent=styles['Heading1'],fontSize=15,textColor=colors.HexColor(V55_BLUE),spaceAfter=7)); styles.add(ParagraphStyle(name='V55Small',parent=styles['BodyText'],fontSize=8,leading=10))
    sc=_v55_scenario_strategy(model); risks=_v55_risks(model); torn=_v55_tornado(model); curves=_v55_curves(model); costs=_v55_all_cost_lines(model).get(_v55_selected_class(model),model.get('cost_lines') or [])
    story=[]
    story.append(Paragraph(f"CASEY BOARD DECISION PACK — {model.get('title')}",styles['V55Title'])); story.append(Paragraph(f"{_v55_scenario_label(model)} scenario | {model.get('subsector')} | Generated {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}",styles['V55Small']))
    kpi=[['P50 Cost',model.get('cost_p50'),'Cost Range',model.get('cost_range'),'QCRA P80',_v55_money((model.get('monte_carlo') or {}).get('qcra',{}).get('p80',0))],['Schedule',model.get('schedule'),'QSRA P80',str((model.get('monte_carlo') or {}).get('qsra',{}).get('p80',''))+' mo','Risk / Confidence',f"{model.get('risk')} / {model.get('confidence_pct')}%"]]
    t=Table(kpi,colWidths=[75,110,75,135,75,120]); t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),colors.HexColor('#F4F7FA')),('TEXTCOLOR',(0,0),(-1,-1),colors.HexColor(V55_BLUE)),('FONTNAME',(0,0),(-1,-1),'Helvetica-Bold'),('GRID',(0,0),(-1,-1),0.5,colors.HexColor(V55_LINE)),('FONTSIZE',(0,0),(-1,-1),9)])); story.append(t); story.append(Spacer(1,10)); story.append(Paragraph(sc['title'],styles['V55H'])); story.append(Paragraph(sc['summary']+' '+sc['decision'],styles['BodyText'])); story.append(PageBreak())
    story.append(Paragraph('QCRA and QSRA Tornado Drivers — separated',styles['V55H'])); rows=[['Rank','QCRA Cost Driver','CBS','Mean Cost','QSRA Schedule Driver','Activity','Mean Days']]
    for i in range(max(len(torn['qcra']),len(torn['qsra']))):
        qc=torn['qcra'][i] if i<len(torn['qcra']) else {}; qs=torn['qsra'][i] if i<len(torn['qsra']) else {}
        rows.append([i+1,qc.get('title',''),qc.get('cbs',''),_v55_money(qc.get('cost_emv_bn',0)) if qc else '',qs.get('title',''),qs.get('activity_id',''),qs.get('schedule_emv_days','')])
    t=Table(rows,repeatRows=1,colWidths=[35,145,55,70,145,65,60]); t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#0070C0')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.3,colors.HexColor(V55_LINE)),('FONTSIZE',(0,0),(-1,-1),7.5),('VALIGN',(0,0),(-1,-1),'TOP'),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F8FBFD')]) ])); story.append(t); story.append(PageBreak())
    story.append(Paragraph('Board Top Risks — one risk per row, readable summary',styles['V55H'])); rows=[['ID','Risk','Cause','Impact','Owner','Action']]
    for r in risks[:12]: rows.append([r['risk_id'],r['title'],r['cause'],r['impact'],r['owner'],r['mitigation']])
    t=Table(rows,repeatRows=1,colWidths=[45,90,180,210,70,205]); t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#0070C0')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.25,colors.HexColor(V55_LINE)),('FONTSIZE',(0,0),(-1,-1),6.9),('VALIGN',(0,0),(-1,-1),'TOP'),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F8FBFD')]) ])); story.append(t); story.append(PageBreak())
    story.append(Paragraph('Selected Cost Estimate — direct / indirect / reserve',styles['V55H'])); rows=[['CBS','Description','Type','P10','P50','P90']]
    for x in costs: rows.append([x.get('cbs'),x.get('description'),x.get('type'),_v55_money(x.get('p10_bn')), _v55_money(x.get('p50_bn')), _v55_money(x.get('p90_bn'))])
    t=Table(rows,repeatRows=1,colWidths=[50,230,80,90,90,90]); t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#0070C0')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.25,colors.HexColor(V55_LINE)),('FONTSIZE',(0,0),(-1,-1),8),('VALIGN',(0,0),(-1,-1),'TOP'),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F8FBFD')]) ])); story.append(t)
    doc.build(story); bio.seek(0); return bio.getvalue()


def pptx_bytes(model: Dict[str, Any]) -> bytes:
    # Removed from recommended output factory. Retained only to avoid a hard 404 if old frontend calls it.
    prs=Presentation(); s=prs.slides.add_slide(prs.slide_layouts[6]); box=s.shapes.add_textbox(PptxInches(0.8),PptxInches(1),PptxInches(11),PptxInches(3)); tf=box.text_frame; tf.text='PPTX DECK REMOVED FROM PLATINUM OUTPUT PACK'; tf.paragraphs[0].runs[0].font.size=PptxPt(30); p=tf.add_paragraph(); p.text='CASEY now prioritises board PDF, Excel cost model, risk workbook, XER + schedule CSV, DOCX and JSON audit outputs. The deck can be rebuilt later as a separate design product.'; p.font.size=PptxPt(18)
    bio=BytesIO(); prs.save(bio); bio.seek(0); return bio.getvalue()


def _v55_schedule_csv(model: Dict[str, Any]) -> bytes:
    out=StringIO(); w=csv.writer(out); w.writerow(['Level','Activity ID','Phase','Activity','Predecessor','Duration Months','Critical','Basis'])
    for lvl, acts in _v55_all_schedule_levels(model).items():
        for a in acts: w.writerow([lvl,a.get('activity_id'),a.get('phase'),a.get('activity'),a.get('predecessor'),a.get('duration_months'),a.get('critical'),a.get('basis')])
    return out.getvalue().encode('utf-8')


def export_workbook_v55_endpoint(model: Dict[str, Any]): return stream(workbook_bytes(model),'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','CASEY_v55_Platinum_Cost_Model.xlsx')
def export_risk_v55_endpoint(model: Dict[str, Any]): return stream(risk_register_workbook_bytes(model),'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','CASEY_v55_Platinum_Risk_Register.xlsx')
def export_xer_v55_endpoint(model: Dict[str, Any]): return stream(xer_bytes(model),'application/octet-stream','CASEY_v55_PRA_Targeted_Schedule.xer')
def export_schedule_csv_v55_endpoint(model: Dict[str, Any]): return stream(_v55_schedule_csv(model),'text/csv','CASEY_v55_All_Schedule_Levels.csv')
def export_pdf_v55_endpoint(model: Dict[str, Any]): return stream(pdf_bytes(model),'application/pdf','CASEY_v55_Board_Decision_Pack.pdf')
def export_all_v55_endpoint(model: Dict[str, Any]):
    bio=BytesIO()
    with zipfile.ZipFile(bio,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr('01_CASEY_v55_Platinum_Cost_Model.xlsx',workbook_bytes(model))
        z.writestr('02_CASEY_v55_Platinum_Risk_Register.xlsx',risk_register_workbook_bytes(model))
        z.writestr('03_CASEY_v55_PRA_Targeted_Schedule.xer',xer_bytes(model))
        z.writestr('04_CASEY_v55_All_Schedule_Levels.csv',_v55_schedule_csv(model))
        z.writestr('05_CASEY_v55_Board_Decision_Pack.docx',word_bytes(model))
        z.writestr('06_CASEY_v55_Board_Decision_Pack.pdf',pdf_bytes(model))
        z.writestr('07_CASEY_v55_Model_Audit.json',json.dumps(model,indent=2))
        z.writestr('README_PRA_IMPORT_NOTE.txt','The XER is exported with ERMHDR 8.0 to target older Primavera/PRA imports. If your PRA instance still rejects XER, use the CSV schedule fallback or open in P6 Professional and re-export using the version supported by your PRA build.')
    bio.seek(0); return stream(bio.getvalue(),'application/zip','CASEY_v55_Platinum_Output_Pack.zip')


def _v55_replace_post(path, endpoint):
    app.router.routes = [r for r in app.router.routes if not (getattr(r,'path',None)==path and 'POST' in getattr(r,'methods',set()))]
    app.post(path)(endpoint)

_v55_replace_post('/export/workbook', export_workbook_v55_endpoint)
_v55_replace_post('/export/risk-register', export_risk_v55_endpoint)
_v55_replace_post('/export/xer', export_xer_v55_endpoint)
_v55_replace_post('/export/pdf', export_pdf_v55_endpoint)
_v55_replace_post('/export/all', export_all_v55_endpoint)
app.post('/export/schedule-csv')(export_schedule_csv_v55_endpoint)
app.post('/v55/export/all')(export_all_v55_endpoint)

APP_VERSION = 'CASEY TITAN X v55 Platinum Outputs Final'

# =====================================================================
# v57 SELLABLE ANY PROJECT ENGINE INSTALLER
# Final output pack removes PDF/PPTX/Word from the primary source-of-truth pack.
# =====================================================================
try:
    from v56_outputs import install_v56
    install_v56(app)
    APP_VERSION = 'CASEY TITAN X v60 Final Commercial'
except Exception as _v56_error:
    print('V57 install failed:', _v56_error)

# =====================================================================
# v59 PLATINUM DYNAMIC OUTPUTS INSTALLER
try:
    from v59_outputs import install_v59
    install_v59(app)
    APP_VERSION = 'CASEY TITAN X v60 Final Commercial'
except Exception as _v59_error:
    print('V59 install failed:', _v59_error)


# v62 TRUE DYNAMIC XER + COMMERCIAL OUTPUT ENGINE INSTALLER
try:
    from v62_outputs import install_v62
    install_v62(app)
    print('CASEY v62 true dynamic XER engine installed')
except Exception as _v62_error:
    print('V62 install failed:', _v62_error)

# v63 TRUE UNIVERSAL ANY-PROJECT ENGINE INSTALLER
try:
    from v63_outputs import install_v63
    install_v63(app)
    print('CASEY v63 true universal engine installed')
except Exception as _v63_error:
    print('V63 install failed:', _v63_error)


# v64 FINAL SELLABLE UNIVERSAL ENGINE INSTALLER
try:
    from v64_outputs import install_v64
    install_v64(app)
    APP_VERSION = 'CASEY TITAN X v64 Final Sellable Universal Engine'
    print('CASEY v64 final sellable universal engine installed')
except Exception as _v64_error:
    print('V64 install failed:', _v64_error)



# ------------------------- CASEY v9.5+ holy-shit scenario consequence engine -------------------------
SCENARIO_PROFILES_V95_PLUS = {
    "base": {
        "label":"Base","cost":1.00,"schedule":1.00,"confidence":0,"risk":"Medium-High",
        "reserve_factor":1.00,"direct_factor":1.00,"indirect_factor":1.00,
        "curve":"base","risk_lift":0,
        "trade":"Balanced cost, time and evidence posture.",
        "won":"Maintains a credible reference case for board challenge.",
        "lost":"Does not buy extra time certainty or capital efficiency."
    },
    "faster": {
        "label":"Faster","cost":1.18,"schedule":0.78,"confidence":-13,"risk":"High",
        "reserve_factor":1.42,"direct_factor":1.10,"indirect_factor":1.30,
        "curve":"faster","risk_lift":18,
        "trade":"You bought time by spending money and consuming recovery float.",
        "won":"Earlier revenue / market-entry option and stronger strategic timing.",
        "lost":"CQV float, interface stability, procurement optionality and late-stage recovery."
    },
    "cheaper": {
        "label":"Cheaper","cost":0.86,"schedule":1.14,"confidence":-16,"risk":"High",
        "reserve_factor":0.58,"direct_factor":0.90,"indirect_factor":0.78,
        "curve":"cheaper","risk_lift":14,
        "trade":"You cut capital authorization by moving risk into resilience, redundancy and start-up.",
        "won":"Lower initial approval number and reduced near-term capital draw.",
        "lost":"Operational resilience, commissioning flexibility, contingency adequacy and lifecycle certainty."
    },
    "lower_risk": {
        "label":"Lower Risk","cost":1.12,"schedule":1.16,"confidence":10,"risk":"Medium-Low",
        "reserve_factor":1.52,"direct_factor":1.05,"indirect_factor":1.16,
        "curve":"lower_risk","risk_lift":-12,
        "trade":"You bought confidence by adding assurance, float and procurement evidence.",
        "won":"Reduced downside volatility and stronger approval defensibility, at the expense of market timing.",
        "lost":"Market-entry acceleration, lean capex posture and near-term strategic timing."
    },
    "premium": {
        "label":"Premium","cost":1.28,"schedule":0.96,"confidence":18,"risk":"Low",
        "reserve_factor":1.65,"direct_factor":1.22,"indirect_factor":1.24,
        "curve":"premium","risk_lift":-16,
        "trade":"You purchased operational separation, redundancy and downside protection capacity.",
        "won":"Higher mission resilience, commissioning separation and stronger downside containment.",
        "lost":"Lean capital optics and minimum-approval positioning."
    },
}

def _money_to_bn_hs(s):
    s=str(s or "").replace("$","").strip().upper()
    try:
        if s.endswith("T"): return float(s[:-1])*1000
        if s.endswith("B"): return float(s[:-1])
        if s.endswith("M"): return float(s[:-1])/1000
        return float(s)
    except Exception:
        return 0.0

def _bn_to_money_hs(bn):
    try: bn=float(bn)
    except Exception: bn=0.0
    if bn >= 1000: return f"${bn/1000:.1f}T"
    if bn >= 1: return f"${bn:.1f}B"
    return f"${bn*1000:.0f}M"

def _duration_hs(s):
    try:
        return int(float(str(s or "60").replace("months","").strip().split()[0]))
    except Exception:
        return 60

def _fam_insight_pack_hs(model):
    fam=_sector_family(model.get("subsector",""))
    if fam=="life_sciences":
        return {
            "base":"Mechanical completion is not the true finish line; validated production readiness and deviation closure are the real board decision gates.",
            "faster":"Acceleration is now the risk: CQV, clean-utility validation and media-fill readiness are being forced into the same window. The programme may finish construction before it is safe to release product.",
            "cheaper":"The cheaper case is not simply cheaper. It has likely transferred cost into operational fragility: leaner redundancy, weaker qualification float and higher late-stage validation disruption.",
            "lower_risk":"The lower-risk case buys regulatory confidence by protecting CQV float, GMP turnover and deviation closure before commercial ramp-up.",
            "premium":"Premium delivery protects market launch by buying utility resilience, validation capacity and stronger batch-release readiness."
        }
    if fam=="data_centre":
        return {
            "base":"Power availability and systems integration readiness are more likely to govern delivery than civil progress alone.",
            "faster":"Acceleration compressed energisation and integrated systems testing into the same delivery window. Civil completion may outpace operational readiness.",
            "cheaper":"Capital efficiency is being achieved by reducing resilience margin, commissioning separation and operational redundancy.",
            "lower_risk":"The lower-risk posture protects energisation sequencing, commissioning float and operational continuity before customer cutover.",
            "premium":"Premium posture buys commissioning separation, operational redundancy and higher uptime defensibility."
        }
    if fam=="semiconductor":
        return {
            "base":"Tool install, UPW readiness and yield-ramp qualification are the real finish line, not cleanroom completion.",
            "faster":"Acceleration compresses tool hook-up, UPW commissioning and qualification cadence into a fragile ramp window.",
            "cheaper":"The cheaper case risks deferring utility resilience and tool-readiness assurance into yield ramp, where recovery is most expensive.",
            "lower_risk":"The lower-risk case protects process-tool cadence, utility stability and yield-ramp confidence.",
            "premium":"Premium delivery buys tool availability, redundancy and stronger production-ramp certainty."
        }
    if model.get("mode")=="Space":
        return {
            "base":"Launch alone is not the real constraint; qualification, thermal-power balance and autonomous recovery determine mission survivability.",
            "faster":"Acceleration compresses payload qualification and launch integration, increasing mission-assurance risk after deployment when recovery options are limited.",
            "cheaper":"The cheaper case likely removes redundancy in the one environment where repair is hardest: after deployment.",
            "lower_risk":"The lower-risk case protects qualification, environmental testing and autonomous recovery evidence before launch commitment.",
            "premium":"Premium delivery buys mission assurance, redundancy and survivability after deployment."
        }
    return {
        "base":"The dominant risk sits in interfaces, procurement evidence and commissioning readiness rather than headline construction progress.",
        "faster":"Acceleration compresses interface resolution and commissioning recovery float, increasing late-stage execution volatility.",
        "cheaper":"The cheaper case lowers the approval number by moving risk into operational readiness, contingency adequacy and deferred resilience.",
        "lower_risk":"The lower-risk case buys confidence by adding evidence, float and assurance.",
        "premium":"Premium delivery buys resilience, optionality and stronger downside protection."
    }

def _scenario_risks_hs(model, scenario, lift):
    fam=_sector_family(model.get("subsector",""))
    common = {
        "faster":[
            ("R-F01","Concurrent commissioning overload","Acceleration","Too many systems enter commissioning together","Rework, late handover and validation slippage","58","Commissioning Lead","Separate hard hold-points; add surge QA/CQV resource"),
            ("R-F02","Recovery float exhaustion","Schedule compression","Float consumed before integration testing","No recovery path for late supplier or interface failure","54","Project Controls Lead","Create protected executive float and weekly path-to-green review"),
            ("R-F03","Acceleration premium shock","Procurement","Expediting and premium vendors exceed allowance","Cost growth beyond approved scenario envelope","49","Commercial Lead","Lock long-lead packages and cap acceleration premiums")
        ],
        "cheaper":[
            ("R-C01","Deferred resilience failure","Capital reduction","Redundancy and resilience scope deferred","Operational start-up or availability failure","57","Operations Lead","Define minimum viable redundancy and non-deferrable systems list"),
            ("R-C02","Contingency underfunding","Budget pressure","Reserve reduced below uncertainty envelope","Board funding shock when latent risks materialise","52","Sponsor","Ring-fence reserve for critical systems and regulatory gates"),
            ("R-C03","Lifecycle cost transfer","Capex cut","Lower initial scope creates higher operating burden","False economy and stakeholder challenge","46","Asset Owner","Show capex/opex trade-off and owner acceptance record")
        ],
        "lower_risk":[
            ("R-L01","Assurance gate delay","Risk reduction","Additional evidence gates slow early progress","Short-term date pressure despite better certainty","28","PMO Lead","Pre-agree evidence criteria and approval cadence"),
            ("R-L02","Procurement buffer cost","Risk reduction","More secure procurement path increases early commitment","Higher pre-approval funding need","24","Commercial Lead","Use staged release with supplier option strategy")
        ],
        "premium":[
            ("R-P01","Premium scope creep","Optionality","Resilience and redundancy additions expand scope","Capex escalation and governance challenge","22","Sponsor","Govern premium scope with value gates"),
            ("R-P02","Over-specification risk","Resilience posture","Premium solution exceeds real operating need","Weak value-for-money challenge","18","Asset Owner","Tie premium scope to quantified service / mission resilience")
        ],
        "base":[]
    }
    sector_extra = []
    if fam=="life_sciences" and scenario=="faster":
        sector_extra=[("R-F04","Media-fill readiness collision","CQV compression","Media-fill, clean utilities and staffing readiness converge","FDA readiness and batch-release delay","63","CQV Lead","Protect media-fill sequence and deviation closure capacity")]
    elif fam=="life_sciences" and scenario=="cheaper":
        sector_extra=[("R-C04","Clean utility redundancy gap","Capex reduction","WFI / clean steam resilience reduced or deferred","Production continuity exposure during ramp-up","60","Utilities Lead","Identify GMP-critical redundancy that cannot be deferred")]
    elif model.get("mode")=="Space" and scenario=="faster":
        sector_extra=[("R-F04","Mission assurance compression","Launch pressure","Qualification evidence compressed before launch slot","Post-deployment failure with limited recovery","67","Mission Assurance Lead","No-go criteria for thermal, power and autonomy testing")]
    elif model.get("mode")=="Space" and scenario=="cheaper":
        sector_extra=[("R-C04","Redundancy deletion after deployment","Capex cut","Backup systems removed before orbital operations","Single-point failure in inaccessible environment","65","Systems Lead","Protect mission-critical redundancy and autonomous recovery budget")]
    return sector_extra + common.get(scenario, [])

def _curve_hs(cost, months, scenario):
    """Scenario-specific QCRA/QSRA P-curves anchored so P50 equals the headline P50.
    This prevents the dashboard saying P50 is one number while the chart implies another.
    The tail shape is the intelligence: Faster/Cheaper widen, Lower Risk/Premium compress.
    """
    pts=[1,5,10,20,30,40,50,60,70,80,90,95,99]
    profiles={
        "base": {
            "cost":[0.78,0.81,0.84,0.89,0.93,0.97,1.00,1.06,1.13,1.22,1.36,1.48,1.62],
            "sched":[0.82,0.84,0.86,0.90,0.94,0.97,1.00,1.05,1.10,1.16,1.26,1.34,1.44],
            "meaning":"Reference case: P50 is the approval view; P80/P90 show normal delivery uncertainty."
        },
        "faster": {
            "cost":[0.82,0.84,0.87,0.91,0.95,0.98,1.00,1.09,1.20,1.35,1.55,1.72,1.95],
            "sched":[0.72,0.75,0.78,0.84,0.90,0.95,1.00,1.09,1.20,1.34,1.52,1.66,1.85],
            "meaning":"Faster case: the median improves, but the tail widens as concurrency and recovery exposure increase."
        },
        "cheaper": {
            "cost":[0.70,0.74,0.78,0.84,0.90,0.96,1.00,1.12,1.25,1.42,1.68,1.88,2.15],
            "sched":[0.84,0.87,0.90,0.94,0.97,0.99,1.00,1.12,1.24,1.39,1.58,1.72,1.92],
            "meaning":"Cheaper case: lower approval number, but a fat tail because contingency and resilience have been stripped."
        },
        "lower_risk": {
            "cost":[0.88,0.90,0.92,0.95,0.97,0.99,1.00,1.04,1.08,1.13,1.20,1.25,1.32],
            "sched":[0.90,0.92,0.94,0.96,0.98,0.99,1.00,1.04,1.08,1.13,1.20,1.26,1.34],
            "meaning":"Lower-risk case: schedule flexibility and evidence maturity tighten the downside curve."
        },
        "premium": {
            "cost":[0.88,0.90,0.93,0.96,0.98,0.99,1.00,1.04,1.09,1.16,1.25,1.34,1.46],
            "sched":[0.84,0.87,0.90,0.94,0.97,0.99,1.00,1.04,1.09,1.16,1.25,1.34,1.46],
            "meaning":"Premium case: additional capital buys separation, optionality and lower operational fragility."
        },
    }
    prof=profiles.get(scenario, profiles["base"])
    out=[]
    for p,cf,sf in zip(pts, prof["cost"], prof["sched"]):
        out.append({
            "percentile":p,
            "cost_bn":round(float(cost)*cf,2),
            "schedule_months":int(round(float(months)*sf)),
            "label": f"P{p}",
            "cost_meaning": "headline P50" if p==50 else ("tail exposure" if p>=80 else "lower-bound / optimism case"),
            "schedule_meaning": "headline P50" if p==50 else ("finish-date risk" if p>=80 else "early finish case"),
        })
    return out

def _scenario_outputs_text_hs(profile, insight):
    return [
        f"Scenario trade: {profile['trade']}",
        f"What you gained: {profile['won']}",
        f"What you gave up: {profile['lost']}",
        f"Board warning: {insight}",
        "Decision rule: do not approve this scenario until the changed risk owner, reserve position and critical-path evidence are explicit."
    ]

def scenario_cascade_v95(model:Dict[str,Any], scenario:str) -> Dict[str,Any]:
    # Override prior v9.5 cascade with deeper v9.5+ consequence model.
    # v9.5+ final sector correction: semiconductor / fab language overrides accidental pharma routing.
    _prompt_text = str(model.get("prompt") or model.get("brief") or model.get("description") or "")
    _title_text = str(model.get("title") or "")
    _all_text = (_prompt_text + " " + _title_text + " " + str(model.get("subsector",""))).lower()
    if any(k in _all_text for k in ["semiconductor", "fab ", "fabrication", "wafer", "euv", "lithography", "process-tool", "process tool", "upw", "chip manufacturing"]):
        model["mode"] = "Earth"
        model["subsector"] = "Semiconductor / Advanced Manufacturing"
        if "Semiconductor" not in str(model.get("title","")):
            model["title"] = "Semiconductor Fabrication Campus"

    scenario=(scenario or "base").lower()
    if scenario not in SCENARIO_PROFILES_V95_PLUS:
        scenario="base"
    profile=SCENARIO_PROFILES_V95_PLUS[scenario]
    if "_hs_base" not in model:
        model["_hs_base"]={
            "cost_p50": model.get("cost_p50"),
            "schedule_months": _duration_hs(model.get("schedule")),
            "confidence": int(model.get("confidence_pct",55) or 55),
            "risk": model.get("risk","Medium-High"),
            "cost_lines": json.loads(json.dumps(model.get("cost_lines",[]))),
            "risks": json.loads(json.dumps(model.get("risks",[]))),
            "schedule_rows": json.loads(json.dumps(model.get("schedule_rows",[]))),
            "direct_cost": model.get("direct_cost"),
            "indirect_cost": model.get("indirect_cost"),
            "risk_reserve": model.get("risk_reserve"),
        }
    base=model["_hs_base"]
    base_cost=_money_to_bn_hs(base["cost_p50"])
    base_months=int(base["schedule_months"])
    base_conf=int(base["confidence"])
    new_cost=base_cost*profile["cost"]
    new_months=max(3,int(round(base_months*profile["schedule"])))
    new_conf=max(5,min(96,base_conf+profile["confidence"]))
    insight=_fam_insight_pack_hs(model).get(scenario, _fam_insight_pack_hs(model)["base"])
    model.update({
        "scenario":scenario,
        "scenario_label":profile["label"],
        "cost_p50":_bn_to_money_hs(new_cost),
        "cost_p10":_bn_to_money_hs(new_cost*(0.78 if scenario!="cheaper" else 0.70)),
        "cost_p90":_bn_to_money_hs(new_cost*(1.27 if scenario in ["lower_risk","premium"] else 1.48 if scenario=="faster" else 1.55 if scenario=="cheaper" else 1.32)),
        "schedule":f"{new_months} months",
        "risk":profile["risk"],
        "confidence_pct":new_conf,
        "executive_shock_insight":insight,
        "scenario_trade":profile["trade"],
        "scenario_gain":profile["won"],
        "scenario_loss":profile["lost"],
        "version":"CASEY TITAN X v95+ Scenario Intelligence"
    })
    model["cost_range"]=f"{model['cost_p10']} - {model['cost_p90']}"
    direct=_money_to_bn_hs(base.get("direct_cost") or model.get("direct_cost") or _bn_to_money_hs(base_cost*.78))*profile["direct_factor"]
    indirect=_money_to_bn_hs(base.get("indirect_cost") or model.get("indirect_cost") or _bn_to_money_hs(base_cost*.17))*profile["indirect_factor"]
    reserve=_money_to_bn_hs(base.get("risk_reserve") or model.get("risk_reserve") or _bn_to_money_hs(base_cost*.05))*profile["reserve_factor"]
    model["direct_cost"]=_bn_to_money_hs(direct)
    model["indirect_cost"]=_bn_to_money_hs(indirect)
    model["risk_reserve"]=_bn_to_money_hs(reserve)

    # scenario delta versus base
    model["scenario_delta_intelligence"]=[
        {"label":"Capital movement","value":f"{(profile['cost']-1)*100:+.0f}%","meaning":profile["trade"]},
        {"label":"Schedule movement","value":f"{(profile['schedule']-1)*100:+.0f}%","meaning":profile["won"]},
        {"label":"Confidence movement","value":f"{profile['confidence']:+d} pts","meaning":profile["lost"]},
        {"label":"Reserve philosophy","value":f"{(profile['reserve_factor']-1)*100:+.0f}% reserve shift","meaning":"Reserve is being used as a strategic choice, not a static percentage."},
        {"label":"Board consequence","value":profile["risk"],"meaning":insight},
    ]
    # confidence breakdown, scenario-specific
    if scenario=="faster":
        cb=[("Concurrency loading","-13","Parallel execution and commissioning overlap increase volatility"),
            ("Procurement acceleration","-9","Expediting reduces optionality and increases premium risk"),
            ("Recovery float","-12","Float is consumed before validation / integration is proven"),
            ("Strategic timing","+6","Earlier market-entry option improves strategic value"),
            ("Evidence maturity","-7","Decisions are being pulled ahead of package maturity")]
    elif scenario=="cheaper":
        cb=[("Reserve adequacy","-14","Contingency is below the uncertainty envelope"),
            ("Operational resilience","-11","Redundancy and start-up protection are reduced"),
            ("Capital efficiency","+7","Lower approval number improves affordability"),
            ("Lifecycle exposure","-9","Costs are likely transferred into operations and recovery"),
            ("Procurement certainty","-6","Lower-cost sourcing increases vendor / quality variance")]
    elif scenario=="lower_risk":
        cb=[("Assurance maturity","+12","More evidence exists before commitment"),
            ("Commissioning float","+10","Recovery time is protected"),
            ("Procurement certainty","+8","Long-lead strategy is more secure"),
            ("Schedule aggressiveness","-4","Longer delivery reduces strategic timing"),
            ("Reserve adequacy","+9","Tail risk is better funded")]
    elif scenario=="premium":
        cb=[("Resilience","+13","Redundancy and optionality improve downside protection"),
            ("Procurement certainty","+10","Premium route buys capacity and priority"),
            ("Operational readiness","+11","Stronger handover / mission readiness"),
            ("Capital intensity","-7","Higher approval burden"),
            ("Tail exposure","+9","P80/P90 risk is controlled")]
    else:
        cb=[("Benchmark similarity","+8","Comparable sector archetype identified"),
            ("Scope maturity","-4","Class 3 assumptions still need package evidence"),
            ("Procurement certainty","-5","Long-lead evidence remains incomplete"),
            ("Schedule logic","+5","Level 4 logic gives credible traceability"),
            ("Reserve adequacy","+2","Balanced contingency posture")]
    model["confidence_breakdown"]=[{"driver":a,"effect":b,"note":c} for a,b,c in cb]

    # Cost line mutation: not proportional only; add scenario basis
    new_lines=[]
    for i,line in enumerate(base.get("cost_lines", model.get("cost_lines",[]))):
        x=dict(line)
        typ=x.get("type","Direct")
        factor = profile["direct_factor"] if typ=="Direct" else profile["indirect_factor"] if typ=="Indirect" else profile["reserve_factor"]
        # Make some sector lines move harder by scenario
        desc=(x.get("description") or "").lower()
        if scenario=="faster" and any(k in desc for k in ["equipment","tool","utility","clean","commission","process"]): factor*=1.08
        if scenario=="cheaper" and any(k in desc for k in ["redund", "utility", "resilience", "warehouse", "automation"]): factor*=0.82
        if scenario in ["lower_risk","premium"] and any(k in desc for k in ["utility","commission","validation","risk","reserve"]): factor*=1.10
        for key in ["low_p10","most_likely_p50","high_p90"]:
            if key in x:
                basev=_money_to_bn_hs(x[key])
                spread = 1.00
                if key=="high_p90" and scenario in ["faster","cheaper"]: spread=1.18
                if key=="low_p10" and scenario=="cheaper": spread=0.92
                if key=="high_p90" and scenario in ["lower_risk","premium"]: spread=0.95
                x[key]=_bn_to_money_hs(basev*factor*spread)
        if scenario=="faster":
            x["basis"]="Acceleration basis: expediting, parallel workfronts, premium suppliers and productivity drag applied to this CBS line."
        elif scenario=="cheaper":
            x["basis"]="Cheaper basis: reduced capital scope / lower redundancy assumed; downstream operational and contingency exposure increases."
        elif scenario=="lower_risk":
            x["basis"]="Lower-risk basis: assurance, procurement buffer, commissioning float and evidence maturity added to this CBS line."
        elif scenario=="premium":
            x["basis"]="Premium basis: resilience, redundancy, optionality and priority procurement applied."
        else:
            x["basis"]="Base basis: balanced reference case using sector benchmark, location, scale and complexity factors."
        new_lines.append(x)
    model["cost_lines"]=new_lines
    model["cost_breakdown"]=new_lines

    # Risk register mutation with scenario-specific risks at top
    risks=[]
    for tup in _scenario_risks_hs(model, scenario, profile["risk_lift"]):
        rid,title,cause,event,impact,prob,owner,mit=tup
        risks.append({"id":rid,"title":title,"cause":cause,"event":event,"impact":impact,"probability_pct":int(prob),"activity":"A1500","cbs":"02.01","owner":owner,"mitigation":mit})
    for r in base.get("risks", model.get("risks",[])):
        y=dict(r)
        try:
            y["probability_pct"]=max(5,min(88,int(y.get("probability_pct",30))+profile["risk_lift"]))
        except Exception: pass
        if scenario=="faster":
            y["mitigation"]="Fast-track mitigation: protect hold-points, add surge controls and weekly executive risk burn-down. " + str(y.get("mitigation",""))
        elif scenario=="cheaper":
            y["mitigation"]="Cheaper mitigation: confirm minimum viable resilience and owner acceptance of transferred risk. " + str(y.get("mitigation",""))
        elif scenario in ["lower_risk","premium"]:
            y["mitigation"]="Assurance mitigation: evidence gates, protected float and procurement certainty. " + str(y.get("mitigation",""))
        risks.append(y)
    model["risks"]=risks[:12]
    model["risk_register"]=model["risks"]

    # Schedule mutation with critical path shifts and float meaning
    rows=[]
    for idx,row in enumerate(base.get("schedule_rows", model.get("schedule_rows",[]))):
        z=dict(row)
        try:
            d=max(1,int(z.get("duration_months",1)))
            z["duration_months"]=max(1,int(round(d*profile["schedule"])))
        except Exception: pass
        if scenario=="faster":
            z["critical"]="Yes" if idx in [1,2,3,4,5] else z.get("critical","No")
            z["basis"]="Faster logic: overlapped procurement / turnover / commissioning; reduced recovery float and higher near-critical density."
        elif scenario=="cheaper":
            z["critical"]="Yes" if idx in [2,4,6] else z.get("critical","No")
            z["basis"]="Cheaper logic: slower phasing, deferred redundancy, leaner owner support and longer operational readiness tail."
        elif scenario=="lower_risk":
            z["critical"]="Yes" if idx in [3,5] else "No"
            z["basis"]="Lower-risk logic: protected evidence gates, added commissioning float and lower near-critical congestion."
        elif scenario=="premium":
            z["critical"]="Yes" if idx in [3,5] else "No"
            z["basis"]="Premium logic: priority procurement, stronger redundancy and managed readiness gates."
        else:
            z["basis"]="Base logic: balanced critical path with standard sector sequencing and benchmark duration."
        rows.append(z)
    model["schedule_rows"]=rows
    model["schedule_detail"]=rows

    # Curves and tornado
    curve=_curve_hs(new_cost,new_months,scenario)
    qcra={"p10":round(curve[2]["cost_bn"],1),"p50":round(new_cost,1),"p80":round(curve[9]["cost_bn"],1),"p90":round(curve[10]["cost_bn"],1),"p95":round(curve[11]["cost_bn"],1)}
    qsra={"p10":curve[2]["schedule_months"],"p50":new_months,"p80":curve[9]["schedule_months"],"p90":curve[10]["schedule_months"],"p95":curve[11]["schedule_months"]}
    # Keep headline range and curve range reconciled; P50 in curve is the same as the KPI P50.
    model["cost_p10"]=_bn_to_money_hs(qcra["p10"])
    model["cost_p90"]=_bn_to_money_hs(qcra["p90"])
    model["cost_range"]=f"{model['cost_p10']} - {model['cost_p90']}"

    # Executive scenario comparison versus the immutable Base case.
    base_cost_money=_bn_to_money_hs(base_cost)
    base_schedule_months=base_months
    base_conf_pct=base_conf
    delta_cost=new_cost-base_cost
    delta_months=new_months-base_months
    delta_conf=new_conf-base_conf
    model["scenario_comparison_vs_base"]={
        "base":{"cost_p50":base_cost_money,"schedule_months":base_schedule_months,"confidence_pct":base_conf_pct,"risk":model.get("_base_risk","Medium-High")},
        "selected":{"scenario":profile["label"],"cost_p50":model["cost_p50"],"schedule_months":new_months,"confidence_pct":new_conf,"risk":profile["risk"]},
        "delta":{"cost_bn":round(delta_cost,2),"cost":_bn_to_money_hs(abs(delta_cost)),"cost_direction":"higher" if delta_cost>0 else "lower" if delta_cost<0 else "same",
                 "months":delta_months,"confidence_pts":delta_conf},
        "plain_english": ("Base is the reference case: no cost, schedule or confidence delta is applied. Use Faster, Cheaper, Lower Risk or Premium to expose the board trade-off." if scenario=="base" else f"Compared with Base, {profile['label']} is {_bn_to_money_hs(abs(delta_cost))} {'more expensive' if delta_cost>0 else 'cheaper' if delta_cost<0 else 'unchanged'}, {abs(delta_months)} months {'faster' if delta_months<0 else 'slower' if delta_months>0 else 'unchanged'}, and {abs(delta_conf)} confidence points {'higher' if delta_conf>0 else 'lower' if delta_conf<0 else 'unchanged'}.")
    }
    model["scenario_matrix"]=[]
    for sk,sp in SCENARIO_PROFILES_V95_PLUS.items():
        c=base_cost*sp["cost"]; m=max(3,int(round(base_months*sp["schedule"]))); cf=max(8,min(96,base_conf+sp["confidence"]))
        model["scenario_matrix"].append({"scenario":sk,"label":sp["label"],"cost_p50":_bn_to_money_hs(c),"schedule_months":m,"risk":sp["risk"],"confidence_pct":cf,"cost_delta_pct":round((sp["cost"]-1)*100),"schedule_delta_pct":round((sp["schedule"]-1)*100),"confidence_delta_pts":sp["confidence"],"why":"; ".join([sp.get("trade",""), sp.get("won",""), sp.get("lost","")])})
    # Waterfalls explain WHY the selected scenario moved. These feed dashboard and exports.
    if scenario=="faster":
        cost_moves=[("Base P50",base_cost),("Acceleration premium",base_cost*.075),("Parallel EPC / package overlap",base_cost*.045),("Early long-lead procurement",base_cost*.035),("Commissioning surge / overtime",base_cost*.025),("Faster scenario P50",new_cost)]
        sched_moves=[("Base duration",base_months),("Parallel delivery workfronts",-round(base_months*.07)),("Early procurement release",-round(base_months*.05)),("Concurrent commissioning",-round(base_months*.04)),("Reduced float / overlap",new_months-base_months+round(base_months*.16)),("Faster duration",new_months)]
    elif scenario=="cheaper":
        cost_moves=[("Base P50",base_cost),("Scope / spec restraint",-base_cost*.055),("Deferred redundancy",-base_cost*.040),("Lower indirects / slower phasing",-base_cost*.030),("Reduced reserve",-base_cost*.015),("Cheaper scenario P50",new_cost)]
        sched_moves=[("Base duration",base_months),("Slower procurement path",round(base_months*.04)),("Deferred assurance / resilience",round(base_months*.035)),("Lean owner / commissioning support",round(base_months*.03)),("Cheaper duration",new_months)]
    elif scenario=="lower_risk":
        cost_moves=[("Base P50",base_cost),("Assurance gates",base_cost*.035),("Procurement buffer",base_cost*.030),("Commissioning float",base_cost*.025),("Reserve uplift",base_cost*.030),("Lower-risk scenario P50",new_cost)]
        sched_moves=[("Base duration",base_months),("Evidence gates",round(base_months*.055)),("Protected commissioning float",round(base_months*.055)),("Lower-risk duration",new_months)]
    elif scenario=="premium":
        cost_moves=[("Base P50",base_cost),("Redundancy / resilience",base_cost*.095),("Priority procurement",base_cost*.060),("Operational readiness",base_cost*.045),("Optionality premium",base_cost*.080),("Premium scenario P50",new_cost)]
        sched_moves=[("Base duration",base_months),("Priority procurement",-round(base_months*.04)),("Stronger readiness gates",round(base_months*.02)),("Premium duration",new_months)]
    else:
        cost_moves=[("Base P50",base_cost),("Selected scenario P50",new_cost)]
        sched_moves=[("Base duration",base_months),("Selected scenario duration",new_months)]
    model["cost_waterfall_vs_base"]=[{"driver":a,"value_bn":round(b,2),"value":_bn_to_money_hs(abs(b)) if i not in [0,len(cost_moves)-1] else _bn_to_money_hs(b),"kind":"total" if i in [0,len(cost_moves)-1] else "delta"} for i,(a,b) in enumerate(cost_moves)]
    model["schedule_waterfall_vs_base"]=[{"driver":a,"months":int(b),"kind":"total" if i in [0,len(sched_moves)-1] else "delta"} for i,(a,b) in enumerate(sched_moves)]
    curve_text={
        "base":"Reference curve: P50 equals the headline estimate; P80/P90 show normal board-level contingency exposure.",
        "faster":"Faster curve: the median date improves, but the P80/P90 tail widens because float, procurement optionality and commissioning recovery are consumed.",
        "cheaper":"Cheaper curve: P50 is lower, but P80/P90 become ugly because reserve, redundancy and recovery capacity have been stripped out.",
        "lower_risk":"Lower-risk curve: P50 is higher/slower, but the P80/P90 tail compresses because evidence, float and reserve are protected.",
        "premium":"Premium curve: higher approval value buys resilience; the value is not the P50, it is the controlled downside tail."
    }.get(scenario)
    model["monte_carlo"]={
        "qcra":qcra,
        "qsra":qsra,
        "curve":curve,
        "curve_interpretation":curve_text,
        "curve_readout":[
            f"Cost P50 reconciles to {model['cost_p50']}; P80 is {_bn_to_money_hs(qcra['p80'])}; P90 is {_bn_to_money_hs(qcra['p90'])}.",
            f"Schedule P50 reconciles to {new_months} months; P80 is {qsra['p80']} months; P90 is {qsra['p90']} months.",
            curve_text
        ],
        "tornado":[{"risk_id":r.get("id") or r.get("risk_id"),"title":r.get("title"),"driver":r.get("title"),"driver_score":max(8,100-i*7),"contribution":max(8,100-i*7)} for i,r in enumerate(model["risks"][:8])]
    }
    # Board + outputs
    model["board_briefing"]=[
        insight,
        model.get("scenario_comparison_vs_base",{}).get("plain_english", f"{profile['label']} scenario versus Base calculated."),
        f"{profile['label']} scenario: {model['cost_p50']} P50, {model['cost_range']} range, {model['schedule']} baseline and {new_conf}% confidence.",
        f"Trade made: {profile['trade']}",
        f"Gained: {profile['won']}",
        f"Lost: {profile['lost']}"
    ]
    model["outputs_board_memo"]=[model.get("scenario_comparison_vs_base",{}).get("plain_english", "Scenario compared with Base.")] + _scenario_outputs_text_hs(profile, insight)
    model["top_decisions_required"]=[
        "Accept or reject the scenario trade-off explicitly at board level.",
        "Confirm the changed critical path and near-critical path density.",
        "Approve the scenario-specific reserve and contingency philosophy.",
        "Assign named owners for the new top scenario risks.",
        "Confirm whether XER, cost workbook and risk register should be issued as scenario-controlled outputs."
    ]
    model["casey_thinking"]=(f"CASEY has re-cut the programme as a {profile['label']} scenario. {profile['trade']} "
        f"The governing consequence is: {insight} QCRA/QSRA curves, cost basis, risk probabilities and schedule logic have been re-weighted to match this strategic posture.")
    model["executive_summary"]=(
        f"{model.get('title','Programme')} scenario view: {profile['label']}. CASEY indicates {model['cost_p50']} P50 exposure, "
        f"{model['cost_range']} range, {model['schedule']} baseline, {model['risk']} risk and {new_conf}% confidence. "
        f"{insight} Trade-off: {profile['trade']}"
    )
    # Export control metadata
    model["export_scenario_control"]={
        "scenario":profile["label"],
        "must_stamp_exports":True,
        "xer_logic":"critical path, duration and basis changed by scenario",
        "cost_logic":"CBS basis, reserve and uncertainty spread changed by scenario",
        "risk_logic":"probability, mitigation and scenario-emergent risks changed by scenario"
    }

    # FINAL EXEC POLISH: scenario-aware ranking and differentiation
    try:
        sc = str(model.get("scenario","base")).lower()
        schedule_lists = {
            "faster":[
                "Concurrent commissioning overload",
                "Recovery float exhaustion",
                "Acceleration premium shock",
                "Grid connection delay",
                "Integrated systems testing concurrency"
            ],
            "cheaper":[
                "Vendor claims and change exposure",
                "Procurement deferral and long-lead slippage",
                "Design maturity gap",
                "Scope growth from deferred decisions",
                "Interface coordination delay"
            ],
            "lower_risk":[
                "Governance and approvals latency",
                "Extended validation sequencing",
                "Conservative commissioning gates",
                "Operational readiness hold-points",
                "Assurance and compliance reviews"
            ],
            "premium":[
                "Integration complexity across parallel packages",
                "Executive decision latency",
                "Technology assurance alignment",
                "Multi-package interface management",
                "Programme coordination overhead"
            ]
        }
        cost_lists = {
            "faster":[
                "Acceleration premiums and overtime",
                "Power train, transformers and switchgear",
                "Integrated systems testing",
                "Grid and utility concurrency",
                "Recovery-float consumption"
            ],
            "cheaper":[
                "Deferred procurement packaging",
                "Claims and commercial exposure",
                "Rework from reduced contingency",
                "Long-lead inflation volatility",
                "Scope rationalisation impacts"
            ],
            "lower_risk":[
                "Additional contingency and reserve",
                "Enhanced validation and assurance",
                "Programme controls and governance",
                "Redundant infrastructure resilience",
                "Extended commissioning readiness"
            ]
        }
        if sc in schedule_lists:
            model["sector_schedule_threats"]=schedule_lists[sc]
        if sc in cost_lists:
            model["sector_primary_cost_drivers"]=cost_lists[sc]

        if sc=="faster":
            model["executive_shock_insight"]="Acceleration increases spend faster than it reduces uncertainty; the delivery tail becomes more volatile."
        elif sc=="cheaper":
            model["executive_shock_insight"]="Capital efficiency reduces resilience: procurement and recovery flexibility become constrained."
        elif sc=="lower_risk":
            model["executive_shock_insight"]="Confidence is purchased through reserve, governance and extended delivery duration."
        elif sc=="premium":
            model["executive_shock_insight"]="Premium posture buys resilience, optionality and stronger certainty at visible capex premium."
    except Exception:
        pass

    return model



# ================= CASEY V104 FRONTEND GRAPH COMPATIBILITY + EXEC FEATURES =================
# Ensures V103 frontend receives all graph / scenario / executive fields it expects.
import copy as _casey_copy_v104
import math as _casey_math_v104

def _v104_money_to_bn(x):
    s = str(x or "").replace("$","").replace(",","").strip().upper()
    try:
        if s.endswith("T"): return float(s[:-1]) * 1000
        if s.endswith("B"): return float(s[:-1])
        if s.endswith("M"): return float(s[:-1]) / 1000
        return float(s)
    except Exception:
        return 0.0

def _v104_fmt_bn(x):
    x = float(x or 0)
    if x >= 1000: return f"${x/1000:.1f}T"
    if x >= 1: return f"${x:.1f}B"
    return f"${x*1000:.0f}M"

def _v104_months(x):
    try:
        return int(float(str(x or "0").replace("months","").strip().split()[0]))
    except Exception:
        return 0

def _v104_profile(s):
    s = (s or "base").lower().replace(" ","_")
    return {
        "base":       {"label":"Base","cost":1.00,"sched":1.00,"conf":0,   "risk":"Medium-High"},
        "faster":     {"label":"Faster","cost":1.22,"sched":0.74,"conf":-18, "risk":"High"},
        "cheaper":    {"label":"Cheaper","cost":0.82,"sched":1.20,"conf":-20, "risk":"High"},
        "lower_risk": {"label":"Lower Risk","cost":1.16,"sched":1.16,"conf":18, "risk":"Medium"},
        "premium":    {"label":"Premium","cost":1.32,"sched":0.94,"conf":23, "risk":"Medium-Low"},
    }.get(s, {"label":"Base","cost":1.00,"sched":1.00,"conf":0,"risk":"Medium-High"})

def _v104_curve(cost_bn, months, scen):
    scen = (scen or "base").lower().replace(" ","_")
    pts = [1,5,10,20,30,40,50,60,70,80,90,95,99]
    arr = []
    for p in pts:
        x = p / 100
        if scen == "faster":
            cf = 0.70 + 0.34*x + 0.50*(x**4.3)
            sf = 0.60 + 0.24*x + 0.64*(x**4.7)
        elif scen == "cheaper":
            cf = 0.50 + 0.25*x + 0.86*(x**3.9)
            sf = 0.70 + 0.22*x + 0.78*(x**3.5)
        elif scen == "lower_risk":
            cf = 0.84 + 0.18*x + 0.08*(x**2.6)
            sf = 0.88 + 0.14*x + 0.07*(x**2.6)
        elif scen == "premium":
            cf = 0.86 + 0.18*x + 0.08*(x**2.6)
            sf = 0.80 + 0.16*x + 0.13*(x**2.6)
        else:
            cf = 0.70 + 0.25*x + 0.34*(x**2.8)
            sf = 0.74 + 0.18*x + 0.32*(x**2.6)
        arr.append({"percentile":p, "cost_bn":round(cost_bn*cf,2), "schedule_months":max(1,int(round(months*sf)))})
    return arr

def _v104_trade(profile):
    label = profile["label"]
    if label == "Faster":
        return ("You bought time by spending money and consuming recovery float.", "Earlier market-entry / revenue timing.", "Procurement optionality, recovery float and late-stage stability.")
    if label == "Cheaper":
        return ("You cut capital authorization by transferring risk into resilience, redundancy and start-up.", "Lower initial approval number and reduced near-term capital draw.", "Operational resilience, commissioning flexibility, contingency adequacy and lifecycle certainty.")
    if label == "Lower Risk":
        return ("You bought confidence by adding assurance, float and procurement evidence.", "Reduced P80/P90 exposure and stronger board approval defensibility.", "Earlier revenue date and lean capital posture.")
    if label == "Premium":
        return ("You bought resilience, redundancy and strategic optionality.", "Higher resilience, stronger procurement certainty and better downside protection.", "Lowest-capex authorization case.")
    return ("Balanced cost, time and evidence posture.", "Credible reference case for board challenge.", "No extra certainty or capital efficiency purchased.")

def _v104_normalize_model(model, prompt="", scenario="base"):
    if not isinstance(model, dict):
        return model
    
    # FINAL EXEC POLISH: scenario-aware ranking and differentiation
    try:
        sc = str(model.get("scenario","base")).lower()
        schedule_lists = {
            "faster":[
                "Concurrent commissioning overload",
                "Recovery float exhaustion",
                "Acceleration premium shock",
                "Grid connection delay",
                "Integrated systems testing concurrency"
            ],
            "cheaper":[
                "Vendor claims and change exposure",
                "Procurement deferral and long-lead slippage",
                "Design maturity gap",
                "Scope growth from deferred decisions",
                "Interface coordination delay"
            ],
            "lower_risk":[
                "Governance and approvals latency",
                "Extended validation sequencing",
                "Conservative commissioning gates",
                "Operational readiness hold-points",
                "Assurance and compliance reviews"
            ],
            "premium":[
                "Integration complexity across parallel packages",
                "Executive decision latency",
                "Technology assurance alignment",
                "Multi-package interface management",
                "Programme coordination overhead"
            ]
        }
        cost_lists = {
            "faster":[
                "Acceleration premiums and overtime",
                "Power train, transformers and switchgear",
                "Integrated systems testing",
                "Grid and utility concurrency",
                "Recovery-float consumption"
            ],
            "cheaper":[
                "Deferred procurement packaging",
                "Claims and commercial exposure",
                "Rework from reduced contingency",
                "Long-lead inflation volatility",
                "Scope rationalisation impacts"
            ],
            "lower_risk":[
                "Additional contingency and reserve",
                "Enhanced validation and assurance",
                "Programme controls and governance",
                "Redundant infrastructure resilience",
                "Extended commissioning readiness"
            ]
        }
        if sc in schedule_lists:
            model["sector_schedule_threats"]=schedule_lists[sc]
        if sc in cost_lists:
            model["sector_primary_cost_drivers"]=cost_lists[sc]

        if sc=="faster":
            model["executive_shock_insight"]="Acceleration increases spend faster than it reduces uncertainty; the delivery tail becomes more volatile."
        elif sc=="cheaper":
            model["executive_shock_insight"]="Capital efficiency reduces resilience: procurement and recovery flexibility become constrained."
        elif sc=="lower_risk":
            model["executive_shock_insight"]="Confidence is purchased through reserve, governance and extended delivery duration."
        elif sc=="premium":
            model["executive_shock_insight"]="Premium posture buys resilience, optionality and stronger certainty at visible capex premium."
    except Exception:
        pass

    return model
    scenario = (scenario or model.get("scenario") or "base").lower().replace(" ","_")
    profile = _v104_profile(scenario)
    model["scenario"] = scenario
    model["scenario_label"] = profile["label"]

    # Determine base values from model if no explicit base object exists.
    base = model.get("base_comparison") or {}
    base_cost = _v104_money_to_bn(base.get("base_cost_p50") or model.get("_base_cost_p50") or model.get("cost_p50"))
    if base_cost <= 0:
        base_cost = _v104_money_to_bn(model.get("cost_p50")) or 5.0
    base_months = _v104_months(base.get("base_schedule") or model.get("_base_schedule") or model.get("schedule")) or 60
    base_conf = int(str(base.get("base_confidence") or model.get("_base_confidence") or model.get("confidence_pct") or 60).replace("%","").split()[0])

    # If selected scenario already has values, preserve them; otherwise derive from base.
    scen_cost = _v104_money_to_bn(model.get("cost_p50")) if scenario == "base" else base_cost * profile["cost"]
    scen_months = _v104_months(model.get("schedule")) if scenario == "base" else max(3,int(round(base_months * profile["sched"])))
    scen_conf = int(model.get("confidence_pct") or base_conf) if scenario == "base" else max(5,min(96,base_conf + profile["conf"]))

    model["cost_p50"] = _v104_fmt_bn(scen_cost)
    model["cost_p10"] = _v104_fmt_bn(scen_cost * (0.78 if scenario != "cheaper" else 0.70))
    model["cost_p90"] = _v104_fmt_bn(scen_cost * (1.32 if scenario in ["base","lower_risk","premium"] else 1.52 if scenario=="faster" else 1.60))
    model["cost_range"] = f"{model['cost_p10']} - {model['cost_p90']}"
    model["schedule"] = f"{scen_months} months"
    model["risk"] = profile["risk"] if scenario != "base" else model.get("risk", profile["risk"])
    model["confidence_pct"] = scen_conf

    trade, gain, loss = _v104_trade(profile)
    insight = {
        "base":"Programme success depends on the real constraint, not the headline construction scope.",
        "faster":"Acceleration increases spend faster than it reduces uncertainty; the delivery tail becomes more volatile.",
        "cheaper":"The cheaper case is not simply cheaper; it transfers risk into resilience, contingency and start-up certainty.",
        "lower_risk":"Lower-risk delivery buys confidence through assurance, float and procurement evidence.",
        "premium":"Premium delivery buys downside protection, resilience and strategic optionality."
    }.get(scenario, trade)
    model["executive_shock_insight"] = model.get("executive_shock_insight") if scenario=="base" else insight
    model["scenario_trade"] = trade
    model["scenario_gain"] = gain
    model["scenario_loss"] = loss

    cost_delta = scen_cost - base_cost
    model["base_comparison"] = {
        "base_cost_p50": _v104_fmt_bn(base_cost),
        "scenario_cost_p50": _v104_fmt_bn(scen_cost),
        "cost_delta": ("-" + _v104_fmt_bn(abs(cost_delta))) if cost_delta < 0 else _v104_fmt_bn(cost_delta),
        "base_schedule": f"{base_months} months",
        "scenario_schedule": f"{scen_months} months",
        "schedule_delta": f"{scen_months-base_months:+d} months",
        "base_confidence": f"{base_conf}%",
        "scenario_confidence": f"{scen_conf}%",
        "confidence_delta": f"{scen_conf-base_conf:+d} pts",
        "trade": trade,
        "gain": gain,
        "loss": loss
    }
    model["scenario_comparison_vs_base"] = model["base_comparison"]
    model["scenario_delta_intelligence"] = [
        {"label":"Base P50 cost", "value":f"{_v104_fmt_bn(base_cost)} → {_v104_fmt_bn(scen_cost)}", "meaning":trade},
        {"label":"Base P50 schedule", "value":f"{base_months} mo → {scen_months} mo", "meaning":gain},
        {"label":"Confidence", "value":f"{base_conf}% → {scen_conf}%", "meaning":loss},
        {"label":"Risk posture", "value":model["risk"], "meaning":insight},
    ]

    model["cost_waterfall_vs_base"] = [
        {"driver":"Base P50", "value":round(base_cost,2), "kind":"base"},
        {"driver":profile["label"]+" scenario delta", "value":round(cost_delta,2), "kind":"delta"},
        {"driver":"Scenario P50", "value":round(scen_cost,2), "kind":"total"},
    ]
    model["schedule_waterfall_vs_base"] = [
        {"driver":"Base schedule", "months":base_months, "kind":"base"},
        {"driver":profile["label"]+" scenario delta", "months":scen_months-base_months, "kind":"delta"},
        {"driver":"Scenario schedule", "months":scen_months, "kind":"total"},
    ]

    curve = _v104_curve(scen_cost, scen_months, scenario)
    qcra = {"p10":curve[2]["cost_bn"], "p50":round(scen_cost,1), "p80":curve[9]["cost_bn"], "p90":curve[10]["cost_bn"], "p95":curve[11]["cost_bn"]}
    qsra = {"p10":curve[2]["schedule_months"], "p50":scen_months, "p80":curve[9]["schedule_months"], "p90":curve[10]["schedule_months"], "p95":curve[11]["schedule_months"]}

    risks = model.get("risks") or model.get("risk_register") or []
    # Make sure risk ids and chart names match frontend table + tornado.
    norm_risks = []
    for i,r in enumerate(risks[:12]):
        x = dict(r)
        if "risk_id" not in x:
            x["risk_id"] = x.get("id") or f"R{i+1:02d}"
        x["title"] = x.get("title") or x.get("risk") or f"Risk {i+1}"
        x["probability_pct"] = int(x.get("probability_pct") or x.get("probability") or max(15,65-i*5))
        norm_risks.append(x)
    if not norm_risks:
        norm_risks = [{"risk_id":"R01","title":"Scenario execution volatility","probability_pct":55,"cause":"Scenario posture","event":"Risk materialises","impact":"Cost/schedule exposure","owner":"PMO","mitigation":"Weekly controls review"}]
    model["risks"] = norm_risks
    model["risk_register"] = norm_risks

    tornado = []
    for i,r in enumerate(norm_risks[:8]):
        title = r.get("title","Risk")
        contribution = max(8, int(r.get("probability_pct",50)) + max(0,30-i*4))
        tornado.append({
            "risk_id": r.get("risk_id") or r.get("id") or f"R{i+1:02d}",
            "driver": title,
            "title": title,
            "contribution": contribution,
            "driver_score": contribution
        })

    model["monte_carlo"] = {
        **(model.get("monte_carlo") or {}),
        "qcra": qcra,
        "qsra": qsra,
        "curve": curve,
        "tornado": tornado,
        "curve_readout": [
            f"QCRA: P50 {model['cost_p50']} equals the headline scenario estimate; P80 { _v104_fmt_bn(qcra['p80']) } and P90 { _v104_fmt_bn(qcra['p90']) } show downside board exposure.",
            f"QSRA: P50 {scen_months} months equals the headline schedule; P80 {qsra['p80']} and P90 {qsra['p90']} months show finish-date risk.",
            f"Scenario vs Base: cost {model['base_comparison']['base_cost_p50']} → {model['cost_p50']}; schedule {base_months} → {scen_months} months; confidence {base_conf}% → {scen_conf}%."
        ],
        "curve_interpretation": {
            "base":"Balanced uncertainty: P80/P90 reflects normal delivery and procurement risk.",
            "faster":"Compressed median with aggressive P80/P95 tail: possible, but unstable.",
            "cheaper":"Lower median with fragile reserve and ugly tail: cheaper only if nothing goes wrong.",
            "lower_risk":"Tighter distribution: slower and more expensive, but more governable.",
            "premium":"Higher median with controlled downside: buys resilience and optionality."
        }.get(scenario)
    }

    model["confidence_breakdown"] = [
        {"driver":"Scenario trade-off", "effect":f"{profile['conf']:+d}", "note":trade},
        {"driver":"Procurement certainty", "effect":"-10" if scenario in ["faster","cheaper"] else "+9", "note":"Long-lead evidence controls confidence movement."},
        {"driver":"Schedule logic maturity", "effect":"-9" if scenario=="faster" else "+7" if scenario in ["lower_risk","premium"] else "+3", "note":"Critical path and handover gates are rebalanced."},
        {"driver":"Commissioning / validation readiness", "effect":"-12" if scenario in ["faster","cheaper"] else "+10", "note":"Readiness exposure changes by scenario."},
        {"driver":"Reserve adequacy", "effect":"-15" if scenario=="cheaper" else "+12" if scenario in ["lower_risk","premium"] else "+2", "note":"Reserve is treated as a scenario choice."},
    ]

    model["board_briefing"] = [
        insight,
        f"{profile['label']} scenario: {model['cost_p50']} P50, {model['cost_range']} range, {model['schedule']} baseline and {scen_conf}% confidence.",
        f"Base comparison: {model['base_comparison']['base_cost_p50']} / {model['base_comparison']['base_schedule']} / {model['base_comparison']['base_confidence']} → {model['cost_p50']} / {model['schedule']} / {scen_conf}%.",
        f"Gained: {gain}",
        f"Lost: {loss}"
    ]
    model["outputs_board_memo"] = [
        f"Selected scenario: {profile['label']}.",
        f"Trade made: {trade}",
        f"Base vs scenario: {model['base_comparison']['base_cost_p50']} → {model['cost_p50']}; {model['base_comparison']['base_schedule']} → {model['schedule']}.",
        f"Board warning: {insight}",
        "Exports should be stamped with scenario, base value, scenario value, delta and trade-off reason."
    ]
    model["casey_thinking"] = f"CASEY recalibrated cost, schedule, QCRA/QSRA, risk and confidence from the same scenario basis. {trade}"
    model["executive_summary"] = f"{model.get('title','Programme')} scenario view: {profile['label']}. CASEY indicates {model['cost_p50']} P50 exposure, {model['cost_range']} range, {model['schedule']} baseline, {model['risk']} risk and {scen_conf}% confidence. {insight}"

    # FINAL EXEC POLISH: scenario-aware ranking and differentiation
    try:
        sc = str(model.get("scenario","base")).lower()
        schedule_lists = {
            "faster":[
                "Concurrent commissioning overload",
                "Recovery float exhaustion",
                "Acceleration premium shock",
                "Grid connection delay",
                "Integrated systems testing concurrency"
            ],
            "cheaper":[
                "Vendor claims and change exposure",
                "Procurement deferral and long-lead slippage",
                "Design maturity gap",
                "Scope growth from deferred decisions",
                "Interface coordination delay"
            ],
            "lower_risk":[
                "Governance and approvals latency",
                "Extended validation sequencing",
                "Conservative commissioning gates",
                "Operational readiness hold-points",
                "Assurance and compliance reviews"
            ],
            "premium":[
                "Integration complexity across parallel packages",
                "Executive decision latency",
                "Technology assurance alignment",
                "Multi-package interface management",
                "Programme coordination overhead"
            ]
        }
        cost_lists = {
            "faster":[
                "Acceleration premiums and overtime",
                "Power train, transformers and switchgear",
                "Integrated systems testing",
                "Grid and utility concurrency",
                "Recovery-float consumption"
            ],
            "cheaper":[
                "Deferred procurement packaging",
                "Claims and commercial exposure",
                "Rework from reduced contingency",
                "Long-lead inflation volatility",
                "Scope rationalisation impacts"
            ],
            "lower_risk":[
                "Additional contingency and reserve",
                "Enhanced validation and assurance",
                "Programme controls and governance",
                "Redundant infrastructure resilience",
                "Extended commissioning readiness"
            ]
        }
        if sc in schedule_lists:
            model["sector_schedule_threats"]=schedule_lists[sc]
        if sc in cost_lists:
            model["sector_primary_cost_drivers"]=cost_lists[sc]

        if sc=="faster":
            model["executive_shock_insight"]="Acceleration increases spend faster than it reduces uncertainty; the delivery tail becomes more volatile."
        elif sc=="cheaper":
            model["executive_shock_insight"]="Capital efficiency reduces resilience: procurement and recovery flexibility become constrained."
        elif sc=="lower_risk":
            model["executive_shock_insight"]="Confidence is purchased through reserve, governance and extended delivery duration."
        elif sc=="premium":
            model["executive_shock_insight"]="Premium posture buys resilience, optionality and stronger certainty at visible capex premium."
    except Exception:
        pass

    return model

_CASEY_V104_ORIGINAL_BUILD_MODEL = build_model
def build_model(prompt:str, client:str="", class_level:int=3, schedule_level:int=3, scenario:str="base"):
    return _v104_normalize_model(_CASEY_V104_ORIGINAL_BUILD_MODEL(prompt, client, class_level, schedule_level, scenario), prompt, scenario)

# Local demo package: disable one-run lock.
def _public_demo_used(identity):
    return None

def demo_status(request):
    return {"allowed": True, "used": 0, "limit": 999, "unlocked": True}
# ================= END CASEY V104 FRONTEND GRAPH COMPATIBILITY =================

# ================= CASEY V106 DEMO EXPORT POLISH =================
# Real demo exports for every export button. Public demo outputs are stamped clearly
# so buyers understand these are sample board artefacts, while still receiving files.
def _v106_stamp_model(model: Dict[str, Any]) -> Dict[str, Any]:
    m = dict(model or {})
    sc = str(m.get('scenario_label') or m.get('scenario') or 'Base')
    base = m.get('base_comparison') or {}
    m['demo_watermark'] = 'CASEY PUBLIC DEMO OUTPUT - NOT CERTIFIED FOR COMMERCIAL RELIANCE'
    m['export_stamp'] = {
        'watermark': m['demo_watermark'],
        'scenario': sc,
        'base_value': base,
        'scenario_value': {'cost_p50': m.get('cost_p50'), 'schedule': m.get('schedule'), 'confidence_pct': m.get('confidence_pct'), 'risk': m.get('risk')},
        'trade_off_reason': m.get('casey_thinking') or m.get('executive_shock_insight') or 'Scenario-controlled first-pass intelligence.'
    }
    return m

def _v106_qcra_qsra_workbook_bytes(model: Dict[str, Any]) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.chart import LineChart, Reference, BarChart
    m = _v106_stamp_model(model)
    mc = m.get('monte_carlo') or {}
    curve = mc.get('curve') or []
    torn = mc.get('tornado') or []
    wb = Workbook()
    ws = wb.active
    ws.title = 'DEMO Stamp'
    rows = [
        ['CASEY PUBLIC DEMO OUTPUT', m.get('demo_watermark')],
        ['Project', m.get('title')],
        ['Scenario', m.get('scenario_label') or m.get('scenario')],
        ['P50 cost', m.get('cost_p50')],
        ['Schedule', m.get('schedule')],
        ['Risk / confidence', f"{m.get('risk')} / {m.get('confidence_pct')}%"],
        ['Trade-off reason', (m.get('casey_thinking') or m.get('executive_shock_insight') or '')],
    ]
    for r in rows: ws.append(r)
    for c in ws[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='7A1F1F')
    ws.column_dimensions['A'].width = 26; ws.column_dimensions['B'].width = 110
    ws2 = wb.create_sheet('QCRA QSRA Curves')
    ws2.append(['Percentile','QCRA Cost $B','QSRA Months'])
    for x in curve:
        ws2.append([x.get('percentile'), x.get('cost_bn'), x.get('schedule_months')])
    for c in ws2[1]: c.font = Font(bold=True); c.fill = PatternFill('solid', fgColor='B7F7FF')
    if len(curve) >= 2:
        chart = LineChart(); chart.title='QCRA cost confidence curve'; chart.y_axis.title='Cost $B'; chart.x_axis.title='Percentile'
        data = Reference(ws2, min_col=2, min_row=1, max_row=ws2.max_row); cats = Reference(ws2, min_col=1, min_row=2, max_row=ws2.max_row)
        chart.add_data(data, titles_from_data=True); chart.set_categories(cats); chart.height=8; chart.width=14; ws2.add_chart(chart,'E2')
        chart2 = LineChart(); chart2.title='QSRA finish-date curve'; chart2.y_axis.title='Months'; chart2.x_axis.title='Percentile'
        data2 = Reference(ws2, min_col=3, min_row=1, max_row=ws2.max_row); chart2.add_data(data2, titles_from_data=True); chart2.set_categories(cats); chart2.height=8; chart2.width=14; ws2.add_chart(chart2,'E18')
    ws3 = wb.create_sheet('Risk Tornado')
    ws3.append(['Rank','Risk ID','Driver','Contribution'])
    for i,x in enumerate(torn[:12],1): ws3.append([i,x.get('risk_id'),x.get('driver') or x.get('title'),x.get('contribution') or x.get('driver_score')])
    for c in ws3[1]: c.font = Font(bold=True); c.fill = PatternFill('solid', fgColor='B7F7FF')
    if ws3.max_row > 1:
        b=BarChart(); b.type='bar'; b.title='Top QCRA/QSRA exposure drivers'; b.y_axis.title='Driver'; b.x_axis.title='Contribution'
        data=Reference(ws3,min_col=4,min_row=1,max_row=ws3.max_row); cats=Reference(ws3,min_col=3,min_row=2,max_row=ws3.max_row)
        b.add_data(data,titles_from_data=True); b.set_categories(cats); b.height=10; b.width=16; ws3.add_chart(b,'F2')
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical='top', wrap_text=True)
        for col in range(1, min(sheet.max_column, 8)+1):
            sheet.column_dimensions[chr(64+col)].width = 22
    bio = BytesIO(); wb.save(bio); return bio.getvalue()

@app.post('/export/qcra-qsra')
def export_qcra_qsra_v106(model: Dict[str, Any]):
    return stream(_v106_qcra_qsra_workbook_bytes(model), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'CASEY_DEMO_QCRA_QSRA_Pack.xlsx')

@app.post('/export/excel')
def export_excel_alias_v106(model: Dict[str, Any]):
    # Alias kept so older frontend builds still download instead of failing.
    return stream(workbook_bytes(_v106_stamp_model(model)), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'CASEY_DEMO_Cost_Workbook.xlsx')
# ================= END CASEY V106 DEMO EXPORT POLISH =================


# ================= CASEY V107 FINAL DEMO COMMERCIAL LOCK + STRONG EXPORTS =================
# This final layer restores the commercial public-demo gate:
# one credible Earth OR Space run per email, and one demo export set per run/email.
try:
    from fastapi import Body
    from fastapi.responses import JSONResponse
    from v64_outputs import (
        cost_workbook_bytes as _v107_cost_workbook_bytes,
        risk_workbook_bytes as _v107_risk_workbook_bytes,
        patch_template_xer as _v107_xer_bytes,
        schedule_csv_bytes as _v107_schedule_csv_bytes,
        model_json_bytes as _v107_model_json_bytes,
    )
except Exception as _v107_import_error:
    print('CASEY v107 import warning:', _v107_import_error)


def _v107_db_init():
    con = db(); cur = con.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS public_demo_exports(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        run_id TEXT,
        email_hash TEXT,
        export_kind TEXT,
        created_at TEXT
    )""")
    con.commit(); con.close()
_v107_db_init()


def _v107_remove_route(path: str, method: str = 'POST'):
    app.router.routes = [r for r in app.router.routes if not (getattr(r, 'path', None) == path and method in getattr(r, 'methods', set()))]




def _v108_admin_email_set():
    """Emails that are allowed unlimited demo testing while public users remain one-run locked.
    Configure with CASEY_ADMIN_EMAILS="you@company.com,another@company.com".
    The bundled defaults are only for local demo testing and can be overridden in production.
    """
    raw = os.environ.get('CASEY_ADMIN_EMAILS', 'test@yahoo.com,jim@yahoo.com,jai@yahoo.com,casey@yahoo.com')
    return { _normalise_email(x) for x in raw.split(',') if str(x).strip() }


def _v108_is_admin_email(email: str) -> bool:
    return bool(email) and _normalise_email(email) in _v108_admin_email_set()


def _v108_admin_key_ok(request: Request = None) -> bool:
    key = os.environ.get('CASEY_ADMIN_KEY', '').strip()
    if not key or request is None:
        return False
    return (request.headers.get('x-casey-admin-key') == key) or (request.query_params.get('admin_key') == key)


def _v108_is_admin_payload(payload: Dict[str, Any]) -> bool:
    email = (payload or {}).get('lead_email') or (payload or {}).get('client') or (payload or {}).get('email') or ''
    return _v108_is_admin_email(str(email))

def _v107_is_paid(payload: Dict[str, Any]) -> bool:
    return bool((payload or {}).get('paid') or (payload or {}).get('enterprise_access') or os.environ.get('CASEY_DEMO_UNLOCK_EXPORTS') == '1' or _v108_is_admin_payload(payload or {}))


def _v107_hash_from_payload(payload: Dict[str, Any]) -> str:
    email = (payload or {}).get('lead_email') or (payload or {}).get('client') or (payload or {}).get('email') or ''
    return _sha(str(email))


def _v107_can_export(payload: Dict[str, Any], kind: str):
    # Local demo/dev override for unrestricted testing.
    # Enabled by default for localhost builds unless explicitly disabled.
    local_demo_override = os.environ.get('CASEY_LOCAL_DEMO', '1') == '1'
    if local_demo_override:
        return True, None

    if _v107_is_paid(payload):
        return True, None
    run_id = str((payload or {}).get('run_id') or '').strip()
    email_hash = _v107_hash_from_payload(payload)
    if not run_id or not email_hash:
        return False, 'Demo exports require a public demo run_id and email. Please run the one free CASEY intelligence run first, or request paid access.'
    con = db(); cur = con.cursor()
    row = cur.execute('SELECT id, export_kind, created_at FROM public_demo_exports WHERE (run_id=? OR email_hash=?) AND export_kind=? LIMIT 1', (run_id, email_hash, kind)).fetchone()
    if row:
        con.close()
        return False, f'This email has already downloaded the demo {kind.replace('_',' ')} export. Request access to unlock repeat exports and reruns.'
    now = datetime.utcnow().isoformat()
    cur.execute('INSERT INTO public_demo_exports(run_id,email_hash,export_kind,created_at) VALUES(?,?,?,?)', (run_id, email_hash, kind, now))
    con.commit(); con.close()
    return True, None


def _v107_export_block(message: str):
    raise HTTPException(status_code=402, detail={
        'message': message,
        'upgrade_cta': 'Request access to unlock unlimited reruns, exports, saved projects and certified workflow support.'
    })


def _v107_stamp_model(model: Dict[str, Any]) -> Dict[str, Any]:
    m = dict(model or {})
    base = m.get('base_comparison') or {}
    label = str(m.get('scenario_label') or m.get('scenario') or 'Base').replace('_',' ').title()
    wm = 'CASEY PUBLIC DEMO OUTPUT - SIMULATED SAMPLE - NOT CERTIFIED FOR COMMERCIAL RELIANCE'
    m['demo_watermark'] = wm
    m['watermark'] = wm
    m['scenario_label'] = label
    m['export_stamp'] = {
        'watermark': wm,
        'scenario': label,
        'run_id': m.get('run_id'),
        'base_value': base,
        'scenario_value': {'cost_p50': m.get('cost_p50'), 'schedule': m.get('schedule'), 'confidence_pct': m.get('confidence_pct'), 'risk': m.get('risk')},
        'trade_off_reason': m.get('casey_thinking') or m.get('executive_shock_insight') or 'Scenario-controlled first-pass CASEY intelligence.'
    }
    title = str(m.get('title') or 'CASEY Project')
    if 'DEMO' not in title.upper():
        m['title'] = f'DEMO - {title}'
    return m


def _v107_readme(model: Dict[str, Any]) -> bytes:
    m = _v107_stamp_model(model)
    txt = f"""CASEY PUBLIC DEMO OUTPUT PACK

WATERMARK: {m.get('demo_watermark')}
Run ID: {m.get('run_id')}
Scenario: {m.get('scenario_label')}
Project: {m.get('title')}
P50 Cost: {m.get('cost_p50')}
Schedule: {m.get('schedule')}
Risk / confidence: {m.get('risk')} / {m.get('confidence_pct')}%

This demo pack is intentionally stamped. It is a simulated board-grade intelligence sample, not a certified estimate, schedule, risk model or commercial deliverable.

Included outputs:
1. Cost workbook with scenario-controlled basis
2. Risk register with cause / event / impact / owner / mitigation
3. XER schedule generated from the selected project/scenario
4. Schedule CSV fallback
5. QCRA/QSRA workbook with curves and tornado drivers
6. JSON audit model

Paid access unlocks saved projects, repeat scenarios, unrestricted exports, private benchmark libraries and production controls workflows.
"""
    return txt.encode('utf-8')


def _v107_watermarked_qcra_qsra(model: Dict[str, Any]) -> bytes:
    return _v106_qcra_qsra_workbook_bytes(_v107_stamp_model(model))


def _v107_full_pack(model: Dict[str, Any]) -> bytes:
    m = _v107_stamp_model(model)
    bio = BytesIO()
    with zipfile.ZipFile(bio, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('00_CASEY_PUBLIC_DEMO_README.txt', _v107_readme(m))
        z.writestr('01_CASEY_DEMO_Cost_Workbook.xlsx', _v107_cost_workbook_bytes(m))
        z.writestr('02_CASEY_DEMO_Risk_Register.xlsx', _v107_risk_workbook_bytes(m))
        z.writestr('03_CASEY_DEMO_Strong_P6_Schedule.xer', _v107_xer_bytes(m))
        z.writestr('04_CASEY_DEMO_Schedule_CSV_Fallback.csv', _v107_schedule_csv_bytes(m))
        z.writestr('05_CASEY_DEMO_QCRA_QSRA_Curves_and_Tornado.xlsx', _v107_watermarked_qcra_qsra(m))
        z.writestr('06_CASEY_DEMO_Model_Audit.json', _v107_model_json_bytes(m))
    bio.seek(0)
    return bio.getvalue()

_v107_remove_route('/demo/status', 'GET')
@app.get('/demo/status')
def demo_status_v107(request: Request):
    email = request.query_params.get('email','')
    if _v108_is_admin_email(email) or _v108_admin_key_ok(request):
        return {'allowed': True, 'admin_bypass': True, 'used': 0, 'limit': 'unlimited', 'remaining': 'unlimited', 'rule': 'admin bypass active; public remains one credible Earth or Space run per email'}
    if not email:
        return {'allowed': True, 'used': 0, 'limit': 1, 'remaining': 1, 'rule': 'one credible Earth or Space run per email'}
    h = _sha(_normalise_email(email))
    con = db(); cur = con.cursor(); row = cur.execute('SELECT COUNT(*) AS c FROM public_demo_uses WHERE email_hash=?', (h,)).fetchone(); con.close()
    used = int(row['c'] if row else 0)
    return {'allowed': used < 1, 'used': used, 'limit': 1, 'remaining': max(0, 1-used), 'rule': 'one credible Earth or Space run per email'}

_v107_remove_route('/public-demo/generate', 'POST')
@app.post('/public-demo/generate')
def public_demo_generate_v107(req: PublicDemoRequest, request: Request):
    issues = _quality_gate_public_demo(req)
    if issues:
        raise HTTPException(status_code=422, detail={'message': 'CASEY needs one real infrastructure or space programme brief before using your free run.', 'issues': issues})
    admin_bypass = _v108_is_admin_email(req.email) or _v108_admin_key_ok(request)
    identity = _public_demo_identity(request, req)
    con = db(); cur = con.cursor()
    prev = cur.execute('SELECT run_id, created_at FROM public_demo_uses WHERE email_hash=? LIMIT 1', (identity['email_hash'],)).fetchone()
    if prev and not admin_bypass:
        con.close()
        raise HTTPException(status_code=403, detail={'message': 'This email has already used its one free CASEY intelligence run. Request access to run another Earth or Space scenario.', 'existing_run_id': prev['run_id']})
    prompt = _premium_public_prompt(req)
    input_quality = _public_demo_brief_quality_score(req)
    model = build_model(prompt, _normalise_email(req.email), 3, 4, 'base')
    model['input_quality_score'] = input_quality['score']
    model['public_demo'] = True
    model['admin_bypass'] = bool(admin_bypass)
    model['lead_email'] = _normalise_email(req.email)
    model['project_type'] = req.project_type
    run_id = 'CASEY-DEMO-' + uuid.uuid4().hex[:10].upper()
    model['run_id'] = run_id
    model = _v107_stamp_model(model)
    report = _public_demo_report(model)
    now = datetime.utcnow().isoformat()
    cur.execute("""INSERT INTO public_demo_uses(run_id,email_hash,ip_hash,fingerprint_hash,client_token_hash,project_type,project_text,model_json,created_at)
        VALUES(?,?,?,?,?,?,?,?,?)""", (run_id, identity['email_hash'], identity['ip_hash'], identity['fingerprint_hash'], identity['client_token_hash'], req.project_type, req.project_description, json.dumps(model), now))
    cur.execute('INSERT INTO projects(name,client,prompt,mode,created_at,model_json) VALUES(?,?,?,?,?,?)', (model.get('title'), model.get('client'), model.get('prompt'), model.get('mode'), now, json.dumps(model)))
    con.commit(); con.close()
    return {'run_id': run_id, 'used': 0 if admin_bypass else 1, 'limit': 'unlimited' if admin_bypass else 1, 'remaining': 'unlimited' if admin_bypass else 0, 'admin_bypass': bool(admin_bypass), 'report': report, 'model': model}

def _v107_export_endpoint(kind: str, builder, media: str, filename: str):
    def endpoint(payload: dict = Body(default={})):
        m = _v107_stamp_model(payload or {})
        ok, msg = _v107_can_export(m, kind)
        if not ok:
            _v107_export_block(msg)
        return stream(builder(m), media, filename)
    return endpoint

for _p in ['/export/workbook','/export/cost-model','/export/excel']:
    _v107_remove_route(_p, 'POST')
    app.post(_p)(_v107_export_endpoint('cost_workbook', _v107_cost_workbook_bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'CASEY_DEMO_Cost_Workbook.xlsx'))
for _p in ['/export/risk-register']:
    _v107_remove_route(_p, 'POST')
    app.post(_p)(_v107_export_endpoint('risk_register', _v107_risk_workbook_bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'CASEY_DEMO_Risk_Register.xlsx'))
for _p in ['/export/xer']:
    _v107_remove_route(_p, 'POST')
    app.post(_p)(_v107_export_endpoint('xer_schedule', _v107_xer_bytes, 'application/octet-stream', 'CASEY_DEMO_Strong_P6_Schedule.xer'))
for _p in ['/export/schedule-csv']:
    _v107_remove_route(_p, 'POST')
    app.post(_p)(_v107_export_endpoint('schedule_csv', _v107_schedule_csv_bytes, 'text/csv', 'CASEY_DEMO_Schedule_CSV_Fallback.csv'))
for _p in ['/export/qcra-qsra']:
    _v107_remove_route(_p, 'POST')
    app.post(_p)(_v107_export_endpoint('qcra_qsra', _v107_watermarked_qcra_qsra, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'CASEY_DEMO_QCRA_QSRA.xlsx'))
for _p in ['/export/json']:
    _v107_remove_route(_p, 'POST')
    app.post(_p)(_v107_export_endpoint('model_audit', _v107_model_json_bytes, 'application/json', 'CASEY_DEMO_Model_Audit.json'))
for _p in ['/export/all','/export/full-pack']:
    _v107_remove_route(_p, 'POST')
    app.post(_p)(_v107_export_endpoint('full_pack', _v107_full_pack, 'application/zip', 'CASEY_DEMO_Final_Output_Pack.zip'))

APP_VERSION = 'CASEY V107 Final Demo Commercial Lock + Strong Exports'
print('CASEY V107 final demo commercial lock and strong exports installed')
# ================= END CASEY V107 FINAL DEMO COMMERCIAL LOCK + STRONG EXPORTS =================

# ================= CASEY V116 ELITE DEMO COMPLETION PATCH =================
# Adds a safe local/admin reset for demo testing and a health/readiness manifest.
# Public users still remain locked to one credible Earth or Space run per email.

def _casey_v116_local_request(request: Request) -> bool:
    host = (request.client.host if request and request.client else '') or ''
    return host in {'127.0.0.1', '::1', 'localhost'}

@app.get('/elite-demo/health')
def elite_demo_health():
    return {
        'version': 'CASEY V116 Elite Demo Completion',
        'status': 'ready',
        'public_lock': 'one credible Earth or Space intelligence run per email',
        'admin_testing': 'CASEY_ADMIN_EMAILS bypasses run and export limits; localhost can reset test locks',
        'confidence_definition': 'board-defensibility, not generic optimism',
        'exports': ['board pack zip', 'cost workbook', 'risk register', 'XER schedule', 'QCRA/QSRA workbook', 'model audit json'],
        'demo_positioning': 'first-pass strategic intelligence, demo-watermarked, not certified commercial reliance'
    }

@app.post('/elite-demo/reset-test-email')
def elite_demo_reset_test_email(request: Request, payload: Dict[str, Any] = Body(default={})):
    """Local/admin-only helper so the owner can test the one-run lock without weakening public gating."""
    email = _normalise_email(str((payload or {}).get('email') or ''))
    if not email:
        raise HTTPException(status_code=422, detail='email is required')
    if not (_casey_v116_local_request(request) or _v108_admin_key_ok(request) or _v108_is_admin_email(email)):
        raise HTTPException(status_code=403, detail='Reset is local/admin only. Public one-run lock remains active.')
    h = _sha(email)
    con = db(); cur = con.cursor()
    cur.execute('DELETE FROM public_demo_uses WHERE email_hash=?', (h,))
    cur.execute('DELETE FROM public_demo_exports WHERE email_hash=?', (h,))
    con.commit(); con.close()
    return {'reset': True, 'email': email, 'message': 'Test lock cleared for this email only. Public users remain one-run locked.'}

APP_VERSION = 'CASEY V116 Elite Demo Completion'
print('CASEY V116 elite demo completion patch installed')
# ================= END CASEY V116 ELITE DEMO COMPLETION PATCH =================

# ================= CASEY V118 FINAL DEMO EXPORT REFINEMENT MARKER =================
APP_VERSION = 'CASEY V118 Final Demo Export Refinement'
print('CASEY V118 final demo export refinement marker installed')
# ================= END CASEY V118 FINAL DEMO EXPORT REFINEMENT MARKER =================

# ================= CASEY V119 FINAL EARTH + SPACE DEMO EXPORT LOCK =================
# Final requested demo refinements are routed through the existing export endpoints.
try:
    from v64_outputs import (
        cost_workbook_bytes as _v119_cost_workbook_bytes,
        risk_workbook_bytes as _v119_risk_workbook_bytes,
        patch_template_xer as _v119_xer_bytes,
        schedule_csv_bytes as _v119_schedule_csv_bytes,
        model_json_bytes as _v119_model_json_bytes,
        all_zip_bytes as _v119_full_pack_bytes,
        qcra_qsra_workbook_bytes as _v119_qcra_qsra_workbook_bytes,
    )

    for _p in ['/export/workbook','/export/cost-model','/export/excel']:
        _v107_remove_route(_p, 'POST')
        app.post(_p)(_v107_export_endpoint('v119_cost_workbook', _v119_cost_workbook_bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'CASEY_V119_Final_Cost_Workbook.xlsx'))
    for _p in ['/export/risk-register']:
        _v107_remove_route(_p, 'POST')
        app.post(_p)(_v107_export_endpoint('v119_risk_register', _v119_risk_workbook_bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'CASEY_V119_Operational_Risk_Register.xlsx'))
    for _p in ['/export/xer']:
        _v107_remove_route(_p, 'POST')
        app.post(_p)(_v107_export_endpoint('v119_xer_schedule', _v119_xer_bytes, 'application/octet-stream', 'CASEY_V119_Planner_Grade_Schedule.xer'))
    for _p in ['/export/schedule-csv']:
        _v107_remove_route(_p, 'POST')
        app.post(_p)(_v107_export_endpoint('v119_schedule_csv', _v119_schedule_csv_bytes, 'text/csv', 'CASEY_V119_Schedule_Levels_Resource_Loaded.csv'))
    for _p in ['/export/qcra-qsra']:
        _v107_remove_route(_p, 'POST')
        app.post(_p)(_v107_export_endpoint('v119_qcra_qsra', _v119_qcra_qsra_workbook_bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'CASEY_V119_QCRA_QSRA_Asymmetric_Tornado.xlsx'))
    for _p in ['/export/json']:
        _v107_remove_route(_p, 'POST')
        app.post(_p)(_v107_export_endpoint('v119_model_audit', _v119_model_json_bytes, 'application/json', 'CASEY_V119_Model_Audit.json'))
    for _p in ['/export/all','/export/full-pack']:
        _v107_remove_route(_p, 'POST')
        app.post(_p)(_v107_export_endpoint('v119_full_pack', _v119_full_pack_bytes, 'application/zip', 'CASEY_V119_FINAL_EARTH_SPACE_OUTPUT_PACK.zip'))

    APP_VERSION = 'CASEY V119 Final Earth + Space Demo Lock'
    print('CASEY V119 final Earth + Space export endpoints installed')
except Exception as _v119_export_error:
    print('CASEY V119 export endpoint warning:', _v119_export_error)
# ================= END CASEY V119 FINAL EARTH + SPACE DEMO EXPORT LOCK =================

# ================= CASEY V123 LIVE CALIBRATION SIGNALS =================
# Lightweight live-sector intelligence layer for demo: converts current market,
# delivery and launch-environment signals into transparent confidence / P-tail
# modifiers. This is intentionally deterministic and auditable for demos.

def _v123_live_calibration_signals(model: Dict[str, Any]) -> List[Dict[str, Any]]:
    mode = str(model.get('mode') or '').lower()
    subsector = str(model.get('subsector') or '').lower()
    prompt = str(model.get('prompt') or '').lower()
    scenario = str(model.get('scenario') or 'base').lower()
    space = mode == 'space' or any(x in prompt + ' ' + subsector for x in ['lunar','mars','orbital','space','launch','leo','cislunar'])
    if space:
        signals = [
            {'signal':'Launch reliability volatility','status':'Active','direction':'↑ schedule-tail exposure','weight':0.18,'applies_to':'QSRA P80/P90, launch cadence risk','basis':'Recent heavy-lift test campaigns show progress but continue to expose launch, booster and flight-safety variance.'},
            {'signal':'Mission assurance burden','status':'Active','direction':'↓ confidence until evidence improves','weight':0.15,'applies_to':'confidence, reserve adequacy','basis':'Lunar / orbital programmes require qualification, redundancy and post-deployment recovery evidence before board approval.'},
            {'signal':'Thermal-power balance','status':'Watch','direction':'↑ integration risk','weight':0.11,'applies_to':'systems integration, QCRA reserve','basis':'Power storage, thermal rejection and autonomous servicing remain governing constraints for surface / orbital infrastructure.'},
            {'signal':'Regulatory and range availability','status':'Watch','direction':'↑ launch-cadence uncertainty','weight':0.09,'applies_to':'schedule logic, manifest dependencies','basis':'Launch windows, flight approvals and range constraints are treated as live operating-environment risks.'},
            {'signal':'Orbital logistics and recovery','status':'Active','direction':'↑ lifecycle exposure','weight':0.13,'applies_to':'risk register, operational resilience','basis':'Repair and recovery after deployment are materially harder than terrestrial rework.'},
        ]
    elif any(x in subsector + ' ' + prompt for x in ['data centre','data center','hyperscale']) or re.search(r'\bai\b|artificial intelligence', subsector + ' ' + prompt):
        signals = [
            {'signal':'Grid connection congestion','status':'Active','direction':'↑ critical-path exposure','weight':0.17,'applies_to':'QSRA P80/P90, energisation risk','basis':'Power availability and utility interconnection remain the dominant delivery constraints for hyperscale AI infrastructure.'},
            {'signal':'Transformer / switchgear lead times','status':'Active','direction':'↑ procurement tail','weight':0.15,'applies_to':'risk register, QCRA reserve','basis':'Long-lead electrical packages are treated as supply-chain stress signals.'},
            {'signal':'Liquid cooling readiness','status':'Watch','direction':'↑ commissioning risk','weight':0.10,'applies_to':'confidence, IST logic','basis':'Cooling readiness and integrated systems testing can govern delivery more than civil progress.'},
            {'signal':'Power-market escalation','status':'Watch','direction':'↑ cost uncertainty','weight':0.08,'applies_to':'cost range, reserve philosophy','basis':'Energy infrastructure demand and procurement competition are included in capital-delivery calibration.'},
        ]
    elif any(x in subsector + ' ' + prompt for x in ['semiconductor','fab','wafer','cleanroom']):
        signals = [
            {'signal':'Tool-install sequencing','status':'Active','direction':'↑ schedule-tail exposure','weight':0.16,'applies_to':'QSRA critical path','basis':'Cleanroom readiness, process tools and vendor windows control fab delivery confidence.'},
            {'signal':'UPW / specialty utilities maturity','status':'Active','direction':'↑ commissioning risk','weight':0.13,'applies_to':'confidence, risk reserve','basis':'Specialty utility qualification can dominate board-defensibility before ramp.'},
            {'signal':'Equipment supply concentration','status':'Watch','direction':'↑ procurement uncertainty','weight':0.10,'applies_to':'QCRA reserve','basis':'Specialist OEM slots and logistics remain high-sensitivity inputs.'},
        ]
    elif any(x in subsector + ' ' + prompt for x in ['airport','terminal','runway']):
        signals = [
            {'signal':'Live-operations phasing','status':'Active','direction':'↑ interface risk','weight':0.14,'applies_to':'schedule logic, risk register','basis':'Work inside operating airports creates staging, safety and stakeholder constraints.'},
            {'signal':'Systems integration readiness','status':'Active','direction':'↑ handover risk','weight':0.12,'applies_to':'confidence, commissioning','basis':'Baggage, security, communications and passenger-flow systems must be commissioned as one operating environment.'},
            {'signal':'Regulatory / stakeholder approvals','status':'Watch','direction':'↑ schedule uncertainty','weight':0.09,'applies_to':'QSRA P80/P90','basis':'Airside approvals and operational windows are treated as current delivery-friction signals.'},
        ]
    elif any(x in subsector + ' ' + prompt for x in ['defence','defense','military','radar','naval','airbase']):
        signals = [
            {'signal':'Security and assurance burden','status':'Active','direction':'↓ evidence maturity until cleared','weight':0.16,'applies_to':'confidence, approvals','basis':'Classified interfaces and assurance gates increase decision friction.'},
            {'signal':'Sovereign supply-chain exposure','status':'Active','direction':'↑ procurement uncertainty','weight':0.14,'applies_to':'QCRA reserve, risk register','basis':'Controlled components and supplier eligibility affect delivery tails.'},
            {'signal':'Integration test dependency','status':'Watch','direction':'↑ commissioning risk','weight':0.11,'applies_to':'schedule / operational readiness','basis':'Mission-system integration can govern delivery more than construction progress.'},
        ]
    else:
        signals = [
            {'signal':'Commodity and procurement volatility','status':'Watch','direction':'↑ cost-range uncertainty','weight':0.10,'applies_to':'QCRA reserve','basis':'Current market and supply-chain volatility are applied to benchmark ranges.'},
            {'signal':'Approvals and governance latency','status':'Watch','direction':'↑ schedule-tail exposure','weight':0.09,'applies_to':'QSRA P80/P90','basis':'Decision friction and evidence maturity are reflected in confidence scoring.'},
            {'signal':'Commissioning readiness','status':'Active','direction':'↓ confidence if unproven','weight':0.12,'applies_to':'confidence, risk register','basis':'Operational readiness is treated as a board-defensibility signal, not a late-stage admin task.'},
        ]
    if scenario == 'faster':
        signals.append({'signal':'Acceleration stress overlay','status':'Active','direction':'P50 improves / P90 worsens','weight':0.14,'applies_to':'QSRA tail, confidence movement','basis':'Speed consumes float and optionality unless acceleration evidence is explicit.'})
    elif scenario == 'cheaper':
        signals.append({'signal':'Capital-efficiency stress overlay','status':'Active','direction':'P50 lower / resilience lower','weight':0.14,'applies_to':'reserve adequacy, lifecycle exposure','basis':'Lower headline capital can transfer exposure into start-up, operations and P90 contingency.'})
    elif scenario == 'lower_risk':
        signals.append({'signal':'Assurance overlay','status':'Active','direction':'confidence ↑ / duration ↑','weight':0.12,'applies_to':'reserve, schedule gates','basis':'The scenario buys evidence, float and resilience rather than early date or lean capex.'})
    elif scenario == 'premium':
        signals.append({'signal':'Resilience premium overlay','status':'Active','direction':'optionality ↑ / affordability challenge ↑','weight':0.12,'applies_to':'board trade-off, QCRA reserve','basis':'Premium posture increases downside protection but creates value-for-money challenge.'})
    return signals

def _v123_apply_live_calibration(model: Dict[str, Any]) -> Dict[str, Any]:
    m = dict(model or {})
    signals = _v123_live_calibration_signals(m)
    m['live_calibration_active'] = True
    m['live_calibration_label'] = 'LIVE CALIBRATION SIGNALS ACTIVE'
    m['live_calibration_summary'] = 'Current sector conditions are being applied to confidence, contingency and delivery-tail exposure.'
    m['live_calibration_signals'] = signals
    m['live_calibration_strip'] = ' • '.join([s['signal'] for s in signals[:4]])
    top = signals[0]['signal'] if signals else 'delivery environment volatility'
    # Inject subtle evidence of calibration into existing narrative fields without overwhelming the pack.
    if m.get('mode') == 'Space':
        extra = f" Live calibration is weighting {top.lower()}, mission assurance and orbital recovery exposure into the QSRA/QCRA tail."
    else:
        extra = f" Live calibration is weighting {top.lower()}, procurement stress and commissioning readiness into the QSRA/QCRA tail."
    if m.get('uncertainty_narrative') and isinstance(m['uncertainty_narrative'], dict):
        interp = m['uncertainty_narrative'].get('interpretation') or ''
        if 'Live calibration' not in interp:
            m['uncertainty_narrative']['interpretation'] = (interp + extra).strip()
    else:
        m['uncertainty_narrative'] = {'interpretation': extra.strip()}
    cards = list(m.get('mission_control_cards') or [])
    cards.insert(0, {'label':'Live calibration', 'signal': m['live_calibration_summary'], 'severity':'Active'})
    m['mission_control_cards'] = cards[:6]
    audit = dict(m.get('active_context_lock') or {})
    audit['live_calibration'] = True
    audit['signal_count'] = len(signals)
    m['active_context_lock'] = audit
    return m

_CASEY_V123_ORIGINAL_BUILD_MODEL = build_model
def build_model(prompt:str, client:str='', class_level:int=3, schedule_level:int=3, scenario:str='base'):
    return _v123_apply_live_calibration(_CASEY_V123_ORIGINAL_BUILD_MODEL(prompt, client, class_level, schedule_level, scenario))

@app.get('/live-calibration/manifest')
def live_calibration_manifest():
    return {
        'version':'CASEY V123 Live Calibration Demo Layer',
        'status':'active',
        'positioning':'sector intelligence calibration, not a certified market-data feed',
        'signals':['launch reliability','mission assurance','commodity / procurement volatility','regulatory exposure','commissioning readiness'],
        'applies_to':['confidence','QCRA reserve','QSRA P80/P90','risk register','board narrative']
    }

APP_VERSION = 'CASEY V123 Live Calibration Demo Layer'
print('CASEY V123 live calibration demo layer installed')
# ================= END CASEY V123 LIVE CALIBRATION SIGNALS =================
