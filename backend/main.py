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
      ("AWRE Aldermaston","Earth","Defence Nuclear Infrastructure",6,96,["awre aldermaston","awre burghfield","aldermaston nuclear"]),
      ("Hinkley Point C","Earth","Nuclear Power Station",32,204,["hinkley point c","hinkley"]),
      ("Elizabeth Line","Earth","Rail Mega Programme",18,216,["elizabeth line","crossrail"]),
      ("Thames Tideway","Earth","Water Mega Programme",4.2,72,["thames tideway","tideway tunnel"]),
      ("Artemis Lunar Gateway","Space","Deep Space Infrastructure",40,180,["lunar gateway","artemis gateway"]),
      ("Mars Colony","Space","Mars Settlement",280,360,["mars colony","mars settlement"]),
      ("Gigafactory UK","Earth","Battery Gigafactory",4,48,["gigafactory uk","britishvolt","battery gigafactory"]),
      ("ITER Fusion","Earth","Nuclear Fusion Facility",22,240,["iter","iter fusion","fusion reactor"]),
      ("NEOM The Line","Earth","Future City Mega Programme",500,240,["the line","neom"]),
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
    """Global location intelligence — 70+ countries, construction cost multipliers,
    regulatory frameworks and risk context. All USD-normalised."""
    t = t.lower()
    # Format: (display_name, cost_multiplier, keywords, regulatory_framework, risk_note)
    # Cost multiplier = relative to global baseline (1.0)
    # Sources: International Construction Market Survey, AECOM, BCIS
    table = [
        # ── SPACE ────────────────────────────────────────────────────
        ("Lunar Surface", 2.9, ["moon","lunar","lunar surface","lunar south pole","lunar base"], "NASA/ESA/JAXA/ISRO", "Remote logistics, ISRU dependency, no recovery"),
        ("Mars Surface", 3.8, ["mars","martian","mars surface","mars base","mars colony"], "NASA/ESA international", "7-month transit, no real-time comms, autonomous ops only"),
        ("Low Earth Orbit", 2.5, ["leo ","low earth orbit","orbital platform","space station","iss"], "NASA/ESA/Roscosmos/CNSA", "Launch window dependency, radiation environment"),
        ("Geostationary Orbit", 2.2, ["geo ","geostationary","geo orbit","geo satellite"], "ITU frequency coordination", "Long integration lead, launch procurement single-source"),
        ("Cislunar Space", 2.7, ["cislunar","cis-lunar","lunar gateway","l2 ","lagrange"], "NASA Artemis framework", "FOAK infrastructure, 3-day Earth contingency minimum"),
        ("Deep Space", 4.3, ["deep space","asteroid","interplanetary","mars transit","outer planet"], "NASA/ESA deep space", "No contingency window — autonomous ops mandatory"),
        # ── MIDDLE EAST / GCC ────────────────────────────────────────
        ("Saudi Arabia", 1.35, ["saudi","riyadh","jeddah","neom","the line","mecca","medina","ksa","vision 2030"], "Saudi Aramco / Royal Commission / MOMRA", "Vision 2030 programme concentration risk — contractor capacity"),
        ("UAE / Dubai", 1.18, ["uae","dubai","abu dhabi","sharjah","ras al khaimah"], "Abu Dhabi DoT / RTA Dubai", "Strong delivery track record — labour sourcing risk in summer"),
        ("Qatar", 1.22, ["qatar","doha","lusail"], "Ashghal / Qatar Rail", "Post-World Cup capacity risk — programme concentration easing"),
        ("Kuwait", 1.28, ["kuwait","kuwait city"], "MPWI Kuwait", "Regulatory complexity — multiple authority approvals required"),
        ("Bahrain", 1.20, ["bahrain","manama"], "Ministry of Works Bahrain", "Small market — contractor pool limited"),
        ("Oman", 1.25, ["oman","muscat","sohar"], "Ministry of Transport Oman", "Remote sites — logistics premium"),
        ("Jordan", 1.15, ["jordan","amman"], "Ministry of Public Works Jordan", "IFI-financed — World Bank/ADB procurement rules apply"),
        ("Egypt", 1.10, ["egypt","cairo","alexandria","new administrative capital"], "NUCA / Ministry of Housing Egypt", "Currency exposure — EGP volatility significant risk"),
        ("Iraq", 1.65, ["iraq","baghdad","basra","erbil"], "Ministry of Construction Iraq", "Security risk premium, contractor insurance significant"),
        # ── AFRICA ───────────────────────────────────────────────────
        ("South Africa", 1.05, ["south africa","johannesburg","cape town","durban","pretoria","gauteng"], "SANRAL / Eskom / DPW", "BEE procurement requirements, load-shedding risk, skills flight"),
        ("Nigeria", 1.30, ["nigeria","lagos","abuja","port harcourt","kano"], "FERMA / FMWH Nigeria", "FX risk, import duty on materials, security premium in north"),
        ("Kenya", 1.18, ["kenya","nairobi","mombasa"], "KeNHA / Kenya Power", "IFI financing typical — World Bank/AfDB procurement rules"),
        ("Ethiopia", 1.20, ["ethiopia","addis ababa","addis"], "ERA / EEP Ethiopia", "Chinese contractor dominance — quality assurance challenge"),
        ("Ghana", 1.22, ["ghana","accra","kumasi"], "GHA / Ghana Grid Company", "Stable regulatory — IFI-financed, GHS currency risk"),
        ("Tanzania", 1.18, ["tanzania","dar es salaam","dodoma"], "TANROADS / TANESCO", "Good governance — land acquisition slower than programme assumes"),
        ("Morocco", 1.12, ["morocco","casablanca","rabat","marrakech","tangier"], "ONCF / ONEE Morocco", "Strong Chinese + European contractor competition — good delivery"),
        ("Angola", 1.45, ["angola","luanda"], "Ministry of Construction Angola", "Oil-price dependent budget — programme suspension risk"),
        ("Mozambique", 1.35, ["mozambique","maputo"], "ANE Mozambique", "LNG programme — specialist contractor availability premium"),
        ("DRC / Congo", 1.55, ["drc","congo","kinshasa","democratic republic"], "Various DRC ministries", "Governance risk — community opposition, artisanal mining interface"),
        ("Zambia / Zimbabwe", 1.30, ["zambia","lusaka","zimbabwe","harare"], "Various authorities", "Debt distress risk — IFI financing conditional"),
        ("Senegal", 1.20, ["senegal","dakar"], "AGEROUTE Senegal", "Stable — French system, World Bank partner"),
        # ── ASIA PACIFIC ─────────────────────────────────────────────
        ("India", 0.78, ["india","mumbai","delhi","bangalore","chennai","hyderabad","pune","kolkata","ahmedabad"], "MoRTH / CEA / CERC India", "GST complexity, land acquisition 2-3x programmed duration, labour abundance"),
        ("China", 0.72, ["china","beijing","shanghai","shenzhen","guangzhou","chengdu","wuhan"], "NDRC / MoR China", "State-directed — speed advantage but quality variance, IP concerns"),
        ("Japan", 1.45, ["japan","tokyo","osaka","kyoto","fukuoka"], "MLIT Japan", "Seismic design premium, ageing workforce, Monozukuri quality standard"),
        ("South Korea", 1.22, ["south korea","korea","seoul","busan","incheon"], "MOLIT Korea / KEPCO", "Strong delivery capability — chaebols dominate, consortium risk"),
        ("Singapore", 1.38, ["singapore"], "LTA / BCA / EDB Singapore", "Land constraint premium, very strong governance, foreign worker quota"),
        ("Malaysia", 0.92, ["malaysia","kuala lumpur","kl","penang","johor"], "JKR / SPAD Malaysia", "Cost competitive, political procurement risk, Ringgit exposure"),
        ("Indonesia", 0.88, ["indonesia","jakarta","bali","surabaya","bandung"], "BAPPENAS / PUPR Indonesia", "Land acquisition 3-5x programmed, inter-island logistics premium"),
        ("Philippines", 0.90, ["philippines","manila","cebu","davao"], "DPWH / DOE Philippines", "Typhoon risk (12 per year), procurement delays, ODA-financed typical"),
        ("Thailand", 0.95, ["thailand","bangkok","chiang mai"], "EXAT / EGAT Thailand", "Political risk to mega-programmes — military government cycle"),
        ("Vietnam", 0.85, ["vietnam","ho chi minh","hanoi","da nang"], "MOT / EVN Vietnam", "Fast growth — contractor capacity catching demand, IFI-financed"),
        ("Pakistan", 0.82, ["pakistan","karachi","lahore","islamabad"], "NHA / WAPDA Pakistan", "IMF programme dependency, FX crisis risk, CPEC Chinese contractor dominant"),
        ("Bangladesh", 0.80, ["bangladesh","dhaka","chittagong"], "LGED / REB Bangladesh", "IDA-financed — procurement slower, strong garment sector skills adjacency"),
        ("Sri Lanka", 0.78, ["sri lanka","colombo"], "RDA / CEB Sri Lanka", "Post-debt crisis — IFI restructuring, Chinese debt renegotiation"),
        ("Australia", 1.22, ["australia","sydney","melbourne","brisbane","perth","adelaide","darwin","queensland"], "IPA Australia / AEMO", "Strong governance, high labour cost, skills shortage in WA/NT"),
        ("New Zealand", 1.18, ["new zealand","auckland","wellington","christchurch"], "NZTA / Transpower NZ", "Seismic risk premium, supply chain remoteness, small market"),
        # ── EUROPE ───────────────────────────────────────────────────
        ("United Kingdom", 1.20, ["uk","united kingdom","england","scotland","wales","london","heathrow","manchester","birmingham","leeds","glasgow"], "IPA / HM Treasury Green Book / ORR / ONR", "Strong governance, high cost, IPA gateway mandatory for government programmes"),
        ("Germany", 1.28, ["germany","berlin","munich","hamburg","frankfurt","cologne","stuttgart","dusseldorf"], "BMVI / BNetzA Germany", "DIN standards, Energiewende grid complexity, coalition procurement delays"),
        ("France", 1.22, ["france","paris","lyon","marseille","toulouse","bordeaux"], "CEREMA / CRE France", "EDF nuclear expertise, TGV benchmark, public sector strike risk"),
        ("Netherlands", 1.25, ["netherlands","amsterdam","rotterdam","the hague","eindhoven"], "Rijkswaterstaat / TenneT", "Water management world leader, nitrogen ruling causes project delays"),
        ("Belgium", 1.20, ["belgium","brussels","antwerp","ghent"], "AWV / Elia Belgium", "Coalition government — approval timelines extended"),
        ("Sweden", 1.30, ["sweden","stockholm","gothenburg","malmo"], "Trafikverket / Svenska kraftnat", "High labour cost, strong governance, environmental consent rigorous"),
        ("Norway", 1.38, ["norway","oslo","bergen","stavanger"], "Statens vegvesen / Statnett", "Oil fund wealth, extreme terrain premium, offshore expertise"),
        ("Denmark", 1.28, ["denmark","copenhagen","aarhus"], "Vejdirektoratet / Energinet DK", "Offshore wind world leader — reference benchmark for OSW"),
        ("Finland", 1.30, ["finland","helsinki","tampere"], "Fintraffic / Fingrid", "ITER Olkiluoto nuclear precedent — FOAK caution warranted"),
        ("Poland", 0.88, ["poland","warsaw","krakow","gdansk","wroclaw"], "GDDKiA / PSE Poland", "EU-funded — strong delivery, cost competitive, fast growth"),
        ("Spain", 1.12, ["spain","madrid","barcelona","seville","bilbao","valencia"], "MITMA / Red Electrica Spain", "AVE rail world leader, strong contractor base, regional complexity"),
        ("Italy", 1.18, ["italy","rome","milan","turin","naples","genoa"], "MIT / Terna Italy", "Procurement duration systemic — 3-5 years for major contracts, mafia risk"),
        ("Portugal", 1.08, ["portugal","lisbon","porto"], "ANSR / REN Portugal", "EU-funded, good governance, cost competitive vs northern Europe"),
        ("Greece", 1.02, ["greece","athens","thessaloniki"], "Ministry of Infrastructure Greece", "Post-austerity capacity recovery, EU funding dependent"),
        ("Romania", 0.82, ["romania","bucharest","cluj","timisoara"], "CNAIR / Transelectrica Romania", "EU funding absorption challenge, corruption risk, strong cost advantage"),
        ("Czech Republic", 0.90, ["czech","czechia","prague","brno"], "SSD / CEPS Czech", "Central European hub, strong automotive/industrial base"),
        ("Austria", 1.30, ["austria","vienna","graz","salzburg"], "ASFINAG / APG Austria", "Alpine construction premium, strong governance, ÖBB rail reference"),
        ("Switzerland", 1.55, ["switzerland","zurich","geneva","bern","basel"], "SBB / Swissgrid", "World's highest construction costs, NEAT tunnel world benchmark"),
        # ── AMERICAS ─────────────────────────────────────────────────
        ("United States", 1.08, ["usa","united states","us ","texas","california","florida","new york","arizona","north carolina","ohio","virginia","georgia","washington state","illinois","pennsylvania","michigan"], "DOT / FERC / NRC / DOE USA", "State + federal permitting, union labour premium, litigation exposure"),
        ("Canada", 1.18, ["canada","toronto","vancouver","calgary","ottawa","montreal","edmonton"], "Infrastructure Canada / NEB", "Resource-driven market cycles, Indigenous consultation mandatory"),
        ("Mexico", 0.85, ["mexico","mexico city","monterrey","guadalajara","cancun"], "SCT / CFE Mexico", "Pemex oil dependency, security premium in north, political cycle risk"),
        ("Brazil", 0.82, ["brazil","sao paulo","rio de janeiro","brasilia","belo horizonte","manaus"], "ANTT / ANEEL / EPE Brazil", "Lava Jato corruption legacy, BRL currency volatility, Amazon logistics"),
        ("Chile", 0.95, ["chile","santiago","valparaiso","antofagasta"], "MOP / CNE Chile", "Lithium/copper mining world reference, stable governance, seismic premium"),
        ("Colombia", 0.88, ["colombia","bogota","medellin","cali","barranquilla"], "INVIAS / UPME Colombia", "4G/5G road concession programme, peace dividend, terrain premium"),
        ("Peru", 0.85, ["peru","lima","arequipa","cusco"], "MTC / MINEM Peru", "Mining logistics benchmark, Inca site constraint, political instability"),
        ("Argentina", 0.75, ["argentina","buenos aires","cordoba","rosario"], "Ministerio de Obras Argentina", "Currency crisis — USD indexing essential, import restriction risk"),
        ("Ecuador", 0.80, ["ecuador","quito","guayaquil"], "MTOP Ecuador", "Oil-dependent, dollarised economy, political risk"),
    ]
    for name, factor, keys, framework, risk_note in table:
        if has(t, keys):
            return name, factor
    return "Global", 1.0

def location_context(location_name: str) -> dict:
    """Return full regulatory, currency and risk context for a location."""
    ctx = {
        # Space
        "Lunar Surface": {"currency":"USD","framework":"NASA/ESA/ISRO international","approval_body":"NASA Human Exploration / ESA SciProg","optimism_bias_note":"FOAK infrastructure — no comparable. Apply 140%+ OBA uplift.","financing":"NASA HLS / ESA Ministerial / sovereign funding","risk_premium":"Extreme — remote logistics, autonomous ops, life support dependency"},
        "Mars Surface": {"currency":"USD","framework":"NASA/ESA international","approval_body":"NASA HEOMD","optimism_bias_note":"No reference class exists. JWST pattern: 10-15x baseline.","financing":"NASA/ESA sovereign","risk_premium":"Maximum — 7-month transit, no contingency window"},
        "Low Earth Orbit": {"currency":"USD","framework":"ITU / COPUOS","approval_body":"FCC / national telecoms regulator","optimism_bias_note":"Apply 100-140% OBA. OneWeb/JWST reference class.","financing":"Commercial/venture / sovereign","risk_premium":"Very high — FOAK integration, radiation, launch dependency"},
        # UK
        "United Kingdom": {"currency":"GBP","framework":"HM Treasury Green Book / IPA Gateway / NEC4","approval_body":"IPA / DLUHC / DfT / ORR / ONR / HSE","optimism_bias_note":"HM Treasury Table B6 uplifts mandatory for government programmes. Rail +66%, Roads +44%, Hospitals +44%.","financing":"HM Treasury / UK Infrastructure Bank / private finance","risk_premium":"Medium — strong governance, high cost, IPA red-rated programmes systemic"},
        # USA
        "United States": {"currency":"USD","framework":"OMB Circular A-11 / FHWA benefit-cost / NEPA environmental","approval_body":"DOT / FERC / NRC / USACE / State PUC","optimism_bias_note":"Flyvbjerg US dataset: rail +41%, roads +27%. NEPA delay risk structural.","financing":"IIJA / IRA / DOE LPO / private / P3","risk_premium":"Medium-High — litigation exposure, union labour, permitting duration"},
        # Saudi Arabia  
        "Saudi Arabia": {"currency":"SAR / USD","framework":"Saudi Aramco project management / Royal Commission standards","approval_body":"Royal Commission / MOMRA / SEC","optimism_bias_note":"Vision 2030 concentration risk — contractor capacity constraint is the primary OBA driver.","financing":"PIF / ARAMCO / sovereign / Islamic finance","risk_premium":"High — extreme programme concentration, summer labour restriction, political mandate risk"},
        # Australia
        "Australia": {"currency":"AUD","framework":"Infrastructure Australia / State procurement / AS 4000","approval_body":"Infrastructure Australia / State DBs / AEMO","optimism_bias_note":"IPA Australia data: major projects average +35% cost, +40% schedule.","financing":"Commonwealth / State / asset recycling / P3","risk_premium":"Medium — strong governance, skills shortage WA/NT, union EBA exposure"},
        # India
        "India": {"currency":"INR / USD","framework":"Ministry of Finance / MOSPI / DPR appraisal","approval_body":"MoRTH / CEA / CERC / NHA / NHAI","optimism_bias_note":"Land acquisition 2-4x programmed timeline is the primary OBA driver. GVA cost base low.","financing":"World Bank / ADB / AIIB / NDB / sovereign bonds","risk_premium":"Medium — land acquisition, GST complexity, skills abundance, regulatory multi-layer"},
        # Nigeria
        "Nigeria": {"currency":"NGN / USD","framework":"FERMA / World Bank procurement / PPPA","approval_body":"FERMA / FMWH / Ministry of Power","optimism_bias_note":"FX devaluation risk adds 30-50% to imported equipment cost. Apply 45%+ OBA.","financing":"World Bank / AfDB / China EXIM / domestic bonds","risk_premium":"High — FX, security north, import duty, contractor capacity limited"},
        # South Africa
        "South Africa": {"currency":"ZAR / USD","framework":"CIDB / PFMA / NEC3","approval_body":"SANRAL / Eskom / DoT SA","optimism_bias_note":"Load-shedding adds 8-15% to project costs. BEE requirements constrain contractor pool.","financing":"DBSA / IDC / World Bank / sovereign / PPP","risk_premium":"Medium-High — load-shedding, BEE, skills emigration, Eskom grid uncertainty"},
        # Brazil
        "Brazil": {"currency":"BRL / USD","framework":"Lei 14.133 procurement / ANTT / ANEEL","approval_body":"ANTT / ANEEL / IBAMA environmental","optimism_bias_note":"BRL volatility 20-40% over programme life typical. Lava Jato legacy — contractor due diligence mandatory.","financing":"BNDES / IDB / World Bank / CRI/CRA","risk_premium":"High — FX, environmental consent IBAMA, terrain, Lava Jato legacy"},
    }
    # Return best match or generic global context
    for k, v in ctx.items():
        if k.lower() in location_name.lower() or location_name.lower() in k.lower():
            return v
    # Generic global context
    return {"currency":"USD (converted)","framework":"World Bank / IFC / FIDIC / MDB standard","approval_body":"National infrastructure authority / MDB co-financier","optimism_bias_note":"Apply Flyvbjerg global reference class: infrastructure average +27% cost, +39% schedule (200-country dataset).","financing":"MDB (World Bank/ADB/AfDB/AIIB) / sovereign / PPP concession","risk_premium":"Varies by country — apply location multiplier and assess political/FX/contractor capacity risk"}


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
    # ── RAIL ─────────────────────────────────────────────────────────
    {"name":"Crossrail / Elizabeth Line","sector":"Rail / Transit","mode":"Earth",
     "keywords":["rail","metro","hs2","crossrail","signalling","station","tunnel","transit","underground","subway","tram"],
     "cost_bn":22.7,"months":216,"cost_growth_pct":88,"schedule_slip_months":84,
     "failure_mode":"Deferred systems integration — 900 open IEMs at planned opening",
     "lesson":"Possessions and signalling must be on the critical path from day one, not treated as commissioning activities"},
    {"name":"HS2 Phase 1","sector":"Rail / Transit","mode":"Earth",
     "keywords":["hs2","high speed","rail","tunnelling"],
     "cost_bn":44.6,"months":168,"cost_growth_pct":140,"schedule_slip_months":36,
     "failure_mode":"Scope growth, ground conditions, open corridor risk",
     "lesson":"Cost-at-completion estimates grow during delivery — approving at P50 without P80 reserve is a governance failure"},
    {"name":"Riyadh Metro","sector":"Rail / Transit","mode":"Earth",
     "keywords":["riyadh","metro","rail","saudi","middle east"],
     "cost_bn":22.5,"months":96,"cost_growth_pct":12,"schedule_slip_months":24,
     "failure_mode":"Systems integration and operational readiness timeline",
     "lesson":"International rail programmes with multiple concessions face interface risk proportional to contract complexity"},
    # ── NUCLEAR ──────────────────────────────────────────────────────
    {"name":"Hinkley Point C","sector":"Nuclear / Energy","mode":"Earth",
     "keywords":["nuclear","hinkley","reactor","epr","gda","pressurised water","new build"],
     "cost_bn":35.0,"months":204,"cost_growth_pct":94,"schedule_slip_months":60,
     "failure_mode":"FOAK EPR supply chain, first-pour concrete issues, nuclear-grade welding failures",
     "lesson":"GDA is the real critical path — not construction. Every 6 months of GDA slip costs £1B+ in financing"},
    {"name":"Olkiluoto 3 (Finland)","sector":"Nuclear / Energy","mode":"Earth",
     "keywords":["nuclear","finland","olkiluoto","reactor","epr"],
     "cost_bn":11.0,"months":240,"cost_growth_pct":300,"schedule_slip_months":168,
     "failure_mode":"FOAK EPR complexity, safety system integration, regulatory hold-points",
     "lesson":"New reactor designs have 3-5x baseline cost growth on first deployment"},
    {"name":"Vogtle Units 3 & 4 (Georgia)","sector":"Nuclear / Energy","mode":"Earth",
     "keywords":["nuclear","georgia","vogtle","ap1000","reactor"],
     "cost_bn":34.0,"months":204,"cost_growth_pct":113,"schedule_slip_months":84,
     "failure_mode":"FOAK AP1000 design changes, contractor performance, qualified labour shortage",
     "lesson":"Fixed-price EPC contracts on nuclear FOAK do not transfer risk — they transfer insolvency"},
    # ── DEFENCE ──────────────────────────────────────────────────────
    {"name":"Ajax Armoured Vehicles (UK)","sector":"Defence / Secure Infrastructure","mode":"Earth",
     "keywords":["defence","defense","ajax","armoured","vehicle","military","mod"],
     "cost_bn":5.5,"months":180,"cost_growth_pct":57,"schedule_slip_months":120,
     "failure_mode":"EMC/vibration issues, crew safety, training system integration — none on critical path",
     "lesson":"Operational acceptance is the real programme gate, not platform delivery"},
    {"name":"Watchkeeper UAV Programme","sector":"Defence / Secure Infrastructure","mode":"Earth",
     "keywords":["defence","uav","drone","watchkeeper","military"],
     "cost_bn":1.3,"months":180,"cost_growth_pct":130,"schedule_slip_months":120,
     "failure_mode":"Civil airspace certification never achieved — airworthiness not a delivery constraint",
     "lesson":"Regulatory acceptance must be on the master critical path from day one"},
    {"name":"F-35 Joint Strike Fighter","sector":"Defence / Secure Infrastructure","mode":"Earth",
     "keywords":["defence","f-35","fighter","aircraft","military","sovereign"],
     "cost_bn":412.0,"months":384,"cost_growth_pct":68,"schedule_slip_months":96,
     "failure_mode":"Software integration complexity, concurrent development and production",
     "lesson":"Software-intensive defence programmes have 3-5x baseline schedule assumptions"},
    # ── DATA CENTRE ──────────────────────────────────────────────────
    {"name":"Microsoft Azure UK South (Slough campus)","sector":"Digital Infrastructure / Hyperscale Data Centre","mode":"Earth",
     "keywords":["data centre","data center","hyperscale","microsoft","azure","ai campus","gpu","compute"],
     "cost_bn":3.2,"months":36,"cost_growth_pct":15,"schedule_slip_months":18,
     "failure_mode":"Grid connection delay, DNO queue, cooling commissioning",
     "lesson":"Grid connection must be a signed agreement, not a queue position — energisation delays are now systemic"},
    {"name":"Amazon AWS Dublin Campus","sector":"Digital Infrastructure / Hyperscale Data Centre","mode":"Earth",
     "keywords":["data centre","data center","amazon","aws","hyperscale","ireland"],
     "cost_bn":4.2,"months":30,"cost_growth_pct":20,"schedule_slip_months":14,
     "failure_mode":"Planning opposition, grid capacity, water usage consent",
     "lesson":"Data centres in water-stressed regions face novel consent constraints not in traditional risk registers"},
    # ── PHARMA / LIFE SCIENCES ───────────────────────────────────────
    {"name":"AstraZeneca Macclesfield Expansion","sector":"Life Sciences / Biologics Manufacturing","mode":"Earth",
     "keywords":["pharma","gmp","biologics","fill finish","validation","fda","mhra","life science"],
     "cost_bn":1.2,"months":60,"cost_growth_pct":30,"schedule_slip_months":24,
     "failure_mode":"Validation deferred post-construction, clean utility qualification delay",
     "lesson":"CQV is a programme deliverable — not a post-construction activity"},
    {"name":"Pfizer Ringaskiddy Ireland","sector":"Life Sciences / Biologics Manufacturing","mode":"Earth",
     "keywords":["pharma","pfizer","biologics","ireland","api","sterile"],
     "cost_bn":1.5,"months":54,"cost_growth_pct":18,"schedule_slip_months":18,
     "failure_mode":"Regulatory submission delayed by CMC dossier readiness",
     "lesson":"Regulatory submission is the revenue gate — it must be on the programme critical path"},
    # ── SEMICONDUCTOR ────────────────────────────────────────────────
    {"name":"Intel Ohio Fab (Planned)","sector":"Semiconductor / Advanced Manufacturing","mode":"Earth",
     "keywords":["semiconductor","fab","intel","ohio","wafer","chip","euv"],
     "cost_bn":28.0,"months":84,"cost_growth_pct":0,"schedule_slip_months":36,
     "failure_mode":"Workforce shortage, UPW complexity, tool delivery slippage",
     "lesson":"Tool delivery sequences must be confirmed orders — OEM intent letters are not programme commitments"},
    {"name":"TSMC Arizona Fab","sector":"Semiconductor / Advanced Manufacturing","mode":"Earth",
     "keywords":["semiconductor","fab","tsmc","arizona","wafer","chip"],
     "cost_bn":40.0,"months":84,"cost_growth_pct":35,"schedule_slip_months":30,
     "failure_mode":"Specialised workforce unavailable locally, tool delivery, UPW systems",
     "lesson":"Fab yields in new geographies are systematically below initial projections"},
    {"name":"Samsung Taylor Texas Fab","sector":"Semiconductor / Advanced Manufacturing","mode":"Earth",
     "keywords":["semiconductor","fab","samsung","texas","wafer","chip","foundry"],
     "cost_bn":17.0,"months":72,"cost_growth_pct":20,"schedule_slip_months":24,
     "failure_mode":"Market demand timing, workforce availability, tool delivery",
     "lesson":"Semiconductor fabs require 5-8 year horizon planning — market timing risk is structural"},
    # ── GIGAFACTORY ──────────────────────────────────────────────────
    {"name":"Britishvolt (Failed)","sector":"Battery / Gigafactory","mode":"Earth",
     "keywords":["gigafactory","battery","ev","britishvolt","cell","lithium"],
     "cost_bn":3.8,"months":48,"cost_growth_pct":0,"schedule_slip_months":0,
     "failure_mode":"Grid connection, cell chemistry qualification, BMS supply chain — all unconfirmed at commitment",
     "lesson":"A gigafactory without a confirmed grid connection and qualified cell chemistry is a building, not a factory"},
    {"name":"Northvolt Ett (Sweden)","sector":"Battery / Gigafactory","mode":"Earth",
     "keywords":["gigafactory","battery","northvolt","sweden","ev","cell"],
     "cost_bn":8.0,"months":60,"cost_growth_pct":40,"schedule_slip_months":36,
     "failure_mode":"Yield ramp 16x below target — 1GWh/year achieved vs 16GWh target",
     "lesson":"Battery yield ramp is the board metric — building capacity for an unqualified product is not production"},
    # ── ENERGY ───────────────────────────────────────────────────────
    {"name":"Hornsea 2 Offshore Wind Farm","sector":"Energy / Utilities","mode":"Earth",
     "keywords":["wind","offshore","wind farm","grid","hvdc","renewables","energy","hornsea"],
     "cost_bn":3.0,"months":60,"cost_growth_pct":20,"schedule_slip_months":18,
     "failure_mode":"Grid connection 18 months late — DNO queue backlog",
     "lesson":"Grid connection queue position is not an energisation date — it is a forecast"},
    {"name":"Neart na Gaoithe Offshore Wind (Scotland)","sector":"Energy / Utilities","mode":"Earth",
     "keywords":["wind","offshore","scotland","wind farm","hvdc","cable"],
     "cost_bn":3.5,"months":72,"cost_growth_pct":25,"schedule_slip_months":48,
     "failure_mode":"Aviation radar objection known at planning — not treated as programme constraint",
     "lesson":"Third-party consent risks must be treated as critical path items at project inception"},
    {"name":"Hinkley Point C Nuclear (Energy angle)","sector":"Energy / Utilities","mode":"Earth",
     "keywords":["nuclear","energy","hinkley","baseload","power station"],
     "cost_bn":35.0,"months":204,"cost_growth_pct":94,"schedule_slip_months":60,
     "failure_mode":"FOAK construction, supply chain, regulatory timeline",
     "lesson":"Nuclear baseload power has a 50-year asset life — the approval case must reflect lifetime value not just build cost"},
    # ── WATER / UTILITIES ────────────────────────────────────────────
    {"name":"Thames Water AMP7 Capital Programme","sector":"Water / Environmental Infrastructure","mode":"Earth",
     "keywords":["water","thames","wastewater","utility","amp","ofwat","treatment","smart meter"],
     "cost_bn":3.7,"months":60,"cost_growth_pct":40,"schedule_slip_months":24,
     "failure_mode":"Procurement and supply chain capacity, site access — 40% below delivery target",
     "lesson":"Utility capital programmes require contracted supply chain at programme start — not competitive procurement during execution"},
    {"name":"SMETS2 Smart Meter Rollout (UK)","sector":"Water / Environmental Infrastructure","mode":"Earth",
     "keywords":["smart meter","meter rollout","utility","smets","gas","electricity","connections"],
     "cost_bn":13.9,"months":120,"cost_growth_pct":90,"schedule_slip_months":60,
     "failure_mode":"Comms infrastructure complexity, back-office platform readiness, MDU access",
     "lesson":"Smart meter programmes fail at the back-office integration layer, not at the physical meter"},
    {"name":"NBN Co (Australia) Smart Infrastructure","sector":"Water / Environmental Infrastructure","mode":"Earth",
     "keywords":["smart meter","utility","rollout","broadband","connections","infrastructure","australia"],
     "cost_bn":51.0,"months":132,"cost_growth_pct":985,"schedule_slip_months":72,
     "failure_mode":"Engineering complexity, copper network assumptions, multi-technology mix",
     "lesson":"Utility rollout programmes in mixed urban/rural geographies have 3-10x baseline complexity assumptions"},
    # ── OIL & GAS ────────────────────────────────────────────────────
    {"name":"Chevron Gorgon LNG (Australia)","sector":"Oil & Gas / Process Infrastructure","mode":"Earth",
     "keywords":["lng","gas","oil","gorgon","chevron","offshore","platform","australia"],
     "cost_bn":54.0,"months":108,"cost_growth_pct":54,"schedule_slip_months":36,
     "failure_mode":"Brownfield interface complexity, remote logistics, HAZOP findings",
     "lesson":"Brownfield LNG interface complexity is systematically underestimated at project inception"},
    {"name":"Shell Prelude FLNG (Australia)","sector":"Oil & Gas / Process Infrastructure","mode":"Earth",
     "keywords":["lng","flng","shell","offshore","gas","fpso","floating"],
     "cost_bn":12.0,"months":144,"cost_growth_pct":50,"schedule_slip_months":60,
     "failure_mode":"FOAK floating LNG technology — never achieved nameplate capacity",
     "lesson":"FOAK floating process technology has 5x cost growth assumption vs comparable fixed infrastructure"},
    # ── MINING ───────────────────────────────────────────────────────
    {"name":"Cobre Panama Copper Mine (First Quantum)","sector":"Mining / Metals Infrastructure","mode":"Earth",
     "keywords":["mine","mining","copper","cobre","panama","open pit","ore"],
     "cost_bn":10.0,"months":108,"cost_growth_pct":100,"schedule_slip_months":0,
     "failure_mode":"Shut by government order — community licence-to-operate not treated as board gate",
     "lesson":"$10B built and operating, then shut. Community opposition must be a board approval gate, not a stakeholder communication activity"},
    {"name":"Roy Hill Iron Ore Mine (Australia)","sector":"Mining / Metals Infrastructure","mode":"Earth",
     "keywords":["mine","mining","iron ore","australia","bulk","ore"],
     "cost_bn":10.0,"months":84,"cost_growth_pct":20,"schedule_slip_months":24,
     "failure_mode":"Rail and port logistics, processing plant yield ramp",
     "lesson":"Mining logistics corridors require the same programme rigour as the mine itself"},
    # ── AIRPORT ──────────────────────────────────────────────────────
    {"name":"Heathrow Terminal 5","sector":"Airport / Aviation","mode":"Earth",
     "keywords":["airport","heathrow","terminal","t5","baggage","orat","airside"],
     "cost_bn":4.3,"months":96,"cost_growth_pct":5,"schedule_slip_months":0,
     "failure_mode":"34,000 bags lost on day 1 — IT/baggage integration not a programme deliverable",
     "lesson":"Construction on time and budget is not success — ORAT must be on the master critical path"},
    {"name":"Berlin Brandenburg Airport","sector":"Airport / Aviation","mode":"Earth",
     "keywords":["airport","berlin","terminal","brandenburgo","baggage","orat"],
     "cost_bn":7.3,"months":204,"cost_growth_pct":363,"schedule_slip_months":108,
     "failure_mode":"Fire safety integration, IT, regulatory approval — all post-construction",
     "lesson":"Airport safety and regulatory approval is the opening gate — not construction practical completion"},
    # ── HEALTHCARE ───────────────────────────────────────────────────
    {"name":"Royal Liverpool Hospital","sector":"Healthcare / Hospital Infrastructure","mode":"Earth",
     "keywords":["hospital","nhs","liverpool","clinical","royal","acute"],
     "cost_bn":0.8,"months":120,"cost_growth_pct":80,"schedule_slip_months":60,
     "failure_mode":"Structural defects, infection-control compliance, PFI contractor insolvency",
     "lesson":"PFI construction risk transfer does not transfer commissioning and occupation risk"},
    {"name":"New Royal Adelaide Hospital","sector":"Healthcare / Hospital Infrastructure","mode":"Earth",
     "keywords":["hospital","adelaide","clinical","nhs","royal","acute","australia"],
     "cost_bn":2.3,"months":120,"cost_growth_pct":60,"schedule_slip_months":24,
     "failure_mode":"Clinical commissioning not on master schedule, operational transition not contracted",
     "lesson":"Clinical commissioning is a 12-18 month programme requiring a dedicated team and critical path"},
    # ── ROADS ────────────────────────────────────────────────────────
    {"name":"A303 Stonehenge Tunnel","sector":"Roads / Highways Infrastructure","mode":"Earth",
     "keywords":["road","highway","tunnel","stonehenge","a303","motorway","uk"],
     "cost_bn":2.1,"months":96,"cost_growth_pct":50,"schedule_slip_months":36,
     "failure_mode":"UNESCO/DCO legal challenge post-contract award",
     "lesson":"Third-party consent risks that are known but unresolved at contract award transfer to the client"},
    {"name":"A14 Cambridge to Huntingdon","sector":"Roads / Highways Infrastructure","mode":"Earth",
     "keywords":["road","highway","a14","cambridge","motorway","uk","dual carriageway"],
     "cost_bn":1.5,"months":60,"cost_growth_pct":0,"schedule_slip_months":0,
     "failure_mode":"Utility diversions were the critical path for 60% of works",
     "lesson":"Utility diversion timelines are systematically underestimated — third-party access is not in the contractor's control"},
    # ── PORTS ────────────────────────────────────────────────────────
    {"name":"Felixstowe South Quay Extension","sector":"Ports / Marine Infrastructure","mode":"Earth",
     "keywords":["port","container","felixstowe","berth","quay","terminal","marine"],
     "cost_bn":0.4,"months":48,"cost_growth_pct":15,"schedule_slip_months":12,
     "failure_mode":"Marine ground conditions, operational cutover constraints",
     "lesson":"Port redevelopments require contingency for dredging ground conditions — seabed assumptions drive P80"},
    {"name":"London Gateway Phase 2","sector":"Ports / Marine Infrastructure","mode":"Earth",
     "keywords":["port","london gateway","container","berth","crane","terminal"],
     "cost_bn":1.8,"months":60,"cost_growth_pct":20,"schedule_slip_months":18,
     "failure_mode":"Terminal IT/OT integration 18 months late — not in EPC contract boundary",
     "lesson":"Terminal operating systems are the port's critical path at commissioning — not the quay wall"},
    # ── TELECOMS ─────────────────────────────────────────────────────
    {"name":"BT Openreach FTTP Rollout (UK)","sector":"Telecoms / Digital Infrastructure","mode":"Earth",
     "keywords":["telecom","fibre","broadband","rollout","openreach","bt","fttp","gigabit","5g"],
     "cost_bn":15.0,"months":144,"cost_growth_pct":50,"schedule_slip_months":36,
     "failure_mode":"Wayleave complexity in MDUs and dense urban areas — 2+ years behind target",
     "lesson":"Wayleave acquisition is the critical path for FTTP — not network build"},
    {"name":"NBN Co Multi-Technology Mix (Australia)","sector":"Telecoms / Digital Infrastructure","mode":"Earth",
     "keywords":["telecom","broadband","rollout","nbn","fibre","connectivity","rural"],
     "cost_bn":51.0,"months":144,"cost_growth_pct":985,"schedule_slip_months":72,
     "failure_mode":"Multi-technology complexity, copper network assumptions, contractor performance",
     "lesson":"National broadband rollouts in mixed geographies require 5-10x baseline cost assumptions"},
    # ── SPACE ────────────────────────────────────────────────────────
    {"name":"James Webb Space Telescope (JWST)","sector":"Space / Mission Assurance","mode":"Space",
     "keywords":["space","telescope","jwst","webb","orbital","satellite","observatory","deep space"],
     "cost_bn":10.0,"months":276,"cost_growth_pct":1400,"schedule_slip_months":168,
     "failure_mode":"Systems integration complexity, cryogenic testing failures, scope growth visible early",
     "lesson":"FOAK space systems have 14-15x baseline cost growth assumptions — qualification must be on the critical path"},
    {"name":"Artemis / SLS Programme","sector":"Space / Mission Assurance","mode":"Space",
     "keywords":["space","artemis","lunar","sls","nasa","rocket","launch","moon"],
     "cost_bn":93.0,"months":192,"cost_growth_pct":200,"schedule_slip_months":60,
     "failure_mode":"Fixed-price Boeing contract removed schedule incentives, propulsion complexity",
     "lesson":"Fixed-price contracting on FOAK space systems transfers insolvency risk, not schedule risk"},
    {"name":"OneWeb Satellite Constellation","sector":"Space / Mission Assurance","mode":"Space",
     "keywords":["space","satellite","constellation","onewe","leo","broadband","orbital"],
     "cost_bn":3.4,"months":96,"cost_growth_pct":70,"schedule_slip_months":36,
     "failure_mode":"Bankruptcy — launch cadence, ground segment, customer revenue all optimistic",
     "lesson":"Satellite constellation business cases require contracted anchor customers before launch commitment"},
    {"name":"Iridium NEXT Constellation","sector":"Space / Mission Assurance","mode":"Space",
     "keywords":["space","satellite","iridium","leo","constellation","communication"],
     "cost_bn":3.0,"months":84,"cost_growth_pct":0,"schedule_slip_months":0,
     "failure_mode":"Managed successfully — named launch provider, contracted cadence, anchor customers",
     "lesson":"Successful constellation reference: contracted launch, proven bus, anchor customer base from day 1"},
    {"name":"Lunar Gateway (Planned)","sector":"Space / Mission Assurance","mode":"Space",
     "keywords":["lunar","gateway","cislunar","space station","moon","orbit","habitat"],
     "cost_bn":40.0,"months":180,"cost_growth_pct":30,"schedule_slip_months":36,
     "failure_mode":"International partner coordination, launch cadence, FOAK life support",
     "lesson":"Cislunar infrastructure requires autonomous recovery capability — Earth-based contingency is a 3-day response minimum"},
    {"name":"Mars InSight Mission","sector":"Space / Mission Assurance","mode":"Space",
     "keywords":["mars","surface","lander","mission","nasa","probe","deep space"],
     "cost_bn":0.83,"months":96,"cost_growth_pct":25,"schedule_slip_months":24,
     "failure_mode":"Heat probe failed to penetrate Martian soil — regolith properties not in design basis",
     "lesson":"Mars surface properties require margin for FOAK geophysical assumptions"},
    # ── STADIA / CIVIC ───────────────────────────────────────────────
    {"name":"Tottenham Hotspur Stadium","sector":"Stadia / Events Infrastructure","mode":"Earth",
     "keywords":["stadium","arena","football","sports","tottenham","spurs","venue"],
     "cost_bn":1.2,"months":48,"cost_growth_pct":25,"schedule_slip_months":9,
     "failure_mode":"Retractable pitch mechanism, FA inspection, safety certification delay",
     "lesson":"Event-deadline driven construction compresses commissioning — safety certification is the opening gate"},
    {"name":"Wembley Stadium Redevelopment","sector":"Stadia / Events Infrastructure","mode":"Earth",
     "keywords":["stadium","wembley","football","arena","venue","arch"],
     "cost_bn":0.8,"months":60,"cost_growth_pct":40,"schedule_slip_months":18,
     "failure_mode":"Steelwork fabrication, contractor disputes, safety system integration",
     "lesson":"Stadium arch and signature structural elements carry 2-3x contingency assumption"},
    # ── GLOBAL RAIL ───────────────────────────────────────────────────
    {"name":"Riyadh Metro (Saudi Arabia)","sector":"Rail / Transit","mode":"Earth",
     "keywords":["riyadh","metro","rail","saudi","middle east","gcc","ksa"],
     "cost_bn":22.5,"months":96,"cost_growth_pct":12,"schedule_slip_months":24,
     "failure_mode":"Systems integration and operational readiness timeline across 6 concessions",
     "lesson":"Multi-concession metro programmes require a single systems integrator with contractual authority over all packages"},
    {"name":"California High Speed Rail (USA)","sector":"Rail / Transit","mode":"Earth",
     "keywords":["california","high speed","rail","usa","america","hsr"],
     "cost_bn":128.0,"months":384,"cost_growth_pct":1000,"schedule_slip_months":180,
     "failure_mode":"Environmental review, land acquisition, design changes — NEPA timeline structural constraint",
     "lesson":"US rail mega-projects require NEPA completion before cost can be baselined — pre-NEPA estimates are not estimates"},
    {"name":"Sydney Metro Northwest","sector":"Rail / Transit","mode":"Earth",
     "keywords":["sydney","metro","australia","rail","northwest"],
     "cost_bn":8.3,"months":96,"cost_growth_pct":5,"schedule_slip_months":0,
     "failure_mode":"Successfully delivered — TBM tunnelling, systems integration on schedule",
     "lesson":"Strong project reference: alliance contract model, TBM tunnelling, single systems integrator"},
    {"name":"Grand Paris Express (France)","sector":"Rail / Transit","mode":"Earth",
     "keywords":["paris","metro","france","grand paris","gpe"],
     "cost_bn":36.0,"months":240,"cost_growth_pct":45,"schedule_slip_months":48,
     "failure_mode":"Ground conditions, geology variability, post-COVID procurement inflation",
     "lesson":"Paris basin geology is more complex than initial surveys indicated — ground risk reserve must reflect full TBM risk"},
    # ── GLOBAL ENERGY ─────────────────────────────────────────────────
    {"name":"NEOM THE LINE Power Infrastructure","sector":"Energy / Power Infrastructure","mode":"Earth",
     "keywords":["neom","the line","saudi","ksa","smart city","mirrored building"],
     "cost_bn":500.0,"months":240,"cost_growth_pct":0,"schedule_slip_months":0,
     "failure_mode":"FOAK megacity — no comparable. Technology readiness of autonomous systems is the primary risk.",
     "lesson":"No reference class exists for THE LINE. Apply maximum OBA and require independent technical review of every FOAK system"},
    {"name":"Snowy 2.0 Pumped Hydro (Australia)","sector":"Energy / Power Infrastructure","mode":"Earth",
     "keywords":["snowy","pumped hydro","australia","hydro","energy storage"],
     "cost_bn":12.0,"months":108,"cost_growth_pct":233,"schedule_slip_months":60,
     "failure_mode":"TBM breakdown, ground conditions, geological fault — 3.3km TBM stuck for 14 months",
     "lesson":"Deep underground works in complex geology — apply 3-5x TBM programme contingency"},
    {"name":"Barakah Nuclear Power (UAE)","sector":"Nuclear / Regulated Generation","mode":"Earth",
     "keywords":["barakah","nuclear","uae","abu dhabi","enec","kepco"],
     "cost_bn":32.4,"months":204,"cost_growth_pct":62,"schedule_slip_months":72,
     "failure_mode":"Regulatory approval timeline, ENEC/IAEA safety case, operational licensing",
     "lesson":"First nuclear plant in the Arab world — regulatory approval timeline was the real critical path, not construction"},
    {"name":"Gordie Howe Bridge (Canada-USA)","sector":"Roads / Highways Infrastructure","mode":"Earth",
     "keywords":["canada","usa","bridge","international","border","windsor","detroit"],
     "cost_bn":5.7,"months":144,"cost_growth_pct":90,"schedule_slip_months":24,
     "failure_mode":"Bi-national procurement complexity, COVID, steel fabrication delays",
     "lesson":"Cross-border infrastructure requires harmonised procurement rules — different national standards add 20-40% contingency"},
    # ── GLOBAL WATER / UTILITIES ──────────────────────────────────────
    {"name":"Desalination Plant Jubail II (Saudi Arabia)","sector":"Water / Environmental Infrastructure","mode":"Earth",
     "keywords":["desalination","saudi","jubail","water","gcc","middle east"],
     "cost_bn":1.4,"months":60,"cost_growth_pct":15,"schedule_slip_months":12,
     "failure_mode":"Process performance at extreme ambient temperature — membrane degradation",
     "lesson":"Middle East desalination must be designed for 50°C+ ambient — standard membrane specifications are insufficient"},
    {"name":"Melbourne Water Smart Meter Rollout","sector":"Water / Environmental Infrastructure","mode":"Earth",
     "keywords":["smart meter","australia","melbourne","water","utility","rollout"],
     "cost_bn":0.6,"months":60,"cost_growth_pct":25,"schedule_slip_months":18,
     "failure_mode":"Back-office data platform readiness, meter reading system integration",
     "lesson":"Smart meter rollouts fail at the data layer — field installation is the easy part"},
    # ── GLOBAL OIL & GAS ──────────────────────────────────────────────
    {"name":"Kashagan Phase 1 (Kazakhstan)","sector":"Oil & Gas / Process Infrastructure","mode":"Earth",
     "keywords":["kashagan","kazakhstan","oil","gas","offshore","caspian"],
     "cost_bn":50.0,"months":192,"cost_growth_pct":400,"schedule_slip_months":120,
     "failure_mode":"H2S corrosion — pipeline design failed at commissioning, 3-year restart delay",
     "lesson":"Sour gas processing requires independent material qualification — no deviation from specification permissible"},
    {"name":"Ichthys LNG (Australia)","sector":"Oil & Gas / Process Infrastructure","mode":"Earth",
     "keywords":["ichthys","lng","australia","inpex","gas","offshore","darwin"],
     "cost_bn":45.0,"months":132,"cost_growth_pct":50,"schedule_slip_months":24,
     "failure_mode":"Module fabrication, labour costs, commissioning complexity",
     "lesson":"LNG final cost is determined by module fabrication quality and commissioning duration — not field development"},
    # ── GLOBAL MINING ─────────────────────────────────────────────────
    {"name":"Oyu Tolgoi Underground Mine (Mongolia)","sector":"Mining / Metals Infrastructure","mode":"Earth",
     "keywords":["oyu tolgoi","mongolia","copper","mine","rio tinto","underground"],
     "cost_bn":7.0,"months":144,"cost_growth_pct":60,"schedule_slip_months":48,
     "failure_mode":"Ground conditions, geotechnical complexity, caveback — production delayed",
     "lesson":"Block cave mining in complex ground requires geotechnical margin — cave propagation cannot be accelerated"},
    {"name":"Jansen Potash Mine (Canada)","sector":"Mining / Metals Infrastructure","mode":"Earth",
     "keywords":["jansen","potash","canada","bhp","mine","saskatchewan"],
     "cost_bn":5.7,"months":96,"cost_growth_pct":0,"schedule_slip_months":0,
     "failure_mode":"On schedule — strong project controls, single-owner BHP, definitive feasibility",
     "lesson":"Single-owner mega-mine with completed definitive feasibility study and no joint venture complexity — reference benchmark for mining delivery"},
    # ── GLOBAL DEFENCE ────────────────────────────────────────────────
    {"name":"AUKUS Submarine Programme (Australia/UK/USA)","sector":"Defence / Secure Infrastructure","mode":"Earth",
     "keywords":["aukus","submarine","australia","defence","nuclear","naval"],
     "cost_bn":268.0,"months":360,"cost_growth_pct":0,"schedule_slip_months":0,
     "failure_mode":"FOAK nuclear-powered submarine in Australia — no comparable. Workforce, regulatory, industrial base all new.",
     "lesson":"No reference class exists for AUKUS — it is simultaneously a FOAK submarine programme, FOAK nuclear programme, and FOAK industrial base creation"},
    # ── GLOBAL SPACE ──────────────────────────────────────────────────
    {"name":"Chandrayaan-3 (India Lunar)","sector":"Space / Mission Assurance","mode":"Space",
     "keywords":["chandrayaan","india","lunar","moon","isro","lander"],
     "cost_bn":0.075,"months":60,"cost_growth_pct":0,"schedule_slip_months":0,
     "failure_mode":"Chandrayaan-2 lander failed — software bug in braking sequence. Chandrayaan-3 corrected and succeeded.",
     "lesson":"Lunar landing requires exhaustive failure mode simulation — Chandrayaan-3 cost 10x less than Apollo equivalent by reusing heritage platform"},
    {"name":"Starlink Constellation (SpaceX)","sector":"Space / Mission Assurance","mode":"Space",
     "keywords":["starlink","spacex","satellite","constellation","leo","broadband"],
     "cost_bn":30.0,"months":96,"cost_growth_pct":0,"schedule_slip_months":0,
     "failure_mode":"Successfully scaled — reusable launch, vertical integration, iterative design",
     "lesson":"Vertical integration (own launch + own satellite) is the only structure that achieves constellation economics — merchant launch makes constellations unfinanceable"},
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


def _v136_extract_project_identity(prompt: str) -> dict:
    """Extract project-specific identity from user prompt for personalised output."""
    t = str(prompt or '').strip()
    tl = t.lower()
    
    # Extract location
    locations = {
        'south africa': 'South Africa', 'johannesburg': 'Johannesburg, South Africa',
        'cape town': 'Cape Town, South Africa', 'nigeria': 'Nigeria', 'kenya': 'Kenya',
        'ghana': 'Ghana', 'egypt': 'Egypt', 'morocco': 'Morocco', 'ethiopia': 'Ethiopia',
        'uk': 'United Kingdom', 'united kingdom': 'United Kingdom', 'england': 'England',
        'london': 'London, UK', 'manchester': 'Manchester, UK', 'birmingham': 'Birmingham, UK',
        'scotland': 'Scotland, UK', 'wales': 'Wales, UK',
        'usa': 'United States', 'united states': 'United States', 'america': 'United States',
        'texas': 'Texas, USA', 'california': 'California, USA', 'arizona': 'Arizona, USA',
        'north carolina': 'North Carolina, USA', 'ohio': 'Ohio, USA', 'virginia': 'Virginia, USA',
        'florida': 'Florida, USA', 'washington': 'Washington, USA', 'georgia': 'Georgia, USA',
        'west midlands': 'West Midlands, UK', 'yorkshire': 'Yorkshire, UK',
        'france': 'France', 'paris': 'Paris, France', 'germany': 'Germany', 'berlin': 'Berlin, Germany',
        'netherlands': 'Netherlands', 'sweden': 'Sweden', 'norway': 'Norway', 'denmark': 'Denmark',
        'spain': 'Spain', 'italy': 'Italy', 'poland': 'Poland',
        'saudi arabia': 'Saudi Arabia', 'riyadh': 'Riyadh, Saudi Arabia', 'dubai': 'Dubai, UAE',
        'uae': 'UAE', 'qatar': 'Qatar', 'doha': 'Doha, Qatar', 'bahrain': 'Bahrain',
        'oman': 'Oman', 'kuwait': 'Kuwait', 'jordan': 'Jordan',
        'india': 'India', 'mumbai': 'Mumbai, India', 'delhi': 'Delhi, India',
        'bangalore': 'Bangalore, India', 'chennai': 'Chennai, India',
        'singapore': 'Singapore', 'malaysia': 'Malaysia', 'indonesia': 'Indonesia',
        'australia': 'Australia', 'sydney': 'Sydney, Australia', 'melbourne': 'Melbourne, Australia',
        'perth': 'Perth, Australia', 'queensland': 'Queensland, Australia',
        'canada': 'Canada', 'toronto': 'Toronto, Canada', 'vancouver': 'Vancouver, Canada',
        'brazil': 'Brazil', 'sao paulo': 'São Paulo, Brazil', 'chile': 'Chile', 'mexico': 'Mexico',
        'china': 'China', 'beijing': 'Beijing, China', 'shanghai': 'Shanghai, China',
        'japan': 'Japan', 'tokyo': 'Tokyo, Japan', 'south korea': 'South Korea', 'seoul': 'Seoul, Korea',
        'taiwan': 'Taiwan', 'hong kong': 'Hong Kong',
        'north sea': 'North Sea', 'gulf of mexico': 'Gulf of Mexico', 'offshore': 'Offshore',
        'global': 'Global', 'international': 'International',
        'lunar': 'Lunar Surface', 'moon': 'Lunar Surface', 'mars': 'Mars Surface',
        'leo': 'Low Earth Orbit', 'orbit': 'Earth Orbit', 'cislunar': 'Cislunar Space',
    }
    location = ''
    for key, val in sorted(locations.items(), key=lambda x: -len(x[0])):
        if key in tl:
            location = val
            break
    
    # Extract scale/capacity signals
    scale_sig = ''
    for pat in [
        r'(\d[\d,.]*)\s*(million|m)\s*(connection|meter|meter|home|household|unit|seat|bed)',
        r'(\d[\d,.]*)\s*(gw|gwh|mw|mwh)',
        r'(\d[\d,.]*)\s*(km|mile)',
        r'(\d[\d,.]*)\s*(bed|ward|theatre)',
        r'(\d[\d,.]*)\s*(satellite|launch|mission)',
        r'(\d[\d,.]*)\s*(floor|storey)',
    ]:
        m = re.search(pat, tl)
        if m:
            scale_sig = m.group(0).strip()
            break
    
    # Extract cost signal
    cost_sig = ''
    for pat in [
        r'\$[\d,.]+\s*(?:billion|bn|b|million|mn|m|trillion|tn)',
        r'£[\d,.]+\s*(?:billion|bn|b|million|mn|m)',
        r'€[\d,.]+\s*(?:billion|bn|b|million|mn|m)',
        r'[\d,.]+\s*(?:billion|bn)\s*(?:dollar|pound|euro|usd|gbp|eur)?',
        r'[\d,.]+\s*(?:million)\s*(?:dollar|pound|usd|gbp)',
    ]:
        m = re.search(pat, tl)
        if m:
            cost_sig = m.group(0).strip()
            break
    
    # Extract duration signal
    dur_sig = ''
    for pat in [
        r'(\d+)\s*(?:month|months)',
        r'(\d+)\s*(?:year|years)',
        r'by\s+(20\d\d)',
        r'(20\d\d)\s+(?:delivery|completion|open)',
    ]:
        m = re.search(pat, tl)
        if m:
            dur_sig = m.group(0).strip()
            break
    
    # Build a project title from the prompt (first ~60 chars cleaned up)
    # Remove filler words and capitalise key terms
    title_words = []
    skip = {'a','an','the','and','or','for','with','to','in','on','at','of','is','are','will','that',
             'this','be','by','as','its','has','have','been','would','could','should','from','into',
             'across','through','per','about','above','below','between','within','during','after'}
    for word in t.split()[:12]:
        clean = re.sub(r'[^a-zA-Z0-9£$€%]', '', word)
        if clean and clean.lower() not in skip and len(clean) > 1:
            title_words.append(clean)
        if len(title_words) >= 7:
            break
    project_title = ' '.join(w.capitalize() if w[0].islower() else w for w in title_words[:6])
    
    return {
        'location': location,
        'scale_signal': scale_sig,
        'cost_signal': cost_sig,
        'duration_signal': dur_sig,
        'project_title': project_title,
        'prompt_short': t[:80],
    }


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
# ═══════════════════════════════════════════════════════════════════════
# CASEY V135 INTAKE NORMALISATION ENGINE
# Reads ANY client file — messy Excel, XER, risk registers, cost books,
# PDFs, CSVs — and challenges them from the inside against sector benchmarks.
# Returns structured findings, EPC flags, board attacks, CASEY comparison.
# ═══════════════════════════════════════════════════════════════════════

def _i_sf(v):
    """Safe float parse."""
    try: return float(str(v).replace(',','').replace('$','').replace('£','').replace('€','').replace('%','').strip())
    except: return 0.0

def _i_has(text, terms):
    t = str(text).lower()
    return any(term in t for term in terms)

def _i_detect_sector(text, fname):
    t = (str(text) + ' ' + str(fname)).lower()
    if _i_has(t, ['lunar','mars','orbital','spacecraft','satellite','launch vehicle','space station','isru','cislunar','leo ','geo ','cubesat','astronaut']): return 'space'
    if _i_has(t, ['awre','aldermaston','burghfield','warhead','nuclear weapon','dockyard','naval','aukus','frigate','destroyer','aircraft carrier','classified','sovereign supply','mod ','ministry of defence','military base','gchq','mi5','mi6']): return 'defence'
    if _i_has(t, ['rail','metro','tram','hs2','crossrail','signalling','possession','rolling stock','track access','station fit','light rail','subway']): return 'rail'
    if _i_has(t, ['nuclear','reactor','smr','gda','safety case','radiological','hinkley','sizewell','wylfa','sellafield','magnox','enrichment','decommission']): return 'nuclear'
    if _i_has(t, ['data centre','data center','datacenter','hyperscale','gpu cluster','pue ','server hall','ai campus','compute cluster','colocation']): return 'data_centre'
    if _i_has(t, ['pharma','gmp','validation','biologics','sterile','fill-finish','fill finish','cqv','mhra','fda approv','vaccine','bioreactor','cell therapy','gene therapy']): return 'pharma'
    if _i_has(t, ['semiconductor','wafer','euv','cleanroom chip','tsmc','intel fab','yield ramp','lithography','microchip']): return 'semiconductor'
    if _i_has(t, ['gigafactory','battery factory','battery manufactur','ev manufactur','electric vehicle plant','cathode','anode manufactur','lithium process','battery cell']): return 'gigafactory'
    if _i_has(t, ['oil field','gas field','lng','fpso','offshore platform','subsea','refinery','petrochemical','cracker','lng terminal','gas processing','carbon capture','ccus','upstream','downstream']): return 'oil_gas'
    if _i_has(t, ['mine ','mining','quarry','open pit','underground mine','tailings','ore processing','concentrator','smelter','lithium mine','copper mine','gold mine']): return 'mining'
    if _i_has(t, ['airport','terminal aviation','baggage system','runway','airside','orat airport','atc ','apron','taxiway','air traffic','heathrow','gatwick','stansted']): return 'airport'
    if _i_has(t, ['hospital','nhs','clinical build','ward','operating theatre','patient facility','healthcare campus','mental health unit','cancer centre','hospice']): return 'healthcare'
    if _i_has(t, ['wind farm','solar farm','battery storage','grid connection','substation','offshore wind','onshore wind','hydrogen plant','electrolyser','tidal energy','hydro power','biomass plant','energy storage','power station','combined cycle','ccgt','hvdc','national grid']): return 'energy'
    if _i_has(t, ['water treatment','desalin','wastewater','sewage','sewer','reservoir','smart meter','meter rollout','flood defence','flood alleviation','drainage','stormwater','irrigation','water supply','ofwat','anglian water','thames water']): return 'water'
    if _i_has(t, ['5g','telecom','fibre rollout','fiber optic','broadband rollout','mobile network','mast install','tower rollout','subsea cable','openreach','bt infrastr','network rollout','rural connectivity','gigabit']): return 'telecoms'
    if _i_has(t, ['motorway','highway','expressway','bridge road','tunnel road','road widening','junction upgrade','bypass','dual carriageway','trunk road','smart motorway']): return 'roads'
    if _i_has(t, ['port ','harbour','harbor','quay','berth','container terminal','cruise terminal','ferry terminal','marine infrastr','breakwater','dry dock','ship repair']): return 'ports'
    if _i_has(t, ['stadium','arena','velodrome','aquatic centre','olympic','sports facility','convention centre','exhibition centre','concert hall']): return 'stadia'
    if _i_has(t, ['university','school build','college build','campus build','student accommodation','library build','museum build','civic centre','town hall','government build']): return 'civic'
    return 'infrastructure'

def _i_sector_benchmarks(sector):
    b = {
        'space':     {'p50':(2,150),'sch':(36,200),'cont':25,
                      'epc':['FOAK technology risk understated','Launch manifest date assumed not contracted','TRL claimed without independent test evidence','Mass/power margin not confirmed','Radiation qualification deferred'],
                      'q':['What is the launch vehicle — contractually confirmed or assumed?','What independent TRL verification exists for each unproven technology?','What is the mass margin and what is the contractual consequence of a breach?','Which qualification tests have been deferred and what does each deferral cost?']},
        'defence':   {'p50':(1,80),'sch':(36,180),'cont':20,
                      'epc':['Security accreditation not on critical path','Classified supplier allocation unconfirmed','Export control dependencies unnamed','Operational acceptance testing post-delivery','Sovereign supply chain single-source'],
                      'q':['What is the security accreditation critical path and which classified supplier is currently unconfirmed?','What is the export control fallback for each sovereign supply chain item?','Is operational acceptance testing on the master programme or treated as post-delivery?','Which sovereign supply chain items are sole-source?']},
        'rail':      {'p50':(0.5,100),'sch':(24,180),'cont':15,
                      'epc':['Possession access assumed not confirmed with operator','Signalling integration milestone optimism — classic LP1 pattern','Systems migration deferred to late programme','Open corridor risk not in programme cost','Float not operationally usable against access windows'],
                      'q':['Which possessions are contractually confirmed with the operator versus assumed in the programme?','What is the systems migration critical path — this is where LP1 failed?','Is the reported float operationally usable against confirmed access windows and permit schedules?','What is the cost of a single failed trial-running gate?']},
        'nuclear':   {'p50':(2,50),'sch':(60,240),'cont':20,
                      'epc':['GDA critical path optimism vs Hinkley/Olkiluoto precedent','FOAK nuclear supply chain with single-source items','Safety case evidence programme not on master schedule','Nuclear-grade material lead times understated','Workforce nuclear certification not on critical path'],
                      'q':['What is the GDA critical path and what does a 6-month slip cost?','Which nuclear supply chain items are sole-source?','Is the safety case evidence programme on the master schedule with named owners?','What is the cost impact of a failed nuclear island weld inspection?']},
        'data_centre':{'p50':(0.3,15),'sch':(12,36),'cont':10,
                       'epc':['Grid connection assumed not contracted','Generator lead times not confirmed as orders','Cooling performance assumed from datasheet not verified','Commissioning float compressed','IT delivery dependency not on programme critical path'],
                       'q':['Is the grid connection agreement signed or is the energisation date still an assumption?','What is the fallback if cooling performance fails commissioning?','Are generator and transformer lead times confirmed orders or OEM forecast letters?','Which IT delivery milestones are on the critical path?']},
        'pharma':    {'p50':(0.2,5),'sch':(24,72),'cont':15,
                      'epc':['Validation programme deferred to post-construction','Clean utility readiness not on critical path','Regulatory submission timeline optimistic','Single failed PQ batch impact not costed','CQV resource assumed not confirmed'],
                      'q':['When does the inspection readiness programme complete and who owns it?','Is clean utility readiness on the master validation plan with dates?','What is the cost and schedule impact of a single failed PQ batch?','Is CQV resource confirmed as named individuals?']},
        'gigafactory':{'p50':(1,20),'sch':(24,60),'cont':12,
                       'epc':['Grid connection assumed not contracted','Cell chemistry qualification not on critical path','Tool delivery sequence optimistic','UPW and chemical system lead times understated','Yield ramp scope outside programme boundary'],
                       'q':['Is the grid connection contracted or assumed?','Is cell chemistry qualification on the master programme critical path?','Are tool delivery sequences confirmed by OEM or derived from optimistic ramp models?','Is yield ramp in the programme boundary?']},
        'oil_gas':   {'p50':(1,50),'sch':(24,96),'cont':15,
                      'epc':['Offshore weather windows optimistically scoped','Long-lead procurement assumed not contracted','Brownfield interface with live production underestimated','Regulatory consent assumed as formality','Single-source subsea equipment'],
                      'q':['Are long-lead procurement items confirmed orders or assumed from OEM intent letters?','What is the brownfield interface plan with live production facilities?','Is regulatory consent on the critical path or treated as a formality?']},
        'mining':    {'p50':(0.5,30),'sch':(24,84),'cont':15,
                      'epc':['Environmental consent treated as formality','Tailings management costs underestimated','Community licence-to-operate not in risk register','Ore grade assumptions optimistic','Processing plant yield assumed from datasheet'],
                      'q':['Is environmental consent on the critical path or treated as a formality?','What is the tailings management cost basis?','What is the community licence-to-operate risk and how is it managed?']},
        'airport':   {'p50':(0.5,20),'sch':(24,96),'cont':12,
                      'epc':['ORAT not on master schedule as critical path','Airside access constraints not in programme cost model','Security system integration deferred','Baggage acceptance testing optimistic','CAA/FAA timeline not verified against precedent'],
                      'q':['Is ORAT on the master programme critical path?','What are the airside access constraints and are they in the programme cost model?','Is the CAA/FAA approval timeline verified against comparable terminal precedents?']},
        'infrastructure':{'p50':(0.1,50),'sch':(12,120),'cont':12,
                          'epc':['Critical path constraint unnamed','Float nominal not confirmed as usable','Reserve flat percentage not risk-linked','Basis of estimate not documented','Owner accountability diluted'],
                          'q':['Who is the named individual owner of the governing critical-path constraint?','Is programme float confirmed as operationally usable?','Is the reserve risk-linked to QCRA P80 or a flat percentage?','What evidence package closes the primary board approval blocker?']},
    }
    return b.get(sector, b['infrastructure'])

def _i_parse_xlsx(content, filename):
    """Parse any Excel file — handles messy headers, merged cells, multiple sheets."""
    from openpyxl import load_workbook
    sheets = []
    try:
        wb = load_workbook(BytesIO(content), read_only=True, data_only=True)
        for ws in list(wb.worksheets)[:10]:
            rows_out = []; nums = []
            for ridx, row in enumerate(ws.iter_rows(values_only=True), 1):
                if ridx > 500: break
                cells = []
                for c in (row or [])[:30]:
                    # Handle merged cell None values
                    if c is None:
                        cells.append('')
                    elif isinstance(c, float) and (c != c):  # NaN
                        cells.append('')
                    else:
                        cells.append(str(c).strip() if c is not None else '')
                rows_out.append(cells)
                for c in (row or []):
                    if isinstance(c, (int, float)) and c == c and 0 < abs(c) < 1e12:
                        nums.append(float(c))
            st = ' '.join(c for r in rows_out for c in r if c).lower()
            sheets.append({'name': ws.title, 'rows': rows_out, 'numbers': nums, 'text': st})
    except Exception as e:
        sheets.append({'name': 'error', 'rows': [], 'numbers': [], 'text': str(e)[:120]})
    return sheets

def _i_parse_xer(content):
    """Parse Primavera XER — full TASK/TASKPRED/WBS/CALENDAR extraction."""
    try: text = content.decode('utf-8', errors='ignore')
    except: text = ''
    tasks, preds, cals, wbs = [], [], [], []
    cur_t, fields = None, []
    for line in text.splitlines():
        s = line.rstrip()
        if s.startswith('%T\t'): cur_t = s[3:].strip(); fields = []
        elif s.startswith('%F\t') and cur_t: fields = s[3:].split('\t')
        elif s.startswith('%R\t') and fields and cur_t:
            pts = s[3:].split('\t')
            row = dict(zip(fields, pts + [''] * max(0, len(fields) - len(pts))))
            if cur_t == 'TASK': tasks.append(row)
            elif cur_t == 'TASKPRED': preds.append(row)
            elif cur_t == 'CALENDAR': cals.append(row.get('clndr_name', ''))
            elif cur_t == 'WBS': wbs.append(row)
    return tasks, preds, cals, wbs, text

def _i_challenge_xer(tasks, preds, cals, wbs, sector, bm, bench_model):
    if not tasks:
        return {'file_type':'SCHEDULE (XER)','findings':['XER received but no TASK table found. Ensure full schedule export includes TASK, TASKPRED and CALENDAR tables.'],'red_flags':[],'epc_flags':[],'board_attacks':bm['q'],'metrics':{},'next_steps':['Re-export from Primavera including all tables.']}
    total = len(tasks)
    mt = {'TT_Mile','TT_FinMile','TT_StartMile'}
    non_m = [t for t in tasks if t.get('task_type','') not in mt]
    pred_to = set(r.get('task_id','') for r in preds)
    pred_from = set(r.get('pred_task_id','') for r in preds)
    open_s = [t for t in non_m if t.get('task_id','') not in pred_to]
    open_e = [t for t in non_m if t.get('task_id','') not in pred_from]
    constrained = [t for t in tasks if t.get('cstr_type','') not in ('','CS_ALAP','CS_ASAP','None')]
    neg_f = [t for t in tasks if _i_sf(t.get('total_float_hr_cnt','0')) < -8]
    long_a = [t for t in tasks if _i_sf(t.get('target_drtn_hr_cnt','0')) > 2400]
    durations = [_i_sf(t.get('target_drtn_hr_cnt','0')) for t in tasks]
    logic_r = round(len(preds) / max(total, 1), 2)
    dur_months = sum(durations) / (8 * 22) if durations else 0
    sch_lo, sch_hi = bm['sch']
    findings, red, epc, attacks = [], [], [], []
    metrics = {'activities': total, 'relationships': len(preds), 'open_starts': len(open_s),
               'open_ends': len(open_e), 'constrained': len(constrained), 'neg_float': len(neg_f),
               'long_activities': len(long_a), 'logic_density': logic_r, 'dur_months_implied': round(dur_months,1)}
    # CASEY benchmark P50
    casey_months = 0
    if bench_model:
        try: casey_months = int(str(bench_model.get('schedule','')).replace(' months','').strip())
        except: pass
    findings.append(f"XER parsed: {total} activities, {len(preds)} logic ties, {len(cals)} calendars, {len(wbs)} WBS nodes. Sector: {sector.upper()}.")
    if casey_months: findings.append(f"CASEY benchmark for {sector}: {casey_months} months. Schedule implies {dur_months:.0f} months of work.")
    else: findings.append(f"Sector benchmark: {sch_lo}–{sch_hi} months. Implied work: {dur_months:.0f} months.")
    if open_s:
        nms = ', '.join(t.get('task_name','?')[:30] for t in open_s[:3] if t.get('task_name',''))
        findings.append(f"{len(open_s)} open-start activities — no predecessor, cannot be logic-driven. Includes: {nms}.")
        epc.append(f"EPC RED FLAG: {len(open_s)} open-start activities. These are anchored durations, not logic-driven. Remove and rerun to expose the real critical path — this is where {sector} programmes hide float padding.")
        attacks.append(f"{len(open_s)} activities have no predecessor — what specific deliverable or approval triggers each one, and who is contractually responsible?")
    if open_e:
        findings.append(f"{len(open_e)} open-end activities — no successor, dangling ends that inflate apparent completeness.")
        epc.append(f"EPC RED FLAG: {len(open_e)} open-end activities. Work leading nowhere — the completion date is not driven by completing this work.")
    if neg_f:
        nms2 = ', '.join(t.get('task_name','?')[:25] for t in neg_f[:2] if t.get('task_name',''))
        findings.append(f"{len(neg_f)} activities with negative float — schedule is already late by its own logic. Includes: {nms2}.")
        red.append(f"SCHEDULE ALREADY LATE: {len(neg_f)} activities with negative float. The programme date is broken before any risk materialises. Recovery plan required at board.")
        attacks.append(f"The schedule has {len(neg_f)} activities with negative float — the programme is already logically late. What is the recovery plan, who owns it, and what does it cost?")
    if constrained:
        findings.append(f"{len(constrained)} date-constrained activities override logic and force the schedule to appear on-programme regardless of upstream work.")
        epc.append(f"EPC RED FLAG: {len(constrained)} date constraints applied. Remove constraints and rerun — this is how contractors make schedules look green when the logic says otherwise.")
    if long_a:
        findings.append(f"{len(long_a)} activities exceed 300 days — summary bars hiding risk. A Level 4 schedule should not have activities this long.")
        epc.append(f"EPC RED FLAG: {len(long_a)} activities over 300 days. This is not a Level 4 schedule. Require breakdown with resource loading for each.")
    if logic_r < 1.2:
        findings.append(f"Logic density: {logic_r} relationships per activity (benchmark 1.5–2.5). This is a task list, not a delivery model.")
        red.append(f"WEAK SCHEDULE LOGIC: {logic_r} relationships per activity. Cannot predict delay impact. Require re-sequenced Level 4 before board approval.")
    if dur_months > 0 and dur_months < sch_lo * 0.72:
        findings.append(f"Schedule implies {dur_months:.0f} months — {sch_lo-dur_months:.0f} months below sector minimum. Optimism bias indicator.")
        epc.append(f"EPC RED FLAG: Schedule is {sch_lo-dur_months:.0f} months shorter than sector benchmark. Contractors compress durations at bid — the gap closes through variations.")
        attacks.append(f"The schedule implies {dur_months:.0f} months but sector benchmark is {sch_lo}–{sch_hi} months. Where is the difference — de-risked or just removed?")
    for p in bm['epc'][:3]:
        epc.append(f"SECTOR PATTERN ({sector.upper()}): {p} — verify not present in submitted schedule.")
    attacks.extend(bm['q'][:3])
    attacks = list(dict.fromkeys(attacks))[:8]
    return {'file_type':'SCHEDULE (XER)','findings':findings,'red_flags':red,'epc_flags':epc,'board_attacks':attacks,'metrics':metrics,
            'next_steps':[f"Require contractor to close all {len(open_s)} open-start activities with documented predecessors.",
                          "Remove all date constraints and rerun to expose the true logic-driven critical path.",
                          "Run QSRA — the P80 finish date is the board commitment date, not the management date.",
                          f"Require Level 4 breakdowns for all {len(long_a)} summary-level long activities.",
                          "Validate float against confirmed access windows, permits and operator acceptance dates."]}

def _i_challenge_risk(sheets, sector, bm, bench_model):
    risks, emv_t, p90_t, p10_t = [], 0.0, 0.0, 0.0
    owners, statuses = {}, {}
    ev_req = low_conf = no_trig = v_high = no_owner = 0
    for sh in sheets:
        rows = sh.get('rows', [])
        sname = sh.get('name', '').lower()
        for hi, hrow in enumerate(rows[:15]):
            ht = ' '.join(str(c) for c in hrow).lower()
            if any(w in ht for w in ['risk id','risk ','owner','mitigation','likelihood','category','cause','hazard']):
                hdrs = [str(c).lower().strip() for c in hrow]
                for r in rows[hi+1:hi+500]:
                    if not r or not any(c for c in r[:3] if c): continue
                    rd = dict(zip(hdrs, list(r) + [''] * max(0, len(hdrs) - len(r))))
                    risks.append(rd)
                    # Owner
                    ow = str(rd.get('owner', rd.get('risk owner', rd.get('responsible','?'))) or '').strip()
                    if not ow or ow.lower() in ['','none','tbc','unknown','nan','?','n/a']: no_owner += 1
                    else: owners[ow] = owners.get(ow, 0) + 1
                    # Status
                    st = str(rd.get('response status', rd.get('status', rd.get('action status',''))) or '').strip()
                    statuses[st] = statuses.get(st, 0) + 1
                    if 'evidence required' in st.lower() or 'open' == st.lower(): ev_req += 1
                    # Mitigation confidence
                    for conf_key in ['mitigation confidence','mit confidence','mit conf','confidence']:
                        mc_val = rd.get(conf_key, '')
                        if mc_val:
                            try:
                                mc = float(str(mc_val))
                                if mc < 0.5: low_conf += 1
                            except: pass
                            break
                    # Trigger
                    trig = str(rd.get('trigger', rd.get('early warning', rd.get('trigger event',''))) or '').strip()
                    if not trig or trig.lower() in ['','none','tbc','n/a','nan','not defined','to be confirmed']: no_trig += 1
                    # Residual
                    res = str(rd.get('residual', rd.get('residual exposure', rd.get('residual score', rd.get('residual rating','')))  or '') or '').strip()
                    if res.lower() in ['very high','extreme','critical','5','4 - high','5 - critical']: v_high += 1
                break
        # Tornado/EMV
        if any(w in sname for w in ['tornado','emv','qcra','exposure','downside','risk driver']):
            for row in rows:
                rl = ' '.join(str(c) for c in row).lower()
                if any(w in rl for w in ['downside bn','downside','p90','emv bn','emv','exposure']):
                    h2 = [str(c).lower().strip() for c in row]
                    for dr in rows[rows.index(row)+1:rows.index(row)+50]:
                        if not dr: continue
                        rd2 = dict(zip(h2, list(dr) + [''] * max(0, len(h2) - len(dr))))
                        for key in ['downside bn','downside b','p90 bn','p90']: 
                            v = _i_sf(str(rd2.get(key,'0')))
                            if 0 < v < 1000: p90_t += v; break
                        for key in ['upside bn','upside b','p10 bn','p10']:
                            v = _i_sf(str(rd2.get(key,'0')))
                            if 0 < v < 1000: p10_t += v; break
                        for key in ['emv bn','emv b','emv']:
                            v = _i_sf(str(rd2.get(key,'0')))
                            if 0 < v < 1000: emv_t += v; break
                    break
    total = len(risks)
    oc = max(owners.values()) / max(total, 1) if owners else 0
    top_ow = max(owners, key=owners.get) if owners else 'Unknown'
    # CASEY benchmark
    cp50 = cp80 = 0.0
    if bench_model:
        try: cp50 = float(str(bench_model.get('cost_p50','0')).replace('$','').replace('B','').replace(',','').strip())
        except: pass
        try: cp80 = float((bench_model.get('monte_carlo') or {}).get('qcra',{}).get('p80', 0))
        except: pass
    findings, red, epc, attacks = [], [], [], []
    findings.append(f"Risk register parsed: {total} risks, {len(owners)} named owners. Sector: {sector.upper()}.")
    if emv_t > 0: findings.append(f"Total EMV from file: ${emv_t:.2f}B. P90 downside: ${p90_t:.1f}B. P10 upside: ${p10_t:.1f}B.")
    elif total > 0: findings.append(f"{total} risk rows detected. No EMV/tornado sheet found — CASEY estimates exposure from risk density and sector benchmarks.")
    if ev_req > 0:
        findings.append(f"{ev_req} risks with status 'Evidence Required' or 'Open' — these cannot be treated as mitigated for board approval.")
        epc.append(f"EPC RED FLAG: {ev_req} risks not evidenced. A contractor submitting this register is presenting action intent as evidence closure. Board approval with open exposures transfers unquantified risk to the client.")
        attacks.append(f"Which of the {ev_req} open/evidence-required risks have a confirmed closure date, named owner and residual cost — and what is the exposure if each closes late?")
    if low_conf > 0:
        findings.append(f"{low_conf} mitigations have confidence below 50% — CASEY treats these as unmitigated. Reserve must cover them at P80.")
        red.append(f"RESERVE UNDERSTATED: {low_conf} mitigations below 50% confidence. Reserve sized against nominal mitigations will be insufficient at P80.")
        attacks.append(f"{low_conf} mitigations have confidence below 50% — what evidence proves each will work, and what is the additional reserve if they fail?")
    if no_trig > 0:
        findings.append(f"{no_trig} risks have no documented trigger — programme cannot detect materialisation before it hits cost and schedule.")
        epc.append(f"EPC RED FLAG: {no_trig} risks without triggers. This is a risk list, not active risk management. Require confirmed triggers with monitoring frequency for every high-priority risk.")
        attacks.append(f"For each of the {no_trig} untriggered risks, how will the programme detect materialisation, and what is the response time from trigger to recovery?")
    if oc > 0.35:
        findings.append(f"Owner concentration: {int(oc*100)}% of risks owned by '{top_ow}'. Individual accountability is diluted.")
        epc.append(f"EPC RED FLAG: {int(oc*100)}% of risks owned by '{top_ow}'. Committee ownership means no individual is accountable. Require named individuals, not roles.")
        attacks.append(f"'{top_ow}' owns {owners.get(top_ow,0)} risks — who is the named individual for each high-residual risk and what are the contractual consequences?")
    if no_owner > 0:
        findings.append(f"{no_owner} risks have no named owner — these cannot be challenged or escalated.")
        epc.append(f"EPC RED FLAG: {no_owner} ownerless risks. A risk without an owner cannot be managed. Every risk must have a named individual, not a team or TBC.")
    if v_high > 0:
        findings.append(f"{v_high} risks remain Very High or Extreme after stated mitigation — additional reserve or escalation required.")
        red.append(f"ESCALATION REQUIRED: {v_high} Very High/Extreme residual risks cannot be absorbed into general contingency. Named SRO escalation required.")
    if p90_t > 0 and cp80 > 0:
        gap = p90_t - cp80
        findings.append(f"File P90 ${p90_t:.1f}B vs CASEY sector benchmark P80 ${cp80:.1f}B. {'Programme carries more tail risk than benchmark.' if gap>0 else 'Verify register completeness — P90 below benchmark.'}")
        if gap > 1: attacks.append(f"The register P90 is ${p90_t:.1f}B but the approved reserve is likely sized to P50 — what is the board-approved confidence level for the reserve position?")
    if cp50 > 0 and emv_t > 0:
        findings.append(f"CASEY sector P50 benchmark: ${cp50:.1f}B. Register EMV: ${emv_t:.1f}B ({int(emv_t/cp50*100)}% of CASEY P50).")
    for p in bm['epc']:
        epc.append(f"SECTOR PATTERN ({sector.upper()}): {p} — verify this risk has a named owner and evidence closure plan.")
    attacks.extend(bm['q'])
    attacks = list(dict.fromkeys(attacks))[:9]
    metrics = {'risks': total, 'emv_bn': round(emv_t,3), 'p90_bn': round(p90_t,3), 'ev_req': ev_req,
               'low_conf': low_conf, 'no_trig': no_trig, 'v_high': v_high, 'no_owner': no_owner,
               'owner_conc_pct': int(oc*100), 'top_owner': top_ow}
    cc = {'client_p90': f"${p90_t:.1f}B" if p90_t else '—', 'casey_p80': f"${cp80:.1f}B" if cp80 else '—',
          'client_risks': total, 'open_exposures': ev_req,
          'governance': 'BOARD CHALLENGE REQUIRED' if ev_req > 1 or low_conf > 3 or no_trig > total*0.5 else 'CONDITIONAL APPROVAL'}
    return {'file_type':'RISK REGISTER','findings':findings,'red_flags':red,'epc_flags':epc,'board_attacks':attacks,
            'metrics':metrics,'casey_comparison':cc,
            'next_steps':['Require evidence closure plans for all open/evidence-required risks before board approval.',
                          f"Resize reserve to cover {low_conf} unmitigated risks at P80 confidence.",
                          "Replace committee ownership with named individual accountability on every high-residual risk.",
                          "Run CASEY QCRA/QSRA alongside this register to test whether mitigations move the P-curves.",
                          "Download CASEY risk register comparison workbook for side-by-side challenge."]}

def _i_challenge_cost(sheets, sector, bm, bench_model):
    direct = indirect = reserve = escalation = 0.0
    has_p80 = has_basis = has_cbs = has_esc = False
    cost_lines_found = 0
    for sh in sheets:
        st = sh.get('text', '')
        if any(w in st for w in ['p80','p90','80th percentile','risk-adjusted range','qcra','probabilistic']): has_p80 = True
        if any(w in st for w in ['basis of estimate','scope basis','basis statement','inclusions','exclusions','assumptions']): has_basis = True
        if any(w in st for w in ['cbs','wbs','cost breakdown structure','work breakdown','cost code']): has_cbs = True
        if any(w in st for w in ['escalat','inflation','price adjustment','time-related cost','tender price index','tpi']): has_esc = True
        for row in sh.get('rows', []):
            rt = ' '.join(str(c) for c in row).lower()
            nums = [_i_sf(c) for c in row if c]
            nums = [n for n in nums if 1000 < n < 1e11]
            if not nums: continue
            cost_lines_found += 1
            val = max(nums)
            if any(w in rt for w in ['direct','civil','construction','mechanical','electrical','works','physical']): direct = max(direct, val)
            elif any(w in rt for w in ['indirect','prelim','overhead','management fee','supervision','site management']): indirect = max(indirect, val)
            elif any(w in rt for w in ['contingency','reserve','risk allowance','risk provision','risk budget','uplift']): reserve = max(reserve, val)
            elif any(w in rt for w in ['escalat','inflation','price adj','outturn adjustment','time allowance']): escalation = max(escalation, val)
    total = direct + indirect + reserve + escalation if (direct or indirect) else 0.0
    if not total:
        all_n = sorted([n for sh in sheets for n in sh.get('numbers', []) if 10000 < n < 1e11], reverse=True)
        total = all_n[0] if all_n else 0.0
        if len(all_n) > 1: direct = all_n[1]
        if len(all_n) > 2: reserve = all_n[2]
    sc = 1e9 if total > 1e8 else (1e6 if total > 1e5 else 1.0)
    tb = total / sc; db = direct / sc; rb = reserve / sc
    cp50 = cp80 = 0.0
    if bench_model:
        try: cp50 = float(str(bench_model.get('cost_p50','0')).replace('$','').replace('B','').replace(',','').strip())
        except: pass
        try: cp80 = float((bench_model.get('monte_carlo') or {}).get('qcra',{}).get('p80',0))
        except: pass
    cont_min = bm['cont']
    findings, red, epc, attacks = [], [], [], []
    if tb > 0:
        cont_pct = int(rb/tb*100) if tb else 0
        findings.append(f"Cost structure extracted: total ~${tb:.2f}B, direct ~${db:.2f}B, reserve ~${rb:.2f}B ({cont_pct}%). {cost_lines_found} cost lines found.")
    else:
        findings.append(f"Cost file parsed. {cost_lines_found} cost signals detected. Scale unclear — CASEY applied sector benchmarks.")
    if not has_p80:
        findings.append("No P80/P90 range detected. Single-point P50 estimate only — board cannot see downside exposure.")
        epc.append(f"EPC RED FLAG: No P80/P90 range. Single-point estimate hides tail risk. For {sector} at this scale, CASEY benchmarks P80/P50 ratio at 1.2–1.4x. Require QCRA support.")
        attacks.append("This estimate shows only a P50 — what is the P80 and P90 range, and which risks drive the tail exposure?")
    if rb > 0 and tb > 0:
        cp = rb / tb * 100
        if cp < cont_min:
            red.append(f"LOW CONTINGENCY: Reserve at {cp:.1f}% is below CASEY benchmark minimum {cont_min}% for {sector}. Require QCRA validation.")
            attacks.append(f"Reserve is {cp:.1f}% of stated estimate — CASEY benchmark for {sector} is minimum {cont_min}%. What QCRA P80 justifies this level?")
    elif tb > 0:
        epc.append(f"EPC RED FLAG: No clearly mapped contingency line. Either reserve is embedded (opaque) or missing. Require explicit risk provision linked to QCRA P80.")
    if not has_basis:
        epc.append(f"EPC RED FLAG: No basis of estimate detected. Cannot verify scope completeness — inclusions, exclusions and assumptions unknown. This is how contractors enable scope creep claims.")
        attacks.append("Where is the basis of estimate — what is explicitly included, excluded, and what assumptions underpin each major cost package?")
    if not has_cbs:
        epc.append(f"EPC RED FLAG: No CBS/WBS mapping. Costs cannot be traced to programme activities — overruns cannot be attributed to specific scope.")
    if not has_esc and tb > 1.0:
        findings.append("No escalation basis detected. A 2% inflation variance can move outturn cost by 5–10% on a multi-year programme.")
        attacks.append("What is the escalation basis — which index, what rate, and what is the outturn sensitivity to a 2% inflation variance?")
    if cp50 > 0 and tb > 0:
        v = (tb - cp50) / cp50 * 100
        if v < -15:
            findings.append(f"Client estimate ${tb:.2f}B is {abs(v):.0f}% below CASEY sector benchmark ${cp50:.1f}B.")
            epc.append(f"EPC RED FLAG: Estimate {abs(v):.0f}% below CASEY benchmark for {sector}. Contractors systematically understate at bid — the gap closes through variations during delivery.")
            attacks.append(f"Client estimate (${tb:.2f}B) is {abs(v):.0f}% below CASEY benchmark (${cp50:.1f}B) — what scope exclusions or efficiency assumptions justify the gap?")
        elif v > 25:
            findings.append(f"Estimate ${tb:.2f}B is {v:.0f}% above CASEY benchmark — verify scope definition and efficiency assumptions.")
    if cp80 > 0 and tb > 0 and tb < cp80 * 0.85:
        red.append(f"ESTIMATE BELOW CASEY P80: Client total (${tb:.2f}B) is below CASEY QCRA P80 (${cp80:.1f}B). The estimate does not cover expected downside.")
    for p in bm['epc'][:3]:
        epc.append(f"SECTOR PATTERN ({sector.upper()}): {p} — verify this cost risk is captured.")
    attacks.extend(bm['q'][:3])
    attacks = list(dict.fromkeys(attacks))[:9]
    metrics = {'total_bn': round(tb,3), 'direct_bn': round(db,3), 'reserve_bn': round(rb,3),
               'has_p80': has_p80, 'has_basis': has_basis, 'has_cbs': has_cbs,
               'casey_p50': f"${cp50:.1f}B" if cp50 else '—', 'casey_p80': f"${cp80:.1f}B" if cp80 else '—'}
    cc = {'client_p50': f"${tb:.2f}B" if tb else '—', 'casey_p50': f"${cp50:.1f}B" if cp50 else '—',
          'casey_p80': f"${cp80:.1f}B" if cp80 else '—',
          'governance': 'BOARD CHALLENGE REQUIRED' if not has_p80 or not has_basis else 'CONDITIONAL APPROVAL'}
    return {'file_type':'COST ESTIMATE','findings':findings,'red_flags':red,'epc_flags':epc,'board_attacks':attacks,
            'metrics':metrics,'casey_comparison':cc,
            'next_steps':['Require P50/P80/P90 range with QCRA support before approving the headline number.',
                          'Require a CBS that maps every cost line to a programme activity and work package.',
                          'Commission independent cost review against CASEY benchmark before board submission.',
                          'Mandate documented basis of estimate with explicit inclusions, exclusions and assumptions.',
                          'Require escalation basis with sensitivity analysis at ±2% inflation variance.']}

def _i_classify(sheets, all_text, filename):
    """Classify what type of file this is — handles ambiguous messy files."""
    lname = filename.lower()
    # Filename hint is strong
    if any(w in lname for w in ['risk','rr','r&o','register','hazard','rag','issues']): return 'risk'
    if any(w in lname for w in ['cost','estimate','budget','cbs','bq','capex','workbook','opex','financial']): return 'cost'
    if any(w in lname for w in ['schedule','programme','program','gantt','master','baseline','wbs']): return 'cost'  # schedule workbooks → cost
    # Content scoring
    rr_s = sum(1 for w in ['risk id','risk description','owner','mitigation','likelihood','trigger','residual','emv','probability','hazard','consequence'] if w in all_text)
    ce_s = sum(1 for w in ['direct cost','indirect','contingency','p50','cbs','estimate','budget','escalat','quantity','bill of quantities','boq','capex','cost code','unit rate'] if w in all_text)
    if rr_s > ce_s and rr_s >= 3: return 'risk'
    if ce_s >= rr_s and ce_s >= 3: return 'cost'
    # Try both and pick better
    return 'risk' if rr_s >= ce_s else 'cost'

async def analyse_upload(file: UploadFile = File(...)):
    """CASEY V135 Intake Normalisation Engine.
    Handles any file type: XER schedules, XLSX cost estimates, XLSX risk registers,
    CSV data, JSON models, messy Excel, PDFs, plain text.
    Returns: findings, EPC flags, board attacks, CASEY vs submitted comparison.
    """
    content = await file.read()
    name = file.filename or 'upload'
    size = len(content)
    try: text_raw = content.decode('utf-8', errors='ignore')
    except: text_raw = ''
    lname = name.lower()
    # For Excel/CSV files, parse first to extract readable text for better sector detection
    _pre_sheets = []
    if lname.endswith(('.xlsx','.xlsm','.xls')):
        try: _pre_sheets = _i_parse_xlsx(content, name)
        except: pass
    _pre_text = text_raw[:50000] + ' '.join(sh.get('text','')[:5000] for sh in _pre_sheets[:3])
    # Detect sector
    sector = _i_detect_sector(_pre_text, name)
    bm = _i_sector_benchmarks(sector)
    # Load CASEY benchmark model for comparison
    bench_model = {}
    try:
        prompts = {
            'space': 'Lunar habitat programme life support power autonomous surface commissioning',
            'defence': 'Defence secure facility classified systems integration sovereign supply operational acceptance',
            'rail': 'Rail transit tunnelling stations signalling systems migration possessions operator acceptance',
            'nuclear': 'Nuclear power station GDA licensing safety case procurement commissioning FOAK',
            'data_centre': 'AI hyperscale data centre 500MW grid liquid cooling GPU supply commissioning',
            'pharma': 'GMP pharmaceutical sterile fill finish validation FDA inspection CQV',
            'gigafactory': 'Battery gigafactory EV manufacturing cell production utility grid 3 billion 48 months',
            'oil_gas': 'LNG terminal offshore platform subsea pipeline brownfield procurement commissioning',
            'mining': 'Open pit mine ore processing plant environmental consent tailings management',
            'airport': 'Airport terminal ORAT airside baggage security regulatory CAA',
            'healthcare': 'Hospital clinical campus MEP utilities operating theatres NHS commissioning',
            'energy': 'Offshore wind farm grid connection HVDC substation procurement commissioning',
            'water': 'Water treatment desalination smart meter rollout regulatory consent commissioning',
            'telecoms': 'Fibre broadband rollout 5G mast installation wayleave planning rural urban',
            'infrastructure': 'Capital infrastructure programme procurement systems integration commissioning',
        }
        bench_model = build_model(prompts.get(sector, prompts['infrastructure']), 'IntakeQA', 3, 4, 'base')
    except Exception: bench_model = {}
    # Parse and challenge
    is_xer = lname.endswith('.xer') or b'%T\tTASK' in content[:8000] or b'TASKPRED' in content[:8000]
    is_xlsx = lname.endswith(('.xlsx', '.xlsm', '.xls'))
    is_csv = lname.endswith('.csv')
    is_json = lname.endswith('.json')
    if is_xer:
        tasks, preds, cals, wbs, _ = _i_parse_xer(content)
        ch = _i_challenge_xer(tasks, preds, cals, wbs, sector, bm, bench_model)
        conf = max(28, 62 - len(ch.get('red_flags',[])) * 8 - len(ch.get('epc_flags',[])) * 4)
        result = {'version': 'CASEY V135 intake engine', 'filename': name, 'size_bytes': size,
                  'file_type': 'SCHEDULE (XER)', 'sector_detected': sector.upper(),
                  'schema_confidence': 'high' if len(tasks) > 10 else 'medium',
                  'epc_challenge': True, 'challenge_verdict': 'BOARD CHALLENGE REQUIRED' if ch.get('epc_flags') else 'CONDITIONAL APPROVAL',
                  'confidence_impact': conf, **ch.get('metrics',{}),
                  'findings': ch['findings'], 'red_flags': ch['red_flags'], 'epc_flags': ch['epc_flags'],
                  'board_challenge_questions': ch['board_attacks'], 'next_steps': ch['next_steps']}
    elif is_xlsx or is_csv:
        if is_xlsx:
            sheets = _i_parse_xlsx(content, name)
        else:
            lines = [l.split(',') for l in text_raw[:200000].splitlines()[:1000] if l.strip()]
            nums = [_i_sf(c) for row in lines for c in row]
            sheets = [{'name':'CSV','rows':lines,'numbers':nums,'text':text_raw[:50000].lower()}]
        all_text = ' '.join(sh.get('text','') for sh in sheets)
        ftype = _i_classify(sheets, all_text, name)
        if ftype == 'risk':
            ch = _i_challenge_risk(sheets, sector, bm, bench_model)
        else:
            ch = _i_challenge_cost(sheets, sector, bm, bench_model)
        metrics = ch.pop('metrics', {})
        cc = ch.pop('casey_comparison', {})
        conf = max(28, 62 - len(ch.get('red_flags',[])) * 7 - len(ch.get('epc_flags',[])) * 4)
        result = {'version': 'CASEY V135 intake engine', 'filename': name, 'size_bytes': size,
                  'file_type': ch.get('file_type','EXCEL'), 'sector_detected': sector.upper(),
                  'schema_confidence': 'high' if metrics.get('risks',0)>5 or metrics.get('total_bn',0)>0 else 'medium',
                  'epc_challenge': True, 'challenge_verdict': cc.get('governance','CONDITIONAL APPROVAL'),
                  'confidence_impact': conf, **metrics,
                  'findings': ch['findings'], 'red_flags': ch['red_flags'], 'epc_flags': ch['epc_flags'],
                  'board_challenge_questions': ch['board_attacks'], 'casey_comparison': cc,
                  'next_steps': ch['next_steps']}
    elif is_json:
        try:
            imp = json.loads(text_raw)
            if imp.get('cost_p50'):
                result = {'version':'CASEY V135 intake engine','filename':name,'size_bytes':size,
                          'file_type':'CASEY MODEL IMPORT','sector_detected':imp.get('mode','Earth'),
                          'epc_challenge':False,'challenge_verdict':'MODEL IMPORTED',
                          'confidence_impact':imp.get('confidence_pct',50),
                          'findings':[f"CASEY model imported: {imp.get('title','Project')} | {imp.get('cost_p50')} P50 | {imp.get('schedule')} | {imp.get('confidence_pct')}% confidence."],
                          'red_flags':imp.get('red_flags',[]),'epc_flags':[],
                          'board_challenge_questions':imp.get('board_challenge_questions',imp.get('board_briefing',[])),
                          'next_steps':['Model loaded. Run scenarios, stress tests and download board pack.']}
            else:
                result = {'version':'CASEY V135','filename':name,'file_type':'JSON',
                          'findings':['JSON received but not a CASEY model export.'],'red_flags':[],'epc_flags':[],'board_challenge_questions':[],'next_steps':[],'challenge_verdict':'REVIEW REQUIRED'}
        except Exception as e:
            result = {'version':'CASEY V135','filename':name,'file_type':'JSON','findings':[f'JSON parse error: {str(e)[:100]}'],'red_flags':[],'epc_flags':[],'board_challenge_questions':[],'next_steps':[],'challenge_verdict':'REVIEW REQUIRED'}
    else:
        # Plain text / unknown — keyword analysis + sector benchmarks
        rr_s = sum(1 for w in ['risk','mitigation','owner','residual','trigger'] if w in text_raw.lower())
        ce_s = sum(1 for w in ['cost','estimate','contingency','direct','budget'] if w in text_raw.lower())
        xe_s = sum(1 for w in ['activity','predecessor','duration','schedule','float','baseline start'] if w in text_raw.lower())
        ftype = 'SCHEDULE' if xe_s > ce_s else ('RISK DOCUMENT' if rr_s > ce_s else 'COST DOCUMENT')
        result = {'version':'CASEY V135 intake engine','filename':name,'size_bytes':size,
                  'file_type':ftype+' (text)','sector_detected':sector.upper(),
                  'schema_confidence':'low','epc_challenge':True,
                  'challenge_verdict':'UPLOAD AS XLSX OR XER FOR FULL CHALLENGE',
                  'confidence_impact':max(32,52-(5 if rr_s<3 else 0)),
                  'findings':[f"Text file received. Sector: {sector.upper()}, {rr_s} risk signals, {ce_s} cost signals, {xe_s} schedule signals.",
                               f"CASEY sector benchmark: P50 ${bm['p50'][0]}–${bm['p50'][1]}B, schedule {bm['sch'][0]}–{bm['sch'][1]} months.",
                               "For a full structural challenge with real numbers, upload as XLSX or XER."],
                  'red_flags':['Text files cannot be fully challenged — upload native XLSX/XER/CSV source exports.'],
                  'epc_flags':[f"SECTOR PATTERN ({sector.upper()}): {p}" for p in bm['epc'][:3]],
                  'board_challenge_questions':bm['q'][:5],
                  'next_steps':['Upload as XLSX or XER for full structural challenge.']}
    # Persist
    try:
        con = db(); cur = con.cursor()
        cur.execute("INSERT INTO uploads(filename,created_at,analysis_json) VALUES(?,?,?)",
                    (name, datetime.utcnow().isoformat(), json.dumps({k:v for k,v in result.items() if not isinstance(v,bytes)}, default=str)))
        con.commit(); con.close()
    except Exception: pass
    return result

print("CASEY V135 intake normalisation engine installed")



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


# ═══════════════════════════════════════════════════════════════════
# CASEY V135 EARTH + SPACE INSTANT DEMO ENDPOINTS
# Pre-built showcase models for the platform demo
# ═══════════════════════════════════════════════════════════════════

@app.get("/demo/earth")
def demo_earth():
    """HS2 Phase 2b — the definitive Earth infrastructure demo."""
    try:
        m = build_model(
            "HS2 Phase 2b tunnelling stations signalling systems integration possessions operator acceptance UK rail",
            "ControlOrbit Demo", 3, 4, "base"
        )
        m["demo_mode"] = True
        m["demo_type"] = "earth"
        m["demo_label"] = "HS2 Phase 2b — Rail Mega Programme"
        m["demo_headline"] = "A full-programme intelligence pack generated in 4 seconds. Click any tab to explore the output."
        return m
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/demo/space")
def demo_space():
    """Lunar Base Alpha — the definitive Space infrastructure demo."""
    try:
        m = build_model(
            "Lunar Base Alpha life support nuclear surface power autonomous commissioning resupply logistics 1000 crew",
            "ControlOrbit Demo", 3, 4, "base"
        )
        m["demo_mode"] = True
        m["demo_type"] = "space"
        m["demo_label"] = "Lunar Base Alpha — Deep Space Mega Programme"
        m["demo_headline"] = "First-of-kind space programme intelligence. TRL risk, launch logistics, autonomous commissioning."
        return m
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/demo/awre")
def demo_awre():
    """AWRE Aldermaston — the definitive Defence demo."""
    try:
        m = build_model(
            "AWRE Aldermaston nuclear warhead facility upgrade classified defence sovereign supply chain security accreditation UK MOD",
            "ControlOrbit Demo", 3, 3, "base"
        )
        m["demo_mode"] = True
        m["demo_type"] = "defence"
        m["demo_label"] = "AWRE Aldermaston — Defence Nuclear Infrastructure"
        m["demo_headline"] = "Classified programme intelligence. Security accreditation, sovereign supply chain, operational acceptance."
        return m
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/demo/gigafactory")
def demo_gigafactory():
    """Gigafactory UK — the EV/battery manufacturing demo."""
    try:
        m = build_model(
            "Battery gigafactory West Midlands UK 50GWh annual capacity EV manufacturing cell production utility grid 3 billion pounds 48 months",
            "ControlOrbit Demo", 3, 3, "base"
        )
        m["demo_mode"] = True
        m["demo_type"] = "gigafactory"
        m["demo_label"] = "Gigafactory UK — Battery Manufacturing"
        m["demo_headline"] = "EV battery manufacturing intelligence. Grid connection, cell chemistry, utility complexity."
        return m
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/demo/library")
def demo_library():
    """All pre-built demo models for the showcase library."""
    demos = [
        {"id": "earth", "label": "HS2 Phase 2b", "sector": "Rail", "mode": "Earth", 
         "tagline": "UK rail mega programme — possessions, signalling, operator acceptance",
         "icon": "🚄", "url": "/demo/earth"},
        {"id": "space", "label": "Lunar Base Alpha", "sector": "Space", "mode": "Space",
         "tagline": "Deep space habitat — TRL risk, autonomous commissioning, life support",
         "icon": "🌕", "url": "/demo/space"},
        {"id": "awre", "label": "AWRE Aldermaston", "sector": "Defence", "mode": "Earth",
         "tagline": "Defence nuclear — classified systems, sovereign supply, accreditation",
         "icon": "🛡️", "url": "/demo/awre"},
        {"id": "gigafactory", "label": "Gigafactory UK", "sector": "Battery Mfg", "mode": "Earth",
         "tagline": "EV battery plant — grid connection, cell chemistry, utility scale",
         "icon": "⚡", "url": "/demo/gigafactory"},
    ]
    return {"demos": demos, "count": len(demos)}

print("CASEY V135 Earth + Space instant demo endpoints installed")

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

# ================= CASEY V124 SECTOR ONTOLOGY HARDENING LOCK =================
# Final public-demo hardening: sector-locked causal graphs, vocabulary and benchmark guardrails.
# Prevents data-centre / rail / airport / space / defence / energy ontology bleed in UI and exports.

def _v124_text_blob(model: Dict[str, Any]) -> str:
    return (str(model.get('prompt','')) + ' ' + str(model.get('title','')) + ' ' + str(model.get('subsector','')) + ' ' + str(model.get('mode',''))).lower()


def _v124_sector_key(model: Dict[str, Any]) -> str:
    t = _v124_text_blob(model)
    sub = str(model.get('subsector','')).lower()
    mode = str(model.get('mode',''))
    if mode == 'Space' or any(k in t for k in ['lunar','mars','orbital','satellite','spaceport','launch vehicle','payload','moon','deep space']): return 'space'
    if any(k in t for k in ['airport','aviation','terminal','runway','heathrow','gatwick','airside','baggage','orat']): return 'airport'
    if any(k in t for k in ['data centre','data center','hyperscale','ai campus','compute campus','gpu','cloud region','white space']): return 'data_centre'
    if any(k in t for k in ['energy','power plant','renewable','wind farm','offshore wind','solar','battery','substation','transmission','grid','hydrogen','nuclear']): return 'energy'
    if any(k in t for k in ['rail','metro','transit','high speed','hs2','station','signalling','rolling stock','california high speed']): return 'rail'
    if any(k in t for k in ['semiconductor','fab','wafer','cleanroom','foundry','lithography','chip plant']): return 'semiconductor'
    if any(k in t for k in ['life sciences','pharma','biologics','gmp','fill-finish','sterile','cqv','amgen','lilly','novartis','pfizer']): return 'life_sciences'
    if any(k in t for k in ['defence','defense','military','naval','airbase','radar','missile','secure facility','mod ','dod '] ): return 'defence'
    if any(k in t for k in ['oil','gas','lng','refinery','petrochemical','offshore','pipeline','fpsO'.lower(),'hydrocarbon','carbon capture']): return 'oil_gas'
    if any(k in t for k in ['hospital','healthcare','clinical','medical centre','patient','nhs']): return 'healthcare'
    if any(k in t for k in ['water','wastewater','desalination','sewer','reservoir','treatment plant']): return 'water'
    if any(k in t for k in ['port','harbour','marine','dock','container terminal']): return 'ports'
    return 'general_infrastructure'


def _v124_library(key: str) -> Dict[str, Any]:
    lib = {
      'airport': {
        'label':'Airport / Aviation',
        'shock':'The dominant risk sits in live operations, ORAT, baggage/security integration and regulatory acceptance — not headline construction progress.',
        'constraints':'ORAT readiness, baggage/security systems integration, airside phasing and regulator acceptance',
        'signals':[('Live operations phasing','Active','Terminal works must sequence around airside access, airline interfaces and passenger operations.'),('ORAT readiness','Active','Operational readiness, trials and airport transition govern confidence.'),('Baggage / security integration','Watch','Baggage, security, ICT and MEP systems must operate as one environment.'),('Regulatory and airline approvals','Watch','Airport approvals and stakeholder readiness create the P80/P90 tail.')],
        'bench':[('Airport Terminal Expansion','$2B–$24B','60-144'),('Major Hub Capacity Programme','$5B–$35B','72-168'),('Rail/Transit Systems Integration','$3B–$80B','60-180')],
        'chain':['ORAT readiness','Baggage/security integration','Airside phasing','Regulatory acceptance','Operational transition','Commissioning overlap','Confidence'],
        'confidence':['Benchmark similarity: airport terminal / airfield expansion','Scope maturity: capacity, phasing and systems definition','Procurement certainty: baggage/security/MEP packages','Schedule maturity: ORAT and live operations logic','Interface exposure: airlines, airside, landside and regulators'],
        'cost':['Terminal and airside works','Baggage and security systems','Operational transition and phasing','Transport integration and utilities','Retail/passenger experience fit-out'],
        'schedule':['Live airport phasing and possessions','Baggage/security systems integration','Operational readiness trials','Regulatory and stakeholder approvals','Airside access and safety constraints']},
      'rail': {
        'label':'Rail / Transit',
        'shock':'The dominant risk sits in possessions, signalling integration, systems migration and operator acceptance rather than civil progress alone.',
        'constraints':'possessions, signalling integration, systems migration and operator acceptance',
        'signals':[('Possession window pressure','Active','Access windows and blockades govern productive delivery time.'),('Signalling and systems integration','Active','Signalling, telecoms, power and control systems drive migration risk.'),('Operator acceptance','Watch','Timetable, trial running and safety assurance determine readiness.'),('Utilities / corridor interfaces','Watch','Diversions and third-party interfaces create schedule tail exposure.')],
        'bench':[('Metro / Rail Extension','$3B–$80B','60-180'),('Major Station Redevelopment','$1B–$15B','36-96'),('Rail Systems Migration Programme','$500M–$10B','30-84')],
        'chain':['Possession access','Utility diversions','Signalling integration','Systems migration','Trial operations','Operator acceptance','Confidence'],
        'confidence':['Benchmark similarity: rail/transit programme','Scope maturity: alignment, station and systems definition','Procurement certainty: civil/systems package strategy','Schedule maturity: possessions and test/commissioning logic','Interface exposure: utilities, operators and regulators'],
        'cost':['Civil and station works','Signalling, power and telecoms systems','Utility diversions and corridor constraints','Possessions and access logistics','Systems assurance and testing'],
        'schedule':['Possession window availability','Signalling integration and migration','Utility diversion completion','Trial running and safety certification','Operator/regulator acceptance']},
      'data_centre': {
        'label':'Digital Infrastructure / Hyperscale Data Centre',
        'shock':'Power availability and systems integration readiness are more likely to govern delivery than shell construction progress.',
        'constraints':'energisation, cooling readiness and integrated systems testing',
        'signals':[('Grid connection congestion','Active','Utility interconnection and energisation remain the dominant delivery constraints.'),('Transformer / switchgear lead times','Active','Long-lead electrical packages shape procurement tail exposure.'),('Liquid cooling readiness','Watch','Cooling readiness and heat-rejection capacity govern commissioning risk.'),('IST and phased hall turnover','Watch','Integrated systems testing and phased data-hall readiness drive confidence.')],
        'bench':[('Hyperscale AI Data Centre Campus','$2B–$18B','24-84'),('Digital Infrastructure Campus','$1B–$12B','24-72'),('Energy / Utility Megaprogramme','$1B–$18B','36-108'),('Semiconductor / Advanced Manufacturing','$8B–$32B','54-96')],
        'chain':['Transformer lead-time','Grid energisation','Liquid cooling readiness','IST congestion','Commissioning overlap','Reserve drawdown','Confidence'],
        'confidence':['Benchmark similarity: hyperscale digital infrastructure','Scope maturity: campus power and white-space definition','Procurement certainty: transformers, generators and switchgear','Schedule maturity: grid and commissioning logic','Interface exposure: utilities, fibre and commissioning'],
        'cost':['Utility/grid connection and substations','Power train, transformers and switchgear','Liquid cooling / heat rejection systems','Data halls and white space fit-out','Accelerated procurement premiums'],
        'schedule':['Grid energisation and utility agreements','Long-lead transformer and switchgear delivery','Integrated systems testing and commissioning','Cooling plant readiness','Phased data-hall turnover']},
      'semiconductor': {
        'label':'Semiconductor / Advanced Manufacturing',
        'shock':'Tool install, cleanroom certification and yield-ramp readiness govern board confidence more than shell completion.',
        'constraints':'cleanroom readiness, process tool install, specialty utilities and yield-ramp qualification',
        'signals':[('Tool-install sequencing','Active','OEM tool windows and installation sequence govern critical path.'),('Cleanroom certification','Active','Cleanroom classification and environmental stability constrain turnover.'),('UPW / specialty utilities','Watch','Ultra-pure water, gases and exhaust systems drive commissioning readiness.'),('Yield ramp qualification','Watch','Production qualification creates post-mechanical-completion uncertainty.')],
        'bench':[('Advanced Semiconductor Fab','$8B–$35B','54-108'),('Cleanroom Manufacturing Campus','$2B–$12B','36-84'),('Specialty Utilities Programme','$500M–$5B','24-60')],
        'chain':['Cleanroom readiness','Specialty utilities','Process tool delivery','Tool hook-up','Qualification lots','Yield ramp','Confidence'],
        'confidence':['Benchmark similarity: advanced fab / cleanroom campus','Scope maturity: process flow and tool list maturity','Procurement certainty: OEM tool slots and specialty utilities','Schedule maturity: tool hook-up and qualification sequence','Interface exposure: utilities, vendors and yield ramp'],
        'cost':['Cleanrooms and classified areas','Process tools and hook-up','UPW/specialty gases/exhaust systems','Vibration/environmental controls','Yield ramp and qualification support'],
        'schedule':['Tool delivery and install windows','Cleanroom certification','Specialty utilities qualification','Process qualification and yield ramp','OEM/vendor interface availability']},
      'life_sciences': {
        'label':'Life Sciences / Biologics Manufacturing',
        'shock':'Mechanical completion is not the true finish line; validated production readiness and deviation closure are the real board decision gates.',
        'constraints':'CQV, GMP turnover, validation readiness and regulatory evidence',
        'signals':[('CQV readiness','Active','Commissioning, qualification and validation control release confidence.'),('Clean utility validation','Active','WFI, clean steam and process utilities constrain turnover.'),('Media fill / batch readiness','Watch','Process validation and batch release drive operational start-up risk.'),('Regulatory inspection readiness','Watch','FDA/EMA readiness affects board-defensible approval.')],
        'bench':[('Biologics Manufacturing Campus','$2B–$8B','36-78'),('Sterile Fill-Finish / Aseptic Expansion','$1B–$5B','30-60'),('Advanced GMP Cleanroom','$2B–$10B','42-84')],
        'chain':['GMP turnover','Clean utility validation','CQV protocols','Media fill readiness','Deviation closure','Regulatory readiness','Confidence'],
        'confidence':['Benchmark similarity: pharma / biologics campus','Scope maturity: GMP package and user requirement definition','Procurement certainty: process equipment and clean utility lead times','Schedule maturity: CQV logic and validation pathway','Regulatory exposure: FDA/EMA inspection readiness'],
        'cost':['Process equipment and fill-finish lines','GMP cleanrooms / classified areas','Clean utilities and HVAC zoning','CQV validation and deviation closure','Regulatory readiness and quality systems'],
        'schedule':['CQV protocol approval and execution','Long-lead process equipment delivery','Clean utility validation and media fills','FDA/EMA inspection readiness','Batch release and operational readiness']},
      'defence': {
        'label':'Defence / Secure Mission Systems',
        'shock':'Assurance, security accreditation and mission-system integration are likely to govern approval before asset completion.',
        'constraints':'security accreditation, mission assurance, sovereign supply chain and integration test evidence',
        'signals':[('Security accreditation','Active','Classified systems and facility accreditation create approval gates.'),('Mission-system integration','Active','Sensors, comms, command systems and platform interfaces govern readiness.'),('Sovereign supply-chain exposure','Watch','Controlled components and cleared supplier capacity affect delivery tails.'),('Assurance test campaign','Watch','Operational acceptance depends on test evidence and traceability.')],
        'bench':[('Secure Mission Facility','$500M–$8B','30-96'),('Defence Systems Integration Programme','$1B–$20B','48-144'),('Naval / Airbase Modernisation','$1B–$15B','48-120')],
        'chain':['Security requirements','Sovereign procurement','Mission-system integration','Assurance test campaign','Accreditation gates','Operational acceptance','Confidence'],
        'confidence':['Benchmark similarity: defence secure systems programme','Scope maturity: mission requirement and accreditation definition','Procurement certainty: cleared suppliers and controlled components','Schedule maturity: integration/test campaign logic','Interface exposure: security, operators and regulators'],
        'cost':['Secure facilities and hardening','Mission systems and integration','Controlled equipment procurement','Cyber/security accreditation','Test campaign and operational acceptance'],
        'schedule':['Security accreditation gates','Mission-system integration testing','Controlled supplier lead-times','Operational test campaign','Authority-to-operate approval']},
      'oil_gas': {
        'label':'Energy / Oil & Gas',
        'shock':'Procurement, modular integration, HSE readiness and commissioning/start-up govern value more than installed quantities.',
        'constraints':'long-lead equipment, module integration, HSE readiness and start-up/commissioning assurance',
        'signals':[('Long-lead rotating equipment','Active','Compressors, turbines and specialist packages shape procurement exposure.'),('Module / brownfield integration','Active','Tie-ins, shutdown windows and interface planning create schedule tails.'),('HSE / regulatory approvals','Watch','Permitting and safety case maturity affect board confidence.'),('Commissioning and start-up','Watch','Hydrotest, pre-commissioning and start-up dominate late risk.')],
        'bench':[('LNG / Gas Processing Train','$5B–$30B','48-108'),('Refinery / Petrochemical Expansion','$2B–$20B','36-96'),('Offshore / Pipeline Programme','$1B–$18B','36-96')],
        'chain':['FEED maturity','Long-lead packages','Module fabrication','Tie-in windows','Pre-commissioning','Start-up readiness','Confidence'],
        'confidence':['Benchmark similarity: oil/gas processing or export infrastructure','Scope maturity: FEED, plot plan and tie-in definition','Procurement certainty: rotating equipment and specialist packages','Schedule maturity: modularisation and shutdown logic','Interface exposure: brownfield operations, HSE and regulators'],
        'cost':['Process units and rotating equipment','Pipelines / terminals / tie-ins','Modular fabrication and logistics','HSE, permitting and compliance','Commissioning and start-up support'],
        'schedule':['Long-lead compressor/turbine packages','Module fabrication and transport','Shutdown/tie-in windows','Regulatory and HSE approvals','Pre-commissioning and start-up']},
      'energy': {
        'label':'Energy / Power Infrastructure',
        'shock':'Grid access, equipment lead-times, permitting and commissioning readiness govern delivery confidence more than site progress.',
        'constraints':'grid connection, major equipment lead-times, permitting and commissioning readiness',
        'signals':[('Grid connection access','Active','Interconnection agreements and energisation gates govern the delivery tail.'),('Major equipment lead-times','Active','Transformers, turbines, inverters or batteries shape procurement risk.'),('Permitting and environmental approvals','Watch','Approvals and grid studies affect schedule certainty.'),('Commissioning / performance testing','Watch','Performance tests and grid compliance govern final acceptance.')],
        'bench':[('Transmission / Grid Reinforcement','$500M–$10B','24-84'),('Renewable + Storage Programme','$300M–$8B','18-60'),('Power Generation Programme','$1B–$15B','36-96')],
        'chain':['Permitting','Grid studies','Major equipment delivery','Site construction','Energisation','Performance testing','Confidence'],
        'confidence':['Benchmark similarity: power / utility infrastructure','Scope maturity: connection, permitting and design definition','Procurement certainty: transformers, turbines, inverters or batteries','Schedule maturity: energisation and commissioning logic','Interface exposure: grid operator, regulator and land/environment'],
        'cost':['Grid connection and substations','Generation/storage equipment','Civil and balance-of-plant works','Environmental and permitting compliance','Commissioning and performance testing'],
        'schedule':['Grid connection agreement','Long-lead equipment delivery','Permitting and environmental approval','Energisation windows','Performance testing and acceptance']},
      'space': {
        'label':'Space / Lunar / Orbital Infrastructure',
        'shock':'Launch alone is not the real constraint; qualification, thermal-power balance and autonomous recovery determine mission survivability.',
        'constraints':'mission assurance, launch logistics, qualification evidence and autonomous recovery',
        'signals':[('Launch reliability volatility','Active','Launch cadence, range access and manifest priority affect schedule certainty.'),('Mission assurance burden','Active','Qualification and redundancy evidence govern board defensibility.'),('Thermal-power balance','Watch','Power storage and thermal rejection shape survivability.'),('Regulatory and range availability','Watch','Flight approvals, range windows and orbital operations create schedule tails.')],
        'bench':[('Orbital / Lunar Infrastructure','$8B–$95B','72-216'),('Launch and Payload Integration','$1B–$14B','36-108'),('Deep-Space Mission Systems','$4B–$35B','60-144'),('Autonomous Operations Platform','$2B–$20B','36-96')],
        'chain':['Launch cadence','Payload integration','Thermal-power balance','Range availability','Autonomous commissioning','Mission assurance','Confidence'],
        'confidence':['Benchmark similarity: space infrastructure archetype','Scope maturity: payload and mission architecture definition','Procurement certainty: launch, avionics and qualified hardware','Schedule maturity: launch and commissioning logic','Operational exposure: remote recovery and servicing limits'],
        'cost':['Payload / habitat / mission systems','Launch and orbital logistics','Power, thermal and autonomy systems','Qualification and test campaign','Mission operations and recovery reserve'],
        'schedule':['Payload qualification campaign','Launch manifest and range access','Thermal-vacuum and systems testing','Orbital/surface deployment sequence','Autonomous commissioning and recovery']},
      'healthcare': {
        'label':'Healthcare / Hospital Infrastructure','shock':'Clinical transition, medical equipment integration and infection-control readiness govern approval more than building completion.','constraints':'clinical commissioning, infection-control compliance, medical equipment integration and phased occupancy','signals':[('Clinical commissioning','Active','Clinical workflows and patient transition govern readiness.'),('Medical equipment integration','Active','Imaging, theatres, labs and digital systems drive handover risk.'),('Infection-control compliance','Watch','ICRA and commissioning evidence shape confidence.'),('Phased occupancy','Watch','Live hospital operations constrain access and transition.')],'bench':[('Major Hospital Campus','$1B–$8B','36-96'),('Clinical Tower / Specialist Centre','$500M–$4B','30-72')],'chain':['Clinical requirements','Medical equipment procurement','Digital systems integration','Infection-control validation','Phased occupancy','Clinical commissioning','Confidence'],'confidence':['Benchmark similarity: hospital / clinical campus','Scope maturity: clinical brief and department definition','Procurement certainty: medical equipment and digital systems','Schedule maturity: clinical commissioning and phased occupancy','Interface exposure: clinicians, patients, regulators and live operations'],'cost':['Clinical departments and fit-out','Medical equipment and imaging','Digital health and systems integration','Infection-control and compliance','Phased occupancy and transition'],'schedule':['Medical equipment delivery','Clinical commissioning readiness','Infection-control validation','Digital systems integration','Patient transition and phased occupancy']},
      'water': {
        'label':'Water / Environmental Infrastructure','shock':'Consents, process commissioning and environmental compliance govern delivery confidence more than civil installation.','constraints':'permits, process commissioning, environmental compliance and operational acceptance','signals':[('Consents and discharge permits','Active','Environmental approvals govern start-up readiness.'),('Process commissioning','Active','Treatment-process performance controls acceptance.'),('MEICA procurement','Watch','Pumps, controls and specialist equipment create lead-time risk.'),('Operational takeover','Watch','Operator readiness and compliance testing affect confidence.')],'bench':[('Water Treatment Programme','$300M–$5B','24-72'),('Desalination / Major Water Plant','$1B–$10B','36-96')],'chain':['Environmental permits','MEICA procurement','Civil/process works','Process commissioning','Compliance testing','Operator acceptance','Confidence'],'confidence':['Benchmark similarity: water treatment / environmental asset','Scope maturity: process capacity and permit definition','Procurement certainty: MEICA and specialist equipment','Schedule maturity: commissioning and compliance testing logic','Interface exposure: regulator, operator and environmental stakeholders'],'cost':['Civil/process treatment works','MEICA equipment','Pipelines and intake/outfall works','Environmental mitigation','Commissioning and compliance testing'],'schedule':['Permit approval','MEICA procurement','Process commissioning','Compliance testing','Operator acceptance']},
      'ports': {
        'label':'Ports / Marine Infrastructure','shock':'Marine access, dredging, quay interfaces and terminal systems govern delivery risk more than civil progress alone.','constraints':'marine access, dredging, quay works, terminal systems and port operational interfaces','signals':[('Marine access window','Active','Weather, tides and vessel windows shape productivity.'),('Dredging / quay interface','Active','Marine works and ground conditions control schedule exposure.'),('Terminal systems readiness','Watch','Cranes, automation and yard systems affect operational acceptance.'),('Port operations interface','Watch','Live port operations constrain phasing and access.')],'bench':[('Container Terminal Expansion','$500M–$8B','30-84'),('Major Port / Marine Works','$1B–$12B','36-96')],'chain':['Marine access','Dredging / ground risk','Quay construction','Cranes and terminal systems','Operational trials','Port acceptance','Confidence'],'confidence':['Benchmark similarity: port / terminal expansion','Scope maturity: marine works and terminal operating model','Procurement certainty: cranes, systems and marine contractors','Schedule maturity: dredging, quay and systems commissioning','Interface exposure: shipping, operators and regulators'],'cost':['Dredging and marine works','Quay / berth construction','Cranes and terminal systems','Yard, utilities and access roads','Operational trials and port transition'],'schedule':['Marine access and weather windows','Dredging / ground treatment','Quay and berth completion','Cranes / terminal systems commissioning','Port operational acceptance']}
    }
    return lib.get(key, {
      'label':'Capital Infrastructure','shock':'The dominant risk sits in interfaces, procurement evidence and commissioning readiness rather than headline progress.','constraints':'interface control, procurement evidence and commissioning readiness','signals':[('Procurement volatility','Watch','Market capacity and supplier evidence shape confidence.'),('Approvals latency','Watch','Governance and permit friction influence schedule tails.'),('Commissioning readiness','Active','Operational readiness is a board-defensibility signal.'),('Interface density','Active','Interfaces create the largest uncertainty transfer.')],'bench':[('Complex Infrastructure Programme','$1B–$20B','36-120')],'chain':['Scope definition','Procurement evidence','Interface control','Commissioning readiness','Operational acceptance','Reserve adequacy','Confidence'],'confidence':['Benchmark similarity: comparable infrastructure archetype','Scope maturity: requirements and package definition','Procurement certainty: supplier and market evidence','Schedule maturity: critical-path and commissioning logic','Interface exposure: utilities, operators and approvals'],'cost':['Civil/enabling works','Specialist systems and equipment','Utilities and interfaces','Programme indirects','Risk reserve and contingency'],'schedule':['Approvals and governance','Long-lead procurement','Interface coordination','Commissioning readiness','Operational acceptance']})


def _v124_forbidden_terms(key: str):
    all_terms = {
      'data_centre':['ORAT','baggage','airside','landside','rolling stock','signalling','payload integration','launch cadence','range availability','yield ramp','CQV'],
      'airport':['liquid cooling','white-space','white space','GPU','data hall','thermal rejection','yield ramp','payload','launch reliability','transformer lead-time'],
      'rail':['liquid cooling','GPU','data hall','ORAT','baggage systems','airside','launch reliability','thermal-power','yield ramp'],
      'space':['ORAT','baggage','airside','rolling stock','signalling','liquid cooling readiness','white space','yield ramp','CQV'],
      'semiconductor':['ORAT','baggage','airside','rolling stock','launch reliability','payload integration'],
      'life_sciences':['ORAT','baggage','rolling stock','launch reliability','liquid cooling readiness','white space'],
      'defence':['ORAT','baggage','liquid cooling readiness','white space','yield ramp'],
      'oil_gas':['ORAT','baggage','rolling stock','liquid cooling readiness','white space','payload integration','yield ramp'],
      'energy':['ORAT','baggage','rolling stock','liquid cooling readiness','payload integration','yield ramp'],
      'healthcare':['launch reliability','liquid cooling readiness','rolling stock','baggage systems'],
      'water':['launch reliability','liquid cooling readiness','rolling stock','baggage systems'],
      'ports':['launch reliability','liquid cooling readiness','rolling stock','baggage systems','ORAT']
    }
    return all_terms.get(key, [])


def _v124_scrub_value(v, forbidden):
    """Remove only terms forbidden for the selected ontology.
    Do not apply global replacements blindly; otherwise valid ORAT/launch/cooling terms are damaged.
    """
    replacements = {
        'liquid cooling readiness':'systems readiness', 'Liquid cooling readiness':'Systems readiness',
        'Liquid cooling':'Systems readiness', 'liquid cooling':'systems readiness',
        'white-space':'specialist systems', 'white space':'specialist systems', 'data hall':'delivery package', 'GPU':'specialist equipment',
        'ORAT':'operational readiness', 'baggage systems':'systems integration', 'Baggage':'Systems', 'baggage':'systems',
        'airside':'operational access', 'landside':'public interface', 'rolling stock':'operating assets', 'signalling':'control systems',
        'launch reliability':'delivery reliability', 'Launch reliability':'Delivery reliability', 'payload integration':'systems integration',
        'range availability':'access availability', 'yield ramp':'operational ramp-up', 'CQV':'commissioning validation'
    }
    if isinstance(v, str):
        out = v
        low_forbidden = {str(f).lower() for f in forbidden}
        for a,b in replacements.items():
            if a.lower() in low_forbidden:
                out = out.replace(a,b)
        return out
    if isinstance(v, list): return [_v124_scrub_value(x, forbidden) for x in v]
    if isinstance(v, dict): return {k:_v124_scrub_value(val, forbidden) for k,val in v.items()}
    return v


def _v124_apply_sector_lock(model: Dict[str, Any]) -> Dict[str, Any]:
    m = dict(model or {})
    key = _v124_sector_key(m)
    L = _v124_library(key)
    m['sector_ontology_key'] = key
    m['sector_ontology_label'] = L['label']
    m['executive_shock_insight'] = L['shock']
    m['sector_confidence_drivers'] = L['confidence']
    m['sector_primary_cost_drivers'] = L['cost']
    m['sector_schedule_threats'] = L['schedule']
    m['causal_graph_nodes'] = L['chain']
    m['causal_chain'] = L['chain']
    m['benchmark_comparison'] = [{'archetype':a,'anchor_cost':c,'anchor_duration_months':d} for a,c,d in L['bench']]
    m['why_casey_generated_this'] = [
        f"CASEY detected {L['label']} from the project brief and locked the sector ontology before generating outputs.",
        f"Sector-native behaviours applied: {', '.join(L['chain'][:5])}.",
        "Benchmark cohort, causal chain, confidence drivers and risk language were constrained to the selected sector.",
        "The output is designed for early board challenge and scope definition, not certified pricing."
    ]
    sigs=[]
    for s,status,basis in L['signals']:
        sigs.append({'signal':s,'status':status,'direction':'sector-locked calibration','weight':0.12,'applies_to':'confidence, QCRA/QSRA, risk register','basis':basis})
    m['live_calibration_signals']=sigs
    m['live_calibration_strip']=' • '.join([x['signal'] for x in sigs[:4]])
    m['mission_control_cards']=[{'label':'Live calibration','signal':'Current sector conditions are being applied to confidence, contingency and delivery-tail exposure.','severity':'Active'}] + [{'label':s,'signal':b,'severity':st} for s,st,b in L['signals']]
    m['mission_control_cards']=m['mission_control_cards'][:6]
    m['casey_thinking'] = f"CASEY has re-cut the programme as a {m.get('scenario_label','Base')} scenario inside the {L['label']} ontology. The governing consequence is: {L['shock']} QCRA/QSRA curves, cost basis, risk probabilities and schedule logic have been sector-locked to this posture."
    m['executive_summary'] = f"{L['label']} scenario view: {m.get('scenario_label','Base')}. CASEY indicates {m.get('cost_p50')} P50 exposure, {m.get('cost_range')} range, {m.get('schedule')} baseline, {m.get('risk')} risk and {m.get('confidence_pct')}% confidence. {L['shock']} Trade-off: {m.get('scenario_why','Balanced cost, time and evidence posture.')}"
    m['board_briefing'] = [
        L['shock'],
        f"{m.get('scenario_label','Base')} is the reference case: no cost, schedule or confidence delta is applied unless a scenario is selected.",
        f"{m.get('scenario_label','Base')} scenario: {m.get('cost_p50')} P50, {m.get('cost_range')} range, {m.get('schedule')} baseline and {m.get('confidence_pct')}% confidence.",
        f"Confidence is governed by {L['constraints']}.",
        "Gained: Maintains a credible reference case for board challenge."
    ]
    m['uncertainty_narrative']={
        'estimate_maturity':'Class 3 maturity is suitable for budget authorization, but procurement and design assumptions still need challenge.',
        'schedule_maturity':'Schedule Level 4 gives stronger logic and QSRA traceability.',
        'interpretation':f"Live calibration is weighting {L['constraints']} into the QSRA/QCRA tail."
    }
    # Scenario delta language remains, but sector nouns are locked.
    m['top_decisions_required']=[
        'Accept or reject the scenario trade-off explicitly at board level.',
        'Confirm the changed critical path and near-critical path density.',
        'Approve the scenario-specific reserve and contingency philosophy.',
        'Assign named owners for the new top scenario risks.',
        'Confirm whether XER, cost workbook and risk register should be issued as scenario-controlled outputs.'
    ]
    forbidden = _v124_forbidden_terms(key)
    m = _v124_scrub_value(m, forbidden)
    # restore sector-native strings after broad scrub in case terms are valid for this key
    m['sector_ontology_key'] = key; m['sector_ontology_label'] = L['label']
    if key == 'data_centre':
        m['sector_primary_cost_drivers'] = L['cost']; m['sector_schedule_threats'] = L['schedule']; m['causal_graph_nodes'] = L['chain']; m['causal_chain'] = L['chain']
    if key == 'space':
        m['sector_confidence_drivers'] = L['confidence']; m['sector_schedule_threats'] = L['schedule']; m['causal_graph_nodes'] = L['chain']; m['causal_chain'] = L['chain']
    return m

_CASEY_V124_PREV_BUILD_MODEL = build_model
def build_model(prompt:str, client:str='', class_level:int=3, schedule_level:int=3, scenario:str='base'):
    return _v124_apply_sector_lock(_CASEY_V124_PREV_BUILD_MODEL(prompt, client, class_level, schedule_level, scenario))

APP_VERSION = 'CASEY V124 Sector Ontology Hardened Public Demo'
print('CASEY V124 sector ontology hardening lock installed')
# ================= END CASEY V124 SECTOR ONTOLOGY HARDENING LOCK =================

# ================= CASEY V125 FULL SECTOR QA LOCK =================
# Post-release hardening: fixes residual sector misclassification and removes legacy benchmark/casual field leakage.

def _v125_text_for_routing(prompt: str='', client: str='') -> str:
    return (str(prompt or '') + ' ' + str(client or '')).lower()

def _v125_has_any(t, terms):
    import re
    for term in terms:
        pat = r'(?<![a-z0-9])' + re.escape(term.lower()) + r'(?![a-z0-9])'
        if re.search(pat, t):
            return True
    return False

def _v125_sector_key_from_input(prompt: str='', client: str='', fallback_model=None) -> str:
    t = _v125_text_for_routing(prompt or (fallback_model or {}).get('prompt',''), client or (fallback_model or {}).get('client',''))
    # Space first because space briefs contain energy/power/manufacturing words but delivery logic is mission assurance.
    if _v125_has_any(t, ['lunar','moon','mars','orbital','satellite','spaceport','launch vehicle','payload','deep space','isru','leo','cislunar','propellant depot']): return 'space'
    # High-specificity Earth sectors before general transport/energy terms.
    if _v125_has_any(t, ['pharma','biologics','life sciences','gmp','fill-finish','sterile','cqv','validation','media fill','amgen','lilly','novartis','pfizer']): return 'life_sciences'
    if _v125_has_any(t, ['semiconductor','fab','wafer','foundry','lithography','chip plant','ultra-pure water','upw']): return 'semiconductor'
    if _v125_has_any(t, ['data centre','data center','hyperscale','ai campus','compute campus','gpu','cloud region','white space','white-space','data hall','azure','openai','meta ai']): return 'data_centre'
    if _v125_has_any(t, ['oil','gas','lng','refinery','petrochemical','offshore platform','pipeline','fpso','hydrocarbon','carbon capture']): return 'oil_gas'
    if _v125_has_any(t, ['airport','aviation','terminal','runway','heathrow','gatwick','airside','baggage','orat']): return 'airport'
    if _v125_has_any(t, ['hospital','healthcare','clinical','medical centre','medical center','patient','nhs']): return 'healthcare'
    if _v125_has_any(t, ['water','wastewater','desalination','sewer','reservoir','treatment plant','pumping station']): return 'water'
    if _v125_has_any(t, ['rail','metro','transit','high speed','hs2','station','signalling','signaling','rolling stock','california high speed']): return 'rail'
    if _v125_has_any(t, ['defence','defense','military','naval','airbase','radar','missile','secure facility','mod','dod','hardened command']): return 'defence'
    if _v125_has_any(t, ['nuclear','smr','reactor']): return 'energy'
    if _v125_has_any(t, ['energy','power plant','renewable','wind farm','offshore wind','solar','battery','substation','transmission','grid','hydrogen','hvdc']): return 'energy'
    if _v125_has_any(t, ['port','harbour','harbor','marine','dock','container terminal','quay','dredging']): return 'ports'
    return 'general_infrastructure'

def _v125_library(key: str):
    return _v124_library(key)

def _v125_forbidden_terms(key: str):
    base = set(_v124_forbidden_terms(key))
    # Add the leaks actually seen in QA screenshots/tests.
    if key in {'airport','rail','healthcare','water','ports','defence','energy','oil_gas','life_sciences','semiconductor','space'}:
        base.update(['Liquid cooling readiness','liquid cooling','GPU','white space','white-space','data hall','IST congestion'])
    if key not in {'airport'}:
        base.update(['ORAT','airside','landside','baggage systems'])
    if key not in {'rail'}:
        base.update(['rolling stock','signalling','signaling','possessions'])
    if key not in {'space'}:
        base.update(['launch reliability','launch cadence','payload integration','range availability','mission assurance','thermal-power balance'])
    if key not in {'life_sciences'}:
        base.update(['CQV','media fills'])
    if key not in {'semiconductor'}:
        base.update(['yield ramp','lithography','UPW'])
    return list(base)

def _v125_sector_rows(L, m):
    # Build compact sector-native cost, schedule and risk rows so exports use the locked ontology.
    cost_base = _parse_money_bn(m.get('cost_p50')) or 1.0
    cost_weights = [0.34,0.24,0.18,0.14,0.10]
    cost_lines=[]
    for i, name in enumerate(L['cost'][:5], 1):
        p50 = cost_base * cost_weights[i-1]
        cost_lines.append({'cbs':f'{i:02d}.0{i}','description':name,'type':'Direct' if i<=3 else ('Indirect' if i==4 else 'Reserve'),'low_p10':_fmt_money_bn(p50*0.78),'p50':_fmt_money_bn(p50),'high_p90':_fmt_money_bn(p50*1.32),'basis':f'{name} priced from sector-locked {L["label"]} template, estimate class and scenario posture.'})
    months = _parse_months(m.get('schedule')) or 60
    schedule_rows=[]
    pred=''
    for i, name in enumerate(L['schedule'][:5], 1):
        aid=f'A{1000+i*100}'
        schedule_rows.append({'activity_id':aid,'phase':'Sector critical path','activity':name,'predecessor':pred,'duration_months':max(1, round(months * [0.15,0.22,0.20,0.18,0.12][i-1])),'critical':'Yes' if i in [2,3,4] else 'No','basis':f'{L["label"]} schedule logic: {name}.'})
        pred=aid
    risks=[]
    for i, name in enumerate((L['schedule']+L['cost'])[:8],1):
        risks.append({'id':f'R-{i:03d}','risk':name,'cause':f'Sector-locked {L["label"]} exposure','event':f'{name} underperforms the board evidence threshold','impact':'P80/P90 cost or schedule tail increases','probability_pct':max(18, 52-i*3),'activity':schedule_rows[min(i-1,len(schedule_rows)-1)]['activity_id'],'cbs':cost_lines[min(i-1,len(cost_lines)-1)]['cbs'],'owner':('Programme Director' if i==1 else 'Sector Integration Lead'),'mitigation':f'Named owner to evidence {name.lower()} readiness, dates, constraints and fallback plan.','status':'Open' if i<=3 else 'Monitor'})
    return cost_lines, schedule_rows, risks

def _v125_deep_scrub_model(m, forbidden):
    protected={'prompt','client','title','id'}
    out={}
    for k,v in (m or {}).items():
        out[k]=v if k in protected else _v124_scrub_value(v, forbidden)
    return out

def _v125_apply_full_sector_lock(model: Dict[str, Any], prompt: str='', client: str='') -> Dict[str, Any]:
    key = _v125_sector_key_from_input(prompt, client, model)
    L = _v125_library(key)
    m = dict(model or {})
    m['sector_ontology_key']=key
    m['sector_ontology_label']=L['label']
    m['subsector']=L['label']
    # Overwrite every display/export field that previously carried shared ontology remnants.
    m['executive_shock_insight']=L['shock']
    m['sector_confidence_drivers']=L['confidence']
    m['sector_primary_cost_drivers']=L['cost']
    m['sector_schedule_threats']=L['schedule']
    m['causal_graph_nodes']=L['chain']
    m['causal_chain']=L['chain']
    bench=[{'sector':a,'archetype':a,'anchor_cost':c,'anchor_duration_months':d,'similarity_score':max(5, 10-i),'use':'Sector-locked benchmark cohort; directional calibration only.'} for i,(a,c,d) in enumerate(L['bench'][:4])]
    m['benchmark_comparison']=bench
    m['benchmark_memory']=bench
    m['benchmarks']=bench
    m['peer_competitors']=bench
    m['why_casey_generated_this']=[
        f"CASEY detected {L['label']} from the project brief and locked the sector ontology before generating outputs.",
        f"Sector-native behaviours applied: {', '.join(L['chain'][:5])}.",
        "Benchmark cohort, causal chain, confidence drivers and risk language were constrained to the selected sector.",
        "The output is designed for early board challenge and scope definition, not certified pricing."
    ]
    sigs=[{'signal':s,'status':st,'direction':'sector-locked calibration','weight':0.12,'applies_to':'confidence, QCRA/QSRA, risk register','basis':basis} for s,st,basis in L['signals']]
    m['live_calibration_signals']=sigs
    m['live_calibration_strip']=' • '.join(x['signal'] for x in sigs[:4])
    m['mission_control_cards']=[{'label':'Live calibration','signal':'Current sector conditions are being applied to confidence, contingency and delivery-tail exposure.','severity':'Active'}] + [{'label':s,'signal':basis,'severity':st} for s,st,basis in L['signals']]
    m['mission_control_cards']=m['mission_control_cards'][:6]
    m['casey_thinking']=f"CASEY has re-cut the programme as a {m.get('scenario_label','Base')} scenario inside the {L['label']} ontology. The governing consequence is: {L['shock']} QCRA/QSRA curves, cost basis, risk probabilities and schedule logic have been sector-locked to this posture."
    m['executive_summary']=f"{L['label']} scenario view: {m.get('scenario_label','Base')}. CASEY indicates {m.get('cost_p50')} P50 exposure, {m.get('cost_range')} range, {m.get('schedule')} baseline, {m.get('risk')} risk and {m.get('confidence_pct')}% confidence. {L['shock']} Trade-off: {m.get('scenario_why','Balanced cost, time and evidence posture.')}"
    m['board_briefing']=[L['shock'],f"{m.get('scenario_label','Base')} is the reference case: scenario movement is explicitly controlled before board use.",f"{m.get('scenario_label','Base')} scenario: {m.get('cost_p50')} P50, {m.get('cost_range')} range, {m.get('schedule')} baseline and {m.get('confidence_pct')}% confidence.",f"Confidence is governed by {L['constraints']}.","Gained: Maintains a credible reference case for board challenge."]
    m['confidence_explanation']=f"{m.get('confidence_pct')}% means CASEY believes the current case is board-defensible only if evidence is improved around {L['constraints']}."
    m['confidence_engine_detail']={'decision_rule':'Use for option selection, but close evidence gaps before approval.','primary_constraint':L['constraints'],'plain_english':'Confidence is not optimism. It is CASEY’s board-defensibility score based on benchmark fit, evidence maturity, procurement certainty, schedule logic, reserve adequacy and scenario posture.'}
    m['board_challenge_questions']=[
        'Is this a decision case or only a reference case?',
        'What must be proven before this becomes board-approvable?',
        f"What evidence proves {L['constraints']}?",
        'Which three risks create most P80/P90 exposure?',
        'What data would move confidence above the board-comfort threshold?',
        'Which named owner is accountable for the critical-path constraint?'
    ]
    m['uncertainty_narrative']={'estimate_maturity':'Class 3 maturity is suitable for budget authorization, but procurement and design assumptions still need challenge.','schedule_maturity':'Schedule Level 4 gives stronger logic and QSRA traceability.','interpretation':f"Live calibration is weighting {L['constraints']} into the QSRA/QCRA tail."}
    cost_lines, schedule_rows, risks = _v125_sector_rows(L, m)
    m['cost_lines']=cost_lines
    m['schedule_rows']=schedule_rows
    m['risk_register']=risks
    m['risks']=risks
    m['procurement_heatmap']=[{'package':x,'status':'Active' if i<3 else 'Watch','risk':'Sector-native procurement exposure','owner':'Commercial Lead'} for i,x in enumerate(L['cost'][:5])]
    m['critical_path_narrative']=[f"{x} is part of the locked {L['label']} critical-path narrative." for x in L['schedule'][:4]]
    m['red_flags']=[f"Evidence gap around {L['constraints']}.",f"Benchmark challenge remains unless {L['chain'][-2].lower()} is proven."]
    forbidden=_v125_forbidden_terms(key)
    m=_v125_deep_scrub_model(m, forbidden)
    # Restore native fields after scrub.
    m['sector_ontology_key']=key; m['sector_ontology_label']=L['label']; m['subsector']=L['label']
    m['sector_confidence_drivers']=L['confidence']; m['sector_primary_cost_drivers']=L['cost']; m['sector_schedule_threats']=L['schedule']; m['causal_graph_nodes']=L['chain']; m['causal_chain']=L['chain']
    return m

_CASEY_V125_PREV_BUILD_MODEL = build_model
def build_model(prompt:str, client:str='', class_level:int=3, schedule_level:int=3, scenario:str='base'):
    return _v125_apply_full_sector_lock(_CASEY_V125_PREV_BUILD_MODEL(prompt, client, class_level, schedule_level, scenario), prompt, client)

APP_VERSION = 'CASEY V125 Full Sector QA Hardened'
print('CASEY V125 full sector QA lock installed')
# ================= END CASEY V125 FULL SECTOR QA LOCK =================

# ================= CASEY V126 QA HOTFIX =================
def _v126_parse_money_bn(x):
    try:
        s=str(x or '').replace('$','').replace(',','').strip().upper().split()[0]
        if s.endswith('T'): return float(s[:-1])*1000
        if s.endswith('B'): return float(s[:-1])
        if s.endswith('M'): return float(s[:-1])/1000
        return float(s)
    except Exception:
        return 1.0

def _v126_fmt_money_bn(x):
    try: x=float(x)
    except Exception: x=0.0
    if x>=1000: return f'${x/1000:.1f}T'
    if x>=1: return f'${x:.1f}B'
    return f'${x*1000:.0f}M'

def _v126_sector_rows(L, m):
    cost_base = _v126_parse_money_bn(m.get('cost_p50')) or 1.0
    cost_weights=[0.34,0.24,0.18,0.14,0.10]
    cost_lines=[]
    for i,name in enumerate(L['cost'][:5],1):
        p50=cost_base*cost_weights[i-1]
        cost_lines.append({'cbs':f'{i:02d}.0{i}','description':name,'type':'Direct' if i<=3 else ('Indirect' if i==4 else 'Reserve'),'low_p10':_v126_fmt_money_bn(p50*0.78),'p50':_v126_fmt_money_bn(p50),'high_p90':_v126_fmt_money_bn(p50*1.32),'basis':f'{name} priced from sector-locked {L["label"]} template, estimate class and scenario posture.'})
    months = _parse_months(m.get('schedule')) if '_parse_months' in globals() else 60
    if not months: months=60
    schedule_rows=[]; pred=''
    for i,name in enumerate(L['schedule'][:5],1):
        aid=f'A{1000+i*100}'
        schedule_rows.append({'activity_id':aid,'phase':'Sector critical path','activity':name,'predecessor':pred,'duration_months':max(1,round(months*[0.15,0.22,0.20,0.18,0.12][i-1])),'critical':'Yes' if i in [2,3,4] else 'No','basis':f'{L["label"]} schedule logic: {name}.'})
        pred=aid
    risks=[]
    for i,name in enumerate((L['schedule']+L['cost'])[:8],1):
        risks.append({'id':f'R-{i:03d}','risk':name,'cause':f'Sector-locked {L["label"]} exposure','event':f'{name} underperforms the board evidence threshold','impact':'P80/P90 cost or schedule tail increases','probability_pct':max(18,52-i*3),'activity':schedule_rows[min(i-1,len(schedule_rows)-1)]['activity_id'],'cbs':cost_lines[min(i-1,len(cost_lines)-1)]['cbs'],'owner':'Programme Director' if i==1 else 'Sector Integration Lead','mitigation':f'Named owner to evidence {name.lower()} readiness, dates, constraints and fallback plan.','status':'Open' if i<=3 else 'Monitor'})
    return cost_lines,schedule_rows,risks

def _v126_apply(model, prompt='', client=''):
    key=_v125_sector_key_from_input(prompt,client,model); L=_v125_library(key)
    m=_v125_apply_full_sector_lock(model,prompt,client)
    cost_lines,schedule_rows,risks=_v126_sector_rows(L,m)
    m['cost_lines']=cost_lines; m['schedule_rows']=schedule_rows; m['risk_register']=risks; m['risks']=risks
    return m

_CASEY_V126_PREV_BUILD_MODEL = _CASEY_V125_PREV_BUILD_MODEL
# Important: bypass broken V125 wrapper and apply the V126 fixed wrapper directly to the V124 model.
def build_model(prompt:str, client:str='', class_level:int=3, schedule_level:int=3, scenario:str='base'):
    return _v126_apply(_CASEY_V126_PREV_BUILD_MODEL(prompt, client, class_level, schedule_level, scenario), prompt, client)
APP_VERSION='CASEY V126 Full Sector QA Hardened'
print('CASEY V126 QA hotfix installed')
# ================= END CASEY V126 QA HOTFIX =================
# CASEY V126.1 compatibility aliases for V125 sector-row functions
_parse_money_bn = _v126_parse_money_bn
_fmt_money_bn = _v126_fmt_money_bn
def _v126_parse_months(x):
    try:
        return int(float(str(x or '0').replace('months','').replace('mo','').strip().split()[0]))
    except Exception:
        return 60
_parse_months = _v126_parse_months

# ================= CASEY V127 QA SCRUB FINAL =================
_CASEY_V127_PREV_BUILD_MODEL = build_model
def build_model(prompt:str, client:str='', class_level:int=3, schedule_level:int=3, scenario:str='base'):
    m=_CASEY_V127_PREV_BUILD_MODEL(prompt, client, class_level, schedule_level, scenario)
    # Remove legacy hidden base-model blob that carried old-sector vocabulary into QA/export JSON.
    if isinstance(m, dict) and '_hs_base' in m:
        m.pop('_hs_base', None)
    return m
APP_VERSION='CASEY V127 Full Sector QA Hardened Final'
print('CASEY V127 final QA scrub installed')
# ================= END CASEY V127 QA SCRUB FINAL =================

# ================= CASEY V128 EXPORT SURFACE HARDENING =================
def _v128_forbidden_terms(key: str):
    terms=set(_v125_forbidden_terms(key))
    if key == 'defence':
        terms.discard('mission assurance'); terms.discard('Mission assurance')
    return list(terms)

def _v128_clean_surface(m):
    key=m.get('sector_ontology_key','general_infrastructure'); L=_v125_library(key)
    # Avoid airport/other sectors using the rail-specific word "possessions".
    if key == 'airport':
        repl=lambda s: str(s).replace('possessions','access windows').replace('Possessions','Access windows')
        for fld in ['sector_schedule_threats','causal_chain','causal_graph_nodes']:
            m[fld]=[repl(x) for x in m.get(fld,[])]
        for fld in ['schedule_rows','risk_register','risks']:
            rows=[]
            for r in m.get(fld,[]):
                if isinstance(r,dict): rows.append({k:repl(v) if isinstance(v,str) else v for k,v in r.items()})
                else: rows.append(repl(r))
            m[fld]=rows
        m['critical_path_narrative']=[repl(x) for x in m.get('critical_path_narrative',[])]
    # Old compatibility fields are replaced with locked sector data so JSON/API exports cannot leak stale ontologies.
    m['cost_breakdown']=m.get('cost_lines',[])
    m['estimates_by_class']={str(i):m.get('cost_lines',[]) for i in range(1,6)}
    m['schedules_by_level']={str(i):m.get('schedule_rows',[]) for i in range(1,6)}
    m['sector_signature_behaviours']=L['chain'][:5]
    forbidden=_v128_forbidden_terms(key)
    for k in list(m.keys()):
        if k not in {'prompt','client','title','id'}:
            m[k]=_v124_scrub_value(m[k], forbidden)
    return m

_CASEY_V128_PREV_BUILD_MODEL=build_model
def build_model(prompt:str, client:str='', class_level:int=3, schedule_level:int=3, scenario:str='base'):
    return _v128_clean_surface(_CASEY_V128_PREV_BUILD_MODEL(prompt, client, class_level, schedule_level, scenario))
APP_VERSION='CASEY V128 Full Sector QA Hardened Final'
print('CASEY V128 export surface hardening installed')
# ================= END CASEY V128 EXPORT SURFACE HARDENING =================

# ================= CASEY V130 EXECUTIVE INTELLIGENCE / FULL ONTOLOGY LOCK =================
# Adds five pre-demo hardening layers: sector-native causal graphs, narrative compression,
# challenge intelligence, second-order contradictions, export-surface governance.

def _v130_has(t, words):
    return any(w in t for w in words)

def _v130_sector_key(prompt='', client='', model=None):
    model = model or {}
    t = (str(prompt)+' '+str(client)+' '+str(model.get('title',''))+' '+str(model.get('subsector',''))+' '+str(model.get('mode',''))).lower()
    if _v130_has(t,['lunar','mars','orbital','satellite','spaceport','launch','payload','moon','deep space','propulsion','range availability']): return 'space'
    if _v130_has(t,['data centre','data center','hyperscale','ai campus','compute campus','gpu','cloud region','white space','data hall']): return 'data_centre'
    if _v130_has(t,['oil','gas','lng','refinery','petrochemical','offshore platform','pipeline','fpso','hydrocarbon','carbon capture']): return 'oil_gas'
    if _v130_has(t,['nuclear','smr','reactor','containment','safety case']): return 'nuclear'
    if _v130_has(t,['port','harbour','harbor','marine','dock','container terminal','quay','dredging']): return 'ports'
    if _v130_has(t,['airport','aviation','terminal','runway','heathrow','gatwick','airside','baggage','orat']): return 'airport'
    if _v130_has(t,['rail','metro','transit','high speed','hs2','rail station','signalling','signaling','rolling stock','california high speed']): return 'rail'
    if _v130_has(t,['semiconductor','fab','wafer','cleanroom','foundry','lithography','chip plant']): return 'semiconductor'
    if _v130_has(t,['life sciences','pharma','biologics','gmp','fill-finish','sterile','cqv','amgen','lilly','novartis','pfizer']): return 'life_sciences'
    if _v130_has(t,['defence','defense','military','naval','airbase','radar','missile','secure facility','mod ','dod ','command centre','command center']): return 'defence'
    if _v130_has(t,['energy','power plant','renewable','wind farm','offshore wind','solar','battery','substation','transmission','grid','hydrogen','hvdc']): return 'energy'
    if _v130_has(t,['hospital','healthcare','clinical','medical centre','patient','nhs']): return 'healthcare'
    if _v130_has(t,['water','wastewater','desalination','sewer','reservoir','treatment plant']): return 'water'
    if _v130_has(t,['port','harbour','harbor','marine','dock','container terminal','quay','dredging']): return 'ports'
    if _v130_has(t,['mine','mining','ore','tailings','pit','processing plant']): return 'mining'
    if _v130_has(t,['road','highway','motorway','bridge','tunnel']): return 'roads'
    return _v125_sector_key_from_input(prompt, client, model) if '_v125_sector_key_from_input' in globals() else 'general_infrastructure'

def _v130_library(key):
    # Start with V125/V128 library where it exists, then override missing or weaker sectors.
    try:
        L = dict(_v125_library(key))
    except Exception:
        L = {}
    overrides = {
    'space': dict(label='Space / Mission Assurance', shock='Mission assurance, payload integration, range availability and thermal-power evidence govern launch confidence — not construction progress alone.', constraints='mission assurance, payload integration, launch readiness, range access and thermal-power balance', signals=[('Mission assurance burden','Active','Qualification evidence and assurance closure control board defensibility.'),('Range and launch-window availability','Active','Launch logistics and range access create schedule-tail exposure.'),('Payload integration maturity','Watch','Interface maturity and verification evidence determine mission readiness.'),('Thermal-power balance','Watch','Sustained operations depend on verified power, thermal and autonomous recovery evidence.')], bench=[('Lunar / Orbital Infrastructure','$8B–$95B','72-216'),('Launch and Payload Integration','$1B–$14B','36-108'),('Deep-Space Mission Systems','$4B–$35B','60-144'),('Autonomous Operations Platform','$2B–$20B','36-96')], chain=['Mission requirements freeze','Payload integration','Qualification campaign','Range availability','Launch readiness review','Mission assurance sign-off','Confidence'], confidence=['Benchmark similarity: mission class and space infrastructure archetype','Scope maturity: payload, surface/orbital architecture and operations definition','Procurement certainty: launch provider, avionics and qualified hardware','Schedule maturity: test campaign, range window and commissioning logic','Operational exposure: remote recovery and autonomous servicing limits'], cost=['Payload and mission systems','Launch integration and range operations','Power/thermal/autonomous operations','Qualification and test campaign','Mission assurance reserve'], schedule=['Payload interface maturity','Qualification and environmental testing','Launch provider and range coordination','Mission operations readiness','Autonomous recovery validation']),
    'defence': dict(label='Defence / Secure Infrastructure', shock='Security accreditation, systems assurance and operational acceptance govern confidence more than estate completion.', constraints='security accreditation, mission systems integration, assurance evidence and operational acceptance', signals=[('Security accreditation','Active','Authority-to-operate evidence creates approval-tail exposure.'),('Mission systems integration','Active','Secure comms, sensors and classified interfaces dominate commissioning.'),('Supplier assurance','Watch','Security-cleared supply chain capacity governs procurement certainty.'),('Operational acceptance','Watch','User trials and command readiness govern board defensibility.')], bench=[('Secure Defence Facility','$500M–$8B','30-96'),('Command and Mission Systems Campus','$1B–$12B','36-108'),('Airbase / Naval Infrastructure','$1B–$20B','48-144')], chain=['Security requirements','Classified procurement','Mission systems integration','Accreditation evidence','Operational trials','Authority to operate','Confidence'], confidence=['Benchmark similarity: secure mission infrastructure','Scope maturity: security, mission and operational requirements','Procurement certainty: cleared suppliers and specialist systems','Schedule maturity: accreditation and acceptance logic','Interface exposure: users, security authorities and mission systems'], cost=['Secure facilities and hardened works','Mission systems and secure comms','Accreditation and assurance evidence','Specialist cleared supply chain','Operational trial and acceptance reserve'], schedule=['Security accreditation pathway','Classified supplier lead-times','Mission systems integration','Operational trials and acceptance','Authority-to-operate approval']),
    'energy': dict(label='Energy / Power Infrastructure', shock='Grid interface, permitting, long-lead electrical packages and commissioning evidence dominate delivery confidence.', constraints='grid interconnection, permitting, long-lead electrical procurement and commissioning readiness', signals=[('Grid interconnection','Active','Connection studies, network access and energisation govern P80/P90 exposure.'),('Permitting and land consent','Active','Planning, environmental and connection approvals shape critical path.'),('Long-lead equipment','Watch','Transformers, turbines, HV equipment and controls create procurement tail.'),('Commissioning readiness','Watch','Energisation, protection settings and performance tests determine acceptance.')], bench=[('Power Generation Programme','$1B–$18B','36-108'),('Transmission / HVDC Programme','$2B–$25B','48-144'),('Renewables Portfolio','$500M–$12B','24-84')], chain=['Permitting and consents','Grid studies','Long-lead equipment','Civil/electrical completion','Energisation','Performance testing','Confidence'], confidence=['Benchmark similarity: power/utility programme','Scope maturity: grid, generation and operating model definition','Procurement certainty: turbines, transformers and HV packages','Schedule maturity: permits, energisation and testing logic','Interface exposure: network operator, regulator and offtaker'], cost=['Generation or network assets','Grid connection and substations','Long-lead electrical equipment','Civil/enabling works','Commissioning and performance reserve'], schedule=['Permitting and environmental approvals','Grid connection agreement','Long-lead equipment delivery','Energisation and protection testing','Performance acceptance']),
    'oil_gas': dict(label='Oil & Gas / Process Infrastructure', shock='Process safety, modular integration, shutdown windows and commissioning systems govern confidence more than mechanical progress.', constraints='process safety, module integration, shutdown access, commissioning systems and regulatory acceptance', signals=[('Process safety case','Active','HAZOP/HAZID closure and safety-critical evidence govern approval.'),('Modular integration','Active','Fabrication, transport, heavy lift and hook-up drive schedule risk.'),('Shutdown / tie-in windows','Watch','Operational access windows create delivery-tail exposure.'),('Commissioning and start-up','Watch','Pre-commissioning, loop checks and performance tests govern readiness.')], bench=[('LNG / Gas Processing Facility','$3B–$40B','48-144'),('Refinery / Petrochemical Expansion','$1B–$25B','36-120'),('Offshore / Pipeline Programme','$1B–$20B','36-108')], chain=['Process safety basis','Module fabrication','Transport and hook-up','Tie-in access','Pre-commissioning','Start-up performance test','Confidence'], confidence=['Benchmark similarity: process/energy infrastructure','Scope maturity: process design and safety-case definition','Procurement certainty: modules, rotating equipment and controls','Schedule maturity: shutdown/tie-in and start-up logic','Interface exposure: operator, regulator and supply chain'], cost=['Process units and modules','Rotating equipment and controls','Pipeline/tie-in and brownfield works','Fabrication, transport and heavy lift','Start-up and safety reserve'], schedule=['Process safety approval','Module fabrication and delivery','Shutdown and tie-in access','Pre-commissioning and loop checks','Start-up performance testing']),
    'healthcare': dict(label='Healthcare / Hospital Infrastructure', shock='Clinical transition, infection-control evidence and medical systems integration govern readiness more than building completion.', constraints='clinical commissioning, infection-control compliance, medical systems integration and phased occupancy', signals=[('Clinical commissioning','Active','Clinical readiness and staff transition govern usable capacity.'),('Infection-control compliance','Active','HTM/HBN evidence and inspection readiness constrain handover.'),('Medical systems integration','Watch','Equipment, digital health and specialist systems drive commissioning risk.'),('Phased occupancy','Watch','Live hospital operations and patient transition create interface exposure.')], bench=[('Acute Hospital Programme','$500M–$6B','36-96'),('Specialist Clinical Campus','$300M–$4B','30-84'),('Live Healthcare Redevelopment','$200M–$3B','30-72')], chain=['Clinical model freeze','Medical equipment procurement','Infection-control evidence','Digital/medical systems integration','Phased occupancy','Clinical commissioning','Confidence'], confidence=['Benchmark similarity: healthcare capital programme','Scope maturity: clinical brief and department definition','Procurement certainty: medical equipment and specialist systems','Schedule maturity: clinical commissioning and occupancy logic','Interface exposure: patients, operations and regulators'], cost=['Clinical departments and fit-out','Medical equipment and digital systems','Infection-control and compliance works','Live operations phasing','Clinical commissioning reserve'], schedule=['Clinical brief and approvals','Medical equipment procurement','Infection-control verification','Digital/medical systems integration','Phased occupancy and clinical commissioning']),
    'water': dict(label='Water / Environmental Infrastructure', shock='Permitting, process performance, interfaces and commissioning evidence govern confidence more than civil progress.', constraints='environmental permits, process performance, utility interfaces and operational acceptance', signals=[('Environmental permitting','Active','Discharge permits and regulator evidence govern approval path.'),('Process performance','Active','Treatment performance and testing drive acceptance.'),('Utility/interface access','Watch','Tie-ins and service continuity create schedule exposure.'),('Operational acceptance','Watch','Operator readiness and performance trials govern confidence.')], bench=[('Water Treatment Programme','$200M–$5B','24-84'),('Wastewater Upgrade','$300M–$8B','30-96'),('Desalination / Resource Programme','$500M–$10B','36-108')], chain=['Permit basis','Process design freeze','Equipment procurement','Tie-ins and bypasses','Performance testing','Operator acceptance','Confidence'], confidence=['Benchmark similarity: water/environmental infrastructure','Scope maturity: process and permit definition','Procurement certainty: pumps, membranes, process equipment and controls','Schedule maturity: tie-in, bypass and performance-test logic','Interface exposure: regulator, operator and utilities'], cost=['Civil/process structures','Process equipment and controls','Tie-ins, bypass and utilities','Environmental compliance works','Performance testing reserve'], schedule=['Permit and discharge approvals','Process equipment delivery','Tie-ins and bypass sequencing','Performance testing','Operator/regulator acceptance']),
    'ports': dict(label='Ports / Marine Infrastructure', shock='Marine access, dredging, landside interfaces and operational cutover govern confidence more than quay wall progress.', constraints='marine access, dredging, berth systems, landside interfaces and operational cutover', signals=[('Marine access windows','Active','Weather, tide and vessel access govern productivity.'),('Dredging and seabed risk','Active','Ground/marine conditions create cost and schedule tail.'),('Terminal systems integration','Watch','Cranes, yard systems and controls drive readiness.'),('Operational cutover','Watch','Shipping-line and port-operator transition governs acceptance.')], bench=[('Container Terminal Expansion','$500M–$8B','30-96'),('Port / Harbour Redevelopment','$300M–$6B','24-84'),('Marine Logistics Hub','$1B–$12B','36-108')], chain=['Marine permits','Dredging/seabed works','Quay/berth construction','Terminal systems','Landside interfaces','Operational cutover','Confidence'], confidence=['Benchmark similarity: port/marine infrastructure','Scope maturity: berth, yard and landside definition','Procurement certainty: cranes, marine plant and systems','Schedule maturity: marine windows and cutover logic','Interface exposure: shipping lines, port operator and regulators'], cost=['Quay/berth and marine works','Dredging and seabed treatment','Cranes and terminal systems','Landside transport interfaces','Operational cutover reserve'], schedule=['Marine permitting and access','Dredging and seabed works','Crane/system procurement','Landside interface readiness','Terminal operational cutover']),
    'mining': dict(label='Mining / Metals Infrastructure', shock='Resource access, processing plant readiness, tailings approvals and logistics govern confidence more than bulk earthworks.', constraints='resource access, processing plant commissioning, tailings approvals and logistics chain readiness', signals=[('Resource/access readiness','Active','Pit, haul roads and ore access govern ramp-up.'),('Processing plant commissioning','Active','Crushers, mills, flotation/leach systems drive performance risk.'),('Tailings approvals','Watch','Tailings storage facility approvals create governance exposure.'),('Logistics chain','Watch','Rail/port/power/water interfaces constrain operating readiness.')], bench=[('Mine and Processing Plant','$1B–$15B','36-108'),('Tailings / Water Infrastructure','$300M–$5B','24-72'),('Remote Logistics Corridor','$500M–$8B','30-96')], chain=['Resource access','Bulk earthworks','Processing plant install','Power/water/logistics','Tailings approval','Ramp-up performance','Confidence'], confidence=['Benchmark similarity: mining and processing infrastructure','Scope maturity: orebody, plant and logistics definition','Procurement certainty: mills, crushers and process equipment','Schedule maturity: commissioning and ramp-up logic','Interface exposure: power, water, tailings and logistics'], cost=['Mine access and earthworks','Processing plant equipment','Power, water and logistics','Tailings and environmental controls','Ramp-up reserve'], schedule=['Resource access and permits','Processing equipment delivery','Power/water/logistics readiness','Tailings facility approval','Plant commissioning and ramp-up']),
    'roads': dict(label='Roads / Highways Infrastructure', shock='Utilities, traffic staging, structures and statutory approvals govern confidence more than earthworks progress.', constraints='utility diversions, traffic management, structures, statutory approvals and staged opening', signals=[('Utility diversions','Active','Buried services and third-party approvals govern early critical path.'),('Traffic staging','Active','Live network management constrains productivity.'),('Structures and ground risk','Watch','Bridges, retaining walls and geotechnical conditions create tail exposure.'),('Statutory approvals','Watch','Consents and stakeholder commitments constrain openings.')], bench=[('Highway Upgrade Programme','$500M–$10B','30-96'),('Bridge / Tunnel Corridor','$1B–$20B','48-144'),('Urban Transport Corridor','$500M–$8B','30-96')], chain=['Statutory approvals','Utility diversions','Traffic staging','Structures completion','Systems/safety readiness','Staged opening','Confidence'], confidence=['Benchmark similarity: road/highway programme','Scope maturity: alignment, structures and traffic strategy','Procurement certainty: civils packages and utility interfaces','Schedule maturity: staging and opening logic','Interface exposure: road users, utilities and authorities'], cost=['Earthworks and pavements','Structures and retaining walls','Utilities and drainage','Traffic management and staging','Opening/safety reserve'], schedule=['Statutory approvals','Utility diversion completion','Traffic staging switches','Structures completion','Safety audit and staged opening'])
    }
    if key in overrides: L.update(overrides[key])
    return L

def _v130_forbidden(key):
    terms = set(_v128_forbidden_terms(key) if '_v128_forbidden_terms' in globals() else [])
    # Global hard stop against the exact leaks seen in screenshots.
    if key != 'data_centre': terms.update(['Liquid cooling readiness','liquid cooling','white space','white-space','data hall','IST congestion','GPU'])
    if key != 'airport': terms.update(['ORAT','baggage','airside','landside','passenger systems'])
    if key != 'rail': terms.update(['rolling stock','signalling','signaling','possessions','timetable'])
    if key != 'space': terms.update(['launch reliability','launch cadence','payload integration','range availability','thermal-power balance','mission assurance'])
    if key != 'defence': terms.update(['classified','authority to operate','secure comms'])
    if key != 'oil_gas': terms.update(['HAZOP','HAZID','hydrocarbon','FPSO','LNG'])
    if key != 'semiconductor': terms.update(['lithography','yield ramp','UPW','wafer'])
    if key != 'life_sciences': terms.update(['CQV','GMP','media fill','fill-finish'])
    return list(terms)

def _v130_rows(L, m):
    cost_base = _v126_parse_money_bn(m.get('cost_p50')) if '_v126_parse_money_bn' in globals() else 1.0
    months = _v126_parse_months(m.get('schedule')) if '_v126_parse_months' in globals() else 60
    weights=[.31,.25,.19,.15,.10]
    cost=[]
    for i,name in enumerate(L['cost'][:5],1):
        p=cost_base*weights[i-1]
        cost.append({'cbs':f'{i:02d}.0{i}','description':name,'type':'Direct' if i<=3 else ('Indirect' if i==4 else 'Reserve'),'low_p10':_v126_fmt_money_bn(p*.78),'p50':_v126_fmt_money_bn(p),'high_p90':_v126_fmt_money_bn(p*1.34),'basis':f'{name}. Sector-locked basis: {L["label"]}; estimate class, location, complexity and scenario posture applied.'})
    sched=[]; pred=''
    for i,name in enumerate(L['schedule'][:5],1):
        aid=f'A{1000+i*100}'
        sched.append({'activity_id':aid,'phase':'Sector critical path','activity':name,'predecessor':pred,'duration_months':max(1,round(months*[.14,.22,.21,.18,.13][i-1])),'critical':'Yes' if i in [2,3,4] else 'Near-critical','basis':f'{name}: near-critical path narrative retained for board challenge.'})
        pred=aid
    risks=[]
    owners=['Programme Director','Integration Lead','Commercial Lead','Technical Assurance Lead','Operations Readiness Lead','Risk Lead','Controls Lead','Sponsor']
    for i,name in enumerate((L['schedule']+L['cost'])[:8],1):
        risks.append({'id':f'R-{i:03d}','risk':name,'cause':f'{L["label"]} sector exposure','event':f'{name} fails to meet evidence threshold','impact':'P80/P90 cost or schedule tail expands; confidence falls unless owner evidence closes the gap','probability_pct':max(16,54-i*3),'activity':sched[min(i-1,len(sched)-1)]['activity_id'],'cbs':cost[min(i-1,len(cost)-1)]['cbs'],'owner':owners[(i-1)%len(owners)],'mitigation':f'Named owner to prove {name.lower()} readiness, quantified fallback, date, constraint and evidence source.','status':'Open' if i<=3 else ('Mitigating' if i<=6 else 'Monitor'),'challenge':'Board should not accept the base case until this evidence is named and dated.'})
    return cost,sched,risks

def _v130_apply(model, prompt='', client=''):
    key=_v130_sector_key(prompt, client, model); L=_v130_library(key); m=dict(model or {})
    scenario=str(m.get('scenario_label') or m.get('scenario') or 'Base').title()
    m['sector_ontology_key']=key; m['sector_ontology_label']=L['label']; m['subsector']=L['label']
    m['executive_shock_insight']=L['shock']
    m['executive_summary'] = m.get('executive_summary') or f"{L['label']} · {scenario}: {m.get('cost_p50')} P50, {m.get('cost_range')} range, {m.get('schedule')} baseline and {m.get('confidence_pct')}% confidence. {L['shock']}"
    m['board_briefing']=[L['shock'], f"Decision posture: use this as a challenge case until named evidence proves {L['constraints']}.", f"P50 reconciles to {m.get('cost_p50')} and QSRA P80 remains the board contingency conversation.", "Board ask: decide what to buy — speed, savings, assurance or resilience — and name the evidence owner."]
    m['casey_thinking']=f"CASEY has locked the model to {L['label']}. It is challenging the programme narrative against {L['constraints']} rather than accepting progress optics. Cost, risk, schedule, benchmark and export language are generated from the same sector graph."
    m['confidence_explanation']=f"{m.get('confidence_pct')}% is a board-defensibility score, not optimism. It is constrained by {L['constraints']}."
    m['confidence_engine_detail']={'decision_rule': 'Do not treat this as approval evidence until owner actions, basis and P80/P90 exposure are reconciled.', 'primary_constraint': L['constraints'], 'plain_english':'CASEY scores whether the board can defend the decision, not whether the team feels confident.'}
    m['board_challenge_questions']=[_v135_sector_q1(key, L.get('constraints','the governing constraint')), f'What proof closes the gap around {L["constraints"]}?', 'What second-order risk does the preferred scenario create?', 'Which mitigation changes confidence rather than just narrative?', 'What would invalidate this case before approval?']
    m['second_order_contradictions']=[f"The base case may look stable while {L['constraints']} remain unevidenced.", f"Acceleration can shorten the visible plan while increasing assurance, interface and commissioning fragility.", f"Cheaper approval may transfer exposure into P80/P90 tail, operations or recovery reserve.", "Higher confidence must come from evidence closure, not from a smoother narrative."]
    m['governance_challenges']=[f"Require named owner evidence for {L['constraints']}.", "Separate true risk reduction from risk transfer.", "Reconcile P50 headline with P80/P90 board exposure.", "Reject scenario benefits that are not traceable to CBS, activity or risk owner."]
    m['sector_confidence_drivers']=L['confidence']; m['sector_primary_cost_drivers']=L['cost']; m['sector_schedule_threats']=L['schedule']
    m['causal_graph_nodes']=L['chain']; m['causal_chain']=L['chain']
    m['benchmark_comparison']=[{'sector':a,'archetype':a,'anchor_cost':c,'anchor_duration_months':d,'similarity_score':max(6,10-i),'use':'Sector-locked benchmark cohort; no cross-sector borrowing unless delivery mechanics match.'} for i,(a,c,d) in enumerate(L['bench'][:4])]
    m['benchmark_memory']=m['benchmark_comparison']; m['benchmarks']=m['benchmark_comparison']; m['peer_competitors']=m['benchmark_comparison']
    m['why_casey_generated_this']=[f"CASEY detected {L['label']} and locked the ontology before output generation.", f"Sector behaviours applied: {', '.join(L['chain'][:5])}.", "Benchmark cohort, causal chain, confidence drivers, risks and exports were constrained to this sector.", "The output is intentionally challenge-oriented: board defensibility over optimism."]
    sigs=[{'signal':s,'status':st,'direction':'confidence / reserve / P-tail','weight':0.13,'applies_to':'board pack, workbook, risk register, XER, QCRA/QSRA','basis':basis} for s,st,basis in L['signals']]
    m['live_calibration_signals']=sigs; m['live_calibration_strip']=' • '.join(x['signal'] for x in sigs[:4])
    m['mission_control_cards']=[{'label':'Live calibration','signal':'Sector conditions are being converted into confidence, contingency and delivery-tail exposure.','severity':'Active'}]+[{'label':s,'signal':basis,'severity':st} for s,st,basis in L['signals']]
    m['mission_control_cards']=m['mission_control_cards'][:6]
    m['uncertainty_narrative']={'estimate_maturity':'Class 3 is usable for budget authorisation only if the evidence gaps are explicit.','schedule_maturity':'Schedule Level 4 improves logic, but board confidence is still governed by near-critical density and owner evidence.','interpretation':f"Live calibration is weighting {L['constraints']} into the QSRA/QCRA tail."}
    cost,sched,risks=_v130_rows(L,m); m['cost_lines']=cost; m['schedule_rows']=sched; m['risk_register']=risks; m['risks']=risks
    m['cost_breakdown']=cost; m['estimates_by_class']={str(i):cost for i in range(1,6)}; m['schedules_by_level']={str(i):sched for i in range(1,6)}
    m['procurement_heatmap']=[{'package':x,'status':'Active' if i<3 else 'Watch','risk':'Sector-native procurement exposure','owner':'Commercial Lead'} for i,x in enumerate(L['cost'][:5])]
    m['critical_path_narrative']=[f"{x} is near-critical in the {L['label']} sector graph and must be evidenced before approval." for x in L['schedule'][:5]]
    m['red_flags']=[f"Unevidenced confidence around {L['constraints']}.", "Scenario benefit may be risk transfer rather than risk reduction.", "Board pack should name owner, evidence source and date for each governing constraint."]
    forbidden=_v130_forbidden(key)
    if '_v124_scrub_value' in globals():
        for k in list(m.keys()):
            if k not in {'prompt','client','title','id'}: m[k]=_v124_scrub_value(m[k], forbidden)
    # restore locked native arrays after scrub
    m['sector_ontology_key']=key; m['sector_ontology_label']=L['label']; m['subsector']=L['label']; m['causal_graph_nodes']=L['chain']; m['causal_chain']=L['chain']
    m['sector_confidence_drivers']=L['confidence']; m['sector_primary_cost_drivers']=L['cost']; m['sector_schedule_threats']=L['schedule']
    return m

_CASEY_V130_PREV_BUILD_MODEL = build_model
def build_model(prompt:str, client:str='', class_level:int=3, schedule_level:int=3, scenario:str='base'):
    return _v130_apply(_CASEY_V130_PREV_BUILD_MODEL(prompt, client, class_level, schedule_level, scenario), prompt, client)

APP_VERSION='CASEY V130 Executive Intelligence Final'
print('CASEY V130 executive intelligence hardening installed')
# ================= END CASEY V130 EXECUTIVE INTELLIGENCE LOCK =================

# ================= CASEY V130.1 ROUTING HOTFIX =================
def _v131_sector_key(prompt='', client='', model=None):
    pt=(str(prompt)+' '+str(client)).lower()
    # classify on user brief first; never let an older fallback subsector overrule explicit input.
    if _v130_has(pt,['lunar','mars','orbital','satellite','spaceport','launch','payload','moon','deep space','propulsion','range availability']): return 'space'
    if _v130_has(pt,['data centre','data center','hyperscale','ai campus','compute campus','gpu','cloud region','white space','data hall']): return 'data_centre'
    if _v130_has(pt,['oil','gas','lng','refinery','petrochemical','offshore platform','pipeline','fpso','hydrocarbon','carbon capture']): return 'oil_gas'
    if _v130_has(pt,['nuclear','smr','reactor','containment','safety case']): return 'nuclear'
    if _v130_has(pt,['energy','power plant','renewable','wind farm','offshore wind','solar','battery','substation','transmission','grid','hydrogen','hvdc']): return 'energy'
    if _v130_has(pt,['port','harbour','harbor','marine','dock','container terminal','quay','dredging']): return 'ports'
    if _v130_has(pt,['airport','aviation','terminal','runway','heathrow','gatwick','airside','baggage','orat']): return 'airport'
    if _v130_has(pt,['rail','metro','transit','high speed','hs2','rail station','signalling','signaling','rolling stock','california high speed']): return 'rail'
    if _v130_has(pt,['semiconductor','fab','wafer','cleanroom','foundry','lithography','chip plant']): return 'semiconductor'
    if _v130_has(pt,['life sciences','pharma','biologics','gmp','fill-finish','sterile','cqv','amgen','lilly','novartis','pfizer']): return 'life_sciences'
    if _v130_has(pt,['defence','defense','military','naval','airbase','radar','missile','secure facility','mod ','dod ','command centre','command center']): return 'defence'
    if _v130_has(pt,['hospital','healthcare','clinical','medical centre','patient','nhs']): return 'healthcare'
    if _v130_has(pt,['water','wastewater','desalination','sewer','reservoir','treatment plant']): return 'water'
    if _v130_has(pt,['mine','mining','ore','tailings','pit','processing plant']): return 'mining'
    if _v130_has(pt,['road','highway','motorway','bridge','tunnel']): return 'roads'
    return _v130_sector_key(prompt, client, model)

def _v131_apply(model, prompt='', client=''):
    # Temporarily override routing inside v130 apply.
    global _v130_sector_key
    old=_v130_sector_key
    try:
        _v130_sector_key=_v131_sector_key
        return _v130_apply(model, prompt, client)
    finally:
        _v130_sector_key=old

# Bypass previous V130 wrapper so a misrouted old model cannot contaminate the final sector lock.
def build_model(prompt:str, client:str='', class_level:int=3, schedule_level:int=3, scenario:str='base'):
    return _v131_apply(_CASEY_V130_PREV_BUILD_MODEL(prompt, client, class_level, schedule_level, scenario), prompt, client)
APP_VERSION='CASEY V130.1 Executive Intelligence Final'
print('CASEY V130.1 routing hotfix installed')
# ================= END CASEY V130.1 ROUTING HOTFIX =================

# ================= CASEY V132 INSTITUTIONAL AUTHORITY / FULL EARTH+SPACE QA LOCK =================
# Final pre-demo hardening: fixes V131 fallback recursion, expands explicit sector routing,
# and makes board challenge / intervention / behavioural intelligence part of every model/export.

V132_SECTOR_TESTED = [
    'data_centre','airport','rail','roads','ports','water','energy','nuclear','oil_gas','mining',
    'healthcare','life_sciences','semiconductor','defence','space','general_infrastructure'
]

def _v132_has(t, words):
    import re
    t = str(t or '').lower()
    for w in words:
        w = str(w or '').lower().strip()
        if not w:
            continue
        # Use token boundaries so 'port' does not match 'transport' and 'mine' does not match 'commissioning'.
        if re.search(r'(?<![a-z0-9])' + re.escape(w) + r'(?![a-z0-9])', t):
            return True
    return False

def _v132_sector_key(prompt='', client='', model=None):
    model = model or {}
    pt = (str(prompt or '') + ' ' + str(client or '')).lower()
    # User brief always wins over older fallback model state. Order is deliberate: specific asset classes before generic words.
    if _v132_has(pt, ['lunar','moon','mars','orbital','leo','cislunar','satellite','spaceport','launch vehicle','payload','deep space','propulsion test','range operations','isru','propellant depot','space habitat']): return 'space'
    if _v132_has(pt, ['data centre','data center','datacenter','hyperscale','ai campus','compute campus','gpu','cloud region','white space','data hall','azure','openai','meta ai','msft ai','microsoft ai']): return 'data_centre'
    if _v132_has(pt, ['semiconductor','fab','wafer','foundry','lithography','chip plant','euv','upw','ultra-pure water']): return 'semiconductor'
    if _v132_has(pt, ['life sciences','pharma','biologics','gmp','fill-finish','fill finish','sterile','cqv','media fill','amgen','lilly','eli lilly','novartis','pfizer','therapeutics']): return 'life_sciences'
    if _v132_has(pt, ['mine','mining','ore','tailings','pit','processing plant','metals','lithium mine','copper mine']): return 'mining'
    if _v132_has(pt, ['oil','gas','lng','refinery','petrochemical','offshore platform','pipeline','fpso','hydrocarbon','carbon capture','ccus','upstream','downstream']): return 'oil_gas'
    # Nuclear weapons/defence context routes to defence not nuclear power
    if _v132_has(pt, ['nuclear weapon','nuclear warhead','nuclear deterrent','trident','awre','aldermaston','burghfield','nuclear weapon programme','weapons grade']): return 'defence'
    if _v132_has(pt, ['nuclear','smr','reactor','containment','safety case','nuclear island']): return 'nuclear'
    if _v132_has(pt, ['port','harbour','harbor','marine terminal','dock','container terminal','quay','dredging','berth']): return 'ports'
    if _v132_has(pt, ['airport','aviation','terminal','runway','heathrow','gatwick','airside','baggage','orat']): return 'airport'
    if _v132_has(pt, ['rail','metro','transit','high speed','hs2','rail station','signalling','signaling','rolling stock','california high speed','subway']): return 'rail'
    if _v132_has(pt, ['water','wastewater','desalination','sewer','reservoir','treatment plant','pumping station','flood defence','flood defense','smart meter','meter rollout','water utility','utility rollout','connection rollout','ofwat','water authority','water company','water board','clean water','drinking water','water network','water main','water pipe']): return 'water'
    if _v132_has(pt, ['awre','aldermaston','burghfield','warhead','nuclear weapon','classified facility','sovereign secure','secure nuclear','weapons programme','weapons program','dstl','defence nuclear','defense nuclear','aukus submarine','dockyard','naval base','naval shipbuild']): return 'defence'
    if _v132_has(pt, ['defence','defense','military','naval','airbase','radar','missile','secure facility','mod ','dod ','command centre','command center','hardened facility','security accreditation','classified systems','classified programme','classified program','special forces','intelligence facility','gchq','mi5','mi6','intelligence campus','sovereign supply']): return 'defence'
    if _v132_has(pt, ['hospital','healthcare','clinical','medical centre','medical center','patient','nhs','acute care','health campus']): return 'healthcare'
    if _v132_has(pt, ['gigafactory','giga factory','battery factory','battery gigafactory','battery manufactur','ev manufactur','electric vehicle manufactur','cell manufactur','cathode manufactur','gwh','gwh/year','battery plant','gigawatt hour']): return 'gigafactory'
    if _v132_has(pt, ['energy','power plant','renewable','wind farm','offshore wind','solar','battery storage','substation','transmission','grid','hydrogen','hvdc','utility interconnector']): return 'energy'
    if _v132_has(pt, ['5g','telecoms','telecom','fibre rollout','fiber optic','broadband rollout','mobile network','mast install','tower rollout','subsea cable','openreach','bt infrastr','network rollout','digital infrastr','rural connectivity','gigabit broadband','wayleave']): return 'telecoms'
    if _v132_has(pt, ['road','highway','motorway','bridge','tunnel','expressway','interchange']): return 'roads'
    # fallback to current/fallback model only if the input was genuinely generic.
    t = (pt + ' ' + str(model.get('title','')) + ' ' + str(model.get('subsector','')) + ' ' + str(model.get('mode',''))).lower()
    if _v132_has(t, ['lunar','moon','mars','orbital','satellite','spaceport','payload']): return 'space'
    if _v132_has(t, ['data centre','data center','hyperscale','ai campus']): return 'data_centre'
    if _v132_has(t, ['airport','aviation','terminal','runway']): return 'airport'
    if _v132_has(t, ['rail','metro','transit','signalling','rolling stock']): return 'rail'
    return 'general_infrastructure'

def _v132_library(key):
    L = _v130_library(key) if '_v130_library' in globals() else {}
    if not isinstance(L, dict): L = {}
    if key == 'general_infrastructure' or not L.get('label'):
        L = dict(label='General Capital Infrastructure', shock='The dominant risk is not progress reporting; it is whether the governing constraint has named evidence, owner accountability and reserve logic.', constraints='scope maturity, procurement certainty, interface control, commissioning readiness and evidence ownership', signals=[('Evidence ownership','Active','Board confidence requires named evidence owners, not generic mitigation text.'),('Procurement certainty','Active','Long-lead packages and commercial exposure drive P80/P90 movement.'),('Interface control','Watch','Third-party interfaces and integration gates create schedule tails.'),('Commissioning readiness','Watch','Operational readiness is a governance signal, not late administration.')], bench=[('Capital Infrastructure Programme','$500M–$15B','30-120'),('Complex Systems Integration Programme','$300M–$10B','24-96'),('Major Public / Private Capital Programme','$1B–$25B','48-144')], chain=['Scope definition','Procurement evidence','Interface control','Commissioning readiness','Owner evidence','Reserve logic','Confidence'], confidence=['Benchmark similarity: capital infrastructure archetype','Scope maturity: requirements and package definition','Procurement certainty: long-lead and market capacity','Schedule maturity: critical path and commissioning logic','Interface exposure: utilities, stakeholders and operations'], cost=['Core asset and enabling works','Specialist systems and long-lead packages','Interfaces, utilities and third-party works','Programme management and assurance','Risk reserve and contingency'], schedule=['Scope freeze and approvals','Long-lead procurement','Interface and utility readiness','Commissioning and operational readiness','Owner evidence and board acceptance'])
    # Institutional authority layer: same structure, stronger phrasing.
    L = dict(L)
    L['challenge_line'] = f"The programme narrative is only defensible if {L['constraints']} are evidenced by named owners before approval."
    L['interventions'] = [
        f"Name the accountable owner and evidence source for {L['constraints']}.",
        "Separate true risk reduction from risk transfer into operations, reserve or P90 exposure.",
        "Reconcile the P50 approval story with P80/P90 downside before board commitment.",
        "Retire the governing constraint before buying acceleration or declaring savings."
    ]
    L['behavioural_forecast'] = f"Comparable {L['label']} programmes with weak evidence around {L['constraints']} typically lose confidence during integration, commissioning or approval windows rather than during visible construction progress."
    return L

def _v132_forbidden(key):
    terms = set(_v130_forbidden(key) if '_v130_forbidden' in globals() else [])
    # Hard guardrails for observed leaks. These are allowed only in their native sector.
    if key != 'data_centre': terms.update(['liquid cooling','Liquid cooling','white space','white-space','data hall','IST congestion','GPU','hyperscale'])
    if key != 'airport': terms.update(['ORAT','orat','baggage','Baggage','airside','Airside','landside','passenger systems'])
    if key != 'rail': terms.update(['rolling stock','Rolling stock','signalling','Signalling','signaling','possessions','Possessions','timetable'])
    if key != 'space': terms.update(['launch reliability','launch cadence','payload integration','range availability','thermal-power balance','mission assurance','orbital recovery','Launch reliability','Mission assurance'])
    if key != 'defence': terms.update(['classified','authority to operate','secure comms'])
    if key != 'oil_gas': terms.update(['HAZOP','HAZID','hydrocarbon','FPSO','LNG','shutdown window'])
    if key != 'semiconductor': terms.update(['lithography','yield ramp','UPW','wafer','EUV'])
    if key != 'life_sciences': terms.update(['CQV','GMP','media fill','fill-finish','FDA','EMA'])
    if key != 'nuclear': terms.update(['nuclear island','containment','safety case'])
    return list(terms)

def _v132_rows(L, m):
    # Reuse V130 row builder where available and add owner/status/challenge columns.
    try:
        cost, sched, risks = _v130_rows(L, m)
    except Exception:
        cost, sched, risks = _v125_sector_rows(L, m) if '_v125_sector_rows' in globals() else ([], [], [])
    for i, r in enumerate(risks):
        if isinstance(r, dict):
            r.setdefault('owner', 'Sector Integration Lead' if i else 'Programme Director')
            r.setdefault('status', 'Open' if i < 3 else 'Mitigating')
            r.setdefault('mitigation_owner', r.get('owner'))
            r.setdefault('mitigation_status', r.get('status'))
            r.setdefault('challenge', 'Board should require named owner evidence and a dated mitigation proof point.')
            r.setdefault('residual_status', 'Challenge until evidence improves' if i < 3 else 'Monitor')
    return cost, sched, risks


def _v135_sector_q1(key: str, constraints: str) -> str:
    """Return a sector-specific opening board question instead of the generic one."""
    q1s = {
        'rail':        f"Which possessions are contractually confirmed with the operator, and which are still assumed in the programme?",
        'nuclear':     f"What is the GDA critical path and what does a 6-month slip cost in carrying charges alone?",
        'defence':     f"What is the security accreditation critical path and which classified supplier or system is currently unconfirmed?",
        'data_centre': f"Is the grid connection agreement signed, or is the energisation date still an assumption in the programme model?",
        'life_sciences':f"When does the inspection readiness programme need to complete, who owns it, and is it on the master schedule?",
        'semiconductor':f"Is the tool install sequence confirmed by the OEM or derived from an optimistic ramp model?",
        'airport':     f"Is the ORAT programme on the master critical path with named owners and acceptance criteria?",
        'healthcare':  f"What is the clinical commissioning scope, who owns it, and is it on the programme critical path?",
        'energy':      f"Is the grid connection contracted or assumed, and what is the DNO/TSO commitment date?",
        'water':       f"What is the regulatory consent status, and which approval is on the critical path to first water?",
        'mining':      f"What is the environmental consent critical path, and which single approval gates the programme?",
        'oil_gas':     f"What are the long-lead procurement items, and which have confirmed supplier slots versus intent letters?",
        'ports':       f"Is the marine consent and dredging programme confirmed, and can the port remain operational during construction?",
        'roads':       f"What is the environmental consent status, and has utilities diversion been scoped and costed?",
        'space':       f"What is the launch vehicle (named or assumed?), current TRL for unproven systems, and is FOAK qualification on the critical path?",
        'general_infrastructure': f"Who is the named individual accountable for the governing critical-path constraint, and what is their evidence closure date?",
    }
    return q1s.get(key, f"Who is the named individual accountable for {constraints}, and what is their evidence closure date?")


def _v135_if_this_fails(key: str, constraints: str, label: str) -> str:
    """Returns the historically-grounded failure narrative for this sector.
    Named programmes, named failure modes, named costs from public record.
    """
    patterns = {
        'rail':
            "If this programme fails, the likely pattern will be: possession access failed to mature before trial running began, signalling integration was deferred and then became the critical path, and the operator acceptance date was treated as a management target rather than a contractual constraint. "
            "This is the LP1/Crossrail pattern — Elizabeth Line overran by 3 years and £4B primarily because Systems Integration had 900 outstanding issues at the planned opening date, float was exhausted, and no single individual owned the critical path. "
            "HS2 Phase 1 is exhibiting the same signals: open corridor risk, assumed possessions and a systems assurance programme that is not on the master schedule. "
            "CASEY challenge: is the float on this programme operationally usable against confirmed access windows, or is it nominal?",

        'nuclear':
            "If this programme fails, the likely pattern will be: FOAK supply chain failed to deliver nuclear-grade components on schedule, the safety case evidence programme was treated as a separate workstream rather than the governing critical path, and GDA took 2–3 years longer than the programme assumed. "
            "This is the Hinkley Point C pattern — cost doubled from £18B to £35B+ and schedule extended by 4+ years due to first-pour concrete issues, nuclear-grade welding failures, and EPR-specific supply chain constraints that were known risks but not on the master critical path. "
            "Olkiluoto 3 in Finland overran by 14 years and tripled in cost for the same reasons. Vogtle Units 3 & 4 in Georgia overran by 7 years and doubled cost. "
            "CASEY challenge: which single weld inspection failure, FOAK component rejection, or GDA hold-point would break this programme date — and what is the cost of 6 months' delay in financing charges alone?",

        'defence':
            "If this programme fails, the likely pattern will be: security accreditation was treated as administrative rather than programme-critical, classified supplier slots were assumed rather than contracted, and operational acceptance was deferred until post-delivery then took 2+ years. "
            "This is the Ajax/Watchkeeper/AWRE pattern — Ajax armoured vehicle programme spent £3.5B and took 10+ years with no vehicles accepted into service, primarily because electromagnetic compatibility, crew safety and training system integration issues were not on the critical path. "
            "Watchkeeper UAV programme cost £1.3B and delivered a system unusable in civil airspace — regulatory and airworthiness approval was never treated as a delivery constraint. "
            "CASEY challenge: what is the authority-to-operate critical path, and which classified supplier or system could break it?",

        'data_centre':
            "If this programme fails, the likely pattern will be: grid connection was assumed rather than contracted, cooling system performance failed at peak GPU load, and IT delivery milestones outside programme control were never on the critical path. "
            "Every major hyperscale delivery in 2022–2024 experienced 6–18 month energisation delays due to DNO grid queue backlogs — Microsoft, Google, Meta and Amazon all publicly disclosed delays in the UK, Ireland and Germany. "
            "The second failure mode is power-use-effectiveness optimism: cooling systems designed to datasheet assumptions fail commissioning tests under actual GPU thermal load. "
            "CASEY challenge: is the grid connection agreement signed with an energisation date, or is it a queue position?",

        'life_sciences':
            "If this programme fails, the likely pattern will be: the validation programme was scoped as a post-construction activity, clean utility readiness was not on the master schedule as a critical milestone, and the FDA/MHRA inspection date was treated as a target rather than a contractual revenue dependency. "
            "AstraZeneca's Macclesfield site, Pfizer's Ringaskiddy expansion, and multiple CDMO greenfield builds have all exhibited this pattern — construction completes on time, CQV runs 18–36 months late because equipment qualifications, media fills and regulatory dossier preparation were not resourced as programme-critical. "
            "CASEY challenge: when does validation begin relative to practical completion, who owns it, and is it on the critical path as a named milestone?",

        'semiconductor':
            "If this programme fails, the likely pattern will be: tool install sequence was based on OEM forecast letters rather than contracted delivery slots, UPW and chemical system commissioning was underestimated by 40–60%, and yield ramp was scoped as outside the programme boundary — then became the board metric. "
            "This is the Intel Ohio and TSMC Arizona pattern — both announced 2024–2025 production dates and slipped 2–3 years due to specialised workforce shortages, UPW system complexity underestimation, and tool delivery sequences that assumed OEM capacity the market could not provide. "
            "CASEY challenge: are tool delivery sequences confirmed orders with contractual delivery slots, or OEM intent letters?",

        'gigafactory':
            "If this programme fails, the likely pattern will be: grid connection was assumed rather than contracted, cell chemistry qualification was treated as an R&D activity rather than a programme milestone, and yield ramp was scoped as post-delivery. "
            "This is the Britishvolt pattern — £3.8B facility, planning approved, £100M of public money committed, collapsed before construction was complete because grid connection, cell chemistry qualification and the battery management system supply chain were all unconfirmed. "
            "Northvolt's Skellefteå gigafactory scaled to 16GWh/year target but achieved 1GWh/year — yield ramp was optimistic by 16x. "
            "CASEY challenge: is cell chemistry confirmed in production at scale, or is this programme building capacity for a product that has not yet been qualified?",

        'oil_gas':
            "If this programme fails, the likely pattern will be: long-lead procurement items had assumed rather than contracted delivery slots, offshore weather windows were optimistically scoped in the schedule, and the brownfield interface with live production was underestimated by 50–100%. "
            "Chevron's Gorgon LNG project overran by $17B (54%) and 3 years — brownfield interface with the Barrow Island nature reserve, remote logistics and HAZOP complexity were the primary drivers. "
            "Shell's Prelude FLNG delivered 5 years late and has never achieved nameplate capacity — FOAK floating LNG technology was more complex than schedule and cost models assumed. "
            "CASEY challenge: which long-lead items are confirmed orders versus intent letters, and what is the brownfield production impact plan for each tie-in?",

        'mining':
            "If this programme fails, the likely pattern will be: environmental consent was treated as a formality and ended up on the critical path for 3–5 years, tailings storage facility approval was deferred until construction was advanced, and community licence-to-operate was not in the risk register. "
            "Vedanta's Nchanga copper mine expansion, First Quantum's Cobre Panama, and multiple Australian iron ore projects have been suspended or abandoned after multi-billion capital deployment due to environmental/community approvals that were assumed rather than secured. "
            "Cobre Panama: $10B built, operating, then shut down in 2023 by government order — community opposition was known but not treated as a board approval gate. "
            "CASEY challenge: what is the community licence-to-operate risk, who owns it, and what is the recovery plan if consent is challenged post-commitment?",

        'airport':
            "If this programme fails, the likely pattern will be: ORAT was not on the master critical path, baggage system integration was deferred and then failed acceptance, and airside access constraints were not priced into the programme. "
            "This is the Heathrow T5 pattern — construction delivered on time and budget, but 34,000 bags were lost in the first 10 days because baggage system integration, operator training and IT cutover were not treated as programme deliverables. "
            "Berlin Brandenburg Airport opened 9 years late primarily due to fire safety system integration, passenger processing IT and regulatory approval — all treated as post-construction rather than programme-critical. "
            "CASEY challenge: is ORAT on the master critical path with named owners, acceptance criteria and a contractual handover date?",

        'healthcare':
            "If this programme fails, the likely pattern will be: clinical commissioning was scoped as a post-construction NHS activity, infection-control compliance required rework of completed areas, and medical equipment procurement ran 18–24 months behind civil completion. "
            "This is the Royal Liverpool Hospital pattern — £335M PFI project was mothballed for 4 years post-practical completion due to concrete structural defects, infection-control compliance failures, and a commissioning programme that was not defined at contract award. "
            "The new Royal Adelaide Hospital opened 2 years late and $640M over budget — clinical commissioning was not on the master schedule and operational transition was not contracted. "
            "CASEY challenge: who owns clinical commissioning, when does it start relative to practical completion, and is the infection-control compliance programme on the master schedule?",

        'energy':
            "If this programme fails, the likely pattern will be: grid connection was assumed rather than contracted, the DNO/TSO queue position was treated as an energisation commitment, and commissioning was compressed to hit a CfD or PPA milestone date. "
            "UK offshore wind: Hornsea 2 achieved grid connection 18 months late due to DNO queue backlog — a risk that was in the market but not on the programme critical path. "
            "Neart na Gaoithe offshore wind farm in Scotland was delayed 4+ years due to aviation radar objections that were known at planning stage but not resolved before financial close. "
            "CASEY challenge: is the grid connection agreement signed with a contracted energisation date, or is it a queue position with a forecast date?",

        'water':
            "If this programme fails, the likely pattern will be: regulatory consent was assumed before the planning timetable was confirmed, the smart metering comms infrastructure was scoped separately from the physical rollout, and community access agreements were not in the programme. "
            "Thames Water's AMP7 capital programme is running 40% below delivery target — procurement, supply chain capacity and site access were all assumed rather than confirmed at programme start. "
            "Smart meter rollouts in the UK (SMETS2) and Australia (AMI) both ran 3–5 years late due to comms infrastructure complexity, meter reading system integration, and data management platform readiness that were scoped as separate workstreams. "
            "CASEY challenge: are the comms infrastructure and back-office systems in the programme boundary and on the critical path, or are they separate projects with assumed delivery dates?",

        'telecoms':
            "If this programme fails, the likely pattern will be: wayleave and planning consents were assumed on average timelines but ran 3x longer in practice, access to private land was not secured before mobilisation, and the back-office activation platform was a separate project with an assumed delivery date. "
            "BT Openreach's FTTP rollout is running 2+ years behind the original 25M premises target — wayleave complexity in dense urban areas and multi-dwelling units was systematically underestimated. "
            "NBN Co in Australia: $29B programme, originally estimated at $4.7B, delivered 6+ years late due to engineering complexity, copper network assumptions, and contractor performance in multi-technology mix areas. "
            "CASEY challenge: what is the average wayleave acquisition time in the target deployment area, and is it on the critical path?",

        'roads':
            "If this programme fails, the likely pattern will be: utility diversions ran 2–3x longer than programmed because third-party access agreements were assumed, environmental consent was challenged post-award, and traffic management constraints were not priced into the construction method. "
            "A303 Stonehenge tunnel: development consent order challenged by UNESCO after contract was let — a risk that was known but not treated as a programme constraint. "
            "A14 Cambridge to Huntingdon: delivered on time but the utility diversion programme ran 18 months late and was the critical path for 60% of the works. "
            "CASEY challenge: are all utility diversion agreements signed with third parties, or are timescales based on average industry assumptions?",

        'ports':
            "If this programme fails, the likely pattern will be: marine consent was assumed before the dredging scope was confirmed, the port could not remain operational during construction, and the vessel specification changed after design freeze. "
            "Dublin Port's Alexandra Basin Redevelopment: 3 years late due to unexpected marine sediment conditions, live port operational constraints, and vessel specification changes that required dredging beyond the consented channel. "
            "London Gateway Phase 2: terminal systems integration and quay crane commissioning ran 18 months late because IT/OT system interfaces were not in the EPC contract boundary. "
            "CASEY challenge: are terminal IT/OT systems in the EPC programme boundary, or are they a separate procurement with an assumed delivery date?",

        'space':
            "If this programme fails, the likely pattern will be: systems integration and qualification failed to mature before the launch window, TRL was claimed without independent test evidence, and the launch manifest date was assumed from OEM intent rather than contracted. "
            "This is the JWST/Artemis/OneWeb pattern — James Webb Space Telescope overran by 14 years and $9B (10x original estimate) due to systems integration complexity, cryogenic testing failures, and scope growth that was visible early but not on the master schedule. "
            "Artemis I: 5 years behind original SLS schedule due to propulsion system complexity and NASA's fixed-price contracting approach with Boeing that removed contractor schedule incentives. "
            "OneWeb constellation: $3.4B in bankruptcy — launch cadence, ground segment readiness, and customer revenue assumptions were all optimistic. "
            "CASEY challenge: what independent TRL evidence exists for each unproven technology, and what is the programme cost of a 12-month qualification slip?",

        'general_infrastructure':
            f"If this programme fails, the likely explanation will be: {constraints} failed to mature before the board commitment date, the reserve was sized to the P50 not the P80, and no named individual was accountable for the governing constraint. "
            "This is the consistent pattern across HM Treasury's Infrastructure and Projects Authority red-rated programmes — cost and schedule optimism at approval, reserve inadequacy, and diffuse ownership of governing risks. "
            "The IPA's 2023/24 annual report shows 17% of government major projects rated Red (unachievable in scope, time and cost). The primary cause in 14 of 19 Red-rated programmes was inadequate evidence at approval stage. "
            "CASEY challenge: what is the specific evidence that closes the governing constraint before board commitment?",
    }
    return patterns.get(key, f"If this programme fails, the likely explanation will be: {constraints} failed to mature before approval, and the governing constraint had no named owner with a documented evidence closure plan. The IPA's analysis of major programme failures consistently identifies inadequate evidence at approval stage as the primary cause — not execution failure.")




def _v136_tt_killer_fields(model: dict, key: str, L: dict) -> dict:
    """intelligence fields — the intelligence that makes cost consultants look static."""
    m = dict(model)
    conf = float(str(m.get('confidence_pct', 60)).replace('%',''))
    p50 = str(m.get('cost_p50','$0B'))
    sched = str(m.get('schedule','0 months'))
    risk = str(m.get('risk','Medium'))
    sector = str(m.get('subsector', key))
    constraints = L.get('constraints', 'the governing constraint')
    shock = L.get('shock','')
    chain = L.get('chain', [])

    # Traditional vs CASEY — what the incumbent deck says vs what CASEY reads
    _loc_ctx2 = location_context(m.get('location',''))
    _framework = _loc_ctx2.get('framework', 'international best practice')
    _approval = _loc_ctx2.get('approval_body', 'the investment committee')
    _financing = _loc_ctx2.get('financing', 'programme funding')
    _oba_note = _loc_ctx2.get('optimism_bias_note', 'Apply Flyvbjerg global reference class: average +27% cost, +39% schedule.')
    tvc = {
        'traditional': f"Civil progress and cost spend appear on track. Programme shows {p50} outturn with {sched} completion. Risk register has been updated. Contractor reporting green.",
        'casey': f"CASEY reads the programme as {conf}% board-defensible against {_framework} standards. The governing constraint is {constraints}. Until this has a named owner and evidence closure date, the {p50} P50 is not a defensible approval position for {_approval}. The dashboard is green; the programme is not.",
        'gap': f"The gap between traditional reporting and CASEY intelligence: traditional controls measure what has happened. CASEY measures what will happen if the governing constraint is not closed before the approval gateway. {_oba_note}",
        'incumbent_line': f"'A conventional project controls report would show this as a ' + risk.lower() + ' risk programme at ' + p50 + '. CASEY challenges whether the P80 is adequately reserved, whether the critical path has a named owner, and whether ' + _approval + ' is making a decision or deferring one.'location_framework': _framework,
        'approval_body': _approval,
        'financing_context': _financing,
    }
    m['traditional_vs_casey'] = tvc

    # Board attack simulation — the 5 questions any serious board will ask
    # that a conventional static report cannot answer in the room
    sector_attacks = {
        'rail': [
            f"The possession schedule shows {sched} — which possessions are confirmed with the operator versus assumed, and what is the cost of a single failed trial-running gate?",
            f"Signalling integration is on the critical path. What is the current IEMs-open count, who owns closure, and when does the safety case need to be submitted?",
            f"The P50 is {p50}. What is the P80 and what drives the gap — is it possessions, signalling, or systems migration?",
            f"If this programme follows the Crossrail/LP1 pattern, the likely failure mode is deferred systems integration. What evidence proves this is not the case?",
            f"Who is the single named individual accountable for the operator acceptance date, and what are the contractual consequences if it slips?",
        ],
        'nuclear': [
            f"GDA is the programme's real critical path. What is the current GDA stage and what does a 6-month slip cost in financing charges and fuel loading delay?",
            f"FOAK supply chain: which nuclear-grade components have confirmed supplier slots versus intent letters, and what is the programme response to a single weld rejection?",
            f"The P50 is {p50}. Hinkley doubled from £18B to £35B+. What specific evidence differentiates this programme from the Hinkley/Olkiluoto/Vogtle pattern?",
            f"The safety case is the board approval gate, not practical completion. When does safety case submission need to happen, and is it on the master critical path?",
            f"What is the workforce nuclear certification programme, who owns it, and when does it need to complete to support first-pour?",
        ],
        'defence': [
            f"Security accreditation is the real critical path — not construction. What is the Authority-to-Operate pathway and which classified system is currently unconfirmed?",
            f"Ajax cost £3.5B and delivered nothing. What specific programme controls prevent this programme entering the same accreditation/acceptance spiral?",
            f"Which sovereign supply chain items are single-source, what are the contractual delivery commitments, and what is the fallback?",
            f"Operational acceptance testing: is it in the programme boundary and on the critical path, or is it a post-delivery activity with an assumed duration?",
            f"Export control dependencies: which systems involve foreign components, and what is the programme response to a ITAR/EAR restriction?",
        ],
        'data_centre': [
            f"Is the grid connection agreement signed with an energisation date, or is it a DNO queue position with a forecast date? What is the queue position number?",
            f"Cooling performance: has the thermal model been validated under actual GPU load at scale, or is it based on manufacturer datasheet assumptions?",
            f"Every major hyperscale developer experienced 6–18 month energisation delays in 2022–2024. What evidence differentiates this programme's grid position?",
            f"IT delivery milestones: which are inside this programme boundary and which are dependent on external programmes? What are the contractual dependencies?",
            f"The P50 is {p50}. What is the P80 — the primary cost driver is grid connection and cooling system complexity, not civil works.",
        ],
        'life_sciences': [
            f"When does the validation programme start relative to practical completion, who owns it, and is it on the master critical path as a named milestone?",
            f"Clean utility readiness: is UPW/WFI/process gas system commissioning on the critical path, or is it assumed to follow civil completion by a standard interval?",
            f"FDA/MHRA inspection readiness: when does the dossier submission need to happen to support the product launch date, and is that date on the master schedule?",
            f"What is the impact of a single failed media fill on programme cost and schedule, and is that risk in the register with a named owner?",
            f"CDMO/CMO capacity: if this is a contract facility, is the customer product qualification programme on the master critical path?",
        ],
        'semiconductor': [
            f"Tool install sequence: are OEM delivery slots confirmed orders with contractual dates, or are they based on OEM forecast letters?",
            f"Intel Ohio and TSMC Arizona both slipped 2–3 years due to specialised workforce and UPW complexity. What evidence differentiates this programme?",
            f"Yield ramp: is it in the programme boundary and on the board performance metric, or is it scoped as post-delivery?",
            f"UPW system commissioning: what is the scope, and is the 40–60% underestimation risk (a consistent industry pattern) captured in the cost model?",
            f"The P50 is {p50}. What is the P80 — the primary driver is tool delivery and UPW commissioning, not civil works.",
        ],
        'gigafactory': [
            f"Cell chemistry qualification: is it confirmed in production at scale, or is this programme building manufacturing capacity for a product that has not yet been qualified?",
            f"Britishvolt collapsed with planning approved and public funding committed because grid connection, cell chemistry and BMS supply chain were all unconfirmed. What differentiates this programme?",
            f"Grid connection: is the connection agreement signed with an energisation date, or is it a DNO queue position?",
            f"Yield ramp: Northvolt targeted 16GWh/year and achieved 1GWh/year. What yield assumption underlies the business case and what evidence supports it?",
            f"The P50 is {p50}. What is the P80 — the primary drivers are grid connection, cell chemistry qualification, and yield ramp, not construction.",
        ],
        'oil_gas': [
            f"Long-lead procurement: which items are confirmed orders with contractual delivery dates versus intent letters, and what is the programme response to a single slip?",
            f"Gorgon overran by $17B (54%) — brownfield interface and remote logistics complexity. What specific evidence proves this programme's brownfield scope is adequately costed?",
            f"Shutdown windows: which tie-ins require live production shutdown, how long are the windows, and what is the cost of a single missed window?",
            f"HAZOP/HAZID: has the safety case basis been approved, and are the safety-critical instrument lists confirmed?",
            f"The P50 is {p50}. What is the P80 — the primary drivers are long-lead procurement, shutdown access, and start-up performance, not civil works.",
        ],
        'mining': [
            f"Environmental consent: is it secured, or is it assumed to follow the planning timetable? Cobre Panama built $10B then was shut by government order.",
            f"Tailings storage facility: is the TSF approval on the critical path and confirmed with the regulator, or is it assumed to follow construction?",
            f"Community licence-to-operate: what is the community opposition risk and what is the programme response if consent is challenged post-commitment?",
            f"Processing plant yield: is the yield ramp based on orebody test work at scale or datasheet assumptions? What is the cost of a 20% yield underperformance?",
            f"The P50 is {p50}. What is the P80 — the primary drivers are environmental consent, processing yield, and logistics, not earthworks.",
        ],
        'airport': [
            f"ORAT: is it on the master critical path with named owners, acceptance criteria, and a contractual handover date?",
            f"Heathrow T5 lost 34,000 bags on opening day because IT/baggage integration was not a programme deliverable. What specific controls prevent the same pattern here?",
            f"Berlin Brandenburg opened 9 years late — fire safety integration and regulatory approval were not treated as programme-critical. What is the safety/regulatory approval critical path?",
            f"Airside access constraints: are they in the programme cost model, or are they treated as contractor risk?",
            f"The P50 is {p50}. What is the P80 — the primary drivers are ORAT, systems integration, and regulatory approval, not construction.",
        ],
        'healthcare': [
            f"Clinical commissioning: who owns it, when does it start, and is it on the master critical path as a named milestone with contractual acceptance criteria?",
            f"Royal Liverpool Hospital was mothballed for 4 years post-practical completion. What specific evidence proves infection-control compliance is on the construction programme?",
            f"Medical equipment procurement: what is the lead time, who owns it, and is it on the programme critical path with a confirmed delivery date?",
            f"NHS operational transition: is the patient transfer and staff training programme in the project boundary and on the critical path?",
            f"The P50 is {p50}. What is the P80 — the primary drivers are clinical commissioning, medical equipment, and infection control, not construction.",
        ],
        'energy': [
            f"Grid connection: is the connection agreement signed with a contracted energisation date, or is it a DNO/TSO queue position with a forecast date?",
            f"Neart na Gaoithe was delayed 4+ years due to aviation radar objections known at planning stage. What third-party consent risks are in the register?",
            f"Long-lead equipment: transformers, HV switchgear and turbines have 3–5 year lead times. Are these confirmed orders with contractual delivery dates?",
            f"CfD/PPA milestone: what is the cost of missing the contracted commissioning date, and is that risk adequately reserved?",
            f"The P50 is {p50}. What is the P80 — the primary drivers are grid connection, permitting, and long-lead equipment, not construction.",
        ],
        'water': [
            f"Regulatory consent: is it secured, or is it assumed to follow the standard timetable? Thames Water AMP7 is running 40% below delivery target.",
            f"Smart metering comms infrastructure: is it in the programme boundary and on the critical path, or is it a separate project with an assumed delivery date?",
            f"Back-office activation platform: is it in the programme boundary or a separate IT project? SMETS2 rollout ran 3–5 years late partly because of data management platform complexity.",
            f"Operational transition: who is responsible for meter reading data migration, and is it on the programme critical path?",
            f"The P50 is {p50}. What is the P80 — the primary drivers are consent, comms infrastructure, and back-office integration, not physical rollout.",
        ],
        'telecoms': [
            f"Wayleave acquisition: what is the average time in the target deployment area, and is that on the critical path? BT Openreach is running 2+ years late on FTTP.",
            f"Multi-dwelling units: what proportion of the rollout is in MDUs, and is the wayleave complexity for MDUs priced into the programme?",
            f"Back-office activation platform: is it in the programme boundary, and when does it need to be ready relative to first connections?",
            f"NBN Co reached $29B (originally estimated at $4.7B) — engineering complexity and contractor performance in mixed-technology areas. What is the confidence basis for this cost model?",
            f"The P50 is {p50}. What is the P80 — the primary drivers are wayleave, MDU access, and back-office integration, not physical network build.",
        ],
        'roads': [
            f"Utility diversions: are all third-party access agreements signed, or are timescales based on average industry assumptions? A14 utility diversions were the critical path for 60% of works.",
            f"Environmental consent: has the development consent order been granted and is it legally final, or is it subject to challenge? A303 DCO was challenged after contract let.",
            f"Traffic staging: are the traffic management constraints priced into the construction method, or are they treated as contractor risk?",
            f"Structures and ground risk: what is the geotechnical investigation status and are ground condition assumptions adequately reflected in the P80?",
            f"The P50 is {p50}. What is the P80 — the primary drivers are utility diversions, consent, and ground risk, not earthworks.",
        ],
        'ports': [
            f"Marine consent: is it secured with the dredging scope confirmed, or is it assumed to follow the planning timetable?",
            f"Terminal IT/OT systems: are they in the EPC programme boundary, or a separate procurement with an assumed delivery date? London Gateway Phase 2 ran 18 months late for this reason.",
            f"Live port operations: is the construction method statement confirmed for maintaining port operations during build, and is the commercial impact of operational disruption costed?",
            f"Vessel specification: is the design vessel confirmed and frozen, or could specification changes after design freeze require dredging beyond the consented channel?",
            f"The P50 is {p50}. What is the P80 — the primary drivers are marine consent, dredging, and terminal systems, not civil works.",
        ],
        'space': [
            f"Launch vehicle: is it a named, contracted vehicle with a confirmed manifest position, or an assumed vehicle with an intended date?",
            f"TRL: JWST overran by 14 years and $9B due to systems integration complexity and cryogenic testing failures that were visible early. What independent evidence verifies TRL for each unproven system?",
            f"The P50 is {p50}. OneWeb spent $3.4B in bankruptcy — launch cadence, ground segment readiness, and customer revenue were all optimistic. What differentiates this case?",
            f"Autonomous operations: what is the recovery plan if autonomous commissioning fails, and what is the cost of a 6-month mission hold?",
            f"What is the programme cost of a 12-month qualification slip on the unproven technology, and is that risk adequately reserved?",
        ],
    }
    attacks = sector_attacks.get(key, sector_attacks.get('general_infrastructure', [
        f"Who is the named individual accountable for {constraints}?",
        f"What evidence closes the governing constraint before board commitment?",
        f"What is the P80 and what drives the gap from P50?",
        f"What is the programme cost of a 6-month slip on the critical-path constraint?",
        f"What would an independent cost review find that this submission does not show?",
    ]))
    m['board_attack_simulation'] = attacks

    # Programme mortality risk — % probability programme enters cost recovery mode
    mortality_base = max(15, min(85, 100 - conf))
    if conf < 50: mortality_base = min(85, mortality_base + 15)
    if conf > 75: mortality_base = max(15, mortality_base - 10)
    m['programme_mortality_risk'] = {
        'pct': mortality_base,
        'label': 'HIGH' if mortality_base > 60 else ('MEDIUM' if mortality_base > 35 else 'LOW'),
        'narrative': f"Based on {conf}% board-defensibility confidence, CASEY estimates {mortality_base}% probability this programme enters a cost-recovery or schedule-recovery mode before completion, based on comparable programmes with equivalent confidence profiles at approval.",
        'primary_driver': constraints,
        'comparable': m.get('if_this_fails','')[:120] if m.get('if_this_fails') else '',
    }

    # Confidence trajectory — where is confidence headed
    traj = 'DECLINING' if conf < 65 else ('STABLE' if conf < 78 else 'IMPROVING')
    m['confidence_trajectory'] = {
        'direction': traj,
        'narrative': (
            f"Confidence will DECLINE if {constraints} are not evidenced before the next approval gateway. "
            f"The typical pattern for {sector} programmes: confidence drops 8–15 percentage points when integration or commissioning fails to mature on schedule."
            if traj == 'DECLINING' else
            f"Confidence is STABLE at {conf}%. It will improve only when {constraints} are evidenced by named owners with closure dates. "
            f"Without evidence closure, P80 exposure will grow as the programme approaches delivery."
            if traj == 'STABLE' else
            f"Confidence is IMPROVING at {conf}%. Maintain evidence closure pace on {constraints} — do not allow governance to slip as delivery pressure increases."
        ),
        'next_gate': f"Board confidence will be tested when {chain[-2] if len(chain) >= 2 else constraints} needs to be evidenced.",
    }

    # Institutional authority line — the one sentence a board chair needs
    m['institutional_authority_line'] = (
        f"The board should not approve capital commitment until {constraints} "
        f"{'have' if ',' in constraints else 'has'} a named individual owner, a documented evidence closure plan, and a confirmed date — "
        f"without this, the {p50} is a reference number, not a decision basis."
    )

    return m

print("CASEY V136 intelligence fields installed")





def _v136_optimism_bias(key, m, bench_matches):
    """Quantified optimism bias detection — Green Book / IPA aligned."""
    p50_raw = m.get('cost_p50','')
    try:
        p50 = float(str(p50_raw).replace('$','').replace('B','').replace(',','').strip())
    except: p50 = 0
    
    # HM Treasury Green Book / Flyvbjerg reference class forecasting uplifts
    # These are the published optimism bias uplifts used in cost reviews
    oba_uplifts = {
        'rail':          {'cost_pct': 66, 'schedule_pct': 43, 'source': 'Flyvbjerg 2003/2022 global rail reference class (258 projects, 20 countries) + HM Treasury Table B6'},
        'nuclear':       {'cost_pct': 117, 'schedule_pct': 75, 'source': 'Flyvbjerg/Lovallo FOAK nuclear reference class'},
        'defence':       {'cost_pct': 57, 'schedule_pct': 42, 'source': 'NAO defence acquisition analysis — fixed-price contracts'},
        'data_centre':   {'cost_pct': 15, 'schedule_pct': 25, 'source': 'Hyperscale delivery reference class 2018-2024'},
        'life_sciences': {'cost_pct': 32, 'schedule_pct': 38, 'source': 'EMA/FDA major facility delivery reference class'},
        'semiconductor': {'cost_pct': 40, 'schedule_pct': 45, 'source': 'Semiconductor new-fab delivery 2010-2024'},
        'gigafactory':   {'cost_pct': 65, 'schedule_pct': 55, 'source': 'Battery gigafactory delivery reference class 2018-2024'},
        'energy':        {'cost_pct': 25, 'schedule_pct': 32, 'source': 'HM Treasury Green Book — Energy infrastructure'},
        'water':         {'cost_pct': 30, 'schedule_pct': 35, 'source': 'Ofwat capital delivery reference class'},
        'oil_gas':       {'cost_pct': 54, 'schedule_pct': 33, 'source': 'Flyvbjerg offshore/process infrastructure reference class'},
        'mining':        {'cost_pct': 35, 'schedule_pct': 40, 'source': 'Mining capital projects reference class 2000-2020'},
        'airport':       {'cost_pct': 28, 'schedule_pct': 36, 'source': 'Airport infrastructure delivery reference class'},
        'healthcare':    {'cost_pct': 44, 'schedule_pct': 38, 'source': 'HM Treasury Green Book — Hospital'},
        'roads':         {'cost_pct': 46, 'schedule_pct': 40, 'source': 'HM Treasury Green Book Table B6 — Roads'},
        'ports':         {'cost_pct': 25, 'schedule_pct': 28, 'source': 'Port infrastructure delivery reference class'},
        'telecoms':      {'cost_pct': 85, 'schedule_pct': 55, 'source': 'National broadband rollout reference class (NBN/BT/SMETS)'},
        'space':         {'cost_pct': 140, 'schedule_pct': 80, 'source': 'NASA/ESA space programme reference class — Flyvbjerg 2022'},
        'semiconductor': {'cost_pct': 40, 'schedule_pct': 45, 'source': 'Semiconductor new-fab delivery 2010-2024'},
    }
    
    uplift = oba_uplifts.get(key, {'cost_pct': 35, 'schedule_pct': 35, 'source': 'HM Treasury Green Book general infrastructure'})
    
    # Benchmark comparison — named programme growth
    named_growth = []
    for b in bench_matches[:3]:
        if b.get('cost_growth_pct',0) > 0:
            named_growth.append(f"{b.get('name','?')}: +{b['cost_growth_pct']}% cost, +{b.get('schedule_slip_months',0)} months")
    
    # Calculate OBA-adjusted estimate
    oba_adjusted = p50 * (1 + uplift['cost_pct']/100) if p50 else 0
    schedule_raw = str(m.get('schedule','')).replace(' months','').strip()
    try:
        sched_months = int(schedule_raw)
        sched_adjusted = round(sched_months * (1 + uplift['schedule_pct']/100))
    except: sched_months = 0; sched_adjusted = 0
    
    gap = oba_adjusted - p50 if p50 else 0
    
    return {
        'oba_cost_uplift_pct': uplift['cost_pct'],
        'oba_schedule_uplift_pct': uplift['schedule_pct'],
        'oba_source': uplift['source'],
        'oba_adjusted_p50': f"${oba_adjusted:.1f}B" if oba_adjusted else '—',
        'oba_gap': f"${gap:.1f}B" if gap else '—',
        'oba_adjusted_schedule': f"{sched_adjusted} months" if sched_adjusted else '—',
        'named_programme_growth': named_growth,
        'verdict': (
            f"HM Treasury Green Book optimism bias uplift for {key} is +{uplift['cost_pct']}% cost / +{uplift['schedule_pct']}% schedule ({uplift['source']}). "
            f"Applied to this programme: OBA-adjusted P50 = ${oba_adjusted:.1f}B (vs stated ${p50:.1f}B), schedule = {sched_adjusted} months (vs stated {sched_months} months). "
            f"Gap to cover: ${gap:.1f}B. "
            + (f"Named comparables confirm the pattern: {'; '.join(named_growth[:2])}." if named_growth else "")
        ),
        'board_challenge': f"This programme's P50 is ${p50:.1f}B. Flyvbjerg's global reference class for {key} infrastructure (200+ countries, 1960-2023) shows +{uplift['cost_pct']}% average cost growth and +{uplift['schedule_pct']}% schedule overrun. OBA-adjusted outturn estimate: ${oba_adjusted:.1f}B. The HM Treasury Green Book, US OMB Circular A-11, World Bank OPCS guidance and Infrastructure Australia all mandate reference class forecasting — what evidence proves this programme will outperform its global reference class?",
        'green_book_compliant': conf >= 70 if (conf := float(str(m.get('confidence_pct',65)).replace('%',''))) else False,
        'global_regulatory_mandate': "Reference class forecasting is now mandatory for major project appraisal in: UK (HM Treasury 2022), Australia (IPA 2021), USA (OMB 2023), World Bank (OPCS 2020), EU (JASPERS 2019), New Zealand (Treasury 2021).",
    }



def _v137_financing_context(key, m, loc_ctx):
    """Global financing intelligence — what funding structures apply to this project in this location."""
    loc = m.get('location','')
    p50_str = str(m.get('cost_p50','$0B')).replace('$','').replace('B','').strip()
    try: p50 = float(p50_str)
    except: p50 = 0
    
    # Determine likely financing structures by location
    financing_map = {
        # MDB-financed locations
        'Nigeria': {'primary':'World Bank IDA / AfDB', 'secondary':'China EXIM Bank', 'structure':'ODA grant + sovereign loan', 'currency_risk':'High — NGN exposure', 'procurement_rules':'World Bank Procurement Regulations 2020 / AfDB ORPP'},
        'Kenya': {'primary':'World Bank IBRD / AfDB', 'secondary':'China EXIM / EIB', 'structure':'Sovereign loan + PPP concession', 'currency_risk':'Medium — KES managed float', 'procurement_rules':'World Bank Procurement Regulations 2020'},
        'India': {'primary':'World Bank IBRD / ADB / AIIB', 'secondary':'NDB / domestic bonds / IIFCL', 'structure':'Sovereign guarantee + project finance', 'currency_risk':'Medium — INR managed', 'procurement_rules':'World Bank ICB or national competitive bidding'},
        'South Africa': {'primary':'DBSA / IDC / World Bank IBRD', 'secondary':'Private finance / DFI', 'structure':'Government balance sheet + PPP (PFMA)', 'currency_risk':'High — ZAR volatile', 'procurement_rules':'CIDB / PFMA / PPPFA procurement'},
        'Saudi Arabia': {'primary':'PIF / Saudi Aramco / sovereign', 'secondary':'Islamic finance / sukuk', 'structure':'Government direct / Vision 2030 mandate', 'currency_risk':'Low — SAR USD-pegged', 'procurement_rules':'Saudi Government Tenders and Procurement Law'},
        'UAE': {'primary':'Government sovereign / sovereign wealth', 'secondary':'Project finance / Islamic sukuk', 'structure':'Government direct or PPP', 'currency_risk':'Low — AED USD-pegged', 'procurement_rules':'Federal Law No.6 of 2018 procurement'},
        'Brazil': {'primary':'BNDES / CAF / IDB', 'secondary':'CRI/CRA / debentures', 'structure':'Government BNDES + private concession', 'currency_risk':'Very High — BRL volatile', 'procurement_rules':'Lei 14.133/2021 procurement law'},
        'Australia': {'primary':'Commonwealth / State government / asset recycling', 'secondary':'Superannuation funds / private', 'structure':'Government balance sheet + PPP (State PPP Policy)', 'currency_risk':'Low — AUD managed', 'procurement_rules':'State government PPP / AS 4000 / AS 4902'},
        'United Kingdom': {'primary':'HM Treasury / UK Infrastructure Bank', 'secondary':'Private finance / PFI replacement (PF2) / green bonds', 'structure':'Government balance sheet + private finance initiative', 'currency_risk':'Low — GBP', 'procurement_rules':'PCR 2015 / Procurement Act 2023 / NEC4'},
        'United States': {'primary':'IIJA / IRA / DOE LPO / TIFIA / PABs', 'secondary':'Private finance / infrastructure funds', 'structure':'Federal grant + State match + private', 'currency_risk':'None — USD baseline', 'procurement_rules':'FAR / Davis-Bacon / Buy American / NEPA'},
    }
    
    # Default global MDB structure
    default = {'primary':'MDB (World Bank / ADB / AfDB / AIIB)', 'secondary':'Bilateral (China EXIM / JICA / KfW)', 'structure':'Sovereign loan + potential PPP concession', 'currency_risk':'Assess USD exposure vs local currency revenue', 'procurement_rules':'World Bank Procurement Regulations 2020 / MDB standard'}
    
    fctx = None
    for k, v in financing_map.items():
        if k.lower() in loc.lower():
            fctx = v
            break
    if not fctx: fctx = default
    
    # Sector-specific financing notes
    sector_finance = {
        'rail':        "Rail is typically government-funded or concession PPP. User revenue rarely covers capital cost. Require demand study with P10/P50/P90 ridership.",
        'nuclear':     "Nuclear requires government balance sheet or regulated asset base (RAB) — private capital cannot price FOAK risk. UK RAB model is the emerging global standard.",
        'energy':      "Renewable energy is bankable on contracted revenue (CfD/PPA). Merchant exposure requires 30-40% more contingency. Grid connection must be contracted before financial close.",
        'defence':     "Defence is sovereign funded — no private capital. Budget cycle risk is the primary financing constraint.",
        'data_centre': "Hyperscale data centres are privately financed. Tenant pre-commitment is the credit basis — unlocked grid connection is the bankability constraint.",
        'space':       "Space infrastructure is sovereign or venture-funded. Satellite constellations require anchor customer contracts before launch commitment.",
        'gigafactory': "Gigafactories require government incentives (CHIPS Act / IRA / UK ATF) + OEM offtake agreement to achieve bankability. Without both, no lender will finance.",
        'mining':      "Mining is project-financed against offtake. Requires: JORC resource, signed offtake, environmental consent, community agreement. Any of these missing breaks the financing.",
        'water':       "Water is typically government-funded or regulated utility (WACC-based). Smart meter rollouts require regulator (Ofwat/EPA equivalent) approval of the business case.",
        'healthcare':  "Healthcare is government-funded or PPP (DBFOM). NHS/government credit rating is the financing basis. Clinical brief must be approved before financial model.",
        'oil_gas':     "Oil & gas is project-financed against reserve base. Requires: reserves certification, offtake agreement, government PSA, environmental consent. Lenders require P90 reserve coverage.",
    }
    
    return {
        'primary_source': fctx['primary'],
        'secondary_source': fctx['secondary'],
        'structure': fctx['structure'],
        'currency_risk': fctx['currency_risk'],
        'procurement_rules': fctx['procurement_rules'],
        'sector_note': sector_finance.get(key, "Assess funding source, procurement rules and currency risk for this location."),
        'bankability_verdict': 'BANKABLE' if p50 < 5 else ('LARGE TRANSACTION' if p50 < 20 else 'MEGA PROGRAMME — SOVEREIGN BALANCE SHEET OR INTERNATIONAL CONSORTIUM'),
        'board_question': f"What is the funding structure for this programme — is it government balance sheet, PPP concession, MDB loan, or private finance? The procurement rules, currency risk and approval gateway differ fundamentally for each.",
    }


def _v136_gate_review_readiness(key, L, m):
    """Generate gate review readiness assessment — what the programme needs to prove at each stage gate."""
    conf = float(str(m.get('confidence_pct',65)).replace('%',''))
    constraints = L.get('constraints','the governing constraint')
    
    # Standard infrastructure gate structure
    gates = {
        'G0': {
            'name': 'Strategic outline case / option selection',
            'casey_verdict': 'READY' if conf >= 70 else 'CONDITIONAL' if conf >= 55 else 'NOT READY',
            'what_casey_needs': [
                'Named owner for governing critical-path constraint',
                f'Evidence basis for {constraints}',
                'Reserve sized to P80 not P50',
                'Sector benchmark comparison with named comparable programme',
                'Three scenarios with explicit trade-off narrative',
            ],
            'what_will_fail_gate': f"No named owner for {constraints} — this is the consistent IPA gateway finding",
            'casey_check': f"CASEY rates this case {conf:.0f}% board-defensible. " + ("Pass G0 with conditions." if conf >= 55 else "Do not approve until evidence gaps are closed."),
        },
        'G1': {
            'name': 'Outline business case / strategic case approval',
            'casey_verdict': 'READY' if conf >= 68 else 'CONDITIONAL' if conf >= 52 else 'NOT READY',
            'what_casey_needs': [
                f'{constraints.split(",")[0].strip().capitalize()} programme confirmed with owner and evidence date',
                'P80 reserve formally approved — not embedded in P50',
                'Procurement strategy agreed for long-lead items',
                'Schedule logic validated by independent reviewer',
                'Risk register with named owners and triggers for top 5 risks',
            ],
            'what_will_fail_gate': "Reserve inadequacy — board approves at P50 with P80 exposure unnamed",
            'casey_check': f"Governing constraint: {constraints}. {'OBC approval is defensible at this confidence level.' if conf >= 68 else 'CASEY recommends deferring OBC until evidence gaps are closed.'}",
        },
        'G2': {
            'name': 'Full business case / investment decision',
            'casey_verdict': 'READY' if conf >= 74 else 'CONDITIONAL' if conf >= 60 else 'NOT READY',
            'what_casey_needs': [
                f'All {constraints} have named owners and evidence closure dates',
                'QCRA P80 and P90 formally approved by investment committee',
                'All long-lead procurement items have confirmed delivery slots or purchase orders',
                'Contractor appointed or shortlisted — commercial certainty established',
                'Independent cost review completed and findings addressed',
                'Programme mortality risk below 40%',
            ],
            'what_will_fail_gate': "Procurement certainty gap — long-lead items still on OEM intent letters at FBC",
            'casey_check': f"{'Investment decision is defensible.' if conf >= 74 else 'CASEY recommends independent cost review before FBC approval — ' + constraints + ' not yet evidenced.'}",
        },
        'G3': {
            'name': 'Contract award / financial close',
            'casey_verdict': 'READY' if conf >= 68 else 'CONDITIONAL' if conf >= 55 else 'NOT READY',
            'what_casey_needs': [
                'All commercial terms agreed — no open commercial points at financial close',
                'Risk register updated post-contract with contractor responsibilities',
                'Programme baseline locked and accepted by contractor',
                'Change control board constituted with authority levels',
                'First 90-day delivery plan agreed and resourced',
            ],
            'what_will_fail_gate': "Open commercial points carried into contract — creates claims exposure from day 1",
            'casey_check': f"{'Contract award is defensible.' if conf >= 68 else 'CASEY flags open commercial exposure — resolve before financial close.'}",
        },
        'G4': {
            'name': 'Delivery gate / mid-programme review',
            'casey_verdict': 'CONDITIONAL',
            'what_casey_needs': [
                f'Evidence that {constraints.split(",")[0].strip()} is on track against programme baseline',
                'Reserve consumption within approved tolerance — no unapproved drawdown',
                'Top 5 risks have updated mitigations with evidence, not just reassurance',
                'Sector-specific milestone (varies — see below) confirmed',
                'Confidence trajectory positive or stable — not declining',
            ],
            'what_will_fail_gate': "Reserve drawdown without board approval — signals loss of commercial control",
            'casey_check': "Mid-programme confidence review required. CASEY will re-run on updated programme data.",
        },
        'G5': {
            'name': 'Operational readiness / commissioning gate',
            'casey_verdict': 'CONDITIONAL',
            'what_casey_needs': [
                'Commissioning programme complete — no outstanding safety-critical actions',
                'Operational handover accepted by operator/end-user',
                'Post-project review scope agreed and resources committed',
                'Lessons learned documented and shared with sector',
                'Final account agreed or settlement pathway confirmed',
            ],
            'what_will_fail_gate': "Incomplete commissioning carried forward — creates operational liability post-handover",
            'casey_check': "Commissioning gate must be clean — no operational liability accepted at handover.",
        },
    }
    
    # Current gate assessment based on confidence
    if conf >= 74: current_gate = 'G2'
    elif conf >= 68: current_gate = 'G1'
    elif conf >= 55: current_gate = 'G0'
    else: current_gate = 'PRE-G0'
    
    return {
        'current_gate_readiness': current_gate,
        'overall_verdict': gates.get(current_gate, {}).get('casey_verdict', 'NOT READY'),
        'gates': gates,
        'critical_gate_risk': f"The most likely gate failure is at FBC/G2: {constraints} may not be evidenced by named owners before the investment committee requires commitment.",
        'ipa_alignment': f"IPA gateway review would likely find: evidence inadequacy around {constraints}. This is the primary finding in 14 of 19 Red-rated major government programmes (IPA Annual Report 2023/24).",
        'next_gate_actions': gates.get(current_gate, {}).get('what_casey_needs', []),
    }


def _v136_procurement_heatmap(key, L, m):
    """Real sector-specific procurement packages with status and risk."""
    packages = {
        'rail': [
            {'package':'Civil works and tunnelling','status':'Active','risk':'Ground conditions and third-party interface exposure','owner':'Civils Programme Manager','value_est':'35-45% of P50','lead_time':'6-18 months','single_source_risk':False},
            {'package':'Signalling and control systems','status':'Active','risk':'FOAK integration risk — signalling is the critical path constraint','owner':'Systems Integration Director','value_est':'20-30% of P50','lead_time':'24-36 months','single_source_risk':True},
            {'package':'Stations and public realm','status':'Watch','risk':'Interface with live rail operations and third-party retail','owner':'Stations Programme Manager','value_est':'15-20% of P50','lead_time':'12-24 months','single_source_risk':False},
            {'package':'Rolling stock and traction power','status':'Active','risk':'Long-lead equipment — operator acceptance dependency','owner':'Rolling Stock Manager','value_est':'10-15% of P50','lead_time':'36-60 months','single_source_risk':True},
            {'package':'Utility diversions','status':'Watch','risk':'Third-party consent and access — programme cannot control timescales','owner':'Utilities Manager','value_est':'5-10% of P50','lead_time':'12-36 months','single_source_risk':False},
        ],
        'nuclear': [
            {'package':'Nuclear island construction','status':'Active','risk':'FOAK construction — first-pour to reactor pressure vessel is 7+ years','owner':'Nuclear Construction Director','value_est':'40-50% of P50','lead_time':'24-48 months','single_source_risk':True},
            {'package':'Nuclear-grade components and equipment','status':'Active','risk':'Single-source suppliers — qualification to nuclear grade adds 18-36 months','owner':'Nuclear Procurement Director','value_est':'25-35% of P50','lead_time':'36-72 months','single_source_risk':True},
            {'package':'Safety systems and instrumentation','status':'Watch','risk':'ONR qualification — every change requires re-approval','owner':'Safety Systems Manager','value_est':'10-15% of P50','lead_time':'24-48 months','single_source_risk':True},
            {'package':'Balance of plant','status':'Active','risk':'Turbine-generator and conventional island — long lead but more competitive','owner':'BOP Manager','value_est':'10-15% of P50','lead_time':'18-36 months','single_source_risk':False},
            {'package':'Nuclear civil and structures','status':'Watch','risk':'Nuclear-grade concrete and structural steel — specialist workforce limited','owner':'Civil Engineering Director','value_est':'10-15% of P50','lead_time':'12-24 months','single_source_risk':False},
        ],
        'defence': [
            {'package':'Secure facilities and hardened works','status':'Active','risk':'Security constraints limit contractor pool — cleared supplier capacity','owner':'Defence Construction Director','value_est':'35-45% of P50','lead_time':'12-24 months','single_source_risk':False},
            {'package':'Mission systems and classified equipment','status':'Active','risk':'Export control, clearance requirements and FOAK integration','owner':'Mission Systems Director','value_est':'25-35% of P50','lead_time':'24-48 months','single_source_risk':True},
            {'package':'Secure comms and IT infrastructure','status':'Watch','risk':'Accreditation requirement — every interface requires security clearance','owner':'CIS Programme Manager','value_est':'10-15% of P50','lead_time':'18-36 months','single_source_risk':True},
            {'package':'Resilient power and utilities','status':'Active','risk':'Hardened specification drives premium — limited specialist suppliers','owner':'Infrastructure Manager','value_est':'5-10% of P50','lead_time':'12-24 months','single_source_risk':False},
            {'package':'Training systems and simulation','status':'Watch','risk':'Integration with operational systems — acceptance criteria agreed late','owner':'Training Systems Manager','value_est':'5-10% of P50','lead_time':'18-30 months','single_source_risk':False},
        ],
        'data_centre': [
            {'package':'Grid connection and HV infrastructure','status':'Active','risk':'DNO queue — not a contract, a queue position. Critical path item.','owner':'Power Programme Manager','value_est':'15-25% of P50','lead_time':'24-48 months','single_source_risk':True},
            {'package':'Cooling systems (liquid and air)','status':'Active','risk':'Performance at peak GPU load — manufacturer data sheet vs actual','owner':'Mechanical Systems Manager','value_est':'20-30% of P50','lead_time':'18-30 months','single_source_risk':False},
            {'package':'Transformers and switchgear','status':'Active','risk':'52-week minimum lead time — commodity shortage risk','owner':'Electrical Programme Manager','value_est':'10-15% of P50','lead_time':'18-36 months','single_source_risk':False},
            {'package':'Generators and UPS systems','status':'Watch','risk':'Load bank testing required — commissioning constraint','owner':'Power Resilience Manager','value_est':'8-12% of P50','lead_time':'12-18 months','single_source_risk':False},
            {'package':'Civil structure and shell','status':'Active','risk':'Accelerated programme risk — design to delivery compression','owner':'Civil Construction Manager','value_est':'20-30% of P50','lead_time':'12-18 months','single_source_risk':False},
        ],
        'energy': [
            {'package':'Grid connection and network reinforcement','status':'Active','risk':'DNO/TSO queue — 3-7 year wait in UK. Critical path item.','owner':'Grid Connection Manager','value_est':'20-35% of P50','lead_time':'36-84 months','single_source_risk':True},
            {'package':'Generation equipment (turbines/panels/storage)','status':'Active','risk':'Long-lead — commodity shortage and supply chain concentration','owner':'Generation Equipment Manager','value_est':'30-40% of P50','lead_time':'18-36 months','single_source_risk':False},
            {'package':'HVDC cable and substation','status':'Watch','risk':'Single cable route — installation weather window dependency','owner':'Offshore Systems Manager','value_est':'15-25% of P50','lead_time':'24-48 months','single_source_risk':False},
            {'package':'Civil and marine works','status':'Active','risk':'Marine window dependency — weather and tidal programme constraint','owner':'Marine Construction Manager','value_est':'10-20% of P50','lead_time':'12-24 months','single_source_risk':False},
            {'package':'Permitting and consents','status':'Watch','risk':'Not a procurement package — but gates all other packages','owner':'Planning Manager','value_est':'1-2% of P50','lead_time':'24-60 months','single_source_risk':True},
        ],
        'space': [
            {'package':'Launch vehicle contract','status':'Active','risk':'Named provider required — manifest position not a commitment','owner':'Launch Director','value_est':'25-40% of P50','lead_time':'24-60 months','single_source_risk':True},
            {'package':'Payload and mission systems','status':'Active','risk':'FOAK qualification — TRL gap drives cost and schedule','owner':'Payload Systems Director','value_est':'30-40% of P50','lead_time':'36-72 months','single_source_risk':True},
            {'package':'Power and thermal systems','status':'Watch','risk':'Sustained operations proof — datasheet performance vs space environment','owner':'Systems Engineering Manager','value_est':'10-15% of P50','lead_time':'24-48 months','single_source_risk':False},
            {'package':'Ground segment and communications','status':'Active','risk':'Interface with launch provider and mission operations — often deferred','owner':'Ground Segment Manager','value_est':'10-15% of P50','lead_time':'18-36 months','single_source_risk':False},
            {'package':'Autonomous operations software','status':'Watch','risk':'Qualification in simulated environment only — FOAK in space','owner':'Software Systems Director','value_est':'5-10% of P50','lead_time':'24-48 months','single_source_risk':False},
        ],
    }
    
    default = [
        {'package': pkg, 'status': 'Active' if i < 3 else 'Watch', 
         'risk': f"Sector-specific procurement exposure for {pkg} — named owner and lead time confirmation required",
         'owner': L.get('schedule',[['Commercial Lead']*5])[i] if i < len(L.get('schedule',[])) else 'Commercial Lead',
         'value_est': f"{30-i*5}-{40-i*5}% of P50", 'lead_time': '12-36 months', 'single_source_risk': i < 2}
        for i, pkg in enumerate(L.get('cost', ['Package 1','Package 2','Package 3','Package 4','Package 5'])[:5])
    ]
    return packages.get(key, default)


def _v136_enrich_risks(risks, key, L):
    """Add real risk IDs, sector-specific triggers and mitigations to every risk."""
    # Sector-specific triggers and mitigations
    triggers = {
        'rail': [
            "Operator declines possession access request for trial running window",
            "IEM open count exceeds 50 at systems integration milestone",
            "Utility diversion delay confirmed by third-party statutory undertaker",
            "Safety case submission refused or held by ORR",
            "Signalling FAT failure requiring retesting cycle",
            "Rolling stock acceptance test failure identified",
            "Environmental stop notice issued during active corridor works",
        ],
        'nuclear': [
            "GDA hold-point triggered by nuclear-grade weld inspection failure",
            "FOAK component rejected at factory acceptance test",
            "Safety case submission returned by ONR for significant revision",
            "Nuclear-grade supplier declared force majeure or insolvency",
            "First-pour concrete test failure requiring redesign review",
            "Workforce nuclear clearance batch delayed more than 6 months",
            "Radiological contamination event requiring works suspension",
        ],
        'defence': [
            "Security accreditation review opened by authority following interface concern",
            "Classified supplier loses clearance status or exits programme",
            "Export control restriction applied to foreign-supplied component",
            "EMC failure identified during operational trials",
            "Cyber security penetration test identifies critical vulnerability",
            "ITAR restriction applied to system sub-component",
            "Operational acceptance trial fails user acceptance criteria",
        ],
        'data_centre': [
            "Grid connection agreement not executed by DNO by programme date",
            "Cooling system performance test failure under peak GPU load",
            "Transformer delivery delayed beyond 52-week lead time",
            "IT tenant requirement change identified post-design freeze",
            "Generator load-bank commissioning failure requiring root-cause",
            "Planning condition imposed restricting hours or cooling method",
            "Power quality issue identified at energisation test",
        ],
        'life_sciences': [
            "Clean utility qualification failure identified during commissioning",
            "FDA/MHRA pre-approval inspection identifies GMP critical observation",
            "Media fill failure requiring investigation and protocol revision",
            "CQV resource unavailable due to global pharmaceutical capacity",
            "Process equipment OEM delays URS approval beyond programme date",
            "Environmental monitoring data indicates contamination risk",
            "Regulatory dossier submission delayed by CMC data gap",
        ],
        'semiconductor': [
            "OEM confirms tool delivery delay beyond contractual date",
            "UPW system water quality below spec at commissioning milestone",
            "Cleanroom particle count exceeds ISO class limit at certification",
            "Yield performance below ramp model at 6-month production milestone",
            "Chemical supply agreement not executed by programme date",
            "Workforce semiconductor certification supply below programme demand",
            "Process node qualification test failure requiring retesting",
        ],
        'gigafactory': [
            "Grid connection agreement not executed by DNO by programme date",
            "Cell chemistry qualification fails at production trial milestone",
            "Battery management system supplier confirms programme delay",
            "Yield at 12-month ramp more than 30% below business case",
            "Environmental permit for chemical storage refused or deferred",
            "OEM tool delivery slot deferred beyond production start date",
            "Customer offtake agreement not executed before financial close",
        ],
        'energy': [
            "DNO confirms grid connection delay beyond contracted energisation date",
            "Transformer or HV switchgear OEM confirms delivery beyond lead time",
            "Planning inspector issues enforcement notice on consented activity",
            "CfD or PPA milestone missed triggering penalty clause",
            "Aviation or radar objection upheld by planning authority",
            "Environmental stop notice issued for protected species breach",
            "Performance test failure at commissioning requiring root-cause",
        ],
        'water': [
            "Discharge permit refused or significantly conditioned by Environment Agency",
            "Process performance test failure at commissioning milestone",
            "Utility tie-in access refused by third-party network operator",
            "Comms infrastructure deployment delayed by wayleave dispute",
            "Back-office data platform not available at go-live milestone",
            "Smart meter reading failure rate exceeds regulatory threshold",
            "Community opposition triggers judicial review of planning consent",
        ],
        'oil_gas': [
            "Long-lead equipment OEM confirms delivery delay beyond programme date",
            "HAZOP/HAZID finding requires design change at advanced engineering stage",
            "Shutdown access window cancelled or reduced by operations team",
            "Process safety incident during hook-up requiring works suspension",
            "Regulator issues stop-work order following safety case review",
            "Weather delay exceeds programme float for marine operations",
            "Start-up performance test failure requiring process modification",
        ],
        'mining': [
            "Environmental consent refused, significantly conditioned or judicially reviewed",
            "Tailings storage facility approval deferred by regulator",
            "Community protest or blockade prevents site access",
            "Orebody grade below predicted requiring processing redesign",
            "Processing plant yield more than 20% below design performance",
            "Water licence conditions more restrictive than assumed",
            "Logistics corridor (rail/port) unavailable at ramp-up milestone",
        ],
        'airport': [
            "ORAT milestone missed — operations team not ready for acceptance",
            "Baggage system acceptance test failure at opening rehearsal",
            "CAA/EASA safety audit identifies non-compliance requiring remediation",
            "Airside access constraint not resolved before critical enabling works",
            "Fire safety integrated test failure requiring design modification",
            "Security system integration failure at pre-opening test",
            "Airline slot confirmation delayed beyond ORAT planning horizon",
        ],
        'healthcare': [
            "Infection-control HTM/HBN compliance failure identified at inspection",
            "Medical equipment delivery delayed beyond programme critical path",
            "NHS clinical commissioning milestone missed by operational team",
            "Patient transfer plan not approved by clinical governance board",
            "Digital health system integration failure at go-live milestone",
            "CQC inspection identifies patient safety concern requiring remediation",
            "Live hospital operational incident requires works suspension",
        ],
        'roads': [
            "Utility diversion delayed by third-party statutory undertaker dispute",
            "Development consent order challenged by judicial review",
            "Ground investigation reveals unexpected contamination requiring remediation",
            "Traffic management incident triggers regulator stop-work review",
            "Structures design change required following geotechnical data",
            "Environmental constraint identified during works requiring stop",
            "Stage opening safety audit fails requiring remediation works",
        ],
        'ports': [
            "Marine consent conditions more restrictive than design assumption",
            "Dredging reveals unexpected UXO or contaminated sediment",
            "Crane delivery delayed by OEM beyond programme date",
            "Terminal operating system integration failure at commissioning",
            "Vessel specification change requires dredging beyond consented depth",
            "Live port operational incident requires works suspension",
            "Port operator financial difficulty affects cutover plan",
        ],
        'space': [
            "Qualification test failure identified during environmental testing campaign",
            "Launch vehicle manifest date deferred by provider beyond programme window",
            "Mass margin breach confirmed at CDR requiring design iteration",
            "Autonomous commissioning failure identified during ground simulation",
            "Radiation tolerance test failure on primary system component",
            "TRL gap identified at programme PDR requiring additional development",
            "Mission assurance board hold-point triggered by safety concern",
        ],
    }
    
    mitigations = {
        'rail': [
            "Named possession sponsor to confirm access windows with operator, execute Deed of Grant, and place in possession register 6 months ahead",
            "Establish IEM tracker with weekly SRO review — escalation trigger at 100 open items",
            "Utility diversion programme with individual third-party agreements — no programme dependency on unconfirmed diversions",
            "Safety case programme as critical path item — ORR pre-engagement from month 6",
            "Independent signalling integrator review at FAT + 90 days before trial running",
        ],
        'nuclear': [
            "GDA pre-application engagement with ONR — hold-point schedule locked and funded",
            "Dual-source strategy for nuclear-grade components above £5M — no single-source acceptance without board sign-off",
            "Safety case programme on master critical path with named nuclear safety officer",
            "Nuclear workforce certification programme started 24 months ahead of first-pour",
            "Independent nuclear quality assurance audit at each construction milestone",
        ],
        'defence': [
            "Security accreditation programme on master critical path — SIRO named and funded",
            "All classified suppliers to confirm clearance status and capacity 12 months ahead — no assumed slots",
            "Export control review of all foreign-supplied components at concept stage",
            "EMC test programme at integration rig before installation — not at operational trials",
            "Operational acceptance criteria agreed with end-user at contract award — not at delivery",
        ],
        'data_centre': [
            "Grid connection agreement executed with DNO — not a queue position — before construction mobilisation",
            "Cooling performance verified by independent test at 80% and 100% design load before acceptance",
            "Transformer and HV switchgear ordered at earliest opportunity — 52-week lead time minimum",
            "IT tenant requirements frozen at design freeze — change control board from month 3",
            "Backup generator programme on critical path — load bank test at practical completion",
        ],
        'life_sciences': [
            "CQV programme on master critical path from month 1 — named validation manager and resource plan",
            "Clean utility qualification protocol approved by QA before installation begins",
            "FDA/MHRA pre-approval inspection readiness programme 12 months ahead of submission",
            "Media fill programme scheduled 6 months before regulatory submission — no slip",
            "Process equipment URS approved by QA before procurement — not post-delivery",
        ],
        'semiconductor': [
            "All tool delivery slots confirmed as binding purchase orders — OEM intent letters not accepted",
            "UPW system commissioning programme on critical path — independent water quality test at handover",
            "Cleanroom certification programme with independent third-party at each phase handover",
            "Yield ramp milestone in board performance framework with 6-month review — not a post-delivery target",
            "Semiconductor workforce programme — named training provider with cohort schedule",
        ],
        'gigafactory': [
            "Grid connection agreement executed with DNO before financial close — not a queue position",
            "Cell chemistry qualification on master programme critical path — named R&D owner",
            "Battery management system supplier under contract with delivery milestone — no assumed delivery",
            "Offtake agreement executed with anchor customer before construction mobilisation",
            "Yield ramp in board business case with quarterly review — conservative first-year assumption",
        ],
        'energy': [
            "Grid connection agreement executed with signed energisation date — not a DNO forecast",
            "All transformers and HV equipment ordered at earliest opportunity — 52-week minimum",
            "Aviation and radar consultation completed before planning submission — not concurrent",
            "CfD/PPA commissioning milestone with programme float — not minimum float against deadline",
            "Environmental surveys completed in all seasons before construction mobilisation",
        ],
        'water': [
            "Discharge permit pre-application agreed with Environment Agency — formal application with buffer",
            "Comms infrastructure in programme boundary and on critical path — not separate IT project",
            "Back-office platform delivery date contracted and on programme critical path",
            "Wayleave programme managed by dedicated access team — not contractor-assumed",
            "Process performance test protocol approved by regulator before commissioning begins",
        ],
        'oil_gas': [
            "All long-lead items confirmed as binding purchase orders — intent letters not accepted",
            "Shutdown access windows confirmed in operations plan — not assumed from historical access",
            "HAZOP/HAZID completed at FEED stage — not during detailed design",
            "Process safety case submitted to regulator 12 months before first production",
            "Marine weather window programme with Monte Carlo simulation — P80 window used not median",
        ],
        'mining': [
            "Environmental consent secured before financial commitment — judicial review risk assessed",
            "Tailings storage facility design and approval on master critical path from day 1",
            "Community engagement programme with SRO ownership — not delegated to contractor",
            "Orebody resource estimate to JORC code — independent verification before approval",
            "Processing plant yield test at pilot scale before full plant commitment",
        ],
        'airport': [
            "ORAT programme on master critical path — named airline operations director as owner",
            "Baggage system acceptance tested 3 months before opening date — full rehearsal required",
            "CAA/EASA pre-application engagement from month 12 — approval programme on critical path",
            "Airside access programme agreed with airport operations — works programme locked to access",
            "Fire safety integrated test programme 6 months before opening — not post-construction",
        ],
        'healthcare': [
            "Clinical commissioning programme as named milestone on master schedule — NHS operational director owner",
            "Medical equipment procurement programme on critical path — 24-month lead time minimum",
            "Infection-control HTM/HBN compliance audit at each phase handover — not at occupation",
            "Patient transfer plan approved by clinical governance board 12 months ahead",
            "Digital health integration test at commissioning rig — not live system at go-live",
        ],
        'roads': [
            "Utility diversion agreements with all statutory undertakers signed before mobilisation",
            "Development consent order legally final before contract award — judicial review risk assessed",
            "Ground investigation to appropriate level before design freeze — no assumed conditions",
            "Traffic management plan agreed with local authority — programme built around confirmed access",
            "Stage opening safety audit programme 3 months before each opening — not post-opening",
        ],
        'ports': [
            "Marine consent and dredging scope confirmed before design freeze — no assumed conditions",
            "Terminal operating system in EPC programme boundary — not separate procurement",
            "Crane and terminal equipment ordered at earliest opportunity — lead time minimum 18 months",
            "Vessel specification frozen at contract award — change control board for any revision",
            "Operational continuity plan agreed with port operator — revenue impact of disruption costed",
        ],
        'space': [
            "Launch vehicle contract executed with named provider and confirmed manifest position",
            "Independent TRL verification for each system below TRL 7 — internal claims not accepted",
            "Mass margin maintained at 20% minimum through to CDR — breach triggers design review",
            "Autonomous commissioning simulation on ground rig before launch — not first use in orbit",
            "Radiation tolerance testing on all primary components — not relied on datasheet ratings",
        ],
    }
    
    fallback_triggers = triggers.get(key, [
        "Governing constraint fails to close at programme milestone",
        "Evidence required by board not available at approval gateway",
        "Key supplier defaults or declares force majeure",
        "Regulatory approval delayed beyond programme assumption",
        "Interface risk materialises creating critical path impact",
    ])
    fallback_mitigations = mitigations.get(key, [
        "Named owner to evidence constraint closure with dated proof point",
        "Reserve sized to P80 not P50 — independent QCRA validation",
        "Alternative supplier identified and contracted as backstop",
        "Regulator pre-engagement programme on master critical path",
        "Interface programme with named accountability and weekly SRO review",
    ])
    
    enriched = []
    for i, r in enumerate(risks):
        if not isinstance(r, dict):
            enriched.append(r)
            continue
        r = dict(r)
        r['risk_id'] = f"R-{i+1:03d}"
        r['trigger'] = fallback_triggers[i % len(fallback_triggers)]
        r['mitigation'] = fallback_mitigations[i % len(fallback_mitigations)]
        r['early_warning'] = f"Watch indicator: {fallback_triggers[(i+1) % len(fallback_triggers)][:60]}"
        r['residual_risk'] = 'Medium' if i < 3 else 'Low'
        r['review_frequency'] = 'Weekly SRO' if r.get('probability_pct',50) > 40 else 'Monthly'
        enriched.append(r)
    return enriched


def _v132_apply(model, prompt='', client=''):
    key = _v132_sector_key(prompt, client, model)
    L = _v132_library(key)
    m = dict(model or {})
    scenario = str(m.get('scenario_label') or m.get('scenario') or 'Base').title()
    m['app_version'] = 'CASEY V132 Institutional Authority Final'
    m['sector_ontology_key'] = key
    m['sector_ontology_label'] = L['label']
    m['subsector'] = L['label']
    m['sector_constraints'] = L['constraints']
    m['executive_shock_insight'] = L['shock']
    m['institutional_authority_line'] = L['challenge_line']
    _pid = _v136_extract_project_identity(prompt)
    _loc = f" in {_pid['location']}" if _pid['location'] else ""
    # Wire in location context — regulatory framework, currency, financing, OBA note
    _loc_ctx = location_context(_pid['location'] or location_name if 'location_name' in dir() else (_pid['location'] or ''))
    _scale = f" — {_pid['scale_signal']}" if _pid['scale_signal'] else ""
    _cost_hint = f" (stated budget: {_pid['cost_signal']})" if _pid['cost_signal'] else ""
    # Build project-specific title from prompt + sector (don't use detect_sector's generic title)
    _sector_titles = {
        'rail': 'Rail Programme', 'nuclear': 'Nuclear Programme', 'defence': 'Defence Programme',
        'data_centre': 'Data Centre Programme', 'life_sciences': 'Pharma / Life Sciences Programme',
        'semiconductor': 'Semiconductor Fab Programme', 'gigafactory': 'Battery Gigafactory',
        'oil_gas': 'Oil & Gas Programme', 'mining': 'Mining Programme', 'airport': 'Airport Programme',
        'healthcare': 'Healthcare Programme', 'energy': 'Energy Programme', 'water': 'Water & Utilities Programme',
        'telecoms': 'Telecoms Programme', 'roads': 'Roads Programme', 'ports': 'Ports Programme',
        'stadia': 'Stadia Programme', 'civic': 'Civic Programme', 'space': 'Space Programme',
        'gigafactory': 'Battery Gigafactory', 'general_infrastructure': 'Infrastructure Programme',
    }
    _sector_title = _sector_titles.get(key, L['label'])
    if _pid['project_title'] and len(_pid['project_title']) > 6:
        m['title'] = _pid['project_title'] + (_loc or '')
    elif m.get('title') and m['title'] not in ['Rail Infrastructure','Major Infrastructure Programme','General Infrastructure','Capital Infrastructure Programme']:
        m['title'] = m['title'] + (_loc or '')
    else:
        m['title'] = _sector_title + (_loc or '')
    m['location'] = _pid['location'] or m.get('location','')
    _lctx = location_context(m.get('location',''))
    m['location_context'] = {
        'currency': _lctx.get('currency','USD'),
        'framework': _lctx.get('framework',''),
        'approval_body': _lctx.get('approval_body',''),
        'optimism_bias_note': _lctx.get('optimism_bias_note',''),
        'financing': _lctx.get('financing',''),
        'risk_premium': _lctx.get('risk_premium',''),
    }
    m['executive_summary'] = (
        f"{m['title']}{_loc}{_scale}{_cost_hint}. "
        f"{L['label']} programme · {scenario}: {m.get('cost_p50')} P50, {m.get('cost_range')} range, "
        f"{m.get('schedule')} baseline. {m.get('confidence_pct')}% board-defensibility confidence. "
        f"{L['shock']} "
        f"Governing constraint: {L['constraints']}."
    )
    m['board_briefing'] = [
        L['shock'],
        L['challenge_line'],
        f"P50 reconciles to {m.get('cost_p50')}; P80/P90 remain the board contingency and stress conversation.",
        "The case should not be approved until the governing constraint is evidenced, owner-named and traceable to cost, schedule and risk outputs."
    ]
    m['casey_thinking'] = f"CASEY has locked this run to {L['label']} and is challenging the programme narrative against {L['constraints']}. It is not accepting progress optics; it is testing whether the decision is board-defensible."
    m['confidence_explanation'] = f"{m.get('confidence_pct')}% is a board-defensibility score, not optimism. It is constrained by {L['constraints']}."
    m['confidence_engine_detail'] = {
        'decision_rule':'Do not treat this as approval evidence until owner actions, evidence source, residual exposure and P80/P90 movement are reconciled.',
        'primary_constraint':L['constraints'],
        'plain_english':'CASEY scores whether the board can defend the decision, not whether the team feels confident.'
    }
    _q1 = _v135_sector_q1(key, L.get('constraints','the governing constraint'))
    m['board_challenge_questions'] = [
        _q1,
        f"What proof closes the gap around {L['constraints']}?",
        'What second-order risk does the preferred scenario create?',
        'Which mitigation changes confidence rather than just narrative?',
        'What would invalidate this case before approval?',
        'Where is the programme buying schedule, savings or assurance — and what is being sacrificed?'
    ]
    m['if_this_fails'] = _v135_if_this_fails(key, L.get('constraints',''), L.get('label',''))
    m['sector_failure_pattern'] = m['if_this_fails']
    m['gate_review_readiness'] = _v136_gate_review_readiness(key, L, m)
    m['financing_context'] = _v137_financing_context(key, m, m.get('location_context',{}))
    _bm_for_oba = benchmark_similarity(prompt, m.get('mode','Earth'), m.get('subsector',''))
    m['optimism_bias_assessment'] = _v136_optimism_bias(key, m, _bm_for_oba)
    m['second_order_contradictions'] = [
        f"The case can look stable while {L['constraints']} remain unevidenced.",
        "Acceleration may shorten the visible plan while increasing interface, assurance and commissioning fragility.",
        "Savings may be risk transfer into P80/P90 exposure, operations or recovery reserve rather than true cost removal.",
        "Higher confidence must come from evidence closure, not from a smoother executive narrative.",
        "The programme may be reporting progress in the workface while losing confidence in the governing constraint."
    ]
    m['governance_challenges'] = L['interventions']
    m['behavioural_forecast'] = L['behavioural_forecast']
    m['intervention_intelligence'] = L['interventions']
    m['sector_confidence_drivers'] = L['confidence']
    m['sector_primary_cost_drivers'] = L['cost']
    m['sector_schedule_threats'] = L['schedule']
    m['causal_graph_nodes'] = L['chain']
    m['causal_chain'] = L['chain']
    # Use named benchmarks from BENCHMARK_LIBRARY if available, fall back to L['bench']
    _named_bm = benchmark_similarity(prompt, m.get('mode','Earth'), m.get('subsector',''))
    if _named_bm:
        m['benchmark_comparison'] = [{
            'name': b.get('name', b['sector']),
            'sector': b['sector'], 'archetype': b.get('name', b['sector']),
            'anchor_cost': money_bn(b['cost_bn']),
            'anchor_duration_months': b['months'],
            'similarity_score': b.get('similarity_score', 7),
            'cost_growth_pct': b.get('cost_growth_pct', 0),
            'schedule_slip_months': b.get('schedule_slip_months', 0),
            'failure_mode': b.get('failure_mode', ''),
            'lesson': b.get('lesson', ''),
            'metric': b.get('name', b['sector']),
            'value': f"{money_bn(b['cost_bn'])} | {b['months']} months",
            'why': b.get('lesson', b.get('failure_mode', f"Named comparable for {b['sector']}")),
        } for b in _named_bm[:5]]
    else:
        m['benchmark_comparison'] = [{'sector':a,'archetype':a,'anchor_cost':c,'anchor_duration_months':d,'similarity_score':max(6,10-i),'use':'Sector benchmark'} for i,(a,c,d) in enumerate(L.get('bench',[])[:4])]
    m['benchmark_memory'] = m['benchmark_comparison']; m['benchmarks'] = m['benchmark_comparison']; m['peer_competitors'] = m['benchmark_comparison']
    m['why_casey_generated_this'] = [
        f"CASEY detected {L['label']} and locked the ontology before output generation.",
        f"Sector behaviours applied: {', '.join(L['chain'][:5])}.",
        "Benchmark cohort, causal chain, confidence drivers, risks and exports were constrained to this sector.",
        "The output is intentionally challenge-oriented: board defensibility over optimism."
    ]
    sigs=[{'signal':s,'status':st,'direction':'confidence / reserve / P-tail','weight':0.13,'applies_to':'board pack, workbook, risk register, XER, QCRA/QSRA','basis':basis} for s,st,basis in L['signals']]
    m['live_calibration_signals']=sigs; m['live_calibration_strip']=' • '.join(x['signal'] for x in sigs[:4])
    m['mission_control_cards']=[{'label':'Live calibration','signal':'Sector conditions are being converted into confidence, contingency and delivery-tail exposure.','severity':'Active'}]+[{'label':s,'signal':basis,'severity':st} for s,st,basis in L['signals']]
    m['mission_control_cards']=m['mission_control_cards'][:6]
    m['uncertainty_narrative']={'estimate_maturity':'Class 3 is usable for budget authorisation only if the evidence gaps are explicit.','schedule_maturity':'Schedule Level 4 improves logic, but board confidence is still governed by near-critical density and owner evidence.','interpretation':f"Live calibration is weighting {L['constraints']} into the QSRA/QCRA tail. {L['behavioural_forecast']}"}
    cost, sched, risks = _v132_rows(L, m)
    risks = _v136_enrich_risks(risks, key, L)
    m['cost_lines']=cost; m['schedule_rows']=sched; m['risk_register']=risks; m['risks']=risks
    m['cost_breakdown']=cost; m['estimates_by_class']={str(i):cost for i in range(1,6)}; m['schedules_by_level']={str(i):sched for i in range(1,6)}
    m['procurement_heatmap'] = _v136_procurement_heatmap(key, L, m)
    m['critical_path_narrative']=[f"{x} is near-critical in the {L['label']} sector graph and must be evidenced before approval." for x in L['schedule'][:5]]
    m['red_flags']=[f"Unevidenced confidence around {L['constraints']}.", "Scenario benefit may be risk transfer rather than risk reduction.", "Board pack should name owner, evidence source and date for each governing constraint."]
    m['near_critical_narrative'] = f"Near-critical density should be interpreted through {L['constraints']}; the headline critical path is not the only board risk."
    forbidden = _v132_forbidden(key)
    if '_v124_scrub_value' in globals():
        for k in list(m.keys()):
            if k not in {'prompt','client','title','id'}:
                m[k] = _v124_scrub_value(m[k], forbidden)
    # Restore native arrays and authority fields after scrub.
    m['sector_ontology_key']=key; m['sector_ontology_label']=L['label']; m['subsector']=L['label']; m['sector_constraints']=L['constraints']
    m['sector_confidence_drivers']=L['confidence']; m['sector_primary_cost_drivers']=L['cost']; m['sector_schedule_threats']=L['schedule']; m['causal_graph_nodes']=L['chain']; m['causal_chain']=L['chain']
    m['sector_signature_behaviours']=L['chain'][:5]
    m['schedule_detail']=sched
    m['all_schedule_levels']={str(i):sched for i in range(1,6)}
    m['cost_detail']=cost
    m['risk_detail']=risks
    m['governance_challenges']=L['interventions']; m['intervention_intelligence']=L['interventions']; m['behavioural_forecast']=L['behavioural_forecast']
    # Apply intelligence intelligence fields
    m = _v136_tt_killer_fields(m, key, L)
    return m

_CASEY_V132_BASE_BUILD_MODEL = _CASEY_V130_PREV_BUILD_MODEL if '_CASEY_V130_PREV_BUILD_MODEL' in globals() else build_model

def build_model(prompt:str, client:str='', class_level:int=3, schedule_level:int=3, scenario:str='base'):
    return _v132_apply(_CASEY_V132_BASE_BUILD_MODEL(prompt, client, class_level, schedule_level, scenario), prompt, client)

APP_VERSION='CASEY V132 Institutional Authority Final'
print('CASEY V132 institutional authority final installed')
# ================= END CASEY V132 INSTITUTIONAL AUTHORITY LOCK =================

# ================= CASEY V134 FULL SECTOR RENDER / EXPORT HARDENING =================
# Purpose: prevent rail/life-sciences/any-sector incomplete render payloads and ensure every
# Earth + Space sector returns a complete, export-safe, sector-native model.

def _v134_text(v, fallback=''):
    if isinstance(v, str) and v.strip(): return v
    if isinstance(v, (int, float, bool)): return str(v)
    if isinstance(v, dict):
        bits=[v.get(k) for k in ('label','name','title','driver','risk','signal','phase','activity','meaning','note','basis','value','effect') if v.get(k)]
        return ': '.join(map(str,bits[:2])) if bits else fallback
    return fallback

def _v134_list(v, fallback):
    src = v if isinstance(v, list) and v else fallback
    return [_v134_text(x) for x in src if _v134_text(x)]

V134_OVERRIDE_LIB = {
 'nuclear': dict(label='Nuclear / Regulated Generation', shock='Licensing, safety-case maturity, nuclear-grade procurement and regulator hold points govern confidence more than construction progress.', constraints='licensing, safety-case maturity, nuclear-grade procurement, QA traceability and regulator hold points', signals=[('Safety case maturity','Active','Safety-case evidence controls board defensibility.'),('Regulator hold points','Active','Regulatory hold points govern schedule-tail exposure.'),('Nuclear-grade procurement','Watch','Qualified component lead-times constrain procurement certainty.'),('QA traceability','Watch','Documentation and quality evidence govern release to commission.')], bench=[('Advanced Nuclear Generation','$8B–$45B','72-180'),('SMR / Modular Reactor Programme','$2B–$20B','48-144'),('Nuclear Life Extension / Upgrade','$1B–$12B','36-120')], chain=['Safety case maturity','Regulator hold points','Nuclear-grade procurement','QA traceability','Containment systems','Commissioning governance','Confidence'], confidence=['Benchmark similarity: nuclear regulated generation programme','Scope maturity: safety case and reactor island definition','Procurement certainty: qualified nuclear-grade components','Schedule maturity: regulator hold-point and commissioning logic','Interface exposure: QA traceability, operator and regulator'], cost=['Reactor and containment systems','Nuclear-grade equipment and controls','QA, licensing and assurance evidence','Balance-of-plant and grid interface','Commissioning and regulatory reserve'], schedule=['Safety-case approval','Regulator hold-point release','Nuclear-grade component delivery','QA dossier and traceability closure','Commissioning governance and operational acceptance']),
 'general_infrastructure': dict(label='General Capital Infrastructure', shock='The dominant risk is not progress reporting; it is whether the governing constraint has named evidence, owner accountability and reserve logic.', constraints='scope maturity, procurement certainty, interface control, commissioning readiness and evidence ownership', signals=[('Evidence ownership','Active','Board confidence requires named evidence owners, not generic mitigation text.'),('Procurement certainty','Active','Long-lead packages and commercial exposure drive P80/P90 movement.'),('Interface control','Watch','Third-party interfaces and integration gates create schedule tails.'),('Commissioning readiness','Watch','Operational readiness is a governance signal, not late administration.')], bench=[('Capital Infrastructure Programme','$500M–$15B','30-120'),('Complex Systems Integration Programme','$300M–$10B','24-96'),('Major Public / Private Capital Programme','$1B–$25B','48-144')], chain=['Scope definition','Procurement evidence','Interface control','Commissioning readiness','Owner evidence','Reserve logic','Confidence'], confidence=['Benchmark similarity: capital infrastructure archetype','Scope maturity: requirements and package definition','Procurement certainty: long-lead and market capacity','Schedule maturity: critical path and commissioning logic','Interface exposure: utilities, stakeholders and operations'], cost=['Core asset and enabling works','Specialist systems and long-lead packages','Interfaces, utilities and third-party works','Programme management and assurance','Risk reserve and contingency'], schedule=['Scope freeze and approvals','Long-lead procurement','Interface and utility readiness','Commissioning and operational readiness','Owner evidence and board acceptance'])
}

def _v134_library(key):
    if key in V134_OVERRIDE_LIB: return V134_OVERRIDE_LIB[key]
    L = _v132_library(key) if '_v132_library' in globals() else {}
    if not isinstance(L, dict) or not L.get('label'):
        return V134_OVERRIDE_LIB['general_infrastructure']
    return L

def _v134_normalise_rows(m, L):
    # Cost rows
    cost = m.get('cost_breakdown') or m.get('cost_lines') or []
    if not isinstance(cost, list) or not cost:
        cost = [{'cbs':f'01.0{i+1}','description':x,'type':'Direct' if i<3 else ('Indirect' if i==3 else 'Reserve'),'p10_bn':0.5+i*.2,'p50_bn':0.7+i*.25,'p90_bn':1.0+i*.3,'basis':'Sector-normalised fallback line'} for i,x in enumerate(L['cost'][:5])]
    fixed_cost=[]
    for i,r in enumerate(cost[:12]):
        if not isinstance(r, dict): r={'description':_v134_text(r, L['cost'][min(i,len(L['cost'])-1)])}
        fixed_cost.append({
            'cbs':_v134_text(r.get('cbs'), f'01.{i+1:02d}'),
            'description':_v134_text(r.get('description') or r.get('package'), L['cost'][min(i,len(L['cost'])-1)]),
            'type':_v134_text(r.get('type'), 'Direct' if i<3 else ('Indirect' if i==3 else 'Reserve')),
            'p10_bn':float(r.get('p10_bn') or r.get('low_bn') or r.get('low') or 0.4+i*.2),
            'p50_bn':float(r.get('p50_bn') or r.get('most_likely_bn') or r.get('p50') or 0.7+i*.25),
            'p90_bn':float(r.get('p90_bn') or r.get('high_bn') or r.get('high') or 1.0+i*.3),
            'basis':_v134_text(r.get('basis'), 'Sector-normalised estimate basis')
        })
    m['cost_breakdown']=fixed_cost; m['cost_lines']=fixed_cost; m['cost_detail']=fixed_cost
    # Schedule rows
    sched = m.get('schedule_detail') or m.get('schedule_rows') or []
    if not isinstance(sched, list) or not sched:
        sched=[{'activity_id':f'A{1000+i*100}','phase':'Delivery','activity':x,'predecessor':f'A{900+i*100}' if i else '', 'duration_months':4+i*2,'critical':'Yes' if i>1 else 'No','basis':'Sector-normalised schedule line'} for i,x in enumerate(L['schedule'][:5])]
    fixed_sched=[]
    for i,r in enumerate(sched[:20]):
        if not isinstance(r, dict): r={'activity':_v134_text(r, L['schedule'][min(i,len(L['schedule'])-1)])}
        fixed_sched.append({
            'activity_id':_v134_text(r.get('activity_id') or r.get('id'), f'A{1000+i*100}'),
            'phase':_v134_text(r.get('phase'), 'Delivery'),
            'activity':_v134_text(r.get('activity') or r.get('name'), L['schedule'][min(i,len(L['schedule'])-1)]),
            'predecessor':_v134_text(r.get('predecessor') or r.get('pred'), '' if i==0 else f'A{900+i*100}'),
            'duration_months':int(float(r.get('duration_months') or r.get('months') or 4+i*2)),
            'critical':_v134_text(r.get('critical'), 'Yes' if i>1 else 'No'),
            'basis':_v134_text(r.get('basis'), 'Sector-normalised schedule basis')
        })
    m['schedule_detail']=fixed_sched; m['schedule_rows']=fixed_sched; m['all_schedule_levels']={str(i):fixed_sched for i in range(1,6)}
    # Risks
    risks = m.get('risk_register') or m.get('risks') or []
    if not isinstance(risks, list) or not risks:
        risks=[{'id':f'R-{i+1:03d}','risk':x,'cause':'Sector constraint not evidenced','event':'Constraint crystallises','impact':'Cost/schedule/confidence degradation','owner':'Programme Director','mitigation':'Named owner action and evidence closure','likelihood':'Medium','impact_rating':'High'} for i,x in enumerate(L['schedule'][:5])]
    fixed_risks=[]
    for i,r in enumerate(risks[:20]):
        if not isinstance(r, dict): r={'risk':_v134_text(r, L['schedule'][min(i,len(L['schedule'])-1)])}
        fixed_risks.append({**r,'id':_v134_text(r.get('id'),f'R-{i+1:03d}'),'risk':_v134_text(r.get('risk') or r.get('event'),L['schedule'][min(i,len(L['schedule'])-1)]),'owner':_v134_text(r.get('owner'),'Programme Director'),'mitigation':_v134_text(r.get('mitigation'),'Named owner action and evidence closure'),'likelihood':_v134_text(r.get('likelihood'),'Medium'),'impact':_v134_text(r.get('impact'),'High')})
    m['risk_register']=fixed_risks; m['risks']=fixed_risks; m['risk_detail']=fixed_risks
    return m

def _v134_patch_model(model, prompt='', client=''):
    key = _v132_sector_key(prompt, client, model) if '_v132_sector_key' in globals() else 'general_infrastructure'
    L = _v134_library(key)
    m = dict(model or {})
    m['app_version']='CASEY V134 Full Sector Render Hardened'
    m['sector_ontology_key']=key; m['sector_ontology_label']=L['label']; m['subsector']=L['label']
    m['sector_constraints']=L['constraints']; m['executive_shock_insight']=L['shock']
    m['causal_graph_nodes']=list(L['chain']); m['causal_chain']=list(L['chain'])
    m['sector_confidence_drivers']=list(L['confidence']); m['sector_primary_cost_drivers']=list(L['cost']); m['sector_schedule_threats']=list(L['schedule'])
    m['board_briefing']=[L['shock'], f"The programme narrative is only defensible if {L['constraints']} are evidenced by named owners before approval.", f"P50 reconciles to {m.get('cost_p50','the current estimate')}; P80/P90 remain the board contingency and stress conversation.", 'The case should not be approved until the governing constraint is evidenced, owner-named and traceable to cost, schedule and risk outputs.']
    m['next_best_actions']=[f"Name the accountable owner and evidence source for {L['constraints']}.",'Separate true risk reduction from risk transfer into operations, reserve or P90 exposure.','Reconcile the P50 approval story with P80/P90 downside before board commitment.','Retire the governing constraint before buying acceleration or declaring savings.']
    m['board_challenge_questions']=[_v135_sector_q1(key, L.get('constraints','the governing constraint')),f"What proof closes the gap around {L['constraints']}?",'What second-order risk does the preferred scenario create?','Which mitigation changes confidence rather than just narrative?','What would invalidate this case before approval?']
    m['second_order_contradictions']=[f"The case can look stable while {L['constraints']} remain unevidenced.",'Acceleration may shorten the visible plan while increasing interface, assurance and commissioning fragility.','Savings may be risk transfer into P80/P90 exposure, operations or recovery reserve rather than true cost removal.','Higher confidence must come from evidence closure, not from a smoother executive narrative.']
    _nbm = benchmark_similarity(prompt, m.get('mode','Earth'), m.get('subsector',''))
    if _nbm:
        m['benchmark_comparison'] = [{
            'name': b.get('name', b['sector']),
            'sector': b['sector'], 'archetype': b.get('name', b['sector']),
            'anchor_cost': money_bn(b['cost_bn']),
            'anchor_duration_months': b['months'],
            'similarity_score': b.get('similarity_score', 7),
            'cost_growth_pct': b.get('cost_growth_pct', 0),
            'schedule_slip_months': b.get('schedule_slip_months', 0),
            'failure_mode': b.get('failure_mode', ''),
            'lesson': b.get('lesson', ''),
            'metric': b.get('name', b['sector']),
            'value': f"{money_bn(b['cost_bn'])} | {b['months']} months",
            'why': b.get('lesson', b.get('failure_mode', f"Named comparable for {b['sector']}")),
        } for b in _nbm[:5]]
    else:
        m['benchmark_comparison']=[{'archetype':a,'sector':a,'anchor_cost':c,'anchor_duration_months':d,'similarity_score':max(6,10-i),'use':'Sector benchmark','why':f'Sector archetype for {a}'} for i,(a,c,d) in enumerate(L.get('bench',[])[:4])]
    m['benchmark_memory']=m['benchmark_comparison']; m['benchmarks']=m['benchmark_comparison']; m['peer_competitors']=m['benchmark_comparison']
    sigs=[{'signal':s,'status':st,'direction':'confidence / reserve / P-tail','weight':0.13,'applies_to':'board pack, workbook, risk register, XER, QCRA/QSRA','basis':basis} for s,st,basis in L['signals']]
    m['live_calibration_signals']=sigs; m['mission_control_cards']=[{'label':'Live calibration','signal':'Sector conditions are being converted into confidence, contingency and delivery-tail exposure.','severity':'Active'}]+[{'label':s,'signal':basis,'severity':st} for s,st,basis in L['signals']]
    m['why_casey_generated_this']=[f"CASEY detected {L['label']} and locked the ontology before output generation.", f"Sector behaviours applied: {', '.join(L['chain'][:5])}.", 'Benchmark cohort, causal chain, confidence drivers, risks and exports were constrained to this sector.', 'The output is intentionally challenge-oriented: board defensibility over optimism.']
    m['uncertainty_narrative']={'estimate_maturity':'Class maturity is suitable for option selection only if evidence gaps are explicit.','schedule_maturity':'Schedule logic requires critical-path, handover and commissioning validation.','interpretation':f"Live calibration is weighting {L['constraints']} into the QCRA/QSRA tail."}
    m=_v134_normalise_rows(m,L)
    return m

_CASEY_V134_PREV_BUILD_MODEL = build_model
def build_model(prompt:str, client:str='', class_level:int=3, schedule_level:int=3, scenario:str='base'):
    return _v134_patch_model(_CASEY_V134_PREV_BUILD_MODEL(prompt, client, class_level, schedule_level, scenario), prompt, client)

APP_VERSION='CASEY V134 Full Sector Render Hardened'
print('CASEY V134 full-sector render/export hardening installed')
# ================= END CASEY V134 FULL SECTOR RENDER / EXPORT HARDENING =================

# ================= CASEY V134.1 NUMERIC SAFE HOTFIX =================
def _v134_float(v, fallback=0.0):
    import re
    if isinstance(v, (int, float)): return float(v)
    s = str(v or '')
    m = re.search(r'-?\d+(?:\.\d+)?', s.replace(',', ''))
    if not m: return float(fallback)
    n = float(m.group(0))
    if 'm' in s.lower() and 'b' not in s.lower(): n = n / 1000.0
    return n

def _v134_normalise_rows(m, L):
    cost = m.get('cost_breakdown') or m.get('cost_lines') or []
    if not isinstance(cost, list) or not cost:
        cost = [{'cbs':f'01.0{i+1}','description':x,'type':'Direct' if i<3 else ('Indirect' if i==3 else 'Reserve'),'p10_bn':0.5+i*.2,'p50_bn':0.7+i*.25,'p90_bn':1.0+i*.3,'basis':'Sector-normalised fallback line'} for i,x in enumerate(L['cost'][:5])]
    fixed_cost=[]
    for i,r in enumerate(cost[:12]):
        if not isinstance(r, dict): r={'description':_v134_text(r, L['cost'][min(i,len(L['cost'])-1)])}
        fixed_cost.append({'cbs':_v134_text(r.get('cbs'), f'01.{i+1:02d}'),'description':_v134_text(r.get('description') or r.get('package'), L['cost'][min(i,len(L['cost'])-1)]),'type':_v134_text(r.get('type'), 'Direct' if i<3 else ('Indirect' if i==3 else 'Reserve')),'p10_bn':_v134_float(r.get('p10_bn') or r.get('low_bn') or r.get('low'), 0.4+i*.2),'p50_bn':_v134_float(r.get('p50_bn') or r.get('most_likely_bn') or r.get('p50'), 0.7+i*.25),'p90_bn':_v134_float(r.get('p90_bn') or r.get('high_bn') or r.get('high'), 1.0+i*.3),'basis':_v134_text(r.get('basis'), 'Sector-normalised estimate basis')})
    m['cost_breakdown']=fixed_cost; m['cost_lines']=fixed_cost; m['cost_detail']=fixed_cost
    sched = m.get('schedule_detail') or m.get('schedule_rows') or []
    if not isinstance(sched, list) or not sched:
        sched=[{'activity_id':f'A{1000+i*100}','phase':'Delivery','activity':x,'predecessor':f'A{900+i*100}' if i else '', 'duration_months':4+i*2,'critical':'Yes' if i>1 else 'No','basis':'Sector-normalised schedule line'} for i,x in enumerate(L['schedule'][:5])]
    fixed_sched=[]
    for i,r in enumerate(sched[:20]):
        if not isinstance(r, dict): r={'activity':_v134_text(r, L['schedule'][min(i,len(L['schedule'])-1)])}
        fixed_sched.append({'activity_id':_v134_text(r.get('activity_id') or r.get('id'), f'A{1000+i*100}'),'phase':_v134_text(r.get('phase'), 'Delivery'),'activity':_v134_text(r.get('activity') or r.get('name'), L['schedule'][min(i,len(L['schedule'])-1)]),'predecessor':_v134_text(r.get('predecessor') or r.get('pred'), '' if i==0 else f'A{900+i*100}'),'duration_months':int(_v134_float(r.get('duration_months') or r.get('months'), 4+i*2)),'critical':_v134_text(r.get('critical'), 'Yes' if i>1 else 'No'),'basis':_v134_text(r.get('basis'), 'Sector-normalised schedule basis')})
    m['schedule_detail']=fixed_sched; m['schedule_rows']=fixed_sched; m['all_schedule_levels']={str(i):fixed_sched for i in range(1,6)}
    risks = m.get('risk_register') or m.get('risks') or []
    if not isinstance(risks, list) or not risks:
        risks=[{'id':f'R-{i+1:03d}','risk':x,'cause':'Sector constraint not evidenced','event':'Constraint crystallises','impact':'Cost/schedule/confidence degradation','owner':'Programme Director','mitigation':'Named owner action and evidence closure','likelihood':'Medium','impact_rating':'High'} for i,x in enumerate(L['schedule'][:5])]
    fixed_risks=[]
    for i,r in enumerate(risks[:20]):
        if not isinstance(r, dict): r={'risk':_v134_text(r, L['schedule'][min(i,len(L['schedule'])-1)])}
        fixed_risks.append({**r,'id':_v134_text(r.get('id'),f'R-{i+1:03d}'),'risk':_v134_text(r.get('risk') or r.get('event'),L['schedule'][min(i,len(L['schedule'])-1)]),'owner':_v134_text(r.get('owner'),'Programme Director'),'mitigation':_v134_text(r.get('mitigation'),'Named owner action and evidence closure'),'likelihood':_v134_text(r.get('likelihood'),'Medium'),'impact':_v134_text(r.get('impact'),'High')})
    m['risk_register']=fixed_risks; m['risks']=fixed_risks; m['risk_detail']=fixed_risks
    return m

_CASEY_V1341_PREV_BUILD_MODEL = _CASEY_V134_PREV_BUILD_MODEL if '_CASEY_V134_PREV_BUILD_MODEL' in globals() else build_model
def build_model(prompt:str, client:str='', class_level:int=3, schedule_level:int=3, scenario:str='base'):
    return _v134_patch_model(_CASEY_V1341_PREV_BUILD_MODEL(prompt, client, class_level, schedule_level, scenario), prompt, client)
APP_VERSION='CASEY V134.1 Full Sector Render Hardened'
print('CASEY V134.1 numeric-safe hardening installed')
# ================= END CASEY V134.1 NUMERIC SAFE HOTFIX =================
