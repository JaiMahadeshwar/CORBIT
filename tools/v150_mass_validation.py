#!/usr/bin/env python3
"""CASEY V150 mass validation harness.
Runs canonical state checks across Earth/Space scenario libraries and optional export smoke tests.
Usage:
  python tools/v150_mass_validation.py --earth 10000 --space 5000 --exports 100
"""
from __future__ import annotations
import argparse, importlib.util, io, json, os, sys, time, zipfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
BACKEND = os.path.join(ROOT, 'backend')
sys.path.insert(0, BACKEND)

def load_main():
    spec = importlib.util.spec_from_file_location('casey_main', os.path.join(BACKEND, 'main.py'))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # type: ignore[union-attr]
    return mod

EARTH = [
    'California High-Speed Rail corridor with tunneling, signalling, station redevelopment, possessions and land acquisition risk',
    'Microsoft AI hyperscale data centre campus with grid power, liquid cooling, GPU supply and accelerated commissioning',
    'AWS sovereign cloud region expansion with fibre backbone, grid redundancy, permits and geopolitical resilience',
    'Eli Lilly GMP manufacturing expansion with sterile fill finish, clean utilities, validation and FDA handover',
    'Novo Nordisk biologics production expansion with cold chain logistics, aseptic manufacturing and regulatory evidence',
    'UK SMR nuclear rollout with licensing, nuclear island procurement, safety case and grid connection',
    'TSMC Arizona semiconductor fab with cleanroom, tool install, ultrapure water and workforce ramp',
    'NEOM transit infrastructure with autonomous mobility, logistics scaling, systems integration and desert delivery',
    'AUKUS naval shipbuilding industrial base with dockyard scaling, workforce, nuclear certification and controlled procurement',
    'LNG export terminal with cryogenic systems, marine works, commissioning and long lead procurement',
    'Offshore wind hub with grid interconnector, marine vessels, weather windows and turbine supply chain exposure',
    'Major airport terminal expansion with baggage systems, live operations, security systems and ORAT readiness',
]
SPACE = [
    'Lunar habitat infrastructure with life support, surface power, thermal survivability, autonomous commissioning and launch windows',
    'SpaceX Starship industrialization with launch cadence, orbital refueling, thermal protection and range approvals',
    'Amazon Kuiper satellite constellation with production ramp, launch manifest, spectrum coordination and ground segment integration',
    'AST SpaceMobile orbital telecom network with satellite deployment, telecom interoperability and regulatory spectrum approval',
    'Orbital data centre with thermal rejection, radiation hardening, autonomous servicing, orbital cooling and data relay',
    'Mars cargo logistics network with ISRU, long-duration autonomy, communication latency, radiation exposure and cargo landers',
    'Cislunar propellant depot with cryogenic storage, docking systems, launch dependency and autonomous operations',
    'Autonomous orbital servicing platform with rendezvous operations, robotic capture, refuelling and debris avoidance',
    'Lunar resource extraction with autonomous mining, regolith processing, ISRU maturity, surface power and launch logistics',
]
SCENARIOS = ['base', 'faster', 'cheaper', 'lower_risk', 'premium']

def validate(mod, total_earth: int, total_space: int):
    failures = []
    start = time.time()
    total = total_earth + total_space
    for i in range(total):
        is_space = i >= total_earth
        library = SPACE if is_space else EARTH
        brief = library[i % len(library)]
        scenario = SCENARIOS[i % len(SCENARIOS)]
        try:
            model = mod.build_model(brief, 'CASEY V150 QA', 3, 3, scenario)
            result = mod._v150_validate_model(model)
            if not result.get('ok'):
                failures.append({'index': i, 'domain': 'Space' if is_space else 'Earth', 'scenario': scenario, 'failures': result.get('failures')})
            for key in ['scenario_signature', 'governance_state', 'scenario_state', 'casey_runtime']:
                if not model.get(key):
                    failures.append({'index': i, 'domain': 'Space' if is_space else 'Earth', 'scenario': scenario, 'missing': key})
        except Exception as exc:
            failures.append({'index': i, 'domain': 'Space' if is_space else 'Earth', 'scenario': scenario, 'error': repr(exc)})
        if failures and len(failures) >= 25:
            break
    return {'checked': total if len(failures) < 25 else i + 1, 'earth_target': total_earth, 'space_target': total_space, 'failure_count': len(failures), 'failures': failures[:25], 'elapsed_sec': round(time.time() - start, 2)}

def export_smoke(mod, count: int):
    failures = []
    start = time.time()
    samples = (EARTH + SPACE)
    for i in range(count):
        brief = samples[i % len(samples)]
        scenario = SCENARIOS[i % len(SCENARIOS)]
        try:
            model = mod.build_model(brief, 'CASEY V150 Export QA', 3, 3, scenario)
            bio = io.BytesIO()
            with zipfile.ZipFile(bio, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('00_CASEY_V150_EXPORT_MANIFEST.json', mod._v150_manifest(model))
                z.writestr('00_CASEY_V150_SCENARIO_STATE.json', json.dumps(mod._v148_jsonable(model.get('scenario_state') or {}), indent=2))
                z.writestr('01_CASEY_Cost_Model_Planet_Class.xlsx', mod.workbook_bytes(model))
                z.writestr('02_CASEY_Risk_Register_Pro.xlsx', mod.risk_register_workbook_bytes(model))
                z.writestr('03_CASEY_P6_Schedule.xer', mod.xer_bytes(model))
                z.writestr('04_CASEY_Executive_Board_Report.docx', mod.word_bytes(model))
                z.writestr('05_CASEY_Board_Intelligence_Pack.pdf', mod.pdf_bytes(model))
                z.writestr('06_CASEY_Board_Deck_Elite.pptx', mod.pptx_bytes(model))
                z.writestr('07_CASEY_Full_Model_Audit.json', json.dumps(mod._v148_jsonable(model), indent=2))
                z.writestr('08_CASEY_Risk_Register_Raw.csv', mod.risk_csv_bytes(model))
            with zipfile.ZipFile(io.BytesIO(bio.getvalue())) as z:
                required = ['00_CASEY_V150_EXPORT_MANIFEST.json','00_CASEY_V150_SCENARIO_STATE.json','01_CASEY_Cost_Model_Planet_Class.xlsx','02_CASEY_Risk_Register_Pro.xlsx','03_CASEY_P6_Schedule.xer','04_CASEY_Executive_Board_Report.docx','05_CASEY_Board_Intelligence_Pack.pdf','06_CASEY_Board_Deck_Elite.pptx','07_CASEY_Full_Model_Audit.json','08_CASEY_Risk_Register_Raw.csv']
                missing = [x for x in required if x not in z.namelist()]
                if missing:
                    failures.append({'index': i, 'scenario': scenario, 'missing': missing})
                manifest = json.loads(z.read('00_CASEY_V150_EXPORT_MANIFEST.json'))
                if not manifest.get('scenario_signature') or not manifest.get('validation', {}).get('ok'):
                    failures.append({'index': i, 'scenario': scenario, 'manifest_validation': manifest.get('validation')})
        except Exception as exc:
            failures.append({'index': i, 'scenario': scenario, 'error': repr(exc)})
        if failures and len(failures) >= 10:
            break
    return {'checked': count if len(failures) < 10 else i + 1, 'failure_count': len(failures), 'failures': failures[:10], 'elapsed_sec': round(time.time() - start, 2)}

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--earth', type=int, default=10000)
    ap.add_argument('--space', type=int, default=5000)
    ap.add_argument('--exports', type=int, default=100)
    ap.add_argument('--out', default='CASEY_V150_MASS_VALIDATION_REPORT.json')
    args = ap.parse_args()
    mod = load_main()
    report = {'version': getattr(mod, 'APP_VERSION', 'unknown'), 'generated_at': time.strftime('%Y-%m-%dT%H:%M:%SZ', time.gmtime()), 'mass_validation': validate(mod, args.earth, args.space), 'export_smoke': export_smoke(mod, args.exports)}
    with open(args.out, 'w', encoding='utf-8') as f:
        json.dump(report, f, indent=2)
    print(json.dumps(report, indent=2))
    if report['mass_validation']['failure_count'] or report['export_smoke']['failure_count']:
        raise SystemExit(1)

if __name__ == '__main__':
    main()
