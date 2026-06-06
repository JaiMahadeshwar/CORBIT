#!/usr/bin/env python3
"""
CASEY V150 — Full QA: 10,000 Earth + 5,000 Space + all showcase + export smoke tests.
Reports pass/fail with full detail on every check.
"""
import sys, os, re, json, time, random, io, zipfile, traceback

sys.path.insert(0, '/tmp'); import fastapi_stub
ROOT = os.path.dirname(os.path.abspath(__file__))
BACKEND = os.path.join(ROOT, 'backend')
os.chdir(BACKEND); sys.path.insert(0, BACKEND)
import main

# ── Helpers ──────────────────────────────────────────────────────────
def bn(x):
    s = str(x or '').replace('$','').replace(',','').strip().upper()
    try:
        if s.endswith('T'): return float(s[:-1]) * 1000
        if s.endswith('B'): return float(s[:-1])
        if s.endswith('M'): return float(s[:-1]) / 1000
        m = re.search(r'-?\d+(?:\.\d+)?', s)
        return float(m.group(0)) if m else 0.0
    except: return 0.0

def mo(x):
    m = re.search(r'-?\d+(?:\.\d+)?', str(x or ''))
    return int(round(float(m.group(0)))) if m else 0

def check_model(m, label):
    fails = []
    p50 = bn(m.get('cost_p50'))
    direct = bn(m.get('direct_cost'))
    indirect = bn(m.get('indirect_cost'))
    reserve = bn(m.get('risk_reserve'))
    bucket = direct + indirect + reserve
    if p50 > 0 and abs(bucket - p50) > 0.02:
        fails.append(f'bucket_sum: p50={p50:.3f} != d+i+r={bucket:.3f}')

    mc = m.get('monte_carlo') or {}
    qcra = mc.get('qcra') or {}; qsra = mc.get('qsra') or {}
    qcra_p50 = float(qcra.get('p50') or 0)
    if p50 > 0 and abs(qcra_p50 - p50) > 0.15:
        fails.append(f'qcra_p50_drift: model={p50:.3f} qcra={qcra_p50:.3f}')

    qsra_p50 = int(qsra.get('p50') or 0)
    sched_mo = mo(m.get('schedule'))
    if sched_mo > 0 and abs(qsra_p50 - sched_mo) > 2:
        fails.append(f'qsra_p50_drift: sched={sched_mo} qsra_p50={qsra_p50}')

    qcra_p80 = float(qcra.get('p80') or 0)
    if qcra_p80 <= p50:
        fails.append(f'qcra_p80_below_p50: p50={p50:.3f} p80={qcra_p80:.3f}')

    qsra_p80 = int(qsra.get('p80') or 0)
    if qsra_p80 < sched_mo:
        fails.append(f'qsra_p80_below_p50: sched={sched_mo} qsra_p80={qsra_p80}')

    risks = m.get('risks') or m.get('risk_register') or []
    if len(risks) < 10:
        fails.append(f'risk_count_low: {len(risks)}')

    sched = m.get('schedule_detail') or m.get('schedule_rows') or []
    if len(sched) < 12:
        fails.append(f'sched_count_low: {len(sched)}')

    if not m.get('scenario_signature'):
        fails.append('missing: scenario_signature')
    if not m.get('subsector'):
        fails.append('missing: subsector')
    if not m.get('board_challenge_questions'):
        fails.append('missing: board_challenge_questions')
    if not m.get('executive_summary'):
        fails.append('missing: executive_summary')

    # Scenario matrix has selected scenario
    sc_val = str(m.get('scenario') or 'base').lower()
    matrix = m.get('scenario_matrix') or []
    matrix_labels = [str(r.get('scenario','')).lower().replace(' ','_') for r in matrix]
    if matrix and sc_val not in matrix_labels:
        fails.append(f'matrix_missing_scenario: {sc_val} not in {matrix_labels}')

    # Validate via _v150
    v = main._v150_validate_model(m)
    if not v.get('ok'):
        for f2 in (v.get('failures') or []):
            fails.append(f'v150: {f2}')

    return fails

# ── SECTION 1: Earth 10,000 ───────────────────────────────────────────
EARTH = [
    'California High-Speed Rail corridor with tunnelling signalling possessions land acquisition operator acceptance',
    'Microsoft AI hyperscale data centre campus with grid power liquid cooling GPU supply accelerated commissioning',
    'AWS sovereign cloud region expansion with fibre backbone grid redundancy permits geopolitical resilience',
    'Eli Lilly GMP manufacturing expansion with sterile fill finish clean utilities validation FDA handover',
    'Novo Nordisk biologics production with cold chain aseptic manufacturing regulatory evidence',
    'UK SMR nuclear rollout with licensing nuclear island procurement safety case grid connection',
    'TSMC Arizona semiconductor fab with cleanroom tool install ultrapure water workforce ramp',
    'NEOM transit infrastructure with autonomous mobility logistics scaling systems integration desert delivery',
    'AUKUS naval shipbuilding with dockyard scaling workforce nuclear certification controlled procurement',
    'LNG export terminal with cryogenic systems marine works commissioning long lead procurement',
    'Offshore wind hub with grid interconnector marine vessels weather windows turbine supply chain',
    'Major airport terminal expansion with baggage systems live operations security ORAT readiness',
    'Hospital clinical campus with MEP utilities clean rooms operating theatres phased handover',
    'Defence base radar command infrastructure with secure power hardened facilities vetting',
    'Hydrogen export corridor with electrolysis storage compression marine loading',
    'High speed rail with brownfield stations systems migration possessions network operations',
    'Port automation with container cranes AGV systems TOS integration berth deepening',
    'Water treatment desalination with reverse osmosis membrane systems discharge consent utilities',
    'Satellite ground station network with secure comms encrypted uplink antenna farms power',
    'Urban metro underground extension with TBM tunnelling ventilation station fitout systems',
]
SPACE = [
    'Lunar habitat infrastructure with life support thermal survivability autonomous commissioning launch windows',
    'SpaceX Starship industrialization with launch cadence orbital refueling thermal protection range approvals',
    'Amazon Kuiper satellite constellation with production ramp launch manifest spectrum coordination ground segment',
    'AST SpaceMobile orbital telecom with satellite deployment telecom interoperability spectrum approval',
    'Orbital data centre with thermal rejection radiation hardening autonomous servicing orbital cooling',
    'Mars cargo logistics network with ISRU long-duration autonomy communication latency radiation exposure',
    'Cislunar propellant depot with cryogenic storage docking systems launch dependency autonomous operations',
    'Autonomous orbital servicing with rendezvous proximity operations robotic capture refuelling debris avoidance',
    'Lunar resource extraction with autonomous mining regolith processing ISRU surface power launch logistics',
    'Space domain awareness ground station with secure comms encrypted uplink antenna farms launch windows',
]
SCENARIOS = ['base', 'faster', 'cheaper', 'lower_risk', 'premium']

print("=" * 70)
print(f"CASEY V150 FULL QA — {main.APP_VERSION}")
print("=" * 70)

# ── EARTH 10,000 ─────────────────────────────────────────────────────
print(f"\n[1/4] Earth models: 10,000 runs across {len(EARTH)} prompts × {len(SCENARIOS)} scenarios")
earth_fails = {}; earth_count = 0; t0 = time.time()
for i in range(10000):
    prompt = EARTH[i % len(EARTH)]
    sc = SCENARIOS[i % len(SCENARIOS)]
    cl = (i % 4) + 2  # class 2-5
    sl = (i % 4) + 2  # level 2-5
    try:
        m = main.build_model(f'{prompt}', 'QA', cl, sl, sc)
        fails = check_model(m, f'earth_{i}')
        for f in fails:
            earth_fails.setdefault(f.split(':')[0], []).append({'i': i, 'sc': sc, 'detail': f})
    except Exception as e:
        earth_fails.setdefault('exception', []).append({'i': i, 'sc': sc, 'detail': str(e)})
    earth_count += 1
    if earth_count % 2000 == 0:
        print(f"  {earth_count:,}/10,000 — fails so far: {sum(len(v) for v in earth_fails.values())}")

earth_total_fails = sum(len(v) for v in earth_fails.values())
print(f"  Earth complete: {earth_count:,} models | {earth_total_fails} failures | {time.time()-t0:.1f}s")
for k, v in earth_fails.items():
    print(f"    {k}: {len(v)} fails — sample: {v[0]['detail'][:80]}")

# ── SPACE 5,000 ───────────────────────────────────────────────────────
print(f"\n[2/4] Space models: 5,000 runs across {len(SPACE)} prompts × {len(SCENARIOS)} scenarios")
space_fails = {}; space_count = 0; t0 = time.time()
for i in range(5000):
    prompt = SPACE[i % len(SPACE)]
    sc = SCENARIOS[i % len(SCENARIOS)]
    cl = (i % 4) + 2
    sl = (i % 4) + 2
    try:
        m = main.build_model(f'Space mission: {prompt}', 'QA', cl, sl, sc)
        fails = check_model(m, f'space_{i}')
        for f in fails:
            space_fails.setdefault(f.split(':')[0], []).append({'i': i, 'sc': sc, 'detail': f})
    except Exception as e:
        space_fails.setdefault('exception', []).append({'i': i, 'sc': sc, 'detail': str(e)})
    space_count += 1
    if space_count % 1000 == 0:
        print(f"  {space_count:,}/5,000 — fails so far: {sum(len(v) for v in space_fails.values())}")

space_total_fails = sum(len(v) for v in space_fails.values())
print(f"  Space complete: {space_count:,} models | {space_total_fails} failures | {time.time()-t0:.1f}s")
for k, v in space_fails.items():
    print(f"    {k}: {len(v)} fails — sample: {v[0]['detail'][:80]}")

# ── SHOWCASE ALL SCENARIOS ────────────────────────────────────────────
print(f"\n[3/4] Showcase library: all projects × all 5 scenarios")
try:
    showcase_projects = json.load(open(os.path.join(ROOT, 'CASEY_SHOWCASE_LIBRARY', 'showcase_projects.json')))
    total_showcase_projects = sum(len(v) for v in showcase_projects.values())
except:
    showcase_projects = {'Rail': ['California HSR', 'HS2', 'Crossrail'], 'Space': ['Lunar Habitat', 'Kuiper'], 'Data': ['AI Hyperscale']}
    total_showcase_projects = 6

showcase_fails = []; showcase_count = 0; t0 = time.time()
for cat, names in showcase_projects.items():
    for name in names:
        for sc in SCENARIOS:
            prompt = f'{name} strategic delivery for {cat}: procurement, schedule, risk, confidence and board assurance'
            try:
                m = main.build_model(prompt, 'Showcase QA', 3, 4, sc)
                fails = check_model(m, f'{name}/{sc}')
                if fails:
                    showcase_fails.append({'project': name, 'scenario': sc, 'failures': fails})
            except Exception as e:
                showcase_fails.append({'project': name, 'scenario': sc, 'error': str(e)})
            showcase_count += 1

print(f"  Showcase: {showcase_count} outputs ({total_showcase_projects} projects × 5 scenarios) | {len(showcase_fails)} failures | {time.time()-t0:.1f}s")
for f in showcase_fails[:10]:
    print(f"    {f.get('project')}/{f.get('scenario')}: {f.get('failures') or f.get('error')}")

# ── EXPORT SMOKE TEST ─────────────────────────────────────────────────
print(f"\n[4/4] Export smoke test: 200 models, all 8 export functions each")
EXPORT_FUNCS = [
    ('workbook', main.workbook_bytes, 5000, b'PK'),
    ('risk_register', main.risk_register_workbook_bytes, 3000, b'PK'),
    ('xer', main.xer_bytes, 100, b'ERMHDR'),
    ('word', main.word_bytes, 3000, b'PK'),
    ('pdf', main.pdf_bytes, 1000, b'%PDF'),
    ('pptx', main.pptx_bytes, 5000, b'PK'),
    ('risk_csv', main.risk_csv_bytes, 200, b'ID,'),
]
export_fails = []; export_count = 0; t0 = time.time()
EXPORT_PROMPTS = EARTH[:10] + SPACE[:5]
for i in range(200):
    prompt = EXPORT_PROMPTS[i % len(EXPORT_PROMPTS)]
    sc = SCENARIOS[i % len(SCENARIOS)]
    domain = 'Space mission' if i >= 150 else 'Earth project'
    try:
        m = main.build_model(f'{domain}: {prompt}', 'Export QA', 3, 4, sc)
        for name, func, min_bytes, magic in EXPORT_FUNCS:
            try:
                b = func(m)
                if len(b) < min_bytes:
                    export_fails.append({'i': i, 'export': name, 'error': f'too small: {len(b)} < {min_bytes}'})
                elif not b.startswith(magic) and not (name == 'risk_csv' and b[:3] == b'ID,'):
                    export_fails.append({'i': i, 'export': name, 'error': f'bad magic: {b[:8]}'})
                export_count += 1
            except Exception as e:
                export_fails.append({'i': i, 'export': name, 'error': str(e)[:120]})
    except Exception as e:
        export_fails.append({'i': i, 'export': 'build_model', 'error': str(e)[:120]})

print(f"  Export: {export_count} exports checked | {len(export_fails)} failures | {time.time()-t0:.1f}s")
for f in export_fails[:15]:
    print(f"    [{f['export']}] i={f['i']}: {f['error']}")

# ── SUMMARY ───────────────────────────────────────────────────────────
print("\n" + "=" * 70)
total_fails = earth_total_fails + space_total_fails + len(showcase_fails) + len(export_fails)
total_runs = earth_count + space_count + showcase_count + export_count
print(f"TOTAL: {total_runs:,} checks | {total_fails} failures")
print(f"  Earth:    {earth_count:,} models  | {earth_total_fails} fails")
print(f"  Space:    {space_count:,} models  | {space_total_fails} fails")
print(f"  Showcase: {showcase_count} outputs | {len(showcase_fails)} fails")
print(f"  Exports:  {export_count} checks  | {len(export_fails)} fails")
verdict = 'PASS' if total_fails == 0 else f'FAIL ({total_fails} issues)'
print(f"\nVERDICT: {verdict}")
print("=" * 70)

# Save
result = {
    'version': main.APP_VERSION,
    'total_checks': total_runs,
    'total_failures': total_fails,
    'earth': {'count': earth_count, 'failures': earth_total_fails, 'by_type': {k: len(v) for k,v in earth_fails.items()}},
    'space': {'count': space_count, 'failures': space_total_fails, 'by_type': {k: len(v) for k,v in space_fails.items()}},
    'showcase': {'count': showcase_count, 'failures': len(showcase_fails), 'samples': showcase_fails[:10]},
    'exports': {'count': export_count, 'failures': len(export_fails), 'samples': export_fails[:15]},
    'verdict': verdict,
}
with open(os.path.join(ROOT, 'CASEY_V150_FULL_QA_RESULTS.json'), 'w') as fh:
    json.dump(result, fh, indent=2, default=str)
print(f"\nFull results saved to CASEY_V150_FULL_QA_RESULTS.json")
