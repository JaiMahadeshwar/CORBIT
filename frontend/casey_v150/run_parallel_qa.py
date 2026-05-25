#!/usr/bin/env python3
"""Parallel QA — 10,000 Earth + 5,000 Space + showcase + export smoke"""
import sys, os, re, json, time, io, zipfile, multiprocessing as mp

sys.path.insert(0, '/tmp'); import fastapi_stub  # noqa
ROOT = os.path.dirname(os.path.abspath(__file__))
BACKEND = os.path.join(ROOT, 'backend')

EARTH = [
    'California High-Speed Rail with tunnelling signalling possessions land acquisition operator acceptance',
    'Microsoft AI hyperscale data centre with grid power liquid cooling GPU supply accelerated commissioning',
    'Eli Lilly GMP manufacturing expansion with sterile fill finish clean utilities validation FDA handover',
    'UK SMR nuclear rollout with licensing nuclear island procurement safety case grid connection',
    'TSMC Arizona semiconductor fab with cleanroom tool install ultrapure water workforce ramp',
    'Major airport terminal expansion with baggage systems live airside operations security ORAT readiness',
    'LNG export terminal with cryogenic systems marine works commissioning long lead procurement',
    'Offshore wind hub with grid interconnector marine vessels weather windows turbine supply',
    'Hospital clinical campus MEP utilities clean rooms operating theatres phased handover',
    'Defence base radar command with secure power hardened facilities vetting classified systems',
    'Hydrogen export corridor electrolysis storage compression marine loading',
    'Urban metro underground extension TBM tunnelling ventilation station fitout systems',
    'Port automation container cranes AGV systems TOS integration berth deepening marine',
    'Water treatment desalination reverse osmosis membrane systems discharge consent',
    'AUKUS naval shipbuilding dockyard scaling nuclear certification controlled procurement',
    'High speed rail brownfield stations systems migration possessions network operations',
    'Novo Nordisk biologics production cold chain aseptic manufacturing regulatory evidence',
    'NEOM transit infrastructure autonomous mobility logistics systems integration desert delivery',
    'AWS sovereign cloud region expansion fibre backbone grid redundancy permits geopolitical',
    'Semiconductor memory fab with EUV lithography advanced packaging cleanroom commissioning',
]
SPACE = [
    'Lunar habitat with life support thermal survivability autonomous commissioning launch windows',
    'SpaceX Starship industrialization launch cadence orbital refueling thermal protection range',
    'Amazon Kuiper constellation production ramp launch manifest spectrum coordination ground segment',
    'Orbital data centre thermal rejection radiation hardening autonomous servicing orbital cooling',
    'Mars cargo logistics ISRU long-duration autonomy communication latency radiation landers',
    'Cislunar propellant depot cryogenic storage docking autonomous operations launch dependency',
    'Autonomous orbital servicing rendezvous proximity robotic capture refuelling debris avoidance',
    'Lunar resource extraction autonomous mining regolith ISRU surface power launch logistics',
    'Space domain awareness ground station secure comms encrypted uplink antenna launch windows',
    'AST SpaceMobile orbital telecom satellite deployment interoperability spectrum regulatory',
]
SCENARIOS = ['base', 'faster', 'cheaper', 'lower_risk', 'premium']

def bn(x):
    s = str(x or '').replace('$','').replace(',','').strip().upper()
    try:
        if s.endswith('T'): return float(s[:-1])*1000
        if s.endswith('B'): return float(s[:-1])
        if s.endswith('M'): return float(s[:-1])/1000
        m = re.search(r'-?\d+(?:\.\d+)?', s)
        return float(m.group(0)) if m else 0.0
    except: return 0.0

def mo(x):
    m = re.search(r'-?\d+(?:\.\d+)?', str(x or ''))
    return int(round(float(m.group(0)))) if m else 0

def _init():
    global main
    sys.path.insert(0, '/tmp'); import fastapi_stub  # noqa
    sys.path.insert(0, BACKEND); os.chdir(BACKEND)
    import main as _m; main = _m

def run_batch(args):
    kind, indices = args
    library = SPACE if kind == 'space' else EARTH
    fails = []
    for i in indices:
        prompt = library[i % len(library)]
        sc = SCENARIOS[i % len(SCENARIOS)]
        cl = (i % 4) + 2; sl = (i % 4) + 2
        prefix = 'Space mission' if kind == 'space' else 'Earth project'
        try:
            m = main.build_model(f'{prefix}: {prompt}', 'QA', cl, sl, sc)
            p50 = bn(m.get('cost_p50'))
            bucket = bn(m.get('direct_cost')) + bn(m.get('indirect_cost')) + bn(m.get('risk_reserve'))
            mc = m.get('monte_carlo') or {}
            qcra = mc.get('qcra') or {}; qsra = mc.get('qsra') or {}
            sched = mo(m.get('schedule'))
            risks = m.get('risks') or m.get('risk_register') or []
            acts = m.get('schedule_detail') or m.get('schedule_rows') or []

            if p50 > 0 and abs(bucket - p50) > 0.025:
                fails.append({'check':'bucket','kind':kind,'i':i,'sc':sc,'p50':p50,'sum':round(bucket,3)})
            qcra_p50 = float(qcra.get('p50') or 0)
            if p50 > 0 and abs(qcra_p50 - p50) > 0.20:
                fails.append({'check':'qcra_drift','kind':kind,'i':i,'sc':sc,'model_p50':round(p50,3),'qcra_p50':qcra_p50})
            qsra_p50 = int(qsra.get('p50') or 0)
            if sched > 0 and abs(qsra_p50 - sched) > 3:
                fails.append({'check':'qsra_drift','kind':kind,'i':i,'sc':sc,'sched':sched,'qsra_p50':qsra_p50})
            if float(qcra.get('p80') or 0) <= p50:
                fails.append({'check':'qcra_p80_inversion','kind':kind,'i':i,'sc':sc})
            if len(risks) < 10:
                fails.append({'check':'risk_count','kind':kind,'i':i,'sc':sc,'count':len(risks)})
            if len(acts) < 12:
                fails.append({'check':'sched_count','kind':kind,'i':i,'sc':sc,'count':len(acts)})
            if not m.get('scenario_signature'):
                fails.append({'check':'no_sig','kind':kind,'i':i,'sc':sc})
            v = main._v150_validate_model(m)
            if not v.get('ok'):
                fails.append({'check':'v150_validate','kind':kind,'i':i,'sc':sc,'details':v.get('failures')})
        except Exception as e:
            fails.append({'check':'exception','kind':kind,'i':i,'sc':sc,'error':str(e)[:120]})
    return fails

def run_exports(n=200):
    prompts = EARTH[:10] + SPACE[:5]
    fails = []; count = 0
    export_funcs = [
        ('workbook', main.workbook_bytes, 5000, b'PK'),
        ('risk_register', main.risk_register_workbook_bytes, 3000, b'PK'),
        ('xer', main.xer_bytes, 100, None),
        ('word', main.word_bytes, 3000, b'PK'),
        ('pdf', main.pdf_bytes, 1000, b'%PDF'),
        ('pptx', main.pptx_bytes, 5000, b'PK'),
        ('risk_csv', main.risk_csv_bytes, 100, None),
    ]
    for i in range(n):
        prompt = prompts[i % len(prompts)]
        sc = SCENARIOS[i % len(SCENARIOS)]
        domain = 'Space' if i >= int(n * 0.7) else 'Earth'
        try:
            m = main.build_model(f'{domain}: {prompt}', 'Export QA', 3, 4, sc)
            for name, func, min_b, magic in export_funcs:
                try:
                    b = func(m)
                    if len(b) < min_b:
                        fails.append({'export':name,'i':i,'error':f'size {len(b)} < {min_b}'})
                    elif magic and not b.startswith(magic):
                        fails.append({'export':name,'i':i,'error':f'bad magic {b[:8]!r}'})
                    count += 1
                except Exception as e:
                    fails.append({'export':name,'i':i,'error':str(e)[:120]})
        except Exception as e:
            fails.append({'export':'build','i':i,'error':str(e)[:120]})
    return count, fails

if __name__ == '__main__':
    N_EARTH = 10000; N_SPACE = 5000; WORKERS = min(8, mp.cpu_count())
    print(f"CASEY V150 FULL QA — {N_EARTH:,} Earth + {N_SPACE:,} Space + showcase + exports")
    print(f"Workers: {WORKERS}")

    # Split into batches
    BATCH = 250
    earth_batches = [('earth', list(range(s, min(s+BATCH, N_EARTH)))) for s in range(0, N_EARTH, BATCH)]
    space_batches = [('space', list(range(s, min(s+BATCH, N_SPACE)))) for s in range(0, N_SPACE, BATCH)]
    all_batches = earth_batches + space_batches

    t0 = time.time()
    all_fails = []
    completed = 0
    with mp.Pool(WORKERS, initializer=_init) as pool:
        for batch_fails in pool.imap_unordered(run_batch, all_batches, chunksize=1):
            all_fails.extend(batch_fails)
            completed += BATCH
            if completed % 5000 == 0:
                print(f"  {min(completed, N_EARTH+N_SPACE):,}/{N_EARTH+N_SPACE:,} — fails: {len(all_fails)}")

    elapsed = time.time() - t0
    earth_fails = [f for f in all_fails if f.get('kind')=='earth']
    space_fails = [f for f in all_fails if f.get('kind')=='space']
    by_check = {}
    for f in all_fails:
        by_check.setdefault(f['check'], []).append(f)

    print(f"\n  Mass validation: {N_EARTH+N_SPACE:,} models in {elapsed:.1f}s ({(N_EARTH+N_SPACE)/elapsed:.0f}/s)")
    print(f"  Earth fails: {len(earth_fails)} | Space fails: {len(space_fails)}")
    for k, v in sorted(by_check.items(), key=lambda x: -len(x[1])):
        sample = v[0]
        print(f"    {k}: {len(v)} — sample: {str(sample)[:100]}")

    # Showcase
    print(f"\n  Showcase QA...")
    _init()
    try:
        showcase_lib = json.load(open(os.path.join(ROOT, 'CASEY_SHOWCASE_LIBRARY', 'showcase_projects.json')))
    except:
        showcase_lib = {'Rail':['California HSR','HS2','Crossrail','Gateway','Sydney Metro','Brightline West','Rail Baltica','Ontario Line'],
                        'Space':['Starship','Lunar Habitat','Mars Cargo','Amazon Kuiper','AST SpaceMobile','Orbital Data Centres'],
                        'AI_Data_Centres':['Microsoft AI Supercluster','AWS Global Region','Meta AI Compute','Google TPU','Oracle Sovereign AI','NVIDIA Ecosystem'],
                        'Energy':['SMR Nuclear','LNG Terminals','Offshore Wind','Hydrogen Corridors'],
                        'Pharma':['Eli Lilly Expansion','Novo Nordisk Expansion','Moderna Biosecurity'],
                        'Defence':['AUKUS','Missile Defence','Autonomous Drones']}

    sc_fails = []; sc_count = 0
    for cat, names in showcase_lib.items():
        for name in names:
            for sc in SCENARIOS:
                try:
                    m = main.build_model(f'{name} strategic delivery for {cat} procurement schedule risk board assurance', 'Showcase', 3, 4, sc)
                    p50 = bn(m.get('cost_p50')); bucket = bn(m.get('direct_cost'))+bn(m.get('indirect_cost'))+bn(m.get('risk_reserve'))
                    risks = m.get('risks') or m.get('risk_register') or []
                    v = main._v150_validate_model(m)
                    row_fails = []
                    if p50 > 0 and abs(bucket-p50) > 0.025: row_fails.append(f'bucket p50={p50:.2f} sum={bucket:.2f}')
                    if len(risks) < 10: row_fails.append(f'risk_count={len(risks)}')
                    if not v.get('ok'): row_fails.append(f"v150: {v.get('failures')}")
                    if row_fails: sc_fails.append({'project':name,'scenario':sc,'fails':row_fails})
                except Exception as e:
                    sc_fails.append({'project':name,'scenario':sc,'error':str(e)[:100]})
                sc_count += 1

    total_showcase = sum(len(v) for v in showcase_lib.values())
    print(f"  Showcase: {sc_count} outputs ({total_showcase} projects × 5 scenarios) | {len(sc_fails)} failures")
    for f in sc_fails[:10]: print(f"    {f}")

    # Exports
    print(f"\n  Export smoke test: 200 models × 7 export functions...")
    exp_count, exp_fails = run_exports(200)
    print(f"  Exports: {exp_count} checks | {len(exp_fails)} failures")
    for f in exp_fails[:10]: print(f"    {f}")

    # Final
    grand_total_fails = len(all_fails) + len(sc_fails) + len(exp_fails)
    grand_total_checks = (N_EARTH + N_SPACE) + sc_count + exp_count
    verdict = 'ALL PASS ✓' if grand_total_fails == 0 else f'FAIL — {grand_total_fails} issues'

    print(f"\n{'='*70}")
    print(f"CASEY V150 QA COMPLETE")
    print(f"  Total checks:  {grand_total_checks:,}")
    print(f"  Earth models:  {N_EARTH:,} | fails: {len(earth_fails)}")
    print(f"  Space models:  {N_SPACE:,} | fails: {len(space_fails)}")
    print(f"  Showcase:      {sc_count} | fails: {len(sc_fails)}")
    print(f"  Exports:       {exp_count} | fails: {len(exp_fails)}")
    print(f"  VERDICT:       {verdict}")
    print(f"{'='*70}")

    out = {'version': main.APP_VERSION, 'total_checks': grand_total_checks, 'total_failures': grand_total_fails,
           'earth': {'count': N_EARTH, 'failures': len(earth_fails), 'by_check': {k:len(v) for k,v in by_check.items() if any(x.get('kind')=='earth' for x in v)}},
           'space': {'count': N_SPACE, 'failures': len(space_fails), 'by_check': {k:len(v) for k,v in by_check.items() if any(x.get('kind')=='space' for x in v)}},
           'showcase': {'count': sc_count, 'failures': len(sc_fails), 'samples': sc_fails[:10]},
           'exports': {'count': exp_count, 'failures': len(exp_fails), 'samples': exp_fails[:15]},
           'verdict': verdict}
    with open(os.path.join(ROOT, 'CASEY_V150_FULL_QA_RESULTS.json'), 'w') as fh:
        json.dump(out, fh, indent=2, default=str)
    print(f"Results saved → CASEY_V150_FULL_QA_RESULTS.json")
