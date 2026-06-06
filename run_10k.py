"""
CASEY V147 — 10,000 Earth + 5,000 Space batch QA runner.
Uses multiprocessing for speed. Results written to run_10k_results.json.
Usage: python run_10k.py [--workers N]
"""
import os, sys, re, random, time, json, argparse, multiprocessing as mp

_HERE = os.path.dirname(os.path.abspath(__file__))
_BACKEND = os.path.join(_HERE, 'backend')

prompts_earth = [
    'rail transit with tunnelling utility relocation signalling systems migration operator acceptance',
    'AI hyperscale data centre with grid interconnection transformers cooling GPU procurement',
    'airport terminal expansion with baggage systems airside phasing ORAT regulator acceptance',
    'SMR nuclear rollout with licensing safety case containment qualification grid integration',
    'GMP pharma manufacturing campus with sterile fill finish validation clean utilities regulatory release',
    'LNG export terminal with cryogenic systems marine berths long lead equipment commissioning',
    'offshore wind mega-hub with grid connection substations marine logistics installation vessels',
    'hospital clinical campus with MEP utilities clean rooms operating theatres phased handover',
    'defence base radar command infrastructure with secure power hardened facilities',
    'hydrogen export corridor with electrolysis storage compression marine loading',
]
prompts_space = [
    'lunar habitat infrastructure with life support thermal survivability autonomous commissioning launch windows',
    'Mars cargo logistics network with radiation hardening communication latency ISRU cargo landers',
    'orbital data centre with radiation hardening autonomous servicing orbital cooling power redundancy',
    'Amazon Kuiper constellation with satellite production launch manifest ground stations spectrum coordination',
    'autonomous orbital servicing platform with rendezvous proximity operations robotic capture refuelling',
    'SpaceX Starship launch infrastructure with orbital refuelling booster recovery pad infrastructure regulatory approvals',
    'cislunar propellant depot with autonomous docking cryogenic storage power management',
    'lunar south pole power grid with solar arrays regolith shielding cable deployment rovers',
    'Mars base ISRU oxygen production with electrolysis compression storage life support integration',
    'space domain awareness ground station network with secure comms encrypted uplink antenna farms',
]
scens = ['base', 'faster', 'cheaper', 'lower_risk', 'premium']


def _worker_init():
    """Called once per process — import main into the worker global namespace."""
    global main
    sys.path.insert(0, _BACKEND)
    os.chdir(_BACKEND)
    import main as _main
    main = _main


def bn(x):
    s = str(x or '').replace('$', '').replace(',', '').strip().upper()
    try:
        if s.endswith('T'): return float(s[:-1]) * 1000
        if s.endswith('B'): return float(s[:-1])
        if s.endswith('M'): return float(s[:-1]) / 1000
        return float(re.search(r'-?\d+(?:\.\d+)?', s).group(0))
    except Exception:
        return 0.0


def months(x):
    m = re.search(r'-?\d+(?:\.\d+)?', str(x or ''))
    return int(round(float(m.group(0)))) if m else 0


def task(args):
    kind, i = args
    pool = prompts_space if kind == 'space' else prompts_earth
    base = random.choice(pool)
    sc = random.choice(scens)
    m = main.build_model(f'{kind} QA {i}: {base}', 'QA', random.randint(2, 5), random.randint(2, 5), sc)
    fails = []
    p50 = bn(m.get('cost_p50'))
    tot = bn(m.get('direct_cost')) + bn(m.get('indirect_cost')) + bn(m.get('risk_reserve'))
    if abs(tot - p50) > 0.015:
        fails.append({'check': 'bucket', 'kind': kind, 'i': i, 'p50': m.get('cost_p50'), 'sum': round(tot, 3)})
    mc = m.get('monte_carlo') or {}
    q = mc.get('qcra') or {}; s = mc.get('qsra') or {}
    if abs(float(q.get('p50') or 0) - round(p50, 1)) > 0.11:
        fails.append({'check': 'qcra', 'kind': kind, 'i': i, 'qcra': q.get('p50'), 'p50': round(p50, 1)})
    sp50 = int(s.get('p50') or 0); sm = months(m.get('schedule'))
    if sp50 != sm:
        fails.append({'check': 'qsra', 'kind': kind, 'i': i, 'qsra': sp50, 'sched': sm})
    summ = str(m.get('executive_summary') or '')
    if m.get('cost_p50') not in summ or str(m.get('confidence_pct')) + '%' not in summ:
        fails.append({'check': 'summary', 'kind': kind, 'i': i, 'p50': m.get('cost_p50')})
    sc_val = m.get('scenario')
    for r in m.get('scenario_matrix') or []:
        if str(r.get('scenario', '')).lower().replace(' ', '_') == str(sc_val):
            if abs(bn(r.get('cost_p50')) - p50) > 0.11 or months(r.get('schedule_months')) != sm:
                fails.append({'check': 'matrix', 'kind': kind, 'i': i})
            break
    return fails


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--workers', type=int, default=min(4, mp.cpu_count()))
    args = parser.parse_args()

    inputs = [('earth', i) for i in range(10000)] + [('space', i) for i in range(5000)]
    random.shuffle(inputs)

    fail = []
    start = time.time()
    print(f'Running {len(inputs):,} models with {args.workers} workers...')

    with mp.Pool(args.workers, initializer=_worker_init) as pool:
        for res in pool.imap_unordered(task, inputs, chunksize=25):
            if res:
                fail.extend(res)
                if len(fail) % 10 == 0:
                    print(f'  Failures so far: {len(fail)}')

    elapsed = time.time() - start
    result = {
        'total': len(inputs),
        'earth': 10000,
        'space': 5000,
        'fail_count': len(fail),
        'by_check': {},
        'elapsed_s': round(elapsed, 1),
        'rate_per_s': round(len(inputs) / elapsed, 1),
    }
    for f2 in fail:
        result['by_check'].setdefault(f2['check'], []).append(f2)
    result['by_check'] = {k: {'count': len(v), 'samples': v[:5]} for k, v in result['by_check'].items()}
    out = os.path.join(_HERE, 'run_10k_results.json')
    with open(out, 'w') as fh:
        json.dump(result, fh, indent=2, default=str)
    print(json.dumps({k: v for k, v in result.items() if k != 'by_check'}, indent=2))
    print(f'Full results written to {out}')
