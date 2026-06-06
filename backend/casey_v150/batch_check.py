"""
CASEY V147 — single-batch QA checker for CI or parallelised runs.
Usage:
  python batch_check.py --kind earth --start 0 --n 1000
  python batch_check.py --kind space --start 0 --n 500
Results printed as JSON to stdout.
"""
import os, sys, re, random, json, argparse, time

_HERE = os.path.dirname(os.path.abspath(__file__))
_BACKEND = os.path.join(_HERE, 'backend')
sys.path.insert(0, _BACKEND)
os.chdir(_BACKEND)
import main

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
]
scens = ['base', 'faster', 'cheaper', 'lower_risk', 'premium']


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


if __name__ == '__main__':
    p = argparse.ArgumentParser()
    p.add_argument('--kind', choices=['earth', 'space'], required=True)
    p.add_argument('--start', type=int, default=0)
    p.add_argument('--n', type=int, default=1000)
    args = p.parse_args()

    pool = prompts_space if args.kind == 'space' else prompts_earth
    fail = []
    t0 = time.time()

    for i in range(args.start, args.start + args.n):
        sc = random.choice(scens)
        m = main.build_model(
            f'{args.kind} QA {i}: {random.choice(pool)}', 'QA',
            random.randint(2, 5), random.randint(2, 5), sc
        )
        p50 = bn(m.get('cost_p50'))
        tot = bn(m.get('direct_cost')) + bn(m.get('indirect_cost')) + bn(m.get('risk_reserve'))
        if abs(tot - p50) > 0.015:
            fail.append({'check': 'bucket', 'i': i, 'p50': m.get('cost_p50'), 'sum': round(tot, 3)})
        mc = m.get('monte_carlo') or {}; q = mc.get('qcra') or {}; s = mc.get('qsra') or {}
        if abs(float(q.get('p50') or 0) - round(p50, 1)) > 0.11:
            fail.append({'check': 'qcra', 'i': i})
        if int(s.get('p50') or 0) != months(m.get('schedule')):
            fail.append({'check': 'qsra', 'i': i})
        summ = str(m.get('executive_summary') or '')
        if m.get('cost_p50') not in summ or str(m.get('confidence_pct')) + '%' not in summ:
            fail.append({'check': 'summary', 'i': i, 'p50': m.get('cost_p50')})

    elapsed = time.time() - t0
    print(json.dumps({
        'kind': args.kind, 'start': args.start, 'n': args.n,
        'fail_count': len(fail),
        'elapsed_s': round(elapsed, 1),
        'rate_per_s': round(args.n / elapsed, 1),
        'first_failures': fail[:5],
    }, indent=2, default=str))
