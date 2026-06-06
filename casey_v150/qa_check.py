"""
CASEY V147 QA Harness — runs all showcase prompts x all 5 scenarios,
then 5000 Earth + 5000 Space random generations, checking:
  - bucket sum (Direct + Indirect + Reserve == P50)
  - QCRA P50 matches model P50
  - QSRA P50 matches schedule months
  - Executive summary contains headline P50 and confidence%
  - Selected scenario matrix row agrees with model headline
  - V147 version stamp present
Usage: python qa_check.py
"""
import sys, os, re, json, random, time
# Support running from project root or backend dir
_HERE = os.path.dirname(os.path.abspath(__file__))
_BACKEND = os.path.join(_HERE, 'backend')
sys.path.insert(0, _BACKEND)
os.chdir(_BACKEND)
import main

def bn(x):
    try:
        if isinstance(x, (int, float)): return float(x)
        s = str(x or '').replace('$', '').replace(',', '').strip().upper()
        if s.endswith('T'): return float(s[:-1]) * 1000
        if s.endswith('B'): return float(s[:-1])
        if s.endswith('M'): return float(s[:-1]) / 1000
        m = re.search(r'-?\d+(?:\.\d+)?', s)
        return float(m.group(0)) if m else 0.0
    except Exception:
        return 0.0

def months(x):
    m = re.search(r'-?\d+(?:\.\d+)?', str(x or ''))
    return int(round(float(m.group(0)))) if m else 0

scenarios = ['base', 'faster', 'cheaper', 'lower_risk', 'premium']

# Extract all showcase prompts from frontend App.jsx
_APP_JSX = os.path.join(_HERE, 'frontend', 'src', 'App.jsx')
showcase_prompts = []
if os.path.exists(_APP_JSX):
    text = open(_APP_JSX, encoding='utf-8').read()
    showcase_prompts = re.findall(r"prompt:'([^']+)'", text)

earth_terms = [
    'rail transit with tunnelling utility relocation signalling systems migration operator acceptance',
    'AI hyperscale data centre with 500MW grid interconnection transformers cooling GPU procurement',
    'airport terminal expansion with baggage systems airside phasing ORAT regulator acceptance',
    'SMR nuclear rollout with licensing safety case containment qualification grid integration',
    'GMP pharma manufacturing campus with sterile fill finish validation clean utilities regulatory release',
    'LNG export terminal with cryogenic systems marine berths long lead equipment commissioning',
]
space_terms = [
    'lunar habitat infrastructure with life support thermal survivability autonomous commissioning launch windows',
    'Mars cargo logistics network with radiation hardening communication latency ISRU cargo landers',
    'orbital data centre with radiation hardening autonomous servicing orbital cooling power redundancy',
    'Amazon Kuiper constellation with satellite production launch manifest ground stations spectrum coordination',
    'autonomous orbital servicing platform with rendezvous proximity operations robotic capture refuelling',
    'SpaceX Starship launch infrastructure with orbital refuelling booster recovery pad infrastructure regulatory approvals',
]

fail = []
checked = 0

def check_model(m, tag):
    global checked, fail
    checked += 1
    p50 = bn(m.get('cost_p50'))
    d = bn(m.get('direct_cost')); ind = bn(m.get('indirect_cost')); res = bn(m.get('risk_reserve'))
    diff = abs((d + ind + res) - p50)
    if diff > 0.015:
        fail.append({'check': 'bucket_sum', 'tag': tag, 'p50': m.get('cost_p50'), 'sum': round(d+ind+res, 3), 'diff': round(diff, 4)})
    mc = m.get('monte_carlo') or {}; q = mc.get('qcra') or {}; s = mc.get('qsra') or {}
    qp50 = float(q.get('p50') or 0)
    if abs(qp50 - round(p50, 1)) > 0.11:
        fail.append({'check': 'qcra_p50', 'tag': tag, 'qcra': qp50, 'model_p50': round(p50, 1)})
    sp50 = int(s.get('p50') or 0); sm = months(m.get('schedule'))
    if sp50 != sm:
        fail.append({'check': 'qsra_p50', 'tag': tag, 'qsra': sp50, 'sched': sm})
    summ = str(m.get('executive_summary') or '')
    if m.get('cost_p50') not in summ or str(m.get('confidence_pct')) + '%' not in summ:
        fail.append({'check': 'summary_missing_headline', 'tag': tag, 'p50': m.get('cost_p50'), 'conf': m.get('confidence_pct')})
    sc = m.get('scenario')
    for r in m.get('scenario_matrix') or []:
        if str(r.get('scenario', '')).lower().replace(' ', '_') == str(sc):
            if abs(bn(r.get('cost_p50')) - p50) > 0.11 or months(r.get('schedule_months')) != sm:
                fail.append({'check': 'matrix_mismatch', 'tag': tag, 'matrix': r.get('cost_p50'), 'model': m.get('cost_p50')})
            break
    if 'V14' not in str(m.get('version', '')):
        fail.append({'check': 'version', 'tag': tag, 'v': m.get('version')})

start = time.time()

# Showcase prompts x all scenarios
for prompt in showcase_prompts:
    for sc in scenarios:
        check_model(main.build_model(prompt, 'QA', 3, 4, sc), f'showcase:{sc}:{prompt[:35]}')

# 5000 Earth
for i in range(5000):
    check_model(main.build_model(
        f'QA Earth {i}: {random.choice(earth_terms)}', 'QA',
        random.randint(2, 5), random.randint(2, 5), random.choice(scenarios)
    ), f'earth:{i}')

# 5000 Space
for i in range(5000):
    check_model(main.build_model(
        f'QA Space {i}: {random.choice(space_terms)}', 'QA',
        random.randint(2, 5), random.randint(2, 5), random.choice(scenarios)
    ), f'space:{i}')

elapsed = time.time() - start
by_check = {}
for f2 in fail:
    by_check.setdefault(f2['check'], []).append(f2)

result = {
    'checked': checked,
    'fail_count': len(fail),
    'by_check': {k: {'count': len(v), 'samples': v[:3]} for k, v in by_check.items()},
    'showcase_prompts': len(showcase_prompts),
    'elapsed_s': round(elapsed, 1),
    'rate_per_s': round(checked / elapsed, 1),
}
print(json.dumps(result, indent=2, default=str))
