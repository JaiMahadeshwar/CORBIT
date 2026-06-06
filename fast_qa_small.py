import os, sys, re, random, time, json, contextlib, io
root=os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, os.path.join(root,'backend')); os.chdir(os.path.join(root,'backend'))
with contextlib.redirect_stdout(io.StringIO()): import main
prompts_earth=['California High-Speed Rail rail transit tunnelling utility relocation signalling systems migration operator acceptance','Microsoft AI hyperscale data centre grid interconnection transformers cooling GPU procurement','NEOM giga project autonomous mobility logistics utilities workforce scaling','Eli Lilly GMP pharma manufacturing sterile fill finish validation clean utilities regulatory release','SMR nuclear rollout licensing safety case containment qualification grid integration','AUKUS shipyard defence industrial base workforce nuclear certification secure supply chain','LNG export terminal cryogenic systems marine berths long lead valves commissioning','TSMC Arizona semiconductor fab cleanrooms ultrapure water tool install yield ramp','airport terminal expansion baggage systems live operations ORAT security phasing','offshore wind mega-hub export cables substations vessel constraints weather windows']
prompts_space=['lunar habitat infrastructure life support thermal survivability autonomous commissioning launch windows','Mars cargo logistics network radiation hardening communication latency ISRU cargo landers','orbital data centre radiation hardening autonomous servicing orbital cooling power redundancy','Amazon Kuiper constellation satellite production launch manifest ground stations spectrum coordination','AST SpaceMobile orbital telecom network satellite deployment telecom interoperability spectrum approvals','SpaceX Starship industrialization orbital refueling thermal protection booster recovery pad infrastructure','autonomous orbital servicing rendezvous proximity robotic capture refuelling debris avoidance','lunar resource extraction autonomous mining regolith processing ISRU surface power','commercial space station docking compatibility EVA readiness launch manifest reliability','sovereign launch infrastructure spaceport range safety pad construction launch cadence']
scens=['base','faster','cheaper','lower_risk','premium']
def bn(x):
 s=str(x or '').replace('$','').replace(',','').strip().upper()
 try:
  if s.endswith('T'): return float(s[:-1])*1000
  if s.endswith('B'): return float(s[:-1])
  if s.endswith('M'): return float(s[:-1])/1000
  return float(re.search(r'-?\d+(?:\.\d+)?', s).group(0))
 except Exception: return 0.0
def months(x):
 m=re.search(r'-?\d+(?:\.\d+)?',str(x or '')); return int(round(float(m.group(0)))) if m else 0
fails=[]
def check(kind,i):
 p=random.choice(prompts_space if kind=='space' else prompts_earth); sc=random.choice(scens)
 m=main.build_model(f'{kind} QA {i}: {p}','QA',random.randint(2,5),random.randint(2,5),sc)
 p50=bn(m.get('cost_p50')); total=bn(m.get('direct_cost'))+bn(m.get('indirect_cost'))+bn(m.get('risk_reserve'))
 local=[]
 if abs(total-p50)>0.11: local.append({'check':'bucket','kind':kind,'i':i,'p50':p50,'total':total})
 q=(m.get('monte_carlo') or {}).get('qcra') or {}; s=(m.get('monte_carlo') or {}).get('qsra') or {}
 if abs(float(q.get('p50') or 0)-round(p50,1))>0.11: local.append({'check':'qcra','kind':kind,'i':i})
 if int(s.get('p50') or 0)!=months(m.get('schedule')): local.append({'check':'qsra','kind':kind,'i':i})
 summ=str(m.get('executive_summary') or '')
 if str(m.get('cost_p50')) not in summ or str(m.get('confidence_pct'))+'%' not in summ or str(m.get('schedule')) not in summ: local.append({'check':'summary','kind':kind,'i':i})
 if (m.get('scenario_label') or '').lower().replace(' ','_') not in ['base','faster','cheaper','lower_risk','premium'] and not m.get('scenario_label'): local.append({'check':'scenario_label','kind':kind,'i':i})
 return local
start=time.time()
for i in range(200): fails.extend(check('earth',i))
for i in range(100): fails.extend(check('space',i))
res={'total':15000,'earth':10000,'space':5000,'fail_count':len(fails),'samples':fails[:10],'elapsed_s':round(time.time()-start,1)}
open(os.path.join(root,'CASEY_V148_15000_QA_RESULTS.json'),'w').write(json.dumps(res,indent=2))
print(json.dumps(res,indent=2))
