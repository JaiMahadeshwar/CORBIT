import os, sys, re, random, time, json, contextlib, io, multiprocessing as mp
root=os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(root,'backend'))
os.chdir(os.path.join(root,'backend'))
with contextlib.redirect_stdout(io.StringIO()):
    import main
EARTH=['California High-Speed Rail rail transit tunnelling utility relocation signalling systems migration operator acceptance','Microsoft AI hyperscale data centre grid interconnection transformers cooling GPU procurement','NEOM giga project autonomous mobility logistics utilities workforce scaling','Eli Lilly GMP pharma manufacturing sterile fill finish validation clean utilities regulatory release','SMR nuclear rollout licensing safety case containment qualification grid integration','AUKUS shipyard defence industrial base workforce nuclear certification secure supply chain','LNG export terminal cryogenic systems marine berths long lead valves commissioning','TSMC Arizona semiconductor fab cleanrooms ultrapure water tool install yield ramp','airport terminal expansion baggage systems live operations ORAT security phasing','offshore wind mega-hub export cables substations vessel constraints weather windows']
SPACE=['lunar habitat infrastructure life support thermal survivability autonomous commissioning launch windows','Mars cargo logistics network radiation hardening communication latency ISRU cargo landers','orbital data centre radiation hardening autonomous servicing orbital cooling power redundancy','Amazon Kuiper constellation satellite production launch manifest ground stations spectrum coordination','AST SpaceMobile orbital telecom network satellite deployment telecom interoperability spectrum approvals','SpaceX Starship industrialization orbital refueling thermal protection booster recovery pad infrastructure','autonomous orbital servicing rendezvous proximity robotic capture refuelling debris avoidance','lunar resource extraction autonomous mining regolith processing ISRU surface power','commercial space station docking compatibility EVA readiness launch manifest reliability','sovereign launch infrastructure spaceport range safety pad construction launch cadence']
SCENS=['base','faster','cheaper','lower_risk','premium']
def bn(x):
 s=str(x or '').replace('$','').replace(',','').strip().upper()
 try:
  if s.endswith('T'): return float(s[:-1])*1000
  if s.endswith('B'): return float(s[:-1])
  if s.endswith('M'): return float(s[:-1])/1000
  m=re.search(r'-?\d+(?:\.\d+)?', s); return float(m.group(0)) if m else 0.0
 except Exception: return 0.0
def months(x):
 m=re.search(r'-?\d+(?:\.\d+)?',str(x or '')); return int(round(float(m.group(0)))) if m else 0
def check(args):
 kind,i=args
 rnd=random.Random(i + (999999 if kind=='space' else 0))
 p=rnd.choice(SPACE if kind=='space' else EARTH); sc=rnd.choice(SCENS)
 m=main.build_model(f'{kind} QA {i}: {p}','QA',rnd.randint(2,5),rnd.randint(2,5),sc)
 fails=[]; p50=bn(m.get('cost_p50')); total=bn(m.get('direct_cost'))+bn(m.get('indirect_cost'))+bn(m.get('risk_reserve'))
 if abs(total-p50)>0.11: fails.append({'check':'bucket','kind':kind,'i':i,'p50':p50,'total':total})
 q=(m.get('monte_carlo') or {}).get('qcra') or {}; s=(m.get('monte_carlo') or {}).get('qsra') or {}
 if abs(float(q.get('p50') or 0)-round(p50,1))>0.11: fails.append({'check':'qcra','kind':kind,'i':i,'qcra':q.get('p50'),'p50':round(p50,1)})
 if int(s.get('p50') or 0)!=months(m.get('schedule')): fails.append({'check':'qsra','kind':kind,'i':i,'qsra':s.get('p50'),'schedule':m.get('schedule')})
 summ=str(m.get('executive_summary') or '')
 for tok in [m.get('cost_p50'), m.get('schedule'), str(m.get('confidence_pct'))+'%']:
  if tok and str(tok) not in summ: fails.append({'check':'summary','kind':kind,'i':i,'missing':tok})
 if not m.get('scenario_signature'): fails.append({'check':'signature','kind':kind,'i':i})
 v=main._v148_validate_model(m)
 if not v.get('ok'): fails.append({'check':'validator','kind':kind,'i':i,'failures':v.get('failures')[:3]})
 return fails
if __name__=='__main__':
 items=[('earth',i) for i in range(10000)]+[('space',i) for i in range(5000)]
 start=time.time(); fails=[]
 with mp.Pool(processes=min(16, os.cpu_count() or 4)) as pool:
  for res in pool.imap_unordered(check, items, chunksize=50):
   if res: fails.extend(res)
 elapsed=time.time()-start
 out={'total':len(items),'earth':10000,'space':5000,'fail_count':len(fails),'elapsed_s':round(elapsed,1),'samples':fails[:20]}
 os.chdir(root); open('CASEY_V148_15000_QA_RESULTS.json','w').write(json.dumps(out,indent=2))
 print(json.dumps(out,indent=2))
