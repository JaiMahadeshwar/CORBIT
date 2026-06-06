import os, sys, re
root=os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(root,'backend'))
os.chdir(os.path.join(root,'backend'))
import main
prompts=['California High-Speed Rail with tunnelling utility relocation signalling operator acceptance','Microsoft AI hyperscale data centre with grid transformer cooling GPU procurement','Lunar habitat infrastructure life support thermal survivability autonomous commissioning launch windows','Starship launch industrialization orbital refueling booster recovery pad infrastructure']
scens=['base','faster','cheaper','lower_risk','premium']
def bn(x):
 s=str(x or '').replace('$','').replace(',','').upper().strip()
 try:
  if s.endswith('T'): return float(s[:-1])*1000
  if s.endswith('B'): return float(s[:-1])
  if s.endswith('M'): return float(s[:-1])/1000
  return float(re.sub(r'[^0-9.-]','',s) or 0)
 except: return 0
fail=[]
for p in prompts:
 for sc in scens:
  m=main.build_model(p,'QA',3,4,sc)
  p50=bn(m.get('cost_p50')); tot=bn(m.get('direct_cost'))+bn(m.get('indirect_cost'))+bn(m.get('risk_reserve'))
  q=(m.get('monte_carlo') or {}).get('qcra') or {}; qs=(m.get('monte_carlo') or {}).get('qsra') or {}
  print(sc, m.get('title'), m.get('cost_p50'), m.get('schedule'), m.get('confidence_pct'), 'sum', round(tot,1), 'q', q, 'qs', qs, 'summary ok', m.get('cost_p50') in m.get('executive_summary',''))
  if abs(tot-p50)>0.11: fail.append((p,sc,p50,tot))
print('fail', fail[:5])
