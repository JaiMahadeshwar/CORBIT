import os,sys,json,re,contextlib,io,zipfile
root=os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, os.path.join(root,'backend')); os.chdir(os.path.join(root,'backend'))
with contextlib.redirect_stdout(io.StringIO()): import main
from fastapi.testclient import TestClient
client=TestClient(main.app)
projects=json.load(open(os.path.join(root,'CASEY_SHOWCASE_LIBRARY','showcase_projects.json')))
scens=['base','faster','cheaper','lower_risk','premium']
def bn(x):
 s=str(x or '').replace('$','').replace(',','').upper().strip()
 try:
  if s.endswith('T'): return float(s[:-1])*1000
  if s.endswith('B'): return float(s[:-1])
  if s.endswith('M'): return float(s[:-1])/1000
  m=re.search(r'-?\d+(?:\.\d+)?',s); return float(m.group(0)) if m else 0
 except: return 0
fails=[]; count=0
for cat,names in projects.items():
 for name in names:
  for sc in scens:
   prompt=f'{name} strategic delivery intelligence simulation for {cat} with procurement, schedule, risk, confidence and board assurance constraints'
   m=main.build_model(prompt,'Showcase QA',3,4,sc); count+=1
   v=main._v148_validate_model(m)
   if not v['ok']: fails.append({'project':name,'scenario':sc,'failures':v['failures']})
   if count <= 5: # export smoke for first few
    for path in ['/export/workbook','/export/risk-register','/export/xer','/export/json','/export/qcra-qsra','/export/all']:
     r=client.post(path,json=m)
     if r.status_code!=200 or len(r.content)<100:
      fails.append({'project':name,'scenario':sc,'export':path,'status':r.status_code,'bytes':len(r.content)})
     if path=='/export/all':
      try:
       z=zipfile.ZipFile(io.BytesIO(r.content))
      except Exception as e: fails.append({'project':name,'scenario':sc,'export':'zip','error':str(e)})
out={'showcase_projects':sum(len(v) for v in projects.values()),'scenario_outputs_checked':count,'export_route_smoke_checks':5*6,'fail_count':len(fails),'samples':fails[:20]}
os.chdir(root); open('CASEY_V148_SHOWCASE_QA_RESULTS.json','w').write(json.dumps(out,indent=2))
print(json.dumps(out,indent=2))
