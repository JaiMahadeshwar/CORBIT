import os, sys, json, zipfile, io
root=os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, os.path.join(root,'backend')); os.chdir(os.path.join(root,'backend'))
from fastapi.testclient import TestClient
import main
c=TestClient(main.app)
m=main.build_model('California High-Speed Rail with tunnelling utility relocation signalling systems migration operator acceptance','QA',3,4,'faster')
for path in ['/export/workbook','/export/risk-register','/export/xer','/export/word','/export/pdf','/export/pptx','/export/json','/export/qcra-qsra','/export/all']:
 r=c.post(path,json=m)
 print(path,r.status_code,r.headers.get('content-type'),len(r.content),r.content[:4])
 if path=='/export/all':
  z=zipfile.ZipFile(io.BytesIO(r.content)); print(z.namelist())
