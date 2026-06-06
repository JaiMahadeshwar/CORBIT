import os, sys, zipfile, json, re, io
root=os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, os.path.join(root,'backend')); os.chdir(os.path.join(root,'backend'))
import main
m=main.build_model('California High-Speed Rail with tunnelling utility relocation signalling systems migration operator acceptance','QA',3,4,'faster')
print(m['cost_p50'],m['schedule'],m['confidence_pct'],m['executive_summary'])
for name, func in [('workbook',main.workbook_bytes),('risk',main.risk_register_workbook_bytes),('xer',main.xer_bytes),('word',main.word_bytes),('pdf',main.pdf_bytes),('pptx',main.pptx_bytes)]:
 try:
  b=func(m); print(name, len(b), b[:4])
 except Exception as e:
  print('ERR', name, type(e), e)
try:
 resp=main.export_all(m); print('export_all resp', type(resp), getattr(resp,'media_type',None), hasattr(resp,'body'), getattr(resp,'body',None) and len(resp.body))
except Exception as e: print('export_all err', e)
