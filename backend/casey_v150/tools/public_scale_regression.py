
"""
CASEY public-scale regression harness.

Run after backend is available locally:
  cd backend
  uvicorn main:app --reload

Then:
  python tools/public_scale_regression.py

This checks broad prompt coverage and output structure. It does not certify commercial estimates.
"""
import json, random, urllib.request, statistics

EARTH = [
  "Class 5 concept for a 500MW AI hyperscale data centre campus in Riyadh with grid substation, liquid cooling, phased halls, backup generation and accelerated delivery.",
  "Class 4 feasibility for a semiconductor fab in Arizona with 60,000m2 cleanroom, UPW, chemicals, process tools and grid upgrade.",
  "Class 3 budget estimate for an airport terminal expansion in London with baggage systems, security, airside works and live operational phasing.",
  "Class 3 estimate for a defence command and control facility in Qatar with secure comms, hardened operations rooms and resilient power.",
]
SPACE = [
  "Class 5 concept for a lunar south pole logistics hub with landing pads, regolith roads, autonomous cargo rovers, power storage and launch cadence dependency.",
  "Class 4 feasibility for a Mars ISRU fuel plant producing methane and oxygen with mining rovers, reactors, storage and long-duration autonomy.",
  "Class 3 budget estimate for an orbital compute platform in LEO with solar arrays, compute modules, thermal control and ground segment integration.",
]

def post(payload):
    data=json.dumps(payload).encode()
    req=urllib.request.Request("http://127.0.0.1:8000/public-demo/generate", data=data, headers={"Content-Type":"application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read().decode())

def validate(model):
    needed=["cost_p10","cost_p50","cost_p90","schedule","risks","cost_lines","schedule_rows","estimate_quality_index","benchmark_memory","procurement_heatmap","critical_path_narrative"]
    return [k for k in needed if not model.get(k)]

if __name__ == "__main__":
    fails=[]
    cases=[]
    for i in range(20):
        kind="Space" if i%5==0 else "Earth"
        brief=random.choice(SPACE if kind=="Space" else EARTH)
        payload={
          "email":f"regression{i}@example.com",
          "project_type":kind,
          "project_description":brief,
          "location":"Riyadh" if kind=="Earth" else "Lunar surface",
          "size_or_capacity":"500MW" if kind=="Earth" else "Phase one",
          "stage":"Concept / early feasibility",
          "biggest_concern":"Cost, schedule and risk confidence",
          "fingerprint":f"regression-{i}",
          "client_token":f"token-{i}"
        }
        try:
            r=post(payload)
            missing=validate(r["model"])
            if missing: fails.append((i,missing,brief))
            print(i, kind, "OK" if not missing else "FAIL", missing)
        except Exception as e:
            fails.append((i,str(e),brief))
            print(i, "ERROR", e)
    print("FAILURES", len(fails))
    print(json.dumps(fails[:10], indent=2))
