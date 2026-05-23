"""
CASEY public demo quality harness.
Run from project root after backend is live locally:
  cd backend
  uvicorn main:app --reload
Then in another terminal:
  python tools/quality_harness.py --n 100

For 100,000 examples, run on a machine/server and export results:
  python tools/quality_harness.py --n 100000 --fast

This does not train the model live. It checks structural quality and logs weak cases for prompt/model improvement.
"""
import argparse, json, random, urllib.request

EARTH = [
    "500MW AI hyperscale data centre campus in Riyadh with grid connection, liquid cooling and accelerated 2027 delivery",
    "New airport terminal expansion with baggage systems, live operations, security screening and phased passenger transition",
    "Advanced semiconductor fab in Arizona with cleanrooms, ultrapure water, grid upgrades and fast-track procurement",
    "Life sciences GMP manufacturing campus in Boston with cleanrooms, validation, automation and regulatory handover",
    "Green hydrogen production hub with electrolyzers, grid connection, storage, export terminal and water treatment",
    "High-speed rail station redevelopment with tunnelling interfaces, possessions, signalling and urban stakeholder constraints",
]
SPACE = [
    "Lunar surface logistics hub with landing pads, power storage, regolith roads, autonomous vehicles and crew safety systems",
    "Orbital compute platform in low Earth orbit with solar arrays, thermal systems, launch logistics and ground segment integration",
    "Mars ISRU fuel refinery producing methane and oxygen with nuclear power, robotic mining and long-duration maintenance",
    "Cislunar propellant depot with cryogenic storage, docking systems, launch dependency and autonomous operations",
    "Orbital hospital platform with life support, medical modules, crew transfer, radiation protection and emergency return capability",
]

def post(payload):
    data=json.dumps(payload).encode()
    req=urllib.request.Request('http://127.0.0.1:8000/public-demo/generate', data=data, headers={'Content-Type':'application/json'}, method='POST')
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read().decode())

def validate(r):
    model=r.get('model',{})
    needed=['cost_p50','cost_range','schedule','risk','confidence_pct','risks','schedule_rows','cost_lines']
    missing=[x for x in needed if not model.get(x)]
    return missing

if __name__ == '__main__':
    ap=argparse.ArgumentParser(); ap.add_argument('--n', type=int, default=20); ap.add_argument('--fast', action='store_true')
    args=ap.parse_args(); fails=[]
    for i in range(args.n):
        kind='Space' if i % 11 == 0 else 'Earth'
        brief=random.choice(SPACE if kind=='Space' else EARTH)
        payload={'email':f'test{i}@example.com','project_type':kind,'project_description':brief,'location':'Global','size_or_capacity':'Concept scale','stage':'Concept / early feasibility','biggest_concern':'Cost, schedule and risk confidence','fingerprint':f'test-{i}','client_token':f'token-{i}'}
        try:
            r=post(payload); m=validate(r)
            if m: fails.append({'i':i,'missing':m,'brief':brief})
            print(i, kind, 'OK' if not m else 'FAIL', brief[:60])
        except Exception as e:
            fails.append({'i':i,'error':str(e),'brief':brief}); print(i,'ERROR',e)
    print('FAILURES', len(fails)); print(json.dumps(fails[:20], indent=2))
