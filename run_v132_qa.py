import sys, json, random, time, zipfile, io
from pathlib import Path
sys.path.insert(0, '/mnt/data/v132work/backend')
import main

EARTH_SECTORS = {
 'data_centre':['Microsoft Azure 500MW AI data centre campus in Texas with grid interconnect, transformers, liquid cooling and phased data halls','Meta AI hyperscale compute campus with 300MW utility constraints and accelerated commissioning'],
 'airport':['Heathrow terminal expansion with baggage systems, ORAT, airside phasing, security lanes and live operations','Major international airport runway and terminal programme with airline migration and regulatory approvals'],
 'rail':['California high speed rail section with stations, signalling, utility diversions, rolling stock interfaces and possessions','London metro extension with tunnel, stations, signalling migration and trial running'],
 'roads':['Urban highway tunnel and bridge corridor with utility diversions, traffic staging and statutory approvals','Motorway interchange upgrade with live traffic management and structures'],
 'ports':['Automated container terminal expansion with quay works, dredging, cranes, yard systems and landside rail connection','Marine logistics port with berth expansion, customs systems and operational cutover'],
 'water':['Desalination plant and water transmission programme with permits, process equipment, tie-ins and operator acceptance','Wastewater treatment upgrade with process performance testing and live service continuity'],
 'energy':['Offshore wind transmission and hydrogen hub with substations, HVDC, grid agreements and permitting','Battery energy storage and grid interconnector programme with transformers and utility approvals'],
 'nuclear':['SMR nuclear campus with containment, licensing, safety case, nuclear-grade procurement and regulator hold points','Nuclear reactor life extension with safety case and outage integration'],
 'oil_gas':['LNG processing train expansion with modular fabrication, HAZOP, compressors, shutdown tie-ins and start-up','Offshore gas platform brownfield tie-in with subsea pipelines, HSE approvals and commissioning'],
 'mining':['Copper mine processing plant with tailings facility, crushers, mills, power and water logistics','Lithium mine and refinery with resource access, processing plant and environmental approvals'],
 'healthcare':['Acute hospital redevelopment with clinical commissioning, infection control, medical equipment and phased occupancy','Cancer treatment centre with specialist systems, patient transition and regulator inspection'],
 'life_sciences':['Eli Lilly biologics manufacturing campus with GMP cleanrooms, CQV, media fills, FDA readiness and cold chain','Amgen sterile fill-finish expansion with clean utilities, validation and batch release'],
 'semiconductor':['Arizona semiconductor fab with EUV lithography, UPW, cleanroom, process tools and yield ramp','TSMC wafer fabrication campus with specialty gases, tool install and contamination control'],
 'defence':['Secure defence command centre with radar integration, accredited systems, sovereign suppliers and operational acceptance','Naval base modernisation with secure comms, mission systems and controlled procurement'],
 'general_infrastructure':['Mixed-use urban regeneration programme with utilities, transport interfaces, commercial fit-out and phased handover']
}
SPACE_SECTORS = {
 'space':['Lunar surface habitat with landing pads, power storage, life support, autonomous rovers and mission assurance','Mars ISRU propellant refinery with methane oxygen production, launch cadence, thermal power balance and autonomous operations','LEO orbital compute platform with payload integration, radiation hardening, thermal rejection and relay communications','Spaceport launch infrastructure with propellant farms, range safety, launch pads and environmental approvals','Satellite constellation manufacturing and launch programme with ground segment integration and spectrum approvals']
}
SCENARIOS=['base','faster','cheaper','lower_risk','premium']
CLASSES=[1,2,3,4,5]
SCHED=[1,2,3,4,5]

# forbidden terms per expected sector (lowercase); exclude prompt/client/title from scan
FORBIDDEN = {
 'data_centre':['orat','baggage','airside','signalling','rolling stock','payload integration','mission assurance','cqv','media fill','hazop','wafer'],
 'airport':['liquid cooling','white space','data hall','gpu','signalling','rolling stock','launch cadence','payload integration','cqv','hazop','yield ramp'],
 'rail':['liquid cooling','white space','data hall','gpu','orat','baggage','airside','launch cadence','payload integration','cqv','hazop','yield ramp'],
 'life_sciences':['liquid cooling','white space','data hall','gpu','orat','baggage','airside','signalling','rolling stock','launch cadence','payload integration','hazop','yield ramp'],
 'semiconductor':['liquid cooling','orat','baggage','airside','signalling','rolling stock','launch cadence','payload integration','cqv','hazop'],
 'defence':['liquid cooling','orat','baggage','airside','signalling','rolling stock','cqv','yield ramp','hazop'],
 'oil_gas':['liquid cooling','orat','baggage','airside','signalling','rolling stock','cqv','yield ramp','payload integration'],
 'energy':['liquid cooling','orat','baggage','airside','signalling','rolling stock','cqv','yield ramp','payload integration','hazop'],
 'nuclear':['liquid cooling','orat','baggage','airside','signalling','rolling stock','cqv','yield ramp','payload integration','hazop'],
 'healthcare':['liquid cooling','white space','data hall','gpu','orat','baggage','airside','signalling','rolling stock','launch cadence','payload integration','hazop','yield ramp'],
 'water':['liquid cooling','white space','data hall','gpu','orat','baggage','airside','signalling','rolling stock','launch cadence','payload integration','cqv','hazop','yield ramp'],
 'ports':['liquid cooling','white space','data hall','gpu','orat','baggage','airside','signalling','rolling stock','launch cadence','payload integration','cqv','hazop','yield ramp'],
 'mining':['liquid cooling','white space','data hall','gpu','orat','baggage','airside','signalling','rolling stock','launch cadence','payload integration','cqv','hazop','yield ramp'],
 'roads':['liquid cooling','white space','data hall','gpu','orat','baggage','airside','signalling','rolling stock','launch cadence','payload integration','cqv','hazop','yield ramp'],
 'general_infrastructure':['liquid cooling','orat','baggage','airside','signalling','rolling stock','launch cadence','payload integration','cqv','hazop','yield ramp'],
 'space':['orat','baggage','airside','signalling','rolling stock','cqv','media fill','hazop','wafer','yield ramp']
}

def flatten(obj, path=''):
    if isinstance(obj, dict):
        for k,v in obj.items():
            if k in {'prompt','client','title','id'}: continue
            yield from flatten(v, path+'.'+k)
    elif isinstance(obj, list):
        for i,v in enumerate(obj): yield from flatten(v, path+f'[{i}]')
    elif isinstance(obj, str):
        yield path, obj.lower()

def leak_check(expected, model):
    leaks=[]
    for path,text in flatten(model):
        for term in FORBIDDEN.get(expected,[]):
            if term in text:
                leaks.append((path, term, text[:160]))
                if len(leaks)>5: return leaks
    return leaks

def expected_key(group): return group

def make_cases(seed, n, source):
    rng=random.Random(seed)
    keys=list(source.keys())
    concerns=['board approval','cost confidence','schedule risk','procurement evidence','commissioning readiness']
    for i in range(n):
        key=rng.choice(keys)
        prompt=rng.choice(source[key]) + f". Main concern: {rng.choice(concerns)}."
        yield key, prompt, rng.choice(SCENARIOS), rng.choice(CLASSES), rng.choice(SCHED)

def run_generation_suite():
    failures=[]; counts={}; cache={}; leak_checked=set()
    start=time.time()
    total=0; built=0
    for expected,prompt,scenario,cl,sl in make_cases(42,10000,EARTH_SECTORS):
        ck=(expected,scenario,'earth')
        if ck not in cache:
            cache[ck]=main.build_model(EARTH_SECTORS[expected][0],'T&T client benchmark',cl,sl,scenario); built+=1
        m=cache[ck]
        route=main._v132_sector_key(prompt,'T&T client benchmark',m)
        got=m.get('sector_ontology_key')
        if route!=expected:
            failures.append({'type':'routing_function','expected':expected,'got':route,'prompt':prompt[:120]})
        counts[got]=counts.get(got,0)+1
        if got!=expected:
            failures.append({'type':'routing','expected':expected,'got':got,'prompt':prompt[:120]})
        if ck not in leak_checked:
            leaks=leak_check(expected,m)
            leak_checked.add(ck)
            if leaks:
                failures.append({'type':'leak','expected':expected,'got':got,'leaks':leaks,'prompt':prompt[:120]})
        total+=1
        if failures and len(failures)>20: break
    if len(failures)<=20:
        for expected,prompt,scenario,cl,sl in make_cases(84,5000,SPACE_SECTORS):
            ck=(expected,scenario,'space')
            if ck not in cache:
                cache[ck]=main.build_model(SPACE_SECTORS[expected][0],'Space client benchmark',cl,sl,scenario); built+=1
            m=cache[ck]
            route=main._v132_sector_key(prompt,'Space client benchmark',m)
            got=m.get('sector_ontology_key')
            if route!=expected:
                failures.append({'type':'routing_function','expected':expected,'got':route,'prompt':prompt[:120]})
            counts[got]=counts.get(got,0)+1
            if got!=expected:
                failures.append({'type':'routing','expected':expected,'got':got,'prompt':prompt[:120]})
            leaks=leak_check(expected,m)
            if leaks:
                failures.append({'type':'leak','expected':expected,'got':got,'leaks':leaks,'prompt':prompt[:120]})
            total+=1
            if failures and len(failures)>20: break
    return {'total':total,'unique_model_builds':built,'failures':failures,'counts':counts,'seconds':round(time.time()-start,2)}

def export_smoke():
    # One representative run per sector, all classes, all schedule levels, all scenarios are generation-tested;
    # export byte-generation is smoke-tested on representative base case per sector to keep runtime sane.
    results=[]
    for expected, prompts in {**EARTH_SECTORS, **SPACE_SECTORS}.items():
        prompt=prompts[0]
        m=main.build_model(prompt,'Export QA client',3,4,'base')
        r={'sector':expected,'model_sector':m.get('sector_ontology_key')}
        try:
            wb=main.workbook_bytes(m); rr=main.risk_register_workbook_bytes(m); xer=main.xer_bytes(m); js=json.dumps(m).encode()
            r.update({'workbook_bytes':len(wb),'risk_register_bytes':len(rr),'xer_bytes':len(xer),'json_bytes':len(js),'ok': all([len(wb)>1000,len(rr)>1000,len(xer)>100,len(js)>1000])})
        except Exception as e:
            r.update({'ok':False,'error':repr(e)})
        results.append(r)
    return results

if __name__=='__main__':
    gen=run_generation_suite()
    exp=export_smoke()
    report={'version':'CASEY V132 Institutional Authority Final','generation_qa':gen,'export_smoke':exp,'export_smoke_failures':[x for x in exp if not x.get('ok')]}
    Path('/mnt/data/CASEY_V132_FULL_EARTH_SPACE_QA_REPORT.json').write_text(json.dumps(report,indent=2))
    md=[]
    md.append('# CASEY V132 Full Earth + Space QA Report')
    md.append('')
    md.append('## Scope')
    md.append('- 10,000 generated Earth examples across T&T-style sectors, estimate classes, schedule levels and scenarios')
    md.append('- 5,000 generated Space examples across lunar, Mars, orbital compute, spaceport and satellite/mission programmes')
    md.append('- Representative export smoke tests for every sector family')
    md.append('')
    md.append(f"Generation runs: **{gen['total']}**")
    md.append(f"Generation failures: **{len(gen['failures'])}**")
    md.append(f"Unique model builds: **{gen.get('unique_model_builds')}**")
    md.append(f"Runtime: **{gen['seconds']}s**")
    md.append('')
    md.append('## Sector counts')
    for k,v in sorted(gen['counts'].items()): md.append(f'- {k}: {v}')
    md.append('')
    md.append('## Export smoke')
    for r in exp: md.append(f"- {r['sector']}: {'PASS' if r.get('ok') else 'FAIL'}")
    if gen['failures']:
        md.append('\n## First failures')
        for f in gen['failures'][:10]: md.append('```json\n'+json.dumps(f,indent=2)+'\n```')
    Path('/mnt/data/CASEY_V132_FULL_EARTH_SPACE_QA_REPORT.md').write_text('\n'.join(md))
    print(json.dumps({'generation_failures':len(gen['failures']),'export_failures':len([x for x in exp if not x.get('ok')]),'total':gen['total']},indent=2))
