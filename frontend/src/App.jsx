import React, { useEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { motion, AnimatePresence } from 'framer-motion';
import {
  Activity, AlertTriangle, ArrowRight, BarChart3, Brain, BriefcaseBusiness, Building2, ChevronRight,
  Database, Download, FileSpreadsheet, FileText, Globe2, Mail, Moon, Orbit, Play, Rocket,
  ShieldAlert, Sparkles, Upload, Workflow, Zap
} from 'lucide-react';
import {
  Area, AreaChart, Bar, BarChart, CartesianGrid, Line, LineChart, ResponsiveContainer, Tooltip, XAxis, YAxis
} from 'recharts';
import './style.css';

const API = import.meta.env.VITE_API_URL || 'http://127.0.0.1:8000';

const earthPrompt = 'Riyadh AI Hyperscale Campus 500MW accelerated 2027 with sovereign cloud, grid connection and liquid cooling';
const spacePrompt = 'Lunar Base Alpha with 1000 crew, nuclear power, landing pads, life support and launch logistics';
const examples = [
  earthPrompt,
  'Modern airport expansion with terminal, runway, baggage systems and live operations',
  'Automated port expansion with cranes, logistics yards, customs systems and rail connection',
  'Boston GMP life sciences manufacturing campus with cleanrooms and validation',
  'Arizona advanced semiconductor fab with cleanroom and utility complexity',
  'Mars Fuel Refinery with ISRU methane oxygen production',
  spacePrompt,
];
const scenarios = ['base', 'faster', 'cheaper', 'lower_risk', 'premium'];

function fmt(v) {
  if (v === undefined || v === null || v === '') return '—';
  if (typeof v === 'string') return v;
  return v >= 1000 ? `$${(v / 1000).toFixed(1)}T` : v >= 1 ? `$${v.toFixed(1)}B` : `$${(v * 1000).toFixed(0)}M`;
}
async function post(path, body) {
  const r = await fetch(API + path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  if (!r.ok) throw new Error(await r.text());
  return r.json();
}
async function get(path) {
  const r = await fetch(API + path);
  if (!r.ok) throw new Error(await r.text());
  return r.json();
}
async function download(path, model, name) {
  const r = await fetch(API + path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(model) });
  if (!r.ok) throw new Error(await r.text());
  const blob = await r.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}
function Card({ children, className = '' }) {
  return <motion.div className={`card ${className}`} initial={{ opacity: 0, y: 18 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: 0.35 }}>{children}</motion.div>;
}
function Logo({ large = false }) {
  return <div className={large ? 'v50Logo large' : 'v50Logo'}><img src="/brand/casey_wordmark.png" alt="CASEY" /><span>Mission Control for Capital Projects</span></div>;
}
function Kpi({ icon: Icon, label, value, sub, hot }) {
  return <Card className={`v50Kpi ${hot ? 'hot' : ''}`}><Icon size={21}/><div><p>{label}</p><b>{value}</b><span>{sub}</span></div></Card>;
}
function Table({ rows = [], cols = [], moneyCols = [] }) {
  return <div className="tableWrap"><table><thead><tr>{cols.map(c => <th key={c[0]}>{c[1]}</th>)}</tr></thead><tbody>{rows.map((r, i) => <tr key={i}>{cols.map(c => <td key={c[0]}>{moneyCols.includes(c[0]) ? fmt(r[c[0]]) : String(r[c[0]] ?? '')}</td>)}</tr>)}</tbody></table></div>;
}
function Hero({ onBriefing, onEarth, onSpace, onConsole }) {
  return <section className="v50TakeoverHero">
    <video className="v50HeroVideo" src="https://corbit.b-cdn.net/casey_hero_film.mp4" autoPlay muted loop playsInline preload="auto" />
    <div className="v50HeroShade" />
    <div className="v50TopBar"><Logo/><div className="v50TopActions"><button onClick={onBriefing}><Play size={15}/> Watch briefing</button><button onClick={onEarth}>Run Earth model</button><button onClick={onSpace}>Run Space model</button><button onClick={onConsole}>Open console</button><a className="topBuyLink" href="mailto:hello@casey.ai?subject=CASEY%20Access%20Request">Request access</a></div></div>
    <motion.div className="v50HeroCenter" initial={{ opacity: 0, y: 14 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: .7 }}>
      <Logo large />
      <p className="v50HeroLine">Cost · Schedule · Risk · Delivery</p>
      <h1>Price the future before it gets built.</h1>
      <p className="v50HeroSub">First-pass project controls intelligence for Earth infrastructure and orbital programmes.</p>
      <div className="v50HeroButtons"><button className="heroBtn" onClick={onEarth}><Rocket size={18}/> Run this project</button><button className="ghostBtn" onClick={onBriefing}><Play size={18}/> Play film</button></div>
    </motion.div>
    <div className="v50BottomBar"><span>AI data centres</span><span>Airports</span><span>Ports</span><span>Life sciences</span><span>Semiconductors</span><span>Lunar bases</span><button onClick={onEarth}>Generate full pack</button><button onClick={onConsole}>View pricing</button></div>
  </section>;
}
function Briefing({ open, onClose, onEarth, onSpace }) {
  if (!open) return null;
  return <AnimatePresence><motion.div className="v50Briefing" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}>
    <video className="v50BriefingVideo" src="https://corbit.b-cdn.net/casey_hero_film.mp4" autoPlay controls playsInline preload="auto" />
    <div className="v50BriefingTop"><Logo/><button onClick={onClose}>Exit film</button></div>
    <div className="v50BriefingBottom"><button onClick={onEarth}>Run Earth model</button><button onClick={onSpace}>Run Space model</button><button onClick={onClose}>Open product</button></div>
  </motion.div></AnimatePresence>;
}
function Loading({ text }) {
  return <motion.div className="loading" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}><Rocket size={44}/><b>{text || 'Building connected model...'}</b><span>Cost · Schedule · QCRA · QSRA · Risk Register · Board Pack</span></motion.div>;
}
function App() {
  const [show, setShow] = useState(true);
  const [briefing, setBriefing] = useState(false);
  const [prompt, setPrompt] = useState(earthPrompt);
  const [client, setClient] = useState('Client / operator');
  const [scenario, setScenario] = useState('base');
  const [classLevel, setClassLevel] = useState(3);
  const [scheduleLevel, setScheduleLevel] = useState(4);
  const [model, setModel] = useState(null);
  const [tab, setTab] = useState('overview');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [chatQ, setChatQ] = useState('');
  const [chat, setChat] = useState([{ role: 'assistant', text: 'Ask CASEY why cost, schedule, contingency or risk confidence is moving.' }]);
  const [uploadResult, setUploadResult] = useState(null);

  useEffect(() => { get('/v26/demo-library').catch(() => null); }, []);
  async function generate(nextScenario = scenario, nextPrompt = prompt) {
    setError(''); setLoading(true); setTab('overview'); setShow(false);
    try {
      const m = await post('/generate', { prompt: nextPrompt, client, class_level: Number(classLevel), schedule_level: Number(scheduleLevel), scenario: nextScenario, demo: true });
      setModel(m); setScenario(nextScenario); setPrompt(nextPrompt);
    } catch (e) { setError(String(e.message || e)); }
    finally { setLoading(false); }
  }
  function runEarth() { generate('base', earthPrompt); }
  function runSpace() { generate('base', spacePrompt); }
  async function ask() {
    if (!chatQ.trim() || !model) return;
    const q = chatQ; setChatQ(''); setChat(x => [...x, { role: 'user', text: q }]);
    try { const r = await post('/chat', { question: q, project: model, demo: true }); setChat(x => [...x, { role: 'assistant', text: r.answer || JSON.stringify(r) }]); }
    catch (e) { setChat(x => [...x, { role: 'assistant', text: String(e.message || e) }]); }
  }
  async function upload(e) {
    const f = e.target.files?.[0]; if (!f) return;
    const fd = new FormData(); fd.append('file', f);
    const r = await fetch(API + '/upload', { method: 'POST', body: fd });
    setUploadResult(await r.json());
  }
  const costs = model?.cost_breakdown || [];
  const risks = model?.risk_register || [];
  const schedule = model?.schedule_detail || [];
  const curve = model?.monte_carlo?.curve || [];
  const tornado = model?.monte_carlo?.tornado || [];
  const direct = costs.filter(x => String(x.type || '').toLowerCase().includes('direct')).reduce((a, b) => a + Number(b.p50_bn || 0), 0);
  const indirect = costs.filter(x => String(x.type || '').toLowerCase().includes('indirect')).reduce((a, b) => a + Number(b.p50_bn || 0), 0);
  const reserves = costs.filter(x => String(x.type || '').toLowerCase().includes('reserve')).reduce((a, b) => a + Number(b.p50_bn || 0), 0);
  const emailBody = model ? [
    'Please review this project in CASEY.', '', `Project: ${model.title}`, `Scenario: ${model.scenario_label || scenario}`,
    `P50 Cost: ${model.cost_p50}`, `Cost Range: ${model.cost_range}`, `Schedule: ${model.schedule}`,
    `Risk / Confidence: ${model.risk} / ${model.confidence_pct}%`
  ].join('\n') : 'Please send me CASEY access.';
  const emailLink = `mailto:hello@casey.ai?subject=${encodeURIComponent('CASEY project review')}&body=${encodeURIComponent(emailBody)}`;

  return <div className="app v50EliteApp">
    <Briefing open={briefing} onClose={() => setBriefing(false)} onEarth={runEarth} onSpace={runSpace}/>
    <AnimatePresence>{loading && <Loading text="Building full CASEY intelligence pack..."/>}</AnimatePresence>
    {show && !model && <Hero onBriefing={() => setBriefing(true)} onEarth={runEarth} onSpace={runSpace} onConsole={() => setShow(false)}/>} 
    <header className="v50ConsoleTop"><Logo/><nav><button onClick={() => { setModel(null); setShow(true); }}>Home</button><button onClick={() => setBriefing(true)}>Film</button><button onClick={runEarth}>Earth demo</button><button onClick={runSpace}>Space demo</button><a href={emailLink}>Request access</a></nav></header>
    <main className={model ? 'v50Console' : 'v50Console emptyConsole'}>
      {error && <div className="error">{error}</div>}
      {!model && !show && <section className="commandGrid"><Card className="command"><h1>Generate a live project model</h1><label>Project command</label><textarea value={prompt} onChange={e => setPrompt(e.target.value)} /> <div className="chips">{examples.map(x => <button key={x} onClick={() => setPrompt(x)}>{x}</button>)}</div><div className="grid4"><input value={client} onChange={e => setClient(e.target.value)} placeholder="Client / operator"/><select value={classLevel} onChange={e => setClassLevel(e.target.value)}>{[1,2,3,4,5].map(x => <option key={x} value={x}>Class {x}</option>)}</select><select value={scheduleLevel} onChange={e => setScheduleLevel(e.target.value)}>{[1,2,3,4,5].map(x => <option key={x} value={x}>Level {x}</option>)}</select><select value={scenario} onChange={e => setScenario(e.target.value)}>{scenarios.map(x => <option key={x} value={x}>{x}</option>)}</select></div><button className="primary" onClick={() => generate()}><Sparkles/> Generate full intelligence pack</button></Card><Card><h2>What CASEY will produce</h2>{['Executive summary and recommendation','Direct / indirect / reserve cost view','Scenario-linked estimate, schedule and confidence','Risk register with cause, event, impact and mitigation','QCRA + QSRA curves and tornado drivers','Pricing and next-step contact actions'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card></section>}
      {model && <>
        <section className="kpis"><Kpi icon={Globe2} label="Mode / sector" value={model.mode} sub={model.subsector}/><Kpi icon={Activity} label="P50 cost" value={model.cost_p50} sub={model.cost_range}/><Kpi icon={Zap} label="Schedule" value={model.schedule} sub={`QSRA P80 ${model.monte_carlo?.qsra?.p80 || '—'} months`}/><Kpi icon={ShieldAlert} label="Risk / Confidence" value={`${model.risk} / ${model.confidence_pct}%`} sub={model.scenario_label} hot/></section>
        <nav className="tabs">{[['overview','Overview'],['compare','Scenarios'],['cost','Cost'],['schedule','Schedule'],['risk','Risk'],['monte','QCRA/QSRA'],['outputs','Outputs'],['advisor','Advisor'],['method','Methodology'],['pricing','Pricing']].map(x => <button key={x[0]} className={tab===x[0]?'active':''} onClick={() => setTab(x[0])}>{x[1]}</button>)}</nav>
        {tab === 'overview' && <section className="layout two"><Card><h2>Executive intelligence summary</h2><p className="big">{model.executive_summary || `${model.title} has been classified as ${model.subsector}. CASEY generated a first-pass cost, schedule, risk and confidence model for the selected scenario.`}</p><div className="miniMetrics"><b><span>Direct cost</span>{fmt(direct)}</b><b><span>Indirect cost</span>{fmt(indirect)}</b><b><span>Risk / reserve</span>{fmt(reserves)}</b></div><h3>Recommendation</h3>{(model.next_best_actions || []).slice(0,5).map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card><Card><h2>Board challenge questions</h2>{(model.board_challenge_questions || []).slice(0,7).map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}<h3>Scenario switch</h3><div className="scenarioGrid mini">{scenarios.map(s => <button key={s} onClick={() => generate(s)}>{s.replace('_',' ')}</button>)}</div></Card></section>}

        {tab === 'compare' && <section className="layout two"><Card><h2>Scenario comparison</h2><p className="big">Switch options before paying for another advisory cycle. CASEY keeps cost, schedule, risk and confidence aligned across every export.</p><div className="scenarioCompare">{scenarios.map(s => { const active = s === scenario; return <button key={s} className={active?'active':''} onClick={() => generate(s)}><b>{s.replace('_',' ')}</b><span>{active ? 'current model' : 'run scenario'}</span></button> })}</div></Card><Card><h2>Buyer decision lens</h2>{['Base: balanced reference case for board challenge','Faster: premium acceleration with higher interface exposure','Cheaper: reduced cost target with increased delivery risk','Lower Risk: slower, more assured delivery with higher reserves','Premium: flagship resilience, higher capex and stronger confidence'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card></section>}
        {tab === 'cost' && <section className="layout two"><Card><h2>Premium cost estimate</h2><p>Direct, indirect and reserve categories are separated. The Excel export should be treated as the workbook source of truth.</p><Table rows={costs} cols={[['cbs','CBS'],['description','Description'],['type','Type'],['p10_bn','Low/P10'],['p50_bn','Most likely/P50'],['p90_bn','High/P90'],['impact_basis','Basis']]} moneyCols={['p10_bn','p50_bn','p90_bn']}/></Card><Card><h2>Cost composition</h2><ResponsiveContainer width="100%" height={320}><BarChart data={[{name:'Direct',value:direct},{name:'Indirect',value:indirect},{name:'Reserve',value:reserves}]}><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis dataKey="name"/><YAxis/><Tooltip/><Bar dataKey="value" fill="#8df7ff"/></BarChart></ResponsiveContainer></Card></section>}
        {tab === 'schedule' && <section className="layout two"><Card><h2>Schedule logic</h2><Table rows={schedule} cols={[['activity_id','Activity'],['phase','Phase'],['activity','Name'],['predecessor','Pred'],['duration_months','Months'],['critical','Critical'],['basis','Basis']]}/></Card><Card><h2>Schedule confidence</h2><ResponsiveContainer width="100%" height={320}><LineChart data={curve}><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis dataKey="percentile"/><YAxis/><Tooltip/><Line type="monotone" dataKey="schedule_months" stroke="#b18cff" strokeWidth={3}/></LineChart></ResponsiveContainer></Card></section>}
        {tab === 'risk' && <section className="layout two"><Card><h2>Risk Register Pro</h2><p>Risk output should include cause, event, impact, owner, mitigation and links to WBS/CBS.</p><Table rows={risks} cols={[['risk_id','ID'],['title','Risk'],['cause','Cause'],['event','Event'],['impact','Impact'],['probability_pct','Prob %'],['activity_id','Activity'],['cbs','CBS'],['owner','Owner'],['mitigation','Mitigation']]}/></Card><Card><h2>Top exposure drivers</h2><ResponsiveContainer width="100%" height={380}><BarChart data={tornado} layout="vertical"><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis type="number"/><YAxis dataKey="driver" type="category" width={150}/><Tooltip/><Bar dataKey="contribution" fill="#8df7ff"/></BarChart></ResponsiveContainer></Card></section>}
        {tab === 'monte' && <section className="layout two"><Card><h2>QCRA cost curve</h2><ResponsiveContainer width="100%" height={330}><AreaChart data={curve}><defs><linearGradient id="caseyG" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stopColor="#8df7ff" stopOpacity=".55"/><stop offset="1" stopColor="#8df7ff" stopOpacity="0"/></linearGradient></defs><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis dataKey="percentile"/><YAxis/><Tooltip/><Area type="monotone" dataKey="cost_bn" stroke="#8df7ff" fill="url(#caseyG)"/></AreaChart></ResponsiveContainer></Card><Card><h2>QSRA schedule curve</h2><ResponsiveContainer width="100%" height={330}><LineChart data={curve}><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis dataKey="percentile"/><YAxis/><Tooltip/><Line type="monotone" dataKey="schedule_months" stroke="#b18cff" strokeWidth={3}/></LineChart></ResponsiveContainer></Card></section>}
        {tab === 'outputs' && <section className="layout two"><Card><h2>Executive Control Pack</h2><p>Generate a board-ready project controls pack: decision centre, cost model, risk register, PRA schedule, QCRA/QSRA analysis and audit file — all linked to the selected project, scenario, estimate class and schedule level.</p><div className="exports v50Exports"><button onClick={() => download('/export/workbook', model, 'CASEY_FINAL_Cost_Model.xlsx')}><FileSpreadsheet/> Cost Model XLSX</button><button onClick={() => download('/export/risk-register', model, 'CASEY_FINAL_Risk_Register.xlsx')}><Database/> Risk Register XLSX</button><button onClick={() => download('/export/xer', model, 'CASEY_FINAL_Schedule.xer')}><Workflow/> PRA Schedule XER</button><button onClick={() => download('/export/schedule-csv', model, 'CASEY_FINAL_Schedule_Levels.csv')}><Workflow/> Schedule Levels CSV</button><button onClick={() => download('/export/json', model, 'CASEY_FINAL_Model_Audit.json')}><Brain/> Audit File JSON</button><button onClick={() => download('/export/all', model, 'CASEY_FINAL_Output_Pack.zip')}><Download/> Full Pack ZIP</button><a className="contactBtn" href={emailLink}><Mail/> Email this project</a></div></Card><Card><h2>What the pack delivers</h2>{['Executive control centre with project, scenario, class, level and confidence clearly identified','Scenario comparison covering Base, Faster, Cheaper, Lower Risk, Premium, Investor and Survival cases','Selected estimate class plus all class levels for audit and challenge','Direct, indirect and reserve cost views with QCRA cost curve and cost tornado','All schedule levels with QSRA schedule curve and schedule tornado','Risk register with cause, event, impact, owner, mitigation, trigger and quantified likelihood','Basis of Estimate, assumptions, exclusions and benchmark validation','Commercial next steps: buyer action, procurement challenge and board decision path'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card></section>}

        {tab === 'advisor' && <section className="layout two"><Card><h2>Ask CASEY</h2><div className="chatBox">{chat.map((m,i)=><div key={i} className={`msg ${m.role}`}>{m.text}</div>)}</div><div className="ask"><input value={chatQ} onChange={e=>setChatQ(e.target.value)} onKeyDown={e=>{if(e.key==='Enter')ask()}} placeholder="Ask why the cost or risk moved..."/><button onClick={ask}>Ask</button></div></Card><Card><h2>Upload estimate challenge</h2><p>Use this to show buyers how CASEY can challenge a Tier 1 estimate.</p><label className="upload"><Upload size={18}/> Upload file<input type="file" onChange={upload}/></label><button className="secondary" onClick={()=>setUploadResult({review:'Sample contractor estimate challenge', findings:['Direct costs above benchmark in power train and cooling package','Indirects and preliminaries need clearer split from reserves','Schedule contingency understated against critical path risks','Risk allowance should separate QCRA cost and QSRA schedule exposure'], next_action:'Request rate build-up, supplier quotes, basis of estimate and revised risk register.'})}>Run sample challenge</button>{uploadResult && <pre>{JSON.stringify(uploadResult,null,2)}</pre>}</Card></section>}

        {tab === 'method' && <section className="layout two"><Card><h2>How CASEY calculated this</h2>{['Cost model: selected class estimate, sector template, location factor, complexity factor and scenario modifier.','Schedule model: level-based delivery logic, phase durations, critical path sensitivity and scenario acceleration/delay factors.','QCRA: cost exposure model using low / most likely / high impacts and risk-weighted contingency.','QSRA: schedule exposure model using activity-linked O/M/P delay ranges and critical path sensitivity.','Confidence score: class maturity, schedule detail, scenario risk profile and location/space complexity.'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card><Card><h2>Commercial readiness</h2><p className="big">This is first-pass project controls intelligence. It is designed to accelerate challenge, option testing and board preparation before final contractor tender or signed cost plan.</p><a className="contactBtn huge" href={emailLink}><Mail/> Send project for review</a></Card></section>}
        {tab === 'pricing' && <section className="layout two"><Card><h2>CASEY Access</h2><div className="pricingGrid"><div className="priceCard"><b>Pilot</b><strong>Request pricing</strong><span>Guided project review, sample outputs and executive walkthrough.</span><a href={emailLink}>Request pilot</a></div><div className="priceCard hot"><b>Professional</b><strong>Full project pack</strong><span>Cost, schedule, risk, QCRA/QSRA and export pack.</span><a href={emailLink}>Request access</a></div><div className="priceCard"><b>Enterprise</b><strong>Private deployment</strong><span>SSO, teams, benchmark library, private models and audit trail.</span><a href={emailLink}>Book demo</a></div></div></Card><Card><h2>Send this project</h2><p className="big">Turn demo interest into pipeline immediately.</p><a className="contactBtn huge" href={emailLink}><Mail/> Send project to CASEY</a><button className="primary" onClick={() => download('/export/all', model, 'CASEY_Output_Pack.zip')}>Download full pack</button></Card></section>}
      </>}
    </main>
  </div>;
}

createRoot(document.getElementById('root')).render(<App/>);
