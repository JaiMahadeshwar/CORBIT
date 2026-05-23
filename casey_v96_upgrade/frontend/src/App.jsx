import React, { useEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { motion, AnimatePresence } from 'framer-motion';
import {
  Activity, AlertTriangle, ArrowRight, BarChart3, Brain, BriefcaseBusiness, Building2, ChevronRight,
  Database, Download, FileSpreadsheet, FileText, Globe2, Mail, Moon, Orbit, Play, Rocket,
  ShieldAlert, Sparkles, Upload, Workflow, Zap
} from 'lucide-react';
import {
  Area, AreaChart, Bar, BarChart, CartesianGrid, Line, LineChart, ResponsiveContainer, Tooltip, XAxis, YAxis, ReferenceLine
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
function Hero({ onBriefing, onEarth, onSpace, onConsole, onTryDemo }) {
  return <section className="v50TakeoverHero">
    <video className="v50HeroVideo" src="https://corbit.b-cdn.net/casey_hero_film.mp4" autoPlay muted loop playsInline preload="auto" />
    <div className="v50HeroShade" />
    <div className="v50TopBar"><Logo/><div className="v50TopActions"><button onClick={onBriefing}><Play size={15}/> Watch briefing</button><button onClick={onEarth}>Run Earth model</button><button onClick={onSpace}>Run Space model</button><button className="tryTopBtn" onClick={onTryDemo}>Try one free run</button><button onClick={onConsole}>Open console</button><a className="topBuyLink" href="mailto:hello@casey.ai?subject=CASEY%20Access%20Request">Request access</a></div></div>
    <motion.div className="v50HeroCenter" initial={{ opacity: 0, y: 14 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: .7 }}>
      <Logo large />
      <p className="v50HeroLine">Cost · Schedule · Risk · Delivery</p>
      <h1>Price the future before it gets built.</h1>
      <p className="v50HeroSub">Run one free CASEY intelligence pack. Enter your email, describe one Earth or Space project, and receive a first-pass class estimate, schedule view and risk register.</p>
      <div className="v50HeroButtons"><button className="heroBtn" onClick={onTryDemo}><Rocket size={18}/> Run One Free Intelligence Pack</button><button className="ghostBtn" onClick={onBriefing}><Play size={18}/> Play film</button></div>
    </motion.div>
    <div className="v50BottomBar"><span>AI data centres</span><span>Airports</span><span>Ports</span><span>Life sciences</span><span>Semiconductors</span><span>Lunar bases</span><button onClick={onTryDemo}>One free intelligence run</button><button onClick={onConsole}>View pricing</button></div>
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

function OneShotDemo({ open, onClose, onComplete }) {
  const [form, setForm] = useState({
    email: '',
    project_type: 'Earth',
    project_description: '',
    location: 'Auto-inferred from project brief',
    size_or_capacity: 'Auto-inferred from project brief',
    stage: 'Concept / early feasibility',
    biggest_concern: 'Cost, schedule and risk confidence'
  });
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [result, setResult] = useState(null);
  if (!open) return null;

  function update(k, v) { setForm(x => ({ ...x, [k]: v })); }

  function inferType(text) {
    const t = String(text || '').toLowerCase();
    const earthStrong = /(north carolina|carolina|cambridge|boston|arizona|texas|uk|usa|united states|riyadh|dubai|qatar|gmp|aseptic|fill-finish|fill finish|fda|cqv|biologics|therapeutics|pharma|pharmaceutical|cold-chain|clean utilities|semiconductor|data centre|data center|airport|rail|hospital|nuclear|hydrogen|desalination|defence|defense)/.test(t);
    const spaceStrong = /(moon|lunar|mars|orbit|orbital|leo|cislunar|cis-lunar|spaceport|launch vehicle|launch pad|rocket|satellite constellation|space data|orbital ai|deep space|propellant depot|asteroid)/.test(t);
    const productLaunchOnly = /(commercial launch demand|product launch|market launch|launch demand)/.test(t) && !/(rocket|spaceport|launch vehicle|launch pad|orbital|leo|lunar|mars)/.test(t);
    if (productLaunchOnly || (earthStrong && !spaceStrong)) return 'Earth';
    return spaceStrong ? 'Space' : 'Earth';
  }
  function inferLocation(text) {
    const t = String(text || '').toLowerCase();
    if (/leo|low earth orbit/.test(t)) return 'LEO';
    if (/moon|lunar/.test(t)) return 'Lunar surface';
    if (/mars/.test(t)) return 'Mars';
    if (/cislunar|cis-lunar/.test(t)) return 'Cislunar space';
    if (/riyadh|saudi/.test(t)) return 'Riyadh, Saudi Arabia';
    if (/dubai|uae|abu dhabi/.test(t)) return 'UAE';
    if (/texas|usa|america|arizona/.test(t)) return 'United States';
    if (/uk|london|manchester/.test(t)) return 'United Kingdom';
    return 'Auto-inferred from project brief';
  }
  function inferCapacity(text) {
    const m = String(text || '').match(/(\d+[\s-]?(mw|gw|km|beds|satellites|crew|m2|sqm|terminals|stations|halls|modules))/i);
    return m ? m[1] : 'Auto-inferred from project brief';
  }
  function inferConcern(text) {
    const t = String(text || '').toLowerCase();
    if (/thermal|radiation|launch|orbit|servicing|relay|latency/.test(t)) return 'Space logistics / orbital dependency';
    if (/grid|power|cooling|energy/.test(t)) return 'Power, utilities and commissioning risk';
    if (/cost|budget|estimate/.test(t)) return 'Cost estimate challenge';
    if (/schedule|critical path|delay/.test(t)) return 'Schedule risk / critical path';
    if (/risk|approval|regulatory|consent/.test(t)) return 'Risk register quality';
    return 'Cost, schedule and risk confidence';
  }
  function clientToken() {
    let t = localStorage.getItem('casey_public_demo_token');
    if (!t) { t = crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random(); localStorage.setItem('casey_public_demo_token', t); }
    return t;
  }
  function fingerprint() {
    return [navigator.userAgent, navigator.language, Intl.DateTimeFormat().resolvedOptions().timeZone, screen.width + 'x' + screen.height, screen.colorDepth].join('|');
  }

  function normalisedWords(text) {
    return String(text || '').replace(/\s+/g, ' ').trim().split(' ').filter(Boolean);
  }
  function briefSignals(text) {
    const t = String(text || '').toLowerCase();
    const asset = /(data centre|data center|datacenter|airport|rail|metro|hospital|fab|semiconductor|nuclear|smr|hydrogen|grid|desalination|defence|defense|campus|plant|facility|biologics|therapeutics|gmp|aseptic|fill-finish|fda|cqv|pharma|manufacturing|orbital|leo|lunar|moon|mars|spaceport|satellite|propellant|habitat|thermal rejection|relay communications)/.test(t);
    const place = /(north carolina|carolina|arizona|texas|boston|london|cambridge|manchester|riyadh|dubai|uae|uk|usa|united states|saudi|moon|lunar|mars|leo|orbit|orbital|cislunar|spaceport)/.test(t);
    const concern = /(cost|schedule|risk|procurement|approval|commissioning|regulatory|interface|utilities|critical path|supply chain|phasing|validation|qualification|resilience|thermal|radiation|servicing|power|grid|fda|cqv|inspection|continuity|launch cadence|long-lead|long lead)/.test(t);
    const nonsense = /(asdf|qwerty|lorem ipsum|blah blah|hello world|ignore previous|jailbreak|write a poem|recipe|football match|dating)/.test(t);
    return { asset, place, concern, nonsense };
  }
  const briefWords = normalisedWords(form.project_description).length;
  const hasEmail = form.email.includes('@') && form.email.includes('.');
  const signals = briefSignals(form.project_description);
  const hasBrief = briefWords >= 10 && signals.asset && !signals.nonsense;
  const strongBrief = briefWords >= 18 && signals.asset && (signals.place || signals.concern);
  const canRun = hasEmail && (hasBrief || strongBrief);

  async function submit() {
    setBusy(true); setError(''); setResult(null);
    try {
      const brief = form.project_description;
      const payload = {
        ...form,
        project_type: inferType(brief),
        location: inferLocation(brief),
        size_or_capacity: inferCapacity(brief),
        biggest_concern: inferConcern(brief),
        client_token: clientToken(),
        fingerprint: fingerprint()
      };
      const r = await post('/public-demo/generate', payload);
      setResult(r);
      onComplete?.(r.model);
    } catch (e) {
      let msg = String(e.message || e);
      try {
        const parsed = JSON.parse(msg);
        const detail = parsed.detail;
        msg = typeof detail === 'object' ? (detail.message || (detail.issues ? detail.issues.join(' ') : JSON.stringify(detail))) : (detail || msg);
      } catch {}
      setError(msg);
    } finally { setBusy(false); }
  }

  const example = `Describe the programme.

Examples:
• 500MW hyperscale AI data centre campus in Texas with grid constraints, liquid cooling and accelerated delivery.
• Orbital AI data centre in LEO with modular compute clusters, thermal rejection systems, autonomous servicing and relay communications.
• Lunar logistics hub supporting autonomous cargo operations, landing pads, power storage and surface mobility.
• Semiconductor fab in Arizona with cleanroom, UPW systems, process tools and utility complexity.`;

  return <div className="publicDemoOverlay boomDemoOverlay">
    <div className="publicDemoModal boomDemoModal">
      <button className="publicDemoClose" onClick={onClose}>×</button>
      <div className="boomHeader">
        <p className="demoKicker">CASEY Intelligence Run</p>
        <h2>Describe one programme.</h2>
        <p className="demoSub">Type the project. CASEY infers the sector, location, scale, schedule logic and risk profile.</p>
      </div>
      <label className="boomEmail">Work email
        <input value={form.email} onChange={e=>update('email', e.target.value)} placeholder="you@company.com" />
      </label>
      <label className="projectBriefLabel boomBrief">Programme brief
        <textarea value={form.project_description} onChange={e=>update('project_description', e.target.value)} placeholder={example} autoFocus />
      </label>
      <div className="boomQuality">
        <span className={hasBrief || strongBrief ? 'ok' : ''}>{briefWords} words</span>
        <span className={hasEmail ? 'ok' : ''}>Email</span>
        <span>Earth/Space auto-inferred</span>
        <span>Demo launch mode</span>
      </div>
      {!canRun && !error && form.project_description && <div className="publicDemoHint">Add a real asset, location/environment and main concern. CASEY can infer the rest.</div>}
      {error && <div className="publicDemoError">{error}</div>}
      {busy && <div className="missionProcessing"><Rocket size={22}/><div><b>CASEY is thinking</b><span>Parsing infrastructure archetype · Building benchmark memory · Running schedule intelligence · Mapping risk exposure</span></div></div>}
      {result && <div className="publicDemoSuccess"><b>{result.run_id}</b><span>Intelligence pack generated. The full console is now open behind this window.</span></div>}
      <div className="publicDemoActions boomActions">
        <button className="heroBtn" disabled={!canRun || busy} onClick={submit}>{busy ? 'CASEY is building the pack...' : 'Run CASEY'}</button>
        <button className="ghostBtn" onClick={onClose}>Close</button>
      </div>
      <p className="demoFinePrint">CASEY only runs on credible infrastructure briefs. Outputs are first-pass strategic intelligence, not certified estimates.</p>
    </div>
  </div>;
}

function Loading({ text }) {
  return <motion.div className="loading" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}><Rocket size={44}/><b>{text || 'Building connected model...'}</b><span>Cost · Schedule · QCRA · QSRA · Risk Register · Board Pack</span></motion.div>;
}
function App() {
  const [show, setShow] = useState(true);
  const [briefing, setBriefing] = useState(false);
  const [trialOpen, setTrialOpen] = useState(false);
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
  
const scenarioInsightMap95 = {
  base: 'Mechanical completion is not the true finish line; validated production readiness and deviation closure are the real board decision gates.',
  faster: 'Acceleration is now the risk: CQV, clean-utility validation and media-fill readiness are being forced into the same window. The programme may finish construction before it is safe to release product.',
  cheaper: 'The cheaper case is not simply cheaper. It has likely transferred cost into operational fragility: leaner redundancy, weaker qualification float and higher late-stage validation disruption.',
  lower_risk: 'The lower-risk case buys regulatory confidence by protecting CQV float, GMP turnover and deviation closure before commercial ramp-up.',
  premium: 'Premium delivery protects market launch by buying utility resilience, validation capacity and stronger batch-release readiness.'
};
const scenarioTrade95 = {
  base: ['Balanced cost, time and evidence posture.','Maintains a credible reference case for board challenge.','Does not buy extra time certainty or capital efficiency.'],
  faster: ['You bought time by spending money and consuming recovery float.','Earlier revenue / market-entry option and stronger strategic timing.','CQV float, interface stability, procurement optionality and late-stage recovery.'],
  cheaper: ['You cut capital authorization by moving risk into resilience, redundancy and start-up.','Lower initial approval number and reduced near-term capital draw.','Operational resilience, commissioning flexibility, contingency adequacy and lifecycle certainty.'],
  lower_risk: ['You bought confidence by adding assurance, float and procurement evidence.','Reduced P80/P90 exposure and stronger board approval defensibility.','Earlier revenue date and lean capital posture.'],
  premium: ['You bought resilience, redundancy and strategic optionality.','Higher operational resilience, stronger procurement certainty and better downside protection.','Lowest-capex authorization case.']
};

function scenarioAdjustedModel(currentModel, nextScenario) {
    if (!currentModel) return currentModel;
    const factors = {
      base: { cost: 1.00, schedule: 1.00, conf: 0, risk: currentModel._base_risk || currentModel.risk || 'Medium-High', label: 'Base' },
      faster: { cost: 1.14, schedule: 0.82, conf: -10, risk: 'High', label: 'Faster' },
      cheaper: { cost: 0.88, schedule: 1.10, conf: -14, risk: 'High', label: 'Cheaper' },
      lower_risk: { cost: 1.09, schedule: 1.12, conf: 12, risk: 'Medium', label: 'Lower Risk' },
      premium: { cost: 1.22, schedule: 0.96, conf: 18, risk: 'Medium-Low', label: 'Premium' }
    };
    const f = factors[nextScenario] || factors.base;
    const moneyToBn = (v) => {
      const s = String(v || '').replace('$','').trim().toUpperCase();
      if (s.endsWith('T')) return Number(s.slice(0,-1)) * 1000;
      if (s.endsWith('B')) return Number(s.slice(0,-1));
      if (s.endsWith('M')) return Number(s.slice(0,-1)) / 1000;
      return Number(s) || 0;
    };
    const toMoney = (bn) => bn >= 1000 ? '$' + (bn/1000).toFixed(1) + 'T' : bn >= 1 ? '$' + bn.toFixed(1) + 'B' : '$' + (bn*1000).toFixed(0) + 'M';

    const baseCostBn = currentModel._base_cost_bn || moneyToBn(currentModel.cost_p50);
    const baseMonths = currentModel._base_months || parseFloat(String(currentModel.schedule || '').replace(/[^0-9.]/g,'')) || 60;
    const baseConf = currentModel._base_confidence_pct || Number(currentModel.confidence_pct || 55);
    const baseThinking = currentModel._base_casey_thinking || currentModel.casey_thinking || 'CASEY interprets this as an infrastructure programme.';
    const baseSummary = currentModel._base_executive_summary || currentModel.executive_summary || '';

    const p50 = baseCostBn * f.cost;
    const p10 = p50 * 0.80;
    const p90 = p50 * 1.30;
    const months = Math.max(3, Math.round(baseMonths * f.schedule));
    const conf = Math.max(8, Math.min(96, baseConf + f.conf));
    const scenarioNarrative = {
      base: 'Balanced reference case for board challenge.',
      faster: 'Acceleration reduces duration but increases procurement premiums, commissioning interface risk and delivery pressure.',
      cheaper: 'Reduced capital target lowers initial exposure but increases operational and delivery risk.',
      lower_risk: 'Extended preconstruction, assurance and staged delivery improve confidence but increase duration.',
      premium: 'Flagship resilience and redundancy increase capex while strengthening confidence and optionality.'
    }[nextScenario] || 'Scenario sensitivity generated from the base case.';

    return {
      ...currentModel,
      _base_cost_bn: baseCostBn,
      _base_months: baseMonths,
      _base_confidence_pct: baseConf,
      _base_risk: currentModel._base_risk || currentModel.risk,
      _base_casey_thinking: baseThinking,
      _base_executive_summary: baseSummary,
      scenario: nextScenario,
      scenario_label: f.label,
      scenario_trade: (scenarioTrade95[nextScenario]||[])[0],
      scenario_gain: (scenarioTrade95[nextScenario]||[])[1],
      scenario_loss: (scenarioTrade95[nextScenario]||[])[2],
      cost_p50: toMoney(p50),
      cost_p10: toMoney(p10),
      cost_p90: toMoney(p90),
      cost_range: `${toMoney(p10)} - ${toMoney(p90)}`,
      schedule: `${months} months`,
      confidence_pct: conf,
      risk: f.risk,
      executive_summary: `${currentModel.title} scenario view: ${f.label}. CASEY indicates ${toMoney(p50)} P50 exposure, ${toMoney(p10)} to ${toMoney(p90)} range, ${months} month baseline, ${f.risk} risk and ${conf}% confidence. ${scenarioNarrative}`,
      executive_shock_insight: scenarioInsightMap95[nextScenario] || scenarioNarrative,
      board_briefing: [
        scenarioInsightMap95[nextScenario] || scenarioNarrative,
        `Scenario view indicates ${toMoney(p50)} P50 exposure across approximately ${months} months.`,
        `Confidence moves to ${conf}% because cost, schedule and delivery assumptions have been rebalanced.`,
        currentModel.scenario_contradiction || 'This scenario should be used for decision challenge only until assumptions are validated against package evidence.'
      ],
      casey_thinking: `${baseThinking} Scenario lens: ${scenarioNarrative}`,
      scenario_delta_intelligence: [
        { label: 'Cost movement', value: `${Math.round((f.cost-1)*100)}% vs base`, meaning: scenarioNarrative },
        { label: 'Schedule movement', value: `${Math.round((f.schedule-1)*100)}% vs base`, meaning: 'Schedule logic has been rebalanced for this scenario.' },
        { label: 'Confidence movement', value: `${f.conf} pts`, meaning: 'Confidence moved because scenario assumptions changed.' },
        { label: 'Risk posture', value: f.risk, meaning: 'Risk register, QSRA/QCRA and outputs should be read under this scenario lens.' }
      ],
      confidence_breakdown: [
        { driver: 'Procurement certainty', effect: nextScenario === 'cheaper' || nextScenario === 'faster' ? '-9' : '+7', note: 'Long-lead evidence changes scenario confidence.' },
        { driver: 'Schedule logic maturity', effect: nextScenario === 'faster' ? '-8' : '+6', note: 'Critical path and handover gates rebalanced.' },
        { driver: 'Commissioning / validation readiness', effect: nextScenario === 'faster' || nextScenario === 'cheaper' ? '-11' : '+10', note: 'Scenario changes readiness exposure.' },
        { driver: 'Contingency adequacy', effect: nextScenario === 'cheaper' ? '-12' : '+8', note: 'Reserve and assurance posture updated.' }
      ],
      outputs_board_memo: [
        `Decision posture: ${f.label} scenario.`,
        scenarioNarrative,
        `Confidence moves to ${conf}% under this scenario.`,
        'Outputs remain first-pass strategic intelligence, not certified estimate documents.'
      ],
      top_decisions_required: [
        'Confirm the governing critical-path constraint and evidence owner.',
        'Approve scenario-specific procurement and contingency posture.',
        'Lock handover / commissioning / validation decision gates.',
        'Resolve highest-probability interface and readiness risks.',
        'Decide whether the scenario trade-off is acceptable for board approval.'
      ],
      mission_control_cards: [
        ...(currentModel.mission_control_cards || []).filter(c => c.label !== 'SCENARIO LENS').slice(0,3),
        { label: 'SCENARIO LENS', signal: scenarioNarrative, severity: f.risk }
      ]
    };
  }

  async function generate(nextScenario = scenario, nextPrompt = prompt) {
    setError(''); setShow(false);

    // Scenario changes must re-run the server engine so cost, schedule, risk, QCRA/QSRA, XER
    // and exports all move together versus Base. Do not use a front-end-only sensitivity stub.
    setLoading(true); setTab(nextScenario !== 'base' ? 'compare' : 'overview');
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
  const baseVs = model?.scenario_comparison_vs_base || {};
  const costWaterfall = model?.cost_waterfall_vs_base || [];
  const scheduleWaterfall = model?.schedule_waterfall_vs_base || [];
  const qcra = model?.monte_carlo?.qcra || {};
  const qsra = model?.monte_carlo?.qsra || {};
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
    <OneShotDemo open={trialOpen} onClose={() => setTrialOpen(false)} onComplete={(m) => { setModel(m); setShow(false); setTrialOpen(false); setTab('overview'); }} />
    <AnimatePresence>{loading && <Loading text="Building full CASEY intelligence pack..."/>}</AnimatePresence>
    {show && !model && <Hero onBriefing={() => setBriefing(true)} onEarth={runEarth} onSpace={runSpace} onConsole={() => setShow(false)} onTryDemo={() => setTrialOpen(true)}/>} 
    <header className="v50ConsoleTop"><Logo/><nav><button onClick={() => { setModel(null); setShow(true); }}>Home</button><button onClick={() => setBriefing(true)}>Film</button><button onClick={() => setTrialOpen(true)}>Free run</button><button onClick={runEarth}>Earth demo</button><button onClick={runSpace}>Space demo</button><a href={emailLink}>Request access</a></nav></header>
    <main className={model ? 'v50Console' : 'v50Console emptyConsole'}>
      {error && <div className="error">{error}</div>}
      {!model && !show && <section className="commandGrid"><Card className="command"><h1>Generate a live project model</h1><label>Project command</label><textarea value={prompt} onChange={e => setPrompt(e.target.value)} /> <div className="chips">{examples.map(x => <button key={x} onClick={() => setPrompt(x)}>{x}</button>)}</div><div className="grid4"><input value={client} onChange={e => setClient(e.target.value)} placeholder="Client / operator"/><select value={classLevel} onChange={e => setClassLevel(e.target.value)}>{[1,2,3,4,5].map(x => <option key={x} value={x}>Class {x}</option>)}</select><select value={scheduleLevel} onChange={e => setScheduleLevel(e.target.value)}>{[1,2,3,4,5].map(x => <option key={x} value={x}>Level {x}</option>)}</select><select value={scenario} onChange={e => setScenario(e.target.value)}>{scenarios.map(x => <option key={x} value={x}>{x}</option>)}</select></div><button className="primary" onClick={() => generate()}><Sparkles/> Generate full intelligence pack</button></Card><Card><h2>What CASEY will produce</h2>{['Executive summary and recommendation','Direct / indirect / reserve cost view','Scenario-linked estimate, schedule and confidence','Risk register with cause, event, impact and mitigation','QCRA + QSRA curves and tornado drivers','Pricing and next-step contact actions'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card></section>}
      {model && <>
        <section className="confidenceEngineBadge"><b>{model.confidence_engine_label || 'CASEY Confidence Engine'}</b><span>{model.confidence_engine_detail || 'Benchmark + probabilistic + sector-trained reasoning'}</span></section>
        <section className="kpis"><Kpi icon={Globe2} label="Mode / sector" value={model.mode} sub={model.subsector}/><Kpi icon={Activity} label="P50 cost" value={model.cost_p50} sub={model.cost_range}/><Kpi icon={Zap} label="Schedule" value={model.schedule} sub={`QSRA P80 ${model.monte_carlo?.qsra?.p80 || '—'} months`}/><Kpi icon={ShieldAlert} label="Risk / Confidence" value={`${model.risk} / ${model.confidence_pct}%`} sub={model.scenario_label} hot/></section>
        <nav className="tabs">{[['overview','Overview'],['compare','Scenarios'],['delta','Scenario Intel'],['cost','Cost'],['schedule','Schedule'],['risk','Risk'],['monte','QCRA/QSRA'],['outputs','Outputs'],['advisor','Advisor'],['method','Methodology'],['pricing','Pricing']].map(x => <button key={x[0]} className={tab===x[0]?'active':''} onClick={() => setTab(x[0])}>{x[1]}</button>)}</nav>
        {tab === 'overview' && <>
          {model.executive_shock_insight && <section className="layout one"><Card className="shockCard"><h2>Executive shock insight</h2><p>{model.executive_shock_insight}</p></Card></section>}
          <section className="layout two">
            <Card><h2>Executive intelligence summary</h2><p className="big">{model.executive_summary || `${model.title} has been classified as ${model.subsector}. CASEY generated a first-pass cost, schedule, risk and confidence model for the selected scenario.`}</p><div className="miniMetrics"><b><span>Direct cost</span>{fmt(direct)}</b><b><span>Indirect cost</span>{fmt(indirect)}</b><b><span>Risk / reserve</span>{fmt(reserves)}</b></div><h3>Recommendation</h3>{(model.next_best_actions || []).slice(0,5).map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card>
            <Card><h2>Board briefing</h2>{(model.board_briefing || model.board_challenge_questions || []).slice(0,5).map((x,i)=><div className="reason" key={String(x)}><span>{i+1}</span>{x}</div>)}<h3>CASEY thinking</h3><p className="caseyThinking">{model.casey_thinking || 'CASEY interprets this as a system-of-systems infrastructure programme requiring cost, schedule, risk and decision intelligence.'}</p></Card>
          </section>
          {baseVs?.base && <section className="layout two">
            <Card className="shockCard"><h2>Scenario vs Base</h2><p>{baseVs.plain_english}</p><div className="miniMetrics"><b><span>Base P50</span>{baseVs.base.cost_p50}<small>{baseVs.base.schedule_months} mo · {baseVs.base.confidence_pct}%</small></b><b><span>{baseVs.selected.scenario} P50</span>{baseVs.selected.cost_p50}<small>{baseVs.selected.schedule_months} mo · {baseVs.selected.confidence_pct}%</small></b><b><span>Delta</span>{baseVs.delta.cost_direction === 'same' ? 'No cost move' : `${baseVs.delta.cost} ${baseVs.delta.cost_direction}`}<small>{baseVs.delta.months} mo · {baseVs.delta.confidence_pts} pts</small></b></div></Card>
            <Card><h2>What changed and why</h2>{(model.scenario_delta_intelligence || []).slice(0,5).map((x,i)=><div className="reason" key={i}><span>{i+1}</span><b>{x.label}: {x.value}</b><br/>{x.meaning}</div>)}</Card>
          </section>}
          <section className="layout two">
            <Card><h2>Mission control signals</h2><div className="missionCardGrid">{(model.mission_control_cards || []).slice(0,6).map((c,i)=><div className="intelCard" key={i}><b>{c.label}</b><p>{c.signal}</p><span>{c.severity}</span></div>)}</div></Card>
            <Card><h2>Uncertainty narrative</h2><p>{model.uncertainty_narrative?.estimate_maturity}</p><p>{model.uncertainty_narrative?.schedule_maturity}</p><p>{model.uncertainty_narrative?.interpretation}</p><h3>Benchmark comparison</h3>{(model.benchmark_comparison || []).slice(0,4).map((b,i)=><div className="reason" key={i}><span>{i+1}</span><b>{b.archetype}</b> · {b.anchor_cost} · {b.anchor_duration_months} months</div>)}</Card>
          </section>
          <section className="layout two">
            <Card><h2>Confidence drivers</h2>{(model.sector_confidence_drivers || ['Benchmark similarity: high where comparable infrastructure archetypes exist','Scope maturity: concept / budget level until package evidence is supplied','Procurement certainty: sensitive to long-lead equipment and market capacity','Schedule maturity: improves when critical path and commissioning logic are validated','Interface exposure: controlled by utilities, systems integration and operational constraints']).map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card>
            <Card><h2>Why CASEY generated this</h2>{(model.why_casey_generated_this || ['CASEY detected the infrastructure asset and operating environment from the brief','The programme was mapped to benchmark memory and sector archetypes','Cost, schedule and risk were calibrated against class maturity and delivery complexity','The narrative is designed for early board challenge, not certified pricing']).map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card>
          </section>
          <section className="layout two">
            <Card><h2>Primary cost drivers</h2>{(model.sector_primary_cost_drivers || ['Utility / enabling infrastructure','Specialist systems and long-lead equipment','Commissioning and validation complexity','Programme management, preliminaries and indirects','Risk reserve driven by procurement and interface uncertainty']).map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card>
            <Card><h2>Top schedule threats</h2>{(model.sector_schedule_threats || ['Utility energisation delay','Long-lead equipment procurement and supplier capacity','Design freeze instability and scope movement','Systems integration and commissioning bottlenecks','Approvals, safety case, permitting or operational access constraints']).map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card>
          </section>
        </>}

        {tab === 'compare' && <section className="layout two"><Card><h2>Scenario comparison</h2><p className="big">Switch options before paying for another advisory cycle. CASEY keeps cost, schedule, risk and confidence aligned across every export.</p><div className="scenarioCompare">{scenarios.map(s => { const active = s === scenario; return <button key={s} className={active?'active':''} onClick={() => generate(s)}><b>{s.replace('_',' ')}</b><span>{active ? 'current model' : 'run scenario'}</span></button> })}</div></Card><Card><h2>Buyer decision lens</h2>{['Base: balanced reference case for board challenge','Faster: premium acceleration with higher interface exposure','Cheaper: reduced cost target with increased delivery risk','Lower Risk: slower, more assured delivery with higher reserves','Premium: flagship resilience, higher capex and stronger confidence'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card></section>}
        {tab === 'cost' && <section className="layout two"><Card><h2>Scenario cost bridge vs Base</h2><p className="chartCaption">This explains why the selected scenario is cheaper or more expensive than Base before showing the workbook lines.</p>{costWaterfall.map((x,i)=><div className={`reason ${x.kind==='total'?'deltaReason':''}`} key={i}><span>{i+1}</span><b>{x.driver}</b><br/>{x.kind==='total'?x.value:(x.value_bn>=0?'+':'−') + ' ' + x.value}</div>)}<h3>Cost estimate workbook</h3><Table rows={costs} cols={[["cbs","CBS"],["description","Description"],["type","Type"],["p10_bn","Low/P10"],["p50_bn","Most likely/P50"],["p90_bn","High/P90"],["impact_basis","Basis"]]} moneyCols={["p10_bn","p50_bn","p90_bn"]}/></Card><Card><h2>Cost composition</h2><p className="chartCaption">Direct, indirect and reserve are scenario-controlled. For the detailed uncertainty view use QCRA/QSRA.</p><ResponsiveContainer width="100%" height={320}><BarChart data={[{name:'Direct',value:direct},{name:'Indirect',value:indirect},{name:'Reserve',value:reserves}]}><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis dataKey="name"/><YAxis/><Tooltip/><Bar dataKey="value" fill="#8df7ff"/></BarChart></ResponsiveContainer></Card></section>}
        {tab === 'schedule' && <section className="layout two"><Card><h2>Schedule bridge vs Base</h2><p className="chartCaption">This is the month-by-month reason the scenario becomes faster or slower than Base.</p>{scheduleWaterfall.map((x,i)=><div className={`reason ${x.kind==='total'?'deltaReason':''}`} key={i}><span>{i+1}</span><b>{x.driver}</b><br/>{x.kind==='total'?`${x.months} months`:(x.months>=0?'+':'') + x.months + ' months'}</div>)}<h3>Scenario schedule logic</h3><Table rows={schedule} cols={[["activity_id","Activity"],["phase","Phase"],["activity","Name"],["predecessor","Pred"],["duration_months","Months"],["critical","Critical"],["basis","Basis"]]}/></Card><Card><h2>QSRA finish-date curve</h2><p className="chartCaption">P50 equals the headline schedule. P80/P90 are the finish-risk dates, not another baseline.</p><div className="metrics"><div>P50<b>{qsra.p50} mo</b></div><div>P80<b>{qsra.p80} mo</b></div><div>P90<b>{qsra.p90} mo</b></div></div><ResponsiveContainer width="100%" height={280}><LineChart data={curve}><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis dataKey="percentile"/><YAxis/><Tooltip formatter={(v) => [`${v} months`, "QSRA finish date"]}/><ReferenceLine x={50} stroke="#ffffff88" label="P50 = headline"/><ReferenceLine x={80} stroke="#ffffff55" label="P80 = board risk"/><Line type="monotone" name="QSRA finish date" dataKey="schedule_months" stroke="#b18cff" strokeWidth={3}/></LineChart></ResponsiveContainer>{(model.monte_carlo?.curve_readout || []).slice(1).map((x,i)=><div className="reason" key={i}><span>{i+1}</span>{x}</div>)}</Card></section>}
        {tab === 'risk' && <section className="layout two"><Card><h2>Risk Register Pro</h2><p>Risk output should include cause, event, impact, owner, mitigation and links to WBS/CBS.</p><Table rows={risks} cols={[['risk_id','ID'],['title','Risk'],['cause','Cause'],['event','Event'],['impact','Impact'],['probability_pct','Prob %'],['activity_id','Activity'],['cbs','CBS'],['owner','Owner'],['mitigation','Mitigation']]}/></Card><Card><h2>Top exposure drivers</h2><ResponsiveContainer width="100%" height={380}><BarChart data={tornado} layout="vertical"><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis type="number"/><YAxis dataKey="driver" type="category" width={150}/><Tooltip/><Bar dataKey="contribution" fill="#8df7ff"/></BarChart></ResponsiveContainer></Card></section>}
        {tab === 'monte' && <section className="layout two"><Card><h2>QCRA cost range curve</h2><p className="chartCaption">This is not a spend forecast over time. It is the probability range: P50 matches the headline cost, P80/P90 show board contingency exposure.</p><div className="metrics"><div>P50 headline<b>{model.cost_p50}</b></div><div>P80 risk exposure<b>{fmt(qcra.p80)}</b></div><div>P90 stress case<b>{fmt(qcra.p90)}</b></div></div><ResponsiveContainer width="100%" height={280}><AreaChart data={curve}><defs><linearGradient id="caseyG" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stopColor="#8df7ff" stopOpacity=".55"/><stop offset="1" stopColor="#8df7ff" stopOpacity="0"/></linearGradient></defs><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis dataKey="percentile"/><YAxis/><Tooltip formatter={(v) => [`$${Number(v).toFixed(1)}B`, "QCRA total outturn"]}/><ReferenceLine x={50} stroke="#ffffff88" label="P50 = headline"/><ReferenceLine x={80} stroke="#ffffff55" label="P80 = board risk"/><Area type="monotone" name="QCRA total outturn" dataKey="cost_bn" stroke="#8df7ff" fill="url(#caseyG)"/></AreaChart></ResponsiveContainer>{(model.monte_carlo?.curve_readout || []).slice(0,1).map((x,i)=><div className="reason" key={i}><span>i</span>{x}</div>)}<div className="reason"><span>!</span>This curve is a probability distribution, not spend over time. The x-axis is confidence percentile. P50 equals the headline estimate; P80/P90 are board downside exposure.</div></Card><Card><h2>QSRA schedule range curve</h2><p className="chartCaption">P50 matches the headline duration. P80/P90 show the likely board conversation if critical path risk lands.</p><div className="metrics"><div>P50 headline<b>{qsra.p50} mo</b></div><div>P80 risk date<b>{qsra.p80} mo</b></div><div>P90 stress date<b>{qsra.p90} mo</b></div></div><ResponsiveContainer width="100%" height={280}><LineChart data={curve}><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis dataKey="percentile"/><YAxis/><Tooltip formatter={(v) => [`${v} months`, "QSRA finish date"]}/><ReferenceLine x={50} stroke="#ffffff88" label="P50 = headline"/><ReferenceLine x={80} stroke="#ffffff55" label="P80 = board risk"/><Line type="monotone" name="QSRA finish date" dataKey="schedule_months" stroke="#b18cff" strokeWidth={3}/></LineChart></ResponsiveContainer>{(model.monte_carlo?.curve_readout || []).map((x,i)=><div className="reason" key={i}><span>{i+1}</span>{x}</div>)}</Card></section>}
        {tab === 'delta' && <section className="layout two">
          <Card><h2>Strategic Delta Intelligence</h2><p>What changed because this scenario was selected.</p>
            {(model.scenario_delta_intelligence || []).map((x,i)=><div className="reason" key={i}><span>{i+1}</span><b>{x.label}: {x.value}</b><br/>{x.meaning}</div>)}
          </Card>
          <Card><h2>Confidence Breakdown</h2><p>CASEY explains why confidence moved.</p>
            {(model.confidence_breakdown || []).map((x,i)=><div className="reason" key={i}><span>{i+1}</span><b>{x.driver}: {x.effect}</b><br/>{x.note}</div>)}
          </Card>
          <Card><h2>Top Decisions Required</h2>
            {(model.top_decisions_required || []).map((x,i)=><div className="reason" key={i}><span>{i+1}</span>{x}</div>)}
          </Card>
          <Card><h2>Board Memo Snapshot</h2>
            {(model.outputs_board_memo || []).map((x,i)=><div className="reason" key={i}><span>{i+1}</span>{x}</div>)}
          </Card>
        </section>}

        {tab === 'delta' && <section className="layout two">
          <Card className="shockCard"><h2>Scenario Consequence vs Base</h2><p>{model.scenario_trade || 'Scenario trade-off analysis.'}</p>
            {(model.scenario_delta_intelligence || []).map((x,i)=><div className="reason deltaReason" key={i}><span>{i+1}</span><b>{x.label}: {x.value}</b><br/>{x.meaning}</div>)}
          </Card>
          <Card><h2>What You Gained / What You Gave Up</h2>
            <div className="reason"><span>+</span><b>Gained</b><br/>{model.scenario_gain || 'Scenario benefit.'}</div>
            <div className="reason"><span>−</span><b>Gave up</b><br/>{model.scenario_loss || 'Scenario consequence.'}</div>
            <div className="reason"><span>!</span><b>Curve meaning</b><br/>{model.monte_carlo?.curve_interpretation || 'QCRA/QSRA shape reflects scenario uncertainty.'}</div>
          </Card>
          <Card><h2>Confidence Breakdown</h2>
            {(model.confidence_breakdown || []).map((x,i)=><div className="reason" key={i}><span>{i+1}</span><b>{x.driver}: {x.effect}</b><br/>{x.note}</div>)}
          </Card>
          <Card><h2>Top Decisions Required</h2>
            {(model.top_decisions_required || []).map((x,i)=><div className="reason" key={i}><span>{i+1}</span>{x}</div>)}
          </Card>
        </section>}

        {tab === 'outputs' && <section className="layout two"><Card><h2>Generated Artefacts</h2><p>The public demo previews the intelligence pack. Enterprise access unlocks the live generated controls deliverables.</p><div className="exports v50Exports lockedExports">
          {[
            ['Cost Model XLSX', FileSpreadsheet],
            ['Risk Register XLSX', Database],
            ['PRA Schedule XER', Workflow],
            ['Schedule Levels CSV', Workflow],
            ['Audit File JSON', Brain],
            ['Full Pack ZIP', Download]
          ].map(([label,Icon])=><button key={label} onClick={()=>setError('Enterprise access required to unlock generated artefacts: Cost Workbook, Risk Register, XER Schedule, QRA/QSRA Pack and Board Pack.')}><Icon/> Generate {label}</button>)}
          <a className="contactBtn" href={emailLink}><Mail/> Request Enterprise Pack</a></div></Card><Card><h2>What the pack delivers</h2>{['Executive control centre with project, scenario, class, level and confidence clearly identified','Scenario comparison covering Base, Faster, Cheaper, Lower Risk and Premium cases','Selected estimate class plus all class levels for audit and challenge','Direct, indirect and reserve cost views with QCRA cost curve and cost tornado','All schedule levels with QSRA schedule curve and schedule tornado','Risk register with cause, event, impact, owner, mitigation, trigger and quantified likelihood','Basis of Estimate, assumptions, exclusions and benchmark validation','Commercial next steps: buyer action, procurement challenge and board decision path'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card></section>}

        {tab === 'advisor' && <section className="layout two"><Card><h2>Ask CASEY</h2><div className="chatBox">{chat.map((m,i)=><div key={i} className={`msg ${m.role}`}>{m.text}</div>)}</div><div className="ask"><input value={chatQ} onChange={e=>setChatQ(e.target.value)} onKeyDown={e=>{if(e.key==='Enter')ask()}} placeholder="Ask why the cost or risk moved..."/><button onClick={ask}>Ask</button></div></Card><Card><h2>Upload estimate challenge</h2><p>Use this to show buyers how CASEY can challenge a Tier 1 estimate.</p><label className="upload"><Upload size={18}/> Upload file<input type="file" onChange={upload}/></label><button className="secondary" onClick={()=>setUploadResult({review:'Sample contractor estimate challenge', findings:['Direct costs above benchmark in power train and cooling package','Indirects and preliminaries need clearer split from reserves','Schedule contingency understated against critical path risks','Risk allowance should separate QCRA cost and QSRA schedule exposure'], next_action:'Request rate build-up, supplier quotes, basis of estimate and revised risk register.'})}>Run sample challenge</button>{uploadResult && <pre>{JSON.stringify(uploadResult,null,2)}</pre>}</Card></section>}

        {tab === 'method' && <section className="layout two"><Card><h2>How CASEY calculated this</h2>{['Cost model: selected class estimate, sector template, location factor, complexity factor and scenario modifier.','Schedule model: level-based delivery logic, phase durations, critical path sensitivity and scenario acceleration/delay factors.','QCRA: cost exposure model using low / most likely / high impacts and risk-weighted contingency.','QSRA: schedule exposure model using activity-linked O/M/P delay ranges and critical path sensitivity.','Confidence score: class maturity, schedule detail, scenario risk profile and location/space complexity.'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card><Card><h2>Commercial readiness</h2><p className="big">This is first-pass project controls intelligence. It is designed to accelerate challenge, option testing and board preparation before final contractor tender or signed cost plan.</p><a className="contactBtn huge" href={emailLink}><Mail/> Send project for review</a></Card></section>}
        {tab === 'pricing' && <section className="layout two"><Card><h2>CASEY Access</h2><div className="pricingGrid"><div className="priceCard"><b>Pilot</b><strong>Request pricing</strong><span>Guided project review, sample outputs and executive walkthrough.</span><a href={emailLink}>Request pilot</a></div><div className="priceCard hot"><b>Professional</b><strong>Full project pack</strong><span>Cost, schedule, risk, QCRA/QSRA and export pack.</span><a href={emailLink}>Request access</a></div><div className="priceCard"><b>Enterprise</b><strong>Private deployment</strong><span>SSO, teams, benchmark library, private models and audit trail.</span><a href={emailLink}>Book demo</a></div></div></Card><Card><h2>Send this project</h2><p className="big">Turn demo interest into pipeline immediately.</p><a className="contactBtn huge" href={emailLink}><Mail/> Send project to CASEY</a><button className="primary" onClick={() => download('/export/all', model, 'CASEY_Output_Pack.zip')}>Download full pack</button></Card></section>}
      </>}
    </main>
  </div>;
}

createRoot(document.getElementById('root')).render(<App/>);
