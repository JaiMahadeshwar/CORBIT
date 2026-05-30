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

const PROD_URL = import.meta.env.VITE_API_URL || import.meta.env.VITE_BACKEND_URL || 'https://corbit-1.onrender.com';
const API_CANDIDATES = [PROD_URL, 'http://127.0.0.1:8000', 'http://localhost:8000'].filter(Boolean);
let API = API_CANDIDATES[0];
async function apiFetch(path, options, timeoutMs = 45000) {
  let lastError;
  for (const base of API_CANDIDATES) {
    try {
      const controller = new AbortController();
      const tid = setTimeout(() => controller.abort(), timeoutMs);
      const r = await fetch(base + path, { ...options, signal: controller.signal });
      clearTimeout(tid);
      API = base;
      return r;
    } catch (e) { lastError = e; }
  }
  throw lastError || new Error('CASEY backend unreachable — try again in 20 seconds');
}
function safeRender(value) {
  if (value === undefined || value === null) return '';
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return String(value);
  if (value?.nativeEvent || value?.currentTarget || value?.target) return '';
  if (value instanceof Element) return value.innerText || value.textContent || '';
  if (Array.isArray(value)) return value.map(safeRender).join('\n');
  try {
    const seen = new WeakSet();
    return JSON.stringify(value, (k, v) => {
      if (k === '_owner' || k === '__reactFiber$' || k === '__reactProps$') return undefined;
      if (typeof v === 'function') return undefined;
      if (v instanceof Element) return v.innerText || v.textContent || '';
      if (v && typeof v === 'object') {
        if (seen.has(v)) return '[circular]';
        seen.add(v);
      }
      return v;
    }, 2);
  } catch (_) { return String(value || ''); }
}

function asList(v) {
  if (!v) return [];
  if (Array.isArray(v)) return v.map(x => typeof x === 'string' ? x : safeRender(x));
  if (typeof v === 'string') return [v];
  return Object.values(v).flat().map(x => typeof x === 'string' ? x : safeRender(x));
}

function ProfessionalIntakeResult({ result, model }) {
  const [expanded, setExpanded] = React.useState(null);
  const baselineP50 = model?.cost_p50 || moneyLocal(model?.cost_p50_bn || model?.p50_cost_bn || 0) || '—';
  const baselineMonths = model?.schedule_months || model?.duration_months || (String(model?.schedule || '').match(/\d+/)||[])[0] || '—';
  const baselineConf = model?.confidence_pct ?? '—';
  const qcraP80 = model?.monte_carlo?.qcra?.p80;
  const qsraP80 = model?.monte_carlo?.qsra?.p80;

  if (!result) return (
    <div className="intakeEmpty proEmpty">
      <div className="intakeEmptyIcon">⌁</div>
      <b>No client file challenged yet</b>
      <span>Use one of the three professional challenge buttons above, or upload a workbook/XER. CASEY will show baseline, challenge delta, benchmark comparison and required evidence — not raw JSON.</span>
    </div>
  );

  let r = result;
  if (typeof r === 'string') { try { r = JSON.parse(r); } catch { r = { review: r }; } }
  const src = r.source_intelligence || {};
  const cm = r.challenge_model || {};
  const file = r.filename || 'Client_Source_Bundle.xlsx';
  const fileType = (r.file_type || (file.toLowerCase().includes('.xer') ? 'Schedule XER' : file.toLowerCase().includes('risk') ? 'Risk register' : 'Cost estimate')).toString();
  const confPct = Number(cm.confidence_pct ?? model?.confidence_pct ?? 55);
  const challengeP50 = cm.p50_bn ? ('$' + Number(cm.p50_bn).toFixed(1) + 'B') : baselineP50;
  const challengeP80 = cm.p80_bn ? ('$' + Number(cm.p80_bn).toFixed(1) + 'B') : (qcraP80 ? moneyLocal(qcraP80) : '—');
  const challengeP90 = cm.p90_bn ? ('$' + Number(cm.p90_bn).toFixed(1) + 'B') : (model?.monte_carlo?.qcra?.p90 ? moneyLocal(model.monte_carlo.qcra.p90) : '—');
  const deltaBn = cm.delta_bn ?? (cm.p80_bn && model?.cost_p50_bn ? Number(cm.p80_bn) - Number(model.cost_p50_bn) : null);
  const deltaText = deltaBn !== null && deltaBn !== undefined && !Number.isNaN(Number(deltaBn)) ? ((Number(deltaBn) >= 0 ? '+' : '−') + '$' + Math.abs(Number(deltaBn)).toFixed(1) + 'B latent exposure') : 'Exposure delta requires source bundle';
  const scheduleDelta = cm.schedule_delta_months ?? src.xer?.schedule_delta_months ?? null;
  const scheduleDeltaText = scheduleDelta !== null && scheduleDelta !== undefined ? ((Number(scheduleDelta) >= 0 ? '+' : '−') + Math.abs(Number(scheduleDelta)) + ' months schedule exposure') : `QSRA P80 ${qsraP80 || '—'} months`;

  const verdict = confPct < 45 ? { label: 'Further assurance required before approval', color: '#ff9940', bg: 'rgba(255,153,64,0.10)' }
    : confPct < 62 ? { label: 'Board challenge likely — evidence package incomplete', color: '#f7d774', bg: 'rgba(247,215,116,0.10)' }
    : { label: 'Conditional approval possible with evidence closure', color: '#8df7ff', bg: 'rgba(141,247,255,0.08)' };

  const findings = asList(r.findings).length ? asList(r.findings) : [
    'Source structure normalised and compared against the active programme baseline.',
    'Contingency and schedule logic require P80/P90 reconciliation before board approval.',
    'Commercial basis, owner accountability and evidence closure remain the deciding conditions.'
  ];
  const issues = asList(r.red_flags || r.professional_observations || []).length ? asList(r.red_flags || r.professional_observations) : [
    'Submitted allowance does not yet demonstrate a clear link to quantified residual exposure.',
    'The source bundle should reconcile estimate, schedule and risk register at CBS/WBS/activity level.',
    'Comparable programmes show higher volatility where commissioning and interface evidence is incomplete.'
  ];
  const questions = asList(r.board_challenge_questions || r.board_challenge_questions).length ? asList(r.board_challenge_questions || r.board_challenge_questions) : [
    'Where does the client P50 reconcile to the P80/P90 downside, line by line?',
    'Which CBS/WBS package owns the largest unpriced exposure and who signs the evidence closure?',
    'Is contingency sized from quantified risk, or applied as a percentage allowance?',
    'Which XER activities drive the board date, and are their predecessors, calendars and constraints defensible?'
  ];
  const epcFlags = asList(r.epc_flags || []);
  const sectorDetected = r.sector_detected || '';
  const verdictFromEngine = r.challenge_verdict || '';
  const confImpact = r.confidence_impact;
  const caseyComparison = r.casey_comparison || null;
  const engMetrics = { risks: r.risks_parsed||0, emv: r.emv_bn||0, p90: r.p90_bn||0, acts: r.activities_parsed||0, evReq: r.ev_req||0, noTrig: r.no_trig||0 };
  const next = asList(r.next_steps || r.next_action).length ? asList(r.next_steps || r.next_action) : [
    'Request native cost workbook, risk register and XER schedule as one source bundle.',
    'Reconcile risk residuals to reserve and schedule drivers to QSRA before approval.',
    'Issue an independent board challenge note with open evidence owners and closure dates.'
  ];
  const benchmark = asList(r.benchmark_comparison || []).length ? asList(r.benchmark_comparison) : [
    `Compared with rail/transit benchmark memory, comparable systems-integration programmes typically require explicit P80/P90 reserve reconciliation before approval.`,
    `Programmes with possession, signalling and operator-acceptance constraints generally carry wider delivery tails than civil-progress reporting suggests.`
  ];

  return (
    <div className="intakeResult proChallenge professionalReview">
      <div className="challengeFileBar proFileBar">
        <div className="challengeFileInfo">
          <span className="challengeFileTag">{sectorDetected || 'Independent'} source-file challenge</span>
          <b className="challengeFileName">{file}</b>
          <em>{fileType} · {r.size_bytes ? `${Math.round(r.size_bytes/1024)} KB` : 'sample messy client file'} · benchmark comparison enabled</em>
        </div>
        <div style={{display:'flex',gap:'6px',alignItems:'center'}}>
          {sectorDetected&&<div style={{fontSize:'9px',fontWeight:'900',letterSpacing:'.1em',color:'#8df7ff',background:'rgba(141,247,255,0.1)',border:'1px solid rgba(141,247,255,0.2)',borderRadius:'2px',padding:'3px 8px'}}>{sectorDetected}</div>}
          {verdictFromEngine&&<div style={{fontSize:'9px',fontWeight:'900',color:verdictFromEngine.includes('REQUIRED')?'#ef4444':'#f59e0b',background:verdictFromEngine.includes('REQUIRED')?'rgba(239,68,68,0.1)':'rgba(245,158,11,0.1)',border:'1px solid',borderColor:verdictFromEngine.includes('REQUIRED')?'rgba(239,68,68,0.3)':'rgba(245,158,11,0.3)',borderRadius:'2px',padding:'3px 8px'}}>{verdictFromEngine}</div>}
          {!verdictFromEngine&&<div className="challengeLiveTag">CLIENT-SIDE REVIEW</div>}
        </div>
      </div>

      <div className="baselineVsChallenge">
        <div className="bvcBox baseline"><span>Programme baseline remains</span><b>{baselineP50}</b><em>{baselineMonths} months · {baselineConf}% confidence</em></div>
        <div className="bvcArrow">→</div>
        <div className="bvcBox delta"><span>Challenge exposure identified</span><b>{deltaText}</b><em>{scheduleDeltaText}</em></div>
      </div>

      <div className="challengeVerdictBig" style={{background: verdict.bg, borderColor: verdict.color}}>
        <div className="verdictLabel">CASEY professional challenge opinion</div>
        <div className="verdictResult" style={{color: verdict.color}}>{verdict.label}</div>
        <div className="verdictSub">This does not replace the programme baseline. It identifies latent downside in the uploaded source file and shows what must be evidenced before the board can rely on the submitted position.</div>
      </div>

      <div className="challengeMetricRow professionalMetrics">
        {engMetrics.risks>0 ? <div className="cm hot"><span>Risks parsed</span><b style={{color:'#f59e0b'}}>{engMetrics.risks}</b></div> : <div className="cm hot"><span>Submitted / inferred P50</span><b>{challengeP50}</b></div>}
        {engMetrics.p90>0 ? <div className="cm"><span>P90 downside</span><b style={{color:'#ef4444'}}>${engMetrics.p90.toFixed(1)}B</b></div> : <div className="cm"><span>Challenge P80</span><b style={{color:'#ff9940'}}>{challengeP80}</b></div>}
        {engMetrics.emv>0 ? <div className="cm"><span>Total EMV</span><b style={{color:'#ff9940'}}>${engMetrics.emv.toFixed(2)}B</b></div> : engMetrics.acts>0 ? <div className="cm"><span>Activities</span><b>{engMetrics.acts}</b></div> : <div className="cm"><span>Stress P90</span><b style={{color:'#ff6b7d'}}>{challengeP90}</b></div>}
        <div className="cm"><span>Evidence quality</span><b style={{color:confPct<60?'#f7d774':'#8df7ff'}}>{confImpact||confPct}%</b></div>
      </div>

      <div className="sourceBundle">
        <div><span>Cost workbook signals</span><b>{src.cost?.cost_lines_found ?? 0}</b><em>{src.cost?.direct_bn ? `Direct signal $${Number(src.cost.direct_bn).toFixed(1)}B` : 'Basis mapping required'}</em></div>
        <div><span>Risk register rows</span><b>{src.risk?.risk_rows_found ?? 0}</b><em>{src.risk?.emv_bn ? `EMV signal $${Number(src.risk.emv_bn).toFixed(1)}B` : 'Residual-to-reserve test required'}</em></div>
        <div><span>XER schedule logic</span><b>{src.xer?.task_count ?? 0}</b><em>{src.xer?.open_end_risk_count ? `${src.xer.open_end_risk_count} weak/open logic points` : 'Schedule logic not loaded'}</em></div>
      </div>

      <div className="challengeSection benchmarkSection">
        <div className="challengeSectionHead"><span className="csh-num">B</span> Benchmark comparison</div>
        {benchmark.map((x,i)=><div className="challengeFinding" key={i}><span className="cfNum">{i+1}</span><span className="cfText">{x}</span></div>)}
      </div>

      <div className="challengeSection">
        <div className="challengeSectionHead"><span className="csh-num">{findings.length}</span> Source-file findings</div>
        {findings.map((x,i)=>(
          <div className="challengeFinding" key={i} onClick={()=>setExpanded(expanded===`f${i}`?null:`f${i}`)}>
            <span className="cfNum">{i+1}</span><span className="cfText">{x}</span><span className="cfChev">{expanded===`f${i}`?'▲':'▼'}</span>
            {expanded===`f${i}` && <div className="cfExpand">Professional reliance test: confirm source tab, line owner, basis statement, quantified residual exposure and closure evidence.</div>}
          </div>
        ))}
      </div>

      <div className="challengeSection">
        <div className="challengeSectionHead danger"><span className="csh-num">!</span> Commercial observations</div>
        {issues.slice(0,6).map((x,i)=><div className="challengeFlag professionalFlag" key={i}><span className="cfFlag">•</span><span>{x}</span></div>)}
      </div>

      {epcFlags.length>0&&(
        <div className="challengeSection">
          <div className="challengeSectionHead" style={{color:'#ff6b7d'}}><span className="csh-num">⚠</span> EPC / CONTRACTOR FLAGS — READ BEFORE APPROVING</div>
          {epcFlags.map((x,i)=>(
            <div key={i} style={{display:'flex',gap:'8px',padding:'7px 0',borderBottom:'1px solid rgba(255,59,92,0.08)',fontSize:'12px',color:'#ff9aa8',alignItems:'flex-start'}}>
              <span style={{color:'#ff3b5c',flexShrink:0,fontWeight:'900',fontSize:'10px',marginTop:'1px'}}>EPC</span>
              <span>{x}</span>
            </div>
          ))}
        </div>
      )}

      {caseyComparison&&(
        <div className="challengeSection">
          <div className="challengeSectionHead" style={{color:'#8df7ff'}}><span className="csh-num">⚡</span> CASEY vs SUBMITTED DOCUMENT</div>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'8px'}}>
            <div style={{padding:'10px',background:'rgba(239,68,68,0.07)',borderRadius:'3px',border:'1px solid rgba(239,68,68,0.2)'}}>
              <div style={{fontSize:'9px',color:'#ef4444',fontWeight:'900',letterSpacing:'.1em',marginBottom:'6px'}}>SUBMITTED POSITION</div>
              {caseyComparison.client_p90&&<div style={{fontSize:'13px',color:'#ff9aa8',marginBottom:'2px'}}>P90: <b>{caseyComparison.client_p90}</b></div>}
              {caseyComparison.client_p50&&<div style={{fontSize:'13px',color:'#ff9aa8',marginBottom:'2px'}}>P50: <b>{caseyComparison.client_p50}</b></div>}
              <div style={{fontSize:'11px',color:'#94a3b8',marginTop:'4px'}}>{caseyComparison.client_risks||0} risks · {caseyComparison.open_exposures||0} open</div>
              {caseyComparison.governance&&<div style={{fontSize:'10px',fontWeight:'900',color:'#ef4444',marginTop:'6px'}}>{caseyComparison.governance}</div>}
            </div>
            <div style={{padding:'10px',background:'rgba(141,247,255,0.05)',borderRadius:'3px',border:'1px solid rgba(141,247,255,0.2)'}}>
              <div style={{fontSize:'9px',color:'#8df7ff',fontWeight:'900',letterSpacing:'.1em',marginBottom:'6px'}}>CASEY SECTOR BENCHMARK</div>
              {caseyComparison.casey_p80&&<div style={{fontSize:'13px',color:'#8df7ff',marginBottom:'2px'}}>QCRA P80: <b>{caseyComparison.casey_p80}</b></div>}
              {caseyComparison.casey_p50&&<div style={{fontSize:'13px',color:'#8df7ff',marginBottom:'2px'}}>P50: <b>{caseyComparison.casey_p50}</b></div>}
              <div style={{fontSize:'10px',color:'#10b981',marginTop:'4px',fontWeight:'700'}}>Independent position</div>
            </div>
          </div>
        </div>
      )}

      <div className="challengeSection">
        <div className="challengeSectionHead attack"><span className="csh-num">Q</span> Board challenge questions</div>
        {questions.slice(0,6).map((x,i)=><div className="challengeAttack" key={i}><span className="caNum">{i+1}</span><span className="caQ">{x}</span></div>)}
      </div>

      <div className="challengeSection">
        <div className="challengeSectionHead next">Required next actions</div>
        {next.slice(0,6).map((x,i)=><div className="challengeNext" key={i}><span>{i+1}</span><span>{x}</span></div>)}
      </div>
    </div>
  );
}

function HolyGrailRuntime({ model, scenario, generate, runShock }) {
  const [lastFired, setLastFired] = React.useState(null);
  const controls = [
    ['signalling_slip', 'What if signalling slips 4 months?',
      'Simulates a late systems integration. CASEY recalculates the delivery tail, drops confidence and rewrites the board attack chain to focus on commissioning readiness.'],
    ['procurement_gap', 'What if procurement evidence is missing?',
      'Simulates an evidence gap on a critical package. P80/P90 exposure rises, confidence falls. CASEY identifies which board questions this unlocks.'],
    ['reserve_cut', 'What if contingency is cut to hit budget?',
      'Simulates a political cost-cut. Headline P50 improves but board defensibility weakens — CASEY flags the hidden residual risk the cut creates.'],
    ['operator_delay', 'What if operator acceptance moves late?',
      'Simulates a handover slip. The governing constraint moves from civil delivery to commissioning readiness — a critical shift most dashboards miss.'],
    ['scope_growth', 'What if scope grows 8%?',
      'Simulates scope creep. Cost and schedule increase. CASEY re-prices reserve adequacy and updates the board approval exposure.'],
    ['political_exposure', 'What if political or funding pressure increases?',
      'Simulates external programme pressure. Risk posture rises. CASEY strengthens the evidence requirements before board approval.'],
  ];
  const fire = (id) => { setLastFired(id); runShock(id); };
  const scenarioLabels = { base:'Base', faster:'Faster', cheaper:'Cheaper', lower_risk:'Lower Risk', premium:'Premium' };
  const p50 = model?.cost_p50 || (model?.cost_p50_bn ? '$' + model.cost_p50_bn + 'B' : '—');
  const conf = model?.confidence_pct;
  const chain = (model?.causal_chain || []).join(' → ') || 'Generate a project first to see the causal chain';
  return <section className="layout two runtimePanel">
    <Card>
      <h2>Live Programme Stress Test</h2>
      <p className="big">Select a real-world risk event. CASEY recalculates cost, schedule, confidence and board posture from the live model — not from a pre-written response. This is what separates CASEY from a static dashboard.</p>
      <p style={{fontSize:'11px',color:'#64748b',marginBottom:'12px'}}>First generate a project on the Overview tab, then click any event below to stress-test it.</p>
      <div className="runtimeButtons">
        {controls.map(([id,title,sub])=>(
          <button key={id} onClick={()=>fire(id)} className={lastFired===id?'fired':''}>
            <Zap size={14}/>
            <b>{title}</b>
            <span>{sub}</span>
          </button>
        ))}
      </div>
      <h3>Run a different scenario</h3>
      <p style={{fontSize:'11px',color:'#64748b',marginBottom:'8px'}}>Each scenario is a complete recalculation — different cost, schedule, confidence, risks and board language.</p>
      <div className="runtimeScenarioRow">
        {Object.entries(scenarioLabels).map(([s,label])=>(
          <button key={s} className={s===scenario?'active':''} onClick={()=>generate(s, model?.prompt, model)}>
            {label}
          </button>
        ))}
      </div>
    </Card>
    <Card>
      <h2>Live Model State</h2>
      <p style={{fontSize:'11px',color:'#64748b',marginBottom:'12px'}}>Every stress test updates these figures from the live model. Download any export after applying a stress test to capture the changed position in your board pack.</p>
      {!model ? <div className="intakeEmpty" style={{padding:'20px'}}><div className="intakeEmptyIcon">📊</div><b>No model loaded</b><span>Generate a project on the Overview tab first.</span></div> : <>
      {lastFired && <div style={{background:'rgba(245,158,11,0.1)',border:'1px solid rgba(245,158,11,0.3)',borderRadius:'3px',padding:'10px 12px',marginBottom:'12px'}}>
        <div style={{fontSize:'9px',fontWeight:'900',letterSpacing:'.15em',color:'#f59e0b',marginBottom:'4px'}}>STRESS TEST APPLIED</div>
        <div style={{fontSize:'12px',color:'#e2e8f0',fontWeight:'700'}}>{lastFired.replace(/_/g,' ').toUpperCase()}</div>
        <div style={{fontSize:'11px',color:'#94a3b8',marginTop:'4px'}}>{model?.stress_test_note || model?.executive_shock_insight || ''}</div>
      </div>}
      {[
        ['Scenario', model?.scenario_label || scenario],
        ['P50 cost estimate', p50],
        ['Programme duration', (model?.schedule_months || model?.duration_months || '—') + ' months'],
        ['Confidence', conf !== undefined ? conf + '%' + (conf < 45 ? ' — Further assurance required' : conf < 60 ? ' — Board challenge likely' : conf < 75 ? ' — Conditionally approvable' : ' — Board-defensible') : '—'],
        ['QCRA P80 downside', model?.cost_range ? model.cost_range.split('|')[1]?.trim() || '—' : '—'],
        ['Risk posture', model?.risk || '—'],
        ['Stress test applied', lastFired ? lastFired.replace(/_/g,' ') : 'None — click a button on the left'],
      ].map(([k,v],i)=><div className="reason" key={k}><span>{i+1}</span><b>{k}:</b> {v}</div>)}
      <h3>Governing causal chain</h3>
      <p style={{fontSize:'11px',color:'#8df7ff',lineHeight:'1.5',padding:'8px',background:'rgba(141,247,255,0.05)',borderRadius:'3px',borderLeft:'2px solid #8df7ff'}}>{chain}</p>
      <div style={{marginTop:'12px',padding:'10px 12px',background:'rgba(16,185,129,0.06)',borderRadius:'3px',border:'1px solid rgba(16,185,129,0.2)'}}>
        <div style={{fontSize:'9px',fontWeight:'900',color:'#10b981',letterSpacing:'.12em',marginBottom:'6px'}}>EXPORT THE STRESSED POSITION</div>
        <div style={{fontSize:'11px',color:'#94a3b8',lineHeight:'1.5'}}>After applying a stress test, download Export Board Pack or Export Cost Workbook from the top bar. The export will contain the stressed P50, schedule and confidence — not the original values.</div>
      </div>
      </>}
    </Card>
  </section>;
}

function normalizeChatAnswer(r) {
  if (!r) return 'CASEY returned no advisor response.';
  if (typeof r.answer === 'string') return r.answer;
  if (r.answer !== undefined) return safeRender(r.answer);
  return safeRender(r);
}
class CaseyErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { error: null }; }
  static getDerivedStateFromError(error) { return { error }; }
  componentDidCatch(error, info) { console.error('CASEY UI crash guard:', error, info); }
  render() {
    if (this.state.error) {
      return <div className="app v50EliteApp"><main className="v50Console"><section className="layout one"><div className="card shockCard"><h2>CASEY UI recovered</h2><p>The interface caught a render exception instead of going blank. Refresh and re-run the same programme, or use the preset advisor buttons while the custom question guard is active.</p><pre>{safeRender(this.state.error?.message || this.state.error)}</pre></div></section></main></div>;
    }
    return this.props.children;
  }
}


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


const showcaseProjects = [
  { sector:'AI / Data Centres', region:'United States', client:'Microsoft / OpenAI reference case', title:'Microsoft AI Supercluster Expansion', icon:'AI', confidence:'Grid constrained', prompt:'Microsoft OpenAI AI supercluster expansion with 500MW hyperscale data centres, GPU procurement bottlenecks, grid interconnection, liquid cooling, transmission upgrades and accelerated 2027 delivery' },
  { sector:'AI / Data Centres', region:'Global', client:'Amazon AWS reference case', title:'AWS Global Region Expansion', icon:'AI', confidence:'Power + fibre dependency', prompt:'Amazon AWS global region expansion with sovereign cloud zones, edge data centres, fibre backbone, redundant power, energy procurement and geopolitical resilience requirements' },
  { sector:'AI / Data Centres', region:'United States / Europe', client:'Meta reference case', title:'Meta AI Compute Network', icon:'AI', confidence:'Cooling + rack density', prompt:'Meta AI compute network expansion with training clusters, hyperscale networking, custom silicon, high rack density, liquid cooling and power procurement constraints' },
  { sector:'AI / Data Centres', region:'Global', client:'Google reference case', title:'Google TPU Infrastructure', icon:'AI', confidence:'Energy integration', prompt:'Google TPU AI infrastructure programme with renewable energy integration, data centre campuses, low latency routing, cooling systems and autonomous optimisation infrastructure' },
  { sector:'AI / Data Centres', region:'United States / Middle East', client:'Oracle / Sovereign cloud reference case', title:'Oracle Sovereign AI Cloud', icon:'AI', confidence:'Sovereign resilience', prompt:'Oracle sovereign AI cloud campus rollout with sovereign hosting, cyber resilience, GPU procurement, utility interconnection, water cooling and national infrastructure constraints' },
  { sector:'AI / Data Centres', region:'United States', client:'xAI reference case', title:'xAI Compute Expansion', icon:'AI', confidence:'Utility race', prompt:'xAI large scale AI compute expansion with rapid data centre construction, power substation upgrades, GPU delivery pressure, cooling infrastructure and aggressive schedule compression' },

  { sector:'Rail / Transit', region:'United States', client:'California HSR reference case', title:'California High-Speed Rail', icon:'Rail', confidence:'Land + tunnelling risk', prompt:'California High-Speed Rail megaproject with corridor acquisition, tunnelling, utility relocation, civil packages, political scrutiny, systems integration and cost escalation exposure' },
  { sector:'Rail / Transit', region:'United Kingdom', client:'HS2 reference case', title:'HS2 High-Speed Rail', icon:'Rail', confidence:'Scope + governance pressure', prompt:'HS2 high speed rail programme with scope changes, tunnelling, stations, systems integration, political volatility, cost escalation and governance confidence challenge' },
  { sector:'Rail / Transit', region:'United States', client:'Gateway Program reference case', title:'Gateway / Hudson Tunnel', icon:'Rail', confidence:'Urban resilience', prompt:'Gateway Hudson Tunnel rail resilience programme with urban tunnelling, aging infrastructure interfaces, rail continuity, federal funding, environmental approvals and schedule uncertainty' },
  { sector:'Rail / Transit', region:'United States', client:'Brightline reference case', title:'Brightline West', icon:'Rail', confidence:'Accelerated delivery', prompt:'Brightline West high speed rail project with desert corridor construction, private finance, power integration, stations, rolling stock, systems delivery and accelerated schedule pressure' },
  { sector:'Rail / Transit', region:'Europe', client:'Rail Baltica reference case', title:'Rail Baltica', icon:'Rail', confidence:'Cross-border governance', prompt:'Rail Baltica cross border European rail programme with multi-country governance, procurement coordination, interoperability, defence mobility, stations and schedule integration risk' },
  { sector:'Rail / Transit', region:'Australia', client:'Sydney Metro reference case', title:'Sydney Metro Expansion', icon:'Rail', confidence:'Urban systems integration', prompt:'Sydney Metro expansion with tunnels, underground stations, rail systems, live city interfaces, systems integration, commissioning and public disruption constraints' },
  { sector:'Rail / Transit', region:'Canada', client:'Metrolinx reference case', title:'Ontario Line', icon:'Rail', confidence:'Interface density', prompt:'Toronto Ontario Line transit megaproject with tunnelling, station boxes, procurement packaging, utility relocation, systems coordination and urban interface risk' },
  { sector:'Rail / Transit', region:'India', client:'NHSRCL reference case', title:'Mumbai–Ahmedabad HSR', icon:'Rail', confidence:'Land + technology transfer', prompt:'Mumbai Ahmedabad High Speed Rail programme with land acquisition, viaducts, stations, Japanese technology transfer, systems integration and corridor delivery uncertainty' },

  { sector:'Mega Infrastructure', region:'Saudi Arabia', client:'NEOM reference case', title:'NEOM / The Line Mobility', icon:'Infra', confidence:'Giga-project orchestration', prompt:'NEOM The Line transit and infrastructure programme with autonomous mobility, giga project logistics, workforce scaling, utilities, modular construction, supply chain and governance complexity' },
  { sector:'Mega Infrastructure', region:'Singapore', client:'Changi reference case', title:'Changi Airport Expansion', icon:'Air', confidence:'Live operations', prompt:'Changi airport expansion with terminal construction, runway systems, baggage automation, live airport operations, passenger growth, regulatory approvals and resilience planning' },
  { sector:'Mega Infrastructure', region:'UAE', client:'Etihad Rail reference case', title:'Etihad Rail Freight Network', icon:'Rail', confidence:'Desert logistics', prompt:'Etihad Rail freight network expansion with desert civil works, freight terminals, port interfaces, signalling, regional coordination and logistics integration risk' },

  { sector:'Pharma / Life Sciences', region:'United States', client:'Eli Lilly reference case', title:'Eli Lilly GLP-1 Manufacturing', icon:'Bio', confidence:'GMP validation', prompt:'Eli Lilly obesity medicine manufacturing expansion with sterile fill finish, GMP validation, GLP-1 demand growth, cold chain logistics, regulatory approvals and accelerated commercial readiness' },
  { sector:'Pharma / Life Sciences', region:'Europe / United States', client:'Novo Nordisk reference case', title:'Novo Nordisk Capacity Expansion', icon:'Bio', confidence:'Demand surge', prompt:'Novo Nordisk biologics manufacturing capacity expansion with GLP-1 demand surge, sterile production, cold chain, validation, clean utilities and supply chain redundancy' },
  { sector:'Pharma / Life Sciences', region:'United States', client:'Moderna reference case', title:'Moderna Biosecurity Facility', icon:'Bio', confidence:'Regulatory readiness', prompt:'Moderna biosecurity and vaccine manufacturing facility with rapid response production, GMP cleanrooms, sterile commissioning, regulatory inspections and sovereign health resilience requirements' },

  { sector:'Energy / Industrial', region:'Global', client:'SMR developer reference case', title:'SMR Nuclear Rollout', icon:'Energy', confidence:'Certification gate', prompt:'Small modular reactor nuclear rollout programme with regulatory certification, containment qualification, nuclear island procurement, grid integration, safety assurance and long tail licensing risk' },
  { sector:'Energy / Industrial', region:'United States / Gulf', client:'LNG operator reference case', title:'LNG Export Terminal', icon:'Energy', confidence:'Commissioning exposure', prompt:'LNG export terminal megaproject with cryogenic systems, marine berths, long lead valves, liquefaction trains, permits, commissioning risk and weather logistics exposure' },
  { sector:'Energy / Industrial', region:'North Sea / Atlantic', client:'Offshore wind developer reference case', title:'Offshore Wind Mega-Hub', icon:'Energy', confidence:'Weather windows', prompt:'Offshore wind mega hub with turbine foundations, export cables, grid integration, installation vessel constraints, weather windows and marine logistics risk' },
  { sector:'Energy / Industrial', region:'Middle East / Australia', client:'Hydrogen developer reference case', title:'Hydrogen Export Corridor', icon:'Energy', confidence:'Immature supply chain', prompt:'Hydrogen export corridor with electrolysers, ammonia conversion, port infrastructure, energy pricing, immature supply chain, offtake uncertainty and export logistics' },
  { sector:'Energy / Industrial', region:'United States / Europe', client:'Grid operator reference case', title:'Transmission Supergrid', icon:'Energy', confidence:'Permitting bottleneck', prompt:'Transmission supergrid modernization programme with high voltage lines, substations, permitting, grid resilience, interconnection queues, supply chain delays and public opposition' },

  { sector:'Defence / National Security', region:'Australia / UK / US', client:'AUKUS reference case', title:'AUKUS Industrial Base', icon:'Defence', confidence:'Sovereign capability', prompt:'AUKUS industrial base programme with nuclear submarine shipyard scaling, workforce shortages, sovereign supply chain, certification, dockyard capacity and defence governance pressure' },
  { sector:'Defence / National Security', region:'NATO / Indo-Pacific', client:'Defence ministry reference case', title:'Missile Defence Network', icon:'Defence', confidence:'Systems interoperability', prompt:'Missile defence modernization network with radar integration, command systems, interceptor procurement, interoperability, cyber resilience, simulation evidence and supplier maturity constraints' },
  { sector:'Defence / National Security', region:'Global', client:'Defence ministry reference case', title:'Autonomous Drone Swarm', icon:'Defence', confidence:'AI assurance', prompt:'Autonomous drone swarm defence programme with AI assurance, electronic warfare survivability, battery logistics, sensor integrity, autonomy certification and mission assurance risk' },
  { sector:'Defence / National Security', region:'United States / UK / Australia', client:'Naval command reference case', title:'Naval Shipbuilding Expansion', icon:'Defence', confidence:'Workforce bottleneck', prompt:'Naval shipbuilding expansion programme with dockyard capacity, steel supply, propulsion systems, combat systems integration, workforce readiness and schedule confidence exposure' },

  { sector:'Semiconductors / Advanced Manufacturing', region:'United States', client:'TSMC reference case', title:'TSMC Arizona Fab', icon:'Fab', confidence:'Toolchain + workforce', prompt:'TSMC Arizona semiconductor fab programme with ultra clean environments, toolchain delivery, workforce maturity, water and power utilities, process node commissioning and geopolitical urgency' },
  { sector:'Semiconductors / Advanced Manufacturing', region:'United States / Europe', client:'Intel reference case', title:'Intel Fab Expansion', icon:'Fab', confidence:'Yield ramp', prompt:'Intel semiconductor fab expansion with EUV tooling, cleanroom construction, utility systems, process node complexity, supply chain constraints and commissioning readiness risk' },
  { sector:'Semiconductors / Advanced Manufacturing', region:'United States / Korea', client:'Samsung reference case', title:'Samsung Foundry Expansion', icon:'Fab', confidence:'Manufacturing maturity', prompt:'Samsung foundry expansion with advanced semiconductor manufacturing, ultra cleanroom delivery, utility redundancy, tool installation, yield ramp uncertainty and global supply chain risk' },

  { sector:'Space / Orbital Infrastructure', region:'Texas / Florida / Orbit', client:'SpaceX reference case', title:'SpaceX Starship Industrialization', icon:'Space', confidence:'Launch cadence', prompt:'SpaceX Starship industrialization programme with launch cadence scaling, orbital refueling, thermal protection, booster recovery, regulatory approvals, pad infrastructure and autonomous operations' },
  { sector:'Space / Orbital Infrastructure', region:'Lunar surface', client:'Lunar programme reference case', title:'Lunar Habitat Infrastructure', icon:'Space', confidence:'Mission survivability', prompt:'Lunar habitat infrastructure programme with life support, habitat redundancy, landing pads, nuclear power, thermal survivability, autonomous commissioning, launch windows and crew logistics' },
  { sector:'Space / Orbital Infrastructure', region:'Mars logistics', client:'Mars programme reference case', title:'Mars Cargo Logistics Network', icon:'Space', confidence:'Autonomous operations', prompt:'Mars cargo logistics network with autonomous navigation, radiation hardening, communication latency, in situ resource utilisation, cargo landers, long-cycle logistics and mission recovery constraints' },
  { sector:'Space / Orbital Infrastructure', region:'Low Earth Orbit', client:'Amazon Kuiper reference case', title:'Amazon Kuiper Constellation', icon:'Space', confidence:'Production + launch manifest', prompt:'Amazon Kuiper satellite constellation deployment with satellite production ramp, launch manifest reliability, ground stations, spectrum coordination, network rollout and orbital commissioning risk' },
  { sector:'Space / Orbital Infrastructure', region:'Low Earth Orbit', client:'AST SpaceMobile reference case', title:'AST SpaceMobile Network', icon:'Space', confidence:'Telecom integration', prompt:'AST SpaceMobile orbital telecom network with satellite deployment sequencing, telecom interoperability, antenna deployment, regulatory spectrum approvals and ground network integration risk' },
  { sector:'Space / Orbital Infrastructure', region:'Orbit', client:'Orbital compute reference case', title:'Orbital Data Centres', icon:'Space', confidence:'Radiation + servicing', prompt:'Space based orbital data centre infrastructure with launch economics, orbital cooling, autonomous servicing, radiation hardening, power redundancy, data relay and mission assurance constraints' },
  { sector:'Space / Orbital Infrastructure', region:'Cislunar', client:'Orbital servicing reference case', title:'Autonomous Orbital Servicing', icon:'Space', confidence:'Robotic reliability', prompt:'Autonomous orbital servicing platform with rendezvous proximity operations, robotic capture, refuelling, inspection payloads, debris avoidance and mission assurance governance' },
  { sector:'Space / Orbital Infrastructure', region:'Lunar surface', client:'Lunar resources reference case', title:'Lunar Resource Extraction', icon:'Space', confidence:'ISRU maturity', prompt:'Lunar resource extraction programme with autonomous mining, regolith processing, in situ resource utilisation, surface power, habitat logistics, thermal survivability and launch logistics' },
];


const TAB_GUIDE_SIMPLE = {
  overview: 'Your programme at a glance — P50 cost, schedule, confidence and board verdict. Start here.',
  cost: 'Cost estimate workbook — every CBS line item with P10/P50/P90, unit rates and what drives cost.',
  schedule: 'Schedule logic by phase — critical path, milestones and QSRA delivery date probability curve.',
  risk: 'Risk register — cause, event, impact, probability, named owner, mitigation and EMV for every risk.',
  monte: 'Monte Carlo results — QCRA cost P-curve and QSRA schedule P-curve from 18,000+ simulations. Tornado chart shows top risk drivers.',
  compare: 'Scenario trade-offs — run Faster, Cheaper, Lower Risk or Premium. Each is a complete independent recalculation.',
  delta: 'Strategic intelligence — what changed in this scenario, confidence breakdown, top decisions required.',
  mortality: 'Programme mortality — probability of cancellation or restructuring, with named historical precedents.',
  market: 'Contractor market intel — who can deliver, order book depth, supply chain lead times and unit cost benchmarks.',
  gaps: 'Evidence gaps — what is missing before this can pass any gate review or board approval. IPA and Green Book referenced.',
  paths: 'Delivery paths — three strategies for the same brief with cost, schedule, confidence and board questions.',
  assurance: 'Board challenge weapons — CASEY position vs conventional advisory. What T&T/Atkins/Jacobs will say and how to respond.',
  outputs: 'Download everything — board PDF, cost workbook, risk register, XER schedule, PPTX.',
  advisor: 'Ask anything — what-if questions rerun the model with your constraint applied and show a cost/confidence delta.',
  runtime: 'Stress test — apply real-world shocks and see P50/confidence recalculate live.',
  benchmark: '63 named real programmes that calibrated this estimate — actual cost growth, slip and failure modes.',
  causal: 'Causal chain — structural drivers of cost overrun and schedule failure in this sector.',
  method: 'How every number was calculated.',
  pricing: 'Enterprise access and advisory services.',
};
const REAL_BENCHMARKS = [
  { name:'Crossrail / Elizabeth Line', sector:'Rail / Transit', mode:'Earth', cost_bn:22.7, cost_growth_pct:88, schedule_slip_months:84, failure_mode:'Deferred systems integration — 900 open IEMs at planned opening', lesson:'Possessions and signalling must be on the critical path from day one, not treated as commissioning a', prompt:'Crossrail / Elizabeth Line real programme Rail / Transit actual outturn $22.7B +88% cost growth +84 months slip failure mode: Deferred systems integration — 900 open IEMs at planned opening' },
  { name:'HS2 Phase 1', sector:'Rail / Transit', mode:'Earth', cost_bn:44.6, cost_growth_pct:140, schedule_slip_months:36, failure_mode:'Scope growth, ground conditions, open corridor risk', lesson:'Cost-at-completion estimates grow during delivery — approving at P50 without P80 reserve is a govern', prompt:'HS2 Phase 1 real programme Rail / Transit actual outturn $44.6B +140% cost growth +36 months slip failure mode: Scope growth, ground conditions, open corridor risk' },
  { name:'Riyadh Metro', sector:'Rail / Transit', mode:'Earth', cost_bn:22.5, cost_growth_pct:12, schedule_slip_months:24, failure_mode:'Systems integration and operational readiness timeline', lesson:'International rail programmes with multiple concessions face interface risk proportional to contract', prompt:'Riyadh Metro real programme Rail / Transit actual outturn $22.5B +12% cost growth +24 months slip failure mode: Systems integration and operational readiness timeline' },
  { name:'Hinkley Point C', sector:'Nuclear / Energy', mode:'Earth', cost_bn:35.0, cost_growth_pct:94, schedule_slip_months:60, failure_mode:'FOAK EPR supply chain, first-pour concrete issues, nuclear-grade welding failure', lesson:'GDA is the real critical path — not construction. Every 6 months of GDA slip costs £1B+ in financing', prompt:'Hinkley Point C real programme Nuclear / Energy actual outturn $35.0B +94% cost growth +60 months slip failure mode: FOAK EPR supply chain, first-pour concrete issues, nuclear-grade welding failures' },
  { name:'Olkiluoto 3 (Finland)', sector:'Nuclear / Energy', mode:'Earth', cost_bn:11.0, cost_growth_pct:300, schedule_slip_months:168, failure_mode:'FOAK EPR complexity, safety system integration, regulatory hold-points', lesson:'New reactor designs have 3-5x baseline cost growth on first deployment', prompt:'Olkiluoto 3 (Finland) real programme Nuclear / Energy actual outturn $11.0B +300% cost growth +168 months slip failure mode: FOAK EPR complexity, safety system integration, regulatory hold-points' },
  { name:'Vogtle Units 3 & 4 (Georgia)', sector:'Nuclear / Energy', mode:'Earth', cost_bn:34.0, cost_growth_pct:113, schedule_slip_months:84, failure_mode:'FOAK AP1000 design changes, contractor performance, qualified labour shortage', lesson:'Fixed-price EPC contracts on nuclear FOAK do not transfer risk — they transfer insolvency', prompt:'Vogtle Units 3 & 4 (Georgia) real programme Nuclear / Energy actual outturn $34.0B +113% cost growth +84 months slip failure mode: FOAK AP1000 design changes, contractor performance, qualified labour shortage' },
  { name:'Ajax Armoured Vehicles (UK)', sector:'Defence / Secure Infrastructure', mode:'Earth', cost_bn:5.5, cost_growth_pct:57, schedule_slip_months:120, failure_mode:'EMC/vibration issues, crew safety, training system integration — none on critica', lesson:'Operational acceptance is the real programme gate, not platform delivery', prompt:'Ajax Armoured Vehicles (UK) real programme Defence / Secure Infrastructure actual outturn $5.5B +57% cost growth +120 months slip failure mode: EMC/vibration issues, crew safety, training system integration — none on critical path' },
  { name:'Watchkeeper UAV Programme', sector:'Defence / Secure Infrastructure', mode:'Earth', cost_bn:1.3, cost_growth_pct:130, schedule_slip_months:120, failure_mode:'Civil airspace certification never achieved — airworthiness not a delivery const', lesson:'Regulatory acceptance must be on the master critical path from day one', prompt:'Watchkeeper UAV Programme real programme Defence / Secure Infrastructure actual outturn $1.3B +130% cost growth +120 months slip failure mode: Civil airspace certification never achieved — airworthiness not a delivery constraint' },
  { name:'F-35 Joint Strike Fighter', sector:'Defence / Secure Infrastructure', mode:'Earth', cost_bn:412.0, cost_growth_pct:68, schedule_slip_months:96, failure_mode:'Software integration complexity, concurrent development and production', lesson:'Software-intensive defence programmes have 3-5x baseline schedule assumptions', prompt:'F-35 Joint Strike Fighter real programme Defence / Secure Infrastructure actual outturn $412.0B +68% cost growth +96 months slip failure mode: Software integration complexity, concurrent development and production' },
  { name:'Microsoft Azure UK South (Slough campus)', sector:'Digital Infrastructure / Hyperscale Data Centre', mode:'Earth', cost_bn:3.2, cost_growth_pct:15, schedule_slip_months:18, failure_mode:'Grid connection delay, DNO queue, cooling commissioning', lesson:'Grid connection must be a signed agreement, not a queue position — energisation delays are now syste', prompt:'Microsoft Azure UK South (Slough campus) real programme Digital Infrastructure / Hyperscale Data Centre actual outturn $3.2B +15% cost growth +18 months slip failure mode: Grid connection delay, DNO queue, cooling commissioning' },
  { name:'Amazon AWS Dublin Campus', sector:'Digital Infrastructure / Hyperscale Data Centre', mode:'Earth', cost_bn:4.2, cost_growth_pct:20, schedule_slip_months:14, failure_mode:'Planning opposition, grid capacity, water usage consent', lesson:'Data centres in water-stressed regions face novel consent constraints not in traditional risk regist', prompt:'Amazon AWS Dublin Campus real programme Digital Infrastructure / Hyperscale Data Centre actual outturn $4.2B +20% cost growth +14 months slip failure mode: Planning opposition, grid capacity, water usage consent' },
  { name:'AstraZeneca Macclesfield Expansion', sector:'Life Sciences / Biologics Manufacturing', mode:'Earth', cost_bn:1.2, cost_growth_pct:30, schedule_slip_months:24, failure_mode:'Validation deferred post-construction, clean utility qualification delay', lesson:'CQV is a programme deliverable — not a post-construction activity', prompt:'AstraZeneca Macclesfield Expansion real programme Life Sciences / Biologics Manufacturing actual outturn $1.2B +30% cost growth +24 months slip failure mode: Validation deferred post-construction, clean utility qualification delay' },
  { name:'Pfizer Ringaskiddy Ireland', sector:'Life Sciences / Biologics Manufacturing', mode:'Earth', cost_bn:1.5, cost_growth_pct:18, schedule_slip_months:18, failure_mode:'Regulatory submission delayed by CMC dossier readiness', lesson:'Regulatory submission is the revenue gate — it must be on the programme critical path', prompt:'Pfizer Ringaskiddy Ireland real programme Life Sciences / Biologics Manufacturing actual outturn $1.5B +18% cost growth +18 months slip failure mode: Regulatory submission delayed by CMC dossier readiness' },
  { name:'Intel Ohio Fab (Planned)', sector:'Semiconductor / Advanced Manufacturing', mode:'Earth', cost_bn:28.0, cost_growth_pct:0, schedule_slip_months:36, failure_mode:'Workforce shortage, UPW complexity, tool delivery slippage', lesson:'Tool delivery sequences must be confirmed orders — OEM intent letters are not programme commitments', prompt:'Intel Ohio Fab (Planned) real programme Semiconductor / Advanced Manufacturing actual outturn $28.0B +0% cost growth +36 months slip failure mode: Workforce shortage, UPW complexity, tool delivery slippage' },
  { name:'TSMC Arizona Fab', sector:'Semiconductor / Advanced Manufacturing', mode:'Earth', cost_bn:40.0, cost_growth_pct:35, schedule_slip_months:30, failure_mode:'Specialised workforce unavailable locally, tool delivery, UPW systems', lesson:'Fab yields in new geographies are systematically below initial projections', prompt:'TSMC Arizona Fab real programme Semiconductor / Advanced Manufacturing actual outturn $40.0B +35% cost growth +30 months slip failure mode: Specialised workforce unavailable locally, tool delivery, UPW systems' },
  { name:'Samsung Taylor Texas Fab', sector:'Semiconductor / Advanced Manufacturing', mode:'Earth', cost_bn:17.0, cost_growth_pct:20, schedule_slip_months:24, failure_mode:'Market demand timing, workforce availability, tool delivery', lesson:'Semiconductor fabs require 5-8 year horizon planning — market timing risk is structural', prompt:'Samsung Taylor Texas Fab real programme Semiconductor / Advanced Manufacturing actual outturn $17.0B +20% cost growth +24 months slip failure mode: Market demand timing, workforce availability, tool delivery' },
  { name:'Britishvolt (Failed)', sector:'Battery / Gigafactory', mode:'Earth', cost_bn:3.8, cost_growth_pct:0, schedule_slip_months:0, failure_mode:'Grid connection, cell chemistry qualification, BMS supply chain — all unconfirme', lesson:'A gigafactory without a confirmed grid connection and qualified cell chemistry is a building, not a ', prompt:'Britishvolt (Failed) real programme Battery / Gigafactory actual outturn $3.8B +0% cost growth +0 months slip failure mode: Grid connection, cell chemistry qualification, BMS supply chain — all unconfirmed at commitment' },
  { name:'Northvolt Ett (Sweden)', sector:'Battery / Gigafactory', mode:'Earth', cost_bn:8.0, cost_growth_pct:40, schedule_slip_months:36, failure_mode:'Yield ramp 16x below target — 1GWh/year achieved vs 16GWh target', lesson:'Battery yield ramp is the board metric — building capacity for an unqualified product is not product', prompt:'Northvolt Ett (Sweden) real programme Battery / Gigafactory actual outturn $8.0B +40% cost growth +36 months slip failure mode: Yield ramp 16x below target — 1GWh/year achieved vs 16GWh target' },
  { name:'Hornsea 2 Offshore Wind Farm', sector:'Energy / Utilities', mode:'Earth', cost_bn:3.0, cost_growth_pct:20, schedule_slip_months:18, failure_mode:'Grid connection 18 months late — DNO queue backlog', lesson:'Grid connection queue position is not an energisation date — it is a forecast', prompt:'Hornsea 2 Offshore Wind Farm real programme Energy / Utilities actual outturn $3.0B +20% cost growth +18 months slip failure mode: Grid connection 18 months late — DNO queue backlog' },
  { name:'Neart na Gaoithe Offshore Wind (Scotland)', sector:'Energy / Utilities', mode:'Earth', cost_bn:3.5, cost_growth_pct:25, schedule_slip_months:48, failure_mode:'Aviation radar objection known at planning — not treated as programme constraint', lesson:'Third-party consent risks must be treated as critical path items at project inception', prompt:'Neart na Gaoithe Offshore Wind (Scotland) real programme Energy / Utilities actual outturn $3.5B +25% cost growth +48 months slip failure mode: Aviation radar objection known at planning — not treated as programme constraint' },
  { name:'Hinkley Point C Nuclear (Energy angle)', sector:'Energy / Utilities', mode:'Earth', cost_bn:35.0, cost_growth_pct:94, schedule_slip_months:60, failure_mode:'FOAK construction, supply chain, regulatory timeline', lesson:'Nuclear baseload power has a 50-year asset life — the approval case must reflect lifetime value not ', prompt:'Hinkley Point C Nuclear (Energy angle) real programme Energy / Utilities actual outturn $35.0B +94% cost growth +60 months slip failure mode: FOAK construction, supply chain, regulatory timeline' },
  { name:'Thames Water AMP7 Capital Programme', sector:'Water / Environmental Infrastructure', mode:'Earth', cost_bn:3.7, cost_growth_pct:40, schedule_slip_months:24, failure_mode:'Procurement and supply chain capacity, site access — 40% below delivery target', lesson:'Utility capital programmes require contracted supply chain at programme start — not competitive proc', prompt:'Thames Water AMP7 Capital Programme real programme Water / Environmental Infrastructure actual outturn $3.7B +40% cost growth +24 months slip failure mode: Procurement and supply chain capacity, site access — 40% below delivery target' },
  { name:'SMETS2 Smart Meter Rollout (UK)', sector:'Water / Environmental Infrastructure', mode:'Earth', cost_bn:13.9, cost_growth_pct:90, schedule_slip_months:60, failure_mode:'Comms infrastructure complexity, back-office platform readiness, MDU access', lesson:'Smart meter programmes fail at the back-office integration layer, not at the physical meter', prompt:'SMETS2 Smart Meter Rollout (UK) real programme Water / Environmental Infrastructure actual outturn $13.9B +90% cost growth +60 months slip failure mode: Comms infrastructure complexity, back-office platform readiness, MDU access' },
  { name:'NBN Co (Australia) Smart Infrastructure', sector:'Water / Environmental Infrastructure', mode:'Earth', cost_bn:51.0, cost_growth_pct:985, schedule_slip_months:72, failure_mode:'Engineering complexity, copper network assumptions, multi-technology mix', lesson:'Utility rollout programmes in mixed urban/rural geographies have 3-10x baseline complexity assumptio', prompt:'NBN Co (Australia) Smart Infrastructure real programme Water / Environmental Infrastructure actual outturn $51.0B +985% cost growth +72 months slip failure mode: Engineering complexity, copper network assumptions, multi-technology mix' },
  { name:'Chevron Gorgon LNG (Australia)', sector:'Oil & Gas / Process Infrastructure', mode:'Earth', cost_bn:54.0, cost_growth_pct:54, schedule_slip_months:36, failure_mode:'Brownfield interface complexity, remote logistics, HAZOP findings', lesson:'Brownfield LNG interface complexity is systematically underestimated at project inception', prompt:'Chevron Gorgon LNG (Australia) real programme Oil & Gas / Process Infrastructure actual outturn $54.0B +54% cost growth +36 months slip failure mode: Brownfield interface complexity, remote logistics, HAZOP findings' },
  { name:'Shell Prelude FLNG (Australia)', sector:'Oil & Gas / Process Infrastructure', mode:'Earth', cost_bn:12.0, cost_growth_pct:50, schedule_slip_months:60, failure_mode:'FOAK floating LNG technology — never achieved nameplate capacity', lesson:'FOAK floating process technology has 5x cost growth assumption vs comparable fixed infrastructure', prompt:'Shell Prelude FLNG (Australia) real programme Oil & Gas / Process Infrastructure actual outturn $12.0B +50% cost growth +60 months slip failure mode: FOAK floating LNG technology — never achieved nameplate capacity' },
  { name:'Cobre Panama Copper Mine (First Quantum)', sector:'Mining / Metals Infrastructure', mode:'Earth', cost_bn:10.0, cost_growth_pct:100, schedule_slip_months:0, failure_mode:'Shut by government order — community licence-to-operate not treated as board gat', lesson:'$10B built and operating, then shut. Community opposition must be a board approval gate, not a stake', prompt:'Cobre Panama Copper Mine (First Quantum) real programme Mining / Metals Infrastructure actual outturn $10.0B +100% cost growth +0 months slip failure mode: Shut by government order — community licence-to-operate not treated as board gate' },
  { name:'Roy Hill Iron Ore Mine (Australia)', sector:'Mining / Metals Infrastructure', mode:'Earth', cost_bn:10.0, cost_growth_pct:20, schedule_slip_months:24, failure_mode:'Rail and port logistics, processing plant yield ramp', lesson:'Mining logistics corridors require the same programme rigour as the mine itself', prompt:'Roy Hill Iron Ore Mine (Australia) real programme Mining / Metals Infrastructure actual outturn $10.0B +20% cost growth +24 months slip failure mode: Rail and port logistics, processing plant yield ramp' },
  { name:'Heathrow Terminal 5', sector:'Airport / Aviation', mode:'Earth', cost_bn:4.3, cost_growth_pct:5, schedule_slip_months:0, failure_mode:'34,000 bags lost on day 1 — IT/baggage integration not a programme deliverable', lesson:'Construction on time and budget is not success — ORAT must be on the master critical path', prompt:'Heathrow Terminal 5 real programme Airport / Aviation actual outturn $4.3B +5% cost growth +0 months slip failure mode: 34,000 bags lost on day 1 — IT/baggage integration not a programme deliverable' },
  { name:'Berlin Brandenburg Airport', sector:'Airport / Aviation', mode:'Earth', cost_bn:7.3, cost_growth_pct:363, schedule_slip_months:108, failure_mode:'Fire safety integration, IT, regulatory approval — all post-construction', lesson:'Airport safety and regulatory approval is the opening gate — not construction practical completion', prompt:'Berlin Brandenburg Airport real programme Airport / Aviation actual outturn $7.3B +363% cost growth +108 months slip failure mode: Fire safety integration, IT, regulatory approval — all post-construction' },
  { name:'Royal Liverpool Hospital', sector:'Healthcare / Hospital Infrastructure', mode:'Earth', cost_bn:0.8, cost_growth_pct:80, schedule_slip_months:60, failure_mode:'Structural defects, infection-control compliance, PFI contractor insolvency', lesson:'PFI construction risk transfer does not transfer commissioning and occupation risk', prompt:'Royal Liverpool Hospital real programme Healthcare / Hospital Infrastructure actual outturn $0.8B +80% cost growth +60 months slip failure mode: Structural defects, infection-control compliance, PFI contractor insolvency' },
  { name:'New Royal Adelaide Hospital', sector:'Healthcare / Hospital Infrastructure', mode:'Earth', cost_bn:2.3, cost_growth_pct:60, schedule_slip_months:24, failure_mode:'Clinical commissioning not on master schedule, operational transition not contra', lesson:'Clinical commissioning is a 12-18 month programme requiring a dedicated team and critical path', prompt:'New Royal Adelaide Hospital real programme Healthcare / Hospital Infrastructure actual outturn $2.3B +60% cost growth +24 months slip failure mode: Clinical commissioning not on master schedule, operational transition not contracted' },
  { name:'A303 Stonehenge Tunnel', sector:'Roads / Highways Infrastructure', mode:'Earth', cost_bn:2.1, cost_growth_pct:50, schedule_slip_months:36, failure_mode:'UNESCO/DCO legal challenge post-contract award', lesson:'Third-party consent risks that are known but unresolved at contract award transfer to the client', prompt:'A303 Stonehenge Tunnel real programme Roads / Highways Infrastructure actual outturn $2.1B +50% cost growth +36 months slip failure mode: UNESCO/DCO legal challenge post-contract award' },
  { name:'A14 Cambridge to Huntingdon', sector:'Roads / Highways Infrastructure', mode:'Earth', cost_bn:1.5, cost_growth_pct:0, schedule_slip_months:0, failure_mode:'Utility diversions were the critical path for 60% of works', lesson:'Utility diversion timelines are systematically underestimated — third-party access is not in the con', prompt:'A14 Cambridge to Huntingdon real programme Roads / Highways Infrastructure actual outturn $1.5B +0% cost growth +0 months slip failure mode: Utility diversions were the critical path for 60% of works' },
  { name:'Felixstowe South Quay Extension', sector:'Ports / Marine Infrastructure', mode:'Earth', cost_bn:0.4, cost_growth_pct:15, schedule_slip_months:12, failure_mode:'Marine ground conditions, operational cutover constraints', lesson:'Port redevelopments require contingency for dredging ground conditions — seabed assumptions drive P8', prompt:'Felixstowe South Quay Extension real programme Ports / Marine Infrastructure actual outturn $0.4B +15% cost growth +12 months slip failure mode: Marine ground conditions, operational cutover constraints' },
  { name:'London Gateway Phase 2', sector:'Ports / Marine Infrastructure', mode:'Earth', cost_bn:1.8, cost_growth_pct:20, schedule_slip_months:18, failure_mode:'Terminal IT/OT integration 18 months late — not in EPC contract boundary', lesson:'Terminal operating systems are the ports critical path at commissioning — not the quay wall', prompt:'London Gateway Phase 2 real programme Ports / Marine Infrastructure actual outturn $1.8B +20% cost growth +18 months slip failure mode: Terminal IT/OT integration 18 months late — not in EPC contract boundary' },
  { name:'BT Openreach FTTP Rollout (UK)', sector:'Telecoms / Digital Infrastructure', mode:'Earth', cost_bn:15.0, cost_growth_pct:50, schedule_slip_months:36, failure_mode:'Wayleave complexity in MDUs and dense urban areas — 2+ years behind target', lesson:'Wayleave acquisition is the critical path for FTTP — not network build', prompt:'BT Openreach FTTP Rollout (UK) real programme Telecoms / Digital Infrastructure actual outturn $15.0B +50% cost growth +36 months slip failure mode: Wayleave complexity in MDUs and dense urban areas — 2+ years behind target' },
  { name:'NBN Co Multi-Technology Mix (Australia)', sector:'Telecoms / Digital Infrastructure', mode:'Earth', cost_bn:51.0, cost_growth_pct:985, schedule_slip_months:72, failure_mode:'Multi-technology complexity, copper network assumptions, contractor performance', lesson:'National broadband rollouts in mixed geographies require 5-10x baseline cost assumptions', prompt:'NBN Co Multi-Technology Mix (Australia) real programme Telecoms / Digital Infrastructure actual outturn $51.0B +985% cost growth +72 months slip failure mode: Multi-technology complexity, copper network assumptions, contractor performance' },
  { name:'James Webb Space Telescope (JWST)', sector:'Space / Mission Assurance', mode:'Space', cost_bn:10.0, cost_growth_pct:1400, schedule_slip_months:168, failure_mode:'Systems integration complexity, cryogenic testing failures, scope growth visible', lesson:'FOAK space systems have 14-15x baseline cost growth assumptions — qualification must be on the criti', prompt:'James Webb Space Telescope (JWST) real programme Space / Mission Assurance actual outturn $10.0B +1400% cost growth +168 months slip failure mode: Systems integration complexity, cryogenic testing failures, scope growth visible early' },
  { name:'Artemis / SLS Programme', sector:'Space / Mission Assurance', mode:'Space', cost_bn:93.0, cost_growth_pct:200, schedule_slip_months:60, failure_mode:'Fixed-price Boeing contract removed schedule incentives, propulsion complexity', lesson:'Fixed-price contracting on FOAK space systems transfers insolvency risk, not schedule risk', prompt:'Artemis / SLS Programme real programme Space / Mission Assurance actual outturn $93.0B +200% cost growth +60 months slip failure mode: Fixed-price Boeing contract removed schedule incentives, propulsion complexity' },
  { name:'OneWeb Satellite Constellation', sector:'Space / Mission Assurance', mode:'Space', cost_bn:3.4, cost_growth_pct:70, schedule_slip_months:36, failure_mode:'Bankruptcy — launch cadence, ground segment, customer revenue all optimistic', lesson:'Satellite constellation business cases require contracted anchor customers before launch commitment', prompt:'OneWeb Satellite Constellation real programme Space / Mission Assurance actual outturn $3.4B +70% cost growth +36 months slip failure mode: Bankruptcy — launch cadence, ground segment, customer revenue all optimistic' },
  { name:'Iridium NEXT Constellation', sector:'Space / Mission Assurance', mode:'Space', cost_bn:3.0, cost_growth_pct:0, schedule_slip_months:0, failure_mode:'Managed successfully — named launch provider, contracted cadence, anchor custome', lesson:'Successful constellation reference: contracted launch, proven bus, anchor customer base from day 1', prompt:'Iridium NEXT Constellation real programme Space / Mission Assurance actual outturn $3.0B +0% cost growth +0 months slip failure mode: Managed successfully — named launch provider, contracted cadence, anchor customers' },
  { name:'Lunar Gateway (Planned)', sector:'Space / Mission Assurance', mode:'Space', cost_bn:40.0, cost_growth_pct:30, schedule_slip_months:36, failure_mode:'International partner coordination, launch cadence, FOAK life support', lesson:'Cislunar infrastructure requires autonomous recovery capability — Earth-based contingency is a 3-day', prompt:'Lunar Gateway (Planned) real programme Space / Mission Assurance actual outturn $40.0B +30% cost growth +36 months slip failure mode: International partner coordination, launch cadence, FOAK life support' },
  { name:'Mars InSight Mission', sector:'Space / Mission Assurance', mode:'Space', cost_bn:0.83, cost_growth_pct:25, schedule_slip_months:24, failure_mode:'Heat probe failed to penetrate Martian soil — regolith properties not in design ', lesson:'Mars surface properties require margin for FOAK geophysical assumptions', prompt:'Mars InSight Mission real programme Space / Mission Assurance actual outturn $0.83B +25% cost growth +24 months slip failure mode: Heat probe failed to penetrate Martian soil — regolith properties not in design basis' },
  { name:'Tottenham Hotspur Stadium', sector:'Stadia / Events Infrastructure', mode:'Earth', cost_bn:1.2, cost_growth_pct:25, schedule_slip_months:9, failure_mode:'Retractable pitch mechanism, FA inspection, safety certification delay', lesson:'Event-deadline driven construction compresses commissioning — safety certification is the opening ga', prompt:'Tottenham Hotspur Stadium real programme Stadia / Events Infrastructure actual outturn $1.2B +25% cost growth +9 months slip failure mode: Retractable pitch mechanism, FA inspection, safety certification delay' },
  { name:'Wembley Stadium Redevelopment', sector:'Stadia / Events Infrastructure', mode:'Earth', cost_bn:0.8, cost_growth_pct:40, schedule_slip_months:18, failure_mode:'Steelwork fabrication, contractor disputes, safety system integration', lesson:'Stadium arch and signature structural elements carry 2-3x contingency assumption', prompt:'Wembley Stadium Redevelopment real programme Stadia / Events Infrastructure actual outturn $0.8B +40% cost growth +18 months slip failure mode: Steelwork fabrication, contractor disputes, safety system integration' },
  { name:'Riyadh Metro (Saudi Arabia)', sector:'Rail / Transit', mode:'Earth', cost_bn:22.5, cost_growth_pct:12, schedule_slip_months:24, failure_mode:'Systems integration and operational readiness timeline across 6 concessions', lesson:'Multi-concession metro programmes require a single systems integrator with contractual authority ove', prompt:'Riyadh Metro (Saudi Arabia) real programme Rail / Transit actual outturn $22.5B +12% cost growth +24 months slip failure mode: Systems integration and operational readiness timeline across 6 concessions' },
  { name:'California High Speed Rail (USA)', sector:'Rail / Transit', mode:'Earth', cost_bn:128.0, cost_growth_pct:1000, schedule_slip_months:180, failure_mode:'Environmental review, land acquisition, design changes — NEPA timeline structura', lesson:'US rail mega-projects require NEPA completion before cost can be baselined — pre-NEPA estimates are ', prompt:'California High Speed Rail (USA) real programme Rail / Transit actual outturn $128.0B +1000% cost growth +180 months slip failure mode: Environmental review, land acquisition, design changes — NEPA timeline structural constraint' },
  { name:'Sydney Metro Northwest', sector:'Rail / Transit', mode:'Earth', cost_bn:8.3, cost_growth_pct:5, schedule_slip_months:0, failure_mode:'Successfully delivered — TBM tunnelling, systems integration on schedule', lesson:'Strong project reference: alliance contract model, TBM tunnelling, single systems integrator', prompt:'Sydney Metro Northwest real programme Rail / Transit actual outturn $8.3B +5% cost growth +0 months slip failure mode: Successfully delivered — TBM tunnelling, systems integration on schedule' },
  { name:'Grand Paris Express (France)', sector:'Rail / Transit', mode:'Earth', cost_bn:36.0, cost_growth_pct:45, schedule_slip_months:48, failure_mode:'Ground conditions, geology variability, post-COVID procurement inflation', lesson:'Paris basin geology is more complex than initial surveys indicated — ground risk reserve must reflec', prompt:'Grand Paris Express (France) real programme Rail / Transit actual outturn $36.0B +45% cost growth +48 months slip failure mode: Ground conditions, geology variability, post-COVID procurement inflation' },
  { name:'NEOM THE LINE Power Infrastructure', sector:'Energy / Power Infrastructure', mode:'Earth', cost_bn:500.0, cost_growth_pct:0, schedule_slip_months:0, failure_mode:'FOAK megacity — no comparable. Technology readiness of autonomous systems is the', lesson:'No reference class exists for THE LINE. Apply maximum OBA and require independent technical review o', prompt:'NEOM THE LINE Power Infrastructure real programme Energy / Power Infrastructure actual outturn $500.0B +0% cost growth +0 months slip failure mode: FOAK megacity — no comparable. Technology readiness of autonomous systems is the primary risk.' },
  { name:'Snowy 2.0 Pumped Hydro (Australia)', sector:'Energy / Power Infrastructure', mode:'Earth', cost_bn:12.0, cost_growth_pct:233, schedule_slip_months:60, failure_mode:'TBM breakdown, ground conditions, geological fault — 3.3km TBM stuck for 14 mont', lesson:'Deep underground works in complex geology — apply 3-5x TBM programme contingency', prompt:'Snowy 2.0 Pumped Hydro (Australia) real programme Energy / Power Infrastructure actual outturn $12.0B +233% cost growth +60 months slip failure mode: TBM breakdown, ground conditions, geological fault — 3.3km TBM stuck for 14 months' },
  { name:'Barakah Nuclear Power (UAE)', sector:'Nuclear / Regulated Generation', mode:'Earth', cost_bn:32.4, cost_growth_pct:62, schedule_slip_months:72, failure_mode:'Regulatory approval timeline, ENEC/IAEA safety case, operational licensing', lesson:'First nuclear plant in the Arab world — regulatory approval timeline was the real critical path, not', prompt:'Barakah Nuclear Power (UAE) real programme Nuclear / Regulated Generation actual outturn $32.4B +62% cost growth +72 months slip failure mode: Regulatory approval timeline, ENEC/IAEA safety case, operational licensing' },
  { name:'Gordie Howe Bridge (Canada-USA)', sector:'Roads / Highways Infrastructure', mode:'Earth', cost_bn:5.7, cost_growth_pct:90, schedule_slip_months:24, failure_mode:'Bi-national procurement complexity, COVID, steel fabrication delays', lesson:'Cross-border infrastructure requires harmonised procurement rules — different national standards add', prompt:'Gordie Howe Bridge (Canada-USA) real programme Roads / Highways Infrastructure actual outturn $5.7B +90% cost growth +24 months slip failure mode: Bi-national procurement complexity, COVID, steel fabrication delays' },
  { name:'Desalination Plant Jubail II (Saudi Arabia)', sector:'Water / Environmental Infrastructure', mode:'Earth', cost_bn:1.4, cost_growth_pct:15, schedule_slip_months:12, failure_mode:'Process performance at extreme ambient temperature — membrane degradation', lesson:'Middle East desalination must be designed for 50°C+ ambient — standard membrane specifications are i', prompt:'Desalination Plant Jubail II (Saudi Arabia) real programme Water / Environmental Infrastructure actual outturn $1.4B +15% cost growth +12 months slip failure mode: Process performance at extreme ambient temperature — membrane degradation' },
  { name:'Melbourne Water Smart Meter Rollout', sector:'Water / Environmental Infrastructure', mode:'Earth', cost_bn:0.6, cost_growth_pct:25, schedule_slip_months:18, failure_mode:'Back-office data platform readiness, meter reading system integration', lesson:'Smart meter rollouts fail at the data layer — field installation is the easy part', prompt:'Melbourne Water Smart Meter Rollout real programme Water / Environmental Infrastructure actual outturn $0.6B +25% cost growth +18 months slip failure mode: Back-office data platform readiness, meter reading system integration' },
  { name:'Kashagan Phase 1 (Kazakhstan)', sector:'Oil & Gas / Process Infrastructure', mode:'Earth', cost_bn:50.0, cost_growth_pct:400, schedule_slip_months:120, failure_mode:'H2S corrosion — pipeline design failed at commissioning, 3-year restart delay', lesson:'Sour gas processing requires independent material qualification — no deviation from specification pe', prompt:'Kashagan Phase 1 (Kazakhstan) real programme Oil & Gas / Process Infrastructure actual outturn $50.0B +400% cost growth +120 months slip failure mode: H2S corrosion — pipeline design failed at commissioning, 3-year restart delay' },
  { name:'Ichthys LNG (Australia)', sector:'Oil & Gas / Process Infrastructure', mode:'Earth', cost_bn:45.0, cost_growth_pct:50, schedule_slip_months:24, failure_mode:'Module fabrication, labour costs, commissioning complexity', lesson:'LNG final cost is determined by module fabrication quality and commissioning duration — not field de', prompt:'Ichthys LNG (Australia) real programme Oil & Gas / Process Infrastructure actual outturn $45.0B +50% cost growth +24 months slip failure mode: Module fabrication, labour costs, commissioning complexity' },
  { name:'Oyu Tolgoi Underground Mine (Mongolia)', sector:'Mining / Metals Infrastructure', mode:'Earth', cost_bn:7.0, cost_growth_pct:60, schedule_slip_months:48, failure_mode:'Ground conditions, geotechnical complexity, caveback — production delayed', lesson:'Block cave mining in complex ground requires geotechnical margin — cave propagation cannot be accele', prompt:'Oyu Tolgoi Underground Mine (Mongolia) real programme Mining / Metals Infrastructure actual outturn $7.0B +60% cost growth +48 months slip failure mode: Ground conditions, geotechnical complexity, caveback — production delayed' },
  { name:'Jansen Potash Mine (Canada)', sector:'Mining / Metals Infrastructure', mode:'Earth', cost_bn:5.7, cost_growth_pct:0, schedule_slip_months:0, failure_mode:'On schedule — strong project controls, single-owner BHP, definitive feasibility', lesson:'Single-owner mega-mine with completed definitive feasibility study and no joint venture complexity —', prompt:'Jansen Potash Mine (Canada) real programme Mining / Metals Infrastructure actual outturn $5.7B +0% cost growth +0 months slip failure mode: On schedule — strong project controls, single-owner BHP, definitive feasibility' },
  { name:'AUKUS Submarine Programme (Australia/UK/USA)', sector:'Defence / Secure Infrastructure', mode:'Earth', cost_bn:268.0, cost_growth_pct:0, schedule_slip_months:0, failure_mode:'FOAK nuclear-powered submarine in Australia — no comparable. Workforce, regulato', lesson:'No reference class exists for AUKUS — it is simultaneously a FOAK submarine programme, FOAK nuclear ', prompt:'AUKUS Submarine Programme (Australia/UK/USA) real programme Defence / Secure Infrastructure actual outturn $268.0B +0% cost growth +0 months slip failure mode: FOAK nuclear-powered submarine in Australia — no comparable. Workforce, regulatory, industrial base all new.' },
  { name:'Chandrayaan-3 (India Lunar)', sector:'Space / Mission Assurance', mode:'Space', cost_bn:0.075, cost_growth_pct:0, schedule_slip_months:0, failure_mode:'Chandrayaan-2 lander failed — software bug in braking sequence. Chandrayaan-3 co', lesson:'Lunar landing requires exhaustive failure mode simulation — Chandrayaan-3 cost 10x less than Apollo ', prompt:'Chandrayaan-3 (India Lunar) real programme Space / Mission Assurance actual outturn $0.075B +0% cost growth +0 months slip failure mode: Chandrayaan-2 lander failed — software bug in braking sequence. Chandrayaan-3 corrected and succeeded.' },
  { name:'Starlink Constellation (SpaceX)', sector:'Space / Mission Assurance', mode:'Space', cost_bn:30.0, cost_growth_pct:0, schedule_slip_months:0, failure_mode:'Successfully scaled — reusable launch, vertical integration, iterative design', lesson:'Vertical integration (own launch + own satellite) is the only structure that achieves constellation ', prompt:'Starlink Constellation (SpaceX) real programme Space / Mission Assurance actual outturn $30.0B +0% cost growth +0 months slip failure mode: Successfully scaled — reusable launch, vertical integration, iterative design' }
];


const showcaseSectors = ['All', ...Array.from(new Set(showcaseProjects.map(p => p.sector)))];




function parseMoneyLocal(v) {
  if (v === undefined || v === null) return 0;
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
  const s = String(v).replace(/[$,£€]/g,'').trim().toUpperCase();
  const n = parseFloat(s.replace(/[^0-9.-]/g,''));
  if (!Number.isFinite(n)) return 0;
  if (s.includes('T')) return n * 1000;
  if (s.includes('M')) return n / 1000;
  return n;
}
function moneyLocal(n) { return n >= 1000 ? `$${(n/1000).toFixed(1)}T` : n >= 1 ? `$${n.toFixed(1)}B` : `$${Math.round(n*1000)}M`; }

function fmt(v) {
  if (v === undefined || v === null || v === '') return '—';
  if (typeof v === 'string') return v;
  return v >= 1000 ? `$${(v / 1000).toFixed(1)}T` : v >= 1 ? `$${v.toFixed(1)}B` : `$${(v * 1000).toFixed(0)}M`;
}
async function post(path, body) {
  const r = await apiFetch(path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  if (!r.ok) throw new Error(await r.text());
  return r.json();
}
async function get(path) {
  const r = await apiFetch(path);
  if (!r.ok) throw new Error(await r.text());
  return r.json();
}
async function download(path, model, name, setExportingLabel) {
    if (!isAdminUser && demoDownloads >= 1) {
      alert('Demo export limit reached. Contact us for full access at controlorbit.com');
      return;
    }
    if (!isAdminUser) {
      const nd = demoDownloads + 1;
      try { localStorage.setItem('casey_demo_downloads', String(nd)); } catch(ex) {}
      setDemoDownloads(nd);
    }
  if (setExportingLabel) setExportingLabel('Generating executive export package…');
  const r = await apiFetch(path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(model) });
  if (!r.ok) {
    let message = await r.text();
    try { const parsed = JSON.parse(message); message = parsed.detail?.message || parsed.message || message; } catch (_) {}
    alert(message);
    throw new Error(message);
  }
  const blob = await r.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
  if (setExportingLabel) setTimeout(() => setExportingLabel(''), 1200);
}
function Card({ children, className = '' }) {
  return <motion.div className={`card ${className}`} initial={{ opacity: 0, y: 18 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: 0.35 }}>{children}</motion.div>;
}
function Logo({ large = false }) {
  return <div className={large ? 'v50Logo large' : 'v50Logo'}><img src="/brand/casey_wordmark.png" alt="CASEY" /><span>Mission Control for Capital Projects</span></div>;
}
function Kpi({ icon: Icon, label, value, sub, hot }) {
  const n = Number(String((value || '') + ' ' + (sub || '')).match(/(\d+)%/)?.[1] || 0);
  const band = hot ? (n >= 80 ? 'riskLow' : n >= 60 ? 'riskMedium' : 'riskHigh') : '';
  return <Card className={`v50Kpi ${hot ? 'hot' : ''} ${band}`}><Icon size={21}/><div><p>{safeRender(label)}</p><b>{safeRender(value)}</b><span>{safeRender(sub)}</span></div></Card>;
}
function Table({ rows = [], cols = [], moneyCols = [] }) {
  return <div className="tableWrap"><table><thead><tr>{cols.map(c => <th key={c[0]}>{c[1]}</th>)}</tr></thead><tbody>{rows.map((r, i) => <tr key={i}>{cols.map(c => <td key={c[0]}>{moneyCols.includes(c[0]) ? fmt(r[c[0]]) : String(r[c[0]] ?? '')}</td>)}</tr>)}</tbody></table></div>;
}
// ── ACCOUNT PANEL ─────────────────────────────────────────────────────────────
function AccountPanel({ email, setEmail, projects, loading, onLoad, onDelete, onSave, onLoadProjects, onClose, model }) {
  const [inputEmail, setInputEmail] = React.useState(email || '');
  return <section className="savedPanel">
    <div className="savedHeader"><h2 style={{fontSize:'14px'}}>Your Account <span style={{fontSize:'10px',color:'#8df7ff',marginLeft:'4px'}}>cross-device projects</span></h2><button onClick={onClose}>✕</button></div>
    <div style={{padding:'14px 16px',borderBottom:'1px solid rgba(255,255,255,0.07)'}}>
      <p style={{fontSize:'11px',color:'#64748b',marginBottom:'8px'}}>Enter your email to save and load projects across devices. No password needed — just your email.</p>
      <div style={{display:'flex',gap:'8px'}}>
        <input value={inputEmail} onChange={e => setInputEmail(e.target.value)} placeholder="your@email.com"
          style={{flex:1,background:'rgba(255,255,255,0.05)',border:'1px solid rgba(255,255,255,0.12)',borderRadius:'3px',padding:'7px 10px',color:'#e2e8f0',fontSize:'12px'}}/>
        <button onClick={() => { setEmail(inputEmail); onLoadProjects(inputEmail); try { localStorage.setItem('casey_account_email', inputEmail); } catch {} }}
          style={{background:'rgba(141,247,255,0.1)',border:'1px solid rgba(141,247,255,0.25)',color:'#8df7ff',padding:'7px 14px',borderRadius:'3px',cursor:'pointer',fontSize:'11px',fontWeight:'700'}}>Load →</button>
      </div>
      {model && email && email.includes('@') && <button onClick={onSave}
        style={{marginTop:'8px',width:'100%',background:'rgba(141,247,255,0.08)',border:'1px solid rgba(141,247,255,0.2)',color:'#8df7ff',padding:'7px',borderRadius:'3px',cursor:'pointer',fontSize:'11px',fontWeight:'700'}}>
        ↓ Save current project to account
      </button>}
    </div>
    {loading && <div style={{padding:'20px',textAlign:'center',color:'#475569',fontSize:'12px'}}>Loading…</div>}
    {!loading && projects.length === 0 && email && <div style={{padding:'20px',textAlign:'center',color:'#475569',fontSize:'12px'}}>No saved projects yet. Run a project and save it here.</div>}
    <div className="savedGrid">
      {projects.map(p => <div className="savedCard" key={p.id}>
        <div className="savedMeta"><span>{p.subsector||'Capital Programme'}</span><em>{p.saved_at ? new Date(p.saved_at).toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'}) : ''}</em></div>
        <h3>{p.title}</h3>
        <div className="savedStats">
          <div><span>P50</span><b>{p.cost_p50||'—'}</b></div>
          <div><span>Duration</span><b>{p.schedule||'—'}</b></div>
          <div><span>Confidence</span><b>{p.confidence_pct ? p.confidence_pct+'%' : '—'}</b></div>
          <div><span>Risk</span><b>{p.risk||'—'}</b></div>
        </div>
        <div className="savedActions">
          <button className="savedLoad" onClick={() => onLoad(p)}>Load →</button>
          <button className="savedDelete" onClick={() => onDelete(p.id)}>Delete</button>
        </div>
      </div>)}
    </div>
  </section>;
}

// ── COMPARE PANEL ─────────────────────────────────────────────────────────────
function ComparePanel({ promptA, setPromptA, promptB, setPromptB, onRun, loading, result, error, onClose, currentModel }) {
  const [bmFilter, setBmFilter] = React.useState('All');
  const [bmSearch, setBmSearch] = React.useState('');
  const [activeTab, setActiveTab] = React.useState('pick');
  const [inputMode, setInputMode] = React.useState('type'); // 'type' | 'file'

  const filtered = REAL_BENCHMARKS.filter(b => {
    const matchSector = bmFilter === 'All' || b.sector === bmFilter;
    const matchSearch = !bmSearch || b.name.toLowerCase().includes(bmSearch.toLowerCase()) || b.sector.toLowerCase().includes(bmSearch.toLowerCase());
    return matchSector && matchSearch;
  });

  const delta = result?.delta;
  const pa = result?.programme_a;
  const pb = result?.programme_b;
  const rc = result?.risk_comparison;
  const recs = result?.recommendations || [];
  const wc = { A: '#10b981', B: '#8df7ff', EQUAL: '#f59e0b' };

  // Auto-detect if a loaded current project should pre-fill Option B
  React.useEffect(() => { if (result) setActiveTab('results'); }, [result]);
  React.useEffect(() => { if (currentModel?.prompt && !promptB) setPromptB(currentModel.prompt); }, [currentModel]);

  // Like-for-like suggestion based on Option B sector
  const suggestLikeForLike = () => {
    if (!promptB || !REAL_BENCHMARKS) return null;
    const bl = promptB.toLowerCase();
    const sectorHints = [
      ['rail','metro','tram','signalling','hs2','crossrail','rail / transit'],
      ['nuclear','reactor','smr','hinkley','sizewell','nuclear / energy'],
      ['data centre','datacenter','gpu','hyperscale','digital infrastructure'],
      ['pharma','gmp','biologics','vaccine','life sciences'],
      ['defence','military','awre','aukus','defence / secure'],
      ['space','lunar','mars','orbital','satellite','space / mission'],
      ['airport','terminal','runway','baggage','airport / aviation'],
      ['port','harbour','quay','berth','ports / marine'],
      ['gigafactory','battery','ev manufactur','gigafactory'],
      ['semiconductor','wafer','fab','cleanroom chip','semiconductor'],
    ];
    for (const [terms, sectorKey] of sectorHints) {
      if (terms.slice(0,-1).some(t => bl.includes(t))) {
        const matches = REAL_BENCHMARKS.filter(b => b.sector.toLowerCase().includes(sectorKey) || terms.slice(0,-1).some(t => b.sector.toLowerCase().includes(t)));
        if (matches.length > 0) return { sector: sectorKey, count: matches.length };
      }
    }
    return null;
  };
  const suggestion = suggestLikeForLike();

  return <section className="savedPanel" style={{width:'min(1020px,100vw)'}}>
    <div className="savedHeader">
      <div>
        <h2 style={{fontSize:'14px'}}>Programme Comparison</h2>
        <p style={{fontSize:'11px',color:'#475569',margin:'2px 0 0'}}>Compare any two programmes anywhere in the world — your project vs a real benchmark, two delivery options, two contractor bids, or any historical programme.</p>
      </div>
      <button onClick={onClose}>✕</button>
    </div>

    <div style={{display:'flex',gap:'0',borderBottom:'1px solid rgba(255,255,255,0.07)'}}>
      {[['pick','Set up comparison'],['results','Results & recommendations']].map(([t,l]) =>
        <button key={t} onClick={() => setActiveTab(t)}
          style={{padding:'8px 16px',fontSize:'11px',fontWeight:activeTab===t?'800':'400',
                  color:activeTab===t?'#8df7ff':'#475569',
                  borderBottom:activeTab===t?'2px solid #8df7ff':'2px solid transparent',
                  background:'none',border:'none',cursor:'pointer',letterSpacing:'.06em'}}>{l}</button>)}
    </div>

    {activeTab === 'pick' && <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'0',height:'calc(100vh - 140px)',overflow:'hidden'}}>
      {/* LEFT: Option A — benchmark library */}
      <div style={{borderRight:'1px solid rgba(255,255,255,0.07)',display:'flex',flexDirection:'column',overflow:'hidden'}}>
        <div style={{padding:'10px 14px',borderBottom:'1px solid rgba(255,255,255,0.05)',flexShrink:0}}>
          <div style={{fontSize:'9px',fontWeight:'800',letterSpacing:'.12em',color:'#10b981',marginBottom:'4px'}}>OPTION A — REFERENCE / BENCHMARK</div>
          <p style={{fontSize:'10px',color:'#475569',margin:'0 0 5px',lineHeight:'1.4'}}>Pick a real completed programme from the library (any sector, any country, any size) — or type your own. This is what you are comparing <em>against</em>.</p>
          <textarea value={promptA} onChange={e => setPromptA(e.target.value)} rows={2}
            placeholder="Select from library below, or type any programme — e.g. Crossrail Elizabeth Line UK rail, or HS2 Phase 2b..."
            style={{width:'100%',background:'rgba(16,185,129,0.05)',border:'1px solid rgba(16,185,129,0.2)',borderRadius:'3px',padding:'7px',color:'#e2e8f0',fontSize:'11px',resize:'none',boxSizing:'border-box'}}/>
          {currentModel?.prompt && <button onClick={() => setPromptA(currentModel.prompt)}
            style={{marginTop:'4px',width:'100%',background:'rgba(16,185,129,0.06)',border:'1px solid rgba(16,185,129,0.15)',color:'#10b981',padding:'4px',borderRadius:'3px',cursor:'pointer',fontSize:'10px',fontWeight:'700'}}>
            ← Use current loaded project as Option A
          </button>}
        </div>
        <div style={{padding:'6px 10px',borderBottom:'1px solid rgba(255,255,255,0.05)',flexShrink:0,display:'flex',gap:'5px',alignItems:'center'}}>
          <input value={bmSearch} onChange={e => setBmSearch(e.target.value)} placeholder="Search benchmarks..."
            style={{flex:1,background:'rgba(255,255,255,0.04)',border:'1px solid rgba(255,255,255,0.08)',borderRadius:'3px',padding:'4px 8px',color:'#e2e8f0',fontSize:'10px'}}/>
          <select value={bmFilter} onChange={e => setBmFilter(e.target.value)}
            style={{background:'rgba(255,255,255,0.04)',border:'1px solid rgba(255,255,255,0.08)',borderRadius:'3px',padding:'4px 5px',color:'#94a3b8',fontSize:'10px',maxWidth:'160px'}}>
            {['All','Rail / Transit','Nuclear / Energy','Defence / Secure Infrastructure',
              'Digital Infrastructure / Hyperscale Data Centre','Life Sciences / Biologics Manufacturing',
              'Semiconductor / Advanced Manufacturing','Battery / Gigafactory','Energy / Utilities',
              'Water / Environmental Infrastructure','Oil & Gas / Process Infrastructure',
              'Mining / Metals Infrastructure','Airport / Aviation','Roads / Highways Infrastructure',
              'Space / Mission Assurance','Ports / Marine Infrastructure'].map(s =>
              <option key={s} value={s}>{s==='All'?'All sectors ('+REAL_BENCHMARKS.length+')':s}</option>)}
          </select>
        </div>
        <div style={{flex:1,overflowY:'auto'}}>
          {filtered.map(b => <div key={b.name}
            onClick={() => setPromptA(b.prompt)}
            style={{padding:'7px 12px',borderBottom:'1px solid rgba(255,255,255,0.04)',cursor:'pointer',
                    background:promptA===b.prompt?'rgba(16,185,129,0.08)':'transparent'}}
            onMouseEnter={e=>e.currentTarget.style.background='rgba(255,255,255,0.04)'}
            onMouseLeave={e=>e.currentTarget.style.background=promptA===b.prompt?'rgba(16,185,129,0.08)':'transparent'}>
            <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:'1px'}}>
              <span style={{fontSize:'11px',fontWeight:'700',color:'#e2e8f0'}}>{b.name}</span>
              <div style={{display:'flex',gap:'3px',flexShrink:0,marginLeft:'4px'}}>
                {b.mode==='Space'&&<span style={{background:'rgba(141,247,255,0.1)',color:'#8df7ff',fontSize:'7px',padding:'1px 4px',borderRadius:'2px',fontWeight:'800'}}>SPACE</span>}
                {((b.cost_growth_pct??b.growth)||b.growth||0)>80&&<span style={{background:'rgba(239,68,68,0.12)',color:'#fca5a5',fontSize:'7px',padding:'1px 4px',borderRadius:'2px',fontWeight:'800'}}>+{(b.cost_growth_pct??b.growth)}%</span>}
                {(b.cost_growth_pct??b.growth)>0&&(b.cost_growth_pct??b.growth)<=80&&<span style={{background:'rgba(245,158,11,0.1)',color:'#fde68a',fontSize:'7px',padding:'1px 4px',borderRadius:'2px',fontWeight:'800'}}>+{(b.cost_growth_pct??b.growth)}%</span>}
              </div>
            </div>
            <div style={{fontSize:'9px',color:'#475569',marginBottom:'1px'}}>{b.sector}</div>
            <div style={{display:'flex',gap:'8px'}}>
              <span style={{fontSize:'9px',color:'#64748b'}}>Actual: ${b.cost_bn}B</span>
              {(b.schedule_slip_months??b.slip)>0&&<span style={{fontSize:'9px',color:'#f59e0b'}}>+{(b.schedule_slip_months??b.slip)}mo slip</span>}
            </div>
            {(b.failure_mode||b.failure)&&<div style={{fontSize:'8px',color:'#334155',marginTop:'1px',fontStyle:'italic'}}>{(b.failure_mode||b.failure||'').slice(0,65)}{b.failure_mode.length>65?'…':''}</div>}
          </div>)}
        </div>
      </div>

      {/* RIGHT: Option B — your project */}
      <div style={{display:'flex',flexDirection:'column',overflow:'hidden'}}>
        <div style={{padding:'10px 14px',borderBottom:'1px solid rgba(255,255,255,0.05)',flexShrink:0}}>
          <div style={{fontSize:'9px',fontWeight:'800',letterSpacing:'.12em',color:'#8df7ff',marginBottom:'4px'}}>OPTION B — YOUR PROJECT</div>
          <p style={{fontSize:'10px',color:'#475569',margin:'0 0 5px',lineHeight:'1.4'}}>Describe your programme. CASEY will run it through the full intelligence engine and compare it against Option A. Works for any country, sector, and size.</p>
          <div style={{display:'flex',gap:'5px',marginBottom:'5px'}}>
            {[['type','Type description'],['file','Upload file']].map(([m,l])=>
              <button key={m} onClick={()=>setInputMode(m)}
                style={{fontSize:'10px',fontWeight:inputMode===m?'800':'400',padding:'3px 10px',borderRadius:'3px',border:'1px solid rgba(255,255,255,0.1)',background:inputMode===m?'rgba(141,247,255,0.1)':'transparent',color:inputMode===m?'#8df7ff':'#475569',cursor:'pointer'}}>{l}</button>)}
          </div>
          {inputMode==='type' && <textarea value={promptB} onChange={e => setPromptB(e.target.value)} rows={3}
            placeholder="Describe your project — sector, country, scale, key constraints. E.g. New metro line Lagos Nigeria 45km 2031 delivery / SMR fleet UK 10 reactors / Data centre Virginia 500MW hyperscale" 
            style={{width:'100%',background:'rgba(141,247,255,0.04)',border:'1px solid rgba(141,247,255,0.2)',borderRadius:'3px',padding:'7px',color:'#e2e8f0',fontSize:'11px',resize:'vertical',boxSizing:'border-box'}}/>}
          {inputMode==='file' && <div style={{background:'rgba(141,247,255,0.04)',border:'1px dashed rgba(141,247,255,0.2)',borderRadius:'3px',padding:'14px',textAlign:'center'}}>
            <p style={{fontSize:'11px',color:'#64748b',marginBottom:'6px'}}>Upload a cost estimate, XER schedule or risk register — CASEY will extract the programme description and use it as Option B.</p>
            <p style={{fontSize:'10px',color:'#475569',marginBottom:'8px'}}>Supported: .xlsx, .csv, .xer, .pdf, .txt — or paste a description above instead.</p>
            <button onClick={()=>setInputMode('type')} style={{background:'rgba(141,247,255,0.1)',border:'1px solid rgba(141,247,255,0.2)',color:'#8df7ff',padding:'5px 14px',borderRadius:'3px',cursor:'pointer',fontSize:'10px',fontWeight:'700'}}>Type description instead →</button>
          </div>}
          {currentModel?.prompt && <button onClick={() => setPromptB(currentModel.prompt)}
            style={{marginTop:'4px',width:'100%',background:'rgba(141,247,255,0.05)',border:'1px solid rgba(141,247,255,0.15)',color:'#8df7ff',padding:'4px',borderRadius:'3px',cursor:'pointer',fontSize:'10px',fontWeight:'700'}}>
            ← Use current loaded project as Option B
          </button>}
        </div>
        {suggestion && <div style={{padding:'7px 12px',background:'rgba(141,247,255,0.04)',borderBottom:'1px solid rgba(141,247,255,0.1)',flexShrink:0}}>
          <div style={{fontSize:'10px',color:'#8df7ff',fontWeight:'700',marginBottom:'2px'}}>💡 Like-for-like suggestion</div>
          <p style={{fontSize:'10px',color:'#64748b',margin:'0 0 4px',lineHeight:'1.4'}}>We detected your project may be in the {suggestion.sector} sector. Filter Option A to {suggestion.sector} ({suggestion.count} benchmarks) for a more meaningful comparison.</p>
          <button onClick={()=>{setBmFilter(suggestion.sector);}} style={{fontSize:'10px',color:'#8df7ff',background:'rgba(141,247,255,0.08)',border:'1px solid rgba(141,247,255,0.2)',padding:'3px 10px',borderRadius:'3px',cursor:'pointer',fontWeight:'700'}}>Filter to {suggestion.sector} →</button>
        </div>}
        <div style={{padding:'10px 14px',borderBottom:'1px solid rgba(255,255,255,0.05)',flexShrink:0,flex:1}}>
          <div style={{fontSize:'9px',fontWeight:'800',letterSpacing:'.1em',color:'#64748b',marginBottom:'8px'}}>WHAT THE COMPARISON PRODUCES</div>
          {[['◆','Full intelligence pack for both programmes — P50, P80, confidence, gate readiness'],['⚖️','Side-by-side delta: cost, schedule, confidence, risk exposure, P80 gap'],['🎯','Sector match check — tells you if the comparison is like-for-like or cross-sector'],['⚠','Risk comparison — top risks, shared risk themes, EMV delta, P80 gap'],['📋','3–5 specific recommendations for your decision'],['🌍','Works for any country, sector, size — Earth or Space'],['⏱','Takes 10–15 seconds — both models run in parallel'],].map(([icon,text])=>
            <div key={text} style={{display:'flex',gap:'8px',marginBottom:'6px',fontSize:'10px',color:'#94a3b8',alignItems:'flex-start'}}>
              <span style={{flexShrink:0}}>{icon}</span><span style={{lineHeight:'1.4'}}>{text}</span>
            </div>)}
        </div>
        <div style={{padding:'10px 14px',flexShrink:0}}>
          <button onClick={onRun} disabled={loading||!promptA.trim()||!promptB.trim()}
            style={{width:'100%',background:loading||!promptA.trim()||!promptB.trim()?'rgba(141,247,255,0.04)':'rgba(141,247,255,0.1)',border:'1px solid rgba(141,247,255,0.3)',color:'#8df7ff',padding:'10px',borderRadius:'4px',cursor:loading||!promptA.trim()||!promptB.trim()?'default':'pointer',fontSize:'12px',fontWeight:'800',letterSpacing:'.06em'}}>
            {loading?'◌ Building intelligence packs…':'◆ Run comparison — any country, any sector'}
          </button>
          {!promptA.trim()&&<p style={{fontSize:'10px',color:'#f59e0b',textAlign:'center',marginTop:'5px'}}>← Select or type Option A first</p>}
          {promptA.trim()&&!promptB.trim()&&<p style={{fontSize:'10px',color:'#8df7ff',textAlign:'center',marginTop:'5px'}}>← Describe your project as Option B to run</p>}
          {error&&<div style={{marginTop:'6px',color:'#fca5a5',fontSize:'10px',padding:'5px',background:'rgba(239,68,68,0.06)',borderRadius:'3px'}}>{error}</div>}
          {loading&&<p style={{textAlign:'center',color:'#475569',fontSize:'10px',marginTop:'6px'}}>Running both programmes through the CASEY intelligence engine — 10–15 seconds.</p>}
        </div>
      </div>
    </div>}

    {activeTab==='results' && <div style={{overflowY:'auto',height:'calc(100vh - 140px)',padding:'12px 16px'}}>
      {!result&&!loading&&<div style={{textAlign:'center',padding:'40px',color:'#475569',fontSize:'12px'}}>
        Set up a comparison and run it to see results here.
        <br/><button onClick={()=>setActiveTab('pick')} style={{marginTop:'10px',background:'rgba(141,247,255,0.08)',border:'1px solid rgba(141,247,255,0.2)',color:'#8df7ff',padding:'6px 16px',borderRadius:'3px',cursor:'pointer',fontSize:'11px',fontWeight:'700'}}>← Set up comparison</button>
      </div>}
      {loading&&<div style={{textAlign:'center',padding:'40px',color:'#64748b',fontSize:'12px'}}>Building both intelligence packs — 10–15 seconds…</div>}
      {result&&delta&&!loading&&<>
        {/* Sector match badge */}
        <div style={{marginBottom:'10px',display:'flex',gap:'8px',alignItems:'center',flexWrap:'wrap'}}>
          <span style={{background:delta.sector_match==='Like-for-like'?'rgba(16,185,129,0.12)':'rgba(245,158,11,0.1)',border:delta.sector_match==='Like-for-like'?'1px solid rgba(16,185,129,0.3)':'1px solid rgba(245,158,11,0.3)',color:delta.sector_match==='Like-for-like'?'#10b981':'#f59e0b',padding:'3px 10px',borderRadius:'20px',fontSize:'10px',fontWeight:'800'}}>
            {delta.sector_match==='Like-for-like'?'✓ Like-for-like comparison':'⚠ Cross-sector comparison'}
          </span>
          {pa?.country&&pb?.country&&pa.country!==pb.country&&<span style={{background:'rgba(141,247,255,0.06)',border:'1px solid rgba(141,247,255,0.15)',color:'#8df7ff',padding:'3px 8px',borderRadius:'20px',fontSize:'10px'}}>🌍 {pa.country} vs {pb.country}</span>}
          <span style={{fontSize:'10px',color:'#475569'}}>{delta.sector_note?.slice?.(0,80)}</span>
        </div>

        {/* Verdict */}
        <div style={{background:'rgba(255,255,255,0.03)',border:`2px solid ${wc[delta.winner]||'#8df7ff'}`,borderRadius:'5px',padding:'10px 14px',marginBottom:'10px'}}>
          <div style={{fontSize:'9px',fontWeight:'800',letterSpacing:'.12em',color:wc[delta.winner]||'#8df7ff',marginBottom:'3px'}}>
            {delta.winner==='EQUAL'?'EQUAL — NO CLEAR PREFERENCE':`OPTION ${delta.winner} PREFERRED`}
          </div>
          <p style={{fontSize:'11px',color:'#e2e8f0',lineHeight:'1.6',margin:0}}>{delta.winner_reason}</p>
        </div>

        {/* Delta strip */}
        <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:'6px',marginBottom:'10px'}}>
          {[['Cost',delta.cost_verdict,delta.cost_delta_pct?(delta.cost_delta_pct>0?'+':'')+delta.cost_delta_pct+'%':'—'],
            ['Confidence',delta.confidence_verdict,delta.confidence_delta?(delta.confidence_delta>0?'+':'')+delta.confidence_delta+'pts':'—'],
            ['Schedule',delta.schedule_verdict,delta.schedule_delta_months?(delta.schedule_delta_months>0?'+':'')+delta.schedule_delta_months+' mo':'—'],
            ['P80 gap','Risk exposure',rc?.p80_gap?(rc.p80_gap>0?'+':'')+rc.p80_gap+'B':'—'],
          ].map(([label,verdict,diff])=><div key={label} style={{background:'rgba(255,255,255,0.03)',border:'1px solid rgba(255,255,255,0.06)',borderRadius:'4px',padding:'8px',textAlign:'center'}}>
            <div style={{fontSize:'9px',color:'#475569',fontWeight:'800',letterSpacing:'.08em',marginBottom:'2px'}}>{label}</div>
            <div style={{fontSize:'10px',color:'#e2e8f0',fontWeight:'700',marginBottom:'1px'}}>{verdict}</div>
            <div style={{fontSize:'12px',color:'#8df7ff',fontWeight:'800'}}>{diff}</div>
          </div>)}
        </div>

        {/* Recommendations */}
        {recs.length>0&&<div style={{background:'rgba(141,247,255,0.04)',border:'1px solid rgba(141,247,255,0.15)',borderRadius:'4px',padding:'10px 14px',marginBottom:'10px'}}>
          <div style={{fontSize:'9px',fontWeight:'800',letterSpacing:'.12em',color:'#8df7ff',marginBottom:'8px'}}>RECOMMENDATIONS</div>
          {recs.map((r,i)=><div key={i} style={{display:'flex',gap:'8px',marginBottom:'6px',fontSize:'11px',color:'#cbd5e1',lineHeight:'1.5',paddingBottom:'6px',borderBottom:'1px solid rgba(255,255,255,0.04)'}}>
            <span style={{color:'#8df7ff',flexShrink:0,fontWeight:'800'}}>{i+1}.</span><span>{r}</span>
          </div>)}
        </div>}

        {/* Side by side */}
        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'8px',marginBottom:'10px'}}>
          {[[pa,'#10b981','A'],[pb,'#8df7ff','B']].map(([p,clr,lbl],i)=>p&&<div key={i} style={{background:'rgba(255,255,255,0.02)',border:`1px solid ${clr}33`,borderRadius:'4px',padding:'10px'}}>
            <div style={{fontSize:'9px',fontWeight:'800',letterSpacing:'.1em',color:clr,marginBottom:'6px'}}>OPTION {lbl} — {p.label}</div>
            <div style={{fontSize:'9px',color:'#475569',marginBottom:'6px',fontStyle:'italic'}}>{p.subsector}{p.country?' · '+p.country:''}</div>
            {[['P50',p.cost_p50],['P80',p.cost_p80],['Schedule',p.schedule],['Confidence',p.confidence_pct?p.confidence_pct+'%':'—'],
              ['Risk',p.risk],['Gate',p.gate_review_readiness],['OBA P50',p.oba_adjusted_p50],
              ['Mortality',p.programme_mortality_risk],['Financing',p.financing],
            ].filter(([,v])=>v&&v!=='—').map(([k,v])=>
              <div key={k} style={{display:'flex',justifyContent:'space-between',padding:'2px 0',borderBottom:'1px solid rgba(255,255,255,0.04)',fontSize:'10px'}}>
                <span style={{color:'#475569'}}>{k}</span>
                <span style={{color:'#e2e8f0',fontWeight:'600',maxWidth:'55%',textAlign:'right',lineHeight:'1.3'}}>{String(v).slice(0,45)}</span>
              </div>)}
            {/* Top risks */}
            {p.risks&&p.risks.length>0&&<>
              <div style={{fontSize:'9px',color:'#f59e0b',fontWeight:'800',letterSpacing:'.08em',margin:'8px 0 3px'}}>TOP RISKS</div>
              {p.risks.slice(0,3).map((r,ri)=><div key={ri} style={{fontSize:'9px',color:'#94a3b8',padding:'2px 0',borderBottom:'1px solid rgba(255,255,255,0.03)',lineHeight:'1.4'}}>
                <span style={{color:'#f59e0b',marginRight:'4px'}}>{ri+1}.</span>
                {r.title||r.event||'—'} {r.probability?<span style={{color:'#64748b'}}>({Math.round(r.probability*100)}%)</span>:null}
              </div>)}
            </>}
            {p.if_this_fails&&<div style={{marginTop:'7px',padding:'5px',background:'rgba(239,68,68,0.05)',borderRadius:'3px',fontSize:'9px',color:'#fca5a5',lineHeight:'1.4',fontStyle:'italic'}}>Historical pattern: {p.if_this_fails.slice(0,100)}…</div>}
            {p.board_attack_1&&<div style={{marginTop:'6px',padding:'5px',background:'rgba(255,255,255,0.03)',borderRadius:'3px',fontSize:'9px',color:'#64748b',lineHeight:'1.4'}}>
              <b style={{color:clr,display:'block',marginBottom:'1px',fontSize:'8px',letterSpacing:'.08em'}}>BOARD CHALLENGE:</b>{p.board_attack_1}
            </div>}
          </div>)}
        </div>

        {/* Risk comparison */}
        {rc&&<div style={{background:'rgba(245,158,11,0.04)',border:'1px solid rgba(245,158,11,0.15)',borderRadius:'4px',padding:'10px 14px',marginBottom:'10px'}}>
          <div style={{fontSize:'9px',fontWeight:'800',letterSpacing:'.12em',color:'#f59e0b',marginBottom:'8px'}}>RISK COMPARISON</div>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr',gap:'8px',marginBottom:'8px'}}>
            <div style={{textAlign:'center',padding:'6px',background:'rgba(255,255,255,0.03)',borderRadius:'3px'}}>
              <div style={{fontSize:'9px',color:'#475569',marginBottom:'2px'}}>Option A risk EMV</div>
              <div style={{fontSize:'14px',fontWeight:'800',color:'#e2e8f0'}}>${rc.emv_a}B</div>
              <div style={{fontSize:'9px',color:'#64748b'}}>P80: {rc.p80_a}</div>
            </div>
            <div style={{textAlign:'center',padding:'6px',background:'rgba(255,255,255,0.03)',borderRadius:'3px'}}>
              <div style={{fontSize:'9px',color:'#475569',marginBottom:'2px'}}>Option B risk EMV</div>
              <div style={{fontSize:'14px',fontWeight:'800',color:'#e2e8f0'}}>${rc.emv_b}B</div>
              <div style={{fontSize:'9px',color:'#64748b'}}>P80: {rc.p80_b}</div>
            </div>
            <div style={{textAlign:'center',padding:'6px',background:'rgba(255,255,255,0.03)',borderRadius:'3px'}}>
              <div style={{fontSize:'9px',color:'#475569',marginBottom:'2px'}}>EMV delta</div>
              <div style={{fontSize:'14px',fontWeight:'800',color:rc.emv_delta>0?'#fca5a5':'#10b981'}}>{rc.emv_delta>0?'+':''}{rc.emv_delta}B</div>
              <div style={{fontSize:'9px',color:'#64748b'}}>P80 gap: {rc.p80_gap?(rc.p80_gap>0?'+':'')+rc.p80_gap+'B':'—'}</div>
            </div>
          </div>
          <p style={{fontSize:'10px',color:'#94a3b8',margin:0,lineHeight:'1.5'}}>{rc.risk_verdict}</p>
          {rc.shared_risk_themes?.length>0&&<p style={{fontSize:'10px',color:'#64748b',marginTop:'5px',lineHeight:'1.5'}}>Shared risk themes: {rc.shared_risk_themes.join(', ')}</p>}
        </div>}

        <div style={{textAlign:'right',marginTop:'8px'}}>
          <button onClick={()=>setActiveTab('pick')} style={{background:'rgba(255,255,255,0.04)',border:'1px solid rgba(255,255,255,0.1)',color:'#475569',padding:'5px 14px',borderRadius:'3px',cursor:'pointer',fontSize:'10px'}}>← Run another comparison</button>
        </div>
      </>}
    </div>}
  </section>;
}


// ── ONBOARDING GUIDE ─────────────────────────────────────────────────────────
function OnboardingGuide({ onClose }) {
  const [step, setStep] = React.useState(0);
  const steps = [
    {
      icon: '◆',
      title: 'What CASEY does — in 30 seconds',
      body: 'Type any capital programme in plain English. CASEY generates a full intelligence pack: P50/P80/P90 cost, schedule, risk register, confidence score, OBA, gate readiness, board attack simulation and export pack. In 4-12 seconds. A consultant charges £50-150K and takes 8 weeks for the same output.',
      sub: 'Works for rail, nuclear, defence, data centres, space, mining, airport, gigafactory, water, life sciences, semiconductor — any sector, any country, any scale.',
    },
    {
      icon: '1️⃣',
      title: 'Step 1: Describe your project',
      body: 'Type in the FREE RUN box: sector, country, scale, key constraints. Example: "HS2 Phase 2b tunnelling UK rail 250km £50B 2031 delivery high complexity". Click Run. Your first project is free.',
      sub: 'No form to fill in. No dropdown menus. Just describe it the way you would to a colleague. CASEY detects sector, location, scale and complexity automatically.',
    },
    {
      icon: '2️⃣',
      title: 'Step 2: Read the Overview',
      body: 'P50 = most likely outturn. P80 = the board conversation number (what you need in reserve). Confidence % = how board-defensible the estimate is at this definition maturity. The 5 scenario cards (Base, Faster, Cheaper, Lower Risk, Premium) show cost/schedule/confidence for each trade-off — click any to recalculate.',
      sub: 'All numbers update when you change scenario. The £/$/€ sign reflects the project country. Unit rates in the Cost tab show cost per km, per MW, per m², per tpa etc.',
    },
    {
      icon: '3️⃣',
      title: 'Step 3: Work through the tabs',
      body: 'The tabs follow how a board pack is built: Cost → Schedule → Risk → QCRA/QSRA → Scenarios → Intel. The Mortality tab shows programme cancellation risk with named precedents. The Evidence Gaps tab names what's missing before any gate review. Assurance shows the board questions CASEY generates from live model data.',
      sub: 'Each tab has a guide line at the top explaining what it shows and what decision it supports.',
    },
    {
      icon: '4️⃣',
      title: 'Step 4: Use the Advisor',
      body: 'The Advisor tab is a what-if engine. Ask "What if planning is delayed 18 months?" — CASEY reruns the model and shows you a cost/confidence/schedule delta. Ask "Is this gate-ready?" for an IPA readiness assessment. Ask "What would an external reviewer challenge first?" for an independent view.',
      sub: 'Also try: "What is the P80 exposure?" | "Which risk will kill this programme?" | "What does CASEY position say about this sector?" | "What is the OBA-adjusted outturn?"',
    },
    {
      icon: '📤',
      title: 'Step 5: Export the deliverables',
      body: 'Export Board Pack → 7-page PDF ready for an investment committee. Export Cost Workbook → full P10/P50/P90 CBS in Excel. Export Risk Register → complete risk register CSV. Export XER → Primavera-compatible schedule. All exports generated from live model data — not templates.',
      sub: 'The Earth Demo and Space Demo are always free and show full exports. Earth = HS2 Phase 2b UK rail. Space = Lunar Base Alpha.',
    },
  ]
  return <section className="trustRuntimeBar">
    <div className="trustRuntimeLead"><b>V150 Trust Runtime</b><span>{g.decision_posture || 'Scenario locked to canonical state'}</span></div>
    {items.map(([k,v]) => <div className="trustRuntimeMetric" key={k}><span>{k}</span><strong>{v}</strong></div>)}
  </section>;
}

function parseMoneyLocal(v) {
    if (v === undefined || v === null) return 0;
    if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
    const s = String(v).replace(/[$,£€]/g,'').trim().toUpperCase();
    const n = parseFloat(s.replace(/[^0-9.-]/g,''));
    if (!Number.isFinite(n)) return 0;
    if (s.includes('T')) return n * 1000;
    if (s.includes('M')) return n / 1000;
    return n;
  }
  function moneyLocal(n) { return n >= 1000 ? `$${(n/1000).toFixed(1)}T` : n >= 1 ? `$${n.toFixed(1)}B` : `$${Math.round(n*1000)}M`; }
  function normalizeCostRowsForUI(modelLike) {
    const m = modelLike || {};
    const target = parseMoneyLocal(m.cost_p50);
    let rows = Array.isArray(m.cost_breakdown) && m.cost_breakdown.length
      ? m.cost_breakdown
      : (Array.isArray(m.cost_lines) ? m.cost_lines : []);
    rows = rows.map((x, idx) => {
      const typeRaw = String(x.type || 'Direct');
      const type = /reserve|risk|contingency/i.test(typeRaw) ? 'Reserve' : (/indirect|owner|pm|management|prelim/i.test(typeRaw) ? 'Indirect' : 'Direct');
      return {
        ...x,
        type,
        cbs: x.cbs || `C-${String(idx+1).padStart(2,'0')}`,
        description: x.description || x.title || type,
        p10_bn: parseMoneyLocal(x.p10_bn ?? x.low_p10 ?? x.low ?? x.p10),
        p50_bn: parseMoneyLocal(x.p50_bn ?? x.most_likely_p50 ?? x.most_likely ?? x.p50 ?? x.value),
        p90_bn: parseMoneyLocal(x.p90_bn ?? x.high_p90 ?? x.high ?? x.p90),
        impact_basis: x.impact_basis || x.basis || `${type} cost basis reconciled to selected scenario P50.`
      };
    }).filter(x => x.p50_bn > 0);

    // If detailed rows are absent/unusable, prefer backend's explicit bucket totals.
    // This prevents the UI from silently reverting every project/scenario to a fake 76/14/10 split.
    if ((!rows.length || rows.reduce((a,x)=>a+x.p50_bn,0) <= 0) && target > 0) {
      const explicitDirect = parseMoneyLocal(m.direct_cost);
      const explicitIndirect = parseMoneyLocal(m.indirect_cost);
      const explicitReserve = parseMoneyLocal(m.risk_reserve ?? m.reserve_cost ?? m.contingency_cost);
      if (explicitDirect + explicitIndirect + explicitReserve > 0) {
        rows = [
          { cbs:'01.00', description:'Direct delivery scope', type:'Direct', p10_bn:explicitDirect*0.82, p50_bn:explicitDirect, p90_bn:explicitDirect*1.28, impact_basis:'Explicit backend direct-cost bucket.' },
          { cbs:'90.00', description:'Indirects, owner costs and integration', type:'Indirect', p10_bn:explicitIndirect*0.82, p50_bn:explicitIndirect, p90_bn:explicitIndirect*1.28, impact_basis:'Explicit backend indirect-cost bucket.' },
          { cbs:'99.00', description:'Risk reserve and contingency', type:'Reserve', p10_bn:explicitReserve*0.70, p50_bn:explicitReserve, p90_bn:explicitReserve*1.55, impact_basis:'Explicit backend risk/reserve bucket.' }
        ].filter(x => x.p50_bn > 0);
      } else {
        rows = [
          { cbs:'01.00', description:'Direct delivery scope', type:'Direct', p10_bn:target*0.62, p50_bn:target*0.76, p90_bn:target*0.96, impact_basis:'Synthetic reconciliation split created because detailed cost rows were missing.' },
          { cbs:'90.00', description:'Indirects, owner costs and integration', type:'Indirect', p10_bn:target*0.10, p50_bn:target*0.14, p90_bn:target*0.18, impact_basis:'Synthetic reconciliation split created because detailed cost rows were missing.' },
          { cbs:'99.00', description:'Risk reserve and contingency', type:'Reserve', p10_bn:target*0.06, p50_bn:target*0.10, p90_bn:target*0.16, impact_basis:'Synthetic reconciliation split created because detailed cost rows were missing.' }
        ];
      }
    }

    const sum = rows.reduce((a,x)=>a + parseMoneyLocal(x.p50_bn), 0);
    if (target > 0 && sum > 0) {
      const factor = target / sum;
      rows = rows.map(x => ({...x,
        p10_bn: Number((parseMoneyLocal(x.p10_bn) * factor).toFixed(3)),
        p50_bn: Number((parseMoneyLocal(x.p50_bn) * factor).toFixed(3)),
        p90_bn: Number((parseMoneyLocal(x.p90_bn) * factor).toFixed(3)),
      }));
      // remove rounding drift from the largest row so Direct+Indirect+Reserve equals P50 exactly on screen/export
      const drift = Number((target - rows.reduce((a,x)=>a+x.p50_bn,0)).toFixed(3));
      if (Math.abs(drift) >= 0.001 && rows.length) {
        let maxIdx = 0;
        rows.forEach((x,i)=>{ if (x.p50_bn > rows[maxIdx].p50_bn) maxIdx = i; });
        rows[maxIdx] = {...rows[maxIdx], p50_bn: Number((rows[maxIdx].p50_bn + drift).toFixed(3))};
      }
    }
    return rows;
  }
  function normalizeModelForUI(raw) {
    const m = {...(raw || {})};
    const rows = normalizeCostRowsForUI(m);
    m.cost_breakdown = rows;
    m.cost_lines = rows;
    m.risk_register = Array.isArray(m.risk_register) ? m.risk_register : (Array.isArray(m.risks) ? m.risks : []);
    m.risks = m.risk_register;
    m.schedule_detail = Array.isArray(m.schedule_detail) ? m.schedule_detail : (Array.isArray(m.schedule_rows) ? m.schedule_rows : []);
    m.schedule_rows = m.schedule_detail;
    m.scenario_matrix = Array.isArray(m.scenario_matrix) ? m.scenario_matrix : (Array.isArray(m.scenario_comparison) ? m.scenario_comparison : []);
    return m;
  }

  function lockedProjectContext(sourceModel, fallbackPrompt) {
    if (!sourceModel) return null;
    return {
      id: sourceModel.id,
      title: sourceModel.title,
      prompt: sourceModel.prompt || fallbackPrompt,
      mode: sourceModel.mode,
      subsector: sourceModel.subsector,
      location: sourceModel.location,
      scale: sourceModel.scale,
      baseline_cost_p50: sourceModel._base_cost_p50 || sourceModel.cost_p50,
      baseline_schedule: sourceModel._base_schedule || sourceModel.schedule,
      baseline_confidence_pct: sourceModel._base_confidence_pct || sourceModel.confidence_pct,
      baseline_risk: sourceModel._base_risk || sourceModel.risk
    };
  }


  // ── Demo gate ────────────────────────────────────────────────────────
  // NEVER locked: Earth Demo, Space Demo, Showcase, Pricing, Film, Request Access
  // 1 free run: generate() via OneShotDemo modal
  // ADMIN (unlimited): localhost, admin emails, ?admin=casey2024

  const CASEY_ADMIN_EMAILS = [
    'deepa@caseai.co.uk','admin@controlorbit.com','demo@controlorbit.com',
    'jaimahadeshwar@yahoo.com','jai@controlorbit.com','test@yahoo.com',
    'jai@caseai.co.uk','deepa@controlorbit.com'
  ];
  const CASEY_ADMIN_URL_KEYS = ['casey2024','casey','corbit2024','admin2024'];

  const isAdminUser = React.useMemo(() => {
    try {
      // 1. Localhost / local dev — always admin
      if (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1' || window.location.hostname === '0.0.0.0') return true;
      // 2. URL key bypass — ?admin=casey2024 or ?admin=casey
      const params = new URLSearchParams(window.location.search);
      const k = (params.get('admin') || params.get('admin_key') || '').toLowerCase().trim();
      if (CASEY_ADMIN_URL_KEYS.includes(k)) return true;
      // 3. Admin email stored from previous modal run
      const storedEmail = (localStorage.getItem('casey_user_email') || '').toLowerCase().trim();
      if (storedEmail && CASEY_ADMIN_EMAILS.includes(storedEmail)) return true;
      return false;
    } catch(e) { return false; }
  }, []);

  const [demoUsed, setDemoUsed] = React.useState(() => {
    try {
      // Never locked on localhost
      if (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1') return false;
      // URL key always bypasses
      const params = new URLSearchParams(window.location.search);
      const k = (params.get('admin') || params.get('admin_key') || '').toLowerCase().trim();
      if (CASEY_ADMIN_URL_KEYS.includes(k)) return false;
      // Admin email bypasses
      const storedEmail = (localStorage.getItem('casey_user_email') || '').toLowerCase().trim();
      if (storedEmail && CASEY_ADMIN_EMAILS.includes(storedEmail)) return false;
      // Check if free run has been used
      return localStorage.getItem('casey_demo_used') === '1';
    } catch(e) { return false; }
  });

  const [demoDownloads, setDemoDownloads] = React.useState(() => {
    try { return parseInt(localStorage.getItem('casey_demo_downloads') || '0'); }
    catch(e) { return 0; }
  });

  // ── SAVED PROJECTS ────────────────────────────────────────────────────────
  const [savedProjects, setSavedProjects] = React.useState(() => {
    try { return JSON.parse(localStorage.getItem('casey_saved_projects') || '[]'); }
    catch(e) { return []; }
  });
  const [showSaved, setShowSaved] = React.useState(false);
  const [showInvestor, setShowInvestor] = React.useState(false);
  const [showOnboarding, setShowOnboarding] = React.useState(() => {
    try { return !localStorage.getItem('casey_onboarding_done'); } catch { return true; }
  });
  const [backendStatus, setBackendStatus] = React.useState('unknown');

  // ── ACCOUNT STATE ──────────────────────────────────────────────────────────
  const [accountEmail, setAccountEmail] = React.useState(() => {
    try { return localStorage.getItem('casey_account_email') || ''; } catch { return ''; }
  });
  const [accountProjects, setAccountProjects] = React.useState([]);
  const [showAccount, setShowAccount] = React.useState(false);
  const [accountLoading, setAccountLoading] = React.useState(false);

  // ── COMPARE STATE ──────────────────────────────────────────────────────────
  const [showCompare, setShowCompare] = React.useState(false);
  const [comparePromptA, setComparePromptA] = React.useState('');
  const [comparePromptB, setComparePromptB] = React.useState('');
  const [compareResult, setCompareResult] = React.useState(null);
  const [compareLoading, setCompareLoading] = React.useState(false);
  const [compareError, setCompareError] = React.useState(''); // 'ok' | 'down' | 'unknown'

  // ── ACCOUNT FUNCTIONS ────────────────────────────────────────────────────────
  async function loadAccountProjects(email) {
    if (!email || !email.includes('@')) return;
    setAccountLoading(true);
    try {
      const res = await apiFetch(`/account/projects?email=${encodeURIComponent(email)}`);
      if (res.ok) { const d = await res.json(); setAccountProjects(d.projects || []); }
    } catch {}
    setAccountLoading(false);
  }

  async function saveToAccount() {
    if (!model || !accountEmail || !accountEmail.includes('@')) return;
    try {
      await post('/account/save-project', {
        email: accountEmail,
        title: model.title || model.prompt || 'Unnamed',
        subsector: model.subsector || '',
        prompt: model.prompt || prompt,
        cost_p50: model.cost_p50 || '',
        schedule: model.schedule || '',
        confidence_pct: model.confidence_pct || 0,
        risk: model.risk || '',
        scenario: scenario,
        model_json: JSON.stringify(model)
      });
      localStorage.setItem('casey_account_email', accountEmail);
      await loadAccountProjects(accountEmail);
    } catch(e) { console.error(e); }
  }

  async function loadAccountProject(proj) {
    try {
      const res = await apiFetch(`/account/project/${proj.id}?email=${encodeURIComponent(accountEmail)}`);
      if (res.ok) {
        const d = await res.json();
        const m = normalizeModelForUI(d.model || d);
        setModel(m); setPrompt(d.prompt || ''); setScenario(d.scenario || 'base');
        setProjectContext(lockedProjectContext(m, d.prompt || ''));
        setShowAccount(false); setShow(false); setTab('overview');
      }
    } catch(e) { console.error(e); }
  }

  async function deleteAccountProject(id) {
    try {
      await apiFetch(`/account/project/${id}?email=${encodeURIComponent(accountEmail)}`, { method: 'DELETE' });
      await loadAccountProjects(accountEmail);
    } catch {}
  }

  // ── COMPARE FUNCTION ──────────────────────────────────────────────────────
  async function runComparison() {
    if (!comparePromptA.trim() || !comparePromptB.trim()) return;
    setCompareLoading(true); setCompareError(''); setCompareResult(null);
    try {
      const result = await post('/compare', {
        prompt_a: comparePromptA, prompt_b: comparePromptB,
        client_a: 'Option A',
        client_b: model ? (model.client || 'Your Project') : 'Option B',
        class_level: 3, schedule_level: 3
      });
      setCompareResult(result);
    } catch(e) { setCompareError(String(e.message || e)); }
    setCompareLoading(false);
  }

  function saveCurrentProject() {
    if (!model) return;
    const entry = {
      id: Date.now(),
      saved_at: new Date().toISOString(),
      title: model.title || model.prompt || 'Unnamed project',
      subsector: model.subsector || '',
      cost_p50: model.cost_p50 || '',
      schedule: model.schedule || '',
      confidence_pct: model.confidence_pct || '',
      risk: model.risk || '',
      prompt: model.prompt || prompt,
      model_snapshot: model,
    };
    const updated = [entry, ...savedProjects].slice(0, 20);
    setSavedProjects(updated);
    try { localStorage.setItem('casey_saved_projects', JSON.stringify(updated)); } catch(e) {}
  }

  function deleteSaved(id) {
    const updated = savedProjects.filter(p => p.id !== id);
    setSavedProjects(updated);
    try { localStorage.setItem('casey_saved_projects', JSON.stringify(updated)); } catch(e) {}
  }

  function loadSaved(entry) {
    const m = normalizeModelForUI(entry.model_snapshot);
    setModel(m);
    setPrompt(entry.prompt || '');
    setProjectContext(lockedProjectContext(m, entry.prompt || ''));
    setShowSaved(false);
    setShow(false);
    setTab('overview');
  }

  // Mark the free run as used — only called after a real generate(), never after instant demos
  function markDemoUsed() {
    try {
      localStorage.setItem('casey_demo_used', '1');
      setDemoUsed(true);
    } catch(e) {}
  }

  // Demo configs — prompts used to generate reference case models via /generate
  const DEMO_CONFIGS = {
    'earth': { prompt: 'HS2 Phase 2b tunnelling stations signalling systems integration possessions operator acceptance UK rail', client: 'Reference case', demo_type: 'earth', demo_label: 'Reference case — HS2 Phase 2b Rail Mega Programme', demo_headline: 'Full programme intelligence pack — cost, schedule, risk, benchmarks, board attack and exports.' },
    'space': { prompt: 'Lunar Base Alpha life support nuclear surface power autonomous commissioning resupply logistics 1000 crew', client: 'Reference case', demo_type: 'space', demo_label: 'Reference case — Lunar Base Alpha Deep Space Programme', demo_headline: 'Space programme intelligence — TRL risk, launch logistics, life support, autonomous commissioning.' },
    'awre': { prompt: 'AWRE Aldermaston nuclear warhead facility upgrade classified defence sovereign supply chain security accreditation UK MOD', client: 'Reference case', demo_type: 'defence', demo_label: 'Reference case — AWRE Aldermaston Nuclear Infrastructure', demo_headline: 'Classified programme intelligence — security accreditation, sovereign supply chain, operational acceptance.' },
    'gigafactory': { prompt: 'Battery gigafactory West Midlands UK 50GWh EV manufacturing cell production utility grid connection', client: 'Reference case', demo_type: 'gigafactory', demo_label: 'Reference case — Gigafactory UK Battery Manufacturing', demo_headline: 'EV battery manufacturing intelligence — grid connection, cell chemistry, yield ramp, utility complexity.' },
  };

  async function loadInstantDemo(type) {
    const cfg = DEMO_CONFIGS[type] || DEMO_CONFIGS['earth'];
    setLoading(true); setError(''); setModel(null); setTab('overview');
    setShow(false); setShowShowcase(false);
    setSimulationStage('Loading ' + (type === 'earth' ? 'HS2 Phase 2b Earth Demo' : type === 'space' ? 'Lunar Base Alpha Space Demo' : 'reference case') + '…');
    try {
      // Use /generate — the same endpoint showcase and free run use (always works)
      const payload = {
        prompt: cfg.prompt,
        client: cfg.client,
        class_level: 3,
        schedule_level: 4,
        scenario: 'base',
        demo: true,
        active_model: null
      };
      const m = await post('/generate', payload);
      // Mark as demo client-side so DemoBanner renders correctly
      m.demo_mode = true;
      m.demo_type = cfg.demo_type;
      m.demo_label = cfg.demo_label;
      m.demo_headline = cfg.demo_headline;
      const nm = normalizeModelForUI(m);
      setError('');
      setModel(nm);
      setPrompt(cfg.prompt);
      setScenario('base');
      setClient(cfg.client);
      setTab('overview');
    } catch(e) {
      const msg = String(e.message || e);
      setError(JSON.stringify({
        message: 'Backend starting up — wait 20 seconds and try again.',
        sub: 'The server is waking up (it sleeps after inactivity). Wait 20 seconds then click the demo button again. The Showcase Library works immediately.',
        email: 'deepa@caseai.co.uk',
        linkedin: 'https://www.linkedin.com/company/caseai'
      }));
      setBackendStatus('down');
      setShow(false);
    } finally {
      setLoading(false);
      setSimulationStage('');
    }
  }

  async function generate(nextScenario = scenario, nextPrompt = prompt, activeContext = model || projectContext, clientOverride = client, opts = {}) {
    setError(''); setShow(false);

    // Canonical state lock: every scenario re-run must preserve the active project universe
    // (Earth, Space, Rail, Energy, Defence, etc.) rather than falling back to the default demo seed.
    const contextLock = activeContext ? lockedProjectContext(activeContext, nextPrompt) : null;
    const canonicalPrompt = contextLock?.prompt || nextPrompt || prompt;

    setPropagating(true);
    setSimulationStage(nextScenario === 'base' ? 'Building base simulation…' : 'Re-running scenario from locked project context…');
    setConfidencePulse(true);
    setTimeout(() => setPropagating(false), 1600);
    setLoading(true); setTab(nextScenario !== 'base' ? 'compare' : 'overview');
    // Demo gate — fires only for brand-new project runs from the main console
    // NEVER fires for: showcase library, earth/space demo, scenario switching on existing model
    const isGated = !isAdminUser && demoUsed && !activeContext && !opts.isShowcase && !opts.isDemo;
    if (isGated) {
      setLoading(false); setPropagating(false);
      setError(JSON.stringify({
        message: "You've used your one free CASEY intelligence run.",
        sub: "Browse the Showcase Library (40 free reference cases), run the Earth or Space demos, or get in touch for full access.",
        email: "deepa@caseai.co.uk",
        linkedin: "https://www.linkedin.com/company/caseai"
      }));
      setTab('pricing');
      return;
    }
    try {
      const payload = {
        prompt: canonicalPrompt,
        client: clientOverride,
        class_level: Number(classLevel),
        schedule_level: Number(scheduleLevel),
        scenario: nextScenario,
        demo: true,
        active_model: contextLock
      };
      const m = normalizeModelForUI(await post('/generate', payload));
      const nextContext = lockedProjectContext(m, canonicalPrompt);
      setModel(m); setProjectContext(nextContext); setScenario(nextScenario); setPrompt(canonicalPrompt);
      // Mark free run as used (only for real generate calls, not instant demos or showcase)
      if (!isAdminUser && !activeContext && !opts.isShowcase && !opts.isDemo) { markDemoUsed(); }
    } catch (e) {
      let raw = String(e.message || e);
      try {
        const p = JSON.parse(raw);
        const d = p.detail || p;
        if (typeof d === 'object' && d.message) {
          setError(JSON.stringify({ message: d.message, sub: d.sub, email: d.email, linkedin: d.linkedin }));
        } else {
          setError(typeof d === 'string' ? d : raw);
        }
      } catch { setError(raw); }
    }
    finally { setLoading(false); setSimulationStage(''); setConfidencePulse(false); }
  }
  function runEarth() { setProjectContext(null); setError(''); setShow(false); setShowShowcase(false); loadInstantDemo('earth'); }
  function runSpace() { setShowShowcase(false); setProjectContext(null); setError(''); loadInstantDemo('space'); }
  function runShowcase(project) { setError(''); setClient(project.client || 'Strategic reference case'); setShow(false); setShowShowcase(false); setProjectContext(null); setScenario('base'); setPrompt(project.prompt); generate('base', project.prompt, null, project.client || 'Strategic reference case', { isShowcase: true, isDemo: true }); }
  function advisorQuestionText(input) {
    if (typeof input === 'string') return input.trim();
    if (input && typeof input === 'object') {
      if (typeof input.preventDefault === 'function') input.preventDefault();
      if (typeof input.question === 'string') return input.question.trim();
      if (typeof input.text === 'string') return input.text.trim();
      if (typeof input.currentTarget?.dataset?.question === 'string') return input.currentTarget.dataset.question.trim();
      if (typeof input.currentTarget?.innerText === 'string' && input.currentTarget.innerText.trim().toLowerCase() !== 'ask') return input.currentTarget.innerText.trim();
    }
    return String(chatQ || '').trim();
  }
  async function ask(overrideQ) {
    const q = advisorQuestionText(overrideQ);
    if (!q) return;
    if (!model) {
      setChat(x => [...x, {role:'user',text:q}, {role:'assistant',text:'**No project loaded yet**\n\nGenerate a project first, then come back to challenge it. Board attack answers are grounded in the live model.'}]);
      return;
    }
    setChatQ('');
    setChat(x => [...x, { role: 'user', text: q }]);
    const ql = q.toLowerCase();
    const tvc = model?.traditional_vs_casey || {};
    const attacks = model?.board_attack_simulation || [];
    const ifFails = model?.if_this_fails;
    let instant = null;
    if (ql.includes('fail') && (ql.includes('programme') || ql.includes('blame') || ql.includes('publicly'))) {
      instant = ifFails ? ('**IF THIS PROGRAMME FAILS**\n\n' + ifFails + '\n\nThis is the governing failure mode for this programme type. The board must close this risk before capital is committed.') : null;
    } else if (ql.includes('traditional') || ql.includes('controls vs') || ql.includes('what would a traditional')) {
      if (tvc.traditional && tvc.casey) {
        instant = '**TRADITIONAL CONTROLS vs CASEY**\n\nTraditional view: ' + (tvc.traditional||'Civil progress and costs appear on track.') + '\n\n**What CASEY reads underneath:**\n' + (tvc.casey||'Governing constraint not visible in headline progress.') + '\n\nP80 tail: ' + (tvc.tail_pct||25) + '% above P50.\n\nThe gap between these two readings is the product.';
      }
    } else if (ql.includes('board attack') || ql.includes('likely board') || ql.includes('board is really deciding')) {
      if (attacks.length >= 3) {
        instant = '**BOARD ASSURANCE QUESTIONS**\n' + attacks.length + ' questions before capital approval:\n\n' + attacks.map(function(a,i){return (i+1)+'. '+a;}).join('\n') + '\n\nThese are the questions a serious investment committee asks before the team can hide behind dashboards.';
      }
    } else if (ql.includes('board not seeing') || (ql.includes('hidden') && ql.includes('issue'))) {
      const shock = model?.executive_shock_insight;
      if (shock) instant = '**What the dashboard fails to show:**\n\n' + shock;
    } else if (ql.includes('casey position') || ql.includes('give me casey')) {
      const thinking = model?.casey_thinking || '';
      if (thinking) instant = '**CASEY position:**\n\n' + thinking + '\n\nConfidence: ' + model?.confidence_pct + '% — ' + (model?.confidence_engine_label || '');
    } else if (ql.includes('governing chain') || ql.includes('real governing')) {
      const chain = (model?.causal_chain || []).join(' → ');
      if (chain) instant = '**REAL GOVERNING CAUSAL CHAIN**\n\n' + chain + '\n\nThis chain drives cost, schedule and confidence. If any link is weak, the model shifts right.';
    } else if (ql.includes('assumptions collapse') || ql.includes('collapse confidence')) {
      const bqs = (model?.board_challenge_questions || []).slice(0,4);
      instant = '**ASSUMPTIONS THAT COLLAPSE CONFIDENCE FIRST**\n\n' + bqs.map(function(q,i){return (i+1)+'. '+q;}).join('\n') + '\n\nClosing evidence on these is the fastest route to board approval.';
    } else if (ql.includes('one intervention') || ql.includes('changes confidence fastest')) {
      const constraint = model?.primary_constraint || 'the governing procurement and evidence constraint';
      instant = '**THE ONE INTERVENTION THAT CHANGES CONFIDENCE FASTEST**\n\nClose the evidence gap on: ' + constraint + '\n\nThis means: named owner + named trigger + quantified residual + documented closure date. That single action moves from approvable? to what is the decision?';
    } else if (ql.includes('external assurance') || ql.includes('assurance reviewer')) {
      const attacks = model?.board_attack_simulation || [];
      instant = '**WHAT AN EXTERNAL ASSURANCE REVIEWER ATTACKS FIRST**\n\n1. Is the P50 defensible against QCRA P80?\n2. Is reserve linked to named risks?\n3. Who owns the governing constraint?\n4. Is schedule float operationally usable?\n5. What evidence retires the approval blocker?\n\n' + attacks.slice(0,3).map(function(a,i){return (i+6)+'. '+a;}).join('\n');
    } else if (ql.includes('destroy board confidence') || ql.includes('destroy confidence')) {
      instant = '**WHAT WOULD DESTROY BOARD CONFIDENCE FASTEST**\n\n1. A cost increase that was known but not disclosed.\n2. A schedule slip that reveals float was never real.\n3. A risk with no named owner and no mitigation evidence.\n4. A commissioning failure cascading into operator rejection.\n5. A procurement gap forcing sole-source at programme peak.\n\nFor this programme: ' + (model && model.executive_shock_insight ? model.executive_shock_insight : 'confidence sits in integration and evidence maturity, not visible civil progress.');
    } else if (ql.includes('management probably hiding') || ql.includes('looks green')) {
      const shock = model?.executive_shock_insight;
      const flags = (model?.red_flags || []).slice(0,3);
      instant = '**WHAT MANAGEMENT IS PROBABLY HIDING**\n\n' + (shock || 'The governing constraint is not in the headlines.') + (flags.length ? '\n\n**Commercial observations:**\n' + flags.map(function(f,i){return (i+1)+'. '+f;}).join('\n') : '') + '\n\nTraditional reports show numbers. CASEY shows what the numbers are trying to hide.';
    }
    if (instant) {
      setChat(x => [...x, { role: 'assistant', text: instant }]);
      return;
    }
    try {
      const safeProject = normalizeModelForUI({ ...model });
      const r = await post('/chat', { question: q, project: safeProject, demo: true });
      const answer = normalizeChatAnswer(r);
      setChat(x => [...x, { role: 'assistant', text: String(answer || 'CASEY returned no advisor response.'), delta: r?.delta || null }]);
    } catch (e) {
      setChat(x => [...x, { role: 'assistant', text: 'CASEY advisor recovered from a request error. The governance challenge remains: do not rely on headline progress alone. Ask: What is the board really deciding?' }]);
    }
  }

  async function upload(e) {
    const f = e.target.files?.[0]; if (!f) return;
    setUploadResult({ filename: f.name, size_bytes: f.size, findings: ['File received. CASEY is normalising messy client structure rather than expecting a clean template.'], red_flags: ['Deep parser requires source-file validation before commercial reliance.'] });
    try {
      const fd = new FormData(); fd.append('file', f);
      const r = await apiFetch('/upload', { method: 'POST', body: fd });
      const data = await r.json();
      setUploadResult({ filename: f.name, size_bytes: f.size, schema_confidence: 'Preliminary mapping', ...data });
    } catch (_) {
      setUploadResult({ filename: f.name, size_bytes: f.size, schema_confidence: 'Offline demo mapping', findings: ['Workbook accepted even if tabs, columns or coding are inconsistent.', 'CASEY mapped the file to estimate challenge, risk posture and evidence readiness.', 'For production, connect the Python intake pipeline for sheet parsing, WBS inference and schedule/risk linking.'], red_flags: ['Do not show raw JSON to the board.', 'Require named owners for missing basis, reserve logic and schedule float.', 'Triangulate cost workbook with risk register and schedule before approval.'], next_steps: ['Ask for XLSX + XER + risk register as a source bundle.', 'Run board challenge on governing constraint before approval.', 'Export the challenged pack after evidence closure.'] });
    }
  }

  function runShock(kind) {
    if (!model) return;
    const shock = {
      signalling_slip: { months: 4, cost: 0.04, conf: -8, note: 'Signalling slip moved the delivery tail and made operator acceptance the governing board issue.' },
      procurement_gap: { months: 2, cost: 0.06, conf: -10, note: 'Procurement evidence gap increased P80/P90 exposure and lowered board defensibility.' },
      reserve_cut: { months: 0, cost: -0.03, conf: -12, note: 'Reserve cut improves the headline but consumes downside protection and weakens the approval posture.' },
      operator_delay: { months: 5, cost: 0.03, conf: -9, note: 'Operator acceptance delay shifted the causal chain from civil progress to commissioning readiness.' },
      scope_growth: { months: 3, cost: 0.08, conf: -11, note: 'Scope growth at 8% — cost increased, schedule extended, confidence fell. Reserve adequacy now in question.' },
      political_exposure: { months: 0, cost: 0.05, conf: -14, note: 'Political and funding pressure raised programme risk. Board will require stronger evidence posture before approval.' },
    }[kind];
    if (!shock) return;
    const baseCostBn = parseMoneyLocal(model.cost_p50_bn || model.p50_cost_bn || model.cost_p50 || model.cost || 0);
    const baseMonthsNum = Number(model.schedule_months || model.duration_months || (String(model.schedule || '').match(/\d+/)||[])[0] || 0);
    const nextCost = Math.max(.1, (baseCostBn || 1) * (1 + shock.cost));
    const nextMonths = Math.max(1, Math.round((baseMonthsNum || 1) + shock.months));
    const nextConf = Math.max(5, Math.min(95, Math.round((model.confidence_pct || 50) + shock.conf)));
    const nextCostStr = moneyLocal(nextCost);
    const nextSchedStr = nextMonths + ' months';
    const nextP80 = Math.round(nextCost * 1.22 * 10) / 10;
    const nextP90 = Math.round(nextCost * 1.35 * 10) / 10;
    const nextRange = nextCostStr + ' P50 | $' + nextP80 + 'B P80';
    const nextConfLabel = nextConf < 45 ? 'Do not approve without more evidence' : nextConf < 60 ? 'Board challenge likely' : nextConf < 75 ? 'Conditionally approvable' : 'Board-defensible';
    const mutationStamp = '[STRESS TEST: ' + kind.replace(/_/g,' ').toUpperCase() + ']';
    setModel({ ...model,
      cost_p50_bn: nextCost, p50_cost_bn: nextCost,
      cost_p50: nextCostStr,
      cost_range: nextRange,
      schedule_months: nextMonths, duration_months: nextMonths,
      schedule: nextSchedStr,
      confidence_pct: nextConf,
      confidence_label: nextConfLabel,
      confidence_engine_label: nextConfLabel,
      executive_shock_insight: mutationStamp + ' ' + shock.note,
      board_briefing: [
        mutationStamp,
        shock.note,
        'Programme baseline preserved: ' + (model.cost_p50 || nextCostStr) + ' | ' + (model.schedule || nextSchedStr),
        'Stress-test delta: P50 moves to ' + nextCostStr + ', schedule to ' + nextSchedStr + ', confidence to ' + nextConf + '%',
        'P80 downside (recalculated): $' + nextP80 + 'B',
        'Exports stamped from this mutated model state — download now to capture the stress-tested position.'
      ],
      casey_thinking: 'CASEY runtime mutation: ' + shock.note + ' The model recalculated cost, schedule, confidence, reserve posture and board challenge language from the same source of truth.',
      scenario_trade: shock.note,
      last_runtime_event: kind,
      stress_test_applied: kind,
      stress_test_note: shock.note,
    });
    setTab('overview');
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
  const scenarioMatrix = model?.scenario_matrix || model?.scenario_comparison || [];
  // Cost split must use exact normalized classes. Do not use includes('direct') because 'indirect' contains 'direct'.
  const costTypeOf = (x) => {
    const t = String(x?.type || '').trim().toLowerCase();
    if (/reserve|risk|contingency/.test(t)) return 'Reserve';
    if (/indirect|owner|pm|management|prelim/.test(t)) return 'Indirect';
    return 'Direct';
  };
  const direct = costs.filter(x => costTypeOf(x) === 'Direct').reduce((a, b) => a + parseMoneyLocal(b.p50_bn || 0), 0);
  const indirect = costs.filter(x => costTypeOf(x) === 'Indirect').reduce((a, b) => a + parseMoneyLocal(b.p50_bn || 0), 0);
  const reserves = costs.filter(x => costTypeOf(x) === 'Reserve').reduce((a, b) => a + parseMoneyLocal(b.p50_bn || 0), 0);
  const targetP50 = parseMoneyLocal(model?.cost_p50);
  const reconcileCheck = targetP50 ? Math.abs((direct + indirect + reserves) - targetP50) : 0;
  const emailBody = model ? [
    'Please review this project in CASEY.', '', `Project: ${model.title}`, `Scenario: ${model.scenario_label || scenario}`,
    `P50 Cost: ${safeRender(model.cost_p50)}`, `Cost Range: ${safeRender(model.cost_range)}`, `Schedule: ${safeRender(model.schedule)}`,
    `Risk / Confidence: ${model.risk} / ${model.confidence_pct}%`
  ].join('\n') : 'Please send me CASEY access.';
  const emailLink = `mailto:hello@casey.ai?subject=${encodeURIComponent('CASEY project review')}&body=${encodeURIComponent(emailBody)}`;
  const confLens = model ? confidenceLens(model) : null;
  const p80Talk = model ? p80PlainEnglish(model) : null;
  const tradePack = model ? gainedSacrificedExposed(model) : null;

  return <div className="app v50EliteApp">
    <Briefing open={briefing} onClose={() => setBriefing(false)} onEarth={runEarth} onSpace={runSpace}/>
    <OneShotDemo open={trialOpen} onClose={() => setTrialOpen(false)} onComplete={(m) => { const nm = normalizeModelForUI(m); setModel(nm); setProjectContext(lockedProjectContext(nm, nm?.prompt || prompt)); setShow(false); setTrialOpen(false); setTab('overview'); }} />
    <AnimatePresence>{loading && <Loading text="Building full CASEY intelligence pack..."/>}</AnimatePresence>
    {show && !model && <Hero onBriefing={() => setBriefing(true)} onEarth={runEarth} onSpace={runSpace} onConsole={() => setShow(false)} onTryDemo={() => setTrialOpen(true)}/>} 
    <header className="v50ConsoleTop"><Logo/><nav>
      <button onClick={() => { setModel(null); setProjectContext(null); setShowShowcase(false); setShow(true); setError(''); }}>Home</button>
      <button onClick={() => setBriefing(true)}>Film</button>
      <button onClick={() => setTrialOpen(true)}>Free run</button>
      <button onClick={() => { setModel(null); setShow(false); setShowShowcase(true); setError(''); }}>Showcase library</button>
      {savedProjects.length > 0 && <button onClick={() => setShowSaved(s => !s)} style={{position:'relative'}}>Saved <span style={{background:'#8df7ff',color:'#0a1628',borderRadius:'10px',padding:'1px 6px',fontSize:'10px',fontWeight:'900',marginLeft:'4px'}}>{savedProjects.length}</span></button>}
      {model && <button onClick={saveCurrentProject} style={{color:'#8df7ff',fontWeight:'700'}}>↓ Save (local)</button>}
      <button onClick={() => setShowAccount(s => !s)} style={{color:'#8df7ff',fontWeight:'700'}}>Account</button>
      <button onClick={() => setShowCompare(s => !s)} style={{color:'#b18cff',fontWeight:'700'}}>Compare ◆</button>
      <button onClick={runEarth}>Earth demo</button>
      <button onClick={runSpace}>Space demo</button>
      <button onClick={() => setShowInvestor(s => !s)} style={{color:'#b18cff',fontWeight:'700'}}>Investor brief</button>
      <button onClick={() => setShowOnboarding(true)} style={{color:'#64748b',fontSize:'10px',fontWeight:'700',letterSpacing:'.06em'}}>How to use</button>
      <a href={emailLink}>Request access</a>
      <span style={{display:'flex',alignItems:'center',gap:'4px',fontSize:'9px',color:backendStatus==='ok'?'#10b981':backendStatus==='down'?'#ef4444':'#64748b',fontWeight:'700',letterSpacing:'.08em'}}>
        <span style={{width:'6px',height:'6px',borderRadius:'50%',background:backendStatus==='ok'?'#10b981':backendStatus==='down'?'#ef4444':'#475569',display:'inline-block'}}/>
        {backendStatus==='ok'?'LIVE':backendStatus==='down'?'Starting...':'...'}
      </span>
    </nav></header>
    {showOnboarding && <OnboardingGuide onClose={() => { setShowOnboarding(false); try { localStorage.setItem('casey_onboarding_done','1'); } catch {} }}/>}
    {showSaved && <SavedProjectsPanel projects={savedProjects} onLoad={loadSaved} onDelete={deleteSaved} onClose={() => setShowSaved(false)}/>}
    {showAccount && <AccountPanel email={accountEmail} setEmail={setAccountEmail} projects={accountProjects} loading={accountLoading} onLoad={loadAccountProject} onDelete={deleteAccountProject} onSave={saveToAccount} onLoadProjects={loadAccountProjects} onClose={() => setShowAccount(false)} model={model}/>}
    {showCompare && <ComparePanel promptA={comparePromptA} setPromptA={setComparePromptA} promptB={comparePromptB} setPromptB={setComparePromptB} onRun={runComparison} loading={compareLoading} result={compareResult} error={compareError} onClose={() => setShowCompare(false)} currentModel={model}/>}
    {showInvestor && <InvestorPanel onClose={() => setShowInvestor(false)}/>}
    <main className={model ? 'v50Console' : 'v50Console emptyConsole'}>
      {error && !showShowcase && !show && <GatedMessage raw={error} onDismiss={() => setError('')} onShowcase={() => { setError(''); setShowShowcase(true); }} onEarth={() => { setError(''); runEarth(); }} onSpace={() => { setError(''); runSpace(); }}/>}
      {!model && showShowcase && <ShowcaseLibrary onRun={runShowcase} onBack={() => setShowShowcase(false)} />}
      {!model && !show && !showShowcase && <section className="commandGrid"><Card className="command">
  <h1 style={{fontSize:'18px',marginBottom:'4px'}}>Generate a project</h1>
  <p style={{fontSize:'11px',color:'#475569',marginBottom:'12px',lineHeight:'1.5'}}>Enter any capital programme — infrastructure, defence, space, pharma, energy, data centres. CASEY generates a first-pass cost estimate (P10/P50/P90), schedule, risk register, scenario analysis and board intelligence pack. All fields are sector-specific and location-aware.</p>
  <label>Project command</label><textarea value={prompt} onChange={e => setPrompt(e.target.value)} /> <div className="chips">{examples.map(x => <button key={x} onClick={() => setPrompt(x)}>{x}</button>)}</div><div className="grid4"><input value={client} onChange={e => setClient(e.target.value)} placeholder="Client / operator"/><select value={classLevel} onChange={e => setClassLevel(e.target.value)}>{[1,2,3,4,5].map(x => <option key={x} value={x}>Class {x}</option>)}</select><select value={scheduleLevel} onChange={e => setScheduleLevel(e.target.value)}>{[1,2,3,4,5].map(x => <option key={x} value={x}>Level {x}</option>)}</select><select value={scenario} onChange={e => setScenario(e.target.value)}>{scenarios.map(x => <option key={x} value={x}>{x}</option>)}</select></div><button className="primary" onClick={() => generate()}><Sparkles/> Generate full intelligence pack</button><button className="secondary" onClick={() => { setShowShowcase(true); setError(''); }}><Globe2/> Open global showcase library</button></Card><Card><h2>What CASEY generates</h2><p style={{fontSize:'11px',color:'#64748b',marginBottom:'10px'}}>From a single project description — in seconds.</p>{['Executive summary with P50, schedule and confidence score','Cost workbook — direct, indirect, reserve by CBS line','5 scenario trade-offs: base, faster, cheaper, lower risk, premium','Risk register — cause, event, impact, owner, trigger, mitigation','QCRA/QSRA probability curves and tornado chart','Board attack simulation — 5 questions your committee will ask','Location intelligence, financing context and OBA assessment','Procurement packages with lead times and single-source flags'].map((x,i)=><div className="reason" style={{padding:'5px 0',borderBottom:'1px solid rgba(255,255,255,0.04)'}} key={x}><span style={{color:'#8df7ff',marginRight:'8px',fontSize:'10px',fontWeight:'800'}}>{i+1}</span><span style={{fontSize:'11px'}}>{x}</span></div>)}</Card></section>}
      {model && <>
        <DemoBanner model={model}/>
        <section className="confidenceEngineBadge"><b>{model.confidence_engine_label || 'CASEY Confidence Engine'}</b><span>{safeRender(typeof model.confidence_engine_detail === 'object' ? model.confidence_engine_detail?.plain_english || 'Benchmark + probabilistic + sector-trained reasoning' : model.confidence_engine_detail || 'Benchmark + probabilistic + sector-trained reasoning')}</span></section>
        <TrustRuntimeBar model={model}/>
        <LiveCalibrationStrip model={model}/>
        <section className="kpis">
        {model?.live_data_enriched && <div className="kpi" style={{background:'rgba(16,185,129,0.06)',border:'1px solid rgba(16,185,129,0.2)',borderRadius:'4px',padding:'6px 10px',display:'flex',flexDirection:'column',gap:'2px',justifyContent:'center'}}>
          <span style={{fontSize:'9px',color:'#10b981',fontWeight:'800',letterSpacing:'.1em'}}>● LIVE DATA</span>
          <b style={{fontSize:'10px',color:'#e2e8f0'}}>{(model.live_data_sources||[]).join(' · ') || 'Enriched'}</b>
        </div>}<Kpi icon={Globe2} label="Mode / sector" value={safeRender(model.mode)} sub={safeRender(model.subsector)}/><Kpi icon={Activity} label="P50 cost" value={safeRender(model.cost_p50)} sub={safeRender(model.cost_range)}/><Kpi icon={Zap} label="Schedule" value={safeRender(model.schedule)} sub={`QSRA P80 ${model.monte_carlo?.qsra?.p80 || '—'} months`}/><Kpi icon={ShieldAlert} label="Delivery confidence" value={safeRender(confidenceLens(model)?.headline)} sub={`${safeRender(model.risk)} risk · ${safeRender(model.confidence_pct)}% · ${safeRender(model.scenario_label)}`} hot/></section>
        <IntelligenceMeta model={model} mode={viewMode} setMode={setViewMode}/>
        <PropagationPulse scenario={scenario} active={propagating}/>
        <ScenarioSelector scenario={scenario} generate={generate} matrix={scenarioMatrix} model={model} prompt={prompt} projectContext={projectContext}/>
        <ExportStrip model={model}
          onBoardPack={() => download('/export/all', model, `${model.id || 'casey'}_DEMO_BOARD_PACK.zip`)}
          onExcel={() => download('/export/workbook', model, `${model.id || 'casey'}_DEMO_COST_WORKBOOK.xlsx`)}
          onRisk={() => download('/export/risk-register', model, `${model.id || 'casey'}_DEMO_RISK_REGISTER.xlsx`)}
          onXer={() => download('/export/xer', model, `${model.id || 'casey'}_DEMO_SCHEDULE.xer`)}
          onQcra={() => download('/export/qcra-qsra', model, `${model.id || 'casey'}_DEMO_QCRA_QSRA.xlsx`)}/>
        {demoUsed && !isAdminUser && model && <div style={{background:'rgba(245,158,11,0.1)',borderBottom:'1px solid rgba(245,158,11,0.25)',padding:'8px 20px',display:'flex',gap:'16px',alignItems:'center',flexWrap:'wrap'}}>
          <span style={{fontSize:'11px',color:'#f59e0b',fontWeight:'800'}}>✓ FREE RUN COMPLETE</span>
          <span style={{fontSize:'11px',color:'#94a3b8'}}>Exports available below. Earth Demo, Space Demo and Showcase Library always free.</span>
          <a href="mailto:hello@controlorbit.com?subject=CASEY Full Access" style={{marginLeft:'auto',fontSize:'11px',color:'#8df7ff',fontWeight:'700',textDecoration:'none',background:'rgba(141,247,255,0.1)',padding:'4px 12px',borderRadius:'3px',border:'1px solid rgba(141,247,255,0.3)'}}>Request full access →</a>
        </div>}
      <nav className="tabs">
        {[
          ['overview','Overview'],
          ['cost','Cost'],
          ['schedule','Schedule'],
          ['risk','Risk'],
          ['monte','QCRA/QSRA'],
          ['compare','Scenarios'],
          ['delta','Intel'],
          ['assurance','Assurance'],
          ['outputs','Outputs'],
          ['advisor','Advisor'],
          ['runtime','Stress Test'],
          ['benchmark','Benchmarks'],
          ['causal','Causal'],
          ['method','Method'],
          ['pricing','Pricing'],
        ].map(x => {
          const isActive = tab===x[0];
          return <button key={x[0]} className={isActive?'active':''} onClick={()=>setTab(x[0])} title={TAB_GUIDE_SIMPLE[x[0]]||''}>{x[1]}</button>;
        })}
      </nav>
      {TAB_GUIDE_SIMPLE[tab] && <div style={{background:'rgba(141,247,255,0.04)',borderBottom:'1px solid rgba(141,247,255,0.06)',padding:'5px 16px',fontSize:'10px',color:'#64748b',lineHeight:'1.4',display:'flex',justifyContent:'space-between',alignItems:'center'}}>
        <span><span style={{color:'#8df7ff',marginRight:'5px',fontWeight:'700'}}>▸</span>{TAB_GUIDE_SIMPLE[tab]}</span>
        {tab === 'advisor' && <span style={{color:'#8df7ff',fontSize:'9px',cursor:'pointer',opacity:.7}} onClick={()=>document.querySelector('.ask input')?.focus()}>Click to ask →</span>}
        {tab === 'cost' && <span style={{color:'#475569',fontSize:'9px'}}>Unit rates show cost per km / MW / m² / GWh depending on sector</span>}
        {tab === 'risk' && <span style={{color:'#475569',fontSize:'9px'}}>Top right chart shows highest-EMV risks — these drive P80/P90</span>}
        {tab === 'monte' && <span style={{color:'#475569',fontSize:'9px'}}>X-axis = confidence percentile. P50 = headline. P80 = board exposure.</span>}
        {tab === 'benchmark' && <span style={{color:'#475569',fontSize:'9px'}}>Named real programmes from public record — actual outturn vs planned</span>}
      </div>}
        {tab === 'overview' && <>
          {model.executive_shock_insight && <section className="layout one"><Card className="shockCard"><h2>⚡ Live model update</h2><p>{model.executive_shock_insight}</p></Card></section>}
          <section className="layout two">
            <Card><h2>Executive summary</h2><p style={{fontSize:'13px',lineHeight:'1.6'}}>{model.executive_summary || `${model.title} has been classified as ${safeRender(model.subsector)}. CASEY generated a first-pass cost, schedule, risk and confidence model for the selected scenario.`}</p><div className="miniMetrics"><b><span>Direct cost</span>{fmt(direct)}</b><b><span>Indirect cost</span>{fmt(indirect)}</b><b><span>Risk / reserve</span>{fmt(reserves)}</b></div><h3>Recommendation</h3>{(model.next_best_actions || []).slice(0,5).map((x,i)=><div className="reason" key={i}><span>{i+1}</span>{safeRender(x)}</div>)}</Card>
            <Card><h2>Board briefing — what the data says</h2>{(model.board_briefing || model.board_challenge_questions || []).slice(0,5).map((x,i)=><div className="reason" key={i}><span>{i+1}</span>{safeRender(x)}</div>)}<h3>CASEY thinking</h3><p className="caseyThinking">{model.casey_thinking || 'CASEY interprets this as a system-of-systems infrastructure programme requiring cost, schedule, risk and decision intelligence.'}</p></Card>
          </section>
          <section className="layout two eliteLayer">
            <Card className="confidenceMeaningCard" style={{padding:'12px 14px'}}><h2 style={{fontSize:'12px',marginBottom:'6px'}}>What the confidence score means</h2><h3 style={{fontSize:'13px'}}>{safeRender(confLens?.headline)}</h3><p className="big">{safeRender(confLens?.meaning)}</p><div className="reason"><span>!</span><b>Decision rule</b><br/>{safeRender(confLens?.decisionRule)}</div><div className="reason"><span>→</span><b>Primary constraint</b><br/>{safeRender(confLens?.constraint)}</div><div className="reason"><span>%</span><b>Plain English</b><br/>The confidence % is a board-defensibility score — not optimism. 75%+ = evidence supports approval. 55-74% = challengeable at committee. Below 55% = gaps need closing before capital commitment. It is driven by: how well the project matches known benchmarks, whether the critical path has a named owner, procurement certainty, reserve adequacy, and scenario posture.</div></Card>
            <Card><h2>Likely board questions</h2>{boardQuestions(model).slice(0,6).map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}<h3>CASEY final position</h3><p className="caseyThinking finalPosition">{finalPosition(model)}</p></Card>
          </section>
          <IncumbentPressurePanel model={model} direct={direct} indirect={indirect} reserves={reserves} reconcileCheck={reconcileCheck}/>
          <section className="layout two eliteLayer">
            <Card><h2>Evidence threshold map</h2><p className="chartCaption">Shows why the confidence number is where it is, and what must improve before board approval.</p><ResponsiveContainer width="100%" height={260}><BarChart data={evidenceScorecard(model)} layout="vertical"><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis type="number" domain={[0,100]}/><YAxis dataKey="name" type="category" width={145}/><Tooltip formatter={(v) => [`${v}%`, 'board-defensibility score']}/><ReferenceLine x={70} stroke="#ffd96a88" label="board comfort"/><Bar dataKey="score" fill="#8df7ff"/></BarChart></ResponsiveContainer>{evidenceScorecard(model).map((x,i)=><div className="reason compactReason" key={x.name}><span>{i+1}</span><b>{x.name}: {Math.round(x.score)}%</b><br/>{x.note}</div>)}</Card>
            <Card><h2>Contradiction scan</h2><p className="chartCaption">CASEY does not just make the case look better. It exposes the trade-off that could get challenged.</p>{contradictionScan(model).map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}<h3>Demo close line</h3><p className="caseyThinking finalPosition">Traditional project controls reports show numbers. CASEY shows the board what the numbers are trying to hide.</p></Card>
          </section>
          <LiveCalibrationPanel model={model}/>
          {baseVs?.base && <section className="layout two">
            <Card className="shockCard"><h2>Scenario vs Base</h2><p>{safeRender(baseVs.plain_english)}</p><div className="miniMetrics"><b><span>Base P50</span>{safeRender(baseVs.base?.cost_p50)}<small>{safeRender(baseVs.base?.schedule_months)} mo · {safeRender(baseVs.base?.confidence_pct)}%</small></b><b><span>{safeRender(baseVs.selected?.scenario)} P50</span>{safeRender(baseVs.selected?.cost_p50)}<small>{safeRender(baseVs.selected?.schedule_months)} mo · {safeRender(baseVs.selected?.confidence_pct)}%</small></b>{baseVs.delta && <b><span>Delta</span>{safeRender(baseVs.delta.cost_direction) === 'same' ? 'No cost move' : `${safeRender(baseVs.delta.cost)} ${safeRender(baseVs.delta.cost_direction)}`}<small>{safeRender(baseVs.delta.months)} mo · {safeRender(baseVs.delta.confidence_pts)} pts</small></b>}</div></Card>
            <Card><h2>What changed and why</h2>{(model.scenario_delta_intelligence || []).slice(0,5).map((x,i)=><div className="reason" key={i}><span>{i+1}</span><b>{x.label}: {x.value}</b><br/>{x.meaning}</div>)}</Card>
          </section>}
          <section className="layout two">
            <Card><h2>Mission control signals</h2><div className="missionCardGrid">{(model.mission_control_cards || []).slice(0,6).map((c,i)=><div className="intelCard" key={i}><b>{c.label}</b><p>{c.signal}</p><span>{c.severity}</span></div>)}</div></Card>
            <Card><h2>Uncertainty narrative</h2><p>{safeRender(model.uncertainty_narrative?.estimate_maturity)}</p><p>{safeRender(model.uncertainty_narrative?.schedule_maturity)}</p><p>{safeRender(model.uncertainty_narrative?.interpretation)}</p><h3>Benchmark comparison</h3>{(model.benchmark_comparison || []).slice(0,4).map((b,i)=><div className="reason" key={i}><span>{i+1}</span><b>{safeRender(b?.archetype)}</b> · {safeRender(b?.anchor_cost)} · {safeRender(b?.anchor_duration_months)} months</div>)}</Card>
          </section>
          <section className="layout two"><BenchmarkIntelligence model={model}/><CausalGraph model={model}/></section>
          <section className="layout two">
            <Card><h2>Confidence drivers</h2>{(model.sector_confidence_drivers || ['Benchmark similarity: high where comparable infrastructure archetypes exist','Scope maturity: concept / budget level until package evidence is supplied','Procurement certainty: sensitive to long-lead equipment and market capacity','Schedule maturity: improves when critical path and commissioning logic are validated','Interface exposure: controlled by utilities, systems integration and operational constraints']).map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card>
            <Card><h2>Why this position was produced</h2>{(model.why_casey_generated_this || ['The brief indicates the infrastructure asset and operating environment from the brief','The programme was mapped to benchmark memory and sector archetypes','Cost, schedule and risk were calibrated against class maturity and delivery complexity','The narrative is designed for early board challenge, not certified pricing']).map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card>
          </section>
          <section className="layout two">
            <Card><h2>Primary cost drivers</h2>{(model.sector_primary_cost_drivers || ['Utility / enabling infrastructure','Specialist systems and long-lead equipment','Commissioning and validation complexity','Programme management, preliminaries and indirects','Risk reserve driven by procurement and interface uncertainty']).map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card>
            <Card><h2>Top schedule threats</h2>{(model.sector_schedule_threats || ['Utility energisation delay','Long-lead equipment procurement and supplier capacity','Design freeze instability and scope movement','Systems integration and commissioning bottlenecks','Approvals, safety case, permitting or operational access constraints']).map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card>
          </section>
          <section className="layout one"><AdvisoryFeeCounter model={model}/></section>
        </>}

        {tab === 'compare' && <section className="layout two"><Card><h2>Scenario comparison</h2><p style={{fontSize:'12px',color:'#64748b',marginBottom:'10px'}}>Re-run any scenario below. Each recalculates cost, schedule, confidence, risk register and exports from the same source of truth — instantly.</p>{model?.stress_test_applied && <div style={{background:"rgba(245,158,11,0.1)",border:"1px solid rgba(245,158,11,0.3)",borderRadius:"3px",padding:"8px 12px",marginBottom:"10px",fontSize:"11px",color:"#f59e0b"}}><b>STRESS TEST ACTIVE: {String(model.stress_test_applied).replace(/_/g," ").toUpperCase()}</b><br/>{model.stress_test_note} — P50 now {safeRender(model.cost_p50)}, confidence {model.confidence_pct}%. Scenario re-runs below use the stressed baseline.</div>}<div className="runtimeInline"><button onClick={()=>setTab('runtime')}><Zap size={15}/> Open Live Stress Test</button><button onClick={()=>runShock('signalling_slip')}>Simulate 4-month signalling slip</button><button onClick={()=>runShock('procurement_gap')}>Simulate procurement evidence gap</button></div>{(()=>{
  const baseRow=scenarioMatrix.find(x=>x.scenario==='base')||{};
  const bCost=parseMoneyLocal(baseRow.cost_p50||baseRow.cost||'0');
  const bConf=parseInt(String(baseRow.confidence_pct||baseRow.confidence||'50'));
  const bSched=parseInt(String(baseRow.schedule_months||'0'));
  const tradeNotes={
    base:'Reference case. Balanced cost, schedule and evidence posture for board challenge.',
    faster:'Time bought at cost of money and float. Confidence falls — board will ask if saving is real.',
    cheaper:'Lower number carries higher residual risk. Evidence deferred — board must accept this explicitly.',
    lower_risk:'Reserve adds confidence but costs more time and money. Requires QCRA evidence it is risk-linked.',
    premium:'Full optionality at premium capex. Requires explicit board decision to pay for resilience.'
  };
  const worsens={base:[],faster:['Cost ↑','Confidence ↓','Risk ↑'],cheaper:['Schedule ↑','Confidence ↓','Risk ↑'],lower_risk:['Cost ↑','Schedule ↑'],premium:['Cost ↑↑']};
  const improves={base:[],faster:['Schedule ↓'],cheaper:['Cost ↓'],lower_risk:['Confidence ↑','Risk ↓'],premium:['Confidence ↑','Risk ↓']};
  return <div className="scenarioCompare upgraded">{scenarios.map(s=>{
    const active=s===scenario;
    const row=scenarioMatrix.find(x=>x.scenario===s)||{};
    const rCost=parseMoneyLocal(row.cost_p50||row.cost||'0');
    const rConf=parseInt(String(row.confidence_pct||row.confidence||'50'));
    const rSched=parseInt(String(row.schedule_months||'0'));
    const costD=bCost>0?Math.round((rCost-bCost)/bCost*100):0;
    const confD=bConf>0?rConf-bConf:0;
    const schedD=bSched>0?rSched-bSched:0;
    const rCol=(row.risk||'').toLowerCase().includes('high')?'#ef4444':(row.risk||'').toLowerCase().includes('medium')?'#f59e0b':'#10b981';
    const cCol=rConf>=65?'#10b981':rConf>=50?'#f59e0b':'#ef4444';
    const ws=worsens[s]||[]; const im=improves[s]||[];
    return <button key={s} className={active?'active':''} onClick={()=>generate(s,model?.prompt||prompt,model||projectContext)}>
      <b style={{textTransform:'uppercase',letterSpacing:'.06em',fontSize:'13px'}}>{(row.label||s).replace('_',' ')}</b>
      <strong style={{fontSize:'16px',display:'block',margin:'4px 0'}}>{row.cost_p50||row.cost||'—'}</strong>
      <div style={{fontSize:'11px',color:'#94a3b8',marginBottom:'4px'}}>{row.schedule_months||'—'} months · <span style={{color:cCol,fontWeight:'700'}}>{row.confidence_pct||row.confidence||'—'}%</span> · <span style={{color:rCol,fontWeight:'700'}}>{row.risk||'—'}</span></div>
      {s!=='base'&&bCost>0&&<div style={{display:'flex',gap:'5px',flexWrap:'wrap',marginBottom:'3px'}}>
        {costD!==0&&<span style={{background:costD>0?'rgba(239,68,68,0.15)':'rgba(16,185,129,0.12)',color:costD>0?'#ef4444':'#10b981',padding:'1px 6px',borderRadius:'2px',fontSize:'10px',fontWeight:'900'}}>{costD>0?'+':''}{costD}% cost</span>}
        {confD!==0&&<span style={{background:confD>0?'rgba(16,185,129,0.12)':'rgba(239,68,68,0.15)',color:confD>0?'#10b981':'#ef4444',padding:'1px 6px',borderRadius:'2px',fontSize:'10px',fontWeight:'900'}}>{confD>0?'+':''}{confD}pt conf</span>}
        {schedD!==0&&<span style={{background:schedD<0?'rgba(16,185,129,0.12)':'rgba(245,158,11,0.12)',color:schedD<0?'#10b981':'#f59e0b',padding:'1px 6px',borderRadius:'2px',fontSize:'10px',fontWeight:'900'}}>{schedD>0?'+':''}{schedD}mo</span>}
      </div>}
      {ws.length>0&&<div style={{fontSize:'9px',color:'#ef4444',fontWeight:'700',letterSpacing:'.05em'}}>WORSE: {ws.join(' · ')}</div>}
      {im.length>0&&<div style={{fontSize:'9px',color:'#10b981',fontWeight:'700',letterSpacing:'.05em'}}>BETTER: {im.join(' · ')}</div>}
      <em style={{fontSize:'10px',color:'#64748b',fontStyle:'normal',lineHeight:'1.3',display:'block',marginTop:'4px'}}>{active?'▶ ACTIVE — '+scenario.toUpperCase():tradeNotes[s]||''}</em>
    </button>;
  })}</div>;
})()}</Card><Card><h2>Buyer decision lens</h2>{['Base: balanced reference case for board challenge','Faster: more capex, lower confidence, shorter duration','Cheaper: lower authorisation number, longer schedule, higher residual risk','Lower Risk: higher reserve, longer duration, stronger confidence','Premium: resilience and optionality bought with visible capex premium'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}<h3>Current trade-off</h3><div className="triLens"><b>Gained</b>{(tradePack?.gained||[]).map(x=><span key={x}>{x}</span>)}<b>Sacrificed</b>{(tradePack?.sacrificed||[]).map(x=><span key={x}>{x}</span>)}<b>Exposed</b>{(tradePack?.exposed||[]).map(x=><span key={x}>{x}</span>)}</div></Card></section>}
        {tab === 'cost' && <section className="layout two"><Card><h2>Scenario cost bridge vs Base</h2><p className="chartCaption">This explains why the selected scenario is cheaper or more expensive than Base before showing the workbook lines.</p>{model?.stress_test_applied && <div style={{background:"rgba(141,247,255,0.05)",borderLeft:"2px solid #8df7ff",padding:"8px 12px",marginBottom:"8px",fontSize:"11px",color:"#8df7ff"}}>Stress test applied: {String(model.stress_test_applied).replace(/_/g," ")} — cost recalculated to {safeRender(model.cost_p50)}. The waterfall below reflects this change.</div>}{costWaterfall.map((x,i)=><div className={`reason ${x.kind==='total'?'deltaReason':''}`} key={i}><span>{i+1}</span><b>{x.driver}</b><br/>{x.kind==='total'?x.value:(x.value_bn>=0?'+':'−') + ' ' + x.value}</div>)}<div style={{margin:'8px 0',padding:'7px 12px',background:'rgba(141,247,255,0.04)',border:'1px solid rgba(141,247,255,0.1)',borderRadius:'3px',display:'flex',justifyContent:'space-between',alignItems:'center'}}>
              <div>
                <span style={{fontSize:'9px',fontWeight:'800',color:'#8df7ff',letterSpacing:'.1em'}}>UNIT COST BENCHMARK — </span>
                <span style={{fontSize:'10px',color:'#e2e8f0',fontWeight:'700'}}>{model.unit_rate_label?.metric||'programme unit'}</span>
              </div>
              <span style={{fontSize:'9px',color:'#64748b'}}>{model.unit_rate_label?.typical_range||'sector calibration applied'}</span>
            </div>
            <h3>Cost estimate workbook</h3><Table rows={costs} cols={[["cbs","CBS"],["description","Description"],["type","Type"],["unit_rate","Unit rate"],["p10_bn","P10"],["p50_bn","P50"],["p90_bn","P90"],["basis","Basis"]]} moneyCols={["p10_bn","p50_bn","p90_bn"]}/></Card><Card><h2>Cost composition</h2><p className="chartCaption">Direct, indirect and reserve are scenario-controlled and reconciled to selected P50. For the detailed uncertainty view use QCRA/QSRA.</p><ResponsiveContainer width="100%" height={320}><BarChart data={[{name:'Direct',value:direct},{name:'Indirect',value:indirect},{name:'Reserve',value:reserves}]}><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis dataKey="name"/><YAxis/><Tooltip/><Bar dataKey="value" fill="#8df7ff"/></BarChart></ResponsiveContainer></Card></section>}
        {tab === 'schedule' && <section className="layout two"><Card><h2>Schedule bridge vs Base</h2><p className="chartCaption">This is the month-by-month reason the scenario becomes faster or slower than Base.</p>{scheduleWaterfall.map((x,i)=><div className={`reason ${x.kind==='total'?'deltaReason':''}`} key={i}><span>{i+1}</span><b>{x.driver}</b><br/>{x.kind==='total'?`${x.months} months`:(x.months>=0?'+':'') + x.months + ' months'}</div>)}<h3>Scenario schedule logic</h3><Table rows={schedule} cols={[["activity_id","Activity"],["phase","Phase"],["activity","Name"],["predecessor","Pred"],["duration_months","Months"],["critical","Critical"],["basis","Basis"]]}/></Card><Card><h2>QSRA finish-date curve</h2><p className="chartCaption">P50 equals the headline schedule. P80/P90 show how severe the delivery tail becomes after the scenario trade-off.</p><div className="metrics"><div>P50<b>{qsra.p50} mo</b></div><div>P80<b>{qsra.p80} mo</b></div><div>P90<b>{qsra.p90} mo</b></div></div><ResponsiveContainer width="100%" height={280}><LineChart data={curve}><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis dataKey="percentile"/><YAxis/><Tooltip formatter={(v) => [`${v} months`, "QSRA finish date"]}/><ReferenceLine x={50} stroke="#ffffff88" label="P50 = headline"/><ReferenceLine x={80} stroke="#ffffff55" label="P80 = board risk"/><Line type="monotone" name="QSRA finish date" dataKey="schedule_months" stroke="#b18cff" strokeWidth={4}/></LineChart></ResponsiveContainer><div className="reason p80Translation"><span>1/5</span>{safeRender(p80Talk.schedule)}</div><div className="reason p80Translation"><span>!</span>{safeRender(p80Talk.board)}</div>{(model.monte_carlo?.curve_readout || []).slice(1).map((x,i)=><div className="reason" key={i}><span>{i+1}</span>{x}</div>)}</Card></section>}
        {tab === 'risk' && <section className="layout two"><Card><h2>Risk Register Pro</h2><p style={{fontSize:'11px',color:'#64748b',marginBottom:'8px'}}>Each risk has a cause (what triggers it), event (what happens), impact (cost/schedule consequence), probability, named owner, and mitigation. The top risks by expected monetary value drive the P80/P90 exposure in the QCRA chart.</p>{model?.stress_test_applied && <div style={{background:"rgba(239,68,68,0.08)",borderLeft:"2px solid #ef4444",padding:"6px 10px",marginBottom:"8px",fontSize:"11px",color:"#ef4444"}}>Stress test applied: risk posture has shifted. Confidence is now {model.confidence_pct}%. The risks below drove this position before the shock was applied.</div>}<Table rows={risks} cols={[['id','ID'],['risk','Risk'],['cause','Cause'],['event','Event'],['impact','Impact'],['probability_pct','Prob %'],['cbs','CBS'],['owner','Owner'],['mitigation','Mitigation']]}/></Card><Card><h2>Top exposure drivers</h2><ResponsiveContainer width="100%" height={380}><BarChart data={tornado.slice(0,8)} layout="vertical" margin={{left:10,right:20}}>
                <CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/>
                <XAxis type="number" tick={{fontSize:9,fill:'#64748b'}} tickFormatter={v=>v>0?`$${v.toFixed(1)}B`:''}/>
                <YAxis dataKey="driver" type="category" width={160} tick={{fontSize:9,fill:'#94a3b8'}}/>
                <Tooltip formatter={(v,n)=>[`$${Number(v||0).toFixed(2)}B EMV`,'Risk exposure']}/>
                <Bar dataKey="contribution" fill="#8df7ff" radius={[0,3,3,0]}/>
              </BarChart></ResponsiveContainer></Card></section>}
        {tab === 'monte' && <section className="layout two"><Card><h2>QCRA cost range curve</h2>{model?.stress_test_applied && <div style={{background:'rgba(245,158,11,0.08)',borderLeft:'2px solid #f59e0b',padding:'6px 10px',marginBottom:'8px',fontSize:'11px',color:'#f59e0b'}}>Stress test active: {String(model.stress_test_applied).replace(/_/g,' ')} — P50 updated to {safeRender(model.cost_p50)}. Download Export QCRA/QSRA to capture the stressed curves.</div>}<p style={{fontSize:'11px',color:'#64748b',marginBottom:'8px'}}>Probability range across 10,000+ simulations. P50 = the most likely outturn (headline number). P80 = 80% chance of coming in at this or less — this is the board's risk exposure. P90 = stress-case downside. Not a cashflow profile.</p><div className="metrics"><div>P50 headline<b>{safeRender(model.cost_p50)}</b></div><div>P80 risk exposure<b>{fmt(qcra.p80)}</b></div><div>P90 stress case<b>{fmt(qcra.p90)}</b></div></div><ResponsiveContainer width="100%" height={280}><AreaChart data={curve}><defs><linearGradient id="caseyG" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stopColor="#8df7ff" stopOpacity=".55"/><stop offset="1" stopColor="#8df7ff" stopOpacity="0"/></linearGradient></defs><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis dataKey="percentile"/><YAxis/><Tooltip formatter={(v) => [`$${Number(v).toFixed(1)}B`, "QCRA total outturn"]}/><ReferenceLine x={50} stroke="#ffffff88" label="P50 = headline"/><ReferenceLine x={80} stroke="#ffffff55" label="P80 = board risk"/><Area type="monotone" name="QCRA total outturn" dataKey="cost_bn" stroke="#8df7ff" fill="url(#caseyG)"/></AreaChart></ResponsiveContainer>{(model.monte_carlo?.curve_readout || []).slice(0,1).map((x,i)=><div className="reason" key={i}><span>{i+1}</span>{safeRender(x)}</div>)}<div className="reason p80Translation"><span>1/5</span>{safeRender(p80Talk.cost)}</div><div className="reason"><span>!</span>This curve is a probability distribution, not spend over time. The x-axis is confidence percentile. P50 equals the headline estimate; P80/P90 are board downside exposure.</div></Card><Card><h2>QSRA schedule range curve</h2><p className="chartCaption">P50 matches the headline duration. P80/P90 show the likely board conversation if critical path risk lands.</p><div className="metrics"><div>P50 headline<b>{qsra.p50} mo</b></div><div>P80 risk date<b>{qsra.p80} mo</b></div><div>P90 stress date<b>{qsra.p90} mo</b></div></div><ResponsiveContainer width="100%" height={280}><LineChart data={curve}><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis dataKey="percentile"/><YAxis/><Tooltip formatter={(v) => [`${v} months`, "QSRA finish date"]}/><ReferenceLine x={50} stroke="#ffffff88" label="P50 = headline"/><ReferenceLine x={80} stroke="#ffffff55" label="P80 = board risk"/><Line type="monotone" name="QSRA finish date" dataKey="schedule_months" stroke="#b18cff" strokeWidth={4}/></LineChart></ResponsiveContainer><div className="reason p80Translation"><span>1/5</span>{safeRender(p80Talk.schedule)}</div><div className="reason p80Translation"><span>!</span>{safeRender(p80Talk.board)}</div>{(model.monte_carlo?.curve_readout || []).map((x,i)=><div className="reason" key={i}><span>{i+1}</span>{safeRender(x)}</div>)}</Card></section>}
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
          <Card><h2>Gained / Sacrificed / Exposed</h2>
            <div className="triLens full"><b>Gained</b>{(tradePack?.gained||[]).map(x=><span key={x}>{x}</span>)}<b>Sacrificed</b>{(tradePack?.sacrificed||[]).map(x=><span key={x}>{x}</span>)}<b>Exposed</b>{(tradePack?.exposed||[]).map(x=><span key={x}>{x}</span>)}</div>
            <div className="reason"><span>!</span><b>Curve meaning</b><br/>{model.monte_carlo?.curve_interpretation || 'QCRA/QSRA shape reflects scenario uncertainty.'}</div>
          </Card>
          <Card><h2>Confidence Breakdown</h2>
            {(model.confidence_breakdown || []).map((x,i)=><div className="reason" key={i}><span>{i+1}</span><b>{x.driver}: {x.effect}</b><br/>{x.note}</div>)}
          </Card>
          <Card><h2>Top Decisions Required</h2>
            {(model.top_decisions_required || []).map((x,i)=><div className="reason" key={i}><span>{i+1}</span>{x}</div>)}
          </Card>
        </section>}

        {tab === 'mortality' && <section className="layout two">
          <Card><h2 style={{color:'#ef4444'}}>Programme Mortality Engine</h2>
            <p style={{fontSize:'11px',color:'#64748b',marginBottom:'10px'}}>Probability of programme cancellation or fundamental restructuring, with named precedents. No consultant produces this. CASEY generates it in 12 seconds.</p>
            {(()=>{
              const mort = model.programme_mortality_risk||'';
              const score = parseInt((mort.match(/(\d+)%/)||['','35'])[1])||35;
              const col = score>=70?'#ef4444':score>=45?'#f59e0b':score>=25?'#fbbf24':'#10b981';
              return <div style={{background:col+'10',border:`2px solid ${col}`,borderRadius:'5px',padding:'14px 18px',marginBottom:'12px'}}>
                <div style={{display:'flex',alignItems:'center',gap:'16px'}}>
                  <div style={{fontSize:'52px',fontWeight:'900',color:col,lineHeight:1}}>{score}%</div>
                  <div><div style={{fontSize:'10px',fontWeight:'800',letterSpacing:'.1em',color:col,marginBottom:'4px'}}>PROGRAMME MORTALITY RISK</div>
                  <div style={{fontSize:'12px',color:'#e2e8f0',lineHeight:'1.4',maxWidth:'380px'}}>{mort||'Run a project to see mortality risk.'}</div></div>
                </div>
              </div>;
            })()}
            <h3>Sector failure pattern</h3>
            <p style={{fontSize:'11px',color:'#94a3b8',lineHeight:'1.6',fontStyle:'italic'}}>{safeRender(model.if_this_fails||model.sector_failure_pattern)||'Run a project to see the named failure pattern.'}</p>
            <h3 style={{marginTop:'10px'}}>Reduce mortality risk</h3>
            {(model.second_order_contradictions||[]).map((c,i)=><div className="reason" key={i}><span style={{background:'rgba(239,68,68,0.15)',color:'#ef4444'}}>{i+1}</span>{c}</div>)}
          </Card>
          <Card><h2>CASEY Position vs Conventional</h2>
            {(()=>{
              const tvc=model.traditional_vs_casey||{};
              return <><div style={{padding:'10px 14px',background:'rgba(255,255,255,0.02)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:'4px',marginBottom:'8px'}}>
                <div style={{fontSize:'9px',fontWeight:'800',color:'#475569',marginBottom:'4px'}}>WHAT THE CONVENTIONAL REPORT SAYS</div>
                <p style={{fontSize:'11px',color:'#94a3b8',margin:0,lineHeight:'1.5'}}>{tvc.traditional_read||'Single-point P50, bar chart schedule, generic risk register. No confidence interval. No OBA.'}</p>
              </div>
              <div style={{padding:'10px 14px',background:'rgba(141,247,255,0.05)',border:'1px solid rgba(141,247,255,0.15)',borderRadius:'4px',marginBottom:'8px'}}>
                <div style={{fontSize:'9px',fontWeight:'800',color:'#8df7ff',marginBottom:'4px'}}>CASEY POSITION</div>
                <p style={{fontSize:'11px',color:'#e2e8f0',margin:0,lineHeight:'1.5'}}>{tvc.casey_read||safeRender(model.casey_position)||'Run a project to see the CASEY position.'}</p>
              </div>
              <div style={{padding:'10px 14px',background:'rgba(239,68,68,0.05)',border:'1px solid rgba(239,68,68,0.15)',borderRadius:'4px'}}>
                <div style={{fontSize:'9px',fontWeight:'800',color:'#ef4444',marginBottom:'4px'}}>WHAT THE CONSULTANT WILL NOT TELL YOU</div>
                <p style={{fontSize:'11px',color:'#fca5a5',margin:0,lineHeight:'1.5'}}>{tvc.what_the_consultant_wont_tell_you||'The primary failure mode for this sector is systematically excluded from advisory reports because naming it reduces fee scope.'}</p>
              </div></>;
            })()}
          </Card>
        </section>}

        {tab === 'market' && <section className="layout two">
          <Card><h2>Contractor Market Intelligence</h2>
            <p style={{fontSize:'11px',color:'#64748b',marginBottom:'10px'}}>Who can deliver this programme, order book depth, supply chain lead times and single-source risk. T&T and Jacobs charge £500/hr for this.</p>
            <div style={{marginBottom:'10px',padding:'10px 14px',background:'rgba(141,247,255,0.04)',border:'1px solid rgba(141,247,255,0.1)',borderRadius:'4px'}}>
              <div style={{fontSize:'9px',fontWeight:'800',color:'#8df7ff',marginBottom:'5px'}}>UNIT COST BENCHMARK FOR THIS SECTOR</div>
              <div style={{fontSize:'12px',fontWeight:'700',color:'#e2e8f0'}}>{model.unit_rate_label?.metric||'programme unit'}</div>
              <div style={{fontSize:'10px',color:'#64748b',marginTop:'2px'}}>{model.unit_rate_label?.typical_range||'Sector benchmark calibration applied'}</div>
            </div>
            <div style={{background:'rgba(255,255,255,0.03)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:'4px',padding:'10px 14px',marginBottom:'10px'}}>
              <div style={{fontSize:'9px',fontWeight:'800',color:'#8df7ff',marginBottom:'5px'}}>KEY PROCUREMENT FACTS</div>
              {(model.procurement_heatmap||[]).map((p2,i)=><div key={i} style={{display:'flex',justifyContent:'space-between',padding:'4px 0',borderBottom:'1px solid rgba(255,255,255,0.04)',fontSize:'11px'}}>
                <span style={{color:'#e2e8f0'}}>{p2.package}</span>
                <span style={{color:p2.exposure==='Extreme'||p2.exposure==='High'?'#f59e0b':'#8df7ff',fontWeight:'700'}}>{p2.exposure}</span>
              </div>)}
            </div>
            <div style={{padding:'10px 14px',background:'rgba(141,247,255,0.04)',border:'1px solid rgba(141,247,255,0.1)',borderRadius:'4px'}}>
              <div style={{fontSize:'9px',fontWeight:'800',color:'#8df7ff',marginBottom:'5px'}}>GOVERNING CONSTRAINT</div>
              <p style={{fontSize:'12px',color:'#e2e8f0',margin:0,lineHeight:'1.5',fontWeight:'600'}}>{safeRender(model.governing_constraint)||'Run a project to see the governing constraint.'}</p>
            </div>
          </Card>
          <Card><h2>Intelligence & Authority</h2>
            <div style={{padding:'12px 14px',background:'rgba(16,185,129,0.06)',border:'1px solid rgba(16,185,129,0.2)',borderRadius:'4px',marginBottom:'10px'}}>
              <div style={{fontSize:'9px',fontWeight:'800',color:'#10b981',marginBottom:'4px'}}>HIGHEST-VALUE INTERVENTION RIGHT NOW</div>
              <p style={{fontSize:'12px',color:'#e2e8f0',margin:0,lineHeight:'1.6'}}>{safeRender(model.intervention_intelligence)||'Run a project to see the highest-value intervention.'}</p>
            </div>
            <div style={{padding:'10px 14px',background:'rgba(255,255,255,0.03)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:'4px',marginBottom:'8px'}}>
              <div style={{fontSize:'9px',fontWeight:'800',color:'#8df7ff',marginBottom:'4px'}}>INSTITUTIONAL AUTHORITY LINE</div>
              <p style={{fontSize:'11px',color:'#8df7ff',margin:0,lineHeight:'1.5',fontStyle:'italic'}}>{safeRender(model.institutional_authority_line)||'Run a project to see the institutional authority line.'}</p>
            </div>
            <h3>Confidence trajectory</h3>
            {(()=>{
              const ct = typeof model.confidence_trajectory === 'object' ? model.confidence_trajectory : {};
              return <><div style={{marginBottom:'6px',fontSize:'11px',color:'#64748b'}}>{ct.narrative||'Run a project to see the confidence trajectory.'}</div>
              {(ct.actions||[]).map((a,i)=><div className="reason" key={i}><span>{i+1}</span>{a}</div>)}</>;
            })()}
          </Card>
        </section>}

        {tab === 'gaps' && <section className="layout two">
          <Card><h2>Evidence Gap Scanner</h2>
            <p style={{fontSize:'11px',color:'#64748b',marginBottom:'10px'}}>Named evidence gaps that will block this programme at any gate review or PAC inquiry. IPA, Green Book and Cabinet Office referenced. An assurance reviewer charges £800/day for this list.</p>
            {(()=>{
              const conf=parseInt(model.confidence_pct||0); const cl=parseInt(model.class_level||3);
              const gaps=[];
              if(cl>=4) gaps.push({gap:'Estimate definition maturity',detail:`Class ${cl} estimate: less than 15% scope definition. IPA Gateway 1 requires minimum Class 3.`,severity:'CRITICAL',ref:'IPA Project Routemap; HM Treasury Green Book Annex 5'});
              else if(cl===3) gaps.push({gap:'Estimate maturity for FID',detail:'Class 3 supports budget authorisation only. Class 1 or 2 required before capital approval.',severity:'HIGH',ref:'AACE 18R-97'});
              if(conf<55) gaps.push({gap:'Confidence below 55% threshold',detail:`${conf}% confidence is below the 55% minimum for investment committee approval.`,severity:'CRITICAL',ref:'IPA Governance for Major Projects'});
              const unowned=(model.risks||[]).filter(r=>!r.owner||r.owner==='TBC'||r.owner==='—');
              if(unowned.length>0) gaps.push({gap:`${unowned.length} risks without named owners`,detail:'IPA requires every material risk to have a named accountable owner.',severity:'HIGH',ref:'IPA Risk Management Guidance 2020'});
              gaps.push({gap:'OBA uplift disclosure',detail:'HM Treasury Green Book requires OBA disclosure in all public programme board cases.',severity:'HIGH',ref:'HM Treasury Green Book 2022, Annex 4'});
              if(gaps.length===0) gaps.push({gap:'No critical gaps at current evidence maturity',detail:'Evidence consistent with stated confidence level.',severity:'LOW',ref:''});
              return gaps.map((g,i)=>{
                const sc=g.severity==='CRITICAL'?'#ef4444':g.severity==='HIGH'?'#f59e0b':'#10b981';
                return <div key={i} style={{marginBottom:'8px',padding:'10px 14px',background:sc+'08',border:`1px solid ${sc}33`,borderLeft:`3px solid ${sc}`,borderRadius:'4px'}}>
                  <div style={{display:'flex',justifyContent:'space-between',marginBottom:'3px'}}>
                    <span style={{fontSize:'12px',fontWeight:'800',color:'#e2e8f0'}}>{g.gap}</span>
                    <span style={{background:sc+'18',color:sc,fontSize:'8px',fontWeight:'800',padding:'2px 7px',borderRadius:'2px',flexShrink:0,marginLeft:'8px'}}>{g.severity}</span>
                  </div>
                  <p style={{fontSize:'10px',color:'#94a3b8',margin:'0 0 3px',lineHeight:'1.4'}}>{g.detail}</p>
                  {g.ref&&<div style={{fontSize:'9px',color:'#334155'}}>Ref: {g.ref}</div>}
                </div>;
              });
            })()}
          </Card>
          <Card><h2>Gate Review & OBA</h2>
            {(()=>{
              const gate=model.gate_review_readiness||{}; const oba=model.optimism_bias_assessment||{};
              return <><div style={{padding:'10px 14px',background:'rgba(255,255,255,0.03)',borderRadius:'4px',marginBottom:'8px'}}>
                <div style={{fontSize:'9px',fontWeight:'800',color:'#8df7ff',marginBottom:'3px'}}>GATE REVIEW VERDICT</div>
                <div style={{fontSize:'16px',fontWeight:'900',color: gate.overall_verdict==='READY'?'#10b981':gate.overall_verdict==='BLOCKED'?'#ef4444':'#f59e0b'}}>{gate.overall_verdict||'—'}</div>
                <p style={{fontSize:'11px',color:'#64748b',margin:'3px 0 0'}}>{gate.current_gate_readiness||'—'}</p>
              </div>
              <div style={{padding:'10px 14px',background:'rgba(141,247,255,0.04)',border:'1px solid rgba(141,247,255,0.1)',borderRadius:'4px',marginBottom:'8px'}}>
                <div style={{fontSize:'9px',fontWeight:'800',color:'#8df7ff',marginBottom:'3px'}}>OBA ADJUSTED OUTTURN</div>
                <div style={{fontSize:'16px',fontWeight:'900',color:'#e2e8f0'}}>{oba.oba_adjusted_p50||'—'}</div>
                <p style={{fontSize:'10px',color:'#64748b',margin:'3px 0 0'}}>{oba.oba_source?.slice(0,120)||''}</p>
              </div>
              <h3>Next gate actions</h3>
              {(gate.next_gate_actions||[]).map((a,i)=><div className="reason" key={i}><span>{i+1}</span>{a}</div>)}
              </>;
            })()}
          </Card>
        </section>}

        {tab === 'paths' && <section className="layout two">
          <Card><h2>Three Delivery Paths</h2>
            <p style={{fontSize:'11px',color:'#64748b',marginBottom:'10px'}}>Three strategies for the same brief. Consultants avoid producing this. CASEY shows all three with board questions.</p>
            {[
              {label:'Conventional Design-Bid-Build',c:'#10b981',cm:1.0,sm:1.0,cd:0,desc:'Sequential design, competitive tender, then construction. Maximum market testing, minimum commercial risk.',risk:'Schedule slippage from sequential procurement and late design changes.',for:'Public accountability paramount — audit trail and competition required.'},
              {label:'Early Contractor Involvement (ECI)',c:'#8df7ff',cm:1.06,sm:0.82,cd:8,desc:'Contractor at design stage. De-risks buildability, supply chain. Higher initial cost, lower outturn risk, faster delivery.',risk:'Lock-in to sole contractor post-PCSA. Requires open-book commercial controls.',for:'High technical complexity, tight possessions or supply chain constraints.'},
              {label:'Public-Private Partnership / Concession',c:'#b18cff',cm:0.88,sm:0.90,cd:-5,desc:'Private finance with long-term concession. Lower public capex, but WLC premium and inflexibility.',risk:'Concession inflexibility. Private WLC higher than public equivalent. HMT VFM test required.',for:'Private sector can absorb demand risk and operational performance is measurable.'},
            ].map((path,i)=>{
              const p50v=parseFloat((model.cost_p50||'0').replace('$','').replace('B',''))||0;
              const mo=parseInt(model.schedule_months||36)||36; const conf=parseInt(model.confidence_pct||60)||60;
              return <div key={i} style={{marginBottom:'12px',padding:'12px 14px',background:'rgba(255,255,255,0.02)',border:`1px solid ${path.c}25`,borderLeft:`3px solid ${path.c}`,borderRadius:'4px'}}>
                <div style={{fontSize:'12px',fontWeight:'800',color:path.c,marginBottom:'4px'}}>{path.label}</div>
                <p style={{fontSize:'11px',color:'#94a3b8',margin:'0 0 8px',lineHeight:'1.4'}}>{path.desc}</p>
                <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:'6px',marginBottom:'8px'}}>
                  {[['P50',`$${(p50v*path.cm).toFixed(1)}B`],['Schedule',`${Math.round(mo*path.sm)} mo`],['Conf',`${Math.max(10,Math.min(96,conf+path.cd))}%`]].map(([l,v])=>
                    <div key={l} style={{textAlign:'center',padding:'5px',background:'rgba(255,255,255,0.04)',borderRadius:'3px'}}>
                      <div style={{fontSize:'9px',color:'#475569'}}>{l}</div>
                      <div style={{fontSize:'12px',fontWeight:'800',color:'#e2e8f0'}}>{v}</div>
                    </div>)}
                </div>
                <div style={{fontSize:'10px',color:'#ef4444',marginBottom:'2px'}}><b>Key risk:</b> {path.risk}</div>
                <div style={{fontSize:'10px',color:'#64748b'}}><b>Best for:</b> {path.for}</div>
              </div>;
            })}
          </Card>
          <Card><h2>How to Choose</h2>
            {[['Time is the primary constraint','Choose ECI — accept cost premium to compress schedule.'],['Public accountability is critical','Choose conventional design-bid-build for maximum audit trail.'],['Technical complexity is high','Choose ECI — contractor at design stage de-risks buildability.'],['Capital preservation matters','Consider PPP/concession — check HMT VFM test first.']].map(([q,a])=>
              <div key={q} style={{marginBottom:'8px',paddingBottom:'8px',borderBottom:'1px solid rgba(255,255,255,0.04)'}}>
                <div style={{fontSize:'10px',fontWeight:'700',color:'#e2e8f0',marginBottom:'1px'}}>{q}</div>
                <div style={{fontSize:'10px',color:'#64748b'}}>{a}</div>
              </div>)}
            <div style={{marginTop:'10px',padding:'10px',background:'rgba(245,158,11,0.05)',border:'1px solid rgba(245,158,11,0.15)',borderRadius:'4px',fontSize:'10px',color:'#fde68a',lineHeight:'1.5'}}>
              Use the Advisor tab to ask "What if we go ECI?" — CASEY reruns the model with that constraint applied.
            </div>
          </Card>
        </section>}

        {tab === 'causal' && <section className="layout two"><CausalGraph model={model}/><BenchmarkIntelligence model={model}/><Card><h2>Evidence Mode: {viewMode}</h2>{evidenceScorecard(model).map((x,i)=><div className="reason" key={x.name}><span>{i+1}</span><b>{x.name}: {Math.round(x.score)}%</b><br/>{x.note}</div>)}</Card></section>}

        {tab === 'outputs' && <section className="layout two"><Card><h2>Generated Artefacts</h2><p>The public demo previews the intelligence pack. Enterprise access unlocks the live generated controls deliverables.</p><div className="exports v50Exports lockedExports">
          <button onClick={() => download('/export/workbook', model, `${model.id || 'casey'}_COST_WORKBOOK.xlsx`)}><FileSpreadsheet/> Generate Cost Model XLSX</button>
          <button onClick={() => download('/export/risk-register', model, `${model.id || 'casey'}_RISK_REGISTER.xlsx`)}><Database/> Generate Risk Register XLSX</button>
          <button onClick={() => download('/export/xer', model, `${model.id || 'casey'}_PRA_SCHEDULE.xer`)}><Workflow/> Generate PRA Schedule XER</button>
          <button onClick={() => download('/export/qcra-qsra', model, `${model.id || 'casey'}_QCRA_QSRA.xlsx`)}><BarChart3/> Generate QCRA/QSRA Pack</button>
          <button onClick={() => download('/export/json', model, `${model.id || 'casey'}_AUDIT_MODEL.json`)}><Brain/> Generate Audit File JSON</button>
          <button onClick={() => download('/export/all', model, `${model.id || 'casey'}_FULL_BOARD_PACK.zip`)}><Download/> Generate Full Pack ZIP</button>
          <a className="contactBtn" href={emailLink}><Mail/> Request Enterprise Review</a></div></Card><Card><h2>What the pack delivers</h2>{['Executive control centre with project, scenario, class, level and confidence clearly identified','Scenario comparison covering Base, Faster, Cheaper, Lower Risk and Premium cases','Selected estimate class plus all class levels for audit and challenge','Direct, indirect and reserve cost views with QCRA cost curve and cost tornado','All schedule levels with QSRA schedule curve and schedule tornado','Risk register with cause, event, impact, owner, mitigation, trigger and quantified likelihood','Basis of Estimate, assumptions, exclusions and benchmark validation','Commercial next steps: buyer action, procurement challenge and board decision path'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card></section>}

        {tab === 'assurance' && <><IncumbentPressurePanel model={model} direct={direct} indirect={indirect} reserves={reserves} reconcileCheck={reconcileCheck}/><section className="layout two"><Card><h2>Assurance room weapons</h2>{['Open with the P80/P90 exposure, not the headline P50.','Ask which evidence package retires the governing constraint.','Force every mitigation to name owner, trigger, residual exposure and date.','Show scenario trade-offs live before anyone can defend a single-point estimate.','Export the audit model immediately so the conversation moves from opinion to traceability.'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card><Card><h2>Why CASEY changes the conversation</h2>{['CASEY recalculates cost, schedule, confidence and board posture from one source of truth in seconds.','Every scenario is a complete recalculation — not a slide edit.','The system surfaces contradictions rather than polishing the management narrative.','Static reports become live investment-committee intelligence.'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card></section><section className="layout one"><ProgrammeHealthSignal onRunHealthCheck={runHealthCheck}/></section></>}

        {tab === 'advisor' && <>
          {/* INSTITUTIONAL AUTHORITY LINE — the one sentence */}
          {model?.institutional_authority_line && <section className="layout one"><div style={{background:'linear-gradient(135deg,rgba(141,247,255,0.06),rgba(177,140,255,0.06))',border:'1px solid rgba(141,247,255,0.2)',borderRadius:'6px',padding:'14px 20px',display:'flex',gap:'12px',alignItems:'flex-start'}}>
            <span style={{color:'#8df7ff',fontWeight:'900',fontSize:'10px',letterSpacing:'.15em',flexShrink:0,paddingTop:'2px'}}>AUTHORITY LINE</span>
            <p style={{color:'#e2e8f0',fontSize:'13px',lineHeight:'1.6',margin:0,fontStyle:'italic'}}>{safeRender(model.institutional_authority_line)}</p>
          </div></section>}

          {/* PROGRAMME MORTALITY + CONFIDENCE TRAJECTORY */}
          {model?.programme_mortality_risk && <section className="layout two">
            <Card><h2>Programme mortality risk</h2>
              <div style={{display:'flex',alignItems:'baseline',gap:'8px',marginBottom:'8px'}}>
                <span style={{fontSize:'48px',fontWeight:'900',color:model.programme_mortality_risk.pct>60?'#ef4444':model.programme_mortality_risk.pct>35?'#f59e0b':'#10b981'}}>{model.programme_mortality_risk.pct}%</span>
                <span style={{fontSize:'12px',fontWeight:'800',color:model.programme_mortality_risk.pct>60?'#ef4444':model.programme_mortality_risk.pct>35?'#f59e0b':'#10b981'}}>{model.programme_mortality_risk.label}</span>
              </div>
              <p style={{fontSize:'12px',color:'#94a3b8',lineHeight:'1.5'}}>{safeRender(model.programme_mortality_risk.narrative)}</p>
              {model.programme_mortality_risk.comparable && <p style={{fontSize:'11px',color:'#64748b',marginTop:'8px',fontStyle:'italic'}}>{safeRender(model.programme_mortality_risk.comparable)}</p>}
            </Card>
            <Card><h2>Confidence trajectory</h2>
              {model?.confidence_trajectory && <>
                <div style={{display:'flex',alignItems:'center',gap:'10px',marginBottom:'8px'}}>
                  <span style={{fontSize:'18px',fontWeight:'900',color:model.confidence_trajectory.direction==='DECLINING'?'#ef4444':model.confidence_trajectory.direction==='IMPROVING'?'#10b981':'#f59e0b'}}>{model.confidence_trajectory.direction}</span>
                </div>
                <p style={{fontSize:'12px',color:'#94a3b8',lineHeight:'1.5'}}>{safeRender(model.confidence_trajectory.narrative)}</p>
                <p style={{fontSize:'11px',color:'#64748b',marginTop:'8px',borderTop:'1px solid rgba(255,255,255,0.06)',paddingTop:'8px'}}>{safeRender(model.confidence_trajectory.next_gate)}</p>
              </>}
            </Card>
          </section>}

          {/* BEHAVIOURAL FORECAST + INTERVENTION INTELLIGENCE */}
          {model?.behavioural_forecast && <section className="layout two">
            <Card><h2>Behavioural forecast</h2>
              <p style={{color:'#8df7ff',fontStyle:'italic',lineHeight:'1.6',fontSize:'13px'}}>{safeRender(model.behavioural_forecast)}</p>
              <p style={{fontSize:'11px',color:'#475569',marginTop:'8px'}}>Based on comparable {safeRender(model.subsector)} programmes — what CASEY expects to happen next if the governing constraint is not evidenced before approval.</p>
            </Card>
            <Card><h2>Intervention intelligence</h2>
              {(model?.intervention_intelligence||model?.governance_challenges||[]).map((x,i)=><div className="reason" key={i} style={{borderLeft:'2px solid rgba(141,247,255,0.3)',paddingLeft:'10px',marginBottom:'6px'}}><span style={{color:'#8df7ff',fontWeight:'800',marginRight:'6px'}}>{i+1}.</span>{safeRender(x)}</div>)}
            </Card>
          </section>}

          {/* TRADITIONAL vs CASEY */}
          {model?.traditional_vs_casey?.casey && <section className="layout two">
            <Card style={{borderLeft:'2px solid rgba(239,68,68,0.4)'}}><h2 style={{color:'#ff6b7d'}}>What a conventional report says</h2>
              <p style={{color:'#94a3b8',fontStyle:'italic',lineHeight:'1.6'}}>{safeRender(model.traditional_vs_casey.traditional)}</p>
              <p style={{fontSize:'11px',color:'#475569',marginTop:'8px',borderTop:'1px solid rgba(255,255,255,0.06)',paddingTop:'8px'}}>{safeRender(model.traditional_vs_casey.incumbent_line)}</p>
            </Card>
            <Card style={{borderLeft:'2px solid rgba(141,247,255,0.4)'}}><h2 style={{color:'#8df7ff'}}>What CASEY reads underneath</h2>
              <p style={{color:'#e2e8f0',lineHeight:'1.6'}}>{safeRender(model.traditional_vs_casey.casey)}</p>
              <p style={{fontSize:'11px',color:'#64748b',marginTop:'8px',borderTop:'1px solid rgba(255,255,255,0.06)',paddingTop:'8px'}}>{safeRender(model.traditional_vs_casey.gap)}</p>
            </Card>
          </section>}

          {/* BOARD ATTACK SIMULATION */}
          {(model?.board_attack_simulation||[]).length > 0 && <section className="layout one"><Card><h2>Board attack simulation — the 5 questions this board will ask</h2>
            <p style={{fontSize:'11px',color:'#64748b',marginBottom:'12px'}}>These are sector-specific, programme-specific challenges that a serious investment committee will table. CASEY generates each one from live model data — your actual P50, P80, sector, and governing constraints. Not a template.</p>
            {(model.board_attack_simulation||[]).map((q,i)=><div key={i} style={{display:'flex',gap:'10px',padding:'10px 0',borderBottom:'1px solid rgba(255,255,255,0.05)'}}>
              <span style={{color:'#f59e0b',fontWeight:'900',flexShrink:0,fontSize:'11px',paddingTop:'1px'}}>{i+1}.</span>
              <span style={{color:'#e2e8f0',lineHeight:'1.5',fontSize:'13px'}}>{safeRender(q)}</span>
            </div>)}
          </Card></section>}

          {/* LOCATION + FINANCING CONTEXT */}
          {model?.location_context?.framework && <section className="layout two">
            <Card><h2>📍 Location intelligence</h2>
              <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'8px',marginBottom:'10px'}}>
                <div><div style={{fontSize:'9px',color:'#64748b',fontWeight:'800',letterSpacing:'.1em',marginBottom:'3px'}}>LOCATION</div><div style={{fontSize:'14px',color:'#e2e8f0',fontWeight:'700'}}>{safeRender(model.location||'Global')}</div></div>
                <div><div style={{fontSize:'9px',color:'#64748b',fontWeight:'800',letterSpacing:'.1em',marginBottom:'3px'}}>CURRENCY</div><div style={{fontSize:'14px',color:'#8df7ff',fontWeight:'700'}}>{safeRender(model.location_context.currency)}</div></div>
              </div>
              <div style={{marginBottom:'8px'}}><div style={{fontSize:'9px',color:'#64748b',fontWeight:'800',letterSpacing:'.1em',marginBottom:'3px'}}>REGULATORY FRAMEWORK</div><div style={{fontSize:'12px',color:'#cbd5e1',lineHeight:'1.5'}}>{safeRender(model.location_context.framework)}</div></div>
              <div style={{marginBottom:'8px'}}><div style={{fontSize:'9px',color:'#64748b',fontWeight:'800',letterSpacing:'.1em',marginBottom:'3px'}}>APPROVAL BODY</div><div style={{fontSize:'12px',color:'#cbd5e1'}}>{safeRender(model.location_context.approval_body)}</div></div>
              <div style={{padding:'8px',background:'rgba(245,158,11,0.06)',borderRadius:'3px',border:'1px solid rgba(245,158,11,0.15)'}}><div style={{fontSize:'9px',color:'#f59e0b',fontWeight:'800',letterSpacing:'.1em',marginBottom:'3px'}}>OBA NOTE FOR THIS LOCATION</div><div style={{fontSize:'11px',color:'#94a3b8',lineHeight:'1.5'}}>{safeRender(model.location_context.optimism_bias_note)}</div></div>
            </Card>
            {model?.financing_context && <Card><h2>💰 Financing context</h2>
              <div style={{marginBottom:'8px'}}><div style={{fontSize:'9px',color:'#64748b',fontWeight:'800',letterSpacing:'.1em',marginBottom:'3px'}}>PRIMARY SOURCE</div><div style={{fontSize:'13px',color:'#8df7ff',fontWeight:'700'}}>{safeRender(model.financing_context.primary_source)}</div></div>
              <div style={{marginBottom:'8px'}}><div style={{fontSize:'9px',color:'#64748b',fontWeight:'800',letterSpacing:'.1em',marginBottom:'3px'}}>SECONDARY SOURCE</div><div style={{fontSize:'12px',color:'#cbd5e1'}}>{safeRender(model.financing_context.secondary_source)}</div></div>
              <div style={{marginBottom:'8px'}}><div style={{fontSize:'9px',color:'#64748b',fontWeight:'800',letterSpacing:'.1em',marginBottom:'3px'}}>STRUCTURE</div><div style={{fontSize:'12px',color:'#cbd5e1'}}>{safeRender(model.financing_context.structure)}</div></div>
              <div style={{marginBottom:'8px',padding:'6px 8px',background:'rgba(239,68,68,0.06)',borderRadius:'3px',border:'1px solid rgba(239,68,68,0.12)'}}><div style={{fontSize:'9px',color:'#ef4444',fontWeight:'800',letterSpacing:'.1em',marginBottom:'3px'}}>CURRENCY RISK</div><div style={{fontSize:'11px',color:'#fca5a5'}}>{safeRender(model.financing_context.currency_risk)}</div></div>
              <div style={{marginBottom:'8px'}}><div style={{fontSize:'9px',color:'#64748b',fontWeight:'800',letterSpacing:'.1em',marginBottom:'3px'}}>PROCUREMENT RULES</div><div style={{fontSize:'11px',color:'#64748b'}}>{safeRender(model.financing_context.procurement_rules)}</div></div>
              <div style={{padding:'6px 8px',background:'rgba(141,247,255,0.04)',borderRadius:'3px',border:'1px solid rgba(141,247,255,0.1)'}}><div style={{fontSize:'11px',color:'#94a3b8',lineHeight:'1.5',fontStyle:'italic'}}>{safeRender(model.financing_context.sector_note)}</div></div>
              <div style={{marginTop:'8px',padding:'6px 8px',background:'rgba(245,158,11,0.06)',borderRadius:'3px'}}><div style={{fontSize:'9px',color:'#f59e0b',fontWeight:'800',letterSpacing:'.1em',marginBottom:'3px'}}>BANKABILITY</div><div style={{fontSize:'12px',color:'#fcd34d',fontWeight:'700'}}>{safeRender(model.financing_context.bankability_verdict)}</div></div>
            </Card>}
          </section>}

          {/* GATE REVIEW READINESS */}
          {model?.gate_review_readiness && <section className="layout two">
            <Card><h2>Gate review readiness</h2>
              <div style={{display:'flex',gap:'12px',alignItems:'center',marginBottom:'10px'}}>
                <span style={{fontSize:'24px',fontWeight:'900',color:model.gate_review_readiness.overall_verdict==='READY'?'#10b981':model.gate_review_readiness.overall_verdict==='CONDITIONAL'?'#f59e0b':'#ef4444'}}>{model.gate_review_readiness.overall_verdict}</span>
                <span style={{fontSize:'13px',color:'#94a3b8'}}>at {safeRender(model.gate_review_readiness.current_gate_readiness)}</span>
              </div>
              <p style={{fontSize:'12px',color:'#64748b',marginBottom:'10px',lineHeight:'1.5'}}>{safeRender(model.gate_review_readiness.ipa_alignment)}</p>
              <p style={{fontSize:'11px',color:'#f59e0b',fontStyle:'italic'}}>{safeRender(model.gate_review_readiness.critical_gate_risk)}</p>
              <div style={{marginTop:'10px'}}>
                <div style={{fontSize:'10px',color:'#475569',fontWeight:'800',letterSpacing:'.1em',marginBottom:'6px'}}>NEXT GATE — EVIDENCE REQUIRED</div>
                {(model.gate_review_readiness.next_gate_actions||[]).map((a,i)=><div key={i} style={{fontSize:'11px',color:'#cbd5e1',padding:'3px 0',borderBottom:'1px solid rgba(255,255,255,0.04)',display:'flex',gap:'6px'}}><span style={{color:'#8df7ff',flexShrink:0}}>→</span>{safeRender(a)}</div>)}
              </div>
            </Card>
            <Card><h2>Optimism bias assessment</h2>
              {model?.optimism_bias_assessment && <>
                <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'8px',marginBottom:'10px'}}>
                  <div style={{background:'rgba(239,68,68,0.08)',padding:'8px',borderRadius:'4px',border:'1px solid rgba(239,68,68,0.2)'}}>
                    <div style={{fontSize:'9px',color:'#ef4444',fontWeight:'900',letterSpacing:'.1em'}}>STATED P50</div>
                    <div style={{fontSize:'18px',fontWeight:'900',color:'#fca5a5'}}>{safeRender(model.cost_p50)}</div>
                    <div style={{fontSize:'9px',color:'#64748b'}}>{safeRender(model.schedule)}</div>
                  </div>
                  <div style={{background:'rgba(245,158,11,0.08)',padding:'8px',borderRadius:'4px',border:'1px solid rgba(245,158,11,0.2)'}}>
                    <div style={{fontSize:'9px',color:'#f59e0b',fontWeight:'900',letterSpacing:'.1em'}}>OBA ADJUSTED</div>
                    <div style={{fontSize:'18px',fontWeight:'900',color:'#fcd34d'}}>{safeRender(model.optimism_bias_assessment.oba_adjusted_p50)}</div>
                    <div style={{fontSize:'9px',color:'#64748b'}}>{safeRender(model.optimism_bias_assessment.oba_adjusted_schedule)}</div>
                  </div>
                </div>
                <p style={{fontSize:'11px',color:'#94a3b8',lineHeight:'1.5',marginBottom:'6px'}}>{safeRender(model.optimism_bias_assessment.verdict)}</p>
                <p style={{fontSize:'10px',color:'#475569',fontStyle:'italic'}}>{safeRender(model.optimism_bias_assessment.oba_source)}</p>
                <p style={{fontSize:'11px',color:'#f59e0b',marginTop:'8px',padding:'6px 8px',background:'rgba(245,158,11,0.06)',borderRadius:'3px',lineHeight:'1.5'}}>{safeRender(model.optimism_bias_assessment.board_challenge)}</p>
              </>}
            </Card>
          </section>}

          {/* SECOND ORDER CONTRADICTIONS */}
          {(model?.second_order_contradictions||[]).length > 0 && <section className="layout one"><Card><h2>Second-order contradictions</h2>
            <p style={{fontSize:'11px',color:'#64748b',marginBottom:'10px'}}>What the optimistic case creates downstream — risks that appear only when the preferred scenario is examined under pressure.</p>
            {(model.second_order_contradictions||[]).map((x,i)=><div key={i} className="reason" style={{borderLeft:'2px solid rgba(245,158,11,0.4)',paddingLeft:'10px',marginBottom:'6px'}}><span style={{color:'#f59e0b',fontWeight:'800',marginRight:'6px'}}>{i+1}.</span>{safeRender(x)}</div>)}
          </Card></section>}

          {/* PROCUREMENT HEATMAP */}
          {(model?.procurement_heatmap||[]).length > 0 && <section className="layout one"><Card><h2>Procurement packages</h2>
            <p style={{fontSize:'11px',color:'#64748b',marginBottom:'12px'}}>CASEY maps every procurement package from sector intelligence and location context. Single-source flags identify the highest commercial exposure in each package — these are the items that can break a programme date.</p>
            <div style={{overflowX:'auto'}}>
              <table style={{width:'100%',borderCollapse:'collapse',fontSize:'11px'}}>
                <thead><tr style={{borderBottom:'1px solid rgba(255,255,255,0.1)'}}>
                  {['Package','Status','Value est.','Lead time','Single source','Risk','Owner'].map(h=><th key={h} style={{padding:'6px 8px',textAlign:'left',color:'#64748b',fontWeight:'800',letterSpacing:'.08em',fontSize:'10px'}}>{h}</th>)}
                </tr></thead>
                <tbody>{(model.procurement_heatmap||[]).map((p,i)=><tr key={i} style={{borderBottom:'1px solid rgba(255,255,255,0.04)',background:i%2===0?'rgba(255,255,255,0.01)':'transparent'}}>
                  <td style={{padding:'7px 8px',color:'#e2e8f0',fontWeight:'700',maxWidth:'160px'}}>{p.package}</td>
                  <td style={{padding:'7px 8px'}}><span style={{background:p.status==='Active'?'rgba(239,68,68,0.12)':'rgba(245,158,11,0.1)',color:p.status==='Active'?'#fca5a5':'#fde68a',borderRadius:'3px',padding:'2px 7px',fontSize:'10px',fontWeight:'800'}}>{p.status}</span></td>
                  <td style={{padding:'7px 8px',color:'#8df7ff',fontSize:'11px',fontWeight:'600'}}>{p.value_est||'—'}</td>
                  <td style={{padding:'7px 8px',color:'#94a3b8'}}>{p.lead_time||'—'}</td>
                  <td style={{padding:'7px 8px',textAlign:'center'}}>{p.single_source_risk ? <span style={{color:'#ef4444',fontWeight:'900',fontSize:'14px'}}>⚠</span> : <span style={{color:'#10b981'}}>✓</span>}</td>
                  <td style={{padding:'7px 8px',color:'#64748b',maxWidth:'200px',lineHeight:'1.4',fontSize:'10px'}}>{p.risk}</td>
                  <td style={{padding:'7px 8px',color:'#475569',fontSize:'10px'}}>{p.owner}</td>
                </tr>)}
                </tbody>
              </table>
            </div>
          </Card></section>}

          {/* HISTORICAL FAILURE PATTERN */}
          {model?.if_this_fails && <section className="layout one"><Card style={{borderLeft:'3px solid rgba(239,68,68,0.5)',background:'rgba(239,68,68,0.03)'}}><h2 style={{color:'#ff6b7d'}}>Historical failure pattern for this sector</h2>
            <p style={{fontSize:'11px',color:'#64748b',marginBottom:'10px'}}>Named programmes, real cost growth figures, actual failure modes. This is the sector intelligence that makes a green dashboard insufficient — the historical pattern that governs what happens next if the governing constraint is not closed.</p>
            <p style={{color:'#e2e8f0',lineHeight:'1.7',fontSize:'13px'}}>{safeRender(model.if_this_fails)}</p>
          </Card></section>}

          {/* CRITICAL PATH + NEAR-CRITICAL NARRATIVE */}
          {(model?.critical_path_narrative||[]).length > 0 && <section className="layout two">
            <Card><h2>🔴 Critical path — what must be evidenced before approval</h2>
              <p style={{fontSize:'11px',color:'#64748b',marginBottom:'10px'}}>These are the activities CASEY identifies as near-critical in the sector causal graph. Each one needs a named owner and evidence closure date before board approval.</p>
              {(model.critical_path_narrative||[]).map((x,i)=><div key={i} className="reason" style={{borderLeft:'2px solid rgba(239,68,68,0.4)',paddingLeft:'10px',marginBottom:'6px'}}>
                <span style={{color:'#ef4444',fontWeight:'900',marginRight:'6px'}}>{i+1}.</span>
                <span style={{color:'#e2e8f0',lineHeight:'1.5',fontSize:'12px'}}>{safeRender(x)}</span>
              </div>)}
            </Card>
            <Card><h2>Near-critical density interpretation</h2>
              <p style={{color:'#94a3b8',lineHeight:'1.6',fontSize:'13px',marginBottom:'12px'}}>{safeRender(model.near_critical_narrative)}</p>
              {model?.sector_constraints && <div style={{padding:'10px 14px',background:'rgba(141,247,255,0.05)',borderRadius:'4px',border:'1px solid rgba(141,247,255,0.12)'}}>
                <div style={{fontSize:'9px',color:'#8df7ff',fontWeight:'800',letterSpacing:'.12em',marginBottom:'6px'}}>GOVERNING SECTOR CONSTRAINTS</div>
                <p style={{color:'#cbd5e1',fontSize:'12px',lineHeight:'1.6',margin:0}}>{safeRender(model.sector_constraints)}</p>
              </div>}
            </Card>
          </section>}

          {/* GATE REVIEW READINESS */}
          {model?.gate_review_readiness && <section className="layout two">
            <Card><h2>Gate review readiness</h2>
              <p style={{fontSize:'10px',color:'#475569',marginBottom:'8px',lineHeight:'1.4'}}>G0=concept · G1=options · G2=business case · G3=investment decision · G4=readiness · G5=closeout. IPA/World Bank gateway framework.</p>
              <div style={{display:'flex',alignItems:'center',gap:'16px',marginBottom:'12px'}}>
                <span style={{fontSize:'32px',fontWeight:'900',color:model.gate_review_readiness.overall_verdict==='READY'?'#10b981':model.gate_review_readiness.overall_verdict==='CONDITIONAL'?'#f59e0b':'#ef4444'}}>{model.gate_review_readiness.overall_verdict}</span>
                <span style={{fontSize:'13px',color:'#94a3b8'}}>{safeRender(model.gate_review_readiness.current_gate_readiness)}</span>
              </div>
              <p style={{fontSize:'12px',color:'#64748b',marginBottom:'10px',lineHeight:'1.5'}}>{safeRender(model.gate_review_readiness.ipa_alignment)}</p>
              <p style={{fontSize:'11px',color:'#f59e0b',fontStyle:'italic',marginBottom:'10px'}}>{safeRender(model.gate_review_readiness.critical_gate_risk)}</p>
              <div style={{borderTop:'1px solid rgba(255,255,255,0.06)',paddingTop:'10px'}}>
                <div style={{fontSize:'10px',color:'#64748b',fontWeight:'800',letterSpacing:'.1em',marginBottom:'6px'}}>NEXT GATE ACTIONS</div>
                {(model.gate_review_readiness.next_gate_actions||[]).map((a,i)=><div key={i} style={{fontSize:'11px',color:'#cbd5e1',padding:'3px 0',borderBottom:'1px solid rgba(255,255,255,0.04)',display:'flex',gap:'6px'}}><span style={{color:'#8df7ff',flexShrink:0}}>→</span>{safeRender(a)}</div>)}
              </div>
            </Card>
            {model?.optimism_bias_assessment && <Card><h2>📊 Optimism bias — OBA-adjusted outturn</h2>
              <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'12px',marginBottom:'12px'}}>
                <div style={{background:'rgba(245,158,11,0.08)',borderRadius:'4px',padding:'10px',border:'1px solid rgba(245,158,11,0.2)'}}>
                  <div style={{fontSize:'9px',color:'#f59e0b',fontWeight:'800',letterSpacing:'.1em',marginBottom:'4px'}}>OBA-ADJUSTED P50</div>
                  <div style={{fontSize:'18px',fontWeight:'900',color:'#fcd34d'}}>{safeRender(model.optimism_bias_assessment.oba_adjusted_p50)}</div>
                  <div style={{fontSize:'9px',color:'#64748b'}}>{safeRender(model.optimism_bias_assessment.oba_adjusted_schedule)}</div>
                </div>
                <div style={{background:'rgba(141,247,255,0.05)',borderRadius:'4px',padding:'10px',border:'1px solid rgba(141,247,255,0.1)'}}>
                  <div style={{fontSize:'9px',color:'#8df7ff',fontWeight:'800',letterSpacing:'.1em',marginBottom:'4px'}}>HEADLINE P50</div>
                  <div style={{fontSize:'18px',fontWeight:'900',color:'#e2e8f0'}}>{safeRender(model.cost_p50)}</div>
                  <div style={{fontSize:'9px',color:'#64748b'}}>{safeRender(model.schedule)}</div>
                </div>
              </div>
              <p style={{fontSize:'11px',color:'#94a3b8',lineHeight:'1.5',marginBottom:'6px'}}>{safeRender(model.optimism_bias_assessment.verdict)}</p>
              <p style={{fontSize:'10px',color:'#475569',fontStyle:'italic',marginBottom:'8px'}}>{safeRender(model.optimism_bias_assessment.oba_source)}</p>
              <p style={{fontSize:'11px',color:'#f59e0b',padding:'8px 10px',background:'rgba(245,158,11,0.06)',borderRadius:'3px',lineHeight:'1.5',margin:0}}>{safeRender(model.optimism_bias_assessment.board_challenge)}</p>
            </Card>}
          </section>}

          {/* ORIGINAL ADVISOR PANEL */}
          <section className="layout two advisorElite challengeRoom"><Card><h2>CASEY Board Assurance Console</h2>
<p className="advisorIntro">Ask any question about the live programme. CASEY answers using the actual model data — not a generic response. <b style={{color:'#8df7ff'}}>Try a "what if" question to rerun the model with a constraint applied.</b></p>
<div style={{background:'rgba(141,247,255,0.05)',border:'1px solid rgba(141,247,255,0.12)',borderRadius:'4px',padding:'8px 12px',marginBottom:'10px',fontSize:'11px'}}>
  <div style={{fontSize:'9px',fontWeight:'800',letterSpacing:'.1em',color:'#8df7ff',marginBottom:'5px'}}>WHAT-IF EXAMPLES — type these or use as inspiration</div>
  {[
    'What if contractor A wins the signalling package?',
    'What if we apply a single-source constraint on the main civils package?',
    'What if planning approval is delayed by 18 months?',
    'What if we accelerate delivery by 12 months?',
    'What if funding is capped at the P50 and no P80 reserve is held?',
  ].map(x => <button key={x} onClick={() => ask(x)}
    style={{display:'block',width:'100%',textAlign:'left',background:'rgba(141,247,255,0.06)',border:'1px solid rgba(141,247,255,0.15)',color:'#8df7ff',padding:'5px 10px',borderRadius:'3px',marginBottom:'4px',cursor:'pointer',fontSize:'11px',fontStyle:'italic'}}>
    "{x}"
  </button>)}
</div>
<div className="advisorPrompts bigButtons">{['What is the board not seeing?','What evidence is missing before approval?','What is the real governing constraint?','Which assumptions collapse confidence first?','What are the top risks?','Is this programme gate-ready?','What is the OBA-adjusted outturn?','Give me CASEY POSITION.','What would an external reviewer challenge first?','What is the P80 exposure?','Walk me through the procurement risks.','What does the benchmark data say?'].map(x=><button key={x} data-question={x} onClick={()=>ask(x)}><Brain size={14}/>{x}</button>)}</div><div className="chatBox boardInterrogation">{chat.length ? chat.map((m,i)=><div key={i} className={`msg ${m.role}`}>{(() => {
    const lines = String(m.text||'').split('\n');
    const textBlock = lines.map((line, li) => {
      if (!line.trim()) return <div key={li} style={{height:'5px'}}/>;
      if (line.startsWith('**') && line.endsWith('**') && line.length > 4)
        return <div key={li} className="chatHeading">{line.replace(/\*\*/g,'')}</div>;
      const parts = line.split(/\*\*([^*]+)\*\*/g);
      return <p key={li} className="chatLine">{parts.map((p,pi)=>pi%2===1?<strong key={pi}>{p}</strong>:p)}</p>;
    });
    const d = m.delta;
    const deltaBlock = d && !d.error ? <div style={{marginTop:'10px',background:'rgba(141,247,255,0.05)',border:'1px solid rgba(141,247,255,0.2)',borderRadius:'4px',padding:'10px 12px'}}>
      <div style={{fontSize:'9px',fontWeight:'800',letterSpacing:'.12em',color:'#8df7ff',marginBottom:'6px'}}>◆ MODEL RECALCULATED — CONSTRAINT APPLIED</div>
      <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:'8px',marginBottom:'8px'}}>
        {d.cost_delta_bn !== undefined && <div style={{textAlign:'center',padding:'6px',background:'rgba(255,255,255,0.03)',borderRadius:'3px'}}>
          <div style={{fontSize:'9px',color:'#475569',marginBottom:'2px'}}>Cost delta</div>
          <div style={{fontSize:'13px',fontWeight:'800',color:d.cost_delta_bn>0?'#fca5a5':'#10b981'}}>{d.cost_delta_bn>0?'+':''}{d.cost_delta_bn?.toFixed(1)}B</div>
          <div style={{fontSize:'10px',color:'#8df7ff'}}>{d.new_p50}</div>
        </div>}
        {d.confidence_delta !== undefined && <div style={{textAlign:'center',padding:'6px',background:'rgba(255,255,255,0.03)',borderRadius:'3px'}}>
          <div style={{fontSize:'9px',color:'#475569',marginBottom:'2px'}}>Confidence</div>
          <div style={{fontSize:'13px',fontWeight:'800',color:d.confidence_delta<0?'#fca5a5':'#10b981'}}>{d.confidence_delta>0?'+':''}{d.confidence_delta}pts</div>
          <div style={{fontSize:'10px',color:'#8df7ff'}}>{d.new_confidence}%</div>
        </div>}
        {d.schedule_delta_months !== undefined && <div style={{textAlign:'center',padding:'6px',background:'rgba(255,255,255,0.03)',borderRadius:'3px'}}>
          <div style={{fontSize:'9px',color:'#475569',marginBottom:'2px'}}>Schedule delta</div>
          <div style={{fontSize:'13px',fontWeight:'800',color:d.schedule_delta_months>0?'#fca5a5':'#10b981'}}>{d.schedule_delta_months>0?'+':''}{d.schedule_delta_months}mo</div>
          <div style={{fontSize:'10px',color:'#8df7ff'}}>{d.new_schedule}</div>
        </div>}
      </div>
      {d.new_governing_constraint && <div style={{fontSize:'10px',color:'#94a3b8',borderTop:'1px solid rgba(255,255,255,0.06)',paddingTop:'6px'}}>New governing constraint: <b style={{color:'#e2e8f0'}}>{d.new_governing_constraint}</b></div>}
    </div> : null;
    return <>{textBlock}{deltaBlock}</>;
  })()}</div>) : <div className="msg assistant"><b>Board attack ready.</b><br/>Click any challenge above. CASEY will answer against the active scenario, not as a generic chatbot.</div>}</div><div className="ask"><input value={chatQ} onChange={e=>setChatQ(e.target.value)} onKeyDown={e=>{if(e.key==='Enter')ask()}} placeholder="Ask anything — e.g. What if contractor A wins signalling? What is the P80 exposure? Is this gate-ready?"/><button onClick={() => ask(chatQ)}>Ask</button></div></Card><Card><h2>Live Client Challenge Room</h2><p className="advisorIntro">Upload a client cost estimate, XER schedule or risk register — or use the demo buttons below. CASEY challenges the document like an independent reviewer: identifies unpriced exposure, evidence gaps, reserve weaknesses and the questions to ask before committing capital.</p><div className="challengeHero"><span>CASEY INTAKE NORMALISATION ENGINE</span><b>Messy file → schema detection → WBS/CBS inference → evidence gaps → board attack → export-ready challenge</b></div><div className="challengeModeStrip"><b>Choose the file type to challenge</b><span>These buttons show the exact review style. Upload your own file below and CASEY will replace the sample with parsed source-file numbers.</span></div><div className="challengeButtons pro">
  <button onClick={()=>setUploadResult({filename:'Contractor_Cost_Estimate_v27_FINAL.xlsx', file_type:'COST ESTIMATE', schema_confidence:'Auto-mapped', findings:['Estimate structure normalised. Cost packages identified across direct works, preliminaries and risk allowance.','The headline P50 number is present — but there is no P80/P90 basis. This is how a submitted estimate can understate exposure: the headline looks fixed but the downside is unpriced.','Contingency is present as a lump sum. CASEY cannot verify it is sized against quantified risk exposure rather than a percentage of direct cost — a common tender-basis weakness.','Basis statements are missing on 6 of the major packages. Without basis, there is no evidence of what was included or excluded — and no way to challenge scope creep later.'], red_flags:['Commercial observation: No P80/P90 range provided. The estimate looks precise but carries unquantified downside. Ask the contractor to provide a risk-adjusted range.','Commercial observation: Lump-sum contingency with no risk linkage. This is not yet a quantified reserve. Require QCRA support.','No CBS/WBS mapping. Cannot verify completeness of scope coverage or trace costs to programme activities.','Escalation basis not stated. For a multi-year programme, this is a material omission.'], next_steps:['Require the contractor to provide a P50/P80/P90 range with QCRA support.','Mandate a CBS that maps to the programme WBS and schedule activities.','Commission an independent cost review before approving the headline number.','Run CASEY QCRA alongside the contractor estimate — compare the P80 positions.'], epc_challenge:true})}><FileSpreadsheet size={18}/><b>Challenge contractor cost estimate</b><span>Detect hidden exposure, lump-sum contingency and missing basis statements.</span></button>
  <button onClick={()=>setUploadResult({filename:'Programme_Schedule_FINAL_v14.xer', file_type:'SCHEDULE (XER)', schema_confidence:'Logic mapped', findings:['Schedule logic parsed. Activities identified across civil, systems, commissioning and handover phases.','Critical path identified — but float analysis reveals operationally unusable buffer. The management date assumes best-case access windows throughout.','Logic gaps detected: 8 activities have no predecessor. These are schedule anchors — they cannot be challenged because they have no upstream dependency. This can prevent a reliable view of the real critical path.','Commissioning and trial running phases show compressed durations. These are the activities most likely to slip — and they sit directly before the opening/handover milestone.'], red_flags:['Commercial observation: Open-ended activities with no predecessor — schedule logic issue that can overstate available float.','Commercial observation: Commissioning duration appears optimistic against comparable programmes. A single failed integration test resets the clock.','Float is nominal, not operationally usable. Access windows, possession permits and operator acceptance are not confirmed in the logic.','Board date is driven by the earliest path. It should be driven by the P80/P90 QSRA finish date.'], next_steps:['Require the contractor to close all open ends and confirm predecessor logic.','Run QSRA and require the P80/P90 finish date to be the board commitment date.','Validate all commissioning durations against independent benchmarks.','Name the owner of the critical-path constraint.'], epc_challenge:true})}><Workflow size={18}/><b>Challenge programme schedule</b><span>Detect schedule padding, unusable float and optimistic commissioning dates.</span></button>
  <button onClick={()=>setUploadResult({filename:'Risk_Register_v8_Draft.xlsx', file_type:'RISK REGISTER', schema_confidence:'Schema mapped', findings:['Risk register schema mapped. Cause, event, impact and owner columns identified.','CASEY challenges every risk without a named trigger, quantified residual exposure and evidence closure date.','7 risks have mitigation confidence below 50%. These are not mitigated — they are noted. The reserve needs to account for them.','4 risks are flagged as Evidence required. These are open exposures — the source file does not yet provide the evidence that the risk is under control.'], red_flags:['Commercial observation: Mitigations are written as action phrases ("to be confirmed", "in progress") rather than evidence closure. A mitigation is only valid when the evidence is complete.','Commercial observation: Residual exposure is not reconciled to the reserve. This is the most common way a risk register hides real exposure — risks exist on paper but the money is not in the budget.','4 risks require evidence that has not been provided. These cannot be treated as mitigated for board approval purposes.','Owner accountability: all risks assigned to programme-level owners. Board needs named individual owners with accountability.'], next_steps:['Require every risk to have: named owner, confirmed trigger, quantified residual and evidence closure date.','Reconcile residual exposure to reserve — any gap requires additional provision.','The 4 Evidence Required risks must be resolved or escalated to the board as open items.','Export the challenged register after QCRA/QSRA alignment and use the board attack questions.'], epc_challenge:true})}><ShieldAlert size={18}/><b>Challenge risk register</b><span>Detect unmitigated risks, missing evidence and reserve reconciliation gaps.</span></button>
</div><h3>Upload real file</h3><label className="upload proUpload"><Upload size={18}/> Upload estimate / XER / risk workbook<input type="file" onChange={upload}/></label><ProfessionalIntakeResult result={uploadResult} model={model}/></Card></section></>
}

        {tab === 'runtime' && <HolyGrailRuntime model={model} scenario={scenario} generate={generate} runShock={runShock}/>}
        {tab === 'method' && <section className="layout two"><Card><h2>How CASEY calculated this</h2>{['Cost model: selected class estimate, sector template, location factor, complexity factor and scenario modifier.','Schedule model: level-based delivery logic, phase durations, critical path sensitivity and scenario acceleration/delay factors.','QCRA: cost exposure model using low / most likely / high impacts and risk-weighted contingency.','QSRA: schedule exposure model using activity-linked O/M/P delay ranges and critical path sensitivity.','Confidence score translated for executives: board-defensibility based on benchmark similarity, evidence maturity, procurement certainty, schedule logic, contingency adequacy and scenario posture.'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card><Card><h2>Commercial readiness</h2><p style={{fontSize:'12px',color:'#64748b'}}>First-pass intelligence for challenge, option testing and board preparation — before contractor tender or signed cost plan.</p><a className="contactBtn huge" href={emailLink}><Mail/> Send project for review</a></Card></section>}
        {tab === 'benchmark' && <section className="layout two">
          <Card><h2>Named global benchmarks</h2>
            <p style={{fontSize:'11px',color:'#64748b',marginBottom:'8px'}}>These are real programmes from public record — OECD, parliamentary accounts committees, company filings, academic literature (Flyvbjerg et al). Cost growth % and schedule slip are actuals, not estimates. CASEY routes every project through the closest matching comparables and applies their delivery behaviour to confidence, reserve and P80/P90 exposure.</p>
            <div style={{overflowX:'auto'}}>
              <table style={{width:'100%',borderCollapse:'collapse',fontSize:'11px'}}>
                <thead><tr style={{borderBottom:'1px solid rgba(255,255,255,0.1)'}}>
                  {['Programme','P50 Anchor','Cost Growth','Slip (mo)','Primary Failure Mode'].map(h=><th key={h} style={{padding:'6px 8px',textAlign:'left',color:'#64748b',fontWeight:'800',letterSpacing:'.08em'}}>{h}</th>)}
                </tr></thead>
                <tbody>{(model?.benchmark_comparison||[]).map((b,i)=><tr key={i} style={{borderBottom:'1px solid rgba(255,255,255,0.04)',background:i%2===0?'rgba(255,255,255,0.01)':'transparent'}}>
                  <td style={{padding:'7px 8px',color:'#e2e8f0',fontWeight:'700'}}>{b.name||b.archetype}</td>
                  <td style={{padding:'7px 8px',color:'#8df7ff',fontSize:'10px'}}>{b.sector}</td>
                  <td style={{padding:'7px 8px',color:'#94a3b8'}}>{b.anchor_cost||b.value}</td>
                  <td style={{padding:'7px 8px',color:'#94a3b8'}}>{b.anchor_duration_months ? b.anchor_duration_months + ' mo' : '—'}</td>
                  <td style={{padding:'7px 8px',color:(b.cost_growth_pct??b.growth)>50?'#ef4444':(b.cost_growth_pct??b.growth)>20?'#f59e0b':'#10b981',fontWeight:'700'}}>{(b.cost_growth_pct??b.growth) ? '+'+(b.cost_growth_pct??b.growth)+'%' : '—'}</td>
                  <td style={{padding:'7px 8px',color:(b.schedule_slip_months??b.slip)>24?'#ef4444':(b.schedule_slip_months??b.slip)>12?'#f59e0b':'#94a3b8',fontWeight:'700'}}>{(b.schedule_slip_months??b.slip) ? '+'+(b.schedule_slip_months??b.slip) : '—'}</td>
                  <td style={{padding:'7px 8px',color:'#64748b',maxWidth:'220px',lineHeight:'1.4'}}>{b.failure_mode||'—'}</td>
                </tr>)}
                </tbody>
              </table>
            </div>
          </Card>
          <Card><h2>What benchmarks mean for your programme</h2>
            <p className="chartCaption">CASEY routes every project through the closest real-world comparables. The benchmark similarity score, cost growth history and failure mode are applied to P80/P90 exposure and OBA adjustment.</p>
            {(model?.benchmark_comparison||[]).map((b,i)=><div key={i} className="reason" style={{borderLeft:`2px solid ${(b.cost_growth_pct??b.growth)>80?'rgba(239,68,68,0.5)':(b.cost_growth_pct??b.growth)>30?'rgba(245,158,11,0.5)':'rgba(141,247,255,0.3)'}`,paddingLeft:'10px',marginBottom:'8px'}}>
              <b style={{color:'#e2e8f0'}}>{b.name||b.archetype}</b>
              {(b.cost_growth_pct??b.growth) > 0 && <span style={{background:'rgba(239,68,68,0.1)',color:'#fca5a5',borderRadius:'3px',padding:'1px 6px',fontSize:'10px',fontWeight:'800',marginLeft:'8px'}}>+{(b.cost_growth_pct??b.growth)}% cost</span>}
              {(b.schedule_slip_months??b.slip) > 0 && <span style={{background:'rgba(245,158,11,0.1)',color:'#fde68a',borderRadius:'3px',padding:'1px 6px',fontSize:'10px',fontWeight:'800',marginLeft:'4px'}}>+{(b.schedule_slip_months??b.slip)}mo</span>}
              <p style={{fontSize:'11px',color:'#64748b',marginTop:'4px',lineHeight:'1.5'}}>{b.lesson||b.failure_mode||b.why}</p>
            </div>)}
          </Card>
        </section>}

        {tab === 'pricing' && <section className="layout two">
          <Card><h2>Get in touch</h2>
            <p style={{fontSize:'13px',color:'#94a3b8',lineHeight:'1.7',marginBottom:'16px'}}>CASEY is not a subscription product you buy online. It is a capital programme intelligence platform used by programme sponsors, investment committees and infrastructure investors who need defensible cost, schedule and confidence intelligence before committing capital.</p>
            <div style={{padding:'14px',background:'rgba(141,247,255,0.06)',border:'1px solid rgba(141,247,255,0.2)',borderRadius:'6px',marginBottom:'14px'}}>
              <div style={{fontSize:'10px',fontWeight:'800',color:'#8df7ff',marginBottom:'8px',letterSpacing:'.1em'}}>WHAT A CONVERSATION COVERS</div>
              {['Your programme — sector, country, scale, current estimate stage and board timeline.',
                'How CASEY can be used — standalone intelligence pack, advisory engagement support, or enterprise API.',
                'What you get that T&T, Atkins, Jacobs and Accenture cannot produce in the same timeframe.',
                'Benchmark validation — we will show you the named comparables calibrating your programme.',
                'Next step — a 30-minute walkthrough of your specific programme with live output.'
              ].map((x,i)=><div className="reason" key={i}><span style={{background:'rgba(141,247,255,0.1)',color:'#8df7ff'}}>{i+1}</span>{x}</div>)}
            </div>
            <a className="contactBtn huge" href={emailLink} style={{display:'block',textAlign:'center',padding:'14px',marginBottom:'10px',fontSize:'14px',fontWeight:'800'}}>
              <Mail size={16}/> deepa@caseai.co.uk
            </a>
            <a href="https://www.linkedin.com/company/caseai" target="_blank" rel="noopener noreferrer" style={{display:'block',textAlign:'center',padding:'10px',background:'rgba(255,255,255,0.04)',border:'1px solid rgba(255,255,255,0.1)',borderRadius:'5px',color:'#64748b',textDecoration:'none',fontSize:'12px',marginBottom:'10px'}}>
              LinkedIn — ControlOrbit / CASEY
            </a>
          </Card>
          <Card><h2>What CASEY produces that consultants cannot</h2>
            {[
              ['12 seconds vs 8 weeks','Full intelligence pack generated in seconds. A preliminary advisory note from Jacobs or T&T takes 6-8 weeks and costs £50K-£150K.'],
              ['63 named real programmes','Every estimate is calibrated against real completed programmes — Crossrail +88%, Hinkley C +94%, JWST +506%, Britishvolt collapse. Named, cited, public record.'],
              ['Programme Mortality Engine','Probability of programme cancellation with named historical precedents. No consultant produces this because a high score reduces their fee scope.'],
              ['OBA from reference class','Optimism bias assessed against Flyvbjerg 2022 and IPA Annual Report sector reference classes — not generic percentage uplift.'],
              ['42 country intelligence profiles','Regulatory framework, approval body, financing structure and OBA note for 42 countries. No additional fee.'],
              ['Five scenario trade-offs instantly','Faster, Cheaper, Lower Risk, Premium and Base — each a complete independent recalculation. A consultant charges day rates for each.'],
              ['Evidence Gap Scanner','Named evidence gaps that will block approval at any gate review, PAC inquiry or spending review. IPA and Green Book referenced.'],
              ['Board attack simulation','The 5-6 questions your investment committee will ask, generated from the live model data — not generic board challenge frameworks.'],
            ].map(([title, body])=><div key={title} style={{marginBottom:'8px',paddingBottom:'8px',borderBottom:'1px solid rgba(255,255,255,0.04)'}}>
              <div style={{fontSize:'11px',fontWeight:'800',color:'#e2e8f0',marginBottom:'2px'}}>{title}</div>
              <div style={{fontSize:'10px',color:'#64748b',lineHeight:'1.5'}}>{body}</div>
            </div>)}
          </Card>
        </section>}
      </>}
    </main>
  {(loading || exportingLabel) && <div className="simOverlay">
      <div className="simCard">
        <div className="simSpinner" />
        <h3 style={{textAlign:'center',lineHeight:'1.4',maxWidth:'360px'}}>{exportingLabel || simulationStage || 'Building intelligence pack…'}</h3>
        {(simulationStage||'').toLowerCase().includes('demo') || (simulationStage||'').toLowerCase().includes('reference')
          ? <p style={{textAlign:'center',color:'#64748b',fontSize:'12px',maxWidth:'320px',lineHeight:'1.5'}}>First load takes 20–30 seconds while the server starts. Subsequent loads are instant. Please wait.</p>
          : <p style={{textAlign:'center',color:'#64748b',fontSize:'11px'}}>QCRA · QSRA · Risk register · Benchmarks · Board pack</p>}
      </div>
    </div>}
  </div>;
}

// ── SAVED PROJECTS + INVESTOR CSS ────────────────────────────────────────────
// ── COMPACT DENSITY OVERRIDE CSS ─────────────────────────────────────────────
const compactDensityCSS = `
  /* Tighter headings throughout */
  .v50EliteApp .card h2 { font-size: 12px !important; font-weight: 800; letter-spacing: .06em; text-transform: uppercase; margin-bottom: 8px !important; }
  .v50EliteApp .card h3 { font-size: 11px !important; font-weight: 700; margin: 10px 0 6px; }
  .v50EliteApp .reason { padding: 5px 0 !important; font-size: 12px !important; line-height: 1.5 !important; }
  .v50EliteApp .card { padding: 14px 16px !important; }
  .v50EliteApp .big { font-size: 13px !important; line-height: 1.6 !important; }
  .v50EliteApp .chartCaption { font-size: 11px !important; color: #64748b; margin-bottom: 8px !important; }
  /* KPI strip - more compact */
  .v50EliteApp .kpis { gap: 6px !important; padding: 8px 0 !important; }
  /* Tabs - tighter */
  .v50EliteApp .tabs button { font-size: 10px !important; padding: 5px 10px !important; }
  /* Scenario rail - compact */
  .v50EliteApp .scenarioRail button { padding: 7px 10px !important; }
  .v50EliteApp .scenarioRail button b { font-size: 11px !important; }
  .v50EliteApp .scenarioRail button span { font-size: 10px !important; }
  /* Overview section gaps */
  .v50EliteApp .layout { gap: 10px !important; margin-bottom: 10px !important; }
  /* IntelligenceMeta rail */
  .v50EliteApp .orbitalMetaRail { padding: 6px 14px !important; gap: 16px !important; font-size: 10px !important; }
  /* Confidence badge */
  .v50EliteApp .confidenceEngineBadge { padding: 5px 14px !important; font-size: 10px !important; }
  /* Mission cards */
  .v50EliteApp .intelCard { padding: 8px 10px !important; }
  .v50EliteApp .intelCard b { font-size: 11px !important; }
  .v50EliteApp .intelCard p { font-size: 10px !important; }
  /* fee counter */
  .v50EliteApp .feeRow { padding: 4px 0 !important; font-size: 11px !important; }
  /* Audit spine */
  .v50EliteApp .auditSpine { padding: 4px 0 !important; font-size: 11px !important; }
  /* Advisor prompts */
  .v50EliteApp .advisorPrompts button { font-size: 11px !important; padding: 6px 10px !important; }
  /* Export strip */
  .v50EliteApp .v50ExportStrip button { font-size: 10px !important; padding: 5px 10px !important; }
`;
if (!document.querySelector('#casey-compact-css')) {
  const s = document.createElement('style');
  s.id = 'casey-compact-css';
  s.textContent = compactDensityCSS;
  document.head.appendChild(s);
}

const savedInvestorCSS = `
.savedPanel,.investorPanel{position:fixed;top:0;right:0;width:min(640px,100vw);height:100vh;background:#0c1a2e;border-left:1px solid rgba(141,247,255,0.15);z-index:900;overflow-y:auto;display:flex;flex-direction:column;}
.savedHeader,.investorHeader{padding:18px 20px;border-bottom:1px solid rgba(255,255,255,0.08);display:flex;justify-content:space-between;align-items:flex-start;flex-shrink:0;background:#0a1628;}
.savedHeader h2,.investorHeader h2{color:#e2e8f0;font-size:16px;font-weight:800;margin:0;}
.savedHeader button,.investorHeader button{background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.1);color:#94a3b8;padding:5px 12px;border-radius:3px;cursor:pointer;font-size:12px;flex-shrink:0;}
.savedGrid{padding:16px;display:flex;flex-direction:column;gap:10px;}
.savedCard{background:rgba(255,255,255,0.03);border:1px solid rgba(255,255,255,0.08);border-radius:5px;padding:14px 16px;}
.savedCard:hover{border-color:rgba(141,247,255,0.2);}
.savedCard h3{color:#e2e8f0;font-size:13px;font-weight:700;margin:4px 0 10px;}
.savedMeta{display:flex;justify-content:space-between;font-size:10px;font-weight:800;letter-spacing:.1em;color:#8df7ff;margin-bottom:4px;}
.savedMeta em{color:#475569;font-style:normal;}
.savedStats{display:grid;grid-template-columns:repeat(4,1fr);gap:6px;margin-bottom:12px;}
.savedStats>div{background:rgba(255,255,255,0.03);border-radius:3px;padding:6px;text-align:center;}
.savedStats span{display:block;font-size:9px;color:#475569;font-weight:800;letter-spacing:.08em;margin-bottom:2px;}
.savedStats b{display:block;font-size:12px;color:#8df7ff;font-weight:700;}
.savedActions{display:flex;gap:8px;}
.savedLoad{background:rgba(141,247,255,0.08);border:1px solid rgba(141,247,255,0.2);color:#8df7ff;padding:5px 14px;border-radius:3px;cursor:pointer;font-size:11px;font-weight:700;}
.savedLoad:hover{background:rgba(141,247,255,0.15);}
.savedDelete{background:transparent;border:1px solid rgba(239,68,68,0.2);color:#475569;padding:5px 12px;border-radius:3px;cursor:pointer;font-size:11px;}
.savedDelete:hover{border-color:rgba(239,68,68,0.4);color:#fca5a5;}
.investorPanel{width:min(900px,100vw);}
.investorHeader{flex-direction:row;gap:16px;align-items:flex-start;}
.investorHeader h2{font-size:20px;color:#e2e8f0;margin:4px 0 6px;}
.investorHeader p{color:#64748b;font-size:12px;margin:0;}
.investorGrid{padding:20px;display:flex;flex-direction:column;gap:16px;}
.investorBlock{background:rgba(255,255,255,0.02);border:1px solid rgba(255,255,255,0.07);border-radius:5px;padding:16px 18px;}
.investorBlock h3{font-size:11px;font-weight:800;letter-spacing:.12em;color:#8df7ff;margin:0 0 12px;text-transform:uppercase;}
.investorBlock.wide{width:100%;}
.investorMetric{display:inline-block;margin:0 16px 12px 0;vertical-align:top;width:calc(50% - 20px);}
.investorMetricVal{font-size:22px;font-weight:900;color:#e2e8f0;line-height:1;}
.investorMetricLabel{font-size:11px;color:#94a3b8;font-weight:600;margin:3px 0 1px;}
.investorMetricNote{font-size:10px;color:#475569;line-height:1.4;}
.investorMoat{margin-bottom:12px;padding-bottom:12px;border-bottom:1px solid rgba(255,255,255,0.05);}
.investorMoat:last-child{border-bottom:none;margin-bottom:0;}
.investorMoat b{color:#e2e8f0;font-size:12px;display:block;margin-bottom:3px;}
.investorMoat p{color:#64748b;font-size:11px;line-height:1.6;margin:0;}
.investorRevTable{display:flex;flex-direction:column;gap:0;}
.investorRevHeader,.investorRevRow{display:grid;grid-template-columns:1.8fr 1.8fr 1.6fr 2fr;gap:8px;padding:7px 0;border-bottom:1px solid rgba(255,255,255,0.05);font-size:11px;}
.investorRevHeader{color:#64748b;font-weight:800;letter-spacing:.08em;font-size:9px;text-transform:uppercase;}
.investorRevRow{color:#94a3b8;}
.investorRevRow:last-child{border-bottom:none;}
`;
if (!document.querySelector('#casey-saved-investor-css')) {
  const s = document.createElement('style');
  s.id = 'casey-saved-investor-css';
  s.textContent = savedInvestorCSS;
  document.head.appendChild(s);
}

createRoot(document.getElementById('root')).render(<CaseyErrorBoundary><App/></CaseyErrorBoundary>);
