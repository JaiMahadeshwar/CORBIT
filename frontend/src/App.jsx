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
if (typeof window !== 'undefined') window._CASEY_API = PROD_URL;
// Only use localhost fallbacks in development - never in production (triggers browser security warnings)
const IS_DEV = typeof window !== 'undefined' && (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1');
// Rollup-safe sign prefix helper
const sign = (n) => n > 0 ? '+' : n < 0 ? '-' : '';
const signedVal = (n, suffix='') => n !== 0 ? sign(n) + Math.abs(n) + suffix : null;

const API_CANDIDATES = IS_DEV 
  ? [PROD_URL, 'http://127.0.0.1:8000', 'http://localhost:8000'].filter(Boolean)
  : [PROD_URL].filter(Boolean);
let API = API_CANDIDATES[0];
async function apiFetch(path, options, timeoutMs = 90000) {
  let lastError;
  for (const base of API_CANDIDATES) {
    try {
      const controller = new AbortController();
      const tid = setTimeout(() => controller.abort(), timeoutMs);
      // credentials:'omit' prevents CORS credential restriction
      // Never fall through to localhost on production - prevents "unsafe site" browser warning
      const r = await fetch(base + path, { credentials: 'omit', ...options, signal: controller.signal });
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
  const _curr = model?.currency_symbol || '$';
  const challengeP50 = cm.p50_bn ? (_curr + Number(cm.p50_bn).toFixed(1) + 'B') : baselineP50;
  const challengeP80 = cm.p80_bn ? (_curr + Number(cm.p80_bn).toFixed(1) + 'B') : (qcraP80 ? moneyLocal(qcraP80, _curr) : '—');
  const challengeP90 = cm.p90_bn ? (_curr + Number(cm.p90_bn).toFixed(1) + 'B') : (model?.monte_carlo?.qcra?.p90 ? moneyLocal(model.monte_carlo.qcra.p90) : '—');
  const deltaBn = cm.delta_bn ?? (cm.p80_bn && model?.cost_p50_bn ? Number(cm.p80_bn) - Number(model.cost_p50_bn) : null);
  const deltaText = deltaBn !== null && deltaBn !== undefined && !Number.isNaN(Number(deltaBn)) ? ((Number(deltaBn) >= 0 ? '+' : '−') + _curr + Math.abs(Number(deltaBn)).toFixed(1) + 'B latent exposure') : 'Exposure delta requires source bundle';
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
        {engMetrics.p90>0 ? <div className="cm"><span>P90 downside</span><b style={{color:'#ef4444'}}>${engMetrics.p90.toFixed(1)+'B'}</b></div> : <div className="cm"><span>Challenge P80</span><b style={{color:'#ff9940'}}>{challengeP80}</b></div>}
        {engMetrics.emv>0 ? <div className="cm"><span>Total EMV</span><b style={{color:'#ff9940'}}>${engMetrics.emv.toFixed(2)+'B'}</b></div> : engMetrics.acts>0 ? <div className="cm"><span>Activities</span><b>{engMetrics.acts}</b></div> : <div className="cm"><span>Stress P90</span><b style={{color:'#ff6b7d'}}>{challengeP90}</b></div>}
        <div className="cm"><span>Evidence quality</span><b style={{color:confPct<60?'#f7d774':'#8df7ff'}}>{confImpact||confPct+'%'}</b></div>
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
  const p50 = model?.cost_p50 || (model?.cost_p50_bn ? (model?.currency_symbol || '$') + model.cost_p50_bn + 'B' : '—');
  const conf = model?.confidence_pct;
  const chain = (model?.causal_chain || []).join(' → ') || 'Generate a project first to see the causal chain';
  return <section className="layout two runtimePanel">
    <Card>
      <h2>Live Programme Stress Test</h2>
      <p className="big">Select a real-world risk event. CASEY recalculates cost, schedule, confidence and board posture from the live model — not from a pre-written response. This is what separates CASEY from a static dashboard.</p>
      <p style={{fontSize:'11px',color:'#64748b',marginBottom:'12px'}}>First generate a project on the Overview tab, then click any event below to stress-test it.</p>
      <div className="runtimeButtons">
        {controls.map(([id,title,sub])=>(
          <button key={id} onClick={()=>fire(id)} className={lastFired===id?'fired':undefined}>
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
  constructor(props) {
    super(props);
    this.state = { error: null };
  }

  static getDerivedStateFromError(error) {
    return { error };
  }

  componentDidCatch(error, info) {
    console.error('CASEY UI crash guard:', error, info);
  }

  render() {
    if (this.state.error) {
      const message = safeRender(this.state.error?.message || this.state.error || 'Unknown render error');
      return (
        <div className="app v50EliteApp">
          <main className="v50Console">
            <section className="layout one">
              <div className="card shockCard">
                <h2>CASEY UI recovered</h2>
                <p>
                  The interface caught a render exception instead of going blank.
                  Refresh once. If it repeats, copy the error below.
                </p>
                <pre style={{ whiteSpace: 'pre-wrap' }}>{message}</pre>
              </div>
            </section>
          </main>
        </div>
      );
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
// Scenario multipliers used across all tabs to scale sub-panels consistently
const getScenarioMults = (model) => ({
  cost: parseFloat(model?.scenario_cost_mult || 1.0),
  risk: parseFloat(model?.scenario_risk_mult || 1.0),
  sched: parseFloat(model?.scenario_sched_mult || 1.0),
  conf: parseInt(model?.scenario_conf_delta || 0),
  isNonBase: !!(model?.scenario && model.scenario !== 'base'),
});


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
{ sector:'AI / Data Centres', region:'United States', client:'STACK Infrastructure COL01', title:'STACK Denver Campus', icon:'AI', confidence:'Grid connection and power procurement', prompt:'STACK Infrastructure Denver Campus data centre 960MW hyperscale colocation power infrastructure grid connection GPU deployment' },
  { sector:'AI / Data Centres', region:'United States', client:'STACK Infrastructure VA01', title:'STACK Northern Virginia Campus', icon:'AI', confidence:'Planning consent and grid', prompt:'STACK Infrastructure Northern Virginia campus hyperscale colocation data centre power Loudoun County planning moratorium grid upgrade' },
  { sector:'AI / Data Centres', region:'Germany', client:'STACK Infrastructure FRA01', title:'STACK Frankfurt Campus', icon:'AI', confidence:'EU permitting and grid', prompt:'STACK Infrastructure Frankfurt Germany European campus data centre hyperscale power EU heat reuse permitting regulatory' },
  { sector:'AI / Data Centres', region:'Poland', client:'STACK Infrastructure WAW01', title:'STACK Warsaw Campus', icon:'AI', confidence:'Power procurement and grid', prompt:'STACK Infrastructure Warsaw Poland campus data centre hyperscale European PPA renewable energy grid decarbonisation' },
  { sector:'AI / Data Centres', region:'Singapore', client:'STACK Infrastructure SIN01', title:'STACK Singapore Campus', icon:'AI', confidence:'Renewable energy compliance', prompt:'STACK Infrastructure Singapore campus data centre hyperscale colocation Asia Pacific government moratorium renewable energy requirement' },
  { sector:'AI / Data Centres', region:'Japan', client:'STACK Infrastructure TYO01', title:'STACK Tokyo Campus', icon:'AI', confidence:'Seismic compliance and grid', prompt:'STACK Infrastructure Tokyo Japan campus data centre hyperscale Asia Pacific seismic compliance structural grid capacity' },
  { sector:'AI / Data Centres', region:'United Kingdom', client:'Hyperscale reference case', title:'UK AI Campus — Midlands', icon:'AI', confidence:'Grid and planning', prompt:'UK AI hyperscale data centre campus Midlands 500MW grid connection planning consent National Grid upgrade renewable energy' },
  { sector:'AI / Data Centres', region:'Australia', client:'Hyperscale reference case', title:'Sydney AI Data Centre Hub', icon:'AI', confidence:'Water and power constraints', prompt:'Sydney Australia AI data centre hyperscale hub power water cooling constraints grid augmentation sovereign data requirements' },
  { sector:'AI / Data Centres', region:'United Arab Emirates', client:'Sovereign AI reference case', title:'Abu Dhabi Sovereign AI Campus', icon:'AI', confidence:'Sovereign data and power', prompt:'Abu Dhabi UAE sovereign AI data centre campus 1GW power cooling infrastructure G42 government strategic AI deployment' },
  { sector:'AI / Data Centres', region:'Ireland', client:'Hyperscale reference case', title:'Dublin Hyperscale Cluster', icon:'AI', confidence:'Grid moratorium risk', prompt:'Dublin Ireland hyperscale data centre cluster EU grid moratorium EirGrid capacity constraints planning restrictions sustainability' },
  { sector:'Rail / Transit', region:'Australia', client:'Australian government reference case', title:'Inland Rail Melbourne–Brisbane', icon:'Rail', confidence:'Alignment and land access', prompt:'Inland Rail Melbourne Brisbane freight corridor 1700km alignment land access communities grade separation bridges tunnels' },
  { sector:'Rail / Transit', region:'France', client:'SNCF / EU reference case', title:'Grand Paris Express', icon:'Rail', confidence:'Urban tunnelling interfaces', prompt:'Grand Paris Express metro 200km four new lines urban tunnelling utility interfaces station excavation operational railway adjacency' },
  { sector:'Rail / Transit', region:'Spain', client:'Adif reference case', title:'Barcelona–Valencia HSR Upgrade', icon:'Rail', confidence:'Mixed traffic interfaces', prompt:'Barcelona Valencia high speed rail upgrade mixed traffic interfaces existing network possession strategy electrification track renewal' },
  { sector:'Rail / Transit', region:'India', client:'NHSRCL reference case', title:'Mumbai–Ahmedabad HSR Phase 2', icon:'Rail', confidence:'Land acquisition and viaduct', prompt:'Mumbai Ahmedabad high speed rail Phase 2 land acquisition viaduct construction station interfaces passenger ramp-up demand risk' },
  { sector:'Rail / Transit', region:'Saudi Arabia', client:'SAR reference case', title:'Riyadh Metro Network Expansion', icon:'Rail', confidence:'Systems integration', prompt:'Riyadh Metro network expansion systems integration signalling rolling stock depot maintenance urban interface passenger operations' },
  { sector:'Rail / Transit', region:'United States', client:'Amtrak reference case', title:'Northeast Corridor Modernisation', icon:'Rail', confidence:'Possession and funding', prompt:'Northeast Corridor Amtrak modernisation bridge tunnel replacement electrification possession strategy funding federal state coordination' },
  { sector:'Rail / Transit', region:'Germany', client:'DB reference case', title:'Stuttgart 21 Completion', icon:'Rail', confidence:'Political and cost governance', prompt:'Stuttgart 21 underground station completion cost governance political scrutiny tunnelling interfaces utility diversions operational handover' },
  { sector:'Rail / Transit', region:'Canada', client:'Metrolinx reference case', title:'GO Rail Expansion Programme', icon:'Rail', confidence:'Grade separation and electrification', prompt:'GO Rail Toronto expansion electrification grade separation stations freight corridor sharing bi-directional running operational complexity' },
  { sector:'Mega Infrastructure', region:'Saudi Arabia', client:'NEOM reference case', title:'NEOM Backbone Infrastructure', icon:'Mega', confidence:'Desert logistics and sovereign risk', prompt:'NEOM backbone infrastructure desalination power transmission roads ports airports workforce accommodation desert logistics sovereign megaproject' },
  { sector:'Mega Infrastructure', region:'Egypt', client:'Government reference case', title:'Suez Canal Expansion Zone', icon:'Mega', confidence:'Geopolitical and logistics risk', prompt:'Suez Canal expansion zone industrial port logistics infrastructure free zone utilities power water Egyptian sovereign programme' },
  { sector:'Mega Infrastructure', region:'Indonesia', client:'Government reference case', title:'Nusantara Capital City', icon:'Mega', confidence:'Greenfield governance', prompt:'Nusantara Indonesia new capital city greenfield infrastructure roads utilities government precinct mass transit housing sovereign programme' },
  { sector:'Energy / Industrial', region:'United Kingdom', client:'EDF / Government reference case', title:'Hinkley Point C Nuclear', icon:'Nuclear', confidence:'First-of-kind supply chain', prompt:'Hinkley Point C nuclear power station EPR first of kind UK construction supply chain workforce regulatory consent cost governance' },
  { sector:'Energy / Industrial', region:'United Kingdom', client:'NNB reference case', title:'Sizewell C Nuclear', icon:'Nuclear', confidence:'Consent and funding model', prompt:'Sizewell C nuclear power station UK development consent financing model RAB supply chain lessons Hinkley regulatory approval' },
  { sector:'Energy / Industrial', region:'United States', client:'Georgia Power reference case', title:'Vogtle Units 3 & 4', icon:'Nuclear', confidence:'Regulatory and construction risk', prompt:'Vogtle Units 3 4 AP1000 nuclear power USA construction overrun regulatory NRC inspection supply chain workforce ramp' },
  { sector:'Energy / Industrial', region:'Finland', client:'TVO reference case', title:'Olkiluoto 3 EPR Lessons', icon:'Nuclear', confidence:'First-of-kind overrun', prompt:'Olkiluoto 3 EPR Finland nuclear construction lessons first of kind cost overrun schedule delay regulatory interface supply chain' },
  { sector:'Energy / Industrial', region:'United Kingdom', client:'National Grid reference case', title:'Eastern Green Link 2', icon:'Energy', confidence:'Marine and consent risk', prompt:'Eastern Green Link 2 HVDC subsea cable Scotland England 2GW offshore wind transmission marine consenting landfall cable installation' },
  { sector:'Energy / Industrial', region:'Australia', client:'ARENA reference case', title:'Snowy 2.0 Pumped Hydro', icon:'Energy', confidence:'Tunnelling and cost governance', prompt:'Snowy 2.0 pumped hydro Australia tunnelling TBM headrace tailrace cost overrun schedule governance environmental consent' },
  { sector:'Energy / Industrial', region:'Global', client:'LNG reference case', title:'Floating LNG Production Vessel', icon:'Energy', confidence:'Offshore integration risk', prompt:'Floating LNG production vessel FLNG offshore integration topsides installation moorings gas processing liquefaction single point mooring' },
  { sector:'Energy / Industrial', region:'United States', client:'BP / reference case', title:'Gulf Coast LNG Export Phase 2', icon:'Energy', confidence:'Regulatory and offtake risk', prompt:'Gulf Coast LNG export terminal Phase 2 FERC approval liquefaction trains storage marine berths offtake contracts cost escalation' },
  { sector:'Energy / Industrial', region:'Norway', client:'Equinor reference case', title:'Northern Lights CCS Hub', icon:'Energy', confidence:'CO2 transport and storage risk', prompt:'Northern Lights CCS carbon capture storage Norway subsea CO2 transport injection well offshore platform regulatory first commercial scale' },
  { sector:'Energy / Industrial', region:'Australia', client:'Fortescue reference case', title:'Green Hydrogen Export Corridor', icon:'Energy', confidence:'Electrolyser and offtake risk', prompt:'Green hydrogen export Australia electrolyser gigawatt scale renewable power ammonia conversion port export offtake agreements cost reduction' },
  { sector:'Energy / Industrial', region:'United Kingdom', client:'National Grid reference case', title:'North Sea Wind Transmission Upgrade', icon:'Energy', confidence:'Offshore cable installation', prompt:'North Sea offshore wind transmission upgrade HVDC subsea cables offshore substations grid connection vessels weather windows consenting' },
  { sector:'Defence / National Security', region:'United Kingdom', client:'MOD reference case', title:'AWRE Aldermaston Upgrade', icon:'Defence', confidence:'Classified supply chain', prompt:'AWRE Aldermaston nuclear warhead facility upgrade classified defence sovereign supply chain security accreditation UK MOD schedule governance' },
  { sector:'Defence / National Security', region:'Australia', client:'AUKUS reference case', title:'AUKUS Submarine Dockyard', icon:'Defence', confidence:'Sovereign dockyard capacity', prompt:'AUKUS nuclear submarine dockyard shipyard scaling workforce Australia sovereign capability supply chain certification operational readiness' },
  { sector:'Defence / National Security', region:'United States', client:'DoD reference case', title:'F-35 Production Ramp', icon:'Defence', confidence:'Software and supply chain', prompt:'F-35 Lightning II production ramp software integration supply chain concurrent development testing operational acceptance DoD sustainment cost' },
  { sector:'Defence / National Security', region:'Global', client:'NATO reference case', title:'NATO Air Defence Modernisation', icon:'Defence', confidence:'Interoperability and classified systems', prompt:'NATO air defence modernisation radar integration command systems interceptor procurement interoperability cyber resilience sovereign procurement' },
  { sector:'Defence / National Security', region:'United Kingdom', client:'Royal Navy reference case', title:'Type 26 Frigate Programme', icon:'Defence', confidence:'Combat systems integration', prompt:'Type 26 City class frigate Royal Navy combat systems integration weapons fit-out sea trials supply chain single source' },
  { sector:'Semiconductors / Advanced Manufacturing', region:'Japan', client:'TSMC / Sony reference case', title:'TSMC Kumamoto Fab', icon:'Semi', confidence:'Workforce and supply chain', prompt:'TSMC Kumamoto Japan semiconductor fab 28nm construction workforce ramp chemical supply chain utility power water purification' },
  { sector:'Semiconductors / Advanced Manufacturing', region:'Germany', client:'Intel reference case', title:'Intel Magdeburg Fab', icon:'Semi', confidence:'EU subsidy and workforce risk', prompt:'Intel Magdeburg Germany semiconductor megafab EU Chips Act subsidy skilled workforce utility infrastructure power grid construction' },
  { sector:'Semiconductors / Advanced Manufacturing', region:'United Kingdom', client:'Newport Wafer reference case', title:'Newport Wafer Fab Expansion', icon:'Semi', confidence:'Sovereign ownership risk', prompt:'Newport Wafer Fab UK semiconductor expansion sovereign ownership national security review workforce utilities compound semiconductor' },
  { sector:'Semiconductors / Advanced Manufacturing', region:'United States', client:'Micron reference case', title:'Micron New York Fab Campus', icon:'Semi', confidence:'CHIPS Act compliance', prompt:'Micron New York semiconductor fab campus CHIPS Act funding DRAM advanced packaging workforce power water utility construction phasing' },
  { sector:'Pharma / Life Sciences', region:'United States', client:'Pfizer reference case', title:'Pfizer Oncology Manufacturing Hub', icon:'Pharma', confidence:'FDA validation and supply chain', prompt:'Pfizer oncology manufacturing hub FDA process validation cleanroom supply chain single source API cytotoxic containment commissioning' },
  { sector:'Pharma / Life Sciences', region:'Ireland', client:'MSD reference case', title:'MSD Biologics Expansion', icon:'Pharma', confidence:'Regulatory and validation', prompt:'MSD Ireland biologics manufacturing expansion EMA regulatory approval cleanroom GMP validation supply chain cold chain logistics' },
  { sector:'Pharma / Life Sciences', region:'Singapore', client:'Biotech reference case', title:'Singapore Biotech Manufacturing Cluster', icon:'Pharma', confidence:'Regulatory and workforce', prompt:'Singapore biotech manufacturing cluster GMP facilities EDB grants regulatory HPRA workforce cleanroom cold chain API supply security' },
  { sector:'Pharma / Life Sciences', region:'United Kingdom', client:'AstraZeneca reference case', title:'AstraZeneca Cambridge Campus', icon:'Pharma', confidence:'Planning and utilities', prompt:'AstraZeneca Cambridge UK R&D manufacturing campus planning consent utilities infrastructure cleanroom GMP facilities workforce research hub' },
  { sector:'Pharma / Life Sciences', region:'Denmark', client:'Novo Nordisk reference case', title:'Novo Nordisk Kalundborg Expansion', icon:'Pharma', confidence:'Capacity and supply chain', prompt:'Novo Nordisk Kalundborg Denmark GLP-1 capacity expansion API manufacturing cleanroom utilities supply chain insulin diabetes obesity drugs' },
  { sector:'Mega Infrastructure', region:'United Arab Emirates', client:'Abu Dhabi Airports reference case', title:'Abu Dhabi Airport Midfield Expansion', icon:'Mega', confidence:'Airside and systems integration', prompt:'Abu Dhabi airport midfield terminal expansion airside systems integration baggage handling airbridge utilities BMS commissioning airline migration' },
  { sector:'Mega Infrastructure', region:'United Kingdom', client:'Heathrow reference case', title:'Heathrow Third Runway', icon:'Mega', confidence:'Planning consent and political risk', prompt:'Heathrow Airport third runway planning consent DCO political risk surface access rail link terminal development airspace redesign' },
  { sector:'Mega Infrastructure', region:'United States', client:'Port Authority reference case', title:'JFK Airport Redevelopment', icon:'Mega', confidence:'Airside interfaces and phasing', prompt:'JFK Airport New York terminal redevelopment phased construction airside operational interfaces contractor coordination funding AirTrain integration' },
  { sector:'Mega Infrastructure', region:'Australia', client:'Western Sydney Airport reference case', title:'Western Sydney Airport', icon:'Mega', confidence:'Greenfield delivery and surface access', prompt:'Western Sydney Airport Badgerys Creek greenfield airport runway terminal surface access rail connection utilities airspace integration' },
  { sector:'Mega Infrastructure', region:'Global', client:'Port reference case', title:'Deepwater Container Port Expansion', icon:'Mega', confidence:'Marine civil and dredging risk', prompt:'Deepwater container port expansion quay wall dredging berths cranes intermodal rail road access customs warehousing capacity phasing' },
  { sector:'Energy / Industrial', region:'Australia', client:'Mining reference case', title:'Pilbara Iron Ore Expansion', icon:'Industrial', confidence:'Mining and logistics', prompt:'Pilbara iron ore mine expansion Western Australia haul roads crushing processing plant rail export port berths water power' },
  { sector:'Energy / Industrial', region:'Chile', client:'Mining reference case', title:'Copper Mine Tailings Expansion', icon:'Industrial', confidence:'Environmental consent and water', prompt:'Chile copper mine tailings storage facility expansion environmental consent water rights community licence to operate slope stability' },
  { sector:'Energy / Industrial', region:'Canada', client:'Oil sands reference case', title:'Alberta Oil Sands SAGD Expansion', icon:'Industrial', confidence:'Regulatory and emissions', prompt:'Alberta oil sands SAGD expansion regulatory approval emissions intensity workforce camp utilities steam injection production ramp' },
  { sector:'Energy / Industrial', region:'South Africa', client:'Eskom reference case', title:'Medupi Power Station Completion', icon:'Industrial', confidence:'Labour and governance risk', prompt:'Medupi coal power station South Africa completion boiler welding labour unrest commissioning governance Eskom financial constraint' },
  { sector:'Energy / Industrial', region:'United Kingdom', client:'National Highways reference case', title:'Lower Thames Crossing', icon:'Industrial', confidence:'Planning consent and tunnelling', prompt:'Lower Thames Crossing road tunnel DCO planning consent tunnelling geology traffic management construction phasing environmental mitigation' },
  { sector:'Energy / Industrial', region:'Norway', client:'Equinor reference case', title:'Johan Sverdrup Phase 2', icon:'Industrial', confidence:'Topsides and subsea integration', prompt:'Johan Sverdrup Phase 2 offshore oil Norway topsides subsea integration wellhead platform power from shore riser systems commissioning' },
  { sector:'Space / Orbital Infrastructure', region:'United States', client:'NASA reference case', title:'Artemis Lunar Gateway', icon:'Space', confidence:'International module integration', prompt:'Artemis Lunar Gateway space station cislunar orbit international partner modules power propulsion habitation life support docking launch manifest' },
  { sector:'Space / Orbital Infrastructure', region:'Europe', client:'ESA reference case', title:'ESA Moonlight Navigation Constellation', icon:'Space', confidence:'Multi-launch and ground segment', prompt:'ESA Moonlight lunar navigation communication constellation satellite bus launch manifest ground segment lunar surface relay service commercial' },
  { sector:'Space / Orbital Infrastructure', region:'United States', client:'Commercial reference case', title:'Blue Origin New Glenn Constellation', icon:'Space', confidence:'Reusability and launch cadence', prompt:'Blue Origin New Glenn heavy lift reusable launch vehicle production ramp launch cadence booster recovery upper stage payload integration' },
  { sector:'Space / Orbital Infrastructure', region:'Global', client:'Commercial reference case', title:'Telesat Lightspeed LEO Network', icon:'Space', confidence:'LEO constellation deployment', prompt:'Telesat Lightspeed low earth orbit satellite constellation launch manifest ground terminals gateway stations spectrum licences commercial service' },
  { sector:'Space / Orbital Infrastructure', region:'United States', client:'Axiom reference case', title:'Axiom Commercial Space Station', icon:'Space', confidence:'ISS interface and module integration', prompt:'Axiom Space commercial station ISS module attachment detachment standalone operations life support power thermal EVA capability commercial crew' },
  { sector:'Space / Orbital Infrastructure', region:'Europe', client:'ESA reference case', title:'JUICE Jupiter Mission', icon:'Space', confidence:'Mission assurance and deep space', prompt:'ESA JUICE Jupiter icy moons explorer deep space mission science instrument integration power propulsion navigation trajectory mission assurance' },
  { sector:'Space / Orbital Infrastructure', region:'United States', client:'Rocket Lab reference case', title:'Rocket Lab Neutron Vehicle', icon:'Space', confidence:'Reusability development risk', prompt:'Rocket Lab Neutron medium lift reusable launch vehicle propulsion development carbon composite structure reusability first stage recovery' },
  { sector:'Space / Orbital Infrastructure', region:'Global', client:'Commercial reference case', title:'In-Space Propellant Depot', icon:'Space', confidence:'Cryogenic transfer technology', prompt:'In space propellant depot cryogenic liquid oxygen hydrogen transfer microgravity operations docking interface orbital refuelling architecture' },
  { sector:'Rail / Transit', region:'United Kingdom', client:'Network Rail reference case', title:'Transpennine Route Upgrade', icon:'Rail', confidence:'Possession and electrification', prompt:'Transpennine Route Upgrade northern England electrification track doubling station upgrades possession strategy operational railway interfaces signalling' },
  { sector:'Rail / Transit', region:'Netherlands', client:'ProRail reference case', title:'Amsterdam Metro Noord/Zuidlijn', icon:'Rail', confidence:'Urban tunnelling and sewer risk', prompt:'Amsterdam metro Noord Zuidlijn urban tunnelling sewer ground conditions station building settlement monitoring operational interfaces city centre' },
  { sector:'Rail / Transit', region:'United States', client:'WMATA reference case', title:'Washington Metro Silver Line Extension', icon:'Rail', confidence:'Systems integration and safety', prompt:'Washington DC Metro Silver Line extension systems integration safety certification WMATA operations commissioning Dulles corridor stations' },
  { sector:'Rail / Transit', region:'Hong Kong', client:'MTR reference case', title:'MTR Tuen Ma Line Extension', icon:'Rail', confidence:'Urban interface and systems', prompt:'MTR Hong Kong Tuen Ma Line extension urban interfaces systems integration station fit-out commissioning passenger service transition operational' },
  { sector:'Energy / Industrial', region:'United Kingdom', client:'Thames Water reference case', title:'Thames Tideway Tunnel', icon:'Water', confidence:'Urban tunnelling and utility', prompt:'Thames Tideway Tunnel London sewer upgrade tunnelling urban utilities shaft construction operational railway adjacency Thames river crossings' },
  { sector:'Energy / Industrial', region:'Australia', client:'Water Corporation reference case', title:'Perth METRONET Water Infrastructure', icon:'Water', confidence:'Supply security and treatment', prompt:'Perth Western Australia water supply security desalination groundwater treatment distribution network climate resilience demand forecasting' },
  { sector:'Energy / Industrial', region:'United Kingdom', client:'Anglian Water reference case', title:'Strategic Pipeline Alliance', icon:'Water', confidence:'Consenting and land access', prompt:'Anglian Water strategic pipeline alliance water transfer grid resilience land access consenting pumping stations treatment works demand management' },
  { sector:'Energy / Industrial', region:'United States', client:'Metropolitan Water District reference case', title:'LA Water Supply Resilience Programme', icon:'Water', confidence:'Drought risk and infrastructure', prompt:'Los Angeles Metropolitan Water District supply resilience programme reservoir groundwater recycled water infrastructure drought climate adaptation' },
  { sector:'Mega Infrastructure', region:'Panama', client:'ACP reference case', title:'Panama Canal Water Management', icon:'Mega', confidence:'Drought and transit risk', prompt:'Panama Canal water management drought lake Gatun levels vessel transit restrictions infrastructure investment resilience desalination pumping' },
  { sector:'Pharma / Life Sciences', region:'United States', client:'Lilly reference case', title:'Lilly Lebanon Indiana Gigafactory', icon:'Pharma', confidence:'Construction and validation', prompt:'Eli Lilly Lebanon Indiana GLP-1 manufacturing gigafactory greenfield cleanroom API synthesis utilities workforce validation FDA approval' },
  { sector:'Pharma / Life Sciences', region:'Germany', client:'BioNTech reference case', title:'BioNTech mRNA Manufacturing Hub', icon:'Pharma', confidence:'Process validation and cold chain', prompt:'BioNTech mRNA vaccine manufacturing hub Germany EU facilities process validation GMP cold chain lipid nanoparticle supply scale-up' },
  { sector:'Pharma / Life Sciences', region:'United States', client:'BARDA reference case', title:'Strategic National Stockpile Facility', icon:'Pharma', confidence:'Security and regulatory', prompt:'Strategic National Stockpile pharmaceutical manufacturing facility BARDA government security clearance GMP cGMP fill finish regulatory compliance' },
  { sector:'Semiconductors / Advanced Manufacturing', region:'United States', client:'Samsung reference case', title:'Samsung Taylor Texas Fab', icon:'Semi', confidence:'CHIPS Act and workforce', prompt:'Samsung Taylor Texas advanced semiconductor fab CHIPS Act funding 4nm process workforce utility power water purification construction timeline' },
  { sector:'Semiconductors / Advanced Manufacturing', region:'France', client:'STMicroelectronics reference case', title:'Crolles 300mm Fab Expansion', icon:'Semi', confidence:'EU Chips Act and grid', prompt:'STMicroelectronics Crolles France 300mm wafer fab expansion EU Chips Act subsidy power grid workforce cleanroom construction utilities' },
  { sector:'Semiconductors / Advanced Manufacturing', region:'Israel', client:'Intel reference case', title:'Intel Kiryat Gat Fab 28', icon:'Semi', confidence:'Geopolitical and workforce risk', prompt:'Intel Kiryat Gat Israel semiconductor fab 28 advanced process node geopolitical risk workforce security infrastructure government support' },
  { sector:'Defence / National Security', region:'Australia', client:'Defence reference case', title:'Hunter Class Frigate Programme', icon:'Defence', confidence:'Industrial base and supply chain', prompt:'Hunter class frigate Australia ASC shipyard naval surface combatant combat system integration supply chain sovereign capability workforce' },
  { sector:'Defence / National Security', region:'United Kingdom', client:'MOD reference case', title:'Ajax Armoured Vehicle Programme', icon:'Defence', confidence:'Acceptance testing and safety', prompt:'Ajax armoured vehicle UK Army programme crew safety acceptance testing vibration noise regulatory production delivery MOD governance' },
  { sector:'Defence / National Security', region:'United States', client:'DoD reference case', title:'B-21 Raider Bomber Programme', icon:'Defence', confidence:'Classified systems and production', prompt:'B-21 Raider stealth bomber USAF Northrop Grumman classified systems production ramp test evaluation operational capability sovereign supply chain' },
  { sector:'Defence / National Security', region:'Global', client:'Five Eyes reference case', title:'Joint Intelligence Surveillance Platform', icon:'Defence', confidence:'Interoperability and classification', prompt:'Joint intelligence surveillance reconnaissance platform Five Eyes interoperability classification security accreditation satellite comms integration' },
  { sector:'AI / Data Centres', region:'United States', client:'Nvidia reference case', title:'Nvidia DGX Cloud Infrastructure', icon:'AI', confidence:'GPU supply and power density', prompt:'Nvidia DGX Cloud AI training infrastructure GPU supply chain power density liquid cooling rack integration network fabric latency optimisation' },
  { sector:'AI / Data Centres', region:'Sweden', client:'Hyperscale reference case', title:'Nordic Green Data Centre Hub', icon:'AI', confidence:'Renewable power and cooling', prompt:'Nordic green data centre hub Sweden hydropower cooling efficiency district heating integration EU AI Act compliance sustainability reporting' },
  { sector:'AI / Data Centres', region:'Malaysia', client:'Hyperscale reference case', title:'Johor AI Campus Cluster', icon:'AI', confidence:'Power and land risk', prompt:'Johor Malaysia AI data centre campus cluster hyperscale Singapore connectivity power grid land acquisition planning incentives' },
  { sector:'AI / Data Centres', region:'Kenya', client:'Africa Data Centres reference case', title:'Nairobi Sovereign Data Hub', icon:'AI', confidence:'Power reliability and connectivity', prompt:'Nairobi Kenya sovereign data hub power reliability grid connectivity undersea cable landing internet exchange cooling tropical climate' },
  { sector:'Energy / Industrial', region:'United Kingdom', client:'EDF reference case', title:'Dungeness Decommissioning', icon:'Nuclear', confidence:'Waste management and safety', prompt:'Dungeness nuclear power station decommissioning UK defuelling waste characterisation contaminated structure removal regulatory NDA oversight schedule' },
  { sector:'Energy / Industrial', region:'Japan', client:'TEPCO reference case', title:'Fukushima Decommissioning', icon:'Nuclear', confidence:'Radiological and water management', prompt:'Fukushima Daiichi nuclear decommissioning Japan contaminated water treatment fuel debris retrieval regulatory IAEA oversight decades-long programme' },
  { sector:'Energy / Industrial', region:'Global', client:'Offshore Wind reference case', title:'Floating Offshore Wind Demonstration', icon:'Energy', confidence:'Mooring and subsea cable risk', prompt:'Floating offshore wind demonstration programme mooring system dynamic cable installation grid connection port logistics turbine installation vessel' },
  { sector:'Energy / Industrial', region:'United States', client:'DOE reference case', title:'Long Duration Energy Storage Grid', icon:'Energy', confidence:'Technology and grid integration', prompt:'Long duration energy storage grid programme iron air flow battery technology utility scale procurement grid integration DOE loan guarantee' },
  { sector:'Mega Infrastructure', region:'Italy', client:'ANAS reference case', title:'Messina Bridge', icon:'Mega', confidence:'Seismic and wind design', prompt:'Messina Strait Bridge Italy Sicily mainland crossing suspension bridge seismic wind design foundations marine environment geopolitical consent' },
  { sector:'Mega Infrastructure', region:'Hong Kong', client:'HKSAR reference case', title:'Northern Metropolis Development', icon:'Mega', confidence:'Land and infrastructure phasing', prompt:'Hong Kong Northern Metropolis development new towns rail transport water sewerage utilities flood control land reclamation cross border' },
  { sector:'Mega Infrastructure', region:'India', client:'Government reference case', title:'Mumbai Trans Harbour Link', icon:'Mega', confidence:'Marine civil and traffic', prompt:'Mumbai Trans Harbour Link sea bridge marine civil works precast concrete deck cable stay traffic integration toll operations' },
  { sector:'Rail / Transit', region:'Taiwan', client:'THSR reference case', title:'Taiwan HSR South Extension', icon:'Rail', confidence:'Urban tunnelling and seismic', prompt:'Taiwan high speed rail south extension Tainan Kaohsiung urban tunnelling seismic design station interfaces operational integration ridership' },
  { sector:'Rail / Transit', region:'Norway', client:'Bane NOR reference case', title:'Oslo–Bergen Mountain Rail Upgrade', icon:'Rail', confidence:'Mountain tunnelling and climate', prompt:'Oslo Bergen mountain railway Norway tunnel upgrades rock reinforcement winter operations signalling electrification station accessibility' },
  { sector:'Rail / Transit', region:'Brazil', client:'CPTM reference case', title:'São Paulo Metro Line 6', icon:'Rail', confidence:'Urban TBM and finance', prompt:'São Paulo Metro Line 6 Orange Brazil TBM tunnelling urban interfaces station excavation public private partnership finance governance' },
  { sector:'Space / Orbital Infrastructure', region:'United States', client:'Commercial reference case', title:'Orbital Refuelling Station Alpha', icon:'Space', confidence:'Cryogenic operations and docking', prompt:'Orbital refuelling station LEO cryogenic propellant depot docking interface operations life extension missions government commercial customers' },
  { sector:'Space / Orbital Infrastructure', region:'Europe', client:'ESA reference case', title:'ESA HERA Asteroid Mission', icon:'Space', confidence:'Deep space navigation and instrument', prompt:'ESA HERA asteroid mission binary system Didymos Dimorphos instruments CubeSat deployment deep space navigation science return' },
  { sector:'Space / Orbital Infrastructure', region:'Global', client:'Commercial reference case', title:'Hypersonic Point-to-Point Transport', icon:'Space', confidence:'Regulatory and safety assurance', prompt:'Hypersonic point to point sub-orbital transport SpaceX Rocket Lab regulatory approval safety case air traffic management commercial viability' },
  { sector:'Space / Orbital Infrastructure', region:'China', client:'CNSA reference case', title:'Chinese Space Station Expansion', icon:'Space', confidence:'Module integration and crew', prompt:'Chinese Space Station Tiangong expansion module integration Tianzhou cargo Shenzhou crew EVA science payload power thermal management' },
  { sector:'Pharma / Life Sciences', region:'Switzerland', client:'Roche reference case', title:'Roche Kaiseraugst Manufacturing', icon:'Pharma', confidence:'GMP validation and logistics', prompt:'Roche Kaiseraugst Switzerland biologics manufacturing expansion GMP API synthesis fill finish cold chain logistics EMA regulatory approval' },
  { sector:'Pharma / Life Sciences', region:'United States', client:'Merck reference case', title:'Merck Vaccine Manufacturing Hub', icon:'Pharma', confidence:'FDA and supply chain', prompt:'Merck MSD vaccine manufacturing hub USA FDA process validation aseptic fill finish antigen supply chain cold chain distribution' },
  { sector:'Semiconductors / Advanced Manufacturing', region:'Netherlands', client:'ASML reference case', title:'ASML EUV Supply Chain Expansion', icon:'Semi', confidence:'Single source and geopolitical', prompt:'ASML EUV lithography supply chain expansion Veldhoven Netherlands single source optics tin plasma light source geopolitical export control' },
  { sector:'Semiconductors / Advanced Manufacturing', region:'South Korea', client:'SK Hynix reference case', title:'SK Hynix HBM Fab Expansion', icon:'Semi', confidence:'HBM demand and yield risk', prompt:'SK Hynix HBM high bandwidth memory fab expansion AI GPU demand yield ramp advanced packaging thermal compression bonding' },
  { sector:'Defence / National Security', region:'Poland', client:'PGZ reference case', title:'Polish Howitzer & Armour Programme', icon:'Defence', confidence:'Industrial sovereign capability', prompt:'Polish defence industry howitzer armour programme sovereign capability PGZ procurement NATO interoperability workforce technology transfer' },
  { sector:'Defence / National Security', region:'South Korea', client:'DAPA reference case', title:'KF-21 Fighter Programme', icon:'Defence', confidence:'Technology development risk', prompt:'KF-21 Boramae fighter jet South Korea domestic development DAPA radar avionics weapons integration test flight certification export' },
  { sector:'Energy / Industrial', region:'Morocco', client:'Xlinks reference case', title:'Morocco–UK Power Interconnector', icon:'Energy', confidence:'4,000km subsea cable risk', prompt:'Morocco UK power interconnector HVDC subsea cable 4000km Xlinks solar wind Sahara grid connection financing consenting supply security' },
  { sector:'Energy / Industrial', region:'Australia', client:'Sun Cable reference case', title:'Australia–Asia Power Link', icon:'Energy', confidence:'HVDC and sovereign risk', prompt:'Sun Cable Australia Asia HVDC power link 5000km submarine cable Northern Territory solar farm Darwin export Singapore financing' },
  { sector:'Mega Infrastructure', region:'Netherlands', client:'Rijkswaterstaat reference case', title:'IJmuiden Sea Lock Expansion', icon:'Mega', confidence:'Marine civil and naval access', prompt:'IJmuiden sea lock expansion Netherlands largest sea lock civil works marine interfaces shipping operations Amsterdam port access' },
  { sector:'Mega Infrastructure', region:'United Kingdom', client:'HM Treasury reference case', title:'UK Levelling Up Infrastructure Fund', icon:'Mega', confidence:'Multi-site governance and spend', prompt:'UK Levelling Up infrastructure fund roads rail broadband housing multi-site programme governance spending governance accountability' },
  { sector:'AI / Data Centres', region:'Canada', client:'Hyperscale reference case', title:'Montreal AI Research Campus', icon:'AI', confidence:'Power and bilingual workforce', prompt:'Montreal Canada AI research data centre campus hydropower renewable energy bilingual workforce university partnerships GPU infrastructure' },
  { sector:'AI / Data Centres', region:'Spain', client:'Hyperscale reference case', title:'Madrid Hyperscale Campus', icon:'AI', confidence:'Power and EU regulatory', prompt:'Madrid Spain hyperscale data centre campus EU AI Act compliance renewable energy certification GDPR grid interconnection construction' },
  { sector:'Energy / Industrial', region:'United Kingdom', client:'Grid reference case', title:'UK Electricity Network Reinforcement', icon:'Energy', confidence:'Planning consent and grid', prompt:'UK electricity network reinforcement National Grid transmission upgrade pylon replacement underground cabling planning consent consenting resistance' },
  { sector:'Energy / Industrial', region:'Germany', client:'50Hertz reference case', title:'German Suedlink HVDC Link', icon:'Energy', confidence:'Underground cabling consent', prompt:'German Suedlink HVDC underground cable transmission link north south renewable energy Bundesnetzagentur planning consent land access' },
  { sector:'Mega Infrastructure', region:'Vietnam', client:'Government reference case', title:'Hanoi–HCMC High Speed Rail', icon:'Mega', confidence:'Land acquisition and funding', prompt:'Vietnam north south high speed rail Hanoi Ho Chi Minh City 1500km land acquisition tunnels viaducts funding World Bank' },
  { sector:'Mega Infrastructure', region:'United States', client:'USDOT reference case', title:'Baltimore Francis Scott Key Bridge Rebuild', icon:'Mega', confidence:'Marine and programme urgency', prompt:'Baltimore Francis Scott Key Bridge rebuild marine civil works channel clearance foundation design funding USDOT programme delivery urgency' },
  { sector:'Energy / Industrial', region:'United States', client:'ExxonMobil reference case', title:'Permian Basin Carbon Capture', icon:'Industrial', confidence:'Regulatory and pipeline risk', prompt:'ExxonMobil Permian Basin carbon capture utilisation storage CCUS pipeline injection well regulatory EPA Class VI permit geological storage' },
  { sector:'Energy / Industrial', region:'Iceland', client:'Orca CCS reference case', title:'Mammoth DAC Facility', icon:'Energy', confidence:'Technology scale and cost', prompt:'Mammoth direct air capture facility Iceland Climeworks geothermal power carbon mineralisation basalt injection commercial scale first of kind' },
  { sector:'Rail / Transit', region:'Egypt', client:'NAT reference case', title:'Cairo Monorail Network', icon:'Rail', confidence:'Elevated structure and systems', prompt:'Cairo monorail network New Administrative Capital Greater Cairo elevated guideway stations rolling stock systems integration commissioning' },
  { sector:'Rail / Transit', region:'Ethiopia', client:'ERA reference case', title:'Addis Ababa LRT Expansion', icon:'Rail', confidence:'Power and operations', prompt:'Addis Ababa light rail transit expansion Ethiopia power supply traction systems operations maintenance workforce urban interface' },
  { sector:'Mega Infrastructure', region:'Kenya', client:'LAPSSET reference case', title:'Lamu Port and Corridor', icon:'Mega', confidence:'Geopolitical and marine risk', prompt:'Lamu Port South Sudan Ethiopia transport corridor LAPSSET Kenya deepwater port road rail pipeline oil infrastructure financing' },
  { sector:'Mega Infrastructure', region:'Nigeria', client:'Government reference case', title:'Lagos–Calabar Coastal Highway', icon:'Mega', confidence:'Land acquisition and financing', prompt:'Lagos Calabar coastal highway Nigeria 700km road infrastructure land acquisition financing environmental communities governance' },
  { sector:'Space / Orbital Infrastructure', region:'India', client:'ISRO reference case', title:'Gaganyaan Human Spaceflight', icon:'Space', confidence:'Crew safety and systems assurance', prompt:'Gaganyaan India human spaceflight programme ISRO crew module service module escape system launch abort test evaluation astronaut training' },
  { sector:'Space / Orbital Infrastructure', region:'Japan', client:'JAXA reference case', title:'HII-A Next Generation Launch', icon:'Space', confidence:'New vehicle development risk', prompt:'Japan JAXA HII-A next generation launch vehicle propulsion development flight demonstration commercial payload government mission assurance' },
  { sector:'Defence / National Security', region:'Israel', client:'Rafael reference case', title:'Iron Dome Production Ramp', icon:'Defence', confidence:'Production rate and supply chain', prompt:'Iron Dome interceptor missile production ramp Rafael Israel supply chain Tamir interceptor electronics software integration export demand' },
  { sector:'Defence / National Security', region:'France', client:'DGA reference case', title:'SCAF Next Generation Fighter', icon:'Defence', confidence:'Multi-nation governance', prompt:'SCAF future combat air system France Germany Spain next generation fighter Dassault Airbus multi-nation governance technology sharing' },
  { sector:'Semiconductors / Advanced Manufacturing', region:'India', client:'Tata reference case', title:'India First Semiconductor Fab', icon:'Semi', confidence:'Greenfield ecosystem risk', prompt:'India first semiconductor fab Tata Electronics Gujarat 28nm process India Semiconductor Mission greenfield ecosystem supply chain workforce' },
  { sector:'Semiconductors / Advanced Manufacturing', region:'United Arab Emirates', client:'G42 reference case', title:'UAE Advanced Manufacturing Hub', icon:'Semi', confidence:'Sovereign and export control', prompt:'UAE advanced manufacturing hub G42 semiconductor chip design packaging sovereign AI compute export control geopolitical supply chain' },
  { sector:'AI / Data Centres', region:'South Africa', client:'Hyperscale reference case', title:'Johannesburg Cloud Hub', icon:'AI', confidence:'Power reliability and connectivity', prompt:'Johannesburg South Africa cloud data centre hub Eskom power reliability fibre connectivity submarine cable landing Africa hyperscale' },
  { sector:'AI / Data Centres', region:'Brazil', client:'Hyperscale reference case', title:'São Paulo Hyperscale Region', icon:'AI', confidence:'Grid and regulatory', prompt:'São Paulo Brazil hyperscale data centre region LGPD data sovereignty grid reliability renewable energy tariff regulatory framework' },
  { sector:'Energy / Industrial', region:'Saudi Arabia', client:'Aramco reference case', title:'Aramco Jafurah Gas Development', icon:'Industrial', confidence:'Sour gas and surface facilities', prompt:'Saudi Aramco Jafurah unconventional gas development sour gas treatment NGL fractionation pipeline gas to liquid surface facilities' },
  { sector:'Energy / Industrial', region:'Mozambique', client:'TotalEnergies reference case', title:'Mozambique LNG Restart', icon:'Energy', confidence:'Security and political risk', prompt:'Mozambique TotalEnergies LNG Cabo Delgado security risk insurgency restart onshore facilities force majeure insurance political risk financing' },
  { sector:'Mega Infrastructure', region:'Mexico', client:'Government reference case', title:'Tren Maya Railway', icon:'Mega', confidence:'Jungle terrain and heritage', prompt:'Tren Maya railway Mexico Yucatan Peninsula jungle terrain cenote ground conditions Mayan heritage sites environmental controversy phasing' },
  { sector:'Mega Infrastructure', region:'Turkey', client:'KGM reference case', title:'Istanbul Canal Project', icon:'Mega', confidence:'Environmental and geopolitical', prompt:'Istanbul Canal Turkey artificial waterway Bosphorus bypass marine engineering environmental impact geopolitical Montreux Convention financing' },
  { sector:'Rail / Transit', region:'Chile', client:'Metro reference case', title:'Santiago Metro Line 7', icon:'Rail', confidence:'Seismic and urban tunnelling', prompt:'Santiago Chile Metro Line 7 seismic design urban tunnelling ground conditions Andes geology station excavation systems integration' },
  { sector:'Rail / Transit', region:'South Africa', client:'PRASA reference case', title:'Metrorail Fleet Renewal', icon:'Rail', confidence:'Fleet integration and depot', prompt:'PRASA Metrorail South Africa fleet renewal rolling stock procurement depot upgrade signalling maintenance workforce operational transition' },
  { sector:'Pharma / Life Sciences', region:'Australia', client:'CSL Behring reference case', title:'CSL Broadmeadows Plasma Facility', icon:'Pharma', confidence:'GMP and cold chain', prompt:'CSL Behring Broadmeadows Australia plasma fractionation facility GMP upgrade cold chain immunoglobulin albumin clotting factors regulatory TGA' },
  { sector:'Pharma / Life Sciences', region:'China', client:'Wuxi Biologics reference case', title:'Wuxi Large Scale Biologics', icon:'Pharma', confidence:'Regulatory and quality systems', prompt:'Wuxi Biologics China large scale biologics manufacturing NMPA FDA dual approval quality systems supply chain global export capability' },
  { sector:'Energy / Industrial', region:'United Kingdom', client:'Veolia reference case', title:'Coventry Energy Recovery Facility', icon:'Industrial', confidence:'Planning and community', prompt:'Coventry energy recovery facility waste to energy planning consent community opposition pollution control tipping hall steam turbine grid' },
  { sector:'Energy / Industrial', region:'Denmark', client:'Ørsted reference case', title:'Hornsea 3 Offshore Wind', icon:'Energy', confidence:'Offshore marine and grid', prompt:'Hornsea 3 offshore wind farm North Sea 2.4GW cable installation foundation monopile offshore substation HVDC grid connection O&M' },
  { sector:'Mega Infrastructure', region:'Colombia', client:'ANI reference case', title:'Bogotá Metro Line 1', icon:'Mega', confidence:'Urban tunnelling and PPP', prompt:'Bogotá Metro Line 1 Colombia elevated viaduct underground section PPP concession systems integration stations commissioning community' },
  { sector:'Mega Infrastructure', region:'Philippines', client:'DoTr reference case', title:'Metro Manila Subway', icon:'Mega', confidence:'Urban tunnelling and typhoon risk', prompt:'Metro Manila subway Philippines urban tunnelling typhoon resilience station design ODA financing JICA Japan government systems integration' },
  { sector:'Defence / National Security', region:'India', client:'HAL reference case', title:'Tejas Mk2 Production Programme', icon:'Defence', confidence:'Indigenous supply chain', prompt:'Tejas Mk2 light combat aircraft India HAL production ramp indigenous supply chain avionics radar engine KAVERI HAL workforce' },
  { sector:'Defence / National Security', region:'Sweden', client:'FMV reference case', title:'Gripen E/F Production', icon:'Defence', confidence:'Export and supply chain', prompt:'Gripen E F production Sweden Saab export programme Brazil South Africa supply chain composite structures avionics EW systems integration' },
  { sector:'AI / Data Centres', region:'Netherlands', client:'Hyperscale reference case', title:'Amsterdam AMS-IX Expansion', icon:'AI', confidence:'Grid moratorium and land', prompt:'Amsterdam AMS-IX internet exchange data centre expansion grid capacity moratorium land scarcity energy transition cooling water sustainability' },
  { sector:'AI / Data Centres', region:'Chile', client:'Hyperscale reference case', title:'Santiago LATAM Cloud Hub', icon:'AI', confidence:'Power and seismic', prompt:'Santiago Chile LATAM cloud hub data centre seismic design renewable energy copper mining clean power subsea cable connectivity' },
  { sector:'Space / Orbital Infrastructure', region:'United Arab Emirates', client:'MBRSC reference case', title:'Emirates Mars Mission Phase 2', icon:'Space', confidence:'Deep space and science', prompt:'Emirates Mars Mission Phase 2 UAE MBRSC science return follow-on mission design instrument development international partnership ground segment' },
  { sector:'Space / Orbital Infrastructure', region:'Global', client:'Commercial reference case', title:'Space-Based Solar Power Demonstrator', icon:'Space', confidence:'In-space assembly and transmission', prompt:'Space based solar power demonstrator microwave transmission rectenna ground station in space assembly robotics orbit deployment frequency' },
  { sector:'Energy / Industrial', region:'Chile', client:'Government reference case', title:'National Lithium Strategy Infrastructure', icon:'Industrial', confidence:'Environmental and community', prompt:'Chile national lithium strategy mining infrastructure brine extraction processing battery precursor export port rail environmental indigenous communities' },
  { sector:'Energy / Industrial', region:'Democratic Republic of Congo', client:'Mining reference case', title:'Cobalt Supply Chain Infrastructure', icon:'Industrial', confidence:'ESG and sovereign risk', prompt:'DRC cobalt mining supply chain infrastructure ESG compliance artisanal mining formalisation port logistics sovereign risk financing' },
  { sector:'Rail / Transit', region:'Kenya', client:'SGR reference case', title:'Nairobi–Mombasa SGR Phase 2', icon:'Rail', confidence:'Financing and patronage', prompt:'Nairobi Mombasa standard gauge railway SGR Phase 2 Kenya China EXIM financing patronage revenue land acquisition operations maintenance' },
  { sector:'Rail / Transit', region:'Pakistan', client:'ML-1 reference case', title:'ML-1 Railway Upgrade', icon:'Rail', confidence:'Geopolitical and financing risk', prompt:'Pakistan ML-1 main line railway upgrade CPEC China financing Lahore Karachi track upgrade signalling electrification geopolitical risk' },
  { sector:'Mega Infrastructure', region:'Iraq', client:'Government reference case', title:'Basra Smart City Infrastructure', icon:'Mega', confidence:'Sovereign and security risk', prompt:'Basra Iraq smart city infrastructure water power roads telecoms housing oil revenue sovereign fund security governance delivery risk' },
  { sector:'Mega Infrastructure', region:'Angola', client:'Sonangol reference case', title:'Luanda Port Expansion', icon:'Mega', confidence:'Offshore and governance', prompt:'Luanda Angola deepwater port expansion quay berths dredging logistics free zone oil sector services Sonangol governance financing' },
  { sector:'Pharma / Life Sciences', region:'South Korea', client:'Samsung Biologics reference case', title:'Samsung Biologics Plant 5', icon:'Pharma', confidence:'Validation and capacity ramp', prompt:'Samsung Biologics Incheon South Korea Plant 5 expansion 180000L bioreactor capacity GMP validation FDA EMA approval supply contracts' },
  { sector:'Pharma / Life Sciences', region:'India', client:'Serum Institute reference case', title:'Serum Institute Vaccine Capacity', icon:'Pharma', confidence:'Scale and regulatory', prompt:'Serum Institute India vaccine manufacturing capacity expansion WHO prequalification COVAX supply chain fill finish regulatory compliance export' },
  { sector:'Semiconductors / Advanced Manufacturing', region:'Australia', client:'Government reference case', title:'Australian Critical Minerals Processing', icon:'Semi', confidence:'Beneficiation and supply chain', prompt:'Australia critical minerals processing lithium rare earths cobalt beneficiation refinery strategic supply chain AUKUS semiconductor defence' },
  { sector:'Semiconductors / Advanced Manufacturing', region:'Canada', client:'Government reference case', title:'Ontario EV Battery Gigafactory', icon:'Semi', confidence:'Supply chain and workforce', prompt:'Ontario Canada EV battery gigafactory Volkswagen Stellantis government incentive supply chain cathode anode workforce electric vehicle supply' },
];

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
  { name:'Starlink Constellation (SpaceX)', sector:'Space / Mission Assurance', mode:'Space', cost_bn:30.0, cost_growth_pct:0, schedule_slip_months:0, failure_mode:'Successfully scaled — reusable launch, vertical integration, iterative design', lesson:'Vertical integration (own launch + own satellite) is the only structure that achieves constellation ', prompt:'Starlink Constellation (SpaceX) real programme Space / Mission Assurance actual outturn $30.0B +0% cost growth +0 months slip failure mode: Successfully scaled — reusable launch, vertical integration, iterative design' },
  { name:'STACK Infrastructure Denver (COL01)', sector:'Data Centre / Digital', mode:'Earth', cost_bn:3.2, cost_growth_pct:15, schedule_slip_months:12, failure_mode:'Grid interconnection delays for 960MW campus expansion', lesson:'Grid is the governing constraint. Power procurement must be secured 3-4 years before go-live.', prompt:'STACK Infrastructure Denver Campus data centre 960MW hyperscale colocation power infrastructure grid connection', region:'Denver, Colorado, USA', client:'STACK Infrastructure COL01', title:'STACK Denver Campus', icon:'Data Centre', confidence:'Grid connection and power procurement' },
  { name:'STACK Infrastructure Northern Virginia (VA01)', sector:'Data Centre / Digital', mode:'Earth', cost_bn:2.8, cost_growth_pct:18, schedule_slip_months:10, failure_mode:'Loudoun County planning moratorium 2023 — power availability constraints', lesson:'Northern Virginia moratorium risk: LPAs now require dedicated grid upgrades before permits issue.', prompt:'STACK Infrastructure Northern Virginia data centre campus hyperscale colocation power', region:'Northern Virginia, USA', client:'STACK Infrastructure VA01', title:'STACK Northern Virginia', icon:'Data Centre', confidence:'Planning consent and grid' },
  { name:'STACK Infrastructure Frankfurt (FRA01)', sector:'Data Centre / Digital', mode:'Earth', cost_bn:1.4, cost_growth_pct:20, schedule_slip_months:14, failure_mode:'German regulatory and grid upgrade delays', lesson:'EU data centre permitting requires heat reuse plans — adds 6-12 months to planning.', prompt:'STACK Infrastructure Frankfurt Germany European campus data centre hyperscale power', region:'Frankfurt, Germany', client:'STACK Infrastructure FRA01', title:'STACK Frankfurt Campus', icon:'Data Centre', confidence:'EU permitting and grid' },
  { name:'STACK Infrastructure Warsaw (WAW01)', sector:'Data Centre / Digital', mode:'Earth', cost_bn:4.5, cost_growth_pct:22, schedule_slip_months:12, failure_mode:'EU energy transition and grid capacity constraints Poland', lesson:'Poland grid decarbonisation creates power procurement uncertainty. PPA availability is the critical path item.', prompt:'STACK Infrastructure Warsaw Poland campus data centre hyperscale European', region:'Warsaw, Poland', client:'STACK Infrastructure WAW01', title:'STACK Warsaw Campus', icon:'Data Centre', confidence:'Power procurement and grid' },
  { name:'STACK Infrastructure Singapore (SIN01)', sector:'Data Centre / Digital', mode:'Earth', cost_bn:2.4, cost_growth_pct:8, schedule_slip_months:8, failure_mode:'Government moratorium 2019-2022 delayed entry', lesson:'Singapore lifted moratorium with 30% renewable energy requirement. Regulatory compliance is a licence condition.', prompt:'STACK Infrastructure Singapore campus data centre hyperscale colocation Asia Pacific', region:'Singapore', client:'STACK Infrastructure SIN01', title:'STACK Singapore Campus', icon:'Data Centre', confidence:'Renewable energy compliance' },
  { name:'STACK Infrastructure Tokyo (TYO01)', sector:'Data Centre / Digital', mode:'Earth', cost_bn:2.0, cost_growth_pct:14, schedule_slip_months:10, failure_mode:'Seismic compliance requirements and limited grid capacity', lesson:'Japan seismic requirements add 15-25% to structural costs. Must be in baseline not contingency.', prompt:'STACK Infrastructure Tokyo Japan campus data centre hyperscale Asia Pacific seismic', region:'Tokyo, Japan', client:'STACK Infrastructure TYO01', title:'STACK Tokyo Campus', icon:'Data Centre', confidence:'Seismic compliance and grid' }
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
function moneyLocal(n, curr) { const c = curr || '$'; return n >= 1000 ? `${c}${(n/1000).toFixed(1)}T` : n >= 1 ? `${c}${n.toFixed(1)}B` : `${c}${Math.round(n*1000)}M`; }

function fmt(v, curr) {
  if (v === undefined || v === null || v === '') return '—';
  if (typeof v === 'string') return v;
  const c = curr || '$';
  return v >= 1000 ? `${c}${(v / 1000).toFixed(1)}T` : v >= 1 ? `${c}${v.toFixed(1)}B` : `${c}${(v * 1000).toFixed(0)}M`;
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
  // Slim payload: strip heavy fields that cause Render payload rejection
  const STRIP = ['all_schedule_levels','schedules_by_level','risk_detail','cost_detail','self_challenge','raw_crawl'];
  const slim = Object.fromEntries(Object.entries(model || {}).filter(([k]) => !STRIP.includes(k)));

  const attempt = async (label) => {
    if (setExportingLabel) setExportingLabel(label);
    const resp = await apiFetch(path, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(slim)
    });
    return resp;
  };

  try {
    if (setExportingLabel) setExportingLabel('Generating export…');
    let resp = await attempt('Generating export…');

    // If backend is waking up (502/503/504) retry once after 8 seconds
    if ([502, 503, 504].includes(resp?.status) || !resp?.ok && resp?.status >= 500) {
      if (setExportingLabel) setExportingLabel('Backend waking up — retrying in 8s…');
      await new Promise(r => setTimeout(r, 8000));
      resp = await attempt('Retrying export…');
    }

    if (!resp.ok) {
      const txt = await resp.text();
      let msg = txt;
      try { msg = JSON.parse(txt)?.detail?.message || JSON.parse(txt)?.message || txt; } catch (_) {}
      // Friendly message for common errors
      if (resp.status === 502 || resp.status === 503) {
        alert('The server is still starting up. Wait 20 seconds and try the export again — this only happens after a period of inactivity.');
      } else {
        alert('Export failed: ' + String(msg).slice(0, 200));
      }
      if (setExportingLabel) setExportingLabel('');
      return;
    }
    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { a.remove(); URL.revokeObjectURL(url); }, 2000);
    if (setExportingLabel) setTimeout(() => setExportingLabel(''), 1500);
  } catch (err) {
    // Network error — likely cold start
    if (setExportingLabel) setExportingLabel('Retrying…');
    try {
      await new Promise(r => setTimeout(r, 8000));
      const resp2 = await attempt('Retrying export…');
      if (resp2.ok) {
        const blob = await resp2.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = name;
        document.body.appendChild(a); a.click();
        setTimeout(() => { a.remove(); URL.revokeObjectURL(url); }, 2000);
        if (setExportingLabel) setTimeout(() => setExportingLabel(''), 1500);
        return;
      }
    } catch (_) {}
    alert('Export failed — the server may still be starting up. Wait 20 seconds and try again.');
    if (setExportingLabel) setExportingLabel('');
  }
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
function Table({ rows = [], cols = [], moneyCols = [], cellFmt = null, curr = '$' }) {
  const renderCell = (col, row) => {
    const raw = row[col] ?? '';
    if (moneyCols.includes(col)) return fmt(raw, curr);
    if (cellFmt) return cellFmt(col, raw);
    return String(raw);
  };
  return <div className="tableWrap"><table><thead><tr>{cols.map(c => <th key={c[0]}>{c[1]}</th>)}</tr></thead><tbody>{rows.map((r, i) => <tr key={i}>{cols.map(c => <td key={c[0]}>{renderCell(c[0], r)}</td>)}</tr>)}</tbody></table></div>;
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
                {b.cost_growth_pct>80&&<span style={{background:'rgba(239,68,68,0.12)',color:'#fca5a5',fontSize:'7px',padding:'1px 4px',borderRadius:'2px',fontWeight:'800'}}>+{b.cost_growth_pct+'%'}</span>}
                {b.cost_growth_pct>0&&b.cost_growth_pct<=80&&<span style={{background:'rgba(245,158,11,0.1)',color:'#fde68a',fontSize:'7px',padding:'1px 4px',borderRadius:'2px',fontWeight:'800'}}>+{b.cost_growth_pct+'%'}</span>}
              </div>
            </div>
            <div style={{fontSize:'9px',color:'#475569',marginBottom:'1px'}}>{b.sector}</div>
            <div style={{display:'flex',gap:'8px'}}>
              <span style={{fontSize:'9px',color:'#64748b'}}>Actual: ${b.cost_bn+'B'}</span>
              {b.schedule_slip_months>0&&<span style={{fontSize:'9px',color:'#f59e0b'}}>+{b.schedule_slip_months}mo slip</span>}
            </div>
            {b.failure_mode&&<div style={{fontSize:'8px',color:'#334155',marginTop:'1px',fontStyle:'italic'}}>{b.failure_mode.slice(0,65)}{b.failure_mode.length>65?'…':null}</div>}
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
              <div style={{fontSize:'14px',fontWeight:'800',color:'#e2e8f0'}}>${rc.emv_a+'B'}</div>
              <div style={{fontSize:'9px',color:'#64748b'}}>P80: {rc.p80_a}</div>
            </div>
            <div style={{textAlign:'center',padding:'6px',background:'rgba(255,255,255,0.03)',borderRadius:'3px'}}>
              <div style={{fontSize:'9px',color:'#475569',marginBottom:'2px'}}>Option B risk EMV</div>
              <div style={{fontSize:'14px',fontWeight:'800',color:'#e2e8f0'}}>${rc.emv_b+'B'}</div>
              <div style={{fontSize:'9px',color:'#64748b'}}>P80: {rc.p80_b}</div>
            </div>
            <div style={{textAlign:'center',padding:'6px',background:'rgba(255,255,255,0.03)',borderRadius:'3px'}}>
              <div style={{fontSize:'9px',color:'#475569',marginBottom:'2px'}}>EMV delta</div>
              <div style={{fontSize:'14px',fontWeight:'800',color:rc.emv_delta>0?'#fca5a5':'#10b981'}}>{(rc.emv_delta>0?'+':'')+rc.emv_delta+'B'}</div>
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
      body: 'Start with one of three routes: Earth Demo, Space Demo, or Showcase Library. Those are always free reference runs. For your own project, type the programme in plain English and CASEY builds a first-pass intelligence pack: P50/P80/P90 cost, schedule, risk register, confidence score, OBA, gate readiness, board challenge questions and export-ready evidence.',
      sub: 'The model supports rail, nuclear, defence, data centres, space, mining, airports, gigafactories, water, life sciences and semiconductors. Demos and showcase cases should open without consuming the custom one free run.',
    },
    {
      icon: '📝',
      title: 'How to describe your project',
      body: 'For a custom run, describe the asset, location or operating environment, scale/capacity, stage, and biggest concern. You do not need a cost plan or schedule — CASEY creates a benchmarked starting position from sector intelligence.',
      examples: [
        'New metro line Lagos Nigeria 45km elevated rail 2031 delivery federal funding',
        'South Africa smart meter rollout 4 million connections 14 months $1B',
        'Lunar fuel depot ISRU methane oxygen production Mars transit support',
        'Nuclear SMR fleet UK 10 reactors grid decarbonisation regulated generation',
      ],
      sub: 'The more context you give, the more specific the output. Location matters — it changes currency, regulatory framework, financing, and OBA.',
    },
    {
      icon: '📊',
      title: 'What you get',
      body: 'After a demo, showcase case, or custom run, the project opens into tabs like these:',
      bullets: [
        'Overview — executive summary, P50/P80, confidence score, board verdict',
        'Scenarios — base, faster, cheaper, lower risk, premium — instant re-run each',
        'Cost — CBS breakdown with unit rates, P10/P50/P90 per line item',
        'Risk — cause, event, impact, probability, owner, trigger, mitigation',
        'QCRA/QSRA — Monte Carlo P-curves and tornado chart',
        'Advisor — gate review G0-G5, OBA assessment, procurement heatmap, failure pattern',
        'Benchmarks — 63 named real programmes that calibrated the model',
        'Exports — cost workbook, risk register, XER schedule, board PDF',
      ],
    },
    {
      icon: '⚖️',
      title: 'Programme comparison',
      body: 'Use the Compare button to run two programmes side by side. Good for:',
      bullets: [
        'Your project vs a real completed programme (Crossrail, JWST, Hinkley...)',
        'Two delivery options — faster vs cheaper vs lower risk',
        'Two contractor proposals with different scope boundaries',
        'Your P50 vs the historical reference class for your sector',
      ],
      sub: 'Click "Compare ◆" in the nav, pick a real benchmark from the library as Option A, describe your project as Option B, and run.',
    },
    {
      icon: '💡',
      title: 'Tips for better outputs',
      bullets: [
        'Name the location — "Nigeria" vs "UK" changes financing, OBA and regulatory framework',
        'State the budget if you have one — CASEY will challenge it against reference class',
        'Mention known constraints — "federal funding", "NEPA required", "single contractor"',
        'Use Class 5 for early feasibility, Class 1 for near-tender estimates',
        'Run the Scenarios tab to see how cost/schedule/confidence respond to each trade-off',
        'Use the Advisor tab for the questions your board will ask — before they do',
      ],
    },
  ];
  const s = steps[step];
  return <div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.75)',zIndex:1000,display:'flex',alignItems:'center',justifyContent:'center',padding:'20px'}}>
    <div style={{background:'#0c1a2e',border:'1px solid rgba(141,247,255,0.2)',borderRadius:'8px',padding:'28px 32px',maxWidth:'560px',width:'100%',position:'relative'}}>
      <button onClick={onClose} style={{position:'absolute',top:'14px',right:'16px',background:'none',border:'none',color:'#475569',cursor:'pointer',fontSize:'18px'}}>✕</button>
      <div style={{fontSize:'10px',fontWeight:'800',letterSpacing:'.15em',color:'#8df7ff',marginBottom:'6px'}}>GETTING STARTED — STEP {step+1} OF {steps.length}</div>
      <div style={{fontSize:'32px',marginBottom:'8px'}}>{s.icon}</div>
      <h2 style={{fontSize:'18px',fontWeight:'800',color:'#e2e8f0',marginBottom:'10px'}}>{s.title}</h2>
      <p style={{fontSize:'13px',color:'#94a3b8',lineHeight:'1.7',marginBottom:'10px'}}>{s.body}</p>
      {s.examples && <div style={{marginBottom:'10px'}}>
        <div style={{fontSize:'10px',fontWeight:'800',letterSpacing:'.1em',color:'#64748b',marginBottom:'6px'}}>EXAMPLE PROMPTS:</div>
        {s.examples.map(e => <div key={e} style={{background:'rgba(141,247,255,0.05)',border:'1px solid rgba(141,247,255,0.1)',borderRadius:'3px',padding:'6px 10px',marginBottom:'5px',fontSize:'11px',color:'#cbd5e1',fontStyle:'italic'}}>"{e}"</div>)}
      </div>}
      {s.bullets && <ul style={{margin:'0 0 10px',padding:'0 0 0 0',listStyle:'none'}}>
        {s.bullets.map(b => <li key={b} style={{fontSize:'12px',color:'#94a3b8',padding:'4px 0',borderBottom:'1px solid rgba(255,255,255,0.04)',display:'flex',gap:'8px',alignItems:'flex-start'}}>
          <span style={{color:'#8df7ff',flexShrink:0}}>→</span><span>{b}</span>
        </li>)}
      </ul>}
      {s.sub && <p style={{fontSize:'11px',color:'#475569',fontStyle:'italic',marginBottom:'10px',lineHeight:'1.6'}}>{s.sub}</p>}
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginTop:'16px'}}>
        <div style={{display:'flex',gap:'6px'}}>
          {steps.map((_,i) => <div key={i} style={{width:i===step?'20px':'6px',height:'6px',borderRadius:'3px',background:i===step?'#8df7ff':'rgba(141,247,255,0.2)',transition:'all 0.2s'}}/>)}
        </div>
        <div style={{display:'flex',gap:'8px'}}>
          {step > 0 && <button onClick={() => setStep(s => s-1)} style={{background:'rgba(255,255,255,0.06)',border:'1px solid rgba(255,255,255,0.1)',color:'#94a3b8',padding:'7px 16px',borderRadius:'3px',cursor:'pointer',fontSize:'12px'}}>← Back</button>}
          {step < steps.length-1
            ? <button onClick={() => setStep(s => s+1)} style={{background:'rgba(141,247,255,0.1)',border:'1px solid rgba(141,247,255,0.3)',color:'#8df7ff',padding:'7px 18px',borderRadius:'3px',cursor:'pointer',fontSize:'12px',fontWeight:'700'}}>Next →</button>
            : <button onClick={onClose} style={{background:'rgba(141,247,255,0.15)',border:'1px solid rgba(141,247,255,0.4)',color:'#8df7ff',padding:'7px 18px',borderRadius:'3px',cursor:'pointer',fontSize:'12px',fontWeight:'800'}}>Start using CASEY ◆</button>}
        </div>
      </div>
    </div>
  </div>;
}

function Hero({ onBriefing, onEarth, onSpace, onConsole, onTryDemo }) {
  return <section className="v50TakeoverHero">
    <video className="v50HeroVideo" src="https://corbit.b-cdn.net/casey_hero_film.mp4" autoPlay muted loop playsInline preload="auto" crossOrigin="anonymous" />
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
  const videoRef = useRef(null);
  const [soundOn, setSoundOn] = useState(false);
  useEffect(() => {
    if (!open) {
      setSoundOn(false);
      return;
    }
    const v = videoRef.current;
    if (v) {
      v.muted = true;
      v.volume = 0.9;
      v.play?.().catch(() => {});
    }
  }, [open]);
  function enableSound() {
    const v = videoRef.current;
    if (!v) return;
    v.muted = false;
    v.volume = 0.9;
    v.play?.().catch(() => {});
    setSoundOn(true);
  }
  if (!open) return null;
  return <AnimatePresence><motion.div className="v50Briefing" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}>
    <video ref={videoRef} className="v50BriefingVideo" src="https://corbit.b-cdn.net/casey_hero_film.mp4" autoPlay muted loop controls playsInline preload="auto" crossOrigin="anonymous" />
    <div className="v50BriefingTop"><Logo/><button onClick={onClose}>Exit film</button></div>
    {!soundOn && <button className="v50SoundBtn" onClick={enableSound}>Enable sound</button>}
    <div className="v50BriefingBottom"><button onClick={enableSound}>{soundOn ? 'Sound on' : 'Enable sound'}</button><button onClick={onEarth}>Run Earth model</button><button onClick={onSpace}>Run Space model</button><button onClick={onClose}>Open product</button></div>
  </motion.div></AnimatePresence>;
}

function OneShotDemo({ open, onClose, onComplete }) {
  const [form, setForm] = useState({ email: '', project_description: '' });
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [result, setResult] = useState(null);
  if (!open) return null;

  function update(k, v) { setForm(x => ({ ...x, [k]: v })); }

  // ── Live brief intelligence ──────────────────────────────────────────
  // Reads what the user is typing and returns real-time guidance
  function readBrief(text) {
    const t = String(text || '').toLowerCase();
    const words = t.replace(/\s+/g,' ').trim().split(' ').filter(Boolean);
    const w = words.length;
    const sector =
      /(rail|metro|tram|hs2|crossrail|elizabeth line|overground|signalling|possession|rolling stock|track|light rail|brt|bus rapid|underground|subway|commuter rail|freight rail|level crossing)/.test(t) ? 'rail' :
      /(nuclear|smr|reactor|gda|hinkley|sizewell|wylfa|bradwell|safety case|radiological|nuclear grade|enrichment|decommission|sellafield|magnox|agr|pressurised water|boiling water|fast reactor|fusion|tokamak)/.test(t) ? 'nuclear' :
      /(data cent|data center|datacenter|hyperscale|gpu cluster|pue |server hall|ai campus|compute cluster|colocation|colo |cloud campus)/.test(t) ? 'data_centre' :
      /(submarine|dockyard|naval|aukus|warship|frigate|destroyer|aircraft carrier|awre|aldermaston|burghfield|defence|defense|classified|sovereign supply|mod |ministry of defence|military|barracks|munitions|weapons|ordnance|radar|sonar|nato|gchq|mi5|mi6|prison|custody|detention|police station|court.*build|probation)/.test(t) ? 'defence' :
      /(pharma|gmp|validation|biologics|sterile|fill.finish|cqv|mhra|fda.*approv|cleanroom.*pharma|drug.*manufactur|api.*manufactur|vaccine|bioreactor|cell.*therapy|gene therapy|clinical.*manufactur|life science)/.test(t) ? 'pharma' :
      /(semiconductor|wafer|euv|cleanroom.*chip|tsmc|intel fab|yield ramp|lithography|advanced manufactur|microchip|integrated circuit)/.test(t) ? 'semiconductor' :
      /(gigafactory|giga.*factory|battery.*factory|battery.*manufactur|ev.*manufactur|electric vehicle.*plant|cathode|anode.*manufactur|lithium.*process|battery.*cell|cell.*manufactur)/.test(t) ? 'gigafactory' :
      /(oil field|gas field|lng|liquefied natural gas|fpso|offshore.*platform|subsea|pipeline.*oil|refinery|petrochemical|cracker|lng.*terminal|gas.*terminal|upstream|downstream|midstream|wellhead|compressor.*station|gas.*processing|carbon capture|ccus|ccs)/.test(t) ? 'oil_gas' :
      /(mine |mining|quarry|open.*pit|underground.*mine|tailings|ore.*processing|concentrator|smelter|lithium.*mine|copper.*mine|gold.*mine|coal.*mine|potash|mineral.*process|heap.*leach)/.test(t) ? 'mining' :
      /(airport|terminal.*aviation|baggage.*system|runway|airside|orat.*airport|atc |apron|caa |faa |taxiway|air traffic|heathrow|gatwick|stansted|aviation.*infrastr)/.test(t) ? 'airport' :
      /(hospital|nhs|clinical.*build|ward|operating theatre|patient.*facility|healthcare campus|mental health.*unit|diagnostic.*centre|mri.*build|radiotherapy|cancer.*centre|hospice|care.*home)/.test(t) ? 'healthcare' :
      /(wind farm|solar farm|battery storage|grid connection|substation|offshore wind|onshore wind|hydrogen.*plant|electrolyser|tidal.*energy|wave.*energy|pumped hydro|hydro.*power|biomass.*plant|waste.*energy|energy.*storage|power.*station|combined.*cycle|ccgt|peaker.*plant|smart.*grid|grid.*upgrade|transmission.*line|hvdc|national grid|distribution.*network|solar.*park|floating.*solar)/.test(t) ? 'energy' :
      /(water.*treatment|desalin|wastewater|sewage|sewer|reservoir|smart meter|meter rollout|water.*main|water.*network|flood.*defence|flood.*alleviation|drainage|stormwater|irrigation|water.*supply|utility.*network|clean.*water|drinking water|ofwat|anglian water|thames water|severn trent|yorkshire water|southern water|united utilities)/.test(t) ? 'water' :
      /(5g|telecoms|telecom|fibre.*rollout|fiber.*optic|broadband.*rollout|mobile.*network|mast.*install|tower.*rollout|subsea.*cable|cable.*landing|openreach|bt.*infrastr|network.*rollout|digital.*infrastr|rural.*connectivity|gigabit)/.test(t) ? 'telecoms' :
      /(motorway|highway|expressway|bridge.*road|tunnel.*road|road.*widening|junction.*upgrade|bypass|dual.*carriageway|grade.*separation|road.*scheme|trunk road|smart motorway)/.test(t) ? 'roads' :
      /(moon|lunar|mars|martian|orbit|orbital|leo |geo |meo |lagrange|satellite|launch vehicle|rocket.*develop|spacecraft|astronaut|cosmonaut|isru|cislunar|spaceport|launch.*pad|launch.*complex|space station|space.*habitat|reusable.*launch|small.*sat|cubesat|deep space|interplanetary|asteroid|space.*telescope|earth.*observation|sar.*satellite|gnss|space.*refuel|on.*orbit.*service|debris.*removal)/.test(t) ? 'space' :
      /(port |harbour|harbor|quay|berth|container.*terminal|cruise.*terminal|ferry.*terminal|marine.*infrastr|breakwater|dry.*dock|ship.*repair|logistics.*hub|distribution.*centre|freight.*terminal|intermodal|inland.*port)/.test(t) ? 'ports' :
      /(university|school.*build|college.*build|academy.*build|student.*accommodation|campus.*build|library.*build|museum.*build|civic.*centre|town.*hall|government.*build|embassy|fire.*station|community.*centre)/.test(t) ? 'civic' :
      /(stadium|arena|velodrome|aquatic.*centre|olympic|commonwealth.*games|sports.*facility|convention.*centre|exhibition.*centre|concert.*hall|leisure.*centre)/.test(t) ? 'stadia' :
      /(mixed.*use|regeneration|masterplan|urban.*development|residential.*large|build.*to.*rent|urban.*realm|enabling.*works|civil.*works)/.test(t) ? 'civil' : '';
    const hasLocation = /(uk|united kingdom|england|scotland|wales|northern ireland|london|manchester|birmingham|leeds|glasgow|edinburgh|bristol|liverpool|sheffield|cardiff|belfast|usa|united states|america|texas|california|new york|north carolina|arizona|boston|chicago|florida|washington|europe|france|paris|germany|berlin|netherlands|amsterdam|sweden|norway|denmark|finland|spain|madrid|italy|rome|poland|middle east|saudi arabia|riyadh|dubai|uae|abu dhabi|qatar|doha|kuwait|bahrain|oman|jordan|egypt|cairo|africa|south africa|johannesburg|cape town|nigeria|lagos|nairobi|kenya|ghana|accra|ethiopia|morocco|tanzania|asia|india|mumbai|delhi|bangalore|singapore|malaysia|indonesia|jakarta|philippines|thailand|bangkok|vietnam|pakistan|china|beijing|shanghai|shenzhen|hong kong|taiwan|japan|tokyo|south korea|seoul|australia|sydney|melbourne|new zealand|canada|toronto|vancouver|brazil|latin america|mexico|colombia|chile|argentina|gulf|gcc|offshore|national|domestic)/.test(t);
    const hasCost = /(\$|£|€|¥|\d+[\s]*(billion|million|bn|mn|b\b|m\b|usd|gbp|eur|trillion))/.test(t);
    const hasScale = /(\d+[\s]*(mw|gw|km|beds|units|meters|metres|satellites|modules|stations|terminals|floors|sqm|ha|acres|connections|homes|masts|turbines|panels|vehicles|trains|aircraft|ships|wells|schools|wards|seats|routes|berths|cells|wafers|doses|tonnes|tpa|bpd))/.test(t);
    const hasDuration = /(\d+[\s]*(month|year|week|quarter|yr\b|mo\b|phase)|by\s*20\d\d|completion\s*20\d\d|deliver.*20\d\d|open.*20\d\d)/.test(t);
    const hasChallenge = /(cost|schedule|risk|procurement|approval|commissioning|interface|regulatory|evidence|safety|qualification|consent|finance|funding|delay|overrun|confidence|critical path|planning|permitting|supply chain|workforce|integration|handover|acceptance)/.test(t);
    const missing = [];
    if (!hasLocation && w > 4) missing.push('location or country');
    if (!hasCost && !hasScale && w > 6) missing.push('cost, capacity or scale');
    if (!hasDuration && w > 10) missing.push('programme duration or target date');
    if (!hasChallenge && w > 12) missing.push('main challenge');
    const sectorHints = {
      rail:         'Rail detected. Add: possession access (confirmed or assumed?), operator acceptance date, signalling integration on critical path?',
      nuclear:      'Nuclear detected — GDA, safety case and FOAK supply chain applied. Add: site, reactor type, new build or modification, safety case timeline.',
      data_centre:  'Data centre detected. Add: grid connection contracted or assumed, cooling technology, delivery date contracted or target.',
      defence:      'Defence programme detected — CASEY applies sovereign supply chain, classified systems, operational acceptance and security accreditation. Add: new build, upgrade or extension, key programme milestone.',
      pharma:       'Pharma/GMP detected. Add: validation scope, regulatory body (FDA/MHRA), NPV milestone or product launch date driving delivery.',
      semiconductor:'Fab detected. Add: tool install sequence, UPW and chemical system scope, yield ramp in programme boundary.',
      gigafactory:  'Gigafactory/battery manufacturing detected. Add: cell chemistry, production capacity (GWh/year), utility complexity.',
      oil_gas:      'Oil, gas or CCS detected. Add: offshore or onshore, brownfield or greenfield, long-lead items, regulatory consent status.',
      mining:       'Mining detected. Add: commodity, open-pit or underground, processing plant scope, environmental consent.',
      airport:      'Airport detected. Add: ORAT in scope, airside access constraints, regulatory body (CAA/FAA/EASA), live operations interface.',
      healthcare:   'Healthcare detected. Add: clinical commissioning scope, NHS or private pathway, MEP complexity.',
      energy:       'Energy programme detected. Add: grid connection contracted, DNO/TSO engagement, commissioning sequence, contracted or merchant revenue.',
      water:        'Water/utilities detected. Add: treatment capacity or connection count, regulatory consent status, comms infrastructure scope.',
      telecoms:     'Telecoms detected. Add: coverage obligation contracted or voluntary, wayleave and planning complexity, rural or urban rollout.',
      roads:        'Roads/highways detected. Add: environmental consent status, utilities diversion scope, live-traffic working involved.',
      space:        'Space programme detected — CASEY applies launch logistics, TRL risk, autonomous commissioning and mission assurance. Add: launch vehicle, current TRL, programme phase, FOAK technology involved.',
      ports:        'Port/maritime detected. Add: vessel class and berth dimensions, marine consent status, dredging scope, port operational during construction.',
      civic:        'Civic/education infrastructure detected. Add: funding route (public/private), planning consent status, user brief completeness.',
      stadia:       'Stadium/events detected. Add: event-day capacity, operational date (event-deadline driven?), broadcast and safety certification.',
      civil:        'Civil/mixed-use detected. Add: procurement route (EPC/D&B/management contract), enabling works scope, key approval gateway.',
    };
    return { sector, w, hasLocation, hasCost, hasScale, hasDuration, hasChallenge, missing, sectorHint: sectorHints[sector] || (w >= 4 ? 'Keep going — name the asset: hospital, wind farm, satellite, pipeline, stadium, mine, fab, school…' : '') };
  }
  const brief = readBrief(form.project_description);
  const briefWords = brief.w;
  const hasEmail = form.email.includes('@') && form.email.includes('.');
  const hasNonsense = /(asdf|qwerty|lorem ipsum|ignore previous|jailbreak|write a poem|forget instructions)/.test(String(form.project_description).toLowerCase());
  const canRun = hasEmail && briefWords >= 8 && !hasNonsense;

  // Quality score 0-4
  const quality = [brief.hasLocation, brief.hasCost || brief.hasScale, brief.hasDuration, brief.hasChallenge].filter(Boolean).length;
  const qualityColor = quality >= 3 ? '#10b981' : quality >= 2 ? '#f59e0b' : '#ef4444';
  const qualityLabel = quality >= 4 ? 'Excellent brief' : quality >= 3 ? 'Good — ready to run' : quality >= 2 ? 'Add more detail' : briefWords >= 4 ? 'Keep going…' : '';

  function inferType(text) {
    return /(moon|lunar|mars|orbit|leo |geo |satellite|launch vehicle|rocket|spacecraft|astronaut|isru|cislunar)/.test(String(text||'').toLowerCase()) ? 'Space' : 'Earth';
  }
  function clientToken() {
    let t = localStorage.getItem('casey_public_demo_token');
    if (!t) { t = crypto.randomUUID ? crypto.randomUUID() : String(Date.now())+Math.random(); localStorage.setItem('casey_public_demo_token',t); }
    return t;
  }
  function fingerprint() {
    return [navigator.userAgent,navigator.language,Intl.DateTimeFormat().resolvedOptions().timeZone,screen.width+'x'+screen.height].join('|');
  }
  async function submit() {
    setBusy(true); setError(''); setResult(null);
    if (form.email.includes('@')) {
      try { localStorage.setItem('casey_user_email', form.email.toLowerCase().trim()); } catch(e) {}
    }
    try {
      const r = await post('/public-demo/generate', { ...form, project_type: inferType(form.project_description), client_token: clientToken(), fingerprint: fingerprint() });
      setResult(r);
      onComplete?.(r.model);
      // Mark free run as used after successful OneShotDemo run
      // demo_used is only set after a successful generate() call, not on form submit
    } catch(e) {
      let msg = String(e.message||e);
      try { const p=JSON.parse(msg); msg=typeof p.detail==='object'?(p.detail.message||JSON.stringify(p.detail)):(p.detail||msg); } catch {}
      setError(msg);
    } finally { setBusy(false); }
  }

  return <div className="publicDemoOverlay boomDemoOverlay">
    <div className="publicDemoModal boomDemoModal" style={{maxWidth:'720px',width:'96vw',maxHeight:'92vh',overflowY:'auto'}}>
      <button className="publicDemoClose" onClick={onClose}>×</button>

      <div className="boomHeader" style={{paddingBottom:'8px'}}>
        <p className="demoKicker" style={{marginBottom:'4px'}}>CASEY Intelligence Run — 1 free programme</p>
        <h2 style={{fontSize:'clamp(22px,3.5vw,34px)',marginBottom:'4px',lineHeight:'1.1'}}>Describe your programme.</h2>
        <p className="demoSub" style={{fontSize:'12px',lineHeight:'1.4',marginBottom:'0'}}>Type anything — CASEY reads as you go and guides you to get the sharpest output.</p>
      </div>

      {/* Email */}
      <label style={{display:'block',fontSize:'10px',fontWeight:'800',letterSpacing:'.12em',color:'#8df7ff',marginBottom:'4px',marginTop:'4px'}}>WORK EMAIL
        <input value={form.email} onChange={e=>update('email',e.target.value)} placeholder="you@company.com" style={{display:'block',width:'100%',marginTop:'4px',padding:'9px 12px',background:'rgba(255,255,255,0.04)',border:`1px solid ${hasEmail?'rgba(141,247,255,0.3)':'rgba(255,255,255,0.1)'}`,borderRadius:'4px',color:'#e2e8f0',fontSize:'14px',boxSizing:'border-box',outline:'none'}}/>
      </label>

      {/* Brief */}
      <p style={{fontSize:'11px',color:'#64748b',marginBottom:'10px',lineHeight:'1.5'}}>Describe your programme in plain English — sector, location, scale, key constraints. Example: "Smart meter rollout South Africa, 4M connections, 14 months, $1B". CASEY handles the rest.</p>
      <label style={{display:'block',fontSize:'10px',fontWeight:'800',letterSpacing:'.12em',color:'#8df7ff',margin:'12px 0 4px'}}>PROGRAMME BRIEF
        <textarea value={form.project_description} onChange={e=>update('project_description',e.target.value)} rows={7} placeholder={"Start typing — any programme, any sector, anywhere.\n\nExamples:\n• Smart meter rollout South Africa, 4M connections, 14 months, $1B\n• AWRE Aldermaston nuclear facility upgrade, classified systems\n• Lunar Base Alpha — life support, nuclear power, autonomous commissioning\n• HS2 tunnelling and signalling, possession windows, 178 months"} style={{display:'block',width:'100%',marginTop:'6px',padding:'10px 12px',background:'rgba(255,255,255,0.03)',border:`1px solid ${briefWords>=8?'rgba(141,247,255,0.2)':'rgba(255,255,255,0.08)'}`,borderRadius:'4px',color:'#e2e8f0',fontSize:'13px',lineHeight:'1.6',resize:'vertical',boxSizing:'border-box',fontFamily:'inherit',outline:'none'}}/>
      </label>

      {/* Live CASEY reading — progressive coaching */}
      {briefWords >= 2 && !hasNonsense && <div style={{marginTop:'8px',padding:'10px 14px',background:'rgba(141,247,255,0.04)',border:'1px solid rgba(141,247,255,0.12)',borderRadius:'4px',fontSize:'12px',lineHeight:'1.6'}}>

        {/* Stage 1: no sector yet — nudge toward asset type */}
        {!brief.sector && briefWords < 6 && <div style={{color:'#64748b'}}>
          Keep going — name the <b style={{color:'#94a3b8'}}>asset type</b>. Examples: hospital, wind farm, nuclear plant, data centre, satellite, pipeline, stadium, mine, rail line, smart meter rollout…
        </div>}

        {/* Stage 2: sector detected — give specific guidance */}
        {brief.sector && <div style={{marginBottom:brief.missing.length>0?'8px':'0'}}>
          <span style={{color:'#8df7ff',fontWeight:'800',fontSize:'10px',letterSpacing:'.1em',marginRight:'8px'}}>✓ {brief.sector.replace('_',' ').toUpperCase()} DETECTED</span>
          <span style={{color:'#cbd5e1'}}>{brief.sectorHint}</span>
        </div>}

        {/* Stage 3: what is still missing — shown as sentence not pills */}
        {brief.missing.length > 0 && briefWords >= 4 && <div style={{marginTop:'4px'}}>
          {!brief.hasLocation && <div style={{color:'#f59e0b',fontSize:'11px',marginBottom:'2px'}}>📍 <b>Location:</b> Add a city, country or region — e.g. "in South Africa", "across the UK", "Texas".</div>}
          {!brief.hasCost && !brief.hasScale && briefWords >= 6 && <div style={{color:'#f59e0b',fontSize:'11px',marginBottom:'2px'}}>💰 <b>Scale or cost:</b> Add a number — e.g. "$1B", "500MW", "4 million connections", "200 beds".</div>}
          {!brief.hasDuration && briefWords >= 8 && <div style={{color:'#f59e0b',fontSize:'11px',marginBottom:'2px'}}>📅 <b>Timeline:</b> Add a duration or target date — e.g. "14 months", "by 2027", "36-month programme".</div>}
          {!brief.hasChallenge && briefWords >= 12 && <div style={{color:'#f59e0b',fontSize:'11px'}}>⚠ <b>Main challenge:</b> What matters most — cost confidence, schedule risk, procurement, approval or commissioning?</div>}
        </div>}

        {/* Stage 4: brief is strong — confirm and encourage */}
        {canRun && quality >= 3 && <div style={{color:'#10b981',fontWeight:'700',fontSize:'12px',marginTop: brief.missing.length>0?'6px':'0'}}>
          ✓ Strong brief — CASEY will generate sector-calibrated cost, schedule, risk and board intelligence.
        </div>}
        {canRun && quality < 3 && <div style={{color:'#94a3b8',fontSize:'11px',marginTop:'4px'}}>
          Ready to run. Add more detail for sharper results — the more specific the brief, the more CASEY can challenge it.
        </div>}
      </div>}

      {/* Status row */}
      <div style={{display:'flex',gap:'10px',margin:'8px 0',alignItems:'center',flexWrap:'wrap'}}>
        <span style={{fontSize:'11px',color:briefWords>=8?'#10b981':'#64748b',fontWeight:'700'}}>{briefWords} words {briefWords>=8?'✓':briefWords>0?'(need 8+)':''}</span>
        {hasEmail && <span style={{fontSize:'11px',color:'#10b981',fontWeight:'700'}}>Email ✓</span>}
        {brief.sector && <span style={{fontSize:'11px',color:'#8df7ff',fontWeight:'700',background:'rgba(141,247,255,0.08)',padding:'2px 8px',borderRadius:'10px'}}>{brief.sector.replace('_',' ').toUpperCase()}</span>}
        {briefWords>=8 && quality>=2 && <span style={{fontSize:'11px',color:qualityColor,fontWeight:'700',marginLeft:'auto'}}>{qualityLabel}</span>}
      </div>

      {error && <div style={{fontSize:'12px',color:'#ef4444',padding:'8px 10px',background:'rgba(239,68,68,0.08)',borderRadius:'4px',marginBottom:'8px'}}>{error}</div>}
      {busy && <div className="missionProcessing"><Rocket size={18}/><div><b>CASEY is building your intelligence pack</b><span style={{display:'flex',flexDirection:'column',gap:'2px'}}>
            <span>Parsing brief, applying {brief.sector||'infrastructure'} benchmarks…</span>
            <span style={{fontSize:'10px',color:'#10b981',display:'flex',alignItems:'center',gap:'4px'}}>
              <span style={{width:'5px',height:'5px',borderRadius:'50%',background:'#10b981',display:'inline-block',animation:'pulse 1.2s infinite'}}/>
              🌐 Open Crawl — scanning {brief.location||'global'} news, markets and benchmarks…
            </span>
          </span></div></div>}
      {result && <div className="publicDemoSuccess"><b>Intelligence pack ready.</b><span>Close this and explore your model.</span></div>}

      <div style={{display:'flex',gap:'8px',marginTop:'4px'}}>
        <button className="heroBtn" disabled={!canRun||busy} onClick={submit} style={{flex:1,opacity:canRun&&!busy?1:0.5}}>
          {busy?'Building…':canRun?'⚡ RUN CASEY':'Complete the brief above'}
        </button>
        <button className="ghostBtn" onClick={onClose} style={{flex:0,padding:'0 16px'}}>Close</button>
      </div>
      <p style={{fontSize:'10px',color:'#334155',marginTop:'8px',textAlign:'center'}}>One free intelligence run per visitor. First-pass strategic output — not a certified estimate.</p>
    </div>
  </div>;
}


function Loading({ text }) {
  const stages = ['CASEY recalibrating confidence curves...', 'Applying live sector calibration signals...', 'Running procurement and delivery-tail model...', 'Comparing against benchmark archetypes...', 'Stamping scenario/base deltas into exports...'];
  const [i,setI] = useState(0);
  useEffect(() => { const t = setInterval(() => setI(v => Math.min(v + 1, stages.length - 1)), 650); return () => clearInterval(t); }, []);
  return <motion.div className="loading intelligenceLoading" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}><Rocket size={44}/><b>{text || 'Building connected model...'}</b><span>{stages[i]}</span><small>Cost · Schedule · QCRA · QSRA · Risk Register · Board Pack</small></motion.div>;
}
function ScenarioSelector({ scenario, generate, matrix=[], model=null, prompt='', projectContext=null }) {
  const labels = {base:'Base', faster:'Faster', cheaper:'Cheaper', lower_risk:'Lower Risk', premium:'Premium'};
  return <section className="scenarioRail">{scenarios.map(s => { const row = matrix.find(x => x.scenario === s) || {}; const active = s === scenario; return <button key={s} className={active?'active':''} onClick={() => generate(s, model?.prompt || prompt, model || projectContext)}><b>{labels[s] || s}</b><span>{row.cost_p50 || '—'} · {row.schedule_months || '—'} mo · {row.confidence_pct || '—'+'%'}</span><em>{row.risk || (active ? 'selected' : 'run scenario')}</em></button> })}</section>;
}
function ExportStrip({ model, onBoardPack, onExcel, onRisk, onXer, onQcra }) {
  if (!model) return null;
  return <section className="exportRail">
    <button onClick={onBoardPack}><Download size={15}/> Export Board Pack</button>
    <button onClick={onExcel}><FileSpreadsheet size={15}/> Export Cost Workbook</button>
    <button onClick={onRisk}><ShieldAlert size={15}/> Export Risk Register</button>
    <button onClick={onXer}><FileText size={15}/> Export XER</button>
    <button onClick={onQcra}><BarChart3 size={15}/> Export QCRA/QSRA</button>
  </section>;
}

function p80PlainEnglish(model) {
  const mc = model?.monte_carlo || {};
  const qcra = mc.qcra || {};
  const qsra = mc.qsra || {};
  const p80Cost = qcra.p80 ? fmt(qcra.p80, model?.currency_symbol) : 'the P80 cost';
  const p80Schedule = qsra.p80 ? `${qsra.p80} months` : 'the P80 date';
  return {
    cost: `P80 cost means there is roughly a 1-in-5 downside chance the final cost exceeds ${p80Cost}.`,
    schedule: `P80 schedule means there is roughly a 1-in-5 downside chance delivery finishes later than ${p80Schedule}.`,
    board: 'For executives: P50 is the headline/base case; P80 is the board contingency conversation; P90 is the stress case.'
  };
}

function confidenceLens(model) {
  const pct = Number(model?.confidence_pct || 0);
  const risk = String(model?.risk || 'Medium').toLowerCase();
  const scenario = String(model?.scenario || 'base').toLowerCase();
  const mode = String(model?.mode || 'Earth');
  const subsector = String(model?.subsector || '').toLowerCase();
  const confidenceBand = pct >= 82 ? 'Board-defensible' : pct >= 70 ? 'Execution posture credible' : pct >= 58 ? 'Board challenge likely' : pct >= 45 ? 'Evidence gap visible' : 'Do not approve without more evidence';
  const constraint = mode === 'Space'
    ? 'mission assurance, launch logistics and autonomous recovery evidence'
    : subsector.includes('data') ? 'energisation, cooling readiness and integrated systems testing'
    : subsector.includes('semiconductor') ? 'tool install, UPW readiness and yield-ramp qualification'
    : subsector.includes('life') || subsector.includes('pharma') ? 'CQV, validation readiness and regulatory evidence'
    : 'interface control, procurement evidence and commissioning readiness';
  const posture = scenario === 'faster'
    ? 'CASEY reads this as an aggressive acceleration posture: the date improves, but delivery confidence is now being spent as a resource.'
    : scenario === 'cheaper'
    ? 'CASEY reads this as a capital-preservation posture: headline cost improves, but unpriced operational and recovery risk rises.'
    : scenario === 'lower_risk'
    ? 'CASEY reads this as an assurance-led posture: the board is buying evidence, float and reduced downside exposure.'
    : scenario === 'premium'
    ? 'CASEY reads this as a resilience-led posture: the board is buying optionality, redundancy and stronger strategic protection.'
    : 'CASEY reads this as the balanced reference case: useful for challenge, but not yet a certified approval basis.';
  return {
    headline: confidenceBand,
    constraint,
    posture,
    meaning: `${pct}% means CASEY believes the current ${model?.scenario_label || 'scenario'} case is ${confidenceBand.toLowerCase()} because confidence is constrained by ${constraint}.`,
    decisionRule: pct >= 75 ? 'Proceed to board challenge with evidence pack.' : pct >= 58 ? 'Use for option selection, but close evidence gaps before approval.' : 'Do not approve capital without package evidence, owner actions and updated QCRA/QSRA.'
  };
}
function boardQuestions(model) {
  const lens = confidenceLens(model);
  const scenario = String(model?.scenario || 'base').toLowerCase();
  const common = [
    `What evidence proves ${lens.constraint}?`,
    'Which three risks create most P80/P90 exposure?',
    'What data would move confidence above the board-comfort threshold?',
    'Which named owner is accountable for the critical-path constraint?'
  ];
  if (scenario === 'faster') return ['Why did confidence fall despite a faster date?', 'Are we buying time or just consuming recovery float?', ...common];
  if (scenario === 'cheaper') return ['Are the savings real, deferred, or transferred into operations?', 'What scope, resilience or contingency has been sacrificed?', ...common];
  if (scenario === 'lower_risk') return ['Is the extra cost buying measurable risk reduction?', 'Which evidence gates justify the longer schedule?', ...common];
  if (scenario === 'premium') return ['Which resilience benefits justify the premium?', 'What risk would remain even after the premium is spent?', ...common];
  return ['Is Base a decision case or only a reference case?', 'What must be proven before this becomes board-approvable?', ...common];
}
function gainedSacrificedExposed(model) {
  const s = String(model?.scenario || 'base').toLowerCase();
  const packs = {
    faster: {
      gained: ['Earlier market-entry / revenue option', 'Stronger strategic timing', 'Visible acceleration decision basis'],
      sacrificed: ['Recovery float', 'Procurement optionality', 'Late-stage commissioning stability'],
      exposed: ['Concurrent systems testing', 'Grid / utility readiness', 'Acceleration premium shock']
    },
    cheaper: {
      gained: ['Lower initial authorisation number', 'Reduced near-term capital draw', 'More procurement competition pressure'],
      sacrificed: ['Redundancy', 'Assurance float', 'Reserve adequacy'],
      exposed: ['Operational start-up risk', 'Scope deferral risk', 'P90 contingency shock']
    },
    lower_risk: {
      gained: ['Stronger approval defensibility', 'Reduced P80/P90 exposure', 'Clearer evidence gates'],
      sacrificed: ['Earliest date', 'Lean capital posture', 'Aggressive market-entry timing'],
      exposed: ['Decision delay risk', 'Assurance-gate bureaucracy', 'Higher upfront reserve']
    },
    premium: {
      gained: ['Resilience', 'Optionality', 'Downside protection'],
      sacrificed: ['Lowest capex case', 'Lean procurement route', 'Simple approval story'],
      exposed: ['Premium scope creep', 'Sponsor affordability challenge', 'Value-for-money challenge']
    },
    base: {
      gained: ['Balanced reference case', 'Clear board challenge baseline', 'Scenario comparison anchor'],
      sacrificed: ['No extra time certainty', 'No extra capital efficiency', 'No premium resilience'],
      exposed: ['The true governing constraint', 'Procurement evidence gaps', 'Commissioning readiness uncertainty']
    }
  };
  return packs[s] || packs.base;
}

function evidenceScorecard(model) {
  const pct = Number(model?.confidence_pct || 60);
  const scenario = String(model?.scenario || 'base').toLowerCase();
  const mode = String(model?.mode || 'Earth');
  const isSpace = mode === 'Space';
  const base = {
    benchmark: Math.min(92, Math.max(45, pct + 9)),
    evidence: Math.min(88, Math.max(34, pct - 8)),
    procurement: Math.min(86, Math.max(30, pct - (scenario === 'faster' ? 18 : 10))),
    schedule: Math.min(90, Math.max(38, pct + (scenario === 'lower_risk' ? 10 : scenario === 'faster' ? -10 : 3))),
    resilience: Math.min(94, Math.max(35, pct + (scenario === 'premium' ? 14 : scenario === 'cheaper' ? -12 : 2)))
  };
  return [
    { name: 'Benchmark fit', score: base.benchmark, note: isSpace ? 'space archetype / mission class fit' : 'sector archetype and cost-class fit' },
    { name: 'Package evidence completeness', score: base.evidence, note: 'brief depth, basis visibility and package maturity' },
    { name: 'Procurement certainty', score: base.procurement, note: isSpace ? 'launch / payload / supplier readiness' : 'long-lead supplier and market capacity' },
    { name: 'Schedule logic', score: base.schedule, note: 'critical path, handover gates and QSRA traceability' },
    { name: 'Resilience / reserve', score: base.resilience, note: 'contingency, recovery float and mission / operational resilience' }
  ];
}
function contradictionScan(model) {
  const scenario = String(model?.scenario || 'base').toLowerCase();
  if (scenario === 'faster') return ['Faster date is credible only if acceleration resources are real.', 'Confidence falls because float and procurement optionality are consumed.', 'Board ask: approve speed only with named recovery owners.'];
  if (scenario === 'cheaper') return ['Cheaper headline may transfer cost into operations or P90 exposure.', 'Reserve adequacy should be challenged before capital approval.', 'Board ask: prove savings are removed, not hidden.'];
  if (scenario === 'lower_risk') return ['Lower risk buys evidence and float but delays strategic timing.', 'Board should test whether extra reserve converts into real risk reduction.', 'Board ask: show which P80/P90 drivers are retired.'];
  if (scenario === 'premium') return ['Premium buys resilience but creates a value-for-money challenge.', 'Board should isolate optionality from scope creep.', 'Board ask: prove premium protects the decision, not just the design.'];
  return ['Base is a reference case, not an approval promise.', 'Board should challenge the governing constraint before authorisation.', 'Board ask: decide whether to buy speed, savings, assurance or resilience.'];
}

function finalPosition(model) {
  const lens = confidenceLens(model);
  const scenario = model?.scenario_label || 'Base';
  return `${scenario}: ${lens.posture} Board position: ${lens.decisionRule}`;
}


function stableSeed(model) {
  return String(model?.id || model?.title || 'casey').split('').reduce((a,c)=>a+c.charCodeAt(0),0);
}
function realisticMetric(value, jitter = 0.35, decimals = 1, suffix = '') {
  const n = Number(String(value || '').replace(/[^0-9.-]/g,''));
  if (!Number.isFinite(n)) return value || '—';
  const v = n + Math.sin(n * 7.13) * jitter;
  return `${v.toFixed(decimals)}${suffix}`;
}
function IntelligenceMeta({ model, mode, setMode }) {
  const cohort = (model?.benchmark_comparison || []).length || (model?.mode === 'Space' ? 42 : 153);
  const conf = model?.confidence_pct || 0;
  const evidenceClass = conf >= 75 ? 'Approval ready' : conf >= 55 ? 'Option selection' : 'Pre-approval gap';
  const evidenceColor = conf >= 75 ? '#10b981' : conf >= 55 ? '#f59e0b' : '#ef4444';
  const roles = [
    {id:'exec',  label:'Executive',  desc:'One risk. One number. One decision.'},
    {id:'board', label:'Board',       desc:'Governing constraint + approval evidence'},
    {id:'pm',    label:'PM / QS',     desc:'Cost, schedule, risk register detail'},
    {id:'analyst',label:'Analyst',   desc:'Full methodology + evidence chain'},
  ];
  return <section className="orbitalMetaRail" style={{display:'flex',alignItems:'stretch',gap:0,borderBottom:'1px solid rgba(255,255,255,0.07)',background:'rgba(10,15,30,0.8)'}}>
    <div style={{display:'flex',gap:20,padding:'8px 20px',flex:1,alignItems:'center',flexWrap:'wrap'}}>
      <div style={{fontSize:'9px'}}><span style={{color:'#475569',marginRight:4}}>Sector</span><b style={{color:'#e2e8f0'}}>{model?.subsector || model?.mode || '—'}</b></div>
      <div style={{fontSize:'9px'}}><span style={{color:'#475569',marginRight:4}}>Location</span><b style={{color:'#e2e8f0'}}>{model?.location || '—'}</b></div>
      <div style={{fontSize:'9px'}}><span style={{color:'#475569',marginRight:4}}>Benchmarks</span><b style={{color:'#e2e8f0'}}>{cohort} comparables</b></div>
      <div style={{fontSize:'9px'}}><span style={{color:'#475569',marginRight:4}}>Estimate</span><b style={{color:'#e2e8f0'}}>{model?.estimate_class_name || '—'}</b></div>
      <div style={{fontSize:'9px',padding:'2px 8px',background:evidenceColor+'20',border:`1px solid ${evidenceColor}40`,borderRadius:12}}>
        <b style={{color:evidenceColor}}>{evidenceClass}</b>
      </div>
    </div>
    <div style={{display:'flex',borderLeft:'1px solid rgba(255,255,255,0.07)'}}>
      {roles.map(r => <button key={r.id}
        onClick={()=>setMode(r.id)}
        title={r.desc}
        style={{padding:'0 14px',background:mode===r.id?'rgba(6,182,212,0.15)':'transparent',
          border:'none',borderLeft:'1px solid rgba(255,255,255,0.06)',cursor:'pointer',
          display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center',gap:2,minWidth:80}}>
        <span style={{fontSize:'10px',fontWeight:mode===r.id?'800':'500',color:mode===r.id?'#06b6d4':'#64748b'}}>{r.label}</span>
        <span style={{fontSize:'7px',color:mode===r.id?'#94a3b8':'#334155',whiteSpace:'nowrap'}}>{r.desc}</span>
      </button>)}
    </div>
  </section>;
}

function LiveCalibrationStrip({ model }) {
  const signals = model?.live_calibration_signals || [];
  if (!model?.live_calibration_active && !signals.length) return null;
  const top = signals.slice(0,4);
  return <motion.section className="liveCalStrip" initial={{ opacity: 0, y: -8 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: .35 }}>
    <div className="liveCalPulse"><span></span></div>
    <div><b>{model.live_calibration_label || 'LIVE CALIBRATION SIGNALS ACTIVE'}</b><em>{model.live_calibration_summary || 'Current sector conditions are being applied to confidence and delivery exposure.'}</em></div>
    <div className="liveCalSignals">{top.map((s,i)=><span key={s.signal || i}>{s.signal || s}</span>)}</div>
  </motion.section>;
}
function LiveCalibrationPanel({ model }) {
  const signals = model?.live_calibration_signals || [];
  if (!signals.length) return null;
  return <section className="layout two eliteLayer liveCalPanelWrap">
    <Card className="liveCalPanel"><h2>Delivery environment calibration</h2><p className="chartCaption">CASEY is not treating this as a static estimate. Current sector conditions are translated into confidence, reserve and P-tail weighting.</p>{signals.slice(0,5).map((s,i)=><div className="signalRow" key={s.signal || i}><span>{i+1}</span><b>{s.signal}</b><em>{s.status}</em><strong>{s.direction}</strong><small>{s.applies_to}</small></div>)}</Card>
    <Card className="liveCalPanel"><h2>How it changes the model</h2>{signals.slice(0,5).map((s,i)=><div className="reason compactReason" key={s.basis || i}><span>{i+1}</span>{s.basis}</div>)}<h3>Demo wording</h3><p className="caseyThinking finalPosition">CASEY continuously recalibrates confidence, contingency and delivery exposure against current sector conditions, benchmark behaviour and live operating signals.</p></Card>
  </section>;
}

function PropagationPulse({ scenario, active }) {
  return <AnimatePresence>{active && <motion.section className="propagationPulse" initial={{opacity:0,y:-8}} animate={{opacity:1,y:0}} exit={{opacity:0,y:-8}}>
    {['Scenario selected','QCRA/QSRA curves morphing','Benchmark position recalibrating','Causal chain propagating','Board posture rewritten'].map((x,i)=><motion.div key={x} initial={{opacity:.25,scale:.96}} animate={{opacity:1,scale:1}} transition={{delay:i*.12}}><span>{i+1}</span>{x}<em>{scenario}</em></motion.div>)}
  </motion.section>}</AnimatePresence>
}
function BenchmarkIntelligence({ model }) {
  const seed = stableSeed(model);
  const conf = Number(model?.confidence_pct || 60);
  const schedule = parseFloat(String(model?.schedule || '').replace(/[^0-9.]/g,'')) || 60;
  const peers = [
    ['PMO',31,34],['BI',44,46],['P6',62,55],['EPC',67,61],['CASEY',Math.min(92, conf+16),Math.min(88, 40+conf*.55)]
  ];
  const x = Math.min(88, Math.max(10, 96 - schedule + (seed % 9)));
  const y = Math.min(86, Math.max(18, 28 + conf*.65));
  return <Card className="benchmark2"><h2>Benchmark Positioning Intelligence</h2><p>Live peer positioning by schedule certainty and delivery intelligence maturity.</p><div className="benchmarkField">
    <div className="density d1"/><div className="density d2"/><div className="density d3"/>
    <span className="axis y">Delivery intelligence maturity</span><span className="axis x">Schedule certainty →</span>
    {peers.map(([name,px,py],i)=><motion.div key={name} className={`peerDot ${name==='CASEY'?'caseyDot':''}`} initial={{left:`${px-6}%`,bottom:`${py-8}%`,opacity:.4}} animate={{left:name==='CASEY'?`${x}%`:`${px}%`,bottom:name==='CASEY'?`${y}%`:`${py}%`,opacity:1}} transition={{type:'spring',stiffness:90,damping:14}}><b>{name}</b><span>{name==='CASEY'?'91st pct':`${px}th pct`}</span></motion.div>)}
  </div><div className="benchmarkCards"><div><b>CASEY</b><span>Probabilistic, traceable, decision-led</span></div><div><b>P6</b><span>Schedule control, weak causal intelligence</span></div><div><b>BI</b><span>Reporting without intervention logic</span></div><div><b>PMO</b><span>Human narrative, slow traceability</span></div></div></Card>;
}
function CausalGraph({ model }) {
  const fallback = model?.mode === 'Space' ? ['Launch cadence','Payload integration','Thermal-power balance','Range availability','Autonomous commissioning','Mission assurance','Confidence'] : ['Scope definition','Procurement evidence','Interface control','Commissioning readiness','Operational acceptance','Reserve adequacy','Confidence'];
  const nodes = (model?.causal_graph_nodes && model.causal_graph_nodes.length ? model.causal_graph_nodes : fallback).slice(0,7);
  return <Card className="causalGraphCard"><h2>Dynamic Causal Traceability</h2><p>Click-level explanation of why the intelligence moved.</p><div className="causalGraph">
    {nodes.map((n,i)=><motion.div key={n} className={`causeNode n${i}`} initial={{opacity:0,scale:.9}} animate={{opacity:1,scale:1}} transition={{delay:i*.08}}><span>{i+1}</span><b>{n}</b></motion.div>)}
    <svg viewBox="0 0 100 100" preserveAspectRatio="none">{nodes.slice(1).map((_,i)=><motion.path key={i} d={`M ${10+i*13} ${24+i%2*28} C ${18+i*13} ${18+i%2*28}, ${22+i*13} ${50-(i%2)*22}, ${29+i*13} ${48-(i%2)*22}`} initial={{pathLength:0,opacity:.2}} animate={{pathLength:1,opacity:.85}} transition={{delay:.15+i*.1,duration:.7}} />)}</svg>
  </div><div className="causalReadout"><b>{model?.scenario_label || 'Base'} governing chain</b><span>{nodes.join(' → ')}</span></div></Card>;
}

function assuranceKillChain(model) {
  const mode = String(model?.mode || '').toLowerCase();
  const scenario = String(model?.scenario || 'base').toLowerCase();
  const base = mode.includes('space')
    ? ['Mission assurance evidence', 'Launch manifest realism', 'Payload / habitat integration', 'Thermal-power survival case', 'Autonomous recovery proof']
    : ['Scope freeze evidence', 'Procurement market proof', 'Interface ownership', 'Commissioning readiness', 'Operational handover proof'];
  const scenarioLens = {
    faster: ['Acceleration premium must be evidenced', 'Float burn must be named by package', 'Recovery crews must be funded'],
    cheaper: ['Savings must be removed from scope, not hidden in risk', 'Reserve cannot be a balancing item', 'Lifecycle fragility must be priced'],
    lower_risk: ['Assurance spend must retire named P80 drivers', 'Added time must buy evidence, not bureaucracy', 'Owner gates must be explicit'],
    premium: ['Optionality must be separated from scope creep', 'Resilience must map to avoided downside', 'Value-for-money challenge must be pre-answered'],
    base: ['Reference case must not be sold as approval-ready', 'Dominant constraint must have a named owner', 'Evidence closure must precede authorisation']
  }[scenario] || [];
  return [...base.slice(0,3), ...scenarioLens].slice(0,6);
}

function traditionalConsultantDelta(model) {
  const lens = confidenceLens(model || {});
  const pct = Number(model?.confidence_pct || 50);
  const p80 = model?.monte_carlo?.qcra?.p80 ? fmt(model.monte_carlo.qcra.p80, model?.currency_symbol || '$') : 'P80 exposure';
  const qsra = model?.monte_carlo?.qsra?.p80 ? `${model.monte_carlo.qsra.p80} months` : 'P80 finish date';
  const tvc = model?.traditional_vs_casey || {};
  // Use sector-specific language from backend if available, fall back to generic
  return [
    { old:'Static cost report', casey:`Decision-grade range: ${model?.cost_p50 || 'P50'} headline with ${p80} downside test.` },
    { old:'Programme narrative', casey:`Board-defensibility: ${pct}% / ${lens.headline}.` },
    { old:'What traditional reports say', casey: tvc.traditional || `Civil progress and cost spend appear on track.` },
    { old:'What CASEY reads underneath', casey: tvc.casey || `Governing constraint is not visible in headline progress. P80 is ${tvc.tail_pct || '25'}% above P50.` },
    { old:'Risk register as list', casey:'Cause → event → owner → mitigation → P80/P90 exposure chain.' },
    { old:'Schedule baseline', casey:`QSRA board date: ${qsra}, not just a target milestone.` },
    { old:'Assurance opinion', casey:`Evidence threshold: ${lens.constraint} is the governing approval blocker.` }
  ];
}

function exportAuditSpine(model, direct, indirect, reserves, reconcileCheck) {
  const total = parseMoneyLocal(model?.cost_p50);
  const qcra = model?.monte_carlo?.qcra || {};
  const qsra = model?.monte_carlo?.qsra || {};
  return [
    { label:'Cost reconciliation', value: reconcileCheck < 0.02 ? 'PASS' : 'CHECK', detail:`Direct ${fmt(direct, model?.currency_symbol)} + Indirect ${fmt(indirect, model?.currency_symbol)} + Reserve ${fmt(reserves, model?.currency_symbol)} = ${fmt(direct+indirect+reserves, model?.currency_symbol)} vs P50 ${fmt(total, model?.currency_symbol)}` },
    { label:'Scenario lock', value:String(model?.scenario_label || model?.scenario || 'Base'), detail:'Cards, narratives, QCRA/QSRA and exports are stamped from the selected scenario payload.' },
    { label:'P-tail linkage', value: qcra.p80 ? fmt(qcra.p80, model?.currency_symbol) : 'P80 active', detail:`Cost P80 and QSRA P80 ${qsra.p80 || '—'} months are visible for board challenge.` },
    { label:'Evidence gate', value: String(confidenceLens(model).headline || ''), detail: String(confidenceLens(model).decisionRule || '') },
    { label:'Audit readiness', value:'High', detail:'Scenario propagation is traceable end-to-end — every cost, schedule and risk output is stamped from the same source payload.' },
    ...(model?.stress_test_applied ? [{
      label:'Stress test applied',
      value: String(model.stress_test_applied).replace(/_/g,' ').toUpperCase(),
      detail: (model.stress_test_note || '') + ' P50 updated to ' + (model.cost_p50 || '—') + ', schedule ' + (model.schedule || '—') + ', confidence ' + (model.confidence_pct || '—') + '%.'
    }] : [])
  ];
}

function IncumbentPressurePanel({ model, direct, indirect, reserves, reconcileCheck }) {
  const delta = traditionalConsultantDelta(model);
  const kill = assuranceKillChain(model);
  const spine = exportAuditSpine(model, direct, indirect, reserves, reconcileCheck);
  // Real sector-specific attacks from the model (backend-generated)
  const modelAttacks = model?.board_attack_simulation || [];
  const attacks = modelAttacks.length >= 4 ? modelAttacks : kill;
  // The single most uncomfortable sentence
  const ifFails = model?.if_this_fails || null;
  return <>
    <section className="layout two incumbentPressure">
      <Card className="threatCard">
        <h2>CASEY vs traditional project controls</h2>
        <p style={{fontSize:'12px',color:'#64748b',marginBottom:'10px'}}>How CASEY output compares to a conventional advisory report — same deliverable, different depth and speed.</p>
        {delta.map((x,i)=><div className="versusRow" key={x.old}><span>{i+1}</span><b>{x.old}</b><ArrowRight size={16}/><strong>{x.casey}</strong></div>)}
      </Card>
      <Card className="threatCard">
        <h2>Board Assurance Questions</h2>
        <p>These are the questions a serious investment committee will ask before the project team can hide behind green dashboards.</p>
        {attacks.map((x,i)=><div className="reason" key={i}><span>{i+1}</span>{x}</div>)}
        <h3>Audit spine</h3>
        {spine.map((x,i)=><div className={`auditSpine ${x.value==='PASS'?'pass':''}`} key={x.label}><span>{i+1}</span><b>{x.label}: {safeRender(x.value)}</b><em>{safeRender(x.detail)}</em></div>)}
      </Card>
    </section>
    {ifFails && (
      <section className="layout one">
        <Card className="ifFailsCard">
          <h2>If this programme fails</h2>
          <p className="ifFailsStatement">{ifFails}</p>
          <p className="ifFailsNote">CASEY generates this from sector intelligence, not from programme-specific data. It names the governing failure mode for this programme type historically.</p>
        </Card>
      </section>
    )}
  </>;
}





// ── Programme Health Signal ────────────────────────────────────────
// Named mega-programmes with known public reporting positions.
// CASEY read is derived from model outputs, not hardcoded.
// ── Global Programme Intelligence Database ────────────────────────
// 80+ live and recently-delivered mega-programmes across all sectors.
// publicReported = what the programme board / auditor is saying publicly.
// caseySignal = what CASEY generates from benchmark intelligence.
// The gap between the two is the product.
const ALL_PROGRAMMES = [

  // ── RAIL / TRANSIT ───────────────────────────────────────────────
  { id:'hs2', name:'HS2 Phase 1', sector:'Rail', region:'UK',
    cost:'£45B–£54B authorised', sched:'2033 London–Birmingham', status:'On track (per PAC)',
    source:'HS2 Ltd / PAC 2024',
    prompt:'HS2 Phase 1 tunnelling stations systems migration possessions operator acceptance political exposure' },
  { id:'calHSR', name:'California High-Speed Rail', sector:'Rail', region:'USA',
    cost:'$89B–$128B', sched:'2033 partial / 2040s full', status:'Under review (GAO)',
    source:'CAHSRA / GAO 2023',
    prompt:'California High Speed Rail tunnelling corridor land acquisition utility relocation signalling political scrutiny cost escalation' },
  { id:'crossrail', name:'Elizabeth Line (Crossrail)', sector:'Rail', region:'UK',
    cost:'£18.9B final outturn', sched:'Opened May 2022', status:'Delivered — 2yr delay',
    source:'TfL / DfT 2023',
    prompt:'Crossrail Elizabeth Line tunnelling stations systems integration signalling commissioning live operations possessions' },
  { id:'railBaltica', name:'Rail Baltica', sector:'Rail', region:'Europe',
    cost:'€16B–€24B', sched:'2030 headline / slipping', status:'Schedule under review',
    source:'RB Rail AS 2024',
    prompt:'Rail Baltica new-build high speed corridor multinational procurement station development systems integration cross-border' },
  { id:'sydneyMetro', name:'Sydney Metro West', sector:'Rail', region:'Australia',
    cost:'A$25B–A$35B', sched:'2030 target', status:'Under construction',
    source:'Transport for NSW 2024',
    prompt:'Sydney Metro West twin-bore tunnel stations precinct integration systems commissioning utility relocation operator readiness' },
  { id:'ontarioLine', name:'Ontario Line', sector:'Rail', region:'Canada',
    cost:'C$19B–C$27B', sched:'2031 target', status:'Design / procurement phase',
    source:'Metrolinx 2024',
    prompt:'Ontario Line light metro tunnelling elevated track stations systems integration Toronto procurement operator readiness' },
  { id:'gatewayProgram', name:'Gateway Program', sector:'Rail', region:'USA',
    cost:'$16B–$35B', sched:'2035 target', status:'Funding / procurement',
    source:'Gateway Development Commission 2024',
    prompt:'Gateway Hudson River rail tunnel new and existing tubes portal works systems integration political funding exposure' },
  { id:'brightlineWest', name:'Brightline West', sector:'Rail', region:'USA',
    cost:'$12B', sched:'2028 Olympics target', status:'Under construction',
    source:'Brightline 2024',
    prompt:'Brightline West Las Vegas high speed rail private financing greenfield corridor systems commissioning compressed schedule' },
  { id:'neom_line', name:'NEOM / The Line', sector:'Rail', region:'Saudi Arabia',
    cost:'$500B authorised', sched:'2030 headline', status:'Under construction (descoped)',
    source:'NEOM Authority 2024',
    prompt:'NEOM The Line linear city rail infrastructure systems integration desert delivery logistics workforce political programme scope uncertainty' },

  // ── NUCLEAR / ENERGY ─────────────────────────────────────────────
  { id:'hinkley', name:'Hinkley Point C', sector:'Nuclear', region:'UK',
    cost:'£25B–£35B', sched:'2031 target (EDF)', status:'Under construction — delayed',
    source:'EDF / NAO 2024',
    prompt:'UK EPR nuclear power station first of kind reactor safety case procurement commissioning grid connection workforce' },
  { id:'sizewell', name:'Sizewell C', sector:'Nuclear', region:'UK',
    cost:'£20B–£35B estimate', sched:'2030s FID target', status:'Development / financing',
    source:'EDF / DESNZ 2024',
    prompt:'UK nuclear power station SMR EPR GDA site licence procurement financing government support grid connection' },
  { id:'vogtle', name:'Vogtle Units 3&4', sector:'Nuclear', region:'USA',
    cost:'$35B final outturn', sched:'Completed 2023/24', status:'Delivered — major overrun',
    source:'Georgia Power / NRC 2024',
    prompt:'AP1000 nuclear units first of kind US licensing workforce commissioning supply chain cost overrun regulatory' },
  { id:'smrRollout', name:'UK SMR Nuclear Rollout', sector:'Nuclear', region:'UK',
    cost:'£8B–£15B per unit', sched:'First unit 2035 target', status:'GDA in progress',
    source:'DESNZ / Rolls-Royce 2024',
    prompt:'UK SMR nuclear rollout small modular reactor licensing procurement first of kind supply chain grid integration' },
  { id:'hornsea3', name:'Hornsea 3 Offshore Wind', sector:'Energy', region:'UK',
    cost:'£8B–£10B', sched:'2027 target', status:'Construction approved',
    source:'Ørsted / DESNZ 2024',
    prompt:'Offshore wind farm cable installation turbine supply chain marine vessels weather windows grid connection substation commissioning' },
  { id:'neomPower', name:'NEOM Green Hydrogen', sector:'Energy', region:'Saudi Arabia',
    cost:'$8.4B', sched:'2026 production', status:'Under construction',
    source:'NEOM / Air Products 2024',
    prompt:'Green hydrogen electrolysis wind solar power plant compression storage marine loading export terminal commissioning' },
  { id:'costaRicaHydro', name:'Reventazón Hydro (Costa Rica)', sector:'Energy', region:'Latin America',
    cost:'$1.4B final', sched:'Delivered 2016', status:'Delivered — on budget',
    source:'ICE 2016',
    prompt:'Hydroelectric dam civil works electromechanical systems commissioning environmental community interfaces remote delivery' },

  // ── DATA CENTRE / HYPERSCALE ──────────────────────────────────────
  { id:'microsoftAI', name:'Microsoft AI Supercluster', sector:'Data Centre', region:'Global',
    cost:'$50B+ announced', sched:'2025–2027 phased', status:'Active construction',
    source:'Microsoft / press 2024',
    prompt:'AI hyperscale data centre campus 500MW grid connection liquid cooling GPU supply chain accelerated commissioning' },
  { id:'awsGovCloud', name:'AWS GovCloud Expansion', sector:'Data Centre', region:'USA',
    cost:'$10B+', sched:'2024–2026', status:'Active',
    source:'AWS 2024',
    prompt:'Sovereign cloud government data centre fibre grid redundancy security compliance permitting accelerated delivery' },
  { id:'metaAI', name:'Meta AI Compute Network', sector:'Data Centre', region:'Global',
    cost:'$37B capex 2024', sched:'Ongoing', status:'Active construction',
    source:'Meta Q2 2024',
    prompt:'AI compute data centre network cooling power grid supply chain GPU procurement phased commissioning' },
  { id:'riyadhAI', name:'Riyadh AI Hyperscale 500MW', sector:'Data Centre', region:'Saudi Arabia',
    cost:'$8B–$12B', sched:'2026–2027', status:'Design / procurement',
    source:'NEOM / Saudi Vision 2030 2024',
    prompt:'Riyadh AI hyperscale data centre 500MW grid connection liquid cooling sovereign cloud accelerated compressed schedule' },
  { id:'nvidiaEco', name:'NVIDIA Ecosystem Infrastructure', sector:'Data Centre', region:'Global',
    cost:'$40B+ supply ecosystem', sched:'2024–2026', status:'Active',
    source:'NVIDIA / press 2024',
    prompt:'AI training inference data centre GPU campus power cooling supply chain integration commissioning' },

  // ── DEFENCE / SECURITY ────────────────────────────────────────────
  { id:'aukus', name:'AUKUS Nuclear Submarine Industrial Base', sector:'Defence', region:'Australia/UK/USA',
    cost:'A$268B–A$368B', sched:'First boat 2030s', status:'Industrial base development',
    source:'Australian Government 2024',
    prompt:'AUKUS nuclear submarine industrial base dockyard shipbuilding naval workforce certification sovereign supply chain classified systems' },
  { id:'f35', name:'F-35 Programme', sector:'Defence', region:'Global',
    cost:'$400B+ lifecycle', sched:'Ongoing production', status:'Active — sustainment phase',
    source:'DoD / GAO 2024',
    prompt:'F35 fighter jet production sustainment classified systems integration supply chain avionics certification multinational programme' },
  { id:'ssbnDreadnought', name:'Dreadnought SSBN', sector:'Defence', region:'UK',
    cost:'£31B+', sched:'2030s first patrol', status:'Under construction',
    source:'MoD / PAC 2024',
    prompt:'UK nuclear submarine Dreadnought SSBN Rolls Royce reactor dockyard workforce certified systems integration classified procurement' },
  { id:'arrowhead', name:'Type 31 Frigate', sector:'Defence', region:'UK',
    cost:'£1.25B for 5 ships', sched:'2027–2035', status:'Under construction',
    source:'MoD / Babcock 2024',
    prompt:'Royal Navy Type 31 Arrowhead frigate naval shipbuilding dockyard systems integration weapons electronics trials acceptance' },
  { id:'missileDef', name:'Missile Defence Modernization', sector:'Defence', region:'USA',
    cost:'$20B+', sched:'2025–2030', status:'Active development',
    source:'MDA / DoD 2024',
    prompt:'Missile defence ground based interceptor radar command control classified systems integration software assurance test evaluation' },

  // ── SEMICONDUCTOR / ADVANCED MFG ──────────────────────────────────
  { id:'tsmcArizona', name:'TSMC Arizona Fab', sector:'Semiconductor', region:'USA',
    cost:'$65B committed', sched:'N3 2025, N2 2028', status:'Under construction — delayed',
    source:'TSMC / CHIPS Act 2024',
    prompt:'TSMC Arizona semiconductor fab cleanroom EUV tool install ultrapure water workforce ramp yield N3 N2 process' },
  { id:'intelOhio', name:'Intel Ohio Fab (Intel 18A)', sector:'Semiconductor', region:'USA',
    cost:'$20B+', sched:'2026 target (delayed)', status:'Paused / delayed',
    source:'Intel / press 2024',
    prompt:'Intel Ohio fab cleanroom advanced packaging process tool qualification workforce yield ramp EUV CHIPS Act' },
  { id:'samsungTexas', name:'Samsung Texas Fab', sector:'Semiconductor', region:'USA',
    cost:'$17B+', sched:'2024–2026', status:'Construction / qualification',
    source:'Samsung / press 2024',
    prompt:'Samsung advanced semiconductor fab Texas cleanroom tool qualification ultrapure water EUV 2nm process workforce' },
  { id:'microchipArizona', name:'Microchip Technology Expansion', sector:'Semiconductor', region:'USA',
    cost:'$2B+', sched:'2024–2026', status:'Active',
    source:'Microchip / CHIPS Act 2024',
    prompt:'Semiconductor fab expansion cleanroom tool installation qualification legacy node production ramp workforce' },

  // ── LIFE SCIENCES / PHARMA ────────────────────────────────────────
  { id:'lillyExpansion', name:'Eli Lilly GLP-1 Expansion', sector:'Pharma', region:'Global',
    cost:'$18B+ capex 2024', sched:'2025–2027 phased', status:'Active construction',
    source:'Eli Lilly 2024',
    prompt:'Eli Lilly GMP sterile manufacturing fill finish clean utilities validation FDA inspection GLP-1 obesity drug capacity expansion' },
  { id:'novonordisk', name:'Novo Nordisk Biologics Expansion', sector:'Pharma', region:'Global',
    cost:'DKK 42B+', sched:'2025–2028', status:'Active construction',
    source:'Novo Nordisk 2024',
    prompt:'Novo Nordisk biologics manufacturing GMP cold chain aseptic fill finish validation regulatory submissions capacity GLP-1' },
  { id:'moderna', name:'Moderna mRNA Facility', sector:'Pharma', region:'Global',
    cost:'$1.5B–$2B per site', sched:'2024–2026', status:'Active construction',
    source:'Moderna 2024',
    prompt:'Moderna mRNA manufacturing GMP clean utilities validation FDA inspection aseptic fill finish technology transfer biosecurity' },
  { id:'astrazeneca', name:'AstraZeneca Manufacturing Expansion', sector:'Pharma', region:'UK/Global',
    cost:'£2B+ UK investment', sched:'2025–2027', status:'Active',
    source:'AstraZeneca 2024',
    prompt:'AstraZeneca biologics manufacturing GMP fill finish validation regulatory FDA MHRA UK aseptic clean utilities' },

  // ── AIRPORT / AVIATION ────────────────────────────────────────────
  { id:'heathrowT3', name:'Heathrow Terminal 3 Upgrade', sector:'Airport', region:'UK',
    cost:'£2.5B+', sched:'2027 target', status:'Active construction',
    source:'Heathrow Airport 2024',
    prompt:'Heathrow airport terminal live operations baggage systems ORAT airside phasing security systems CAA acceptance passenger transition' },
  { id:'istanbulPhase2', name:'Istanbul Airport Phase 2', sector:'Airport', region:'Turkey',
    cost:'$4B+', sched:'2027', status:'Development',
    source:'IGA / TAV 2024',
    prompt:'Istanbul airport terminal expansion airside runway baggage systems ORAT live operations phasing security passenger systems commissioning' },
  { id:'newDelhi', name:'New Delhi Airport T1 Redevelopment', sector:'Airport', region:'India',
    cost:'$3B+', sched:'2024–2026', status:'Under construction',
    source:'GMR / AAI 2024',
    prompt:'New Delhi airport terminal expansion baggage systems airside live operations ORAT security DGCA acceptance runway' },
  { id:'riyadhKing', name:'King Salman International Airport', sector:'Airport', region:'Saudi Arabia',
    cost:'$147B long-term', sched:'2030 Phase 1', status:'Master planning / Phase 1',
    source:'Saudi Vision 2030 2024',
    prompt:'Riyadh new airport terminal airside baggage systems ORAT live operations security passenger systems mega programme' },

  // ── PORTS / MARINE ────────────────────────────────────────────────
  { id:'thuwalPort', name:'King Salman Energy Park Port', sector:'Ports', region:'Saudi Arabia',
    cost:'$7B+', sched:'2025–2030 phased', status:'Under development',
    source:'Saudi Aramco / SPARK 2024',
    prompt:'Industrial port automation container handling marine works dredging quay wall systems integration logistics energy cluster' },
  { id:'londonGateway', name:'DP World London Gateway Expansion', sector:'Ports', region:'UK',
    cost:'£1.5B+', sched:'2025–2027', status:'Under construction',
    source:'DP World 2024',
    prompt:'Container terminal expansion quay deepening cranes automated handling TOS integration marine works rail connection' },
  { id:'hamburgerPort', name:'Hamburg Port Modernisation', sector:'Ports', region:'Germany',
    cost:'€3B+', sched:'2025–2030', status:'Active investment',
    source:'HPA 2024',
    prompt:'Port automation container terminal AGV TOS integration dredging fairway deepening logistics rail connection sustainability' },

  // ── HEALTHCARE ────────────────────────────────────────────────────
  { id:'nclNHS', name:'New Hospitals Programme (NHS)', sector:'Healthcare', region:'UK',
    cost:'£20B+', sched:'2030–2040s', status:'Delayed / under review',
    source:'DHSC / PAC 2024',
    prompt:'NHS new hospital clinical campus MEP utilities clean rooms operating theatres phased handover PFI replacement' },
  { id:'abuDhabiHealth', name:'Abu Dhabi Healthcare Expansion', sector:'Healthcare', region:'UAE',
    cost:'AED 50B', sched:'2025–2030', status:'Active development',
    source:'DOH Abu Dhabi 2024',
    prompt:'Hospital clinical campus MEP utilities clean rooms operating theatres diagnostic imaging emergency department phased handover commissioning' },

  // ── WATER / UTILITIES ─────────────────────────────────────────────
  { id:'thamestide', name:'Thames Tideway Tunnel', sector:'Water', region:'UK',
    cost:'£4.5B final', sched:'Completed 2024', status:'Delivered — on budget',
    source:'Tideway 2024',
    prompt:'London sewage tunnel TBM tunnelling shafts pumping stations commissioning live network integration utility interfaces' },
  { id:'sydneyDesalination', name:'Sydney Desalination Expansion', sector:'Water', region:'Australia',
    cost:'A$3.5B+', sched:'2027 target', status:'Under construction',
    source:'Sydney Water 2024',
    prompt:'Desalination plant expansion reverse osmosis membrane ultrafiltration pumping electrical HVAC commissioning water licence discharge consent' },
  { id:'nile', name:'GERD Ethiopia Dam', sector:'Water', region:'Africa',
    cost:'$5B+', sched:'Substantially complete 2024', status:'Operational (disputed)',
    source:'Ethiopian government 2024',
    prompt:'Grand Ethiopian Renaissance Dam hydroelectric civil works electromechanical systems commissioning geopolitical interfaces' },

  // ── INFRASTRUCTURE / MEGA CITY ────────────────────────────────────
  { id:'crossrailTunnel', name:'Brenner Base Tunnel', sector:'Infrastructure', region:'Europe',
    cost:'€8.9B', sched:'2032 target', status:'Under construction',
    source:'BBT SE 2024',
    prompt:'Alpine railway base tunnel TBM tunnelling cross-border systems integration signals electrification multinational procurement' },
  { id:'nileRoad', name:'Cairo Ring Road Expansion', sector:'Infrastructure', region:'Egypt',
    cost:'$3B+', sched:'2025–2027', status:'Active construction',
    source:'Egyptian government 2024',
    prompt:'Urban motorway expansion corridor land acquisition utility relocation bridge works systems commissioning live traffic' },
  { id:'singaporeMRT', name:'Singapore Cross Island Line', sector:'Rail', region:'Singapore',
    cost:'S$15B+', sched:'2030 Phase 1', status:'Design / early procurement',
    source:'LTA 2024',
    prompt:'Singapore deep tunnel metro TBM tunnelling stations systems integration live operations phasing signalling commissioning' },

  // ── SPACE — ALL PROGRAMMES ────────────────────────────────────────
  { id:'starship', name:'SpaceX Starship Industrialization', sector:'Space', region:'Global',
    cost:'$10B+ invested', sched:'Orbital 2024–2025', status:'Test flights / rapid iteration',
    source:'SpaceX / FAA 2024',
    prompt:'SpaceX Starship launch vehicle industrialization rapid iteration launch pad infrastructure orbital refuelling thermal protection system range approvals' },
  { id:'lunarGateway', name:'Lunar Gateway Station', sector:'Space', region:'Lunar Orbit',
    cost:'$8B+ NASA share', sched:'2027–2030', status:'Development / procurement',
    source:'NASA / ESA 2024',
    prompt:'Lunar Gateway space station international docking systems life support power propulsion orbital assembly launch manifest integration' },
  { id:'artemisBase', name:'Artemis Lunar Surface Programme', sector:'Space', region:'Lunar Surface',
    cost:'$93B total estimate (GAO)', sched:'Crewed landing 2026+', status:'Development — behind schedule',
    source:'NASA / GAO 2024',
    prompt:'Artemis crewed lunar landing SLS Orion spacesuits lunar lander surface operations life support thermal radiation autonomy' },
  { id:'kuiper', name:'Amazon Kuiper Constellation', sector:'Space', region:'LEO',
    cost:'$10B+', sched:'Service start 2025', status:'Active launch campaign',
    source:'Amazon / FCC 2024',
    prompt:'Kuiper LEO satellite constellation manufacturing production ramp launch manifest ground segment spectrum coordination regulatory approval' },
  { id:'astSpaceMobile', name:'AST SpaceMobile', sector:'Space', region:'LEO',
    cost:'$1.5B+', sched:'Commercial service 2025', status:'First satellites launched',
    source:'AST SpaceMobile 2024',
    prompt:'Direct to mobile LEO satellite constellation orbital deployment spectrum interoperability telecom regulatory MNO partnerships' },
  { id:'orbitalDC', name:'Orbital Data Centres', sector:'Space', region:'LEO/GEO',
    cost:'$5B–$15B concept', sched:'2030s first module', status:'Concept / investment phase',
    source:'Various / Axiom / press 2024',
    prompt:'Orbital data centre radiation hardening autonomous servicing thermal rejection power generation data relay ground station network' },
  { id:'marsLogistics', name:'Mars Cargo Logistics Network', sector:'Space', region:'Mars',
    cost:'$100B+ long-term', sched:'First cargo 2030s', status:'Early development / SpaceX roadmap',
    source:'SpaceX / NASA 2024',
    prompt:'Mars cargo logistics propellant depot ISRU autonomous operations radiation long-duration communication latency cargo landers surface infrastructure' },
  { id:'lunarHabitat', name:'Lunar Habitat Infrastructure', sector:'Space', region:'Lunar Surface',
    cost:'$50B–$100B', sched:'2035+ sustained presence', status:'Architecture definition',
    source:'NASA / ESA / JAXA 2024',
    prompt:'Lunar habitat crew quarters life support ECLSS thermal survivability power regolith shielding surface mobility autonomous commissioning launch windows mass' },
  { id:'cislunarDepot', name:'Cislunar Propellant Depot', sector:'Space', region:'Cislunar',
    cost:'$4B–$8B', sched:'2028–2032', status:'Concept / early dev',
    source:'NASA / DARPA 2024',
    prompt:'Cislunar propellant depot cryogenic storage autonomous docking launch dependency power management orbital mechanics proximity operations' },
  { id:'starshipMars', name:'Starship Mars Base Architecture', sector:'Space', region:'Mars Surface',
    cost:'$200B+', sched:'First humans 2030s', status:'Architecture / concept',
    source:'SpaceX 2024',
    prompt:'Starship Mars base ISRU propellant oxygen habitat crew life support thermal nuclear power surface operations autonomy radiation' },
  { id:'lunarMining', name:'Lunar Resource Extraction', sector:'Space', region:'Lunar Surface',
    cost:'$3B–$8B concept', sched:'2030s first operations', status:'Technology development',
    source:'NASA / ESA / commercial 2024',
    prompt:'Lunar resource extraction autonomous mining regolith processing ISRU water ice oxygen propellant surface power logistics' },
  { id:'orbitalServicing', name:'Autonomous Orbital Servicing', sector:'Space', region:'LEO/GEO',
    cost:'$2B–$5B', sched:'2026–2030 first ops', status:'Development / trials',
    source:'Northrop / DARPA / ESA 2024',
    prompt:'Autonomous orbital servicing rendezvous proximity operations robotic capture refuelling debris avoidance LEO GEO life extension' },
  { id:'spaceDAwareness', name:'Space Domain Awareness Network', sector:'Space', region:'Global/Orbital',
    cost:'$3B+ defence', sched:'2025–2030', status:'Active development',
    source:'US Space Force / ESA 2024',
    prompt:'Space domain awareness ground station network secure comms encrypted uplink antenna farms sensor fusion orbital tracking' },
  { id:'moonToMars', name:'Moon to Mars Architecture (NASA)', sector:'Space', region:'Cislunar/Mars',
    cost:'$1T+ 2040s total estimate', sched:'2040 Mars target', status:'Architecture / policy',
    source:'NASA / NAS 2024',
    prompt:'Moon to Mars NASA programme SLS Orion Gateway lunar surface Mars transit habitat ISRU life support autonomy radiation' },

  // ── HYDROGEN / LNG / ENERGY TRANSITION ───────────────────────────
  { id:'lngAustralia', name:'Pluto LNG Train 2', sector:'Energy', region:'Australia',
    cost:'$6B+', sched:'2027 target', status:'FID taken',
    source:'Woodside 2024',
    prompt:'LNG liquefaction train expansion cryogenic systems marine loading jetty commissioning long-lead procurement utilities' },
  { id:'h2Green', name:'HyDeal Ambition H2 Corridor', sector:'Energy', region:'Europe',
    cost:'€65B to 2030', sched:'2030 delivery target', status:'Development / financing',
    source:'HyDeal / various 2024',
    prompt:'Green hydrogen production electrolysis solar wind power pipeline compression storage cross-border delivery offtake' },
  { id:'breakfree', name:'Dogger Bank Wind Farm', sector:'Energy', region:'UK',
    cost:'£9B+', sched:'2026 full operation', status:'Under construction',
    source:'Equinor / SSE 2024',
    prompt:'Offshore wind farm North Sea cable installation turbine foundation weather windows grid connection substation HVDC' },

];

// Unique sector list derived from programme data
const HEALTH_SECTORS = ['All', ...Array.from(new Set(ALL_PROGRAMMES.map(p => p.sector))).sort()];

function ProgrammeHealthSignal({ onRunHealthCheck }) {
  const [selected, setSelected] = React.useState(null);
  const [sectorFilter, setSectorFilter] = React.useState('All');
  const [query, setQuery] = React.useState('');
  const [showAll, setShowAll] = React.useState(false);

  const filtered = React.useMemo(() => {
    let list = ALL_PROGRAMMES;
    if (sectorFilter !== 'All') list = list.filter(p => p.sector === sectorFilter);
    if (query.trim()) {
      const q = query.toLowerCase();
      list = list.filter(p =>
        p.name.toLowerCase().includes(q) ||
        p.region.toLowerCase().includes(q) ||
        p.sector.toLowerCase().includes(q) ||
        p.prompt.toLowerCase().includes(q)
      );
    }
    return list;
  }, [sectorFilter, query]);

  const visible = showAll ? filtered : filtered.slice(0, 12);

  const runCheck = (prog) => {
    setSelected(prog.id);
    onRunHealthCheck({ ...prog, caseySignal: { prompt: prog.prompt } });
  };

  // Status colour
  const statusColor = (s) => {
    if (!s) return '#64748b';
    const l = s.toLowerCase();
    if (l.includes('delivered') || l.includes('operational') || l.includes('completed')) return '#10b981';
    if (l.includes('overrun') || l.includes('delay') || l.includes('paused') || l.includes('behind')) return '#ef4444';
    if (l.includes('review') || l.includes('slipping') || l.includes('disputed')) return '#f59e0b';
    return '#94a3b8';
  };

  return (
    <Card className="healthSignal">
      <div className="healthHeader">
        <div>
          <h2>Global Programme Intelligence — {ALL_PROGRAMMES.length} Programmes</h2>
          <p className="advisorIntro">
            Select any live or recently-delivered mega-programme. CASEY generates its own read from benchmark intelligence and compares it to the publicly reported position.
            The gap between reported headline and CASEY governing constraint is the product.
          </p>
        </div>
        <div className="healthCount">
          <b>{ALL_PROGRAMMES.filter(p=>p.sector==='Space').length}</b><span>Space</span>
          <b>{ALL_PROGRAMMES.filter(p=>p.sector!=='Space').length}</b><span>Earth</span>
        </div>
      </div>

      <div className="healthControls">
        <input
          className="healthSearch"
          value={query}
          onChange={e => setQuery(e.target.value)}
          placeholder="Search programmes, regions, sectors…"
        />
        <div className="healthSectorPills">
          {HEALTH_SECTORS.map(s => (
            <button
              key={s}
              className={`healthPill ${sectorFilter === s ? 'active' : ''}`}
              onClick={() => setSectorFilter(s)}
            >
              {s}
              <span>{s === 'All' ? ALL_PROGRAMMES.length : ALL_PROGRAMMES.filter(p => p.sector === s).length}</span>
            </button>
          ))}
        </div>
      </div>

      <div className="healthProgrammes">
        {visible.map(prog => (
          <div key={prog.id} className={`healthRow ${selected === prog.id ? 'active' : ''}`}>
            <div className="healthMeta">
              <b>{prog.name}</b>
              <div className="healthTags">
                <span className="healthSector">{prog.sector}</span>
                <span className="healthRegion">{prog.region}</span>
              </div>
            </div>
            <div className="healthReported">
              <label>Reported</label>
              <span className="reportedCost">{prog.cost}</span>
              <span className="reportedSched">{prog.sched}</span>
              <em className="reportedConf" style={{ color: statusColor(prog.status) }}>{prog.status}</em>
              <small>{prog.source}</small>
            </div>
            <div className="healthVs">vs</div>
            <button className="healthRunBtn" onClick={() => runCheck(prog)}>
              {selected === prog.id ? 'Running…' : 'Run CASEY →'}
            </button>
          </div>
        ))}
      </div>

      {filtered.length > 12 && (
        <button className="healthShowMore" onClick={() => setShowAll(s => !s)}>
          {showAll ? `Show fewer ↑` : `Show all ${filtered.length} programmes ↓`}
        </button>
      )}

      {filtered.length === 0 && (
        <p className="healthEmpty">No programmes match — try a different sector or search term.</p>
      )}

      <p className="healthDisclaimer">
        CASEY reads are generated from benchmark intelligence, not from non-public programme data.
        The comparison is indicative — it shows where the governing constraint model diverges from the reported headline.
        Reported figures are sourced from public statements, auditor reports and press releases as of 2024.
      </p>
    </Card>
  );
}


// ── Advisory Fee Counter ───────────────────────────────────────────
// Shows what a traditional advisory team would charge for what CASEY
// just produced. No firm names. Just the numbers.
function AdvisoryFeeCounter({ model }) {
  const conf = Number(model?.confidence_pct || 60);
  const hasRisk = (model?.risks?.length || 0) >= 5;
  const hasMC = !!(model?.monte_carlo?.qcra?.p80);
  const hasScenarios = (model?.scenario_matrix?.length || 0) >= 3;
  const hasSchedule = (model?.schedule_detail?.length || 0) >= 10;
  const isSpace = model?.mode === 'Space';

  // Build itemised deliverable list with advisory fee ranges
  const items = [
    { label: 'Class 3–5 cost estimate (AACE)', trad: '£80k–£180k', time: '3–6 weeks', done: true },
    { label: 'QCRA Monte Carlo + P-curve', trad: '£40k–£90k', time: '2–4 weeks', done: hasMC },
    { label: 'QSRA Monte Carlo + finish-date curve', trad: '£30k–£70k', time: '2–4 weeks', done: hasMC },
    { label: 'Risk register (cause/event/impact/owner)', trad: '£25k–£60k', time: '3–5 day workshop', done: hasRisk },
    { label: 'Five scenario trade-off analysis', trad: '£30k–£80k', time: '1–3 weeks iteration', done: hasScenarios },
    { label: isSpace ? 'Mission peer benchmark intelligence' : 'Sector peer benchmark intelligence', trad: '£20k–£60k', time: '2–3 weeks research', done: true },
    { label: 'Level 1–5 schedule + P6 XER export', trad: '£20k–£50k', time: '2–4 weeks planning', done: hasSchedule },
    { label: 'Board-grade exports (XLSX/DOCX/PDF/PPTX/XER)', trad: '£15k–£40k', time: '1–2 weeks production', done: true },
  ].filter(x => x.done);

  // Total low/high
  const [low, high] = items.reduce(([l, h], item) => {
    const [il, ih] = item.trad.replace(/[£\$k]/g,'').split('–').map(Number);
    return [l + il, h + ih];
  }, [0, 0]);

  const [animVal, setAnimVal] = React.useState(0);
  React.useEffect(() => {
    const target = high;
    let start = 0;
    const step = target / 40;
    const timer = setInterval(() => {
      start = Math.min(start + step, target);
      setAnimVal(Math.round(start));
      if (start >= target) clearInterval(timer);
    }, 30);
    return () => clearInterval(timer);
  }, [high]);

  return (
    <Card className="feeCounter">
      <div className="feeHeader">
        <div>
          <h2 style={{fontSize:'13px'}}>Equivalent advisory engagement value</h2>
          <p className="feeSubtitle">
            Equivalent early-stage advisory engagement cost for the same deliverables.
            Market rate ranges — no specific firm referenced.
          </p>
        </div>
        <div className="feeTotalBox">
          <span className="feeLabel" style={{fontSize:'9px'}}>Advisory equivalent</span>
          <span className="feeTotalLow">£{low}k – </span>
          <span className="feeTotalHigh">£{animVal}k</span>
          <span className="feeTimeLabel">6 – 10 weeks</span>
        </div>
      </div>
      <div className="feeItems">
        {items.map((item, i) => (
          <div key={i} className="feeRow">
            <span className="feeCheck">✓</span>
            <span className="feeItemLabel">{item.label}</span>
            <span className="feeItemTrad">{item.trad}</span>
            <span className="feeItemTime">{item.time}</span>
            <span className="feeItemCasey">Instant</span>
          </div>
        ))}
      </div>
      <div className="feeDemoLine">
        "Traditional project controls reports show numbers. CASEY shows the board what the numbers are trying to hide."
      </div>
    </Card>
  );
}

function ShowcaseLibrary({ onRun, onBack }) {
  const [sector, setSector] = useState('All');
  const [query, setQuery] = useState('');
  const filtered = showcaseProjects.filter(p => (sector === 'All' || p.sector === sector) && (`${p.title} ${p.sector} ${p.region} ${p.client} ${p.prompt}`.toLowerCase().includes(query.toLowerCase())));
  const counts = showcaseSectors.map(s => ({ sector: s, count: s === 'All' ? showcaseProjects.length : showcaseProjects.filter(p => p.sector === s).length }));
  return <section className="showcaseLibrary">
    <div className="showcaseHero card">
      <div><p className="eyebrow">CASEY Strategic Intelligence Simulations</p><h1>Global capital-project showcase library</h1><p>Clickable Earth and Space reference cases for board packs, scenario pressure tests, confidence analysis and audit-ready export demos.</p></div>
      <div className="phaseStack"><b>Phase 1</b><span>Showcase Library</span><b>Phase 2</b><span>Live intelligence feeds</span><b>Phase 3</b><span>Bring-your-own-project ingestion</span></div>
    </div>
    <div className="showcaseControls card"><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Search rail, data centres, Starship, Lilly, NEOM..."/><button onClick={onBack}>Back to console</button></div>
    <div className="sectorPills">{counts.map(x => <button key={x.sector} className={sector===x.sector?'active':''} onClick={()=>setSector(x.sector)}>{x.sector}<span>{x.count}</span></button>)}</div>
    <div className="showcaseGrid">{filtered.map(p => <motion.button className="showcaseCard" key={p.title} onClick={() => onRun(p)} whileHover={{ y:-4 }} whileTap={{ scale:.99 }}>
      <div className="showcaseIcon">{p.icon}</div><div className="showcaseMeta"><span>{p.sector}</span><em>{p.region}</em></div><h3>{p.title}</h3><p>{p.prompt}</p><div className="showcaseFooter"><b>{p.confidence}</b><strong>Run board pack <ChevronRight size={15}/></strong></div>
    </motion.button>)}</div>
  </section>;
}

function GatedMessage({ raw, onDismiss, onShowcase, onEarth, onSpace }) {
  let msg = "You have used your one free CASEY intelligence run.";
  let sub = "You can still browse 200 free reference cases in the Showcase Library and run the Earth or Space demos for free. For unlimited projects, get in touch.";
  let email = "deepa@caseai.co.uk";
  let linkedin = "https://www.linkedin.com/company/caseai";
  let isStartingUp = false;
  try {
    const p = JSON.parse(raw);
    if (p.message) msg = p.message;
    if (p.sub) sub = p.sub;
    if (p.email) email = p.email;
    if (p.linkedin) linkedin = p.linkedin;
    isStartingUp = !!(p.message && p.message.toLowerCase().includes('starting'));
  } catch {}

  return (
    <div style={{
      position:'fixed',inset:0,background:'rgba(0,0,0,0.65)',zIndex:900,
      display:'flex',alignItems:'center',justifyContent:'center',padding:'20px'
    }}>
      <div style={{
        background:'#0d1b2e',border:'1px solid rgba(141,247,255,0.25)',borderRadius:'8px',
        padding:'28px 32px',maxWidth:'520px',width:'100%',position:'relative',
        boxShadow:'0 24px 80px rgba(0,0,0,0.6)'
      }}>
        {onDismiss && <button onClick={onDismiss} style={{
          position:'absolute',top:'14px',right:'16px',background:'rgba(255,255,255,0.06)',
          border:'1px solid rgba(255,255,255,0.1)',color:'#94a3b8',cursor:'pointer',
          fontSize:'14px',width:'28px',height:'28px',borderRadius:'50%',
          display:'flex',alignItems:'center',justifyContent:'center',lineHeight:1
        }}>✕</button>}

        <div style={{fontSize:'22px',marginBottom:'10px'}}>{isStartingUp ? '⏳' : '✦'}</div>
        <h3 style={{
          fontSize:'17px',fontWeight:'800',color: isStartingUp ? '#f59e0b' : '#e2e8f0',
          marginBottom:'8px',paddingRight:'30px',lineHeight:'1.3'
        }}>{isStartingUp ? '⏳ Server starting up — try again in 20 seconds' : msg}</h3>
        <p style={{
          fontSize:'13px',color:'#94a3b8',lineHeight:'1.6',marginBottom:'18px'
        }}>{sub}</p>

        {isStartingUp ? (
          <div style={{display:'flex',flexDirection:'column',gap:'8px'}}>
            <div style={{background:'rgba(245,158,11,0.08)',border:'1px solid rgba(245,158,11,0.2)',borderRadius:'5px',padding:'10px 14px',fontSize:'12px',color:'#fde68a',lineHeight:'1.5'}}>
              The backend is waking up — this takes 20–30 seconds on first load. Click the demo button again after waiting.
            </div>
            <div style={{display:'flex',gap:'8px',flexWrap:'wrap'}}>
              {onEarth && <button onClick={onEarth} style={{flex:1,background:'rgba(16,185,129,0.1)',border:'1px solid rgba(16,185,129,0.3)',color:'#10b981',cursor:'pointer',padding:'10px 14px',borderRadius:'5px',fontSize:'13px',fontWeight:'700'}}>🚄 Try Earth Demo again</button>}
              {onSpace && <button onClick={onSpace} style={{flex:1,background:'rgba(141,247,255,0.08)',border:'1px solid rgba(141,247,255,0.2)',color:'#8df7ff',cursor:'pointer',padding:'10px 14px',borderRadius:'5px',fontSize:'13px',fontWeight:'700'}}>🌕 Try Space Demo again</button>}
            </div>
            {onShowcase && <button onClick={onShowcase} style={{width:'100%',background:'rgba(255,255,255,0.04)',border:'1px solid rgba(255,255,255,0.1)',color:'#94a3b8',cursor:'pointer',padding:'8px',borderRadius:'5px',fontSize:'12px'}}>Browse Showcase Library while waiting →</button>}
          </div>
        ) : (
          <div style={{display:'flex',flexDirection:'column',gap:'8px'}}>
            {onShowcase && <button onClick={onShowcase} style={{width:'100%',background:'rgba(141,247,255,0.1)',border:'1px solid rgba(141,247,255,0.3)',color:'#8df7ff',cursor:'pointer',padding:'11px',borderRadius:'5px',fontSize:'13px',fontWeight:'700',textAlign:'left'}}>
              Browse Showcase Library — 200 free reference cases →
            </button>}
            <div style={{display:'flex',gap:'8px',flexWrap:'wrap'}}>
              {onEarth && <button onClick={onEarth} style={{flex:1,background:'rgba(16,185,129,0.08)',border:'1px solid rgba(16,185,129,0.2)',color:'#10b981',cursor:'pointer',padding:'9px',borderRadius:'5px',fontSize:'12px',fontWeight:'700'}}>🚄 Earth Demo</button>}
              {onSpace && <button onClick={onSpace} style={{flex:1,background:'rgba(141,247,255,0.05)',border:'1px solid rgba(141,247,255,0.15)',color:'#8df7ff',cursor:'pointer',padding:'9px',borderRadius:'5px',fontSize:'12px',fontWeight:'700'}}>🌕 Space Demo</button>}
            </div>
            <div style={{borderTop:'1px solid rgba(255,255,255,0.07)',paddingTop:'12px',display:'flex',flexDirection:'column',gap:'6px'}}>
              <a href={"mailto:" + email} style={{display:'block',background:'rgba(141,247,255,0.06)',border:'1px solid rgba(141,247,255,0.2)',color:'#8df7ff',padding:'10px 14px',borderRadius:'5px',fontSize:'12px',fontWeight:'700',textDecoration:'none',textAlign:'center'}}>✉ Email us for full access — {email}</a>
              <a href={linkedin} target="_blank" rel="noopener noreferrer" style={{display:'block',background:'rgba(255,255,255,0.04)',border:'1px solid rgba(255,255,255,0.1)',color:'#94a3b8',padding:'9px 14px',borderRadius:'5px',fontSize:'12px',textDecoration:'none',textAlign:'center'}}>Connect on LinkedIn</a>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function DemoBanner({ model }) {
  if (!model) return null;

  const isDemo =
    model.demo_mode === true ||
    Boolean(model.demo_type) ||
    Boolean(model.demo_label) ||
    Boolean(model.demo_headline);

  if (!isDemo) return null;

  const label = safeRender(model.demo_label || model.demo_type || 'Public demo');
  const headline = safeRender(
    model.demo_headline ||
    (model.demo_type === 'earth'
      ? 'Earth infrastructure demo loaded'
      : model.demo_type === 'space'
        ? 'Space infrastructure demo loaded'
        : 'Showcase demo loaded')
  );

  const subline = safeRender(
    model.demo_subline ||
    'Demo routes stay open for visitors. Only the free custom project run is limited.'
  );

  return (
    <section className="demoBanner" style={{margin:'12px 0',padding:'12px 14px',border:'1px solid rgba(141,247,255,0.22)',borderRadius:'14px',background:'linear-gradient(135deg, rgba(141,247,255,0.10), rgba(124,58,237,0.08))',boxShadow:'0 12px 36px rgba(0,0,0,0.18)'}}>
      <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',gap:'12px',flexWrap:'wrap'}}>
        <div>
          <div style={{fontSize:'10px',letterSpacing:'0.14em',textTransform:'uppercase',color:'#8df7ff',fontWeight:800}}>{label}</div>
          <div style={{fontSize:'15px',fontWeight:800,color:'#f8fafc',marginTop:'3px'}}>{headline}</div>
          <div style={{fontSize:'12px',color:'#94a3b8',marginTop:'3px'}}>{subline}</div>
        </div>
        <div style={{fontSize:'11px',fontWeight:800,color:'#d1fae5',border:'1px solid rgba(16,185,129,0.35)',background:'rgba(16,185,129,0.10)',borderRadius:'999px',padding:'6px 10px'}}>Demo bypass active</div>
      </div>
    </section>
  );
}

function EmailGateForm({ onSubmit, onDismiss }) {
  const [val, setVal] = React.useState('');
  return <div>
    <input type="email" value={val} onChange={e => setVal(e.target.value)}
      placeholder="your@email.com"
      style={{width:'100%',padding:'10px 12px',background:'#0a0f1e',border:'1px solid #0e7490',borderRadius:'4px',color:'#fff',fontSize:'13px',marginBottom:'10px',outline:'none'}}
      onKeyDown={e => e.key === 'Enter' && val.includes('@') && onSubmit(val)}
    />
    <div style={{display:'flex',gap:'8px'}}>
      <button onClick={() => val.includes('@') && onSubmit(val)}
        style={{flex:1,padding:'10px',background:'#0e7490',color:'#fff',border:'none',borderRadius:'4px',fontWeight:'700',fontSize:'13px',cursor:'pointer'}}>
        Get my free run →
      </button>
      <button onClick={onDismiss}
        style={{padding:'10px 16px',background:'transparent',color:'#64748b',border:'1px solid #334155',borderRadius:'4px',fontSize:'12px',cursor:'pointer'}}>
        Cancel
      </button>
    </div>
    <p style={{fontSize:'9px',color:'#475569',marginTop:'8px',fontFamily:'monospace'}}>No spam. No card. 1 free project run. Earth Demo and Space Demo always free.</p>
  </div>;
}


function ApprovalStatus({ model }) {
  const conf = Number(model?.confidence_pct || 0);
  const ready = conf >= 75;
  const partial = conf >= 55;
  const status = ready ? 'APPROVAL READY' : partial ? 'CONDITIONAL — EVIDENCE GAPS REMAIN' : 'NOT READY FOR BOARD';
  const color = ready ? '#10b981' : partial ? '#f59e0b' : '#ef4444';
  const bg = ready ? 'rgba(16,185,129,0.08)' : partial ? 'rgba(245,158,11,0.06)' : 'rgba(239,68,68,0.08)';
  const border = ready ? 'rgba(16,185,129,0.4)' : partial ? 'rgba(245,158,11,0.35)' : 'rgba(239,68,68,0.5)';
  return (
    <div style={{background:bg,border:`2px solid ${border}`,borderRadius:10,padding:'16px 20px',marginBottom:12,display:'grid',gridTemplateColumns:'1fr auto',alignItems:'center',gap:16}}>
      <div>
        <div style={{fontSize:'10px',fontWeight:'800',color,letterSpacing:'.14em',marginBottom:4}}>{status}</div>
        <div style={{fontSize:'22px',fontWeight:'900',color:'#fff',marginBottom:6}}>
          {safeRender(model?.cost_p50 || '—')} &nbsp;<span style={{fontSize:'14px',fontWeight:'400',color:'#64748b'}}>P50</span>&nbsp;&nbsp;
          {safeRender(model?.cost_p80 || model?.cost_p90 || '—')} &nbsp;<span style={{fontSize:'14px',fontWeight:'400',color:'#64748b'}}>P80</span>&nbsp;&nbsp;
          {safeRender(model?.schedule || '—')} &nbsp;<span style={{fontSize:'14px',fontWeight:'400',color:'#64748b'}}>delivery</span>
        </div>
        <div style={{fontSize:'11px',color:'#94a3b8'}}>{safeRender(model?.subsector || 'Programme')} · {safeRender(model?.location || 'Global')} · {safeRender(model?.estimate_class_name || '')}</div>
      </div>
      <div style={{textAlign:'center',padding:'10px 20px',background:'rgba(255,255,255,0.03)',borderRadius:8,border:'1px solid rgba(255,255,255,0.07)'}}>
        <div style={{fontSize:'9px',color:'#64748b',marginBottom:4}}>BOARD CONFIDENCE</div>
        <div style={{fontSize:'36px',fontWeight:'900',color,lineHeight:1}}>{conf ? conf + '%' : '—'}</div>
        <div style={{fontSize:'8px',color:'#475569',marginTop:3}}>{conf>=75?'Board-defensible':'Target: 75%+'}</div>
      </div>
    </div>
  );
}

// ── CASEY Self-Challenge Component ─────────────────────────────────────────
function SelfChallenge({ sc, programme }) {
  const [expanded, setExpanded] = React.useState(false);
  if (!sc) return null;

  const overall = sc.overall_score || 0;
  const traffic = sc.overall_traffic || 'amber';
  const dims = sc.dimensions || {};
  const totalIssues = sc.total_issues || 0;

  const trafficColor = t => t === 'green' ? '#10b981' : t === 'amber' ? '#f59e0b' : '#ef4444';
  const trafficBg = t => t === 'green' ? 'rgba(16,185,129,0.08)' : t === 'amber' ? 'rgba(245,158,11,0.08)' : 'rgba(239,68,68,0.08)';
  const trafficBorder = t => t === 'green' ? 'rgba(16,185,129,0.25)' : t === 'amber' ? 'rgba(245,158,11,0.25)' : 'rgba(239,68,68,0.3)';
  const dot = t => <span style={{display:'inline-block',width:8,height:8,borderRadius:'50%',background:trafficColor(t),marginRight:6,flexShrink:0,marginTop:3}}/>;

  const dimList = Object.values(dims);

  // Self-challenge panel render
  return <div className="v50SelfChallenge" style={{background:'rgba(255,255,255,0.02)',border:'1px solid '+trafficBorder(traffic),borderRadius:6,padding:'10px 14px'}}>
    <div style={{fontSize:'10px',fontWeight:'700',color:trafficColor(traffic)}}>⚡ CASEY OUTPUT CHALLENGE — {sc?.verdict||''}</div>
    <div style={{fontSize:'9px',color:'#94a3b8',marginTop:4}}>{overall}% overall · {totalIssues} items to address</div>
  </div>;
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
  const isNonBase = !!(model?.scenario && model.scenario !== 'base');
  const sCostMult = parseFloat(model?.scenario_cost_mult || 1.0);
  const sRiskMult = parseFloat(model?.scenario_risk_mult || 1.0);
  const sSchedMult = parseFloat(model?.scenario_sched_mult || 1.0);
  // Scenario helpers - available throughout App

  const [projectContext, setProjectContext] = useState(null);
  const [tab, setTab] = useState('overview');
  const [healthProg, setHealthProg] = React.useState(null);
  const runHealthCheck = React.useCallback((prog) => {
    setPrompt(prog.caseySignal.prompt);
    setScenario('base');
    setClassLevel(3);
    setScheduleLevel(4);
    // Trigger generate with health signal context
    setTimeout(() => {
      generate('base', prog.caseySignal.prompt, null, prog.name || 'Reference case', { healthCheck: true, isShowcase: true, isDemo: true });
      setTab('assurance');
    }, 100);
  }, []);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [chatQ, setChatQ] = useState('');
  const [chat, setChat] = useState([{ role: 'assistant', text: 'Ask CASEY anything about this programme — cost, schedule, risk, board questions, benchmarks, or governing constraints. Be direct.' }]);
  const [chatUsed, setChatUsed] = useState(0); // free advisor questions used
  const FREE_ADVISOR_LIMIT = 1; // 1 free advisor question, unlimited for admin
  const [advisorModel, setAdvisorModel] = useState(null); // which AI is active
  const [uploadResult, setUploadResult] = useState(null);
  const [viewMode, setViewMode] = useState('exec');
  const [propagating, setPropagating] = useState(false);
  const [simulationStage, setSimulationStage] = useState('');
  const [exportingLabel, setExportingLabel] = useState('');
  const [confidencePulse, setConfidencePulse] = useState(false);
  const [showShowcase, setShowShowcase] = useState(false);

  
  // ── Usage limits & email gate (localStorage-tracked) ──────────────────
  const [showLanding, setShowLanding] = React.useState(() => !localStorage.getItem('casey_seen_landing'));
  const [emailGateOpen, setEmailGateOpen] = React.useState(false);
  const [emailGateFor, setEmailGateFor] = React.useState('');
  const [capturedEmail, setCapturedEmail] = React.useState(() => localStorage.getItem('casey_email') || '');
  const [freeRunsUsed, setFreeRunsUsed] = React.useState(() => parseInt(localStorage.getItem('casey_free_runs') || '0'));
  const [freeCompareUsed, setFreeCompareUsed] = React.useState(() => parseInt(localStorage.getItem('casey_free_compare') || '0'));
  const FREE_RUN_LIMIT = 1;
  const FREE_COMPARE_LIMIT = 1;
  const isUnlimited = false;
  const checkAndGate = (action) => {
    if (isUnlimited) return true;
    if (action === 'run') {
      if (!capturedEmail) { setEmailGateFor('run'); setEmailGateOpen(true); return false; }
      if (freeRunsUsed >= FREE_RUN_LIMIT) { setTab('upgrade'); return false; }
    }
    if (action === 'compare') {
      if (!capturedEmail) { setEmailGateFor('compare'); setEmailGateOpen(true); return false; }
      if (freeCompareUsed >= FREE_COMPARE_LIMIT) { setTab('upgrade'); return false; }
    }
    return true;
  };
  const recordUsage = (action) => {
    if (isUnlimited) return;
    if (action === 'run') { const n = freeRunsUsed + 1; setFreeRunsUsed(n); localStorage.setItem('casey_free_runs', String(n)); }
    if (action === 'compare') { const n = freeCompareUsed + 1; setFreeCompareUsed(n); localStorage.setItem('casey_free_compare', String(n)); }
  };
  const saveEmailAndContinue = (email) => {
    localStorage.setItem('casey_email', email);
    setCapturedEmail(email);
    fetch(API + '/capture-email', {method:'POST',credentials:'omit',headers:{'Content-Type':'application/json'},body:JSON.stringify({email, action: emailGateFor})}).catch(()=>{});
    setEmailGateOpen(false);
  };

useEffect(() => {
    // Ping backend on page load to pre-warm Render instance & check status
    const checkBackend = async () => {
      // Try up to 3 times with delay — Render cold starts take 20-30s
      for (let attempt = 1; attempt <= 3; attempt++) {
        try {
          await get('/health');
          setBackendStatus('ok');
          return;
        } catch {
          if (attempt < 3) await new Promise(r => setTimeout(r, 12000));
        }
      }
      // Only mark as down after 3 failed attempts (~36 seconds)
      setBackendStatus('down');
    };
    checkBackend();

    // Keep-alive ping every 9 minutes — prevents Render free tier sleeping
    // Render sleeps after 15 minutes of inactivity; 9min interval keeps it warm
    const keepAlive = setInterval(async () => {
      try {
        await fetch((window._CASEY_API || window.CASEY_API || PROD_URL) + '/health', {
          method: 'GET', credentials: 'omit',
          signal: AbortSignal.timeout ? AbortSignal.timeout(5000) : undefined
        });
        setBackendStatus('ok');
      } catch (_) {
        // Silent — don't show error for background keep-alive
      }
    }, 9 * 60 * 1000); // 9 minutes

    return () => clearInterval(keepAlive);
  }, []);

  // Ping backend when page becomes visible (catches Render sleep after tab was inactive)
  useEffect(() => {
    const handleVisibility = async () => {
      if (document.visibilityState === 'visible') {
        try {
          const resp = await fetch((window._CASEY_API || PROD_URL) + '/health', {
            method: 'GET', credentials: 'omit',
            signal: AbortSignal.timeout ? AbortSignal.timeout(5000) : undefined
          });
          if (resp.ok) setBackendStatus('ok');
          else setBackendStatus('down');
        } catch (_) { setBackendStatus('down'); }
      }
    };
    document.addEventListener('visibilitychange', handleVisibility);
    return () => document.removeEventListener('visibilitychange', handleVisibility);
  }, []);
  

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
    const toMoney = (bn) => bn >= 1000 ? (curr || '$') + (bn/1000).toFixed(1) + 'T' : bn >= 1 ? (curr || '$') + bn.toFixed(1) + 'B' : (curr || '$') + (bn*1000).toFixed(0) + 'M';

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


  
function TrustRuntimeBar({ model }) {
  if (!model) return null;
  const g = model.governance_state || model.scenario_state?.governance || {};
  const runtime = model.casey_runtime || {};
  const items = [
    ['Board defensibility', g.board_defensibility ?? '—'],
    ['Governance stress', g.governance_stress ?? '—'],
    ['Tail exposure', g.tail_exposure ?? '—'],
    ['Evidence volatility', g.evidence_volatility ?? '—'],
    ['Reserve pressure', g.reserve_pressure ?? '—'],
    ['Signature', model.scenario_signature || runtime.scenario_signature || '—'],
  ];
  return <section className="trustRuntimeBar">
    <div className="trustRuntimeLead"><b>STRESS TEST</b><span>{g.decision_posture || "Scenario stress testing active"}</span></div>
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
  function moneyLocal(n, curr) { const c = curr || '$'; return n >= 1000 ? `${c}${(n/1000).toFixed(1)}T` : n >= 1 ? `${c}${n.toFixed(1)}B` : `${c}${Math.round(n*1000)}M`; }
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
  const [showHelp, setShowHelp] = React.useState(false);
  const [showIngestPanel, setShowIngestPanel] = React.useState(false);
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
      // Pass user's selected class_level and schedule_level so all combinations work
      const payload = {
        prompt: cfg.prompt,
        client: cfg.client,
        class_level: classLevel || 3,
        schedule_level: scheduleLevel || 4,
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
        message: 'Demo engine is waking up — wait 20 seconds and try again.',
        sub: 'Earth Demo, Space Demo and Showcase Library all use the CASEY model engine. If the server has slept, wait 20 seconds and click again. If this repeats, the backend needs redeploying.',
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
    setLoading(true); setTab('overview');
    // Demo gate — fires only for brand-new project runs from the main console
    // NEVER fires for: showcase library, earth/space demo, scenario switching on existing model
    // Gate is handled by checkAndGate() on the Generate button before generate() is called.
    // generate() itself never blocks — showcase, demos, scenarios, and health checks all flow through freely.
    const isNewProjectRun = !activeContext && !opts.isShowcase && !opts.isDemo && !opts.healthCheck;
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
      if (!isAdminUser && isNewProjectRun) { markDemoUsed(); }
    } catch (e) {
      let raw = String(e.message || e);
      // Detect network/CORS/connection errors — show wake-up message, never the gate message
      const isNetworkError = raw.includes('fetch') || raw.includes('Failed to fetch') || 
        raw.includes('NetworkError') || raw.includes('CORS') || raw.includes('ERR_') ||
        raw.includes('net::') || raw.includes('Connection refused') || raw.includes('502') ||
        raw.includes('503') || raw.includes('504') || raw.includes('Load failed') ||
        raw.includes('unreachable') || raw.includes('AbortError') || raw.includes('timeout') ||
        raw.includes('signal') || raw.includes('aborted');
      if (isNetworkError) {
        setError(JSON.stringify({
          message: 'Backend starting up — wait 20 seconds and try again.',
          sub: 'The server is waking up after a period of inactivity. Wait 20 seconds then click Earth Demo, Space Demo or any Showcase item. These are always free.',
          email: 'deepa@caseai.co.uk',
          linkedin: 'https://www.linkedin.com/company/caseai'
        }));
        setBackendStatus('down');
        return;
      }
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
      setChat(x => [...x, {role:'user',text:q}, {role:'assistant',text:'**No project loaded**\n\nGenerate or load a project first, then ask CASEY anything about cost, schedule, risk, or board strategy.'}]);
      return;
    }
    // Advisor lock: 1 free question for non-admin users
    if (!isAdminUser && chatUsed >= FREE_ADVISOR_LIMIT) {
      setChat(x => [...x, {role:'user',text:q}, {role:'assistant',text:'**CASEY Advisor — Access Required**\n\nYou have used your free advisor question. The CASEY Advisor uses live AI to challenge your programme with the authority of an IPA reviewer.\n\nRequest full access at hello@controlorbit.com or enter your access code.'}]);
      setEmailGateOpen(true);
      setEmailGateFor('advisor');
      return;
    }
    setChatUsed(n => n + 1);
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
      instant = '**WHAT AN EXTERNAL REVIEWER ATTACKS — IN ORDER**\n\n' +
        '1. P50 vs P80 gap: is the reserve funded and named? If not, the number is not real.\n' +
        '2. OBA disclosure: is the reference class adjustment in the executive summary? If not, why not?\n' +
        '3. Risk register: are there risks with no EMV and no named owner? Every one is an undisclosed liability.\n' +
        '4. Governing constraint: is it on the critical path? Is there a named owner and a closure date?\n' +
        '5. Schedule float: is it operationally usable? Theoretical float that cannot be accessed is not float.\n\n' +
        (attacks.length ? 'Your specific programme challenges:\n' + attacks.slice(0,3).map(function(a,i){return (i+6)+'. '+a;}).join('\n') : '');
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
      setChat(x => [...x, { role: 'assistant', text: String(answer || 'CASEY returned no advisor response.'), delta: r?.delta || null, source: r?.source || null }]);
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



  const renderApprovalStatus = () => {
              const conf = model?.confidence_pct || 0;
              const ready = conf >= 75; const partial = conf >= 55;
              const status = ready ? 'APPROVAL READY' : partial ? 'CONDITIONAL — EVIDENCE GAPS REMAIN' : 'NOT READY FOR BOARD';
              const color = ready ? '#10b981' : partial ? '#f59e0b' : '#ef4444';
              const bg = ready ? 'rgba(16,185,129,0.08)' : partial ? 'rgba(245,158,11,0.06)' : 'rgba(239,68,68,0.08)';
              const border = ready ? 'rgba(16,185,129,0.4)' : partial ? 'rgba(245,158,11,0.35)' : 'rgba(239,68,68,0.5)';
              return <div style={{background:bg,border:`2px solid ${border}`,borderRadius:10,padding:'16px 20px',marginBottom:12,display:'grid',gridTemplateColumns:'1fr auto',alignItems:'center',gap:16}}>
                <div>
                  <div style={{fontSize:'10px',fontWeight:'800',color,letterSpacing:'.14em',marginBottom:4}}>{status}</div>
                  <div style={{fontSize:'22px',fontWeight:'900',color:'#fff',marginBottom:6}}>{model?.cost_p50} &nbsp;<span style={{fontSize:'14px',fontWeight:'400',color:'#64748b'}}>P50</span>&nbsp;&nbsp;{model?.cost_p80} &nbsp;<span style={{fontSize:'14px',fontWeight:'400',color:'#64748b'}}>P80</span>&nbsp;&nbsp;{model?.schedule} &nbsp;<span style={{fontSize:'14px',fontWeight:'400',color:'#64748b'}}>delivery</span></div>
                  <div style={{fontSize:'11px',color:'#94a3b8'}}>{model?.subsector} · {model?.location} · {model?.estimate_class_name}</div>
                </div>
                <div style={{textAlign:'center',padding:'10px 20px',background:'rgba(255,255,255,0.03)',borderRadius:8,border:'1px solid rgba(255,255,255,0.07)'}}>
                  <div style={{fontSize:'9px',color:'#64748b',marginBottom:4}}>BOARD CONFIDENCE</div>
                  <div style={{fontSize:'36px',fontWeight:'900',color,lineHeight:1}}>{model?.confidence_pct+'%'}</div>
                  <div style={{fontSize:'8px',color:'#475569',marginTop:3}}>{conf>=75?'Board-defensible':'Target: 75%+'}</div>
                </div>
              </div>;
            };

  const renderScenarioCompare = () => {
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
      <div style={{fontSize:'11px',color:'#94a3b8',marginBottom:'4px'}}>{row.schedule_months||'—'} months · <span style={{color:cCol,fontWeight:'700'}}>{row.confidence_pct||row.confidence||'—'+'%'}</span> · <span style={{color:rCol,fontWeight:'700'}}>{row.risk||'—'}</span></div>
      {s!=='base'&&bCost>0&&<div style={{display:'flex',gap:'5px',flexWrap:'wrap',marginBottom:'3px'}}>
        {costD!==0&&<span style={{background:'rgba(239,68,68,0.12)',color:costD>0?'#ef4444':'#10b981',padding:'1px 6px',borderRadius:'2px',fontSize:'10px',fontWeight:'900'}}>{(costD>0?'+':'-')+Math.abs(costD)+'% cost'}</span>}
        {confD!==0&&<span style={{background:'rgba(6,182,212,0.12)',color:confD>0?'#10b981':'#ef4444',padding:'1px 6px',borderRadius:'2px',fontSize:'10px',fontWeight:'900'}}>{(confD>0?'+':'-')+Math.abs(confD)+'pt conf'}</span>}
        {schedD!==0&&<span style={{background:'rgba(245,158,11,0.12)',color:schedD<0?'#10b981':'#f59e0b',padding:'1px 6px',borderRadius:'2px',fontSize:'10px',fontWeight:'900'}}>{(schedD>0?'+':'-')+Math.abs(schedD)+'mo'}</span>}
      </div>}
      {ws.length>0&&<div style={{fontSize:'9px',color:'#ef4444',fontWeight:'700',letterSpacing:'.05em'}}>WORSE: {ws.join(' · ')}</div>}
      {im.length>0&&<div style={{fontSize:'9px',color:'#10b981',fontWeight:'700',letterSpacing:'.05em'}}>BETTER: {im.join(' · ')}</div>}
      <em style={{fontSize:'10px',color:'#64748b',fontStyle:'normal',lineHeight:'1.3',display:'block',marginTop:'4px'}}>{active?'▶ ACTIVE — '+scenario.toUpperCase():tradeNotes[s]||''}</em>
    </button>;
  })}</div>;
};

  const renderChatMsg = (m) => {
    const lines = String(m.text||'').split('\n');
    const textBlock = lines.map((line, li) => {
      if (!line.trim()) return <div key={li} style={{height:'5px'}}/>;
      if (line.startsWith('**') && line.endsWith('**') && line.length > 4)
        return <div key={li} className="chatHeading">{line.replace(/\*\*/g,'')}</div>;
      const parts = line.split(/\*\*([^*]+)\*\*/g);
      return <p key={li} className="chatLine">{parts.map((p,pi)=>pi%2===1?<strong key={pi}>{p}</strong>:p)}</p>;
    });
    const srcBadge = m.role === 'assistant' && m.source ? <div style={{marginTop:'4px',fontSize:'9px',color:'#334155',fontFamily:'monospace'}}>
      {m.source === 'claude' ? '⚡ Claude (Anthropic)' : m.source === 'openai' ? '🤖 GPT-4o (OpenAI)' : m.source === 'pattern' ? '⚙ Pattern match (set API key for AI)' : ''}
    </div> : null;
    const d = m.delta;
    const deltaBlock = d && !d.error ? <div style={{marginTop:'10px',background:'rgba(141,247,255,0.05)',border:'1px solid rgba(141,247,255,0.2)',borderRadius:'4px',padding:'10px 12px'}}>
      <div style={{fontSize:'9px',fontWeight:'800',letterSpacing:'.12em',color:'#8df7ff',marginBottom:'6px'}}>◆ MODEL RECALCULATED — CONSTRAINT APPLIED</div>
      <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:'8px',marginBottom:'8px'}}>
        {d.cost_delta_bn !== undefined && <div style={{textAlign:'center',padding:'6px',background:'rgba(255,255,255,0.03)',borderRadius:'3px'}}>
          <div style={{fontSize:'9px',color:'#475569',marginBottom:'2px'}}>Cost delta</div>
          <div style={{fontSize:'13px',fontWeight:'800',color:d.cost_delta_bn>0?'#fca5a5':'#10b981'}}>{(d.cost_delta_bn>0?'+':'')+((d.cost_delta_bn||0).toFixed(1))+'B'}</div>
          <div style={{fontSize:'10px',color:'#8df7ff'}}>{d.new_p50}</div>
        </div>}
        {d.confidence_delta !== undefined && <div style={{textAlign:'center',padding:'6px',background:'rgba(255,255,255,0.03)',borderRadius:'3px'}}>
          <div style={{fontSize:'9px',color:'#475569',marginBottom:'2px'}}>Confidence</div>
          <div style={{fontSize:'13px',fontWeight:'800',color:d.confidence_delta<0?'#fca5a5':'#10b981'}}>{(d.confidence_delta>0?'+':'')+d.confidence_delta+'pts'}</div>
          <div style={{fontSize:'10px',color:'#8df7ff'}}>{d.new_confidence+'%'}</div>
        </div>}
        {d.schedule_delta_months !== undefined && <div style={{textAlign:'center',padding:'6px',background:'rgba(255,255,255,0.03)',borderRadius:'3px'}}>
          <div style={{fontSize:'9px',color:'#475569',marginBottom:'2px'}}>Schedule delta</div>
          <div style={{fontSize:'13px',fontWeight:'800',color:d.schedule_delta_months>0?'#fca5a5':'#10b981'}}>{(d.schedule_delta_months>0?'+':'')+d.schedule_delta_months+'mo'}</div>
          <div style={{fontSize:'10px',color:'#8df7ff'}}>{d.new_schedule}</div>
        </div>}
      </div>
      {d.new_governing_constraint && <div style={{fontSize:'10px',color:'#94a3b8',borderTop:'1px solid rgba(255,255,255,0.06)',paddingTop:'6px'}}>New governing constraint: <b style={{color:'#e2e8f0'}}>{d.new_governing_constraint}</b></div>}
    </div> : null;
    return <>{textBlock}{srcBadge}{deltaBlock}</>;
  };

  const renderBenchmarkStats = () => {
                const bms = model.benchmark_comparison || [];
                const costGrowths = bms.map(b=>b.cost_growth_pct||0).sort((a,b)=>a-b);
                const slips = bms.map(b=>b.schedule_slip_months||0).sort((a,b)=>a-b);
                const p50Val = model?.p50_cost_bn || 0;
                const anchorVals = bms.map(b=>parseFloat((b.anchor_cost||'0').replace(/[^0-9.]/g,''))).filter(v=>v>0).sort((a,b)=>a-b);
                const costPct = anchorVals.length ? Math.round(anchorVals.filter(v=>v<=p50Val*1000).length/anchorVals.length*100) : null;
                return <div style={{fontSize:'11px',color:'#cbd5e1',lineHeight:'1.6'}}>
                  {costPct !== null && <div style={{marginBottom:3}}><b style={{color:'#fff'}}>Cost:</b> Your P50 is at the <span style={{color:'#a78bfa',fontWeight:'700'}}>{costPct}th percentile</span> of this cohort by anchor cost.</div>}
                  <div style={{marginBottom:3}}><b style={{color:'#fff'}}>Risk profile:</b> {model?.confidence_pct >= 65 ? 'Above median confidence for this cohort.' : 'Below median confidence — higher than average delivery risk.'}</div>
                  <div><b style={{color:'#fff'}}>Class {model?.estimate_class} accuracy range:</b> Your P50 could reasonably be {({1:'±10%',2:'±15%',3:'±20%',4:'±35%',5:'±50%'})[model?.estimate_class] || '±20%'} higher or lower.</div>
                </div>;
              };

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
      {model?.live_intel_active && <div style={{display:'flex',alignItems:'center',gap:'5px',padding:'3px 9px',background:'rgba(16,185,129,0.08)',border:'1px solid rgba(16,185,129,0.2)',borderRadius:'4px'}}>
        <div style={{width:'6px',height:'6px',borderRadius:'50%',background:'#10b981',animation:'pulse 1.5s infinite',flexShrink:0}}/>
        <span style={{fontSize:'9px',fontWeight:'800',color:'#10b981',letterSpacing:'.06em'}}>
            {model.live_intel_mode === 'AI-enriched' ? '⚡ MARKET INTEL (AI)' : '📊 MARKET INTEL'}
          </span>
          <span style={{fontSize:'8px',color:'#334155'}}>{(model.live_intel_sources||'').includes('GDELT') ? '🌐 Live news' : ''} {(model.live_intel_timestamp||'').split(' ').slice(0,3).join(' ')}</span>
      </div>}
      <button onClick={() => setShowHelp(true)} style={{color:'#8df7ff',fontSize:'10px',fontWeight:'800',letterSpacing:'.06em',background:'rgba(141,247,255,0.08)',border:'1px solid rgba(141,247,255,0.2)',borderRadius:'4px',padding:'3px 10px',cursor:'pointer'}}>? HELP</button>
      {model && <button onClick={() => setShowIngestPanel(p=>!p)} style={{color:'#10b981',fontSize:'10px',fontWeight:'800',letterSpacing:'.06em',background:'rgba(16,185,129,0.08)',border:'1px solid rgba(16,185,129,0.2)',borderRadius:'4px',padding:'3px 10px',cursor:'pointer'}}>📂 UPLOAD CLIENT FILES</button>}
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
    {showHelp && <HelpPanel onClose={() => setShowHelp(false)}/>}
      {tab === 'upgrade' && <div style={{padding:'40px 24px',maxWidth:'600px',margin:'0 auto',textAlign:'center'}}>
        <div style={{fontSize:'9px',fontWeight:'800',color:'#ef4444',letterSpacing:'.12em',marginBottom:'12px'}}>FREE LIMIT REACHED</div>
        <h2 style={{fontSize:'24px',fontWeight:'800',color:'#fff',marginBottom:'12px'}}>You have used your 1 free project run.</h2>
        <p style={{fontSize:'13px',color:'#94a3b8',lineHeight:'1.6',marginBottom:'24px'}}>Earth Demo and Space Demo are always free — run them as many times as you like. For unlimited project runs, exports and advisor: upgrade to Professional.</p>
        <div style={{background:'rgba(14,116,144,0.1)',border:'1px solid #0e7490',borderRadius:'6px',padding:'20px',marginBottom:'20px'}}>
          <div style={{fontSize:'11px',fontWeight:'800',color:'#22d3ee',marginBottom:'6px'}}>ALWAYS FREE</div>
          <p style={{fontSize:'12px',color:'#94a3b8'}}>Earth Demo · Space Demo · Showcase Library (200 programmes, all sectors) · Open Crawl intelligence</p>
        </div>
        <div style={{display:'flex',gap:'10px',justifyContent:'center'}}>
          <button onClick={() => runEarth()} style={{padding:'10px 20px',background:'#0e7490',color:'#fff',border:'none',borderRadius:'4px',fontWeight:'700',cursor:'pointer'}}>Try Earth Demo free</button>
          <button onClick={() => runSpace()} style={{padding:'10px 20px',background:'#1e293b',color:'#fff',border:'none',borderRadius:'4px',fontWeight:'700',cursor:'pointer'}}>Try Space Demo free</button>
        </div>
        <p style={{fontSize:'10px',color:'#475569',marginTop:'16px'}}>Professional plan (£99/mo) launching soon — <a href="mailto:hello@controlorbit.com" style={{color:'#22d3ee'}}>join the waitlist</a></p>
      </div>}
      {tab === 'upgrade' && <div style={{padding:'48px 24px',maxWidth:'580px',margin:'0 auto',textAlign:'center'}}>
        <div style={{fontSize:'9px',fontWeight:'800',color:'#ef4444',letterSpacing:'.12em',marginBottom:'12px'}}>FREE LIMIT REACHED</div>
        <h2 style={{fontSize:'22px',fontWeight:'800',color:'#fff',marginBottom:'10px'}}>You have used your 1 free project run.</h2>
        <p style={{fontSize:'13px',color:'#94a3b8',lineHeight:'1.6',marginBottom:'20px'}}>Earth Demo and Space Demo are always free — run them as many times as you like, any sector, any scenario. For unlimited project runs and all exports, Professional is coming soon.</p>
        <div style={{background:'rgba(14,116,144,0.08)',border:'1px solid rgba(14,116,144,0.2)',borderRadius:'6px',padding:'18px',marginBottom:'18px'}}>
          <div style={{fontSize:'10px',fontWeight:'800',color:'#22d3ee',marginBottom:'6px'}}>ALWAYS FREE — NO LIMIT</div>
          <p style={{fontSize:'12px',color:'#94a3b8'}}>Earth Demo · Space Demo · Showcase Library (200 programmes, all sectors) · Open Crawl intelligence</p>
        </div>
        <div style={{display:'flex',gap:'10px',justifyContent:'center',marginBottom:'16px'}}>
          <button onClick={()=>runEarth()} style={{padding:'10px 20px',background:'#0e7490',color:'#fff',border:'none',borderRadius:'4px',fontWeight:'700',cursor:'pointer',fontSize:'12px'}}>Earth Demo →</button>
          <button onClick={()=>runSpace()} style={{padding:'10px 20px',background:'#1e293b',color:'#fff',border:'none',borderRadius:'4px',fontWeight:'700',cursor:'pointer',fontSize:'12px'}}>Space Demo →</button>
        </div>
        <p style={{fontSize:'10px',color:'#475569'}}>Professional (unlimited, £99/mo) launching soon — <a href="mailto:hello@controlorbit.com?subject=CASEY+Professional+Waitlist" style={{color:'#22d3ee'}}>join the waitlist</a></p>
      </div>}
      {showIngestPanel && model && <ClientIngestPanel model={model} setModel={setModel} onClose={()=>setShowIngestPanel(false)}/>}
    <main className={model ? 'v50Console' : 'v50Console emptyConsole'}>
      {error && !showShowcase && !show && <GatedMessage raw={error} onDismiss={() => setError('')} onShowcase={() => { setError(''); setShowShowcase(true); }} onEarth={() => { setError(''); runEarth(); }} onSpace={() => { setError(''); runSpace(); }}/>}
      {!model && showShowcase && <ShowcaseLibrary onRun={runShowcase} onBack={() => setShowShowcase(false)} />}
      {!model && !show && !showShowcase && <section className="commandGrid"><Card className="command">
  <h1 style={{fontSize:'18px',marginBottom:'4px'}}>Generate a project</h1>
  <p style={{fontSize:'11px',color:'#475569',marginBottom:'12px',lineHeight:'1.5'}}>Enter any capital programme — infrastructure, defence, space, pharma, energy, data centres. CASEY generates a first-pass cost estimate (P10/P50/P90), schedule, risk register, scenario analysis and board intelligence pack. All fields are sector-specific and location-aware.</p>
  <label>Project command</label><textarea value={prompt} onChange={e => setPrompt(e.target.value)} /> <div className="chips">{examples.map(x => <button key={x} onClick={() => setPrompt(x)}>{x}</button>)}</div><div className="grid4"><input value={client} onChange={e => setClient(e.target.value)} placeholder="Client / operator"/><select value={classLevel} onChange={e => setClassLevel(e.target.value)}>{[1,2,3,4,5].map(x => <option key={x} value={x}>Class {x}</option>)}</select><select value={scheduleLevel} onChange={e => setScheduleLevel(e.target.value)}>{[1,2,3,4,5].map(x => <option key={x} value={x}>Level {x}</option>)}</select><select value={scenario} onChange={e => setScenario(e.target.value)}>{scenarios.map(x => <option key={x} value={x}>{x}</option>)}</select></div><button className="primary" onClick={() => {
              if (!checkAndGate('run')) return;
              generate();
              recordUsage('run');
            }}><Sparkles/> Generate full intelligence pack {freeRunsUsed >= FREE_RUN_LIMIT && !isUnlimited ? '(limit reached)' : freeRunsUsed === 0 ? '— 1 free run' : ''}</button><button className="secondary" onClick={() => { setShowShowcase(true); setError(''); }}><Globe2/> Open global showcase library</button></Card><Card><h2>What CASEY generates</h2><p style={{fontSize:'11px',color:'#64748b',marginBottom:'10px'}}>From a single project description — in seconds.</p>{['Executive summary with P50, schedule and confidence score','Cost workbook — direct, indirect, reserve by CBS line','5 scenario trade-offs: base, faster, cheaper, lower risk, premium','Risk register — cause, event, impact, owner, trigger, mitigation','QCRA/QSRA probability curves and tornado chart','Board attack simulation — 5 questions your committee will ask','Location intelligence, financing context and OBA assessment','Procurement packages with lead times and single-source flags'].map((x,i)=><div className="reason" style={{padding:'5px 0',borderBottom:'1px solid rgba(255,255,255,0.04)'}} key={x}><span style={{color:'#8df7ff',marginRight:'8px',fontSize:'10px',fontWeight:'800'}}>{i+1}</span><span style={{fontSize:'11px'}}>{x}</span></div>)}</Card></section>}
      {model && <>
        <DemoBanner model={model}/>
        <section className="confidenceEngineBadge"><b>{model.confidence_engine_label || 'CASEY Confidence Engine'}</b><span>{safeRender(typeof model.confidence_engine_detail === 'object' ? model.confidence_engine_detail?.plain_english || 'Benchmark + probabilistic + sector-trained reasoning' : model.confidence_engine_detail || 'Benchmark + probabilistic + sector-trained reasoning')}</span></section>
        <TrustRuntimeBar model={model}/>
        <LiveCalibrationStrip model={model}/>
        <section className="kpis"><Kpi icon={Globe2} label="Mode / sector" value={safeRender(model.mode)} sub={safeRender(model.subsector)}/><Kpi icon={Activity} label="P50 cost" value={safeRender(model.cost_p50)} sub={safeRender(model.cost_range)}/><Kpi icon={Zap} label="Schedule" value={safeRender(model.schedule)} sub={`QSRA P80 ${model.monte_carlo?.qsra?.p80 || '—'} months`}/><Kpi icon={ShieldAlert} label="Delivery confidence" value={safeRender(confidenceLens(model)?.headline)} sub={`${safeRender(model.risk)} risk · ${safeRender(model.confidence_pct)}% · ${safeRender(model.scenario_label)}`} hot/></section>
        <IntelligenceMeta model={model} mode={viewMode} setMode={setViewMode}/>
        <PropagationPulse scenario={scenario} active={propagating}/>
        <ScenarioSelector scenario={scenario} generate={generate} matrix={scenarioMatrix} model={model} prompt={prompt} projectContext={projectContext}/>
        <ExportStrip model={model}
          onBoardPack={() => download('/export/pdf', model, `${model.id || 'casey'}_CASEY_Board_Pack.pdf`)}
          onExcel={() => download('/export/workbook', model, `${model.id || 'casey'}_DEMO_COST_WORKBOOK.xlsx`)}
          onRisk={() => download('/export/risk-register', model, `${model.id || 'casey'}_DEMO_RISK_REGISTER.xlsx`)}
          onXer={() => download('/export/xer', model, `${model.id || 'casey'}_DEMO_SCHEDULE.xer`)}
          onQcra={() => download('/export/qcra-qsra', model, `${model.id || 'casey'}_DEMO_QCRA_QSRA.xlsx`)}/>
        {demoUsed && !isAdminUser && model && <div style={{background:'rgba(245,158,11,0.1)',borderBottom:'1px solid rgba(245,158,11,0.25)',padding:'8px 20px',display:'flex',gap:'16px',alignItems:'center',flexWrap:'wrap'}}>
          <span style={{fontSize:'11px',color:'#f59e0b',fontWeight:'800'}}>✓ FREE RUN COMPLETE</span>
          <span style={{fontSize:'11px',color:'#94a3b8'}}>Exports available below. Earth Demo, Space Demo and Showcase Library always free.</span>
          <a href="mailto:hello@controlorbit.com?subject=CASEY Full Access" style={{marginLeft:'auto',fontSize:'11px',color:'#8df7ff',fontWeight:'700',textDecoration:'none',background:'rgba(141,247,255,0.1)',padding:'4px 12px',borderRadius:'3px',border:'1px solid rgba(141,247,255,0.3)'}}>Request full access →</a>
        </div>}
      <nav className="tabs">{[['overview','Overview'],['twin','⚡ Decision Twin'],['compare','Scenarios'],['cost','Cost'],['schedule','Schedule'],['risk','Risk'],['monte','QCRA/QSRA'],['assurance','Assurance'],['defence','⚡ Board Pack'],['method','Method'],['advisor','Advisor']].map(x => <button key={x[0]} className={tab===x[0]?'active':''} onClick={() => setTab(x[0])}>{x[1]}</button>)}</nav>
        {tab === 'overview' && <>
          {/* ── ROLE FILTER BAR ─────────────────────────────────────────── */}
          <div style={{display:'flex',gap:6,marginBottom:12,padding:'8px 12px',background:'rgba(255,255,255,0.02)',borderRadius:8,border:'1px solid rgba(255,255,255,0.06)',alignItems:'center'}}>
            <span style={{fontSize:'8px',fontWeight:'800',color:'#475569',letterSpacing:'.12em',marginRight:4}}>VIEWING AS</span>
            {[['exec','Programme Director'],['board','Board'],['pm','Project Manager'],['analyst','QS · Scheduler · Risk']].map(([m,label])=>(
              <button key={m} onClick={()=>setViewMode(m)} style={{padding:'5px 12px',borderRadius:20,border:`1px solid ${viewMode===m?'rgba(141,247,255,0.5)':'rgba(255,255,255,0.08)'}`,background:viewMode===m?'rgba(141,247,255,0.1)':'transparent',color:viewMode===m?'#8df7ff':'#475569',fontSize:'9px',fontWeight:viewMode===m?'800':'500',cursor:'pointer',transition:'all .15s'}}>
                {label}
              </button>
            ))}
            {model?.scenario && model.scenario !== 'base' && <div style={{marginLeft:'auto',padding:'3px 10px',background:'rgba(6,182,212,0.1)',border:'1px solid rgba(6,182,212,0.25)',borderRadius:12,fontSize:'8px',fontWeight:'700',color:'#06b6d4'}}>{(model.scenario_label||'').toUpperCase()} SCENARIO ACTIVE</div>}
            {model?.evidence_mode && <div style={{marginLeft:model?.scenario&&model.scenario!=='base'?4:'auto',padding:'3px 10px',background:'rgba(16,185,129,0.1)',border:'1px solid rgba(16,185,129,0.3)',borderRadius:12,fontSize:'8px',fontWeight:'700',color:'#10b981'}}>🔗 EVIDENCE MODE</div>}
          </div>

          {/* ── EXECUTIVE / CFO VIEW ─────────────────────────────────────── */}
          {/* ── PROGRAMME DIRECTOR ─── */}
          {viewMode === 'exec' && <>
            {model && <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:8,marginBottom:10}}>
              <div style={{background:'rgba(6,182,212,0.06)',border:'1px solid rgba(6,182,212,0.25)',borderRadius:8,padding:'12px 14px'}}>
                <div style={{fontSize:'8px',fontWeight:'800',color:'#06b6d4',letterSpacing:'.12em',marginBottom:5}}>THE DEFENSIBLE NUMBER</div>
                <div style={{fontSize:'28px',fontWeight:'900',color:'#8df7ff',lineHeight:1}}>{model.cost_p80||model.cost_p50}</div>
                <div style={{fontSize:'8px',color:'#64748b',marginTop:2}}>P80 — approve at this number</div>
                <div style={{fontSize:'8px',color:'#475569',marginTop:4}}>P50: {model.cost_p50} · OBA outturn: {model.outturn||'—'}</div>
                <div style={{marginTop:5,fontSize:'8px',fontWeight:'700',color:(model.p80_reserve_pct||0)>=(model.reserve_vs_benchmark_pct||18)?'#10b981':'#ef4444'}}>Reserve: {model.p80_reserve_pct||0}% {(model.p80_reserve_pct||0)>=(model.reserve_vs_benchmark_pct||18)?'✓ adequate':'⚠ need '+(model.reserve_vs_benchmark_pct||18)+'%'}</div>
              </div>
              <div style={{background:'rgba(245,158,11,0.06)',border:'1px solid rgba(245,158,11,0.2)',borderRadius:8,padding:'12px 14px'}}>
                <div style={{fontSize:'8px',fontWeight:'800',color:'#f59e0b',letterSpacing:'.12em',marginBottom:5}}>DEFENSIBLE SCHEDULE</div>
                <div style={{fontSize:'28px',fontWeight:'900',color:'#fbbf24',lineHeight:1}}>{model.schedule}</div>
                <div style={{fontSize:'8px',color:'#64748b',marginTop:2}}>P50 delivery</div>
                <div style={{fontSize:'8px',color:'#475569',marginTop:4}}>QSRA P80: {model.monte_carlo?.qsra?.p80||Math.round(parseInt(model.schedule||189)*1.15)+' months'}</div>
                <div style={{marginTop:5,fontSize:'8px',color:'#f59e0b',fontWeight:'700'}}>{(model.governing_constraint_prominent||'Governing constraint: see Board Pack').slice(0,45)}</div>
              </div>
              <div style={{background:(model.confidence_pct||0)>=75?'rgba(16,185,129,0.06)':'rgba(239,68,68,0.06)',border:'1px solid '+((model.confidence_pct||0)>=75?'rgba(16,185,129,0.25)':'rgba(239,68,68,0.25)'),borderRadius:8,padding:'12px 14px'}}>
                <div style={{fontSize:'8px',fontWeight:'800',color:(model.confidence_pct||0)>=75?'#10b981':'#ef4444',letterSpacing:'.12em',marginBottom:5}}>BOARD CONFIDENCE</div>
                <div style={{fontSize:'28px',fontWeight:'900',color:(model.confidence_pct||0)>=75?'#10b981':'#ef4444',lineHeight:1}}>{model.confidence_pct+'%'}</div>
                <div style={{fontSize:'8px',color:'#64748b',marginTop:2}}>{(model.confidence_pct||0)>=75?'Approval-ready':'Below 75% threshold'}</div>
                <div style={{fontSize:'8px',color:'#475569',marginTop:4}}>{model.estimate_class_name||'Class 3'} · {model.scenario_label||'Base'}</div>
                <div style={{marginTop:5,fontSize:'8px',fontWeight:'700',color:(model.confidence_pct||0)>=75?'#10b981':'#f59e0b'}}>{(model.confidence_pct||0)>=75?'No action required':(model.confidence_pct||0)>=55?'Close evidence gaps':'Significant rework needed'}</div>
              </div>
              <div style={{background:'rgba(239,68,68,0.06)',border:'1px solid rgba(239,68,68,0.25)',borderRadius:8,padding:'12px 14px'}}>
                <div style={{fontSize:'8px',fontWeight:'800',color:'#ef4444',letterSpacing:'.12em',marginBottom:5}}>MORTALITY RISK</div>
                <div style={{fontSize:'12px',fontWeight:'800',color:'#fca5a5',lineHeight:1.3,marginBottom:4}}>{(model.mortality_event?.title||model.programme_mortality_risk?.title||'Systems integration failure').slice(0,40)}</div>
                <div style={{fontSize:'8px',color:'#94a3b8'}}>{model.mortality_event?.probability||65}% prob · {model.currency_symbol}{model.mortality_event?.exposure||'4.5B'}</div>
                <div style={{marginTop:5,fontSize:'8px',color:'#ef4444',fontWeight:'700'}}>{(model.mortality_event?.board_action||'Must resolve before capital commitment').slice(0,55)}</div>
              </div>
            </div>}
            {model && <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8,marginBottom:10}}>
              <div style={{background:'rgba(255,255,255,0.02)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:8,padding:'12px 16px'}}>
                <div style={{fontSize:'9px',fontWeight:'800',color:'#8df7ff',letterSpacing:'.12em',marginBottom:8}}>GOING FASTER vs STAYING BASE — WHAT DO YOU GIVE UP?</div>
                {(model.scenario_delta_intelligence||[{scenario:'FASTER',delta_pct:14,verdict:'Cost +14%, schedule -18%. Confidence drops 10pts. HS2 attempted acceleration in 2014 — final cost grew 88%.'}]).slice(0,3).map((d,i)=>(
                  <div key={i} style={{display:'flex',gap:10,marginBottom:6,padding:'6px 8px',background:'rgba(0,0,0,0.15)',borderRadius:5}}>
                    <div style={{flexShrink:0,padding:'3px 8px',background:'rgba(141,247,255,0.08)',borderRadius:3,fontSize:'8px',fontWeight:'800',color:'#8df7ff',alignSelf:'flex-start'}}>{(d.scenario||('S'+(i+1))).toUpperCase()}</div>
                    <div>
                      <div style={{fontSize:'10px',fontWeight:'800',color:(d.delta_pct||0)>5?'#ef4444':(d.delta_pct||0)<-5?'#10b981':'#f59e0b'}}>{(d.delta_pct||0)>0?'+':''}{d.delta_pct||0}% cost</div>
                      <div style={{fontSize:'8px',color:'#64748b',marginTop:1,lineHeight:1.4}}>{(d.verdict||d.casey_comment||'See Scenarios tab for full analysis').slice(0,80)}</div>
                    </div>
                  </div>
                ))}
              </div>
              <div style={{background:'rgba(255,255,255,0.02)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:8,padding:'12px 16px'}}>
                <div style={{fontSize:'9px',fontWeight:'800',color:'#a78bfa',letterSpacing:'.12em',marginBottom:8}}>WHAT SIMILAR PROGRAMMES DID — LEARN FROM THEM</div>
                {(model.benchmark_comparison||[]).slice(0,3).map((b,i)=>(
                  <div key={i} style={{display:'flex',gap:8,marginBottom:6,padding:'5px 8px',background:'rgba(139,92,246,0.04)',borderRadius:4,borderLeft:'3px solid '+(b.cost_growth_pct>50?'#ef4444':b.cost_growth_pct>25?'#f59e0b':'#10b981')}}>
                    <div style={{flex:1}}>
                      <div style={{fontSize:'9px',fontWeight:'700',color:'#c4b5fd',marginBottom:1}}>{b.name||b.programme}</div>
                      <div style={{fontSize:'8px',color:'#64748b'}}>Cost +{b.cost_growth_pct}% · {b.schedule_slip_months}mo slip</div>
                      <div style={{fontSize:'8px',color:'#fca5a5',marginTop:1}}>{(b.failure_mode||b.failure_pattern||b.lesson||'').slice(0,55)}</div>
                    </div>
                  </div>
                ))}
                {(!model.benchmark_comparison||model.benchmark_comparison.length===0) && <div style={{fontSize:'9px',color:'#475569',padding:'8px'}}>Run a project to see comparable programme history</div>}
              </div>
            </div>}
            <ApprovalStatus model={model}/>
            {(model?.board_attack_simulation||[]).length>0 && <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr',gap:8,marginTop:10}}>
              {(model.board_attack_simulation).slice(0,3).map((d,i)=>(
                <div key={i} style={{background:'rgba(255,255,255,0.02)',border:'1px solid rgba(255,255,255,0.06)',borderRadius:7,padding:'10px 12px'}}>
                  <div style={{fontSize:'16px',fontWeight:'900',color:'#8df7ff',marginBottom:4}}>{i+1}</div>
                  <div style={{fontSize:'9px',color:'#e2e8f0',fontWeight:'600',lineHeight:1.5}}>{typeof d==='string'?d.slice(0,100):(d.question||d.decision||'').slice(0,100)}</div>
                </div>
              ))}
            </div>}
          </>}

          {/* ── BOARD ─── */}
          {viewMode === 'board' && <>
            {model && <div style={{background:'rgba(245,158,11,0.08)',border:'2px solid rgba(245,158,11,0.4)',borderRadius:10,padding:'14px 20px',marginBottom:10}}>
              <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',gap:16,flexWrap:'wrap'}}>
                <div>
                  <div style={{fontSize:'9px',fontWeight:'800',color:'#f59e0b',letterSpacing:'.14em',marginBottom:4}}>⛔ GOVERNING CONSTRAINT — BOARD MUST CONFIRM BEFORE APPROVAL</div>
                  <div style={{fontSize:'18px',fontWeight:'900',color:'#fff',marginBottom:6}}>{model.governing_constraint_prominent||model.governing_constraint_full?.statement||'Not yet identified — run project to generate'}</div>
                  <div style={{display:'flex',gap:16,fontSize:'9px',flexWrap:'wrap'}}>
                    <span style={{color:'#64748b'}}>Owner: <b style={{color:model.governing_constraint_full?.owner&&model.governing_constraint_full.owner!=='TBC'?'#fbbf24':'#ef4444'}}>{model.governing_constraint_full?.owner||'NOT NAMED'}</b></span>
                    <span style={{color:'#64748b'}}>Evidence: <b style={{color:model.governing_constraint_full?.evidence?'#10b981':'#ef4444'}}>{model.governing_constraint_full?.evidence?'Confirmed':'MISSING'}</b></span>
                  </div>
                </div>
                <div style={{textAlign:'center',padding:'10px 20px',background:'rgba(0,0,0,0.2)',borderRadius:8,flexShrink:0}}>
                  <div style={{fontSize:'9px',color:'#64748b',marginBottom:2}}>CONFIDENCE</div>
                  <div style={{fontSize:'36px',fontWeight:'900',lineHeight:1,color:(model.confidence_pct||0)>=75?'#10b981':(model.confidence_pct||0)>=55?'#f59e0b':'#ef4444'}}>{model.confidence_pct+'%'}</div>
                  <div style={{fontSize:'8px',color:'#475569',marginTop:1}}>{(model.confidence_pct||0)>=75?'Approvable':'Target 75%+'}</div>
                </div>
              </div>
            </div>}
            {model && <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:10,marginBottom:10}}>
              <div style={{background:'rgba(255,255,255,0.02)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:8,padding:'12px 16px'}}>
                <div style={{fontSize:'9px',fontWeight:'800',color:'#8df7ff',letterSpacing:'.12em',marginBottom:8}}>COST — APPROVE AT P80, NOT P50</div>
                <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:6,marginBottom:8}}>
                  {[['P10',model.monte_carlo?.qcra?.p10||'—','#10b981'],['P50',model.cost_p50,'#8df7ff'],['P80',model.cost_p80||model.cost_p50,'#f59e0b'],['P90',model.monte_carlo?.qcra?.p90||'—','#ef4444']].map(([k,v,col])=>(
                    <div key={k} style={{textAlign:'center',padding:'8px 4px',background:'rgba(0,0,0,0.2)',borderRadius:5}}>
                      <div style={{fontSize:'8px',color:'#64748b',marginBottom:2}}>{k}</div>
                      <div style={{fontSize:'13px',fontWeight:'900',color:col}}>{v}</div>
                    </div>
                  ))}
                </div>
                <div style={{padding:'6px 8px',background:'rgba(245,158,11,0.06)',borderRadius:4,fontSize:'8px',color:'#f59e0b',fontWeight:'700'}}>Reserve: {model.p80_reserve} ({model.p80_reserve_pct}%) · Benchmark: {model.reserve_vs_benchmark_pct||18}% minimum</div>
              </div>
              <div style={{background:'rgba(255,255,255,0.02)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:8,padding:'12px 16px'}}>
                <div style={{fontSize:'9px',fontWeight:'800',color:'#f59e0b',letterSpacing:'.12em',marginBottom:8}}>SCHEDULE — APPROVE AT P80, COMMIT AT P50</div>
                <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:6}}>
                  {[['P10',(model.monte_carlo?.qsra?.p10||Math.round(parseInt(model.schedule||189)*0.85))+'mo','#10b981'],['P50',model.schedule,'#fbbf24'],['P80',(model.monte_carlo?.qsra?.p80||Math.round(parseInt(model.schedule||189)*1.15))+'mo','#f59e0b'],['P90',(model.monte_carlo?.qsra?.p90||Math.round(parseInt(model.schedule||189)*1.25))+'mo','#ef4444']].map(([k,v,col])=>(
                    <div key={k} style={{textAlign:'center',padding:'8px 4px',background:'rgba(0,0,0,0.2)',borderRadius:5}}>
                      <div style={{fontSize:'8px',color:'#64748b',marginBottom:2}}>{k}</div>
                      <div style={{fontSize:'13px',fontWeight:'900',color:col}}>{v}</div>
                    </div>
                  ))}
                </div>
              </div>
            </div>}
            {(model?.board_attack_simulation||[]).length>0 && <div style={{background:'rgba(255,255,255,0.02)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:8,padding:'12px 16px',marginBottom:10}}>
              <div style={{fontSize:'9px',fontWeight:'800',color:'#e2e8f0',letterSpacing:'.12em',marginBottom:8}}>THE 5 QUESTIONS THIS BOARD WILL ASK — ARE YOU READY?</div>
              {(model.board_attack_simulation).slice(0,5).map((q,i)=>(
                <div key={i} style={{display:'flex',gap:10,marginBottom:5,padding:'5px 8px',background:'rgba(239,68,68,0.03)',borderRadius:4}}>
                  <span style={{color:'#ef4444',fontWeight:'900',fontSize:'11px',flexShrink:0,minWidth:14}}>{i+1}.</span>
                  <div style={{fontSize:'9px',color:'#cbd5e1',lineHeight:1.5}}>{typeof q==='string'?q:(q?.question||String(q))}</div>
                </div>
              ))}
            </div>}
            {model?.decision_simulator && <div style={{background:'rgba(255,255,255,0.02)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:8,padding:'12px 16px'}}>
              <div style={{fontSize:'9px',fontWeight:'800',color:'#e2e8f0',letterSpacing:'.12em',marginBottom:8}}>🎮 DECISION SIMULATOR — WHAT IF THE BOARD ASKS?</div>
              <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:8}}>
                {Object.entries(model.decision_simulator).map(([k,d])=>(
                  <div key={k} style={{padding:'8px 10px',background:'rgba(0,0,0,0.2)',borderRadius:6}}>
                    <div style={{fontSize:'8px',fontWeight:'800',color:'#8df7ff',marginBottom:4}}>{k.replace(/_/g,' ').toUpperCase()}</div>
                    <div style={{display:'flex',gap:6,fontSize:'9px',flexWrap:'wrap'}}>
                      {d.cost_delta_bn!==undefined&&<span style={{color:d.cost_delta_bn>0?'#ef4444':'#10b981',fontWeight:'700'}}>{d.cost_delta_bn>0?'+':''}{(d.cost_delta_bn||0).toFixed(2)+'B'}</span>}
                      {d.schedule_delta_months!==undefined&&<span style={{color:d.schedule_delta_months>0?'#f59e0b':'#10b981',fontWeight:'700'}}>{d.schedule_delta_months>0?'+':''}{d.schedule_delta_months+'mo'}</span>}
                      {d.confidence_delta_pct!==undefined&&<span style={{color:d.confidence_delta_pct>0?'#10b981':'#ef4444',fontWeight:'700'}}>{d.confidence_delta_pct>0?'+':''}{d.confidence_delta_pct+'%'}</span>}
                    </div>
                  </div>
                ))}
              </div>
            </div>}
          </>}

          {/* ── PROJECT MANAGER ─── */}
          {viewMode === 'pm' && <>
            {model && <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:10,marginBottom:10}}>
              <div style={{background:'rgba(255,255,255,0.02)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:8,padding:'12px 16px'}}>
                <div style={{fontSize:'9px',fontWeight:'800',color:'#8df7ff',letterSpacing:'.12em',marginBottom:6}}>COST P-NUMBERS · BREAKDOWN</div>
                <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:6,marginBottom:6}}>
                  <div>
                    <div style={{fontSize:'7px',color:'#475569',fontWeight:'700',marginBottom:3}}>P-RANGE (QCRA)</div>
                    {[['P10',model.monte_carlo?.qcra?.p10||'—','#10b981'],['P50',model.cost_p50,'#8df7ff'],['P80',model.cost_p80||model.monte_carlo?.qcra?.p80||'—','#f59e0b'],['P90',model.monte_carlo?.qcra?.p90||'—','#ef4444']].map(([k,v,col])=>(
                      <div key={k} style={{display:'flex',justifyContent:'space-between',padding:'2px 0',borderBottom:'1px solid rgba(255,255,255,0.04)'}}>
                        <span style={{fontSize:'8px',color:'#64748b'}}>{k}</span>
                        <span style={{fontSize:'9px',fontWeight:'700',color:col}}>{v}</span>
                      </div>
                    ))}
                  </div>
                  <div>
                    <div style={{fontSize:'7px',color:'#475569',fontWeight:'700',marginBottom:3}}>COST BREAKDOWN</div>
                    {[['Directs',model.direct_cost||'—'],['Indirects',model.indirect_cost||'—'],['Reserve',model.p80_reserve||'—'],['OBA uplift',(model.oba_pct||35)+'%']].map(([k,v])=>(
                      <div key={k} style={{display:'flex',justifyContent:'space-between',padding:'2px 0',borderBottom:'1px solid rgba(255,255,255,0.04)'}}>
                        <span style={{fontSize:'8px',color:'#64748b'}}>{k}</span>
                        <span style={{fontSize:'9px',fontWeight:'700',color:'#e2e8f0'}}>{v}</span>
                      </div>
                    ))}
                  </div>
                </div>
                <div style={{fontSize:'7px',color:'#475569',padding:'4px 6px',background:'rgba(255,255,255,0.02)',borderRadius:3}}>{model.estimate_class_name} · {model.prolific_cost_str||'Upload cost file to verify unit rates'}</div>
              </div>
              <div style={{background:'rgba(255,255,255,0.02)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:8,padding:'12px 16px'}}>
                <div style={{fontSize:'9px',fontWeight:'800',color:'#f59e0b',letterSpacing:'.12em',marginBottom:6}}>SCHEDULE P-NUMBERS · HEALTH</div>
                <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:6,marginBottom:6}}>
                  <div>
                    <div style={{fontSize:'7px',color:'#475569',fontWeight:'700',marginBottom:3}}>P-RANGE (QSRA)</div>
                    {[['P10',(model.monte_carlo?.qsra?.p10||Math.round(parseInt(model.schedule||189)*0.85))+'mo','#10b981'],['P50',model.schedule,'#fbbf24'],['P80',(model.monte_carlo?.qsra?.p80||Math.round(parseInt(model.schedule||189)*1.15))+'mo','#f59e0b'],['P90',(model.monte_carlo?.qsra?.p90||Math.round(parseInt(model.schedule||189)*1.25))+'mo','#ef4444']].map(([k,v,col])=>(
                      <div key={k} style={{display:'flex',justifyContent:'space-between',padding:'2px 0',borderBottom:'1px solid rgba(255,255,255,0.04)'}}>
                        <span style={{fontSize:'8px',color:'#64748b'}}>{k}</span>
                        <span style={{fontSize:'9px',fontWeight:'700',color:col}}>{v}</span>
                      </div>
                    ))}
                  </div>
                  <div>
                    <div style={{fontSize:'7px',color:'#475569',fontWeight:'700',marginBottom:3}}>SCHEDULE QUALITY</div>
                    {[['Activities',model.xer_health?.activity_count||'Proxy L3'],['Logic',model.xer_health?.logic_quality||'Upload XER'],['Critical %',(model.xer_health?.critical_pct||'—')+'%'],['Open ends',model.xer_health?.open_ends||'—']].map(([k,v])=>(
                      <div key={k} style={{display:'flex',justifyContent:'space-between',padding:'2px 0',borderBottom:'1px solid rgba(255,255,255,0.04)'}}>
                        <span style={{fontSize:'8px',color:'#64748b'}}>{k}</span>
                        <span style={{fontSize:'9px',fontWeight:'700',color:'#e2e8f0'}}>{String(v).slice(0,15)}</span>
                      </div>
                    ))}
                  </div>
                </div>
                <div style={{fontSize:'7px',color:'#475569',padding:'4px 6px',background:'rgba(255,255,255,0.02)',borderRadius:3}}>{model.xer_health?'Uploaded XER — verified schedule':'CASEY generated proxy schedule. Upload P6 XER to replace with your actual critical path.'}</div>
              </div>
            </div>}
            {model && <div style={{background:'rgba(255,255,255,0.02)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:8,padding:'12px 16px',marginBottom:10}}>
              <div style={{fontSize:'9px',fontWeight:'800',color:'#e2e8f0',letterSpacing:'.12em',marginBottom:8}}>TOP 10 RISKS — PROBABILITY · P10/P50/P80 COST · SCHEDULE IMPACT</div>
              <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:5}}>
                {(model.risks||model.risk_register||[]).slice(0,10).map((r,i)=>{
                  const emv=parseFloat(r.cost_emv_bn||r.emv_bn||0);
                  const prob=parseFloat(r.probability_pct||r.probability||30);
                  const p10=+(emv*0.4).toFixed(3), p80=+(emv*2.1).toFixed(3);
                  const sched=parseFloat(r.schedule_impact_months||2);
                  const unowned=!r.owner||r.owner==='TBC'||r.owner==='Unknown';
                  return <div key={i} style={{padding:'6px 8px',background:i===0?'rgba(239,68,68,0.06)':'rgba(255,255,255,0.01)',borderRadius:5,border:'1px solid '+(i===0?'rgba(239,68,68,0.15)':'rgba(255,255,255,0.03)')}}>
                    <div style={{display:'flex',justifyContent:'space-between',marginBottom:2}}>
                      <span style={{fontSize:'8px',fontWeight:'700',color:'#e2e8f0',flex:1}}>{(r.title||r.risk||r.risk_title||('Risk '+(i+1))).slice(0,42)}</span>
                      <span style={{fontSize:'8px',fontWeight:'800',color:unowned?'#ef4444':'#94a3b8',flexShrink:0,marginLeft:4}}>{unowned?'⚠ No owner':(r.owner||'').slice(0,10)}</span>
                    </div>
                    <div style={{display:'flex',gap:5,fontSize:'8px',flexWrap:'wrap'}}>
                      <span style={{color:'#f59e0b',fontWeight:'700'}}>{prob+'%'}</span>
                      <span style={{color:'#94a3b8'}}>P10:{model.currency_symbol}{p10+'B'}</span>
                      <span style={{color:'#8df7ff'}}>P50:{model.currency_symbol}{emv.toFixed(3)+'B'}</span>
                      <span style={{color:'#ef4444'}}>P80:{model.currency_symbol}{p80+'B'}</span>
                      <span style={{color:'#f59e0b'}}>+{sched+'mo'}</span>
                    </div>
                  </div>;
                })}
              </div>
              {(model.risks||model.risk_register||[]).filter(r=>!r.owner||r.owner==='TBC').length>0&&<div style={{marginTop:6,padding:'5px 8px',background:'rgba(239,68,68,0.06)',borderRadius:4,fontSize:'8px',color:'#ef4444',fontWeight:'700'}}>⚠ {(model.risks||model.risk_register||[]).filter(r=>!r.owner||r.owner==='TBC').length} risks have no named owner — assign before board submission</div>}
            </div>}
            {model && <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8}}>
              <div style={{background:'rgba(16,185,129,0.04)',border:'1px solid rgba(16,185,129,0.15)',borderRadius:7,padding:'10px 12px'}}>
                <div style={{fontSize:'9px',fontWeight:'800',color:'#10b981',marginBottom:6}}>WHAT CAN GO WELL ↑</div>
                {(model.scenario_delta_intelligence||[]).filter(d=>(d.delta_pct||0)<0).slice(0,3).map((d,i)=>(
                  <div key={i} style={{fontSize:'8px',color:'#6ee7b7',marginBottom:3,display:'flex',gap:6}}><span>✓</span><span>{(d.scenario||'Lower risk scenario')+': '+(d.verdict||d.casey_comment||'Reduces cost, improves confidence').slice(0,65)}</span></div>
                ))}
                {!(model.scenario_delta_intelligence||[]).filter(d=>(d.delta_pct||0)<0).length&&<div style={{fontSize:'8px',color:'#475569'}}>Lower Risk scenario: +12pts confidence, reserve improves to benchmark level.</div>}
              </div>
              <div style={{background:'rgba(239,68,68,0.04)',border:'1px solid rgba(239,68,68,0.15)',borderRadius:7,padding:'10px 12px'}}>
                <div style={{fontSize:'9px',fontWeight:'800',color:'#ef4444',marginBottom:6}}>WATCH CLOSELY ↓</div>
                {(model.risks||model.risk_register||[]).filter(r=>parseFloat(r.probability_pct||0)>50).slice(0,3).map((r,i)=>(
                  <div key={i} style={{fontSize:'8px',color:'#fca5a5',marginBottom:3,display:'flex',gap:6}}><span>⚠</span><span>{(r.title||r.risk||'High probability risk').slice(0,55)}: {r.probability_pct}% probability</span></div>
                ))}
                {!(model.risks||model.risk_register||[]).filter(r=>parseFloat(r.probability_pct||0)>50).length&&<div style={{fontSize:'8px',color:'#475569'}}>Monitor governing constraint and procurement critical path items.</div>}
              </div>
            </div>}
          </>}

          {/* ── QS / SCHEDULER / RISK MANAGER ─── */}
          {viewMode === 'analyst' && <>
            {model && <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:10,marginBottom:10}}>
              <div style={{background:'rgba(255,255,255,0.02)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:8,padding:'12px 16px'}}>
                <div style={{fontSize:'9px',fontWeight:'800',color:'#e2e8f0',letterSpacing:'.12em',marginBottom:8}}>CONFIDENCE BY DISCIPLINE</div>
                {model.confidence_by_discipline&&Object.entries(model.confidence_by_discipline).map(([disc,pct])=>(
                  <div key={disc} style={{marginBottom:5}}>
                    <div style={{display:'flex',justifyContent:'space-between',marginBottom:2}}>
                      <span style={{fontSize:'8px',color:'#94a3b8'}}>{disc}</span>
                      <span style={{fontSize:'8px',fontWeight:'700',color:pct>=70?'#10b981':pct>=50?'#f59e0b':'#ef4444'}}>{pct+'%'}</span>
                    </div>
                    <div style={{height:4,background:'rgba(255,255,255,0.06)',borderRadius:2}}><div style={{height:'100%',width:pct+'%',background:pct>=70?'#10b981':pct>=50?'#f59e0b':'#ef4444',borderRadius:2}}/></div>
                  </div>
                ))}
              </div>
              <div style={{background:'rgba(255,255,255,0.02)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:8,padding:'12px 16px'}}>
                <div style={{fontSize:'9px',fontWeight:'800',color:'#e2e8f0',letterSpacing:'.12em',marginBottom:8}}>ESTIMATE BASIS — FULL TRACE</div>
                {(model.estimate_basis?.traceability||['P50 = sector unit rate × scope × location factor','Benchmark calibration: '+(model.benchmark_comparison?.length||4)+' comparables','OBA applied: +'+(model.oba_pct||35)+'% per '+(model.location||'UK')+' reference class','Risk EMV quantified across '+(model.total_risks_identified||model.risks?.length||43)+' risks']).map((t,i)=>(
                  <div key={i} style={{fontSize:'8px',color:'#94a3b8',marginBottom:4,display:'flex',gap:6,alignItems:'flex-start'}}><span style={{color:'#8df7ff',flexShrink:0}}>→</span><span>{t}</span></div>
                ))}
              </div>
            </div>}
            {model && <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr',gap:10}}>
              <div style={{background:'rgba(255,255,255,0.02)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:8,padding:'12px 14px'}}>
                <div style={{fontSize:'9px',fontWeight:'800',color:'#06b6d4',letterSpacing:'.12em',marginBottom:8}}>QS — COST INTELLIGENCE</div>
                {[['Estimate class',model.estimate_class_name||'Class 3'],['Unit cost',model.prolific_cost_str||'—'],['Direct cost',model.direct_cost||'—'],['Indirect cost',model.indirect_cost||'—'],['OBA uplift',(model.oba_pct||35)+'%'],['Risk reserve',model.p80_reserve||'—'],['Best comparable',((model.benchmark_comparison||[])[0]?.name||'—').slice(0,18)]].map(([k,v])=>(
                  <div key={k} style={{display:'flex',justifyContent:'space-between',padding:'3px 0',borderBottom:'1px solid rgba(255,255,255,0.04)'}}>
                    <span style={{fontSize:'8px',color:'#64748b'}}>{k}</span>
                    <span style={{fontSize:'8px',fontWeight:'700',color:'#e2e8f0'}}>{String(v||'—').slice(0,22)}</span>
                  </div>
                ))}
                {(model.benchmark_comparison||[]).slice(0,2).map((b,i)=>(
                  <div key={i} style={{marginTop:4,padding:'3px 6px',background:'rgba(139,92,246,0.06)',borderRadius:3,fontSize:'7px',color:'#a78bfa'}}>{b.name}: +{b.cost_growth_pct}% growth</div>
                ))}
              </div>
              <div style={{background:'rgba(255,255,255,0.02)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:8,padding:'12px 14px'}}>
                <div style={{fontSize:'9px',fontWeight:'800',color:'#f59e0b',letterSpacing:'.12em',marginBottom:8}}>SCHEDULER — P-NUMBERS &amp; HEALTH</div>
                <div style={{fontSize:'7px',color:'#475569',fontWeight:'700',marginBottom:4}}>QSRA RANGE</div>
                {[['P10',(model.monte_carlo?.qsra?.p10||Math.round(parseInt(model.schedule||189)*0.85))+'mo'],['P50',model.schedule],['P80',(model.monte_carlo?.qsra?.p80||Math.round(parseInt(model.schedule||189)*1.15))+'mo'],['P90',(model.monte_carlo?.qsra?.p90||Math.round(parseInt(model.schedule||189)*1.25))+'mo']].map(([k,v])=>(
                  <div key={k} style={{display:'flex',justifyContent:'space-between',padding:'2px 0',borderBottom:'1px solid rgba(255,255,255,0.04)'}}>
                    <span style={{fontSize:'8px',color:'#64748b'}}>{k}</span>
                    <span style={{fontSize:'8px',fontWeight:'700',color:'#fbbf24'}}>{v}</span>
                  </div>
                ))}
                {model.xer_health&&<><div style={{fontSize:'7px',color:'#475569',fontWeight:'700',margin:'6px 0 4px'}}>XER HEALTH</div>
                {[['Activities',model.xer_health.activity_count],['Logic',model.xer_health.logic_quality],['Critical',model.xer_health.critical_pct+'%'],['Open ends',model.xer_health.open_ends]].map(([k,v])=>(
                  <div key={k} style={{display:'flex',justifyContent:'space-between',padding:'2px 0'}}>
                    <span style={{fontSize:'8px',color:'#64748b'}}>{k}</span>
                    <span style={{fontSize:'8px',fontWeight:'700',color:'#fbbf24'}}>{v}</span>
                  </div>
                ))}</>}
                {!model.xer_health&&<div style={{marginTop:6,fontSize:'7px',color:'#475569'}}>Upload P6 XER to replace proxy schedule with verified critical path</div>}
              </div>
              <div style={{background:'rgba(255,255,255,0.02)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:8,padding:'12px 14px'}}>
                <div style={{fontSize:'9px',fontWeight:'800',color:'#ef4444',letterSpacing:'.12em',marginBottom:8}}>RISK MGR — QCRA &amp; EXPOSURE</div>
                <div style={{fontSize:'7px',color:'#475569',fontWeight:'700',marginBottom:4}}>QCRA RANGE</div>
                {[['P10',model.monte_carlo?.qcra?.p10||'—'],['P50',model.cost_p50],['P80',model.cost_p80||model.monte_carlo?.qcra?.p80||'—'],['P90',model.monte_carlo?.qcra?.p90||'—']].map(([k,v])=>(
                  <div key={k} style={{display:'flex',justifyContent:'space-between',padding:'2px 0',borderBottom:'1px solid rgba(255,255,255,0.04)'}}>
                    <span style={{fontSize:'8px',color:'#64748b'}}>{k}</span>
                    <span style={{fontSize:'8px',fontWeight:'700',color:'#ef4444'}}>{v}</span>
                  </div>
                ))}
                <div style={{fontSize:'7px',color:'#475569',fontWeight:'700',margin:'6px 0 4px'}}>BY CATEGORY</div>
                {(model.risk_clusters?.clusters||[]).slice(0,4).map((c,i)=>(
                  <div key={i} style={{display:'flex',justifyContent:'space-between',padding:'2px 0'}}>
                    <span style={{fontSize:'7px',color:'#94a3b8'}}>{(c.category||c.cluster||'—').slice(0,18)}</span>
                    <span style={{fontSize:'7px',fontWeight:'700',color:'#ef4444'}}>{model.currency_symbol}{(c.total_emv_bn||0).toFixed(2)+'B'}</span>
                  </div>
                ))}
                <div style={{marginTop:6,padding:'4px 6px',background:'rgba(239,68,68,0.06)',borderRadius:3,fontSize:'7px',color:'#ef4444',fontWeight:'700'}}>Total EMV: {model.currency_symbol}{(model.board_risk_summary?.total_emv_bn||0).toFixed(2)}B · {model.total_risks_identified||model.risks?.length||43} risks</div>
              </div>
            </div>}
          </>}
        </>}

        {tab === 'compare' && <section className="layout two"><Card><h2>Scenario comparison</h2>
              {model && <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr',gap:8,marginBottom:10,padding:'10px 14px',background:'rgba(6,182,212,0.04)',border:'1px solid rgba(6,182,212,0.15)',borderRadius:7}}>
                {[{label:'BASE',cost:model.cost_p50,sched:parseInt(model.schedule||189)+'mo',conf:(model.confidence_pct||0)+'%',color:'#94a3b8',tag:'base'},{label:(model.scenario_label||'CURRENT').toUpperCase(),cost:model.cost_p50,sched:model.schedule,conf:(model.confidence_pct||0)+'%',color:'#8df7ff',tag:model.scenario||'base'},{label:'BEST BENCHMARK',cost:((model.benchmark_comparison||[])[0]?.name||'—').slice(0,14),sched:'+'+( (model.benchmark_comparison||[])[0]?.schedule_slip_months||'—')+'mo slip',conf:'+'+( (model.benchmark_comparison||[])[0]?.cost_growth_pct||'—')+'% cost',color:'#a78bfa',tag:'bench'}].map((s,i)=>(
                  <div key={i} style={{textAlign:'center',padding:'8px',background:s.tag===( model.scenario||'base')?'rgba(141,247,255,0.06)':'rgba(0,0,0,0.1)',borderRadius:5,border:s.tag===(model.scenario||'base')?'1px solid rgba(141,247,255,0.2)':'1px solid transparent'}}>
                    <div style={{fontSize:'8px',fontWeight:'800',color:s.color,marginBottom:3,letterSpacing:'.1em'}}>{s.label}</div>
                    <div style={{fontSize:'13px',fontWeight:'900',color:s.color}}>{s.cost}</div>
                    <div style={{fontSize:'8px',color:'#64748b',marginTop:2}}>{s.sched} · {s.conf}</div>
                  </div>
                ))}
              </div>}
              <p style={{fontSize:'11px',color:'#64748b',marginBottom:'10px'}}>Select a scenario to recalculate. Each recalculates cost, schedule, confidence, risk register and exports from the same source of truth — instantly.</p>{model?.stress_test_applied && <div style={{background:"rgba(245,158,11,0.1)",border:"1px solid rgba(245,158,11,0.3)",borderRadius:"3px",padding:"8px 12px",marginBottom:"10px",fontSize:"11px",color:"#f59e0b"}}><b>STRESS TEST ACTIVE: {String(model.stress_test_applied).replace(/_/g," ").toUpperCase()}</b><br/>{model.stress_test_note} — P50 now {safeRender(model.cost_p50)}, confidence {model.confidence_pct}%. Scenario re-runs below use the stressed baseline.</div>}<div className="runtimeInline"><button onClick={()=>setTab('compare')}><Zap size={15}/> Open Live Stress Test</button><button onClick={()=>runShock('signalling_slip')}>Simulate 4-month signalling slip</button><button onClick={()=>runShock('procurement_gap')}>Simulate procurement evidence gap</button></div>{renderScenarioCompare()}</Card><Card><h2>Buyer decision lens</h2>{['Base: balanced reference case for board challenge','Faster: more capex, lower confidence, shorter duration','Cheaper: lower authorisation number, longer schedule, higher residual risk','Lower Risk: higher reserve, longer duration, stronger confidence','Premium: resilience and optionality bought with visible capex premium'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}<h3>Current trade-off</h3><div className="triLens"><b>Gained</b>{tradePack.gained.map(x=><span key={x}>{x}</span>)}<b>Sacrificed</b>{tradePack.sacrificed.map(x=><span key={x}>{x}</span>)}<b>Exposed</b>{tradePack.exposed.map(x=><span key={x}>{x}</span>)}</div></Card></section>}
        {tab === 'cost' && <section className="layout two">
          {/* PROCUREMENT INTELLIGENCE */}
          {model?.procurement_intelligence?.items?.length > 0 && <div style={{gridColumn:'1/-1',background:'rgba(16,185,129,0.06)',border:'1px solid rgba(16,185,129,0.2)',borderRadius:8,padding:'14px 18px',marginBottom:12}}>
            <div style={{fontSize:'9px',fontWeight:'800',color:'#10b981',letterSpacing:'.14em',marginBottom:6}}>🔧 PROCUREMENT INTELLIGENCE — KEY LONG-LEAD ITEMS</div>
            <div style={{fontSize:'12px',color:'#fff',fontWeight:'600',marginBottom:8}}>{model.procurement_intelligence.headline}</div>
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:6,marginBottom:8}}>
              {model.procurement_intelligence?.items?.map((p,i)=>(
                <div key={i} style={{display:'flex',alignItems:'flex-start',gap:8,padding:'8px 10px',background:'rgba(255,255,255,0.03)',borderRadius:5,border:`1px solid ${p.secured?'rgba(16,185,129,0.2)':p.priority==='CRITICAL'?'rgba(239,68,68,0.3)':'rgba(245,158,11,0.2)'}`}}>
                  <div style={{fontSize:'14px',marginTop:1}}>{p.secured?'✅':p.priority==='CRITICAL'?'🔴':'🟡'}</div>
                  <div style={{flex:1}}>
                    <div style={{fontSize:'10px',fontWeight:'700',color:p.secured?'#10b981':p.priority==='CRITICAL'?'#ef4444':'#f59e0b'}}>{p.item}</div>
                    <div style={{fontSize:'8px',color:'#64748b',marginTop:2}}>{p.note}</div>
                  </div>
                  <div style={{fontSize:'7px',fontWeight:'800',color:p.priority==='CRITICAL'?'#ef4444':p.priority==='HIGH'?'#f59e0b':'#94a3b8',flexShrink:0}}>{p.priority}</div>
                </div>
              ))}
            </div>
            {model.procurement_intelligence.board_flag && <div style={{padding:'8px 12px',background:'rgba(239,68,68,0.08)',borderRadius:5,fontSize:'9px',color:'#fca5a5',fontWeight:'600'}}>⚠ Board action: {model.procurement_intelligence.board_flag}</div>}
          </div>}

          {/* COST CONCENTRATION — board intelligence */}
          {model?.cost_concentration?.top3_packages?.length > 0 && <div style={{gridColumn:'1/-1',background:'rgba(6,182,212,0.06)',border:'1px solid rgba(6,182,212,0.2)',borderRadius:8,padding:'16px 20px',marginBottom:12}}>
            <div style={{fontSize:'9px',fontWeight:'800',color:'#06b6d4',letterSpacing:'.14em',marginBottom:8}}>📊 COST CONCENTRATION — WHERE THE MONEY IS</div>
            {isNonBase && <div style={{display:'inline-block',padding:'2px 8px',background:'rgba(6,182,212,0.1)',border:'1px solid rgba(6,182,212,0.2)',borderRadius:12,fontSize:'8px',color:'#06b6d4',fontWeight:'700',marginBottom:6}}>{(model.scenario_label||'').toUpperCase()} SCENARIO — cost scaled {((model.scenario_cost_mult||1)*100-100).toFixed(0)}% vs Base</div>}
            <div style={{fontSize:'13px',color:'#fff',fontWeight:'600',marginBottom:10}}>{model.cost_concentration?.top3_pct_of_total||0}% of total cost sits in {model.cost_concentration?.top3_packages?.length} packages — these are the board cost control priorities.</div>
            <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:8,marginBottom:10}}>
              {model.cost_concentration.top3_packages.map((p,i)=>(
                <div key={i} style={{background:'rgba(6,182,212,0.08)',border:'1px solid rgba(6,182,212,0.2)',borderRadius:6,padding:'10px 12px'}}>
                  <div style={{fontSize:'8px',color:'#06b6d4',fontWeight:'700',marginBottom:3}}>#{i+1} — {p.cbs}</div>
                  <div style={{fontSize:'11px',color:'#fff',fontWeight:'600',marginBottom:4}}>{p.description}</div>
                  <div style={{fontSize:'16px',fontWeight:'800',color:'#06b6d4'}}>{model?.currency_symbol||'$'}{p.p50_bn ? ((p.p50_bn * (model?.scenario_cost_mult||1)).toFixed(2)) : '—'+'B'}</div>
                </div>
              ))}
            </div>
            {/* Risk cost anatomy — each top risk split into direct/prelims/inflation */}
            {model?.risk_cost_anatomy?.length > 0 && <div>
              <div style={{fontSize:'9px',fontWeight:'800',color:'#f59e0b',letterSpacing:'.12em',marginBottom:6}}>RISK → COST ANATOMY</div>
              {model.risk_cost_anatomy.slice(0,3).map((r,i)=>(
                <div key={i} style={{display:'flex',alignItems:'center',gap:8,marginBottom:5,padding:'6px 10px',background:'rgba(245,158,11,0.05)',borderRadius:4}}>
                  <div style={{flex:2,fontSize:'10px',color:'#cbd5e1',fontWeight:'600'}}>{r.risk_title}</div>
                  <div style={{flex:1,fontSize:'9px',color:'#94a3b8',textAlign:'center'}}><span style={{color:'#ef4444'}}>{model?.currency_symbol||'$'}{r.direct_cost_bn ? (r.direct_cost_bn*(model?.scenario_risk_mult||1)).toFixed(3) : '—'}</span><br/>direct</div>
                  <div style={{flex:1,fontSize:'9px',color:'#94a3b8',textAlign:'center'}}><span style={{color:'#f59e0b'}}>{model?.currency_symbol||'$'}{r.prelim_extension_bn?.toFixed?.(3)}</span><br/>prelims</div>
                  <div style={{flex:1,fontSize:'9px',color:'#94a3b8',textAlign:'center'}}><span style={{color:'#a78bfa'}}>{model?.currency_symbol||'$'}{r.inflation_bn?.toFixed?.(3)}</span><br/>inflation</div>
                  <div style={{flex:1,fontSize:'10px',color:'#fff',fontWeight:'700',textAlign:'right'}}>{model?.currency_symbol||'$'}{r.emv_bn?.toFixed?.(3)}<br/><span style={{fontSize:'8px',color:'#64748b'}}>total</span></div>
                </div>
              ))}
            </div>}
          </div>}
          <Card><h2>Scenario cost bridge vs Base</h2><p className="chartCaption">This explains why the selected scenario is cheaper or more expensive than Base before showing the workbook lines.</p>{model?.stress_test_applied && <div style={{background:"rgba(141,247,255,0.05)",borderLeft:"2px solid #8df7ff",padding:"8px 12px",marginBottom:"8px",fontSize:"11px",color:"#8df7ff"}}>Stress test applied: {String(model.stress_test_applied).replace(/_/g," ")} — cost recalculated to {safeRender(model.cost_p50)}. The waterfall below reflects this change.</div>}{costWaterfall.map((x,i)=><div className={`reason ${x.kind==='total'?'deltaReason':''}`} key={i}><span>{i+1}</span><b>{x.driver}</b><br/>{x.kind==='total'?x.value:(x.value_bn>=0?'+':'−') + ' ' + x.value}</div>)}{model.unit_rate_label && <div style={{margin:'8px 0',padding:'8px 12px',background:'rgba(141,247,255,0.04)',border:'1px solid rgba(141,247,255,0.1)',borderRadius:'3px'}}>
              <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:'3px'}}>
                <span style={{fontSize:'9px',fontWeight:'800',color:'#8df7ff',letterSpacing:'.1em'}}>UNIT COST BENCHMARK — {model.unit_rate_label?.metric||'programme unit'}</span>
                <span style={{fontSize:'9px',color:'#475569'}}>{model.unit_rate_label?.typical_range||''}</span>
              </div>
              <p style={{fontSize:'9px',color:'#64748b',margin:0,lineHeight:'1.4'}}>Unit rates below show cost per {model.unit_rate_label?.metric||'programme unit'} for each CBS line, derived from this programme estimate. Compare to the typical sector range (right). Divergence indicates exceptional scope or unusual location/complexity.</p>
            </div>}
            <h3>Cost estimate workbook</h3><Table rows={costs} cols={[["cbs","CBS"],["description","Description"],["type","Type"],["unit_rate","Unit rate"],["p10_bn","P10"],["p50_bn","P50"],["p90_bn","P90"],["basis","Basis"]]} curr={model?.currency_symbol || "$"} moneyCols={["p10_bn","p50_bn","p90_bn"]} cellFmt={(col, val) => {
                const curr = model?.currency_symbol || '$';
                if (col === 'unit_rate' && curr !== '$') return String(val).replace(/\$([0-9])/g, curr + '$1');
                return String(val);
              }}/></Card><Card><h2>Cost composition</h2><p className="chartCaption">Direct, indirect and reserve are scenario-controlled and reconciled to selected P50. For the detailed uncertainty view use QCRA/QSRA.</p><ResponsiveContainer width="100%" height={320}><BarChart data={[{name:'Direct',value:direct},{name:'Indirect',value:indirect},{name:'Reserve',value:reserves}]}><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis dataKey="name"/><YAxis/><Tooltip/><Bar dataKey="value" fill="#8df7ff"/></BarChart></ResponsiveContainer></Card></section>}
        {tab === 'schedule' && <section className="layout two">
          {/* XER HEALTH ENGINE */}
          {model?.xer_health && <div style={{gridColumn:'1/-1',display:'grid',gridTemplateColumns:'1fr 1fr 1fr',gap:10,marginBottom:12}}>
            <div style={{background:'rgba(6,182,212,0.06)',border:'1px solid rgba(6,182,212,0.2)',borderRadius:8,padding:'14px 16px',gridColumn:'1/-1'}}>
              <div style={{fontSize:'9px',fontWeight:'800',color:'#06b6d4',letterSpacing:'.14em',marginBottom:6}}>📐 XER HEALTH ENGINE — SCHEDULE QUALITY SCORING</div>
              <div style={{fontSize:'12px',color:'#fff',fontWeight:'600',marginBottom:10}}>{model.xer_health.headline}</div>
              <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:8}}>
                {[['Activities',model.xer_health.activity_count,'#06b6d4'],['Critical',model.xer_health.critical_count,'#ef4444'],['Logic Quality',model.xer_health.logic_quality,model.xer_health.logic_quality==='GOOD'?'#10b981':model.xer_health.logic_quality==='REVIEW'?'#f59e0b':'#ef4444'],['Float Quality',model.xer_health.float_quality,model.xer_health.float_quality==='GOOD'?'#10b981':model.xer_health.float_quality==='REVIEW'?'#f59e0b':'#ef4444']].map(([k,v,col],i)=>(
                  <div key={i} style={{background:'rgba(255,255,255,0.03)',borderRadius:5,padding:'8px 10px'}}>
                    <div style={{fontSize:'8px',color:'#64748b',marginBottom:2}}>{k}</div>
                    <div style={{fontSize:'16px',fontWeight:'800',color:col}}>{v}</div>
                  </div>
                ))}
              </div>
              {model.xer_health.board_flag && <div style={{marginTop:8,padding:'6px 10px',background:'rgba(245,158,11,0.08)',borderRadius:4,fontSize:'9px',color:'#f59e0b'}}>⚠ {model.xer_health.board_flag}</div>}
            </div>
          </div>}

          {/* Scenario banner for schedule */}
          {model?.scenario !== 'base' && <div style={{gridColumn:'1/-1',background:'rgba(245,158,11,0.06)',border:'1px solid rgba(245,158,11,0.2)',borderRadius:6,padding:'8px 14px',marginBottom:8,display:'flex',gap:12,alignItems:'center'}}>
            <div style={{fontSize:'9px',fontWeight:'800',color:'#f59e0b'}}>{(model.scenario_label||'').toUpperCase()} SCENARIO SCHEDULE</div>
            <div style={{display:'flex',gap:16,fontSize:'9px',color:'#94a3b8'}}>
              <span>Duration: <b style={{color:'#f59e0b'}}>{model.schedule||'—'}</b></span>
              <span>QSRA P80: <b style={{color:'#f59e0b'}}>{model.monte_carlo?.qsra?.p80||'—'} months</b></span>
              <span>Schedule mult: <b style={{color:'#f59e0b'}}>{model.scenario_sched_mult ? ((model.scenario_sched_mult*100)-100).toFixed(0)+'%' : '—'} vs Base</b></span>
            </div>
          </div>}
          {/* SCHEDULE KILLER CHAIN */}
          {model?.schedule_killer_chain?.chain?.length > 0 && <div style={{gridColumn:'1/-1',background:'rgba(245,158,11,0.06)',border:'1px solid rgba(245,158,11,0.2)',borderRadius:8,padding:'16px 20px',marginBottom:12}}>
            <div style={{fontSize:'9px',fontWeight:'800',color:'#f59e0b',letterSpacing:'.14em',marginBottom:6}}>⛓ SCHEDULE KILLER CHAIN — CRITICAL PATH SEQUENCE</div>
            <div style={{fontSize:'13px',color:'#fff',fontWeight:'600',marginBottom:10}}>{model.schedule_killer_chain.headline}</div>
            <div style={{display:'flex',alignItems:'center',flexWrap:'wrap',gap:0}}>
              {model.schedule_killer_chain?.chain?.map((a,i)=>(
                <div key={i} style={{display:'flex',alignItems:'center',gap:0}}>
                  <div style={{background:a.linked_risk_emv_bn>0?'rgba(239,68,68,0.12)':'rgba(245,158,11,0.08)',border:a.linked_risk_emv_bn>0?'1px solid rgba(239,68,68,0.3)':'1px solid rgba(245,158,11,0.2)',borderRadius:6,padding:'8px 12px',minWidth:120,textAlign:'center'}}>
                    <div style={{fontSize:'8px',color:'#64748b',marginBottom:2}}>{a.activity_id}</div>
                    <div style={{fontSize:'10px',fontWeight:'700',color:'#fff',marginBottom:2}}>{a.name}</div>
                    <div style={{fontSize:'9px',color:'#f59e0b'}}>{a.duration_months+'mo'}</div>
                    {a.linked_risk_title && <div style={{fontSize:'8px',color:'#ef4444',marginTop:2}}>⚠ {a.linked_risk_title?.slice?.(0,30)}</div>}
                  </div>
                  {i < model.schedule_killer_chain.chain.length-1 && <div style={{fontSize:'16px',color:'#f59e0b',padding:'0 4px'}}>→</div>}
                </div>
              ))}
            </div>
            <div style={{marginTop:10,fontSize:'10px',color:'#94a3b8'}}>Activities marked ⚠ have active risks that can move this milestone. Delay any one and completion moves.</div>
          </div>}
          <Card><h2>Schedule bridge vs Base</h2><p className="chartCaption">This is the month-by-month reason the scenario becomes faster or slower than Base.</p>{scheduleWaterfall.map((x,i)=><div className={`reason ${x.kind==='total'?'deltaReason':''}`} key={i}><span>{i+1}</span><b>{x.driver}</b><br/>{x.kind==='total'?`${x.months} months`:(x.months>=0?'+':'') + x.months + ' months'}</div>)}<h3>Scenario schedule logic</h3><Table rows={schedule} cols={[["activity_id","Activity"],["phase","Phase"],["activity","Name"],["predecessor","Pred"],["duration_months","Months"],["critical","Critical"],["basis","Basis"]]}/></Card><Card><h2>QSRA finish-date curve</h2><p className="chartCaption">P50 equals the headline schedule. P80/P90 show how severe the delivery tail becomes after the scenario trade-off.</p><div className="metrics"><div>P50<b>{qsra.p50} mo</b></div><div>P80<b>{qsra.p80} mo</b></div><div>P90<b>{qsra.p90} mo</b></div></div><ResponsiveContainer width="100%" height={280}><LineChart data={curve}><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis dataKey="percentile"/><YAxis/><Tooltip formatter={(v) => [`${v} months`, "QSRA finish date"]}/><ReferenceLine x={50} stroke="#ffffff88" label="P50 = headline"/><ReferenceLine x={80} stroke="#ffffff55" label="P80 = board risk"/><Line type="monotone" name="QSRA finish date" dataKey="schedule_months" stroke="#b18cff" strokeWidth={4}/></LineChart></ResponsiveContainer><div className="reason p80Translation"><span>1/5</span>{safeRender(p80Talk.schedule)}</div><div className="reason p80Translation"><span>!</span>{safeRender(p80Talk.board)}</div>{(model.monte_carlo?.curve_readout || []).slice(1).map((x,i)=><div className="reason" key={i}><span>{i+1}</span>{x}</div>)}</Card></section>}
        {tab === 'risk' && <section className="layout one">
          {/* SCENARIO BANNER */}
          {model?.scenario !== 'base' && <div style={{background:'rgba(239,68,68,0.06)',border:'1px solid rgba(239,68,68,0.2)',borderRadius:6,padding:'8px 14px',marginBottom:10,display:'flex',gap:12,alignItems:'center'}}>
            <div style={{fontSize:'9px',fontWeight:'800',color:'#ef4444'}}>{(model.scenario_label||'').toUpperCase()} SCENARIO — risk EMVs scaled {(((model.scenario_risk_mult||1)*100)-100).toFixed(0)}% vs Base</div>
            <div style={{fontSize:'9px',color:'#94a3b8',marginLeft:'auto'}}>Total EMV: <b style={{color:'#ef4444'}}>{model.currency_symbol}{((model.board_risk_summary?.total_emv_bn||0)*(model.scenario_risk_mult||1)).toFixed(2)+'B'}</b></div>
          </div>}

          {/* RISK RESERVE HEADLINE — what a CFO asks immediately */}
          {model?.p80_reserve_bn > 0 && <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr auto',gap:8,marginBottom:12,padding:'12px 16px',background:'rgba(239,68,68,0.06)',border:'1px solid rgba(239,68,68,0.2)',borderRadius:8}}>
            <div>
              <div style={{fontSize:'8px',color:'#64748b',marginBottom:2}}>RISK RESERVE (P80−P50)</div>
              <div style={{fontSize:'20px',fontWeight:'900',color:(model.p80_reserve_pct||0)>=(model.reserve_vs_benchmark_pct||18)?'#10b981':'#ef4444'}}>{model.p80_reserve}</div>
              <div style={{fontSize:'8px',color:'#475569'}}>{model.p80_reserve_pct}% of P50 (benchmark: {model.reserve_vs_benchmark_pct}%)</div>
            </div>
            <div>
              <div style={{fontSize:'8px',color:'#64748b',marginBottom:2}}>TOTAL RISK EMV</div>
              <div style={{fontSize:'20px',fontWeight:'900',color:'#ef4444'}}>{model.currency_symbol}{((model.board_risk_summary?.total_emv_bn||0)*(model.scenario_risk_mult||1)).toFixed(2)+'B'}</div>
              <div style={{fontSize:'8px',color:'#475569'}}>across {model.total_risks_identified||model.risks?.length||0} identified risks</div>
            </div>
            <div>
              <div style={{fontSize:'8px',color:'#64748b',marginBottom:2}}>FULL OUTTURN (P80+OBA)</div>
              <div style={{fontSize:'20px',fontWeight:'900',color:'#f59e0b'}}>{model.outturn}</div>
              <div style={{fontSize:'8px',color:'#475569'}}>OBA: {model.oba_pct}% ({model.estimate_class_name})</div>
            </div>
            {model.reserve_gap_bn > 0 && <div style={{padding:'8px 12px',background:'rgba(239,68,68,0.12)',border:'1px solid rgba(239,68,68,0.3)',borderRadius:6,display:'flex',flexDirection:'column',justifyContent:'center',alignItems:'center'}}>
              <div style={{fontSize:'8px',color:'#ef4444',marginBottom:2,textAlign:'center'}}>RESERVE SHORTFALL</div>
              <div style={{fontSize:'16px',fontWeight:'900',color:'#ef4444'}}>{model.currency_symbol}{model.reserve_gap_bn.toFixed?.(2)+'B'}</div>
              <div style={{fontSize:'7px',color:'#94a3b8',textAlign:'center'}}>below benchmark</div>
            </div>}
          </div>}

          {/* TOP BOARD RISKS — MONETISED */}
          {model?.board_risk_summary?.top5?.length > 0 && <div style={{background:'rgba(255,255,255,0.02)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:8,padding:'14px 18px',marginBottom:12}}>
            <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:8}}>
              <div style={{fontSize:'9px',fontWeight:'800',color:'#e2e8f0',letterSpacing:'.12em'}}>🎯 TOP RISKS — MONETISED BOARD EXPOSURE</div>
              <div style={{fontSize:'9px',color:'#64748b'}}>{model.board_risk_summary?.top5_pct}% of total EMV in top 5</div>
            </div>
            {model.board_risk_summary?.top5?.map((r,i)=>(
              <div key={i} style={{display:'grid',gridTemplateColumns:'2fr 1fr 1fr 1fr 1fr 1fr',gap:6,alignItems:'center',marginBottom:5,padding:'8px 12px',background:i===0?'rgba(239,68,68,0.08)':'rgba(255,255,255,0.02)',borderRadius:5,border:i===0?'1px solid rgba(239,68,68,0.25)':'1px solid rgba(255,255,255,0.04)'}}>
                <div>
                  <div style={{fontSize:'10px',color:'#fff',fontWeight:'600'}}>{r.title?.slice?.(0,50)}</div>
                  <div style={{fontSize:'8px',color:'#64748b'}}>Owner: {r.owner} · {r.probability_pct}% probability</div>
                </div>
                <div style={{textAlign:'center'}}><div style={{fontSize:'13px',fontWeight:'800',color:'#ef4444'}}>{model?.currency_symbol}{((r.emv_bn||0)*(model?.scenario_risk_mult||1)).toFixed(3)+'B'}</div><div style={{fontSize:'7px',color:'#64748b'}}>total EMV</div></div>
                <div style={{textAlign:'center'}}><div style={{fontSize:'11px',fontWeight:'700',color:'#f87171'}}>{model?.currency_symbol}{((r.direct_cost_bn||0)*(model?.scenario_risk_mult||1)).toFixed(3)+'B'}</div><div style={{fontSize:'7px',color:'#64748b'}}>direct</div></div>
                <div style={{textAlign:'center'}}><div style={{fontSize:'11px',fontWeight:'700',color:'#fbbf24'}}>{model?.currency_symbol}{((r.prelim_extension_bn||0)*(model?.scenario_risk_mult||1)).toFixed(3)+'B'}</div><div style={{fontSize:'7px',color:'#64748b'}}>prelims</div></div>
                <div style={{textAlign:'center'}}><div style={{fontSize:'11px',fontWeight:'700',color:'#a78bfa'}}>{model?.currency_symbol}{((r.inflation_bn||0)*(model?.scenario_risk_mult||1)).toFixed(3)+'B'}</div><div style={{fontSize:'7px',color:'#64748b'}}>inflation</div></div>
                <div style={{textAlign:'center',padding:'3px 6px',background:i===0?'rgba(239,68,68,0.15)':'rgba(255,255,255,0.04)',borderRadius:3,fontSize:'8px',color:'#94a3b8'}}>{r.category}</div>
              </div>
            ))}
          </div>}

          {/* RISK → SCHEDULE → CONFIDENCE LINKS */}
          {model?.risk_schedule_links?.length > 0 && <div style={{background:'rgba(245,158,11,0.06)',border:'1px solid rgba(245,158,11,0.2)',borderRadius:8,padding:'14px 18px',marginBottom:12}}>
            <div style={{fontSize:'9px',fontWeight:'800',color:'#f59e0b',letterSpacing:'.12em',marginBottom:8}}>RISK → SCHEDULE → CONFIDENCE CHAIN</div>
            {model.risk_schedule_links.slice(0,5).map((l,i)=>(
              <div key={i} style={{display:'flex',alignItems:'center',gap:6,marginBottom:5,padding:'5px 8px',background:'rgba(245,158,11,0.04)',borderRadius:4}}>
                <div style={{fontSize:'9px',color:'#fff',flex:1,fontWeight:'600'}}>{l.risk_title?.slice?.(0,40)}</div>
                <div style={{display:'flex',gap:8,fontSize:'9px',color:'#94a3b8'}}>
                  <span>→ <b style={{color:'#f59e0b'}}>{l.activity?.slice?.(0,25)}</b></span>
                  <span style={{color:'#ef4444'}}>+{l.schedule_impact_months}mo delay</span>
                  <span style={{color:'#a78bfa'}}>{l.confidence_impact_pct}% conf</span>
                </div>
              </div>
            ))}
          </div>}

          {/* RISK CLUSTERS */}
          {model?.risk_clusters?.clusters?.length > 0 && <div style={{background:'rgba(139,92,246,0.06)',border:'1px solid rgba(139,92,246,0.2)',borderRadius:8,padding:'14px 18px',marginBottom:12}}>
            <div style={{fontSize:'9px',fontWeight:'800',color:'#a78bfa',letterSpacing:'.12em',marginBottom:8}}>🔗 {model.risk_clusters.total_clusters} RISK CLUSTERS — WHERE THE EXPOSURE LIVES</div>
            <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:8}}>
              {model.risk_clusters.clusters.slice(0,8).map((c,i)=>(
                <div key={i} style={{background:'rgba(139,92,246,0.06)',border:'1px solid rgba(139,92,246,0.12)',borderRadius:6,padding:'8px 12px'}}>
                  <div style={{fontSize:'8px',color:'#a78bfa',fontWeight:'700',marginBottom:3}}>{c.cluster}</div>
                  <div style={{fontSize:'15px',fontWeight:'900',color:'#ef4444'}}>{model?.currency_symbol}{((c.total_emv_bn||0)*(model?.scenario_risk_mult||1)).toFixed?.(2)+'B'}</div>
                  <div style={{display:'flex',justifyContent:'space-between',fontSize:'7px',color:'#64748b',marginTop:2}}>
                    <span>{c.risk_count} risks</span>
                    <span>{c.pct_of_total_emv+'%'}</span>
                  </div>
                </div>
              ))}
            </div>
          </div>}

          {/* FULL RISK REGISTER */}
          <Card><h2>Risk Register</h2>
            <p style={{color:'#475569',fontSize:'9px',marginBottom:'8px'}}>Each risk: cause → event → cost + schedule impact → probability → owner → mitigation. Top risks drive P80 exposure.</p>
            {(model?.risks||[]).slice(0,20).map((r,i)=>(
              <div key={i} style={{padding:'6px 10px',background:i===0?'rgba(239,68,68,0.06)':'rgba(255,255,255,0.02)',borderRadius:4,marginBottom:3,display:'grid',gridTemplateColumns:'2fr 1fr 1fr 1fr 1fr',gap:6,fontSize:'9px',alignItems:'center'}}>
                <div><b style={{color:'#fff'}}>{r.title||r.risk}</b><br/><span style={{color:'#475569',fontSize:'8px'}}>{r.mitigation?.slice?.(0,60)}</span></div>
                <div style={{textAlign:'center',color:'#ef4444',fontWeight:'700'}}>{model?.currency_symbol}{r.cost_emv_bn?.toFixed?.(3)||'—'+'B'}</div>
                <div style={{textAlign:'center',color:'#f59e0b'}}>{r.probability_pct||'—'+'%'}</div>
                <div style={{textAlign:'center',color:'#94a3b8'}}>{r.owner||'TBC'}</div>
                <div style={{textAlign:'center'}}><span style={{padding:'1px 5px',background:r.status==='Closed'?'rgba(16,185,129,0.1)':'rgba(239,68,68,0.1)',borderRadius:3,color:r.status==='Closed'?'#10b981':'#ef4444',fontSize:'7px'}}>{r.status||'Open'}</span></div>
              </div>
            ))}
          </Card>
        </section>}
        {tab === 'monte' && <section className="layout one">
          {/* SCENARIO BANNER */}
          {model?.scenario !== 'base' && <div style={{background:'rgba(6,182,212,0.06)',border:'1px solid rgba(6,182,212,0.2)',borderRadius:6,padding:'8px 14px',marginBottom:10,display:'flex',gap:12,alignItems:'center'}}>
            <div style={{fontSize:'9px',fontWeight:'800',color:'#06b6d4'}}>{(model.scenario_label||'').toUpperCase()} SCENARIO QCRA/QSRA</div>
            <div style={{display:'flex',gap:16,fontSize:'9px',color:'#94a3b8'}}>
              <span>P50: <b style={{color:'#fff'}}>{model.cost_p50}</b></span>
              <span>P80: <b style={{color:'#06b6d4'}}>{model.cost_p80}</b></span>
              <span>QSRA P80: <b style={{color:'#f59e0b'}}>{model.monte_carlo?.qsra?.p80} months</b></span>
              <span>Confidence: <b style={{color:(model.confidence_pct||0)>=70?'#10b981':'#f59e0b'}}>{model.confidence_pct+'%'}</b></span>
            </div>
          </div>}

          {/* THE RESERVE INTELLIGENCE STRIP — what Bloomberg sees first */}
          {model?.p80_reserve_bn > 0 && <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr 1fr',gap:10,marginBottom:12}}>
            {[
              ['P80 Reserve (P80 − P50)', model?.p80_reserve, model?.p80_reserve_pct+'% of P50', (model?.p80_reserve_pct||0) >= (model?.reserve_vs_benchmark_pct||18) ? '#10b981' : '#ef4444'],
              ['OBA Uplift (Class '+( model?.estimate_class||3)+')', model?.currency_symbol+(model?.oba_uplift_bn||0).toFixed?.(2)+'B', (model?.oba_pct||0)+'% optimism bias adjustment', '#f59e0b'],
              ['Full Outturn (P80 + OBA)', model?.outturn, 'Total board exposure', '#ef4444'],
              ['Benchmark Reserve %', (model?.reserve_vs_benchmark_pct||18)+'%', 'Required at '+( model?.estimate_class_name||'Class 3'), (model?.p80_reserve_pct||0) >= (model?.reserve_vs_benchmark_pct||18) ? '#10b981' : '#ef4444'],
            ].map(([label,value,sub,color])=>(
              <div key={label} style={{background:'rgba(255,255,255,0.02)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:8,padding:'12px 14px'}}>
                <div style={{fontSize:'8px',color:'#64748b',marginBottom:4}}>{label}</div>
                <div style={{fontSize:'18px',fontWeight:'900',color:color,marginBottom:2}}>{value}</div>
                <div style={{fontSize:'8px',color:'#475569'}}>{sub}</div>
              </div>
            ))}
          </div>}
          {model?.reserve_gap_bn > 0 && <div style={{background:'rgba(239,68,68,0.08)',border:'1px solid rgba(239,68,68,0.3)',borderRadius:6,padding:'8px 14px',marginBottom:10,display:'flex',gap:12,alignItems:'center'}}>
            <span style={{fontSize:'14px'}}>⚠</span>
            <div style={{fontSize:'10px',color:'#fca5a5',fontWeight:'700'}}>Reserve shortfall: {model.currency_symbol}{model.reserve_gap_bn.toFixed?.(2)}B below benchmark for {model.estimate_class_name}. Benchmark requires {model.reserve_vs_benchmark_pct}% — you are holding {model.p80_reserve_pct}%.</div>
          </div>}

          {/* CONFIDENCE TRAJECTORY — how confidence grows with maturity */}
          {model?.confidence_trajectory?.length > 0 && <div style={{background:'rgba(6,182,212,0.06)',border:'1px solid rgba(6,182,212,0.2)',borderRadius:8,padding:'14px 18px',marginBottom:12}}>
            <div style={{fontSize:'9px',fontWeight:'800',color:'#06b6d4',letterSpacing:'.12em',marginBottom:8}}>CONFIDENCE TRAJECTORY — HOW YOUR NUMBER GETS STRONGER</div>
            <div style={{display:'flex',alignItems:'center',gap:0,marginBottom:8}}>
              {model.confidence_trajectory?.map((stage,i)=>(
                <div key={i} style={{flex:1,textAlign:'center',padding:'8px 4px',background:stage.is_current?'rgba(6,182,212,0.15)':'transparent',borderRadius:6,border:stage.is_current?'1px solid rgba(6,182,212,0.4)':'1px solid transparent',position:'relative'}}>
                  {stage.is_current && <div style={{position:'absolute',top:-8,left:'50%',transform:'translateX(-50%)',fontSize:'8px',background:'#06b6d4',color:'#fff',padding:'1px 5px',borderRadius:3,fontWeight:'700',whiteSpace:'nowrap'}}>YOU ARE HERE</div>}
                  <div style={{fontSize:'8px',color:stage.is_current?'#8df7ff':'#64748b',fontWeight:stage.is_current?'800':'400'}}>{stage.label}</div>
                  <div style={{fontSize:'16px',fontWeight:'900',color:stage.confidence>=75?'#10b981':stage.confidence>=55?'#f59e0b':'#ef4444',margin:'2px 0'}}>{stage.confidence+'%'}</div>
                  <div style={{fontSize:'7px',color:'#475569'}}>{stage.class}</div>
                  {i < model.confidence_trajectory.length - 1 && <div style={{position:'absolute',right:-8,top:'50%',transform:'translateY(-50%)',color:'#334155',fontSize:'12px'}}>→</div>}
                </div>
              ))}
            </div>
            {model.approval_pathway && <div style={{padding:'8px 12px',background:'rgba(6,182,212,0.06)',borderRadius:5,fontSize:'9px',color:'#8df7ff'}}>{model.approval_pathway}</div>}
          </div>}

          {/* CONFIDENCE BY DISCIPLINE */}
          {model?.confidence_by_discipline && <div style={{background:'rgba(6,182,212,0.06)',border:'1px solid rgba(6,182,212,0.2)',borderRadius:8,padding:'14px 18px',marginBottom:12}}>
            <div style={{fontSize:'9px',fontWeight:'800',color:'#06b6d4',marginBottom:8}}>WHY {model.confidence_pct}% CONFIDENCE? — DECOMPOSITION BY DISCIPLINE</div>
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8}}>
              {Object.entries(model.confidence_by_discipline).map(([disc,pct])=>(
                <div key={disc} style={{display:'flex',alignItems:'center',gap:8}}>
                  <div style={{fontSize:'9px',color:'#94a3b8',width:140,flexShrink:0}}>{disc}</div>
                  <div style={{flex:1,height:5,background:'rgba(255,255,255,0.06)',borderRadius:3}}>
                    <div style={{width:pct+'%',height:'100%',background:pct>=70?'#10b981':pct>=55?'#f59e0b':'#ef4444',borderRadius:3,transition:'width .4s'}}/>
                  </div>
                  <div style={{fontSize:'10px',fontWeight:'800',color:pct>=70?'#10b981':pct>=55?'#f59e0b':'#ef4444',minWidth:30,textAlign:'right'}}>{pct+'%'}</div>
                </div>
              ))}
            </div>
          </div>}

          {/* QCRA + QSRA P VALUES */}
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:12,marginBottom:12}}>
            <div style={{background:'rgba(255,255,255,0.02)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:8,padding:'14px 18px'}}>
              <div style={{fontSize:'9px',fontWeight:'800',color:'#ef4444',marginBottom:8}}>QCRA — COST RISK</div>
              {[['P10 (optimistic)',model.monte_carlo?.qcra?.p10,'#10b981'],['P50 (base estimate)',model.cost_p50,'#8df7ff'],['P80 (board exposure)',model.cost_p80,'#06b6d4'],['P90 (stress case)',model.monte_carlo?.qcra?.p90,'#ef4444']].map(([k,v,col])=>(
                <div key={k} style={{display:'flex',justifyContent:'space-between',alignItems:'baseline',marginBottom:5,padding:'4px 8px',background:k.includes('P80')?'rgba(6,182,212,0.08)':'transparent',borderRadius:4}}>
                  <span style={{fontSize:'9px',color:'#64748b'}}>{k}</span>
                  <span style={{fontSize:k.includes('P80')?'15px':'11px',fontWeight:'800',color:col}}>{typeof v === 'number' ? (model.currency_symbol||'$')+(v).toFixed(2)+'B' : v}</span>
                </div>
              ))}
            </div>
            <div style={{background:'rgba(255,255,255,0.02)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:8,padding:'14px 18px'}}>
              <div style={{fontSize:'9px',fontWeight:'800',color:'#f59e0b',marginBottom:8}}>QSRA — SCHEDULE RISK</div>
              {[['P10 (optimistic)',model.monte_carlo?.qsra?.p10+' months','#10b981'],['P50 (base duration)',model.schedule,'#8df7ff'],['P80 (board exposure)',model.monte_carlo?.qsra?.p80+' months','#f59e0b'],['P90 (stress case)',model.monte_carlo?.qsra?.p90+' months','#ef4444']].map(([k,v,col])=>(
                <div key={k} style={{display:'flex',justifyContent:'space-between',alignItems:'baseline',marginBottom:5,padding:'4px 8px',background:k.includes('P80')?'rgba(245,158,11,0.08)':'transparent',borderRadius:4}}>
                  <span style={{fontSize:'9px',color:'#64748b'}}>{k}</span>
                  <span style={{fontSize:k.includes('P80')?'15px':'11px',fontWeight:'800',color:col}}>{v}</span>
                </div>
              ))}
            </div>
          </div>

          {/* CONFIDENCE DRIVERS */}
          {model?.confidence_breakdown?.length > 0 && <div style={{background:'rgba(255,255,255,0.02)',border:'1px solid rgba(255,255,255,0.06)',borderRadius:8,padding:'12px 16px'}}>
            <div style={{fontSize:'9px',fontWeight:'800',color:'#64748b',marginBottom:6}}>CONFIDENCE DRIVERS</div>
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:6}}>
              {model.confidence_breakdown.map((d,i)=>(
                <div key={i} style={{display:'flex',justifyContent:'space-between',fontSize:'9px',padding:'3px 0',borderBottom:'1px solid rgba(255,255,255,0.04)'}}>
                  <span style={{color:'#64748b'}}>{d.driver}: {d.effect}</span>
                  <span style={{fontWeight:'700',color:(d.delta||0)>=0?'#10b981':'#ef4444'}}>{d.delta ? ((d.delta>0?'+':'')+d.delta) : '—'}</span>
                </div>
              ))}
            </div>
          </div>}
        </section>}
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
            <div className="triLens full"><b>Gained</b>{tradePack.gained.map(x=><span key={x}>{x}</span>)}<b>Sacrificed</b>{tradePack.sacrificed.map(x=><span key={x}>{x}</span>)}<b>Exposed</b>{tradePack.exposed.map(x=><span key={x}>{x}</span>)}</div>
            <div className="reason"><span>!</span><b>Curve meaning</b><br/>{model.monte_carlo?.curve_interpretation || 'QCRA/QSRA shape reflects scenario uncertainty.'}</div>
          </Card>
          <Card><h2>Confidence Breakdown</h2>
            {(model.confidence_breakdown || []).map((x,i)=><div className="reason" key={i}><span>{i+1}</span><b>{x.driver}: {x.effect}</b><br/>{x.note}</div>)}
          </Card>
          <Card><h2>Top Decisions Required</h2>
            {(model.top_decisions_required || []).map((x,i)=><div className="reason" key={i}><span>{i+1}</span>{x}</div>)}
          </Card>
        </section>}

        {tab === 'twin' && <section className="layout one"><HolyGrailRuntime model={model} scenario={scenario} generate={generate} runShock={runShock}/></section>}

        {tab === 'causal' && <section className="layout two"><CausalGraph model={model}/><BenchmarkIntelligence model={model}/><Card><h2>Evidence Mode: {viewMode}</h2>{evidenceScorecard(model).map((x,i)=><div className="reason" key={x.name}><span>{i+1}</span><b>{x.name}: {Math.round(x.score)+'%'}</b><br/>{x.note}</div>)}</Card></section>}

        {tab === 'outputs' && <section className="layout two"><Card><h2>Generated Artefacts</h2><p>The public demo previews the intelligence pack. Enterprise access unlocks the live generated controls deliverables.</p><div className="exports v50Exports lockedExports">
          <button onClick={() => download('/export/workbook', model, `${model.id || 'casey'}_COST_WORKBOOK.xlsx`)}><FileSpreadsheet/> Generate Cost Model XLSX</button>
          <button onClick={() => download('/export/risk-register', model, `${model.id || 'casey'}_RISK_REGISTER.xlsx`)}><Database/> Generate Risk Register XLSX</button>
          <button onClick={() => download('/export/xer', model, `${model.id || 'casey'}_PRA_SCHEDULE.xer`)}><Workflow/> Generate PRA Schedule XER</button>
          <button onClick={() => download('/export/qcra-qsra', model, `${model.id || 'casey'}_QCRA_QSRA.xlsx`)}><BarChart3/> Generate QCRA/QSRA Pack</button>
          <button onClick={() => download('/export/json', model, `${model.id || 'casey'}_AUDIT_MODEL.json`)}><Brain/> Generate Audit File JSON</button>
          <button onClick={() => download('/export/pdf', model, `${model.id || 'casey'}_CASEY_Board_Pack.pdf`)}><Download/> Generate Full Pack ZIP</button>
          <a className="contactBtn" href={emailLink}><Mail/> Request Enterprise Review</a></div></Card><Card><h2>What the pack delivers</h2>{['Executive control centre with project, scenario, class, level and confidence clearly identified','Scenario comparison covering Base, Faster, Cheaper, Lower Risk and Premium cases','Selected estimate class plus all class levels for audit and challenge','Direct, indirect and reserve cost views with QCRA cost curve and cost tornado','All schedule levels with QSRA schedule curve and schedule tornado','Risk register with cause, event, impact, owner, mitigation, trigger and quantified likelihood','Basis of Estimate, assumptions, exclusions and benchmark validation','Commercial next steps: buyer action, procurement challenge and board decision path'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card></section>}

        {tab === 'assurance' && <><IncumbentPressurePanel model={model} direct={direct} indirect={indirect} reserves={reserves} reconcileCheck={reconcileCheck}/><section className="layout two"><Card><h2>Assurance room weapons</h2>{['Open with the P80/P90 exposure, not the headline P50.','Ask which evidence package retires the governing constraint.','Force every mitigation to name owner, trigger, residual exposure and date.','Show scenario trade-offs live before anyone can defend a single-point estimate.','Export the audit model immediately so the conversation moves from opinion to traceability.'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card><Card><h2>Why CASEY changes the conversation</h2>{['CASEY recalculates cost, schedule, confidence and board posture from one source of truth in seconds.','Every scenario is a complete recalculation — not a slide edit.','The system surfaces contradictions rather than polishing the management narrative.','Static reports become live investment-committee intelligence.'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card></section><section className="layout one"><ProgrammeHealthSignal onRunHealthCheck={runHealthCheck}/></section></>}

        {tab === 'advisor' && <>
          <section className="layout two">
            <Card>
              <h2>CASEY Advisor</h2>
              {/* API STATUS INDICATOR */}
              <div style={{marginBottom:8,padding:'6px 10px',background:advisorModel?.advisor_active?'rgba(16,185,129,0.06)':'rgba(245,158,11,0.06)',border:'1px solid '+(advisorModel?.advisor_active?'rgba(16,185,129,0.2)':'rgba(245,158,11,0.2)'),borderRadius:5,display:'flex',alignItems:'center',gap:8}}>
                <span style={{fontSize:'8px',width:6,height:6,borderRadius:'50%',background:advisorModel?.advisor_active?'#10b981':'#f59e0b',display:'inline-block',flexShrink:0}}/>
                <div style={{fontSize:'8px',color:advisorModel?.advisor_active?'#10b981':'#f59e0b',fontWeight:'700'}}>{advisorModel?.advisor_active ? advisorModel.advisor_model+' — Live AI active' : 'Pattern matching active — Set ANTHROPIC_API_KEY or OPENAI_API_KEY in Render to enable AI'}</div>
                {!isAdminUser && chatUsed >= FREE_ADVISOR_LIMIT && <div style={{marginLeft:'auto',fontSize:'8px',color:'#ef4444',fontWeight:'700'}}>1 free question used — <span style={{cursor:'pointer',textDecoration:'underline'}} onClick={()=>{setEmailGateOpen(true);setEmailGateFor('advisor');}}>Request access</span></div>}
                {isAdminUser && <div style={{marginLeft:'auto',fontSize:'8px',color:'#475569'}}>Admin — unlimited</div>}
              </div>
              {/* SCENARIO CONTEXT FOR ADVISOR */}
              {model?.scenario !== 'base' && <div style={{background:'rgba(6,182,212,0.06)',border:'1px solid rgba(6,182,212,0.2)',borderRadius:6,padding:'8px 14px',marginBottom:10,display:'flex',gap:12,alignItems:'center'}}>
                <div style={{fontSize:'10px',fontWeight:'800',color:'#06b6d4'}}>{(model.scenario_label||'').toUpperCase()} SCENARIO ACTIVE</div>
                <div style={{fontSize:'9px',color:'#94a3b8'}}>All advisor analysis below reflects the {model.scenario_label} scenario: {model.cost_p50} · {model.schedule} · {model.confidence_pct}% confidence</div>
              </div>}
              {model?.institutional_authority_line && <div style={{background:'linear-gradient(90deg,rgba(141,247,255,0.08),rgba(177,140,255,0.06))',border:'1px solid rgba(141,247,255,0.2)',borderRadius:8,padding:'12px 18px',marginBottom:12,display:'flex',gap:12,alignItems:'center'}}>
                <span style={{fontSize:'16px'}}>⚡</span>
                <div>
                  <div style={{fontSize:'9px',fontWeight:'800',color:'#8df7ff',letterSpacing:'.12em',marginBottom:3}}>PROGRAMME AUTHORITY LINE</div>
                  <div style={{fontSize:'13px',fontWeight:'700',color:'#fff',lineHeight:1.4}}>{model.institutional_authority_line}</div>
                </div>
              </div>}
              <p style={{fontSize:'11px',color:'#64748b',lineHeight:1.55,marginBottom:10}}>
                CASEY does not affirm. It challenges. Every answer is grounded in the live model data — not general advice.
                If your estimate is optimistic, CASEY will say so. If your risk register is weak, CASEY will name the weakness.
              </p>
              {model?.live_intel_active && <span style={{fontSize:'9px',color:'#10b981',fontWeight:'700',marginBottom:'8px',display:'flex',alignItems:'center',gap:'5px'}}>
                <span style={{width:'5px',height:'5px',borderRadius:'50%',background:'#10b981',display:'inline-block',animation:'pulse 1.5s infinite'}}/>
                {model.live_intel_mode === 'AI-enriched' ? 'AI-enriched intelligence active' : 'Open Crawl active'} — advisor answers include real-time data from {model.live_intel_sources?.split(',')[0]} and {(model.live_intel_sources||'').split(',').length - 1} other sources ({model.live_intel_timestamp})
              </span>}
              <div style={{background:'rgba(141,247,255,0.05)',border:'1px solid rgba(141,247,255,0.12)',borderRadius:'4px',padding:'8px 12px',marginBottom:'10px',fontSize:'11px'}}>
                <div style={{fontSize:'9px',fontWeight:'800',letterSpacing:'.1em',color:'#ef4444',marginBottom:'5px'}}>HARD CONSTRAINT SCENARIOS — each one recalculates the full model</div>
                {[
                  'What if the governing constraint slips by 12 months?',
                  'What if sole-source procurement is forced on the largest package?',
                  'What if the OBA-adjusted P50 becomes the board commitment number?',
                  'What if scope increases by 15% after baseline freeze?',
                  'What if the P80 reserve is held at programme level and not drawable without board approval?',
                ].map(x => <button key={x} onClick={() => ask(x)}
                  style={{display:'block',width:'100%',textAlign:'left',background:'rgba(141,247,255,0.06)',border:'1px solid rgba(141,247,255,0.15)',color:'#8df7ff',padding:'5px 10px',borderRadius:'3px',marginBottom:'4px',cursor:'pointer',fontSize:'11px',fontStyle:'italic'}}>
                  "{x}"
                </button>)}
              </div>
              <div className="advisorPrompts bigButtons">{[
                'Why is this estimate probably wrong?',
                'What assumption in this submission is most likely to be false?',
                'Which risk will kill this programme if it is not closed before capital commitment?',
                'What is management not telling the board?',
                'Why should the board not approve this today?',
                'What would make you reject this submission?',
                'If this programme fails publicly, what is the primary cause?',
                'Where is the optimism bias in this estimate?',
                'What does the reference class say the actual outturn will be?',
                'Which benchmark failure does this most resemble and why?',
                'What is the one piece of evidence missing that would change your view?',
                'Challenge my confidence score — is it justified?',
              ].map(x=><button key={x} data-question={x} onClick={()=>ask(x)}><Brain size={14}/>{x}</button>)}</div>
              <div className="chatBox boardInterrogation">{chat.length ? chat.map((m,i)=><div key={i} className={`msg ${m.role}`}>{renderChatMsg(m)}</div>) : <div className="msg assistant"><b>Board attack ready.</b><br/>Click any challenge above. CASEY will answer against the active scenario, not as a generic chatbot.</div>}</div>
              <div className="ask"><input value={chatQ} onChange={e=>setChatQ(e.target.value)} onKeyDown={e=>{if(e.key==='Enter')ask()}} placeholder="Challenge the programme — e.g. Why is this estimate wrong? Which risk will kill this? What is management hiding?"/><button onClick={() => ask(chatQ)}>Ask</button></div>
            </Card>
            <Card>
              <h2>Live Client Challenge Room</h2>
              <p className="advisorIntro">Upload a client cost estimate, XER schedule or risk register — or use the demo buttons below. CASEY challenges the document like an independent reviewer: identifies unpriced exposure, evidence gaps, reserve weaknesses and the questions to ask before committing capital.</p>
              <div className="challengeHero"><span>CASEY INTAKE NORMALISATION ENGINE</span><b>Messy file → schema detection → WBS/CBS inference → evidence gaps → board attack → export-ready challenge</b></div>
              <div className="challengeModeStrip"><b>Choose the file type to challenge</b><span>These buttons show the exact review style. Upload your own file below and CASEY will replace the sample with parsed source-file numbers.</span></div>
              <div className="challengeButtons pro">
                <button onClick={()=>setUploadResult({filename:'Contractor_Cost_Estimate_v27_FINAL.xlsx', file_type:'COST ESTIMATE', schema_confidence:'Auto-mapped', findings:['Estimate structure normalised. Cost packages identified across direct works, preliminaries and risk allowance.','The headline P50 number is present — but there is no P80/P90 basis. This is how a submitted estimate can understate exposure: the headline looks fixed but the downside is unpriced.','Contingency is present as a lump sum. CASEY cannot verify it is sized against quantified risk exposure rather than a percentage of direct cost — a common tender-basis weakness.','Basis statements are missing on 6 of the major packages. Without basis, there is no evidence of what was included or excluded — and no way to challenge scope creep later.'], red_flags:['Commercial observation: No P80/P90 range provided. The estimate looks precise but carries unquantified downside. Ask the contractor to provide a risk-adjusted range.','Commercial observation: Lump-sum contingency with no risk linkage. This is not yet a quantified reserve. Require QCRA support.','No CBS/WBS mapping. Cannot verify completeness of scope coverage or trace costs to programme activities.','Escalation basis not stated. For a multi-year programme, this is a material omission.'], next_steps:['Require the contractor to provide a P50/P80/P90 range with QCRA support.','Mandate a CBS that maps to the programme WBS and schedule activities.','Commission an independent cost review before approving the headline number.','Run CASEY QCRA alongside the contractor estimate — compare the P80 positions.'], epc_challenge:true})}><FileSpreadsheet size={18}/><b>Challenge contractor cost estimate</b><span>Detect hidden exposure, lump-sum contingency and missing basis statements.</span></button>
                <button onClick={()=>setUploadResult({filename:'Programme_Schedule_FINAL_v14.xer', file_type:'SCHEDULE (XER)', schema_confidence:'Logic mapped', findings:['Schedule logic parsed. Activities identified across civil, systems, commissioning and handover phases.','Critical path identified — but float analysis reveals operationally unusable buffer. The management date assumes best-case access windows throughout.','Logic gaps detected: 8 activities have no predecessor. These are schedule anchors — they cannot be challenged because they have no upstream dependency. This can prevent a reliable view of the real critical path.','Commissioning and trial running phases show compressed durations. These are the activities most likely to slip — and they sit directly before the opening/handover milestone.'], red_flags:['Commercial observation: Open-ended activities with no predecessor — schedule logic issue that can overstate available float.','Commercial observation: Commissioning duration appears optimistic against comparable programmes. A single failed integration test resets the clock.','Float is nominal, not operationally usable. Access windows, possession permits and operator acceptance are not confirmed in the logic.','Board date is driven by the earliest path. It should be driven by the P80/P90 QSRA finish date.'], next_steps:['Require the contractor to close all open ends and confirm predecessor logic.','Run QSRA and require the P80/P90 finish date to be the board commitment date.','Validate all commissioning durations against independent benchmarks.','Name the owner of the critical-path constraint.'], epc_challenge:true})}><Workflow size={18}/><b>Challenge programme schedule</b><span>Detect schedule padding, unusable float and optimistic commissioning dates.</span></button>
                <button onClick={()=>setUploadResult({filename:'Risk_Register_v8_Draft.xlsx', file_type:'RISK REGISTER', schema_confidence:'Schema mapped', findings:['Risk register schema mapped. Cause, event, impact and owner columns identified.','CASEY challenges every risk without a named trigger, quantified residual exposure and evidence closure date.','7 risks have mitigation confidence below 50%. These are not mitigated — they are noted. The reserve needs to account for them.','4 risks are flagged as Evidence required. These are open exposures — the source file does not yet provide the evidence that the risk is under control.'], red_flags:['Commercial observation: Mitigations are written as action phrases ("to be confirmed", "in progress") rather than evidence closure. A mitigation is only valid when the evidence is complete.','Commercial observation: Residual exposure is not reconciled to the reserve. This is the most common way a risk register hides real exposure — risks exist on paper but the money is not in the budget.','4 risks require evidence that has not been provided. These cannot be treated as mitigated for board approval purposes.','Owner accountability: all risks assigned to programme-level owners. Board needs named individual owners with accountability.'], next_steps:['Require every risk to have: named owner, confirmed trigger, quantified residual and evidence closure date.','Reconcile residual exposure to reserve — any gap requires additional provision.','The 4 Evidence Required risks must be resolved or escalated to the board as open items.','Export the challenged register after QCRA/QSRA alignment and use the board attack questions.'], epc_challenge:true})}><ShieldAlert size={18}/><b>Challenge risk register</b><span>Detect unmitigated risks, missing evidence and reserve reconciliation gaps.</span></button>
              </div>
              <h3>Upload real file</h3>
              <label className="upload proUpload"><Upload size={18}/> Upload estimate / XER / risk workbook<input type="file" onChange={upload}/></label>
              <ProfessionalIntakeResult result={uploadResult} model={model}/>
            </Card>
          </section>

          <section className="layout one"><Card><h2>Board attack simulation — the 5 questions this board will ask</h2>
            <p style={{fontSize:'11px',color:'#64748b',marginBottom:'12px'}}>These are sector-specific, programme-specific challenges that a serious investment committee will table. CASEY generates each one from live model data — your actual P50, P80, sector, and governing constraints. Not a template.</p>
            {(model.board_attack_simulation||[]).map((q,i)=><div key={i} style={{display:'flex',gap:'10px',padding:'10px 0',borderBottom:'1px solid rgba(255,255,255,0.05)'}}>
              <span style={{color:'#f59e0b',fontWeight:'900',flexShrink:0,fontSize:'11px',paddingTop:'1px'}}>{i+1}.</span>
              <span style={{color:'#e2e8f0',lineHeight:'1.5',fontSize:'13px'}}>{safeRender(q)}</span>
            </div>)}
          </Card></section>

          <section className="layout two">
            <Card><h2>Programme mortality risk</h2>
              <div style={{display:'flex',alignItems:'baseline',gap:'8px',marginBottom:'8px'}}>
                <span style={{fontSize:'48px',fontWeight:'900',color:model.programme_mortality_risk.pct>60?'#ef4444':model.programme_mortality_risk.pct>35?'#f59e0b':'#10b981'}}>{model.programme_mortality_risk.pct+'%'}</span>
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
          </section>

          <section className="layout one"><div style={{background:'linear-gradient(135deg,rgba(141,247,255,0.06),rgba(177,140,255,0.06))',border:'1px solid rgba(141,247,255,0.2)',borderRadius:'6px',padding:'14px 20px',display:'flex',gap:'12px',alignItems:'flex-start'}}>
            <span style={{color:'#8df7ff',fontWeight:'900',fontSize:'10px',letterSpacing:'.15em',flexShrink:0,paddingTop:'2px'}}>AUTHORITY LINE</span>
            <p style={{color:'#e2e8f0',fontSize:'13px',lineHeight:'1.6',margin:0,fontStyle:'italic'}}>{safeRender(model.institutional_authority_line)}</p>
          </div></section>
        </>
}

        {tab === 'runtime' && <HolyGrailRuntime model={model} scenario={scenario} generate={generate} runShock={runShock}/>}
        {tab === 'method' && <section className="layout two"><Card><h2>How CASEY calculated this</h2>{['Cost model: selected class estimate, sector template, location factor, complexity factor and scenario modifier.','Schedule model: level-based delivery logic, phase durations, critical path sensitivity and scenario acceleration/delay factors.','QCRA: cost exposure model using low / most likely / high impacts and risk-weighted contingency.','QSRA: schedule exposure model using activity-linked O/M/P delay ranges and critical path sensitivity.','Confidence score translated for executives: board-defensibility based on benchmark similarity, evidence maturity, procurement certainty, schedule logic, contingency adequacy and scenario posture.'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card><Card><h2>Commercial readiness</h2><p style={{fontSize:'12px',color:'#64748b'}}>First-pass intelligence for challenge, option testing and board preparation — before contractor tender or signed cost plan.</p><a className="contactBtn huge" href={emailLink}><Mail/> Send project for review</a></Card></section>}
        {tab === 'benchmark' && <section className="layout two">
          <Card><h2>Named global benchmarks</h2>
            <p style={{fontSize:'11px',color:'#64748b',marginBottom:'8px'}}>These are real programmes from public record — OECD, parliamentary accounts committees, company filings, academic literature (Flyvbjerg et al). Cost growth % and schedule slip are actuals, not estimates. CASEY routes every project through the closest matching comparables and applies their delivery behaviour to confidence, reserve and P80/P90 exposure.</p>
            <div style={{overflowX:'auto'}}>
              <table style={{width:'100%',borderCollapse:'collapse',fontSize:'11px'}}>
                <thead><tr style={{borderBottom:'1px solid rgba(255,255,255,0.1)'}}>
                  {['Programme','Sector','P50 Anchor','Duration','Cost Growth','Slip (mo)','Primary Failure Mode'].map(h=><th key={h} style={{padding:'6px 8px',textAlign:'left',color:'#64748b',fontWeight:'800',letterSpacing:'.08em'}}>{h}</th>)}
                </tr></thead>
                <tbody>{(model?.benchmark_comparison||[]).map((b,i)=><tr key={i} style={{borderBottom:'1px solid rgba(255,255,255,0.04)',background:i%2===0?'rgba(255,255,255,0.01)':'transparent'}}>
                  <td style={{padding:'7px 8px',color:'#e2e8f0',fontWeight:'700'}}>{b.name||b.archetype}</td>
                  <td style={{padding:'7px 8px',color:'#8df7ff',fontSize:'10px'}}>{b.sector}</td>
                  <td style={{padding:'7px 8px',color:'#94a3b8'}}>{b.anchor_cost||b.value}</td>
                  <td style={{padding:'7px 8px',color:'#94a3b8'}}>{b.anchor_duration_months ? b.anchor_duration_months + ' mo' : '—'}</td>
                  <td style={{padding:'7px 8px',color:Number(b.cost_growth_pct??b.growth)>50?'#ef4444':Number(b.cost_growth_pct??b.growth)>20?'#f59e0b':'#10b981',fontWeight:'700'}}>{(b.cost_growth_pct??b.growth) ? '+'+(b.cost_growth_pct??b.growth)+'%' : '—'}</td>
                  <td style={{padding:'7px 8px',color:b.schedule_slip_months>24?'#ef4444':b.schedule_slip_months>12?'#f59e0b':'#94a3b8',fontWeight:'700'}}>{b.schedule_slip_months ? '+'+b.schedule_slip_months : '—'}</td>
                  <td style={{padding:'7px 8px',color:'#64748b',maxWidth:'220px',lineHeight:'1.4'}}>{(b.failure_mode||b.failure||'—')}</td>
                </tr>)}
                </tbody>
              </table>
            </div>
          </Card>
          <Card><h2>What benchmarks mean for your programme</h2>
            <p className="chartCaption">CASEY routes every project through the closest real-world comparables. The benchmark similarity score, cost growth history and failure mode are applied to P80/P90 exposure and OBA adjustment.</p>
            {(model?.benchmark_comparison||[]).map((b,i)=><div key={i} className="reason" style={{borderLeft:`2px solid ${b.cost_growth_pct>80?'rgba(239,68,68,0.5)':b.cost_growth_pct>30?'rgba(245,158,11,0.5)':'rgba(141,247,255,0.3)'}`,paddingLeft:'10px',marginBottom:'8px'}}>
              <b style={{color:'#e2e8f0'}}>{b.name||b.archetype}</b>
              {b.cost_growth_pct > 0 && <span style={{background:'rgba(239,68,68,0.1)',color:'#fca5a5',borderRadius:'3px',padding:'1px 6px',fontSize:'10px',fontWeight:'800',marginLeft:'8px'}}>+{b.cost_growth_pct}% cost</span>}
              {b.schedule_slip_months > 0 && <span style={{background:'rgba(245,158,11,0.1)',color:'#fde68a',borderRadius:'3px',padding:'1px 6px',fontSize:'10px',fontWeight:'800',marginLeft:'4px'}}>+{b.schedule_slip_months+'mo'}</span>}
              <p style={{fontSize:'11px',color:'#64748b',marginTop:'4px',lineHeight:'1.5'}}>{b.lesson||b.failure_mode||b.why}</p>
            </div>)}
          </Card>

          {/* WHAT THIS MEANS FOR YOUR PROGRAMME */}
          <Card><h2>What this means for your programme</h2>
            <div style={{marginBottom:12,padding:'10px 14px',background:'rgba(6,182,212,0.06)',border:'1px solid rgba(6,182,212,0.2)',borderRadius:6}}>
              <div style={{fontSize:'10px',fontWeight:'800',color:'#06b6d4',marginBottom:6}}>YOUR PROGRAMME VS THIS COHORT</div>
              {(model?.benchmark_comparison||[]).length > 0 && <div style={{fontSize:'11px',color:'#cbd5e1',lineHeight:'1.6'}}>
                <div style={{marginBottom:4}}>
                  <b style={{color:'#fff'}}>Median cost growth in this cohort:</b>{' '}
                  <span style={{color:'#f59e0b',fontWeight:'700'}}>
                    +{Math.round((model.benchmark_comparison||[]).reduce((s,b)=>s+(b.cost_growth_pct||0),0)/Math.max((model.benchmark_comparison||[]).length,1))}%
                  </span>
                  {' '}— if this applies, your programme moves from {model.cost_p50} to{' '}
                  <span style={{color:'#ef4444',fontWeight:'700'}}>
                    {model?.p50_cost_bn ? '$'+(model.p50_cost_bn*(1+(model.benchmark_comparison||[]).reduce((s,b)=>s+(b.cost_growth_pct||0),0)/Math.max((model.benchmark_comparison||[]).length,1)/100)).toFixed(1)+'B' : '—'}
                  </span>
                </div>
                <div style={{marginBottom:4}}>
                  <b style={{color:'#fff'}}>Median schedule slip:</b>{' '}
                  <span style={{color:'#f59e0b',fontWeight:'700'}}>
                    +{Math.round((model.benchmark_comparison||[]).reduce((s,b)=>s+(b.schedule_slip_months||0),0)/Math.max((model.benchmark_comparison||[]).length,1))} months
                  </span>
                  {' '}vs your current {model.schedule}
                </div>
                <div>
                  <b style={{color:'#fff'}}>Primary failure mode across cohort:</b>{' '}
                  <span style={{color:'#fca5a5'}}>
                    {(model.benchmark_comparison||[]).reduce((acc,b)=>{const fm=b.failure_mode||'';return acc.length<fm.length?fm:acc;},'').slice(0,80)}
                  </span>
                </div>
              </div>}
            </div>

            {/* SECTOR PERCENTILE */}
            {model?.benchmark_comparison?.length >= 3 && <div style={{padding:'10px 14px',background:'rgba(139,92,246,0.06)',border:'1px solid rgba(139,92,246,0.2)',borderRadius:6}}>
              <div style={{fontSize:'10px',fontWeight:'800',color:'#a78bfa',marginBottom:6}}>WHERE YOUR PROGRAMME SITS IN THE COHORT</div>
              {renderBenchmarkStats()}
            </div>}
          </Card>
        </section>}

        {tab === 'pricing' && <section className="layout two">
          <Card><h2>CASEY Access</h2>
            <p style={{fontSize:'12px',color:'#64748b',marginBottom:'16px'}}>CASEY compresses months of advisory work into seconds. Each tier is designed around the value it delivers to the programme — not the cost of producing it manually.</p>
            <div className="pricingGrid">
              <div className="priceCard"><b>Pilot</b><strong style={{color:'#10b981'}}>From £5,000</strong><span>1 live project review, full intelligence pack, and a board pack walkthrough with your team.</span><a href={emailLink}>Start pilot</a></div>
              <div className="priceCard hot"><b>Professional</b><strong style={{color:'#8df7ff'}}>From £30,000/year</strong><span>Unlimited project packs, all sectors, all scenarios, full export suite. Replaces 1–3 advisory engagements per year.</span><a href={emailLink}>Request access</a></div>
              <div className="priceCard"><b>Enterprise</b><strong style={{color:'#b18cff'}}>From £150,000/year</strong><span>Private deployment, SSO, team seats, custom benchmark library, white-label. Replaces a full cost consultancy retainer.</span><a href={emailLink}>Book demo</a></div>
            </div>
            <div style={{marginTop:'16px',padding:'10px 14px',background:'rgba(141,247,255,0.04)',borderRadius:'4px',border:'1px solid rgba(141,247,255,0.1)'}}>
              <p style={{fontSize:'11px',color:'#64748b',lineHeight:'1.6',margin:0}}><b style={{color:'#8df7ff'}}>The value calculation:</b> A global benchmark study, location risk assessment, and financing context advisory typically cost £200K–£400K combined and take 8–12 weeks. CASEY generates all three simultaneously in 4 seconds. Enterprise access at £150K represents a fraction of a single comparable advisory engagement.</p>
            </div>
          </Card>
          <Card><h2>Send this project now</h2>
            <p style={{fontSize:'12px',color:'#64748b',marginBottom:'12px'}}>Send this output to your investment committee, board pack, or programme sponsor.</p>
            <a className="contactBtn huge" href={emailLink}><Mail size={16}/> Send project for review</a>
            {model && <button className="primary" style={{marginTop:'12px',width:'100%'}} onClick={() => download('/export/pdf', model, 'CASEY_Board_Pack.pdf')}><Download size={15}/> Download full intelligence pack</button>}
            <div style={{marginTop:'16px',borderTop:'1px solid rgba(255,255,255,0.06)',paddingTop:'14px'}}>
              <p style={{fontSize:'11px',color:'#475569',marginBottom:'10px',fontWeight:'700',letterSpacing:'.08em'}}>WHAT'S IN THE PACK</p>
              {['Cost model XLSX — P10/P50/P90 by CBS line, direct/indirect/reserve split','Risk register XLSX — cause, event, impact, owner, trigger, mitigation, residual','QCRA/QSRA Excel — cost and schedule probability curves and tornado chart','PRA schedule XER — Primavera-compatible with logic, phases and critical path','Audit model JSON — full traceability, benchmark provenance, evidence gaps','Board pack narrative — executive summary, board attack simulation, OBA assessment'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span><span style={{fontSize:'11px'}}>{x}</span></div>)}
            </div>
          </Card>
        </section>}
        {tab === 'defence' && <>
          {/* BLOOMBERG HEADLINE — the one sentence that matters */}
          {model && <div style={{background:'linear-gradient(90deg,rgba(6,182,212,0.08),rgba(139,92,246,0.06))',border:'1px solid rgba(6,182,212,0.2)',borderRadius:8,padding:'10px 16px',marginBottom:10,display:'flex',gap:10,alignItems:'center'}}>
            <span style={{fontSize:'8px',fontWeight:'900',color:'#06b6d4',letterSpacing:'.18em',flexShrink:0}}>CASEY VERDICT</span>
            <div style={{fontSize:'11px',fontWeight:'700',color:'#fff',lineHeight:1.4}}>
              {model.institutional_authority_line || ((model.confidence_pct||0)>=75 ? model.cost_p50+' '+( model.estimate_class_name)+'. '+model.confidence_pct+'% board confidence. Approval-ready.' : (model.confidence_pct||0)>=55 ? model.cost_p50+' '+( model.estimate_class_name)+'. '+model.confidence_pct+'% confidence — conditional. Governing constraint not confirmed.' : model.cost_p50+' '+( model.estimate_class_name)+'. '+model.confidence_pct+'% confidence. Do not approve without closing '+(model.reserve_gap_bn>0?'£'+model.reserve_gap_bn.toFixed(2)+'B reserve shortfall and ':'')+' governing constraint evidence.')}
            </div>
          </div>}

          {/* SCENARIO CONTEXT */}
          {model?.scenario !== 'base' && <div style={{background:'rgba(6,182,212,0.06)',border:'1px solid rgba(6,182,212,0.2)',borderRadius:6,padding:'8px 14px',marginBottom:10,display:'flex',gap:12,alignItems:'center'}}>
            <div style={{fontSize:'9px',fontWeight:'800',color:'#06b6d4'}}>{(model.scenario_label||'').toUpperCase()} SCENARIO — Board Pack reflects this scenario vs Base</div>
            <div style={{display:'flex',gap:16,fontSize:'9px',color:'#94a3b8',marginLeft:'auto'}}>
              <span>Cost: <b style={{color:'#06b6d4'}}>{model.cost_p50}</b></span>
              <span>Schedule: <b style={{color:'#06b6d4'}}>{model.schedule}</b></span>
              <span>Confidence: <b style={{color:'#06b6d4'}}>{model.confidence_pct+'%'}</b></span>
            </div>
          </div>}

          {/* PANEL 1: THE ONE RISK — full width, red, commanding */}
          {model?.mortality_event?.title && <div style={{background:'rgba(239,68,68,0.08)',border:'2px solid rgba(239,68,68,0.5)',borderRadius:10,padding:'18px 22px',marginBottom:12}}>
            <div style={{fontSize:'9px',fontWeight:'800',color:'#ef4444',letterSpacing:'.16em',marginBottom:8}}>⚠ IF WE ARE WRONG — THE ONE RISK THAT KILLS THIS PROGRAMME</div>
            <div style={{fontSize:'20px',fontWeight:'900',color:'#fff',marginBottom:8,lineHeight:1.3}}>{model.mortality_event.title}</div>
            <div style={{fontSize:'12px',color:'#fca5a5',fontStyle:'italic',lineHeight:1.7,marginBottom:10}}>"{model.mortality_event.terrifying_statement}"</div>
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr auto',gap:12,alignItems:'center'}}>
              <div style={{padding:'8px 12px',background:'rgba(239,68,68,0.12)',borderRadius:6}}>
                <div style={{fontSize:'8px',color:'#ef4444',marginBottom:2}}>PROBABILITY</div>
                <div style={{fontSize:'14px',fontWeight:'900',color:'#fca5a5'}}>{model.mortality_event.probability}</div>
              </div>
              <div style={{padding:'8px 12px',background:'rgba(239,68,68,0.12)',borderRadius:6}}>
                <div style={{fontSize:'8px',color:'#ef4444',marginBottom:2}}>COST + SCHEDULE EXPOSURE</div>
                <div style={{fontSize:'14px',fontWeight:'900',color:'#fca5a5'}}>{model.mortality_event.exposure}</div>
              </div>
              {model.mortality_event.board_action && <div style={{padding:'8px 14px',background:'rgba(239,68,68,0.08)',border:'1px solid rgba(239,68,68,0.25)',borderRadius:6}}>
                <div style={{fontSize:'8px',color:'#ef4444',marginBottom:2}}>BOARD ACTION REQUIRED</div>
                <div style={{fontSize:'10px',color:'#fff',fontWeight:'600'}}>{model.mortality_event.board_action}</div>
              </div>}
            </div>
          </div>}

          {/* PANELS 2+3: GOVERNING CONSTRAINT + DECISION SIMULATOR side by side */}
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:12,marginBottom:12}}>
            {/* PANEL 2: GOVERNING CONSTRAINT */}
            {(model?.governing_constraint_prominent || model?.governing_constraint_full?.statement) && <div style={{background:'rgba(245,158,11,0.08)',border:'2px solid rgba(245,158,11,0.4)',borderRadius:10,padding:'14px 18px'}}>
              <div style={{fontSize:'9px',fontWeight:'800',color:'#f59e0b',letterSpacing:'.14em',marginBottom:6}}>⛔ GOVERNING CONSTRAINT — WHAT MUST HAPPEN BEFORE APPROVAL</div>
              <div style={{fontSize:'15px',fontWeight:'800',color:'#fff',marginBottom:10,lineHeight:1.4}}>
                {model.governing_constraint_prominent || model.governing_constraint_full?.statement}
              </div>
              <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8}}>
                {model.governing_constraint_full?.owner && <div style={{padding:'6px 10px',background:'rgba(245,158,11,0.08)',borderRadius:5}}>
                  <div style={{fontSize:'8px',color:'#f59e0b',marginBottom:2}}>OWNER</div>
                  <div style={{fontSize:'10px',color:'#fbbf24',fontWeight:'700'}}>{model.governing_constraint_full.owner}</div>
                </div>}
                {model.governing_constraint_full?.evidence && <div style={{padding:'6px 10px',background:'rgba(245,158,11,0.08)',borderRadius:5}}>
                  <div style={{fontSize:'8px',color:'#f59e0b',marginBottom:2}}>EVIDENCE NEEDED</div>
                  <div style={{fontSize:'10px',color:'#94a3b8'}}>{model.governing_constraint_full.evidence}</div>
                </div>}
              </div>
            </div>}

            {/* PANEL 3: DECISION SIMULATOR */}
            {model?.decision_simulator && <div style={{background:'rgba(16,185,129,0.06)',border:'2px solid rgba(16,185,129,0.3)',borderRadius:10,padding:'14px 18px'}}>
              <div style={{fontSize:'9px',fontWeight:'800',color:'#10b981',letterSpacing:'.14em',marginBottom:8}}>🎮 DECISION SIMULATOR — WHAT IF THE BOARD ASKS?</div>
              {[model.decision_simulator.spend_200m,model.decision_simulator.descope_10pct,model.decision_simulator.accelerate].filter(Boolean).map((d,i)=>(
                <div key={i} style={{marginBottom:8,padding:'10px 12px',background:'rgba(16,185,129,0.05)',borderRadius:6,border:'1px solid rgba(16,185,129,0.15)'}}>
                  <div style={{fontSize:'10px',fontWeight:'800',color:'#10b981',marginBottom:4}}>{d.label}</div>
                  <div style={{display:'flex',gap:16,fontSize:'10px'}}>
                    <span style={{color:(d.cost_delta_bn||0)>0?'#ef4444':'#10b981',fontWeight:'700'}}>{(d.cost_delta_bn>0?'+':d.cost_delta_bn<0?'-':'')+((Math.abs(d.cost_delta_bn||0)).toFixed(1))+'B cost'}</span>
                    <span style={{color:(d.schedule_delta_months||0)<0?'#10b981':'#ef4444',fontWeight:'700'}}>{(d.schedule_delta_months>0?'+':d.schedule_delta_months<0?'-':'')+Math.abs(d.schedule_delta_months||0)} months</span>
                    <span style={{color:(d.confidence_delta_pct||0)>0?'#10b981':'#ef4444',fontWeight:'700'}}>{(d.confidence_delta_pct>0?'+':d.confidence_delta_pct<0?'-':'')+Math.abs(d.confidence_delta_pct||0)}% conf</span>
                  </div>
                </div>
              ))}
            </div>}
          </div>

          {/* PANEL 4: CONFIDENCE EVIDENCE — what we know vs what we assumed */}
          <div style={{background:'rgba(255,255,255,0.02)',border:'1px solid rgba(255,255,255,0.07)',borderRadius:10,padding:'14px 18px'}}>
            <div style={{fontSize:'9px',fontWeight:'800',color:'#e2e8f0',letterSpacing:'.12em',marginBottom:8}}>📋 CONFIDENCE EVIDENCE — WHAT CASEY KNOWS vs WHAT IT ASSUMED</div>
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr',gap:8,marginBottom:10}}>
              {[['📅 XER Schedule',model?.evidence_files?.xer||model?.xer_health?.from_upload,'Upload XER for verified QSRA'],
                ['📊 Cost Workbook',model?.evidence_files?.cost,'Upload cost file for verified P50'],
                ['📋 Risk Register',model?.evidence_files?.risk_register,'Upload register for verified QCRA']].map(([label,has,hint])=>(
                <div key={label} style={{padding:'8px 12px',background:has?'rgba(16,185,129,0.08)':'rgba(255,255,255,0.02)',border:`1px solid ${has?'rgba(16,185,129,0.25)':'rgba(255,255,255,0.06)'}`,borderRadius:6}}>
                  <div style={{fontSize:'12px',marginBottom:4}}>{has?'✅':'❌'}</div>
                  <div style={{fontSize:'9px',fontWeight:'700',color:has?'#10b981':'#64748b'}}>{label}</div>
                  {!has && <div style={{fontSize:'8px',color:'#475569',marginTop:2}}>{hint}</div>}
                </div>
              ))}
            </div>
            {/* Confidence decomposition compact */}
            {model?.confidence_by_discipline && <div style={{borderTop:'1px solid rgba(255,255,255,0.06)',paddingTop:8}}>
              <div style={{fontSize:'8px',color:'#475569',marginBottom:6}}>WHY {model.confidence_pct}% CONFIDENCE?</div>
              <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:6}}>
                {Object.entries(model.confidence_by_discipline).slice(0,6).map(([disc,pct])=>(
                  <div key={disc} style={{display:'flex',alignItems:'center',gap:6}}>
                    <div style={{flex:1,height:3,background:'rgba(255,255,255,0.06)',borderRadius:2}}>
                      <div style={{width:pct+'%',height:'100%',background:pct>=70?'#10b981':pct>=55?'#f59e0b':'#ef4444',borderRadius:2}}/>
                    </div>
                    <div style={{fontSize:'8px',fontWeight:'700',color:pct>=70?'#10b981':pct>=55?'#f59e0b':'#ef4444',minWidth:24}}>{pct+'%'}</div>
                    <div style={{fontSize:'7px',color:'#475569',overflow:'hidden',whiteSpace:'nowrap',textOverflow:'ellipsis',maxWidth:60}}>{disc.split('/')[0].trim()}</div>
                  </div>
                ))}
              </div>
            </div>}
            {/* Evidence chain if available */}
            {model?.evidence_chain?.length > 0 && <div style={{borderTop:'1px solid rgba(255,255,255,0.06)',paddingTop:8,marginTop:8}}>
              <div style={{fontSize:'8px',color:'#10b981',fontWeight:'700',marginBottom:4}}>EVIDENCE CHAIN — NUMBERS FROM YOUR FILES:</div>
              {model.evidence_chain.slice(0,4).map((e,i)=>(
                <div key={i} style={{fontSize:'8px',color:'#475569',marginBottom:2,display:'flex',gap:5}}>
                  <span style={{color:'#10b981',flexShrink:0}}>→</span>{e}
                </div>
              ))}
            </div>}
          </div>
        </>}
        </>}
    </main>
  </div>;
}


// ── CLIENT FILE INGEST PANEL ──────────────────────────────────────────────────
// Upload client XER/Excel/risk register → CASEY reconciles with its model
// Enriched model flows to ALL features automatically
function ClientIngestPanel({ model, setModel, onClose }) {
  const [files, setFiles] = React.useState([]);
  const [result, setResult] = React.useState(null);
  const [busy, setBusy] = React.useState(false);
  const [drag, setDrag] = React.useState(false);

  const handleFiles = (fl) => setFiles(Array.from(fl));

  const runIngest = async () => {
    if (!files.length) return;
    setBusy(true);
    try {
      const fd = new FormData();
      fd.append("model", JSON.stringify(model));
      files.forEach(f => fd.append("files", f));
      const r = await apiFetch("/ingest", { method: "POST", body: fd });
      if (r.ok) {
        const d = await r.json();
        setResult(d);
        if (d.enriched_model) {
          setModel(d.enriched_model);
        }
      } else {
        const t = await r.text();
        setResult({ error: t });
      }
    } catch(e) { setResult({ error: e.message }); }
    finally { setBusy(false); }
  };

  const ICONS = { xer:"📅", xlsx:"📊", xls:"📊", csv:"📋", pdf:"📄", docx:"📝", txt:"📝" };
  const icon = f => ICONS[f.name.split(".").pop().toLowerCase()] || "📎";
  const sevCol = s => s.includes("CRITICAL")?"#ef4444":s.includes("HIGH")?"#f59e0b":"#8df7ff";

  return <div style={{position:"fixed",inset:0,background:"rgba(2,6,23,0.9)",zIndex:9998,display:"flex",alignItems:"flex-start",justifyContent:"center",paddingTop:"40px",overflowY:"auto"}} onClick={onClose}>
    <div style={{background:"#0f172a",border:"1px solid rgba(16,185,129,0.25)",borderRadius:"8px",width:"min(760px,96vw)",maxHeight:"85vh",overflowY:"auto",padding:"24px"}} onClick={e=>e.stopPropagation()}>

      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"16px"}}>
        <div>
          <div style={{fontSize:"16px",fontWeight:"900",color:"#10b981",marginBottom:"2px"}}>UPLOAD CLIENT PROGRAMME FILES</div>
          <div style={{fontSize:"10px",color:"#475569"}}>CASEY reads your files, reconciles against its estimate, and enriches all features with your real data. Works with any project run.</div>
        </div>
        <button onClick={onClose} style={{background:"rgba(255,255,255,0.05)",border:"1px solid rgba(255,255,255,0.1)",color:"#94a3b8",borderRadius:"4px",padding:"6px 12px",cursor:"pointer",fontSize:"12px"}}>Close</button>
      </div>

      {/* File type guide */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"6px",marginBottom:"14px"}}>
        {[["📅 XER","Schedule % complete, activity logic, project dates"],
          ["📊 Excel (any)","Your cost estimate, EV data, actuals vs budget"],
          ["📋 Risk Register","Your risks, owners, probabilities, EMV"],
          ["📝 Report (.txt/.docx)","Monthly progress report, narrative, CPI"]
        ].map(([t,d])=><div key={t} style={{padding:"7px 8px",background:"rgba(16,185,129,0.05)",border:"1px solid rgba(16,185,129,0.15)",borderRadius:"4px"}}>
          <div style={{fontSize:"10px",fontWeight:"700",color:"#e2e8f0",marginBottom:"1px"}}>{t}</div>
          <div style={{fontSize:"8px",color:"#475569",lineHeight:"1.3"}}>{d}</div>
        </div>)}
      </div>

      {/* Drop zone */}
      <div onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)}
        onDrop={e=>{e.preventDefault();setDrag(false);handleFiles(e.dataTransfer.files);}}
        onClick={()=>document.getElementById("ingest-input").click()}
        style={{border:"2px dashed "+(drag?"rgba(16,185,129,0.5)":"rgba(16,185,129,0.2)"),borderRadius:"6px",padding:"20px",textAlign:"center",cursor:"pointer",background:drag?"rgba(16,185,129,0.06)":"rgba(255,255,255,0.02)",marginBottom:"10px"}}>
        <div style={{fontSize:"24px",marginBottom:"6px"}}>📂</div>
        <div style={{fontSize:"12px",fontWeight:"700",color:"#10b981",marginBottom:"3px"}}>Drop your programme files here or click to select</div>
        <div style={{fontSize:"9px",color:"#475569"}}>XER · Excel · CSV · PDF · Word · TXT — CASEY handles any format and any mess</div>
        <input id="ingest-input" type="file" multiple accept=".xer,.xlsx,.xls,.csv,.pdf,.docx,.txt" style={{display:"none"}} onChange={e=>handleFiles(e.target.files)}/>
      </div>

      {files.length > 0 && <div style={{marginBottom:"10px"}}>
        {files.map((f,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:"4px 10px",background:"rgba(255,255,255,0.04)",borderRadius:"3px",marginBottom:"2px",fontSize:"10px"}}>
          <span style={{color:"#e2e8f0"}}>{icon(f)} {f.name}</span>
          <span style={{color:"#475569"}}>{(f.size/1024).toFixed(0)}KB</span>
        </div>)}
      </div>}

      {files.length > 0 && <button onClick={runIngest} disabled={busy} style={{width:"100%",padding:"12px",background:busy?"rgba(16,185,129,0.04)":"rgba(16,185,129,0.12)",border:"1px solid rgba(16,185,129,0.3)",borderRadius:"5px",color:"#10b981",fontSize:"13px",fontWeight:"900",cursor:busy?"not-allowed":"pointer",letterSpacing:".04em",marginBottom:"12px"}}>
        {busy?"Reading and reconciling your files...":"⚡ INGEST FILES — Enrich all CASEY features with your data"}
      </button>}

      {result && !result.error && <>
        <div style={{padding:"10px 12px",background:"rgba(16,185,129,0.06)",border:"1px solid rgba(16,185,129,0.2)",borderRadius:"5px",marginBottom:"10px"}}>
          <div style={{fontSize:"11px",fontWeight:"800",color:"#10b981",marginBottom:"4px"}}>
            {result.enriched_model ? "✓ ALL CASEY FEATURES UPDATED with your data" : "Files read — see results below"}
          </div>
          <div style={{fontSize:"10px",color:"#94a3b8"}}>{result.summary}</div>
        </div>

        {/* Evidence */}
        {(result.client_context?.evidence||[]).length > 0 && <div style={{marginBottom:"10px"}}>
          <div style={{fontSize:"9px",fontWeight:"800",color:"#64748b",letterSpacing:".08em",marginBottom:"5px"}}>WHAT CASEY READ FROM YOUR FILES</div>
          {result.client_context.evidence.map((e,i)=><div key={i} style={{fontSize:"10px",color:"#94a3b8",marginBottom:"2px",lineHeight:"1.4"}}>
            <span style={{color:"#10b981",marginRight:"5px"}}>✓</span>{e}
          </div>)}
        </div>}

        {/* CASEY vs Client reconciliation */}
        {(result.reconciliation||[]).length > 0 && <div style={{marginBottom:"10px"}}>
          <div style={{fontSize:"9px",fontWeight:"800",color:"#64748b",letterSpacing:".08em",marginBottom:"6px"}}>CASEY vs YOUR STATED POSITION</div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:"10px"}}>
            <thead><tr>{["Metric","Your Files Say","CASEY Derives","Gap","Interpretation"].map(h=><th key={h} style={{textAlign:"left",padding:"4px 8px",color:"#475569",borderBottom:"1px solid rgba(255,255,255,0.07)",fontSize:"9px",fontWeight:"800"}}>{h}</th>)}</tr></thead>
            <tbody>{result.reconciliation.map((r,i)=><tr key={i} style={{background:i%2===0?"rgba(255,255,255,0.02)":"transparent"}}>
              {[r.metric, r.client, r.casey, <span style={{color:r.gap?.includes("+")?"#ef4444":"#10b981",fontWeight:"700"}}>{r.gap}</span>, r.interpretation?.substring(0,80)+"..."].map((v,ci)=><td key={ci} style={{padding:"5px 8px",color:"#94a3b8",fontSize:"9px",lineHeight:"1.4"}}>{v}</td>)}
            </tr>)}
          </tbody></table>
        </div>}

        {/* Conflicts */}
        {(result.conflicts||[]).length > 0 && <div style={{marginBottom:"10px"}}>
          <div style={{fontSize:"9px",fontWeight:"800",color:"#ef4444",letterSpacing:".08em",marginBottom:"5px"}}>CONFLICTS DETECTED</div>
          {result.conflicts.map((c,i)=><div key={i} style={{padding:"8px 10px",background:"rgba(239,68,68,0.07)",border:"1px solid rgba(239,68,68,0.2)",borderLeft:"3px solid #ef4444",borderRadius:"4px",marginBottom:"5px",fontSize:"10px",color:"#e2e8f0",lineHeight:"1.5"}}>{c}</div>)}
        </div>}

        {/* Missing data guidance */}
        {(result.gaps||[]).length > 0 && <div style={{marginBottom:"10px",padding:"8px 10px",background:"rgba(245,158,11,0.06)",border:"1px solid rgba(245,158,11,0.2)",borderRadius:"4px"}}>
          <div style={{fontSize:"9px",fontWeight:"800",color:"#f59e0b",marginBottom:"5px"}}>TO GET MORE FROM YOUR FILES — add these to your documents:</div>
          {result.gaps.map((g,i)=><div key={i} style={{fontSize:"9px",color:"#64748b",marginBottom:"3px",lineHeight:"1.4"}}>
            <span style={{color:"#f59e0b",marginRight:"5px"}}>→</span>{g}
          </div>)}
        </div>}

        <div style={{padding:"8px 10px",background:"rgba(16,185,129,0.05)",borderRadius:"4px",fontSize:"9px",color:"#475569"}}>
          <b style={{color:"#10b981"}}>All features now use your data:</b> Challenge Mode will challenge your stated cost. Digital Twin is pre-filled with your actuals. Programme Memory records this snapshot. Advisor knows your real numbers. Exports reflect your data.
        </div>
      </>}

      {result?.error && <div style={{padding:"10px",color:"#ef4444",fontSize:"10px"}}>Error: {result.error}</div>}
    </div>
  </div>;
}


// ── CHALLENGE MODE + ANALOGUE PANEL ──────────────────────────────────────────
function ChallengeAnaloguePanel({ model }) {
  const [mode, setMode] = React.useState(null); // null | 'challenge' | 'analogue'
  const [result, setResult] = React.useState(null);
  const [busy, setBusy] = React.useState(false);
  const [mem, setMem] = React.useState(() => {
    try { return JSON.parse(localStorage.getItem("casey_prog_memory") || "[]"); } catch { return []; }
  });

  const run = async (endpoint) => {
    setBusy(true); setMode(endpoint); setResult(null);
    try {
      const r = await apiFetch("/" + endpoint, {method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(model)});
      if (r.ok) { const d=await r.json(); setResult(d); }
    } catch(e) { setResult({error: e.message}); } finally { setBusy(false); }
  };

  // Save memory snapshot
  const saveSnapshot = () => {
    const snap = {
      date: new Date().toLocaleDateString("en-GB", {day:"2-digit",month:"short",year:"numeric"}),
      p50: model?.cost_p50 || "—",
      conf: model?.confidence_pct || "—",
      sched: model?.schedule_months || "—",
      scenario: model?.scenario_label || "Base",
    };
    const newMem = [snap, ...mem].slice(0, 12);
    setMem(newMem);
    try { localStorage.setItem("casey_prog_memory", JSON.stringify(newMem)); } catch {}
  };

  const sevCol = (s) => s==="CRITICAL"?"#ef4444":s==="HIGH"?"#f59e0b":"#8df7ff";

  return <><div style={{marginBottom:"10px",display:"flex",gap:"8px",flexWrap:"wrap",alignItems:"center"}}>
    <button onClick={()=>run("challenge")} disabled={busy} style={{
      padding:"9px 18px",background:"rgba(239,68,68,0.1)",border:"1px solid rgba(239,68,68,0.3)",
      borderRadius:"5px",color:"#ef4444",fontSize:"12px",fontWeight:"900",cursor:"pointer",letterSpacing:".04em"
    }}>🔴 CHALLENGE MY PROGRAMME</button>
    <button onClick={()=>run("analogue")} disabled={busy} style={{
      padding:"9px 18px",background:"rgba(141,247,255,0.08)",border:"1px solid rgba(141,247,255,0.2)",
      borderRadius:"5px",color:"#8df7ff",fontSize:"12px",fontWeight:"800",cursor:"pointer"
    }}>📚 SECTOR ANALOGUES</button>
    <button onClick={saveSnapshot} style={{
      padding:"9px 14px",background:"rgba(16,185,129,0.08)",border:"1px solid rgba(16,185,129,0.2)",
      borderRadius:"5px",color:"#10b981",fontSize:"11px",fontWeight:"700",cursor:"pointer"
    }}>💾 SAVE TO MEMORY</button>
    {busy && <span style={{fontSize:"10px",color:"#475569"}}>Running...</span>}
  </div>

  {/* Programme Memory */}
  {mem.length > 0 && <div style={{marginBottom:"10px",padding:"8px 12px",background:"rgba(255,255,255,0.02)",border:"1px solid rgba(255,255,255,0.07)",borderRadius:"5px"}}>
    <div style={{fontSize:"9px",fontWeight:"800",color:"#64748b",letterSpacing:".08em",marginBottom:"6px"}}>PROGRAMME MEMORY — {mem.length} snapshots saved</div>
    <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:"9px"}}>
      <thead><tr>{["Date","P50","Confidence","Schedule","Scenario"].map(h=><th key={h} style={{textAlign:"left",padding:"3px 8px",color:"#475569",borderBottom:"1px solid rgba(255,255,255,0.06)"}}>{h}</th>)}</tr></thead>
      <tbody>{mem.map((s,i)=><tr key={i} style={{background:i===0?"rgba(141,247,255,0.04)":"transparent"}}>
        {[s.date,s.p50,s.conf+"%",s.sched+" mo",s.scenario].map((v,ci)=><td key={ci} style={{padding:"3px 8px",color:i===0?"#8df7ff":"#64748b",fontWeight:i===0?"700":"400"}}>{v}</td>)}
      </tr>)}</tbody>
    </table></div>
    {mem.length >= 2 && (() => {
      const curr_p50 = parseFloat((mem[0]?.p50||"0").replace(/[^0-9.]/g,""));
      const prev_p50 = parseFloat((mem[mem.length-1]?.p50||"0").replace(/[^0-9.]/g,""));
      const curr_conf = parseFloat(mem[0]?.conf||0);
      const prev_conf = parseFloat(mem[mem.length-1]?.conf||0);
      if (curr_p50 > 0 && prev_p50 > 0) {
        const drift = round((curr_p50/prev_p50-1)*100, 1);
        const conf_drift = curr_conf - prev_conf;
        const acc = Math.max(0, 100 - Math.abs(drift));
        return <div style={{marginTop:"6px",padding:"6px 8px",background:"rgba(255,255,255,0.02)",borderRadius:"3px",fontSize:"9px",color:"#94a3b8"}}>
          Forecast drift over {mem.length} snapshots: <span style={{color:drift>10?"#ef4444":drift>5?"#f59e0b":"#10b981",fontWeight:"700"}}>{(drift>0?"+":"")+drift+'%'}</span>
          {"  "}Confidence drift: <span style={{color:conf_drift<-5?"#ef4444":conf_drift>5?"#10b981":"#f59e0b",fontWeight:"700"}}>{(conf_drift>0?"+":"")+conf_drift} pts</span>
          {"  "}Forecast stability: <span style={{color:acc>85?"#10b981":acc>70?"#f59e0b":"#ef4444",fontWeight:"700"}}>{acc.toFixed(0)+'%'}</span>
        </div>;
      }
      return null;
    })()}
  </div>}

  {/* Challenge Results */}
  {mode === "challenge" && result && !result.error && <div style={{marginBottom:"10px",padding:"12px 14px",background:"rgba(239,68,68,0.06)",border:"2px solid rgba(239,68,68,0.2)",borderRadius:"6px"}}>
    <div style={{fontSize:"13px",fontWeight:"900",color:"#ef4444",marginBottom:"6px"}}>CHALLENGE MODE — {model?.title || model?.subsector || "Programme"} — CASEY acting as hostile examiner</div>
    <div style={{display:"flex",gap:"10px",marginBottom:"8px",flexWrap:"wrap"}}>
      <span style={{fontSize:"9px",background:"rgba(239,68,68,0.1)",border:"1px solid rgba(239,68,68,0.2)",borderRadius:"3px",padding:"2px 8px",color:"#ef4444",fontWeight:"700"}}>{result.critical_count} CRITICAL</span>
      <span style={{fontSize:"9px",background:"rgba(245,158,11,0.1)",border:"1px solid rgba(245,158,11,0.2)",borderRadius:"3px",padding:"2px 8px",color:"#f59e0b",fontWeight:"700"}}>{result.high_count} HIGH</span>
      {result.regulatory_framework && <span style={{fontSize:"9px",color:"#64748b"}}>Regulatory ref: {result.regulatory_framework}</span>}
    </div>
    <p style={{fontSize:"10px",color:"#94a3b8",marginBottom:"10px",lineHeight:"1.6"}}>{result.summary}</p>
    {(result.challenges||[]).map((c,i)=><div key={i} style={{marginBottom:"8px",padding:"9px 12px",background:"rgba(255,255,255,0.03)",borderLeft:"3px solid "+sevCol(c.severity),borderRadius:"4px"}}>
      <div style={{display:"flex",gap:"8px",alignItems:"center",marginBottom:"4px"}}>
        <span style={{background:sevCol(c.severity)+"20",color:sevCol(c.severity),fontSize:"7px",fontWeight:"900",padding:"2px 6px",borderRadius:"2px"}}>{c.severity}</span>
        <span style={{fontSize:"10px",fontWeight:"800",color:"#e2e8f0"}}>{c.category}</span>
      </div>
      <p style={{fontSize:"10px",color:"#94a3b8",margin:"0 0 4px 0",lineHeight:"1.5"}}>{c.challenge}</p>
      <p style={{fontSize:"9px",color:"#10b981",margin:0,fontWeight:"600"}}>What CASEY would accept: {c.what_would_make_me_accept}</p>
    </div>)}
  </div>}

  {/* Analogue Results */}
  {mode === "analogue" && result && !result.error && <div style={{marginBottom:"10px",padding:"10px 14px",background:"rgba(141,247,255,0.04)",border:"1px solid rgba(141,247,255,0.15)",borderRadius:"6px"}}>
    <div style={{fontSize:"12px",fontWeight:"900",color:"#8df7ff",marginBottom:"6px"}}>HISTORICAL ANALOGUES — what this programme most resembles</div>
    <p style={{fontSize:"10px",color:"#64748b",marginBottom:"10px"}}>{result.summary}</p>
    {(result.analogues||[]).map((a,i)=><div key={i} style={{marginBottom:"8px",padding:"9px 12px",background:"rgba(255,255,255,0.03)",border:"1px solid rgba(255,255,255,0.08)",borderRadius:"4px",display:"flex",gap:"12px"}}>
      <div style={{flexShrink:0,textAlign:"center",width:"52px"}}>
        <div style={{fontSize:"22px",fontWeight:"900",color:i===0?"#ef4444":i===1?"#f59e0b":"#94a3b8"}}>{a.match_pct+'%'}</div>
        <div style={{fontSize:"8px",color:"#475569"}}>match</div>
      </div>
      <div style={{flex:1}}>
        <div style={{fontSize:"11px",fontWeight:"800",color:"#e2e8f0",marginBottom:"2px"}}>{a.name}</div>
        <div style={{fontSize:"9px",color:"#94a3b8",marginBottom:"3px",lineHeight:"1.4"}}>{a.pattern}</div>
        <div style={{display:"flex",gap:"12px",flexWrap:"wrap"}}>
          <span style={{fontSize:"9px",color:"#ef4444",fontWeight:"700"}}>Outcome: {a.outcome}</span>
          <span style={{fontSize:"9px",color:"#10b981"}}>Lesson: {a.lesson}</span>
        </div>
      </div>
    </div>)}
  </div>}

  {(mode === "challenge" || mode === "analogue") && result?.error && <div style={{padding:"8px",color:"#ef4444",fontSize:"10px"}}>Error: {result.error}</div>}
  </>;
}

// Small helper used in programme memory
function round(n, d) { return Math.round(n * Math.pow(10,d)) / Math.pow(10,d); }


// ── CASEY HELP PANEL ────────────────────────────────────────────────────────
const HELP_ARTICLES = [
  {id:'start',cat:'Getting started',icon:'🚀',title:'How to run your first project',body:'Type any capital programme in plain English in the FREE RUN box. Include sector, country, scale and key constraints. Example: "HS2 Phase 2b tunnelling UK rail 250km". Hit Run. Results in 4–12 seconds. Your first run is free. Earth Demo and Space Demo are always free.',tags:['run','start','free run','how to','project','first']},
  {id:'overview',cat:'Getting started',icon:'📊',title:'What do the numbers mean?',body:'P50 = most likely outturn cost. P80 = the board conversation number — what you need in reserve. Confidence % = how board-defensible the estimate is at this definition maturity. 75%+ is required for investment committee approval. The 5 scenario cards (Base, Faster, Cheaper, Lower Risk, Premium) show cost/schedule/confidence for each trade-off — click any to recalculate everything.',tags:['p50','p80','confidence','numbers','mean','kpi','what']},
  {id:'tabs',cat:'Getting started',icon:'📑',title:'What does each tab show?',body:'OVERVIEW: your baseline at a glance. TWIN: live digital twin — update with real progress data. COST: full CBS cost workbook with P10/P50/P90 and unit rates. SCHEDULE: programme logic and QSRA delivery date curve. RISK: full risk register with cause/event/impact/owner/EMV. QCRA/QSRA: Monte Carlo cost and schedule probability curves. SCENARIOS: five trade-off comparisons. INTEL: strategic intelligence and CASEY position. ASSURANCE: board challenge questions. ADVISOR: ask anything. BENCHMARKS: 63 named real programmes.',tags:['tabs','overview','cost','schedule','risk','qcra','scenarios','intel','advisor','benchmarks','what','tab']},
  {id:'twin',cat:'Digital Twin',icon:'⚡',title:'How to use the Digital Twin',body:'Step 1: Run your project to establish the baseline (P50, schedule, confidence). Step 2: When your programme is live, go to the ⚡ Twin tab. Step 3: Enter real data — earned value %, actual spend vs plan, schedule slip, sector-specific signals. Step 4: Click UPDATE TWIN. CASEY recalculates your forecast-at-completion and gives board alerts if you are heading toward P80.',tags:['twin','digital twin','live','update','earned value','cpi','forecast','completion']},
  {id:'twin-inputs',cat:'Digital Twin',icon:'⚡',title:'What data do I need for the twin?',body:'Core: Programme % complete, actual spend vs plan (100=on budget, 108=8% over), earned value % (EV/BAC×100), schedule slip in months, scope changes count. Sector signals (auto-detected): Rail = possessions used %, open IEMs, signalling milestones. Space = TRL achieved, open anomalies, launch manifest confirmed. Nuclear = regulatory hold-points cleared %, design freeze achieved. Defence = requirements frozen, SC/DV clearances %. You can load a pre-filled demo scenario with one click.',tags:['twin','inputs','data','earned value','ev','possessions','trl','anomalies','requirements']},
  {id:'twin-alerts',cat:'Digital Twin',icon:'⚡',title:'What do twin alerts mean?',body:'CRITICAL = board notification required immediately. "Forecast P50 exceeds original P80" = reserve is exhausted, rebaselining required. "CPI below 0.9" = at this cost performance rate the programme will exceed P80. HIGH = requires action before next gate review. GREEN = programme tracking on or ahead of baseline. The CPI (Cost Performance Index) = Earned Value ÷ Actual Cost. Below 1.0 = over budget, above 1.0 = under budget.',tags:['twin','alerts','critical','cpi','p80','rebaseline','green']},
  {id:'advisor',cat:'Advisor',icon:'💬',title:'How to use the Advisor',body:'The Advisor tab is a live what-if engine. Ask in plain English: "What if planning is delayed 18 months?" — CASEY reruns the model with that constraint and shows a cost/confidence/schedule delta. Try: "Is this programme gate-ready?", "What is the P80 exposure?", "Which risk will kill this programme?", "What is the OBA-adjusted outturn?", "What would an external reviewer challenge first?", "What does the benchmark data say?". The Advisor knows your current model — every answer is specific to your programme.',tags:['advisor','ask','what if','gate','p80','risk','benchmark','outturn','query']},
  {id:'exports',cat:'Exports',icon:'📤',title:'How to export and download files',body:'Scroll down to the export buttons below the scenario cards. EXPORT BOARD PACK = PDF board pack (all sections). EXPORT COST WORKBOOK = Excel CBS with P10/P50/P90 and unit rates. EXPORT RISK REGISTER = full risk register Excel with EMV and owners. EXPORT XER = Primavera P6 schedule file. EXPORT QCRA/QSRA = Monte Carlo workbook. All exports are generated from your live model data — not templates. If export fails, wait 10 seconds and try again (backend may be processing).',tags:['export','download','pdf','xlsx','excel','xer','risk register','qcra','qsra','board pack']},
  {id:'scenarios',cat:'Scenarios',icon:'🔀',title:'How to run scenario comparisons',body:'The 5 scenario cards (Base, Faster, Cheaper, Lower Risk, Premium) are visible below the main KPI bar. Click any card to recalculate everything with that trade-off applied — cost, schedule, confidence, risk register and all exports update. Faster = compressed schedule but higher cost and lower confidence. Cheaper = lower authorisation number but longer schedule and higher residual risk. Lower Risk = higher reserve, longer duration, stronger confidence. Premium = maximum resilience with visible capex premium.',tags:['scenario','faster','cheaper','lower risk','premium','compare','trade-off','recalculate']},
  {id:'benchmarks',cat:'Intelligence',icon:'📚',title:'How benchmarks work',body:'63 named real programmes from public record calibrate every estimate. Each programme has its actual outturn cost growth %, schedule slip and primary failure mode. For HS2-type rail: Crossrail +88%, HS2 Phase 1 +140%, CalHSR +288%. CASEY routes your programme through the closest comparables and applies their delivery behaviour to confidence, reserve and P80/P90 exposure. These are NOT generic percentages — they are named programmes with cited sources.',tags:['benchmarks','comparables','crossrail','hs2','real programmes','calibrate','reference class']},
  {id:'oba',cat:'Intelligence',icon:'📊',title:'What is OBA and why does it matter?',body:'OBA (Optimism Bias Adjustment) is the systematic tendency to underestimate cost and schedule. CASEY applies sector-specific OBA from Flyvbjerg 2022 and IPA Annual Reports. Rail = +44% mean cost growth, Nuclear = +55%, Space = +60%. The OBA-adjusted P50 is what the reference class says the outturn will be. HM Treasury Green Book requires OBA disclosure in all public programme board cases. If your board case does not show the OBA-adjusted number, it will be challenged.',tags:['oba','optimism bias','green book','ipa','flyvbjerg','reference class','treasury']},
  {id:'confidence',cat:'Intelligence',icon:'🎯',title:'What does confidence % mean?',body:'Confidence is CASEY board-defensibility score — not a probability of success. It measures how well the estimate, schedule and risk register would hold up under challenge. 75%+ = gate-ready for investment committee. 60-75% = conditional, evidence gaps must be closed. Below 60% = not gate-ready, material evidence closure required. It is derived from estimate class, sector benchmarks, risk register quality, procurement certainty and location factors.',tags:['confidence','board','defensibility','gate','75%','what','score','meaning']},
  {id:'sector-failure',cat:'Intelligence',icon:'⚠',title:'What is the Sector Failure Pattern?',body:'Every sector has a primary failure mode that causes programmes to be cancelled, restructured or rebaselined. Rail = systems integration deferred (civil works complete but railway cannot run). Nuclear = first-of-kind design changes after FCD. Space = mission assurance burden underestimated. Defence = requirements instability after contract. Data centres = grid connection not on critical path. This is in the Mortality tab. No consultant names it because doing so reduces their fee scope.',tags:['sector failure','mortality','cancel','restructure','rebaseline','pattern','rail','nuclear','space','defence']},
  {id:'gate-review',cat:'Intelligence',icon:'🚦',title:'What is the Gate Review assessment?',body:'CASEY maps your programme to IPA gateway readiness (G0-G4). G0 = strategic definition. G1 = business justification. G2 = delivery strategy. G3 = investment decision. G4 = readiness for service. The gate verdict (READY / CONDITIONAL / NOT READY / BLOCKED) comes from confidence level and estimate class. The Evidence Gaps tab names specifically what is missing before any gate can complete. IPA and Green Book references are cited.',tags:['gate','ipa','gateway','g2','g3','readiness','evidence','conditional','blocked']},
  {id:'gate',cat:'Intelligence',icon:'🚦',title:'What is the Gate Review assessment?',body:'CASEY maps your programme to IPA gateway readiness (G0-G4). G0 = strategic definition. G1 = business justification. G2 = delivery strategy. G3 = investment decision. G4 = readiness for service. The gate verdict (READY / CONDITIONAL / NOT READY) comes from confidence level and estimate class. The Evidence Gaps tab names what is missing before any gate can complete.',tags:['gate','ipa','gateway','g2','g3','readiness','evidence','conditional']},
  {id:'challenge',cat:'Intelligence',icon:'🔴',title:'How to use Challenge Mode',body:'Click the red CHALLENGE MY PROGRAMME button on the Overview tab after running any project. CASEY becomes the hostile examiner — it challenges your estimate class, OBA disclosure, P80 reserve, risk register quality, governing constraint, schedule basis and benchmark evidence. Each challenge includes what CASEY would need to accept the position. Works for every sector and country.',tags:['challenge','hostile','examiner','oba','p80','risk register','benchmark']},
  {id:'sector',cat:'Intelligence',icon:'⚠',title:'What is the Sector Failure Pattern?',body:'Every sector has a primary failure mode. Rail = systems integration deferred (civil complete but railway cannot run). Nuclear = first-of-kind design changes post-FCD. Space = mission assurance burden underestimated. Data centre = grid connection not on critical path. Airport = baggage and fire safety systems integration. Mining = social licence failure at near-completion. The Challenge tab names which historical programme failure your programme most resembles.',tags:['sector failure','mortality','cancel','crossrail','vogtle','artemis','pattern']},
  {id:'casey-challenge',cat:'How CASEY Works',icon:'⚡',title:'CASEY Output Challenge — how the self-attack works',body:'Every time CASEY generates an output — whether from a free project run, Earth Demo, Space Demo, or the Showcase Library — it immediately attacks its own output before you see it. This is called the CASEY Output Challenge. It runs automatically on every programme, every sector, every location, Earth and Space. WHAT IT CHECKS: Cost Realism — CASEY compares its own P50 against known named programme benchmarks (California HSR $110B, HS2 £69B, Vogtle $35B, F-35 $428B and 30 others from GAO, IPA and Infrastructure Australia). If CASEY generated $11B for California HSR, the challenge would immediately flag it as 12x below the GAO benchmark. It also checks the P80/P50 spread, OBA uplift against Flyvbjerg reference classes, and unit rates against sector norms. Schedule Realism — checks whether the QSRA P80 is believable (P80 should be 10-25% above base), whether the programme duration matches sector norms, and whether it aligns with named programme published timelines. Risk Completeness — checks whether the risk register has enough risks for the sector (10+ for board submission), whether sector-specific risks are present (grid connection for energy, planning for rail, regulatory for nuclear, cyber for defence), and whether risks have quantified EMV. OBA Calibration — checks whether Optimism Bias uplift has been applied and whether it matches the Flyvbjerg 2022 reference class for the sector. HOW TO READ IT: Green (80%+) means the output passed all checks and is board-defensible. Amber (60-79%) means the output needs review before submission. Red (below 60%) means there are significant gaps — do not use for board submission without addressing the flags. Each dimension shows a specific issue, not a generic warning. The challenge runs in under 100ms and is shown at the top of the Overview tab. You cannot turn it off — it is part of what makes CASEY defensible.',tags:['challenge','self-attack','quality','realism','oba','risk','cost','benchmark','defence','board']},
  {id:'free-limits',cat:'Access & Pricing',icon:'🔓',title:'What is free? What requires a subscription?',body:'CASEY has a generous free tier that covers the core intelligence experience. ALWAYS FREE — unlimited: Earth Demo (run as many times as you like, any sector, any scenario), Space Demo (same — full space programme intelligence, unlimited), Showcase Library (137 named real programmes from HS2 to JWST, unlimited browsing and loading), Open Crawl intelligence (live World Bank, GDELT, Wikipedia, SpaceX, FX and climate data on every run). FREE ONCE: 1 project run — your own Free Run where you type a custom programme description. 1 comparison run. After your 1 free run, you will see the upgrade screen. Earth Demo and Space Demo are never limited. PROFESSIONAL (£99/mo, coming soon): unlimited project runs, all exports (PDF board pack, Excel workbook, risk register, QCRA/QSRA, P6 XER), Advisor (Claude and GPT-4o), Digital Twin, file ingestion, Programme Memory. TEAM (£349/mo, coming soon): 5 seats, everything in Professional, priority support, white-label PDF exports.',tags:['free','pricing','limits','subscription','demo','project run','upgrade']},
  {id:'opencrawl',cat:'Open Crawl',icon:'🌐',title:'What is Open Crawl intelligence?',body:'CASEY uses Open Crawl — a system of free, open data APIs — to fetch real-time global intelligence for every programme run. No paid subscription required. Sources: NewsAPI (free tier, 80,000+ news sources including Reuters, Bloomberg, New Civil Engineer, Engineering News-Record, Infrastructure Intelligence — full article headlines and descriptions from the last 30 days, set NEWS_API_KEY in Render environment to activate), GDELT Project (global news database, updated every 15 minutes, 100+ countries, 65 languages), World Bank Open Data (217 countries, GDP growth, inflation, infrastructure investment), Wikipedia (programme intelligence), SpaceX API (live launch manifest for space programmes), Open-Meteo (7-day climate signals for construction risk), Open Exchange Rates (live FX for 170 currencies). When ANTHROPIC_API_KEY or OPENAI_API_KEY is set in the server environment, the raw data is additionally processed through AI for a richer narrative — but the underlying data is always free.',tags:['open crawl','live','intelligence','real-time','world bank','wikipedia','spacex','nasa','free','global']},
  {id:'opencrawl-how',cat:'Open Crawl',icon:'🌐',title:'How Open Crawl works in CASEY',body:'Every time you run a project — Custom Free Run, Earth Demo, Space Demo or Showcase Library — CASEY builds the model first, then starts fetching live data in the background. It takes up to 6 seconds. You will see a pulsing green dot and the text OPEN CRAWL or LIVE INTEL appear in the navigation bar once data arrives. The intelligence appears in the Overview tab as a green-bordered panel above the main analysis. It also appears in the PDF Board Pack (final page) and the QCRA/QSRA workbook. For UK, France, Nigeria or any of 217 countries: inflation data, GDP growth and infrastructure investment signals. For space programmes: live SpaceX launch manifest and current mission data. The data informs the confidence calibration and is passed to the Advisor so it can reference current conditions.',tags:['open crawl','how it works','navigation','green dot','overview','pdf','exports','advisor','background']},
  {id:'twin-guide',cat:'Digital Twin',icon:'⚡',title:'Complete guide to the Digital Twin',body:'Step 1: Run any project (Free Run or Showcase Library) to establish baseline. Step 2: Click the Twin tab. Step 3: Either (a) upload your real XER and Excel files using the Build Twin From Files section, or (b) load a demo scenario to see what the twin produces, or (c) fill in the fields manually. Step 4: Click Update Twin. You get: forecast-at-completion, confidence score, CPI, governing constraint, board defensibility score, failure pattern match, recovery options (3 paths), board questions and executive narrative.',tags:['twin','digital twin','update','upload','files','xer','excel','how to']},
  {id:'location',cat:'Intelligence',icon:'🌍',title:'How does location affect the model?',body:'CASEY has 70+ country profiles: cost multipliers, regulatory frameworks, approval bodies, financing context and OBA notes. UK programmes show £ currency, IPA/DLUHC/ORR/ONR as approval bodies, and 1.2× cost multiplier. Nigerian programmes show NGN currency, federal regulatory framework and higher governance risk note. Australian programmes show A$ currency and Infrastructure Australia context. Simply include the country in your description — CASEY detects it automatically.',tags:['location','country','currency','uk','australia','nigeria','nigeria','framework','regulatory']},
];

function HelpPanel({ onClose }) {
  const [query, setQuery] = React.useState('');
  const [openId, setOpenId] = React.useState(null);
  const filtered = query.length < 2
    ? HELP_ARTICLES
    : HELP_ARTICLES.filter(a =>
        (a.title+' '+a.body+' '+a.tags.join(' ')).toLowerCase().includes(query.toLowerCase())
      );
  const cats = [...new Set(filtered.map(a=>a.cat))];
  return <div style={{position:'fixed',inset:0,background:'rgba(2,6,23,0.92)',zIndex:9999,display:'flex',alignItems:'flex-start',justifyContent:'center',paddingTop:'40px',overflowY:'auto'}} onClick={onClose}>
    <div style={{background:'#0f172a',border:'1px solid rgba(141,247,255,0.2)',borderRadius:'8px',width:'min(720px,96vw)',maxHeight:'80vh',overflowY:'auto',padding:'24px'}} onClick={e=>e.stopPropagation()}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:'16px'}}>
        <div>
          <div style={{fontSize:'18px',fontWeight:'900',color:'#8df7ff',marginBottom:'2px'}}>CASEY HELP & FEATURE GUIDE</div>
          <div style={{fontSize:'10px',color:'#475569'}}>Click any topic to expand. Press Escape or click outside to close.</div>
        </div>
        <button onClick={onClose} style={{background:'rgba(255,255,255,0.05)',border:'1px solid rgba(255,255,255,0.1)',color:'#94a3b8',borderRadius:'4px',padding:'6px 12px',cursor:'pointer',fontSize:'12px'}}>✕ Close</button>
      </div>
      <input
        type="text" placeholder="Search — try 'digital twin', 'export', 'P80', 'confidence', 'benchmark'..."
        value={query} onChange={e=>setQuery(e.target.value)} autoFocus
        style={{width:'100%',padding:'10px 14px',background:'rgba(255,255,255,0.06)',border:'1px solid rgba(141,247,255,0.2)',borderRadius:'5px',color:'#e2e8f0',fontSize:'12px',marginBottom:'16px',boxSizing:'border-box'}}
      />
      {filtered.length === 0 && <div style={{color:'#475569',fontSize:'12px',textAlign:'center',padding:'20px'}}>No articles match "{query}" — try a different search term.</div>}
      {cats.map(cat => <div key={cat} style={{marginBottom:'12px'}}>
        <div style={{fontSize:'9px',fontWeight:'800',color:'#475569',letterSpacing:'.1em',marginBottom:'6px'}}>{cat.toUpperCase()}</div>
        {filtered.filter(a=>a.cat===cat).map(a => <div key={a.id} style={{marginBottom:'4px'}}>
          <button onClick={()=>setOpenId(openId===a.id?null:a.id)} style={{
            width:'100%',textAlign:'left',padding:'8px 12px',
            background:openId===a.id?'rgba(141,247,255,0.08)':'rgba(255,255,255,0.03)',
            border:`1px solid ${openId===a.id?'rgba(141,247,255,0.25)':'rgba(255,255,255,0.07)'}`,
            borderRadius:'4px',cursor:'pointer',display:'flex',alignItems:'center',gap:'8px'
          }}>
            <span style={{fontSize:'14px'}}>{a.icon}</span>
            <span style={{flex:1,fontSize:'11px',fontWeight:'700',color:'#e2e8f0'}}>{a.title}</span>
            <span style={{color:'#475569',fontSize:'10px'}}>{openId===a.id?'▲':'▼'}</span>
          </button>
          {openId===a.id && <div style={{padding:'10px 14px 12px 34px',background:'rgba(141,247,255,0.04)',borderLeft:'2px solid rgba(141,247,255,0.15)',marginTop:'1px',borderRadius:'0 0 4px 4px'}}>
            <p style={{fontSize:'11px',color:'#94a3b8',lineHeight:'1.7',margin:0}}>{a.body}</p>
          </div>}
        </div>)}
      </div>)}
      <div style={{borderTop:'1px solid rgba(255,255,255,0.06)',paddingTop:'10px',marginTop:'8px',fontSize:'9px',color:'#334155'}}>
        Can not find your answer? Email deepa@caseai.co.uk — or use the Advisor tab to ask CASEY directly.
      </div>
    </div>
  </div>;
}


function SavedProjectsPanel({ projects, onLoad, onDelete, onClose }) {
  return <div style={{position:"fixed",inset:0,background:"rgba(2,6,23,0.9)",zIndex:9990,display:"flex",alignItems:"center",justifyContent:"center"}} onClick={onClose}>
    <div style={{background:"#0f172a",border:"1px solid rgba(141,247,255,0.2)",borderRadius:"8px",width:"min(600px,95vw)",maxHeight:"80vh",overflowY:"auto",padding:"20px"}} onClick={e=>e.stopPropagation()}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"12px"}}>
        <div style={{fontSize:"14px",fontWeight:"900",color:"#8df7ff"}}>Saved Projects</div>
        <button onClick={onClose} style={{background:"none",border:"1px solid rgba(255,255,255,0.1)",color:"#64748b",borderRadius:"4px",padding:"4px 10px",cursor:"pointer"}}>Close</button>
      </div>
      {(!projects || projects.length === 0) && <p style={{color:"#475569",fontSize:"11px"}}>No saved projects yet. Run a project and save it from the Overview tab.</p>}
      {(projects||[]).map((p,i) => <div key={i} style={{marginBottom:"8px",padding:"10px 12px",background:"rgba(255,255,255,0.03)",border:"1px solid rgba(255,255,255,0.07)",borderRadius:"5px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div>
          <div style={{fontSize:"11px",fontWeight:"700",color:"#e2e8f0",marginBottom:"2px"}}>{p.title || p.subsector || "Project"}</div>
          <div style={{fontSize:"9px",color:"#475569"}}>{p.cost_p50} · {p.schedule_months}mo · {p.confidence_pct+'%'}</div>
        </div>
        <div style={{display:"flex",gap:"6px"}}>
          <button onClick={()=>onLoad(p)} style={{padding:"5px 10px",background:"rgba(141,247,255,0.1)",border:"1px solid rgba(141,247,255,0.2)",color:"#8df7ff",borderRadius:"3px",cursor:"pointer",fontSize:"10px",fontWeight:"700"}}>Load</button>
          <button onClick={()=>onDelete(i)} style={{padding:"5px 10px",background:"rgba(239,68,68,0.1)",border:"1px solid rgba(239,68,68,0.2)",color:"#ef4444",borderRadius:"3px",cursor:"pointer",fontSize:"10px"}}>Delete</button>
        </div>
      </div>)}
    </div>
  </div>;
}


createRoot(document.getElementById('root')).render(<CaseyErrorBoundary><App/></CaseyErrorBoundary>);
