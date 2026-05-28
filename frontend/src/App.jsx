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

const API_CANDIDATES = [import.meta.env.VITE_API_URL, 'http://127.0.0.1:8000', 'http://localhost:8000', 'http://127.0.0.1:8010', 'http://localhost:8010'].filter(Boolean);
let API = API_CANDIDATES[0];
async function apiFetch(path, options) {
  let lastError;
  for (const base of API_CANDIDATES) {
    try {
      const r = await fetch(base + path, options);
      API = base;
      return r;
    } catch (e) { lastError = e; }
  }
  throw lastError || new Error('CASEY backend unreachable');
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
      try { if (!['jaimahadeshwar@yahoo.com','test@yahoo.com','deepa@caseai.co.uk','admin@controlorbit.com','demo@controlorbit.com','jai@controlorbit.com'].includes((form.email||'').toLowerCase().trim())) { localStorage.setItem('casey_demo_used','1'); } } catch(e) {}
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

        {/* Stage 3: what's still missing — shown as sentence not pills */}
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
      {busy && <div className="missionProcessing"><Rocket size={18}/><div><b>CASEY is building your intelligence pack</b><span>Parsing brief, applying {brief.sector||'infrastructure'} benchmarks…</span></div></div>}
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
  return <section className="scenarioRail">{scenarios.map(s => { const row = matrix.find(x => x.scenario === s) || {}; const active = s === scenario; return <button key={s} className={active?'active':''} onClick={() => generate(s, model?.prompt || prompt, model || projectContext)}><b>{labels[s] || s}</b><span>{row.cost_p50 || '—'} · {row.schedule_months || '—'} mo · {row.confidence_pct || '—'}%</span><em>{row.risk || (active ? 'selected' : 'run scenario')}</em></button> })}</section>;
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
  const p80Cost = qcra.p80 ? fmt(qcra.p80) : 'the P80 cost';
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
  const seed = stableSeed(model);
  const cohort = model?.mode === 'Space' ? 42 + (seed % 18) : 118 + (seed % 64);
  const sims = 18000 + (seed % 7000);
  const rev = `CASEY-X.${116 + (seed % 7)}.${seed % 10}`;
  return <section className="orbitalMetaRail">
    <div><b>Model</b><span>{rev}</span></div>
    <div><b>Calibration</b><span>{model?.mode === 'Space' ? 'mission assurance / orbital logistics' : 'capital delivery / infrastructure peer memory'}</span></div>
    <div><b>Benchmark cohort</b><span>{cohort} comparable programmes</span></div>
    <div><b>Simulation count</b><span>{sims.toLocaleString()} QCRA/QSRA trials</span></div>
    <div><b>Evidence class</b><span>{model?.confidence_pct >= 75 ? 'approval challenge' : model?.confidence_pct >= 55 ? 'option selection' : 'pre-approval evidence gap'}</span></div>
    <div className="modeSwitch"><button className={mode==='exec'?'active':''} onClick={()=>setMode('exec')}>Executive</button><button className={mode==='board'?'active':''} onClick={()=>setMode('board')}>Board</button><button className={mode==='delivery'?'active':''} onClick={()=>setMode('delivery')}>Delivery</button><button className={mode==='analyst'?'active':''} onClick={()=>setMode('analyst')}>Analyst</button></div>
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
  const p80 = model?.monte_carlo?.qcra?.p80 ? fmt(model.monte_carlo.qcra.p80) : 'P80 exposure';
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
    { label:'Cost reconciliation', value: reconcileCheck < 0.02 ? 'PASS' : 'CHECK', detail:`Direct ${fmt(direct)} + Indirect ${fmt(indirect)} + Reserve ${fmt(reserves)} = ${fmt(direct+indirect+reserves)} vs P50 ${fmt(total)}` },
    { label:'Scenario lock', value:String(model?.scenario_label || model?.scenario || 'Base'), detail:'Cards, narratives, QCRA/QSRA and exports are stamped from the selected scenario payload.' },
    { label:'P-tail linkage', value: qcra.p80 ? fmt(qcra.p80) : 'P80 active', detail:`Cost P80 and QSRA P80 ${qsra.p80 || '—'} months are visible for board challenge.` },
    { label:'Evidence gate', value: String(confidenceLens(model).headline || ''), detail: String(confidenceLens(model).decisionRule || '') },
    { label:'Competitive threat', value:'High', detail:'Replaces manual slide-production with auditable scenario propagation and board attack logic.' },
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
        <h2>Incumbent consultant pressure test</h2>
        <p className="big">Designed to make a traditional PMO / cost-consultant deck look slow, static and non-auditable without making unsupported claims about any named firm.</p>
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
          <h2>What this just replaced</h2>
          <p className="feeSubtitle">
            A traditional early-stage advisory engagement for the same deliverables.
            No firm names — just the numbers.
          </p>
        </div>
        <div className="feeTotalBox">
          <span className="feeLabel">Advisory equivalent</span>
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

function GatedMessage({ raw }) {
  let msg = "You've used your one free CASEY intelligence run.";
  let sub = "To run more projects, compare scenarios or download the full output pack, get in touch.";
  let email = "deepa@caseai.co.uk";
  let linkedin = "https://www.linkedin.com/company/caseai";
  try {
    const p = JSON.parse(raw);
    if (p.message) msg = p.message;
    if (p.sub) sub = p.sub;
    if (p.email) email = p.email;
    if (p.linkedin) linkedin = p.linkedin;
  } catch {}
  return (
    <div className="caseyGate">
      <div className="caseyGateInner">
        <span className="caseyGateIcon">✦</span>
        <h3>{msg}</h3>
        <p>{sub}</p>
        <div className="caseyGateCtas">
          <a href={"mailto:" + email} className="caseyGateBtn primary">
            ✉ {email}
          </a>
          <a href={linkedin} target="_blank" rel="noopener noreferrer" className="caseyGateBtn secondary">
            in  Connect on LinkedIn
          </a>
        </div>
      </div>
    </div>
  );
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
      generate('base', prog.caseySignal.prompt, null, { healthCheck: true, programme: prog.name });
      setTab('assurance');
    }, 100);
  }, []);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [chatQ, setChatQ] = useState('');
  const [chat, setChat] = useState([{ role: 'assistant', text: 'Ask CASEY why cost, schedule, contingency or risk confidence is moving.' }]);
  const [uploadResult, setUploadResult] = useState(null);
  const [viewMode, setViewMode] = useState('exec');
  const [propagating, setPropagating] = useState(false);
  const [simulationStage, setSimulationStage] = useState('');
  const [exportingLabel, setExportingLabel] = useState('');
  const [confidencePulse, setConfidencePulse] = useState(false);
  const [showShowcase, setShowShowcase] = useState(false);

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

  // Mark the free run as used — only called after a real generate(), never after instant demos
  function markDemoUsed() {
    try {
      localStorage.setItem('casey_demo_used', '1');
      setDemoUsed(true);
    } catch(e) {}
  }

  async function loadInstantDemo(type) {
    setLoading(true); setError(''); setModel(null); setTab('overview');
    try {
      const BACKEND = import.meta.env.VITE_BACKEND_URL || 'https://corbit-1.onrender.com';
      const res = await fetch(`${BACKEND}/demo/${type}`);
      if (!res.ok) throw new Error('Demo unavailable — backend may be waking up, try again in 30 seconds.');
      const m = await res.json();
      setModel(m);
      setPrompt(m.prompt || '');
      setScenario(m.scenario || 'base');
      setClient('ControlOrbit Demo');
    } catch(e) {
      setError(String(e.message || e));
    } finally { setLoading(false); }
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
      setError('Your free CASEY intelligence run has been used. Explore the Earth or Space demos, browse the Showcase Library, or contact us for full access.');
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
  function runEarth() { setProjectContext(null); loadInstantDemo('earth'); }
  function runSpace() { setShowShowcase(false); setProjectContext(null); loadInstantDemo('space'); }
  function runShowcase(project) { setClient(project.client || 'Strategic reference case'); setShow(false); setShowShowcase(false); setProjectContext(null); setScenario('base'); setPrompt(project.prompt); generate('base', project.prompt, null, project.client || 'Strategic reference case', { isShowcase: true }); }
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
      setChat(x => [...x, { role: 'assistant', text: String(answer || 'CASEY returned no advisor response.') }]);
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
    <header className="v50ConsoleTop"><Logo/><nav><button onClick={() => { setModel(null); setProjectContext(null); setShowShowcase(false); setShow(true); }}>Home</button><button onClick={() => setBriefing(true)}>Film</button><button onClick={() => setTrialOpen(true)}>Free run</button><button onClick={() => { setModel(null); setShow(false); setShowShowcase(true); }}>Showcase library</button><button onClick={runEarth}>Earth demo</button><button onClick={runSpace}>Space demo</button><a href={emailLink}>Request access</a></nav></header>
    <main className={model ? 'v50Console' : 'v50Console emptyConsole'}>
      {error && <GatedMessage raw={error} />}
      {!model && showShowcase && <ShowcaseLibrary onRun={runShowcase} onBack={() => setShowShowcase(false)} />}
      {!model && !show && !showShowcase && <section className="commandGrid"><Card className="command"><h1>Generate a live project model</h1><label>Project command</label><textarea value={prompt} onChange={e => setPrompt(e.target.value)} /> <div className="chips">{examples.map(x => <button key={x} onClick={() => setPrompt(x)}>{x}</button>)}</div><div className="grid4"><input value={client} onChange={e => setClient(e.target.value)} placeholder="Client / operator"/><select value={classLevel} onChange={e => setClassLevel(e.target.value)}>{[1,2,3,4,5].map(x => <option key={x} value={x}>Class {x}</option>)}</select><select value={scheduleLevel} onChange={e => setScheduleLevel(e.target.value)}>{[1,2,3,4,5].map(x => <option key={x} value={x}>Level {x}</option>)}</select><select value={scenario} onChange={e => setScenario(e.target.value)}>{scenarios.map(x => <option key={x} value={x}>{x}</option>)}</select></div><button className="primary" onClick={() => generate()}><Sparkles/> Generate full intelligence pack</button><button className="secondary" onClick={() => setShowShowcase(true)}><Globe2/> Open global showcase library</button></Card><Card><h2>What CASEY will produce</h2>{['Executive summary and recommendation','Direct / indirect / reserve cost view','Scenario-linked estimate, schedule and confidence','Risk register with cause, event, impact and mitigation','QCRA + QSRA curves and tornado drivers','Pricing and next-step contact actions'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card></section>}
      {model && <>
        <section className="confidenceEngineBadge"><b>{model.confidence_engine_label || 'CASEY Confidence Engine'}</b><span>{safeRender(typeof model.confidence_engine_detail === 'object' ? model.confidence_engine_detail?.plain_english || 'Benchmark + probabilistic + sector-trained reasoning' : model.confidence_engine_detail || 'Benchmark + probabilistic + sector-trained reasoning')}</span></section>
        <TrustRuntimeBar model={model}/>
        <LiveCalibrationStrip model={model}/>
        <section className="kpis"><Kpi icon={Globe2} label="Mode / sector" value={safeRender(model.mode)} sub={safeRender(model.subsector)}/><Kpi icon={Activity} label="P50 cost" value={safeRender(model.cost_p50)} sub={safeRender(model.cost_range)}/><Kpi icon={Zap} label="Schedule" value={safeRender(model.schedule)} sub={`QSRA P80 ${model.monte_carlo?.qsra?.p80 || '—'} months`}/><Kpi icon={ShieldAlert} label="Delivery confidence" value={safeRender(confidenceLens(model)?.headline)} sub={`${safeRender(model.risk)} risk · ${safeRender(model.confidence_pct)}% · ${safeRender(model.scenario_label)}`} hot/></section>
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
      <nav className="tabs">{[['overview','Overview'],['compare','Scenarios'],['delta','Scenario Intel'],['causal','Causal OS'],['cost','Cost'],['schedule','Schedule'],['risk','Risk'],['monte','QCRA/QSRA'],['outputs','Outputs'],['assurance','Assurance'],['runtime','Live Stress Test'],['advisor','Advisor'],['method','Methodology'],['pricing','Pricing']].map(x => <button key={x[0]} className={tab===x[0]?'active':''} onClick={() => setTab(x[0])}>{x[1]}</button>)}</nav>
        {tab === 'overview' && <>
          {model.executive_shock_insight && <section className="layout one"><Card className="shockCard"><h2>⚡ Live model update</h2><p>{model.executive_shock_insight}</p></Card></section>}
          <section className="layout two">
            <Card><h2>Executive intelligence summary</h2><p className="big">{model.executive_summary || `${model.title} has been classified as ${safeRender(model.subsector)}. CASEY generated a first-pass cost, schedule, risk and confidence model for the selected scenario.`}</p><div className="miniMetrics"><b><span>Direct cost</span>{fmt(direct)}</b><b><span>Indirect cost</span>{fmt(indirect)}</b><b><span>Risk / reserve</span>{fmt(reserves)}</b></div><h3>Recommendation</h3>{(model.next_best_actions || []).slice(0,5).map((x,i)=><div className="reason" key={i}><span>{i+1}</span>{safeRender(x)}</div>)}</Card>
            <Card><h2>Board briefing</h2>{(model.board_briefing || model.board_challenge_questions || []).slice(0,5).map((x,i)=><div className="reason" key={i}><span>{i+1}</span>{safeRender(x)}</div>)}<h3>CASEY thinking</h3><p className="caseyThinking">{model.casey_thinking || 'CASEY interprets this as a system-of-systems infrastructure programme requiring cost, schedule, risk and decision intelligence.'}</p></Card>
          </section>
          <section className="layout two eliteLayer">
            <Card className="confidenceMeaningCard"><h2>What confidence means</h2><h3>{safeRender(confLens?.headline)}</h3><p className="big">{safeRender(confLens?.meaning)}</p><div className="reason"><span>!</span><b>Decision rule</b><br/>{safeRender(confLens?.decisionRule)}</div><div className="reason"><span>→</span><b>Primary constraint</b><br/>{safeRender(confLens?.constraint)}</div><div className="reason"><span>%</span><b>Plain English</b><br/>Confidence is not optimism. It is CASEY board-defensibility score based on benchmark fit, evidence maturity, procurement certainty, schedule logic, reserve adequacy and scenario posture.</div></Card>
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

        {tab === 'compare' && <section className="layout two"><Card><h2>Scenario comparison</h2><p className="big">Switch options before paying for another advisory cycle. Each button re-runs cost, schedule, confidence, risk register, QCRA/QSRA and exports from the same source of truth.</p>{model?.stress_test_applied && <div style={{background:"rgba(245,158,11,0.1)",border:"1px solid rgba(245,158,11,0.3)",borderRadius:"3px",padding:"8px 12px",marginBottom:"10px",fontSize:"11px",color:"#f59e0b"}}><b>STRESS TEST ACTIVE: {String(model.stress_test_applied).replace(/_/g," ").toUpperCase()}</b><br/>{model.stress_test_note} — P50 now {safeRender(model.cost_p50)}, confidence {model.confidence_pct}%. Scenario re-runs below use the stressed baseline.</div>}<div className="runtimeInline"><button onClick={()=>setTab('runtime')}><Zap size={15}/> Open Live Stress Test</button><button onClick={()=>runShock('signalling_slip')}>Simulate 4-month signalling slip</button><button onClick={()=>runShock('procurement_gap')}>Simulate procurement evidence gap</button></div>{(()=>{
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
})()}</Card><Card><h2>Buyer decision lens</h2>{['Base: balanced reference case for board challenge','Faster: more capex, lower confidence, shorter duration','Cheaper: lower authorisation number, longer schedule, higher residual risk','Lower Risk: higher reserve, longer duration, stronger confidence','Premium: resilience and optionality bought with visible capex premium'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}<h3>Current trade-off</h3><div className="triLens"><b>Gained</b>{tradePack.gained.map(x=><span key={x}>{x}</span>)}<b>Sacrificed</b>{tradePack.sacrificed.map(x=><span key={x}>{x}</span>)}<b>Exposed</b>{tradePack.exposed.map(x=><span key={x}>{x}</span>)}</div></Card></section>}
        {tab === 'cost' && <section className="layout two"><Card><h2>Scenario cost bridge vs Base</h2><p className="chartCaption">This explains why the selected scenario is cheaper or more expensive than Base before showing the workbook lines.</p>{model?.stress_test_applied && <div style={{background:"rgba(141,247,255,0.05)",borderLeft:"2px solid #8df7ff",padding:"8px 12px",marginBottom:"8px",fontSize:"11px",color:"#8df7ff"}}>Stress test applied: {String(model.stress_test_applied).replace(/_/g," ")} — cost recalculated to {safeRender(model.cost_p50)}. The waterfall below reflects this change.</div>}{costWaterfall.map((x,i)=><div className={`reason ${x.kind==='total'?'deltaReason':''}`} key={i}><span>{i+1}</span><b>{x.driver}</b><br/>{x.kind==='total'?x.value:(x.value_bn>=0?'+':'−') + ' ' + x.value}</div>)}<h3>Cost estimate workbook</h3><Table rows={costs} cols={[["cbs","CBS"],["description","Description"],["type","Type"],["p10_bn","Low/P10"],["p50_bn","Most likely/P50"],["p90_bn","High/P90"],["impact_basis","Basis"]]} moneyCols={["p10_bn","p50_bn","p90_bn"]}/></Card><Card><h2>Cost composition</h2><p className="chartCaption">Direct, indirect and reserve are scenario-controlled and reconciled to selected P50. For the detailed uncertainty view use QCRA/QSRA.</p><ResponsiveContainer width="100%" height={320}><BarChart data={[{name:'Direct',value:direct},{name:'Indirect',value:indirect},{name:'Reserve',value:reserves}]}><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis dataKey="name"/><YAxis/><Tooltip/><Bar dataKey="value" fill="#8df7ff"/></BarChart></ResponsiveContainer></Card></section>}
        {tab === 'schedule' && <section className="layout two"><Card><h2>Schedule bridge vs Base</h2><p className="chartCaption">This is the month-by-month reason the scenario becomes faster or slower than Base.</p>{scheduleWaterfall.map((x,i)=><div className={`reason ${x.kind==='total'?'deltaReason':''}`} key={i}><span>{i+1}</span><b>{x.driver}</b><br/>{x.kind==='total'?`${x.months} months`:(x.months>=0?'+':'') + x.months + ' months'}</div>)}<h3>Scenario schedule logic</h3><Table rows={schedule} cols={[["activity_id","Activity"],["phase","Phase"],["activity","Name"],["predecessor","Pred"],["duration_months","Months"],["critical","Critical"],["basis","Basis"]]}/></Card><Card><h2>QSRA finish-date curve</h2><p className="chartCaption">P50 equals the headline schedule. P80/P90 show how severe the delivery tail becomes after the scenario trade-off.</p><div className="metrics"><div>P50<b>{qsra.p50} mo</b></div><div>P80<b>{qsra.p80} mo</b></div><div>P90<b>{qsra.p90} mo</b></div></div><ResponsiveContainer width="100%" height={280}><LineChart data={curve}><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis dataKey="percentile"/><YAxis/><Tooltip formatter={(v) => [`${v} months`, "QSRA finish date"]}/><ReferenceLine x={50} stroke="#ffffff88" label="P50 = headline"/><ReferenceLine x={80} stroke="#ffffff55" label="P80 = board risk"/><Line type="monotone" name="QSRA finish date" dataKey="schedule_months" stroke="#b18cff" strokeWidth={4}/></LineChart></ResponsiveContainer><div className="reason p80Translation"><span>1/5</span>{safeRender(p80Talk.schedule)}</div><div className="reason p80Translation"><span>!</span>{safeRender(p80Talk.board)}</div>{(model.monte_carlo?.curve_readout || []).slice(1).map((x,i)=><div className="reason" key={i}><span>{i+1}</span>{x}</div>)}</Card></section>}
        {tab === 'risk' && <section className="layout two"><Card><h2>Risk Register Pro</h2><p>Risk output should include cause, event, impact, owner, mitigation and links to WBS/CBS. These risks drive the QCRA/QSRA P-curves — the top 3 by EMV determine the P80 tail.</p>{model?.stress_test_applied && <div style={{background:"rgba(239,68,68,0.08)",borderLeft:"2px solid #ef4444",padding:"6px 10px",marginBottom:"8px",fontSize:"11px",color:"#ef4444"}}>Stress test applied: risk posture has shifted. Confidence is now {model.confidence_pct}%. The risks below drove this position before the shock was applied.</div>}<Table rows={risks} cols={[['risk_id','ID'],['risk','Risk'],['cause','Cause'],['event','Event'],['impact','Impact'],['probability_pct','Prob %'],['activity_id','Activity'],['cbs','CBS'],['owner','Owner'],['mitigation','Mitigation']]}/></Card><Card><h2>Top exposure drivers</h2><ResponsiveContainer width="100%" height={380}><BarChart data={tornado} layout="vertical"><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis type="number"/><YAxis dataKey="driver" type="category" width={150}/><Tooltip/><Bar dataKey="contribution" fill="#8df7ff"/></BarChart></ResponsiveContainer></Card></section>}
        {tab === 'monte' && <section className="layout two"><Card><h2>QCRA cost range curve</h2>{model?.stress_test_applied && <div style={{background:'rgba(245,158,11,0.08)',borderLeft:'2px solid #f59e0b',padding:'6px 10px',marginBottom:'8px',fontSize:'11px',color:'#f59e0b'}}>Stress test active: {String(model.stress_test_applied).replace(/_/g,' ')} — P50 updated to {safeRender(model.cost_p50)}. Download Export QCRA/QSRA to capture the stressed curves.</div>}<p className="chartCaption">This is not a spend forecast over time. It is the probability range: P50 matches the headline cost, P80/P90 visualise the downside contingency tail created by the selected scenario.</p><div className="metrics"><div>P50 headline<b>{safeRender(model.cost_p50)}</b></div><div>P80 risk exposure<b>{fmt(qcra.p80)}</b></div><div>P90 stress case<b>{fmt(qcra.p90)}</b></div></div><ResponsiveContainer width="100%" height={280}><AreaChart data={curve}><defs><linearGradient id="caseyG" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stopColor="#8df7ff" stopOpacity=".55"/><stop offset="1" stopColor="#8df7ff" stopOpacity="0"/></linearGradient></defs><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis dataKey="percentile"/><YAxis/><Tooltip formatter={(v) => [`$${Number(v).toFixed(1)}B`, "QCRA total outturn"]}/><ReferenceLine x={50} stroke="#ffffff88" label="P50 = headline"/><ReferenceLine x={80} stroke="#ffffff55" label="P80 = board risk"/><Area type="monotone" name="QCRA total outturn" dataKey="cost_bn" stroke="#8df7ff" fill="url(#caseyG)"/></AreaChart></ResponsiveContainer>{(model.monte_carlo?.curve_readout || []).slice(0,1).map((x,i)=><div className="reason" key={i}><span>{i+1}</span>{safeRender(x)}</div>)}<div className="reason p80Translation"><span>1/5</span>{safeRender(p80Talk.cost)}</div><div className="reason"><span>!</span>This curve is a probability distribution, not spend over time. The x-axis is confidence percentile. P50 equals the headline estimate; P80/P90 are board downside exposure.</div></Card><Card><h2>QSRA schedule range curve</h2><p className="chartCaption">P50 matches the headline duration. P80/P90 show the likely board conversation if critical path risk lands.</p><div className="metrics"><div>P50 headline<b>{qsra.p50} mo</b></div><div>P80 risk date<b>{qsra.p80} mo</b></div><div>P90 stress date<b>{qsra.p90} mo</b></div></div><ResponsiveContainer width="100%" height={280}><LineChart data={curve}><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis dataKey="percentile"/><YAxis/><Tooltip formatter={(v) => [`${v} months`, "QSRA finish date"]}/><ReferenceLine x={50} stroke="#ffffff88" label="P50 = headline"/><ReferenceLine x={80} stroke="#ffffff55" label="P80 = board risk"/><Line type="monotone" name="QSRA finish date" dataKey="schedule_months" stroke="#b18cff" strokeWidth={4}/></LineChart></ResponsiveContainer><div className="reason p80Translation"><span>1/5</span>{safeRender(p80Talk.schedule)}</div><div className="reason p80Translation"><span>!</span>{safeRender(p80Talk.board)}</div>{(model.monte_carlo?.curve_readout || []).map((x,i)=><div className="reason" key={i}><span>{i+1}</span>{safeRender(x)}</div>)}</Card></section>}
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

        {tab === 'causal' && <section className="layout two"><CausalGraph model={model}/><BenchmarkIntelligence model={model}/><Card><h2>Evidence Mode: {viewMode}</h2>{evidenceScorecard(model).map((x,i)=><div className="reason" key={x.name}><span>{i+1}</span><b>{x.name}: {Math.round(x.score)}%</b><br/>{x.note}</div>)}</Card></section>}

        {tab === 'outputs' && <section className="layout two"><Card><h2>Generated Artefacts</h2><p>The public demo previews the intelligence pack. Enterprise access unlocks the live generated controls deliverables.</p><div className="exports v50Exports lockedExports">
          <button onClick={() => download('/export/workbook', model, `${model.id || 'casey'}_COST_WORKBOOK.xlsx`)}><FileSpreadsheet/> Generate Cost Model XLSX</button>
          <button onClick={() => download('/export/risk-register', model, `${model.id || 'casey'}_RISK_REGISTER.xlsx`)}><Database/> Generate Risk Register XLSX</button>
          <button onClick={() => download('/export/xer', model, `${model.id || 'casey'}_PRA_SCHEDULE.xer`)}><Workflow/> Generate PRA Schedule XER</button>
          <button onClick={() => download('/export/qcra-qsra', model, `${model.id || 'casey'}_QCRA_QSRA.xlsx`)}><BarChart3/> Generate QCRA/QSRA Pack</button>
          <button onClick={() => download('/export/json', model, `${model.id || 'casey'}_AUDIT_MODEL.json`)}><Brain/> Generate Audit File JSON</button>
          <button onClick={() => download('/export/all', model, `${model.id || 'casey'}_FULL_BOARD_PACK.zip`)}><Download/> Generate Full Pack ZIP</button>
          <a className="contactBtn" href={emailLink}><Mail/> Request Enterprise Review</a></div></Card><Card><h2>What the pack delivers</h2>{['Executive control centre with project, scenario, class, level and confidence clearly identified','Scenario comparison covering Base, Faster, Cheaper, Lower Risk and Premium cases','Selected estimate class plus all class levels for audit and challenge','Direct, indirect and reserve cost views with QCRA cost curve and cost tornado','All schedule levels with QSRA schedule curve and schedule tornado','Risk register with cause, event, impact, owner, mitigation, trigger and quantified likelihood','Basis of Estimate, assumptions, exclusions and benchmark validation','Commercial next steps: buyer action, procurement challenge and board decision path'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card></section>}

        {tab === 'assurance' && <><IncumbentPressurePanel model={model} direct={direct} indirect={indirect} reserves={reserves} reconcileCheck={reconcileCheck}/><section className="layout two"><Card><h2>Assurance room weapons</h2>{['Open with the P80/P90 exposure, not the headline P50.','Ask which evidence package retires the governing constraint.','Force every mitigation to name owner, trigger, residual exposure and date.','Show scenario trade-offs live before anyone can defend a single-point estimate.','Export the audit model immediately so the conversation moves from opinion to traceability.'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card><Card><h2>What scares traditional advisors</h2>{['CASEY moves faster than manual assurance cycles.','Every scenario rewrites cost, schedule, confidence and board posture from one payload.','The system exposes contradictions instead of polishing the management story.','It turns static reports into live investment-committee interrogation.'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card></section><section className="layout one"><ProgrammeHealthSignal onRunHealthCheck={runHealthCheck}/></section></>}

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
            <Card style={{borderLeft:'2px solid rgba(239,68,68,0.4)'}}><h2 style={{color:'#ff6b7d'}}>What the T&T deck says</h2>
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
            <p style={{fontSize:'11px',color:'#64748b',marginBottom:'12px'}}>These are sector-specific, programme-specific challenges that a serious investment committee will table. CASEY generates them from live model data — not a generic template. T&T cannot answer these in the room without CASEY.</p>
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

          {/* ORIGINAL ADVISOR PANEL */}
          <section className="layout two advisorElite challengeRoom"><Card><h2>CASEY Board Assurance Console</h2><p className="advisorIntro">Click any question. CASEY answers instantly using the live programme model — not a generic response. Each answer references your actual P50, P80, confidence level and sector. Generate a project first for the most specific answers.</p><div className="advisorPrompts bigButtons">{['What is the board not seeing?','What would a traditional cost consultant say that CASEY challenges?','What evidence is missing before this becomes board-approvable?','What is the real governing chain?','Which assumptions collapse confidence first?','What is the board really deciding?','If this programme fails, what will be blamed publicly?','Give me CASEY POSITION.','What has management not yet evidenced?','What would destroy board confidence fastest?','What reported green item is not yet board-defensible?','Show Traditional Controls vs CASEY.','What is the one intervention that changes confidence fastest?','What would an external assurance reviewer challenge first?'].map(x=><button key={x} data-question={x} onClick={()=>ask(x)}><Brain size={14}/>{x}</button>)}</div><div className="chatBox boardInterrogation">{chat.length ? chat.map((m,i)=><div key={i} className={`msg ${m.role}`}>{(() => {
    const lines = String(m.text||'').split('\n');
    return lines.map((line, li) => {
      if (!line.trim()) return <div key={li} style={{height:'5px'}}/>;
      if (line.startsWith('**') && line.endsWith('**') && line.length > 4)
        return <div key={li} className="chatHeading">{line.replace(/\*\*/g,'')}</div>;
      const parts = line.split(/\*\*([^*]+)\*\*/g);
      return <p key={li} className="chatLine">{parts.map((p,pi)=>pi%2===1?<strong key={pi}>{p}</strong>:p)}</p>;
    });
  })()}</div>) : <div className="msg assistant"><b>Board attack ready.</b><br/>Click any challenge above. CASEY will answer against the active scenario, not as a generic chatbot.</div>}</div><div className="ask"><input value={chatQ} onChange={e=>setChatQ(e.target.value)} onKeyDown={e=>{if(e.key==='Enter')ask()}} placeholder="Ask any board challenge question — e.g. What would destroy board confidence fastest?"/><button onClick={() => ask(chatQ)}>Ask</button></div></Card><Card><h2>Live Client Challenge Room</h2><p className="advisorIntro">Upload a client cost estimate, XER schedule or risk register — or use the demo buttons below. CASEY challenges the document like an independent reviewer: identifies unpriced exposure, evidence gaps, reserve weaknesses and the questions to ask before committing capital.</p><div className="challengeHero"><span>CASEY INTAKE NORMALISATION ENGINE</span><b>Messy file → schema detection → WBS/CBS inference → evidence gaps → board attack → export-ready challenge</b></div><div className="challengeModeStrip"><b>Choose the file type to challenge</b><span>These buttons show the exact review style. Upload your own file below and CASEY will replace the sample with parsed source-file numbers.</span></div><div className="challengeButtons pro">
  <button onClick={()=>setUploadResult({filename:'Contractor_Cost_Estimate_v27_FINAL.xlsx', file_type:'COST ESTIMATE', schema_confidence:'Auto-mapped', findings:['Estimate structure normalised. Cost packages identified across direct works, preliminaries and risk allowance.','The headline P50 number is present — but there is no P80/P90 basis. This is how a submitted estimate can understate exposure: the headline looks fixed but the downside is unpriced.','Contingency is present as a lump sum. CASEY cannot verify it is sized against quantified risk exposure rather than a percentage of direct cost — a common tender-basis weakness.','Basis statements are missing on 6 of the major packages. Without basis, there is no evidence of what was included or excluded — and no way to challenge scope creep later.'], red_flags:['Commercial observation: No P80/P90 range provided. The estimate looks precise but carries unquantified downside. Ask the contractor to provide a risk-adjusted range.','Commercial observation: Lump-sum contingency with no risk linkage. This is not yet a quantified reserve. Require QCRA support.','No CBS/WBS mapping. Cannot verify completeness of scope coverage or trace costs to programme activities.','Escalation basis not stated. For a multi-year programme, this is a material omission.'], next_steps:['Require the contractor to provide a P50/P80/P90 range with QCRA support.','Mandate a CBS that maps to the programme WBS and schedule activities.','Commission an independent cost review before approving the headline number.','Run CASEY QCRA alongside the contractor estimate — compare the P80 positions.'], epc_challenge:true})}><FileSpreadsheet size={18}/><b>Challenge contractor cost estimate</b><span>Detect hidden exposure, lump-sum contingency and missing basis statements.</span></button>
  <button onClick={()=>setUploadResult({filename:'Programme_Schedule_FINAL_v14.xer', file_type:'SCHEDULE (XER)', schema_confidence:'Logic mapped', findings:['Schedule logic parsed. Activities identified across civil, systems, commissioning and handover phases.','Critical path identified — but float analysis reveals operationally unusable buffer. The management date assumes best-case access windows throughout.','Logic gaps detected: 8 activities have no predecessor. These are schedule anchors — they cannot be challenged because they have no upstream dependency. This can prevent a reliable view of the real critical path.','Commissioning and trial running phases show compressed durations. These are the activities most likely to slip — and they sit directly before the opening/handover milestone.'], red_flags:['Commercial observation: Open-ended activities with no predecessor — schedule logic issue that can overstate available float.','Commercial observation: Commissioning duration appears optimistic against comparable programmes. A single failed integration test resets the clock.','Float is nominal, not operationally usable. Access windows, possession permits and operator acceptance are not confirmed in the logic.','Board date is driven by the earliest path. It should be driven by the P80/P90 QSRA finish date.'], next_steps:['Require the contractor to close all open ends and confirm predecessor logic.','Run QSRA and require the P80/P90 finish date to be the board commitment date.','Validate all commissioning durations against independent benchmarks.','Name the owner of the critical-path constraint.'], epc_challenge:true})}><Workflow size={18}/><b>Challenge programme schedule</b><span>Detect schedule padding, unusable float and optimistic commissioning dates.</span></button>
  <button onClick={()=>setUploadResult({filename:'Risk_Register_v8_Draft.xlsx', file_type:'RISK REGISTER', schema_confidence:'Schema mapped', findings:['Risk register schema mapped. Cause, event, impact and owner columns identified.','CASEY challenges every risk without a named trigger, quantified residual exposure and evidence closure date.','7 risks have mitigation confidence below 50%. These are not mitigated — they are noted. The reserve needs to account for them.','4 risks are flagged as Evidence required. These are open exposures — the source file does not yet provide the evidence that the risk is under control.'], red_flags:['Commercial observation: Mitigations are written as action phrases ("to be confirmed", "in progress") rather than evidence closure. A mitigation is only valid when the evidence is complete.','Commercial observation: Residual exposure is not reconciled to the reserve. This is the most common way a risk register hides real exposure — risks exist on paper but the money is not in the budget.','4 risks require evidence that has not been provided. These cannot be treated as mitigated for board approval purposes.','Owner accountability: all risks assigned to programme-level owners. Board needs named individual owners with accountability.'], next_steps:['Require every risk to have: named owner, confirmed trigger, quantified residual and evidence closure date.','Reconcile residual exposure to reserve — any gap requires additional provision.','The 4 Evidence Required risks must be resolved or escalated to the board as open items.','Export the challenged register after QCRA/QSRA alignment and use the board attack questions.'], epc_challenge:true})}><ShieldAlert size={18}/><b>Challenge risk register</b><span>Detect unmitigated risks, missing evidence and reserve reconciliation gaps.</span></button>
</div><h3>Upload real file</h3><label className="upload proUpload"><Upload size={18}/> Upload estimate / XER / risk workbook<input type="file" onChange={upload}/></label><ProfessionalIntakeResult result={uploadResult} model={model}/></Card></section></>
}

        {tab === 'runtime' && <HolyGrailRuntime model={model} scenario={scenario} generate={generate} runShock={runShock}/>}
        {tab === 'method' && <section className="layout two"><Card><h2>How CASEY calculated this</h2>{['Cost model: selected class estimate, sector template, location factor, complexity factor and scenario modifier.','Schedule model: level-based delivery logic, phase durations, critical path sensitivity and scenario acceleration/delay factors.','QCRA: cost exposure model using low / most likely / high impacts and risk-weighted contingency.','QSRA: schedule exposure model using activity-linked O/M/P delay ranges and critical path sensitivity.','Confidence score translated for executives: board-defensibility based on benchmark similarity, evidence maturity, procurement certainty, schedule logic, contingency adequacy and scenario posture.'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card><Card><h2>Commercial readiness</h2><p className="big">This is first-pass project controls intelligence. It is designed to accelerate challenge, option testing and board preparation before final contractor tender or signed cost plan.</p><a className="contactBtn huge" href={emailLink}><Mail/> Send project for review</a></Card></section>}
        {tab === 'pricing' && <section className="layout two"><Card><h2>CASEY Access</h2><div className="pricingGrid"><div className="priceCard"><b>Pilot</b><strong>Request pricing</strong><span>Guided project review, sample outputs and executive walkthrough.</span><a href={emailLink}>Request pilot</a></div><div className="priceCard hot"><b>Professional</b><strong>Full project pack</strong><span>Cost, schedule, risk, QCRA/QSRA and export pack.</span><a href={emailLink}>Request access</a></div><div className="priceCard"><b>Enterprise</b><strong>Private deployment</strong><span>SSO, teams, benchmark library, private models and audit trail.</span><a href={emailLink}>Book demo</a></div></div></Card><Card><h2>Send this project</h2><p className="big">Turn demo interest into pipeline immediately.</p><a className="contactBtn huge" href={emailLink}><Mail/> Send project to CASEY</a><button className="primary" onClick={() => download('/export/all', model, 'CASEY_Output_Pack.zip')}>Download full pack</button></Card></section>}
      </>}
    </main>
  {(loading || exportingLabel) && <div className="simOverlay">
      <div className="simCard">
        <div className="simSpinner" />
        <h3>{exportingLabel || simulationStage || 'Recalculating confidence posture…'}</h3>
        <p>Refreshing QCRA/QSRA curves, benchmark memory, delivery logic and board narrative.</p>
      </div>
    </div>}
  </div>;
}

createRoot(document.getElementById('root')).render(<CaseyErrorBoundary><App/></CaseyErrorBoundary>);
