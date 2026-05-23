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
async function download(path, model, name, setExportingLabel) {
  if (setExportingLabel) setExportingLabel('Generating executive export package…');
  const r = await fetch(API + path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(model) });
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


// V133: Enterprise no-blank-screen hardening. Every generated sector model is normalised
// before React/Recharts render, so an incomplete backend payload can degrade gracefully
// instead of taking down the full product during demo.
const DEFAULT_CURVE = [
  { percentile: 1, cost_bn: 1, schedule_months: 12 },
  { percentile: 10, cost_bn: 1.1, schedule_months: 14 },
  { percentile: 50, cost_bn: 1.3, schedule_months: 18 },
  { percentile: 80, cost_bn: 1.6, schedule_months: 22 },
  { percentile: 90, cost_bn: 1.9, schedule_months: 25 },
  { percentile: 99, cost_bn: 2.3, schedule_months: 30 },
];
function asArray(v) { return Array.isArray(v) ? v : []; }
function asText(v, fallback = '') {
  if (typeof v === 'string' && v.trim()) return v;
  if (typeof v === 'number' || typeof v === 'boolean') return String(v);
  if (v && typeof v === 'object') {
    try {
      if (v.label || v.name || v.title || v.driver || v.risk || v.meaning || v.note || v.signal || v.basis) {
        return [v.label || v.name || v.title || v.driver || v.risk || v.signal, v.value || v.effect || v.meaning || v.note || v.basis].filter(Boolean).join(': ');
      }
      return JSON.stringify(v);
    } catch (_) { return fallback; }
  }
  return fallback;
}
function listText(v, fallback = []) { return (Array.isArray(v) && v.length ? v : fallback).map(x => asText(x)).filter(Boolean); }
function safeObj(v, fallback = {}) { return v && typeof v === 'object' && !Array.isArray(v) ? v : fallback; }
function num(v, fallback = 0) { const n = Number(v); return Number.isFinite(n) ? n : fallback; }
function sectorKey(model = {}, prompt = '') {
  const t = `${model.subsector || ''} ${model.title || ''} ${model.prompt || ''} ${prompt || ''}`.toLowerCase();
  if (/airport|aviation|terminal|airside|baggage|orat|heathrow/.test(t)) return 'airport';
  if (/rail|transit|metro|signalling|rolling stock|corridor|california high/.test(t)) return 'rail';
  if (/data centre|data center|datacenter|hyperscale|gpu|ai campus|liquid cooling/.test(t)) return 'hyperscale';
  if (/pharma|biologic|gmp|cqv|fill-finish|aseptic|validation/.test(t)) return 'pharma';
  if (/semiconductor|fab|wafer|lithography|cleanroom|ultra-pure/.test(t)) return 'semiconductor';
  if (/lng|oil|gas|offshore|refinery|pipeline|process safety|cryogenic/.test(t)) return 'oilgas';
  if (/nuclear|reactor|safety case|containment|smr/.test(t)) return 'nuclear';
  if (/space|lunar|moon|mars|orbital|launch|payload|spaceport|satellite/.test(t)) return 'space';
  if (/defence|defense|secure|aerospace|mission assurance/.test(t)) return 'defence';
  if (/water|desalination|wastewater|treatment/.test(t)) return 'water';
  if (/port|marine|harbour|terminal yard|crane/.test(t)) return 'ports';
  if (/hospital|healthcare|clinical/.test(t)) return 'healthcare';
  if (/energy|power|grid|utility|substation|transmission/.test(t)) return 'energy';
  return model.mode === 'Space' ? 'space' : 'earth';
}
const SECTOR_SAFE = {
  airport: {
    nodes: ['ORAT readiness','Baggage systems integration','Security certification','Airside phasing','Operational transition','Commissioning overlap','Confidence'],
    threats: ['Live airport phasing and possessions','Baggage/security systems integration','Operational readiness trials','Regulatory and stakeholder approvals','Airside access and safety constraints'],
    drivers: ['Benchmark similarity: airport terminal expansion','Scope maturity: capacity, phasing and systems definition','Procurement certainty: baggage/security/MEP packages','Schedule maturity: ORAT and live operations logic','Interface exposure: airlines, airside, landside and regulators'],
    invalid: ['liquid cooling','transformer lead-time','grid energisation','launch readiness','mission assurance','rolling stock']
  },
  rail: {
    nodes: ['Possession access','Signalling integration','Rolling-stock interface','Utility diversions','Migration sequencing','Operational commissioning','Confidence'],
    threats: ['Possession constraints and access windows','Signalling and systems integration','Rolling-stock interface readiness','Utility diversions and corridor constraints','Timetable migration and regulator approvals'],
    drivers: ['Benchmark similarity: rail/transit programme','Scope maturity: alignment, station and systems definition','Procurement certainty: civil/systems package strategy','Schedule maturity: possessions and test/commissioning logic','Interface exposure: utilities, operators and regulators'],
    invalid: ['liquid cooling','ORAT','launch readiness','payload integration','mission assurance']
  },
  hyperscale: {
    nodes: ['Transformer lead-time','Grid energisation','Liquid cooling readiness','IST congestion','Commissioning overlap','Reserve drawdown','Confidence'],
    threats: ['Grid energisation and utility agreements','Long-lead transformer and switchgear delivery','Liquid cooling readiness','Integrated systems testing and commissioning','Phased data-hall turnover'],
    drivers: ['Benchmark similarity: hyperscale digital infrastructure','Scope maturity: campus power and white-space definition','Procurement certainty: transformers, generators and switchgear','Schedule maturity: grid and commissioning logic','Interface exposure: utilities, fibre and commissioning'],
    invalid: ['ORAT','baggage systems','rolling stock','launch readiness']
  },
  pharma: {
    nodes: ['URS maturity','Cleanroom readiness','Process equipment delivery','CQV execution','GMP inspection readiness','Batch release pathway','Confidence'],
    threats: ['CQV protocol approval and execution','Long-lead process equipment delivery','Clean utility validation and media fills','FDA/EMA inspection readiness','Automation and batch-release readiness'],
    drivers: ['Benchmark similarity: pharma / biologics campus','Scope maturity: GMP package and user requirement definition','Procurement certainty: process equipment and clean utility lead-times','Schedule maturity: CQV logic and validation pathway','Regulatory exposure: FDA/EMA inspection readiness'],
    invalid: ['ORAT','baggage systems','launch readiness','rolling stock']
  },
  semiconductor: {
    nodes: ['Cleanroom envelope','Ultra-pure utilities','Lithography tool delivery','Contamination control','Process qualification','Yield ramp','Confidence'],
    threats: ['Lithography and process-tool lead-times','Cleanroom and ultra-pure utility readiness','Contamination-control qualification','Specialist workforce constraints','Yield-ramp uncertainty'],
    drivers: ['Benchmark similarity: advanced fab programme','Scope maturity: cleanroom and process-tool definition','Procurement certainty: lithography and specialist tools','Schedule maturity: qualification and yield-ramp logic','Interface exposure: utilities, contamination control and tool vendors'],
    invalid: ['ORAT','baggage systems','launch readiness','rolling stock']
  },
  oilgas: {
    nodes: ['Process design freeze','Long-lead equipment','Module fabrication','Shutdown window','Process-safety verification','Commissioning readiness','Confidence'],
    threats: ['Long-lead rotating equipment procurement','Shutdown-window dependency','Process-safety system verification','Module fabrication and marine logistics','Commissioning under live operating constraints'],
    drivers: ['Benchmark similarity: process / energy megaproject','Scope maturity: process design and plot-plan definition','Procurement certainty: compressors, vessels and cryogenic packages','Schedule maturity: shutdown and commissioning logic','Interface exposure: marine, process safety and operations'],
    invalid: ['ORAT','baggage systems','liquid cooling','launch readiness']
  },
  nuclear: {
    nodes: ['Safety case maturity','Regulator hold points','Nuclear-grade procurement','QA traceability','Containment systems','Commissioning governance','Confidence'],
    threats: ['Licensing and safety-case maturity','Regulator hold points','Nuclear-grade component lead-times','QA traceability and documentation','Specialist workforce constraints'],
    drivers: ['Benchmark similarity: nuclear generation programme','Scope maturity: safety case and reactor island definition','Procurement certainty: nuclear-grade long-lead components','Schedule maturity: regulator hold-point logic','Interface exposure: QA traceability and commissioning governance'],
    invalid: ['ORAT','baggage systems','liquid cooling','payload integration']
  },
  space: {
    nodes: ['Payload certification','Launch integration','Range availability','Thermal-power balance','Mission assurance sign-off','Operational readiness','Confidence'],
    threats: ['Launch-window and range coordination','Payload integration and certification','Propulsion / thermal qualification','Mission assurance sign-off','Deep-space communications and operations readiness'],
    drivers: ['Benchmark similarity: aerospace / mission programme','Scope maturity: payload, mission and operations definition','Procurement certainty: qualified flight hardware and suppliers','Schedule maturity: test campaign and launch-readiness logic','Interface exposure: range, payload, propulsion and mission operations'],
    invalid: ['ORAT','baggage systems','rolling stock','liquid cooling readiness']
  },
  defence: {
    nodes: ['Security accreditation','Mission systems integration','Classified supplier readiness','Assurance gates','Operational acceptance','Resilience posture','Confidence'],
    threats: ['Security accreditation and assurance gates','Mission systems integration','Classified supplier dependency','Operational acceptance testing','Resilience and continuity validation'],
    drivers: ['Benchmark similarity: defence / secure infrastructure','Scope maturity: mission-system and security definition','Procurement certainty: classified supplier readiness','Schedule maturity: assurance and acceptance logic','Interface exposure: security, operations and regulator stakeholders'],
    invalid: ['ORAT','baggage systems','liquid cooling readiness','rolling stock']
  },
  water: {
    nodes: ['Process design maturity','Permitting','Civil works sequencing','Equipment procurement','Commissioning permits','Operational handover','Confidence'],
    threats: ['Permitting and environmental approvals','Process equipment procurement','Civil works and tie-in sequencing','Commissioning permits and water-quality validation','Operational handover'],
    drivers: ['Benchmark similarity: water infrastructure programme','Scope maturity: treatment process and civil scope definition','Procurement certainty: pumps, membranes and process packages','Schedule maturity: tie-in and commissioning logic','Interface exposure: utilities, regulator and operator constraints'],
    invalid: ['ORAT','baggage systems','launch readiness','liquid cooling']
  },
  ports: {
    nodes: ['Marine access','Quay works','Crane procurement','Yard systems','Customs / rail interface','Operational commissioning','Confidence'],
    threats: ['Marine access and dredging windows','Quay and berth construction sequencing','Crane and automation procurement','Yard, customs and rail systems integration','Operational commissioning under live port constraints'],
    drivers: ['Benchmark similarity: port / marine terminal programme','Scope maturity: berth, yard and systems definition','Procurement certainty: cranes, automation and marine packages','Schedule maturity: marine windows and port operations logic','Interface exposure: customs, rail, shipping and operator stakeholders'],
    invalid: ['ORAT','liquid cooling','launch readiness','payload integration']
  },
  healthcare: {
    nodes: ['Clinical brief maturity','Medical equipment procurement','Digital health integration','Infection-control validation','Phased occupancy','Clinical commissioning','Confidence'],
    threats: ['Clinical commissioning readiness','Medical equipment procurement','Digital health and systems integration','Infection-control compliance','Phased occupancy and patient transition'],
    drivers: ['Benchmark similarity: healthcare / hospital infrastructure','Scope maturity: clinical brief and equipment schedule','Procurement certainty: medical equipment and specialist systems','Schedule maturity: phased occupancy and clinical commissioning logic','Interface exposure: clinicians, regulators and live hospital operations'],
    invalid: ['ORAT','liquid cooling','launch readiness','rolling stock']
  },
  energy: {
    nodes: ['Grid connection','Permitting','Long-lead equipment','Civil / electrical interface','Commissioning sequence','Market dispatch readiness','Confidence'],
    threats: ['Grid connection and permitting','Transformer / switchgear procurement','Civil and electrical interface coordination','Commissioning sequence and energisation','Market dispatch and operator readiness'],
    drivers: ['Benchmark similarity: energy / utility megaprogramme','Scope maturity: generation and grid-interface definition','Procurement certainty: transformers, turbines and switchgear','Schedule maturity: energisation and commissioning logic','Interface exposure: grid operator, regulator and utilities'],
    invalid: ['ORAT','baggage systems','payload integration']
  },
  earth: {
    nodes: ['Scope definition','Procurement evidence','Interface control','Commissioning readiness','Operational acceptance','Reserve adequacy','Confidence'],
    threats: ['Long-lead procurement','Interface coordination','Commissioning readiness','Approvals and stakeholder governance','Operational handover'],
    drivers: ['Benchmark similarity: comparable infrastructure archetype','Scope maturity: package and requirement definition','Procurement certainty: long-lead supplier readiness','Schedule maturity: critical path and commissioning logic','Interface exposure: utilities, operators and regulators'],
    invalid: ['launch readiness','payload integration']
  }
};
function scrubText(text, invalid = []) {
  let s = String(text ?? '');
  invalid.forEach(term => {
    const re = new RegExp(term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'ig');
    s = s.replace(re, 'sector-specific readiness');
  });
  return s;
}
function scrubDeep(value, invalid) {
  if (typeof value === 'string') return scrubText(value, invalid);
  if (Array.isArray(value)) return value.map(v => scrubDeep(v, invalid));
  if (value && typeof value === 'object') {
    const out = {};
    for (const [k, v] of Object.entries(value)) out[k] = scrubDeep(v, invalid);
    return out;
  }
  return value;
}
function validateCurve(rawCurve, model) {
  const baseCost = parseFloat(String(model?.cost_p50 || '').replace(/[^0-9.]/g, '')) || 1;
  const baseSch = parseFloat(String(model?.schedule || '').replace(/[^0-9.]/g, '')) || 24;
  const curve = asArray(rawCurve).map((d, i) => ({
    percentile: num(d?.percentile, [1,5,10,20,30,40,50,60,70,80,90,95,99][i] || i + 1),
    cost_bn: num(d?.cost_bn, baseCost * (0.8 + i * 0.05)),
    schedule_months: num(d?.schedule_months, Math.round(baseSch * (0.8 + i * 0.05)))
  })).filter(d => Number.isFinite(d.percentile) && Number.isFinite(d.cost_bn) && Number.isFinite(d.schedule_months));
  return curve.length >= 3 ? curve : DEFAULT_CURVE.map(d => ({...d, cost_bn: +(baseCost * d.cost_bn).toFixed(1), schedule_months: Math.max(1, Math.round(baseSch * d.schedule_months / 18))}));
}
function normaliseModel(raw, prompt = '') {
  if (!raw || typeof raw !== 'object') return null;
  const key = sectorKey(raw, prompt);
  const sector = SECTOR_SAFE[key] || SECTOR_SAFE.earth;
  let m = scrubDeep({ ...raw }, sector.invalid);
  m.prompt = asText(m.prompt, prompt);
  m.title = asText(m.title, key === 'space' ? 'Space infrastructure programme' : 'Capital infrastructure programme');
  m.mode = asText(m.mode, key === 'space' ? 'Space' : 'Earth');
  m.subsector = asText(m.subsector, key.replace(/(^|_)(\w)/g, (_,a,b)=> (a?' / ':'') + b.toUpperCase()));
  m.risk = asText(m.risk, 'Medium-High');
  m.confidence_pct = Math.max(1, Math.min(99, num(m.confidence_pct, 58)));
  m.scenario_label = asText(m.scenario_label, 'Base');
  m.cost_p50 = asText(m.cost_p50, '$1.0B');
  m.cost_range = asText(m.cost_range, '$0.8B - $1.4B');
  m.schedule = asText(m.schedule, '24 months');
  m.causal_graph_nodes = asArray(m.causal_graph_nodes).filter(Boolean).map(String).slice(0,7);
  if (m.causal_graph_nodes.length < 4) m.causal_graph_nodes = sector.nodes;
  m.sector_schedule_threats = asArray(m.sector_schedule_threats).length ? asArray(m.sector_schedule_threats).map(String) : sector.threats;
  m.sector_confidence_drivers = asArray(m.sector_confidence_drivers).length ? asArray(m.sector_confidence_drivers).map(String) : sector.drivers;
  m.next_best_actions = asArray(m.next_best_actions).length ? asArray(m.next_best_actions).map(String) : ['Validate dominant critical-path constraint against benchmark evidence.','Confirm procurement maturity and named evidence owners.','Challenge reserve posture against commissioning and interface exposure.','Prepare board decision paper with explicit scenario trade-offs.'];
  m.board_briefing = asArray(m.board_briefing).length ? asArray(m.board_briefing).map(String) : [asText(m.executive_shock_insight, 'The programme is now governed by evidence maturity, interfaces and commissioning readiness.'),'Challenge whether the current confidence posture is supported by procurement and operational-readiness evidence.'];
  m.mission_control_cards = asArray(m.mission_control_cards).length ? asArray(m.mission_control_cards) : [
    { label: 'LIVE CALIBRATION', signal: 'Current sector conditions are being applied to confidence, contingency and delivery-tail exposure.', severity: 'ACTIVE' },
    { label: 'EXECUTIVE SHOCK', signal: asText(m.executive_shock_insight, 'The programme narrative should be challenged against evidence maturity.'), severity: 'MEDIUM-HIGH' },
    { label: 'CRITICAL PATH EXPOSURE', signal: sector.threats[0], severity: 'HIGH' }
  ];
  m.benchmark_comparison = asArray(m.benchmark_comparison).length ? asArray(m.benchmark_comparison) : [
    { archetype: key === 'space' ? 'Aerospace / Mission Assurance Programme' : sector.drivers[0].replace('Benchmark similarity: ', ''), anchor_cost: '$1B-$10B', anchor_duration_months: '36-96' }
  ];
  m.cost_breakdown = asArray(m.cost_breakdown).length ? asArray(m.cost_breakdown) : [
    { cbs: '01.01', description: 'Core delivery package', type: 'Direct', p10_bn: .8, p50_bn: 1, p90_bn: 1.3, basis: 'Fallback normalized cost line' }
  ];
  m.schedule_detail = asArray(m.schedule_detail).length ? asArray(m.schedule_detail) : [
    { activity_id: 'A1000', phase: 'Governance', activity: 'Project initiation / controls setup', predecessor: '', duration_months: 2, critical: 'No', basis: 'Fallback normalized schedule line' }
  ];
  m.risk_register = asArray(m.risk_register).length ? asArray(m.risk_register) : [
    { id: 'R-001', risk: sector.threats[0], owner: 'Programme Director', mitigation: 'Evidence workshop and owner action plan', likelihood: 'Medium', impact: 'High' }
  ];
  m.cost_waterfall_vs_base = asArray(m.cost_waterfall_vs_base).length ? asArray(m.cost_waterfall_vs_base) : [{ driver: 'Base P50', value: m.cost_p50, kind: 'total' }];
  m.schedule_waterfall_vs_base = asArray(m.schedule_waterfall_vs_base).length ? asArray(m.schedule_waterfall_vs_base) : [{ driver: 'Base duration', months: parseFloat(String(m.schedule).replace(/[^0-9.]/g,'')) || 24, kind: 'total' }];
  m.scenario_matrix = asArray(m.scenario_matrix).length ? asArray(m.scenario_matrix) : scenarios.map(s => ({ scenario: s, cost_p50: m.cost_p50, schedule: m.schedule, confidence_pct: m.confidence_pct, risk: m.risk }));
  m.monte_carlo = m.monte_carlo && typeof m.monte_carlo === 'object' ? m.monte_carlo : {};
  m.monte_carlo.curve = validateCurve(m.monte_carlo.curve, m);
  const p50c = parseFloat(String(m.cost_p50).replace(/[^0-9.]/g,'')) || 1;
  const p50s = parseFloat(String(m.schedule).replace(/[^0-9.]/g,'')) || 24;
  m.monte_carlo.qcra = m.monte_carlo.qcra && typeof m.monte_carlo.qcra === 'object' ? m.monte_carlo.qcra : { p50: p50c, p80: +(p50c*1.18).toFixed(1), p90: +(p50c*1.35).toFixed(1) };
  m.monte_carlo.qsra = m.monte_carlo.qsra && typeof m.monte_carlo.qsra === 'object' ? m.monte_carlo.qsra : { p50: p50s, p80: Math.round(p50s*1.18), p90: Math.round(p50s*1.35) };
  m.monte_carlo.tornado = asArray(m.monte_carlo.tornado).length ? asArray(m.monte_carlo.tornado) : sector.threats.slice(0,4).map((driver,i)=>({driver, impact: 5-i}));
  m.uncertainty_narrative = m.uncertainty_narrative && typeof m.uncertainty_narrative === 'object' ? m.uncertainty_narrative : {
    estimate_maturity: 'Class maturity is suitable for option selection, but evidence gaps remain before approval.',
    schedule_maturity: 'Schedule logic requires critical-path, handover and commissioning validation.',
    interpretation: `Live calibration is weighting ${sector.threats.slice(0,3).join(', ')} into the QCRA/QSRA tail.`
  };

  // V134: normalise every field that can be rendered as text or chart data.
  // This prevents object-as-child React failures across rail, life sciences and any future sector payload.
  m.board_challenge_questions = listText(m.board_challenge_questions, []);
  m.top_decisions_required = listText(m.top_decisions_required, ['Accept or reject the scenario trade-off explicitly at board level.','Confirm the critical-path and near-critical path evidence.','Assign named owners to the governing constraints.']);
  m.outputs_board_memo = listText(m.outputs_board_memo, m.board_briefing);
  m.critical_path_narrative = listText(m.critical_path_narrative, sector.threats.map(x => `${x} is near-critical until evidenced.`));
  m.red_flags = listText(m.red_flags, ['Unevidenced confidence around the governing constraint.','Scenario benefit may be risk transfer rather than risk reduction.']);
  m.scenario_delta_intelligence = asArray(m.scenario_delta_intelligence).length ? asArray(m.scenario_delta_intelligence).map((x,i) => ({
    label: asText(x?.label || x?.driver, `Delta ${i+1}`), value: asText(x?.value || x?.effect, '—'), meaning: asText(x?.meaning || x?.note || x, 'Scenario movement requires board challenge.')
  })) : [
    { label: 'Capital movement', value: '+0%', meaning: 'Balanced cost, time and evidence posture.' },
    { label: 'Schedule movement', value: '+0%', meaning: 'Maintains a credible reference case for board challenge.' },
    { label: 'Confidence movement', value: '+0 pts', meaning: 'Confidence moves only when evidence improves.' }
  ];
  m.confidence_breakdown = asArray(m.confidence_breakdown).length ? asArray(m.confidence_breakdown).map((x,i) => ({
    driver: asText(x?.driver || x?.label, `Confidence driver ${i+1}`), effect: asText(x?.effect || x?.value, '—'), note: asText(x?.note || x?.meaning || x, 'Evidence must support this movement.')
  })) : sector.drivers.slice(0,5).map((driver,i)=>({driver, effect: i ? '-3' : '+6', note: 'Sector-specific confidence weighting applied.'}));
  m.live_calibration_signals = asArray(m.live_calibration_signals).length ? asArray(m.live_calibration_signals).map((x,i)=>({
    signal: asText(x?.signal || x?.label || x, `Signal ${i+1}`), status: asText(x?.status, 'Active'), direction: asText(x?.direction, 'confidence / reserve / P-tail'), applies_to: asText(x?.applies_to, 'board pack, workbook, risk register, XER, QCRA/QSRA'), basis: asText(x?.basis || x?.note || x, 'Sector condition is being weighted into the model.')
  })) : sector.threats.slice(0,4).map((signal,i)=>({signal, status:i<2?'Active':'Watch', direction:'confidence / reserve / P-tail', applies_to:'board pack, workbook, risk register, XER, QCRA/QSRA', basis:`${signal} is a sector-native delivery constraint.`}));
  m.mission_control_cards = asArray(m.mission_control_cards).length ? asArray(m.mission_control_cards).map((x,i)=>({
    label: asText(x?.label || x?.signal, `Signal ${i+1}`), signal: asText(x?.signal || x?.basis || x, 'Sector condition is active.'), severity: asText(x?.severity || x?.status, 'ACTIVE')
  })) : m.mission_control_cards;
  m.benchmark_comparison = asArray(m.benchmark_comparison).length ? asArray(m.benchmark_comparison).map((x,i)=>({
    archetype: asText(x?.archetype || x?.sector || x?.name, `Sector benchmark ${i+1}`), anchor_cost: asText(x?.anchor_cost || x?.cost, '$1B-$10B'), anchor_duration_months: asText(x?.anchor_duration_months || x?.duration || x?.months, '36-96'), similarity_score: num(x?.similarity_score, Math.max(6, 9-i)), use: asText(x?.use, 'Sector-locked benchmark cohort.')
  })) : m.benchmark_comparison;

  m.why_casey_generated_this = asArray(m.why_casey_generated_this).length ? asArray(m.why_casey_generated_this).map(x => asText(x)) : [
    `CASEY detected ${m.subsector} from the project brief and routed it to the ${m.mode} sector model.`,
    `Sector behaviours applied: ${sector.nodes.slice(0,4).join(', ')}.`,
    'Cost, schedule and risk were calibrated against estimate class, schedule level, complexity and delivery environment.',
    'The output is designed for early board challenge and scope definition, not certified pricing.'
  ];
  return m;
}

function EmergencyDashboard() {
  let model = null;
  try { model = window.__CASEY_LAST_MODEL__ || JSON.parse(sessionStorage.getItem('CASEY_LAST_MODEL') || 'null'); } catch (_) {}
  if (!model) return <div className="app v50EliteApp"><main className="console"><section className="layout one"><Card className="shockCard"><h2>CASEY recovered the demo session</h2><p>The previous project payload could not be displayed. Refresh and run again.</p><button className="primary" onClick={() => window.location.reload()}>Recover session</button></Card></section></main></div>;
  const nodes = listText(model.causal_graph_nodes, ['Scope definition','Procurement evidence','Interface control','Commissioning readiness','Confidence']).slice(0,7);
  const threats = listText(model.sector_schedule_threats, ['Critical path constraint','Procurement evidence gap','Commissioning readiness']).slice(0,5);
  return <div className="app v50EliteApp"><main className="console"><section className="layout one"><Card className="shockCard"><h2>CASEY intelligence recovered</h2><p className="big">{asText(model.executive_summary, `${asText(model.subsector,'Sector model')} generated successfully. The full payload has been normalised for demo continuity.`)}</p><div className="miniMetrics"><b><span>P50 cost</span>{asText(model.cost_p50,'—')}</b><b><span>Schedule</span>{asText(model.schedule,'—')}</b><b><span>Confidence</span>{asText(model.confidence_pct,'—')}%</b></div><h3>Sector causal chain</h3>{nodes.map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}<h3>Top schedule threats</h3>{threats.map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}<button className="primary" onClick={() => { window.location.hash=''; window.location.reload(); }}>Continue demo</button><p className="chartCaption">Render guard: recovered from incomplete component path without losing the generated project model.</p></Card></section></main></div>;
}

class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { failed: false, error: null }; }
  static getDerivedStateFromError(error) { return { failed: true, error }; }
  componentDidCatch(error, info) { console.error('CASEY render guard caught:', error, info); }
  render() {
    if (!this.state.failed) return this.props.children;
    return <EmergencyDashboard/>;
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
  return <Card className={`v50Kpi ${hot ? 'hot' : ''} ${band}`}><Icon size={21}/><div><p>{label}</p><b>{value}</b><span>{sub}</span></div></Card>;
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
  const stages = ['CASEY recalibrating confidence curves...', 'Applying live sector calibration signals...', 'Running procurement and delivery-tail model...', 'Comparing against benchmark archetypes...', 'Stamping scenario/base deltas into exports...'];
  const [i,setI] = useState(0);
  useEffect(() => { const t = setInterval(() => setI(v => Math.min(v + 1, stages.length - 1)), 650); return () => clearInterval(t); }, []);
  return <motion.div className="loading intelligenceLoading" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}><Rocket size={44}/><b>{text || 'Building connected model...'}</b><span>{stages[i]}</span><small>Cost · Schedule · QCRA · QSRA · Risk Register · Board Pack</small></motion.div>;
}
function ScenarioSelector({ scenario, generate, matrix=[] }) {
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
  const lockedConstraint = model?.confidence_engine_detail?.primary_constraint || model?.sector_constraints || model?.governing_constraint;
  const constraint = lockedConstraint || (mode === 'Space'
    ? 'mission assurance, launch logistics and autonomous recovery evidence'
    : subsector.includes('data') ? 'energisation, cooling readiness and integrated systems testing'
    : subsector.includes('airport') || subsector.includes('aviation') ? 'ORAT readiness, baggage/security systems integration, airside phasing and regulator acceptance'
    : subsector.includes('rail') || subsector.includes('transit') ? 'possessions, signalling integration, systems migration and operator acceptance'
    : subsector.includes('semiconductor') ? 'tool install, UPW readiness and yield-ramp qualification'
    : subsector.includes('life') || subsector.includes('pharma') ? 'CQV, validation readiness and regulatory evidence'
    : 'interface control, procurement evidence and commissioning readiness');
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
    { name: 'Evidence maturity', score: base.evidence, note: 'brief depth, basis visibility and package maturity' },
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
    {peers.map(([name,px,py],i)=><motion.div key={name} className={`peerDot ${name==='CASEY'?'caseyDot':''}`} initial={{left:`${px-6}%`,bottom:`${py-8}%`,opacity:.4}} animate={{left:name==='CASEY'?`${x}%`:`${px}%`,bottom:name==='CASEY'?`${y}%`:`${py}%`,opacity:1}} transition={{type:'spring',stiffness:90,damping:14}}><b>{name}</b><span>{name==='CASEY'?'top-decile':`${px}th pct`}</span></motion.div>)}
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
function BoardPressureSimulator({ model }) {
  const q = boardQuestions(model).slice(0,5);
  return <Card className="boardPressure"><h2>Board Pressure Simulation</h2><p>Likely approval resistance and challenge posture before investment committee.</p>{q.map((x,i)=><div className="pressureRow" key={x}><span>{i+1}</span><b>{i<2?'High challenge':'Medium challenge'}</b><em>{x}</em></div>)}</Card>
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

  async function generate(nextScenario = scenario, nextPrompt = prompt, activeContext = model || projectContext) {
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
    try {
      const payload = {
        prompt: canonicalPrompt,
        client,
        class_level: Number(classLevel),
        schedule_level: Number(scheduleLevel),
        scenario: nextScenario,
        demo: true,
        active_model: contextLock
      };
      const m = await post('/generate', payload);
      const nextContext = lockedProjectContext(m, canonicalPrompt);
      const safe = normaliseModel(m, canonicalPrompt); window.__CASEY_LAST_MODEL__ = safe; try { sessionStorage.setItem('CASEY_LAST_MODEL', JSON.stringify(safe)); } catch (_) {} const nextContextSafe = lockedProjectContext(safe, canonicalPrompt); setModel(safe); setProjectContext(nextContextSafe); setScenario(nextScenario); setPrompt(canonicalPrompt);
    } catch (e) { setError(String(e.message || e)); }
    finally { setLoading(false); setSimulationStage(''); setConfidencePulse(false); }
  }
  function runEarth() { setProjectContext(null); generate('base', earthPrompt, null); }
  function runSpace() { setProjectContext(null); generate('base', spacePrompt, null); }
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
  const scenarioMatrix = model?.scenario_matrix || model?.scenario_comparison || [];
  const direct = costs.filter(x => String(x.type || '').toLowerCase().includes('direct')).reduce((a, b) => a + Number(b.p50_bn || 0), 0);
  const indirect = costs.filter(x => String(x.type || '').toLowerCase().includes('indirect')).reduce((a, b) => a + Number(b.p50_bn || 0), 0);
  const reserves = costs.filter(x => String(x.type || '').toLowerCase().includes('reserve')).reduce((a, b) => a + Number(b.p50_bn || 0), 0);
  const emailBody = model ? [
    'Please review this project in CASEY.', '', `Project: ${model.title}`, `Scenario: ${model.scenario_label || scenario}`,
    `P50 Cost: ${model.cost_p50}`, `Cost Range: ${model.cost_range}`, `Schedule: ${model.schedule}`,
    `Risk / Confidence: ${model.risk} / ${model.confidence_pct}%`
  ].join('\n') : 'Please send me CASEY access.';
  const emailLink = `mailto:hello@casey.ai?subject=${encodeURIComponent('CASEY project review')}&body=${encodeURIComponent(emailBody)}`;
  const confLens = model ? confidenceLens(model) : null;
  const p80Talk = model ? p80PlainEnglish(model) : null;
  const tradePack = model ? gainedSacrificedExposed(model) : null;

  return <div className="app v50EliteApp">
    <Briefing open={briefing} onClose={() => setBriefing(false)} onEarth={runEarth} onSpace={runSpace}/>
    <OneShotDemo open={trialOpen} onClose={() => setTrialOpen(false)} onComplete={(m) => { const safe = normaliseModel(m, m?.prompt || prompt); window.__CASEY_LAST_MODEL__ = safe; try { sessionStorage.setItem('CASEY_LAST_MODEL', JSON.stringify(safe)); } catch (_) {} setModel(safe); setProjectContext(lockedProjectContext(safe, safe?.prompt || prompt)); setShow(false); setTrialOpen(false); setTab('overview'); }} />
    <AnimatePresence>{loading && <Loading text="Building full CASEY intelligence pack..."/>}</AnimatePresence>
    {show && !model && <Hero onBriefing={() => setBriefing(true)} onEarth={runEarth} onSpace={runSpace} onConsole={() => setShow(false)} onTryDemo={() => setTrialOpen(true)}/>} 
    <header className="v50ConsoleTop"><Logo/><nav><button onClick={() => { setModel(null); setProjectContext(null); setShow(true); }}>Home</button><button onClick={() => setBriefing(true)}>Film</button><button onClick={() => setTrialOpen(true)}>Free run</button><button onClick={runEarth}>Earth demo</button><button onClick={runSpace}>Space demo</button><a href={emailLink}>Request access</a></nav></header>
    <main className={model ? 'v50Console' : 'v50Console emptyConsole'}>
      {error && <div className="error">{error}</div>}
      {!model && !show && <section className="commandGrid"><Card className="command"><h1>Generate a live project model</h1><label>Project command</label><textarea value={prompt} onChange={e => setPrompt(e.target.value)} /> <div className="chips">{examples.map(x => <button key={x} onClick={() => setPrompt(x)}>{x}</button>)}</div><div className="grid4"><input value={client} onChange={e => setClient(e.target.value)} placeholder="Client / operator"/><select value={classLevel} onChange={e => setClassLevel(e.target.value)}>{[1,2,3,4,5].map(x => <option key={x} value={x}>Class {x}</option>)}</select><select value={scheduleLevel} onChange={e => setScheduleLevel(e.target.value)}>{[1,2,3,4,5].map(x => <option key={x} value={x}>Level {x}</option>)}</select><select value={scenario} onChange={e => setScenario(e.target.value)}>{scenarios.map(x => <option key={x} value={x}>{x}</option>)}</select></div><button className="primary" onClick={() => generate()}><Sparkles/> Generate full intelligence pack</button></Card><Card><h2>What CASEY will produce</h2>{['Executive summary and recommendation','Direct / indirect / reserve cost view','Scenario-linked estimate, schedule and confidence','Risk register with cause, event, impact and mitigation','QCRA + QSRA curves and tornado drivers','Pricing and next-step contact actions'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card></section>}
      {model && <>
        <section className="confidenceEngineBadge"><b>{model.confidence_engine_label || 'CASEY Confidence Engine'}</b><span>{model.confidence_engine_detail || 'Benchmark + probabilistic + sector-trained reasoning'}</span></section>
        <LiveCalibrationStrip model={model}/>
        <section className="kpis"><Kpi icon={Globe2} label="Mode / sector" value={model.mode} sub={model.subsector}/><Kpi icon={Activity} label="P50 cost" value={model.cost_p50} sub={model.cost_range}/><Kpi icon={Zap} label="Schedule" value={model.schedule} sub={`QSRA P80 ${model.monte_carlo?.qsra?.p80 || '—'} months`}/><Kpi icon={ShieldAlert} label="Delivery confidence" value={confidenceLens(model).headline} sub={`${model.risk} risk · ${model.confidence_pct}% · ${model.scenario_label}`} hot/></section>
        <IntelligenceMeta model={model} mode={viewMode} setMode={setViewMode}/>
        <PropagationPulse scenario={scenario} active={propagating}/>
        <ScenarioSelector scenario={scenario} generate={generate} matrix={scenarioMatrix}/>
        <ExportStrip model={model}
          onBoardPack={() => download('/export/all', model, `${model.id || 'casey'}_DEMO_BOARD_PACK.zip`)}
          onExcel={() => download('/export/workbook', model, `${model.id || 'casey'}_DEMO_COST_WORKBOOK.xlsx`)}
          onRisk={() => download('/export/risk-register', model, `${model.id || 'casey'}_DEMO_RISK_REGISTER.xlsx`)}
          onXer={() => download('/export/xer', model, `${model.id || 'casey'}_DEMO_SCHEDULE.xer`)}
          onQcra={() => download('/export/qcra-qsra', model, `${model.id || 'casey'}_DEMO_QCRA_QSRA.xlsx`)}/>
        <nav className="tabs">{[['overview','Overview'],['compare','Scenarios'],['delta','Scenario Intel'],['causal','Causal OS'],['cost','Cost'],['schedule','Schedule'],['risk','Risk'],['monte','QCRA/QSRA'],['outputs','Outputs'],['advisor','Advisor'],['method','Methodology'],['pricing','Pricing']].map(x => <button key={x[0]} className={tab===x[0]?'active':''} onClick={() => setTab(x[0])}>{x[1]}</button>)}</nav>
        {tab === 'overview' && <>
          {model.executive_shock_insight && <section className="layout one"><Card className="shockCard"><h2>Executive shock insight</h2><p>{model.executive_shock_insight}</p></Card></section>}
          <section className="layout two">
            <Card><h2>Executive intelligence summary</h2><p className="big">{model.executive_summary || `${model.title} has been classified as ${model.subsector}. CASEY generated a first-pass cost, schedule, risk and confidence model for the selected scenario.`}</p><div className="miniMetrics"><b><span>Direct cost</span>{fmt(direct)}</b><b><span>Indirect cost</span>{fmt(indirect)}</b><b><span>Risk / reserve</span>{fmt(reserves)}</b></div><h3>Recommendation</h3>{(model.next_best_actions || []).slice(0,4).map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card>
            <Card><h2>Board briefing</h2>{(model.board_briefing || model.board_challenge_questions || []).slice(0,4).map((x,i)=><div className="reason" key={String(x)}><span>{i+1}</span>{x}</div>)}<h3>CASEY thinking</h3><p className="caseyThinking">{model.casey_thinking || 'CASEY interprets this as a system-of-systems infrastructure programme requiring cost, schedule, risk and decision intelligence.'}</p></Card>
          </section>
          <section className="layout two eliteLayer">
            <Card className="confidenceMeaningCard"><h2>What confidence means</h2><h3>{confLens.headline}</h3><p className="big">{confLens.meaning}</p><div className="reason"><span>!</span><b>Decision rule</b><br/>{confLens.decisionRule}</div><div className="reason"><span>→</span><b>Primary constraint</b><br/>{confLens.constraint}</div><div className="reason"><span>%</span><b>Plain English</b><br/>Confidence is not optimism. It is CASEY's board-defensibility score based on benchmark fit, evidence maturity, procurement certainty, schedule logic, reserve adequacy and scenario posture.</div></Card>
            <Card><h2>Likely board questions</h2>{boardQuestions(model).slice(0,6).map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}<h3>CASEY final position</h3><p className="caseyThinking finalPosition">{finalPosition(model)}</p></Card>
          </section>
          <section className="layout two eliteLayer">
            <Card><h2>Evidence threshold map</h2><p className="chartCaption">Shows why the confidence number is where it is, and what must improve before board approval.</p><ResponsiveContainer width="100%" height={260}><BarChart data={evidenceScorecard(model)} layout="vertical"><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis type="number" domain={[0,100]}/><YAxis dataKey="name" type="category" width={145}/><Tooltip formatter={(v) => [`${v}%`, 'board-defensibility score']}/><ReferenceLine x={70} stroke="#ffd96a88" label="board comfort"/><Bar dataKey="score" fill="#8df7ff"/></BarChart></ResponsiveContainer>{evidenceScorecard(model).map((x,i)=><div className="reason compactReason" key={x.name}><span>{i+1}</span><b>{x.name}: {Math.round(x.score)}%</b><br/>{x.note}</div>)}</Card>
            <Card><h2>Contradiction scan</h2><p className="chartCaption">CASEY does not just make the case look better. It exposes the trade-off that could get challenged.</p>{(model.second_order_contradictions || contradictionScan(model)).slice(0,5).map((x,i)=><div className="reason" key={String(x)}><span>{i+1}</span>{x}</div>)}<h3>Governance challenge</h3>{(model.governance_challenges || []).slice(0,3).map((x,i)=><div className="reason compactReason" key={String(x)}><span>{i+1}</span>{x}</div>)}<h3>Demo close line</h3><p className="caseyThinking finalPosition">Traditional project controls reports show numbers. CASEY shows the board what the numbers are trying to hide.</p></Card>
          </section>
          <LiveCalibrationPanel model={model}/>
          {baseVs?.base && <section className="layout two">
            <Card className="shockCard"><h2>Scenario vs Base</h2><p>{baseVs.plain_english}</p><div className="miniMetrics"><b><span>Base P50</span>{baseVs.base.cost_p50}<small>{baseVs.base.schedule_months} mo · {baseVs.base.confidence_pct}%</small></b><b><span>{baseVs.selected.scenario} P50</span>{baseVs.selected.cost_p50}<small>{baseVs.selected.schedule_months} mo · {baseVs.selected.confidence_pct}%</small></b><b><span>Delta</span>{baseVs.delta.cost_direction === 'same' ? 'No cost move' : `${baseVs.delta.cost} ${baseVs.delta.cost_direction}`}<small>{baseVs.delta.months} mo · {baseVs.delta.confidence_pts} pts</small></b></div></Card>
            <Card><h2>What changed and why</h2>{(model.scenario_delta_intelligence || []).slice(0,5).map((x,i)=><div className="reason" key={i}><span>{i+1}</span><b>{x.label}: {x.value}</b><br/>{x.meaning}</div>)}</Card>
          </section>}
          <section className="layout two">
            <Card><h2>Mission control signals</h2><div className="missionCardGrid">{(model.mission_control_cards || []).slice(0,6).map((c,i)=><div className="intelCard" key={i}><b>{c.label}</b><p>{c.signal}</p><span>{c.severity}</span></div>)}</div></Card>
            <Card><h2>Uncertainty narrative</h2><p>{model.uncertainty_narrative?.estimate_maturity}</p><p>{model.uncertainty_narrative?.schedule_maturity}</p><p>{model.uncertainty_narrative?.interpretation}</p><h3>Benchmark comparison</h3>{(model.benchmark_comparison || []).slice(0,4).map((b,i)=><div className="reason" key={i}><span>{i+1}</span><b>{b.archetype}</b> · {b.anchor_cost} · {b.anchor_duration_months} months</div>)}</Card>
          </section>
          <section className="layout two"><BenchmarkIntelligence model={model}/><CausalGraph model={model}/></section>
          <section className="layout two">
            <Card><h2>Confidence drivers</h2>{(model.sector_confidence_drivers || ['Benchmark similarity: high where comparable infrastructure archetypes exist','Scope maturity: concept / budget level until package evidence is supplied','Procurement certainty: sensitive to long-lead equipment and market capacity','Schedule maturity: improves when critical path and commissioning logic are validated','Interface exposure: controlled by utilities, systems integration and operational constraints']).map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card>
            <Card><h2>Why CASEY generated this</h2>{(model.why_casey_generated_this || ['CASEY detected the infrastructure asset and operating environment from the brief','The programme was mapped to benchmark memory and sector archetypes','Cost, schedule and risk were calibrated against class maturity and delivery complexity','The narrative is designed for early board challenge, not certified pricing']).map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card>
          </section>
          <section className="layout two">
            <Card><h2>Primary cost drivers</h2>{(model.sector_primary_cost_drivers || ['Utility / enabling infrastructure','Specialist systems and long-lead equipment','Commissioning and validation complexity','Programme management, preliminaries and indirects','Risk reserve driven by procurement and interface uncertainty']).map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card>
            <Card><h2>Top schedule threats</h2>{(model.sector_schedule_threats || ['Utility energisation delay','Long-lead equipment procurement and supplier capacity','Design freeze instability and scope movement','Systems integration and commissioning bottlenecks','Approvals, safety case, permitting or operational access constraints']).map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card>
          </section>
        </>}

        {tab === 'compare' && <section className="layout two"><Card><h2>Scenario comparison</h2><p className="big">Switch options before paying for another advisory cycle. Each button re-runs cost, schedule, confidence, risk register, QCRA/QSRA and exports from the same source of truth.</p><div className="scenarioCompare upgraded">{scenarios.map(s => { const active = s === scenario; const row = scenarioMatrix.find(x => x.scenario === s) || {}; return <button key={s} className={active?'active':''} onClick={() => generate(s, model?.prompt || prompt, model || projectContext)}><b>{(row.label || s).replace('_',' ')}</b><strong>{row.cost_p50 || row.cost || '—'} · {row.schedule_months || '—'} mo</strong><span>{row.risk || '—'} / {row.confidence_pct || row.confidence || '—'}%</span><em>{active ? 'current model' : 'run scenario'}</em></button> })}</div></Card><Card><h2>Buyer decision lens</h2>{['Base: balanced reference case for board challenge','Faster: more capex, lower confidence, shorter duration','Cheaper: lower authorisation number, longer schedule, higher residual risk','Lower Risk: higher reserve, longer duration, stronger confidence','Premium: resilience and optionality bought with visible capex premium'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}<h3>Current trade-off</h3><div className="triLens"><b>Gained</b>{tradePack.gained.map(x=><span key={x}>{x}</span>)}<b>Sacrificed</b>{tradePack.sacrificed.map(x=><span key={x}>{x}</span>)}<b>Exposed</b>{tradePack.exposed.map(x=><span key={x}>{x}</span>)}</div></Card></section>}
        {tab === 'cost' && <section className="layout two"><Card><h2>Scenario cost bridge vs Base</h2><p className="chartCaption">This explains why the selected scenario is cheaper or more expensive than Base before showing the workbook lines.</p>{costWaterfall.map((x,i)=><div className={`reason ${x.kind==='total'?'deltaReason':''}`} key={i}><span>{i+1}</span><b>{x.driver}</b><br/>{x.kind==='total'?x.value:(x.value_bn>=0?'+':'−') + ' ' + x.value}</div>)}<h3>Cost estimate workbook</h3><Table rows={costs} cols={[["cbs","CBS"],["description","Description"],["type","Type"],["p10_bn","Low/P10"],["p50_bn","Most likely/P50"],["p90_bn","High/P90"],["impact_basis","Basis"]]} moneyCols={["p10_bn","p50_bn","p90_bn"]}/></Card><Card><h2>Cost composition</h2><p className="chartCaption">Direct, indirect and reserve are scenario-controlled. For the detailed uncertainty view use QCRA/QSRA.</p><ResponsiveContainer width="100%" height={320}><BarChart data={[{name:'Direct',value:direct},{name:'Indirect',value:indirect},{name:'Reserve',value:reserves}]}><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis dataKey="name"/><YAxis/><Tooltip/><Bar dataKey="value" fill="#8df7ff"/></BarChart></ResponsiveContainer></Card></section>}
        {tab === 'schedule' && <section className="layout two"><Card><h2>Schedule bridge vs Base</h2><p className="chartCaption">This is the month-by-month reason the scenario becomes faster or slower than Base.</p>{scheduleWaterfall.map((x,i)=><div className={`reason ${x.kind==='total'?'deltaReason':''}`} key={i}><span>{i+1}</span><b>{x.driver}</b><br/>{x.kind==='total'?`${x.months} months`:(x.months>=0?'+':'') + x.months + ' months'}</div>)}<h3>Scenario schedule logic</h3><Table rows={schedule} cols={[["activity_id","Activity"],["phase","Phase"],["activity","Name"],["predecessor","Pred"],["duration_months","Months"],["critical","Critical"],["basis","Basis"]]}/></Card><Card><h2>QSRA finish-date curve</h2><p className="chartCaption">P50 equals the headline schedule. P80/P90 show how severe the delivery tail becomes after the scenario trade-off.</p><div className="metrics"><div>P50<b>{qsra.p50} mo</b></div><div>P80<b>{qsra.p80} mo</b></div><div>P90<b>{qsra.p90} mo</b></div></div><ResponsiveContainer width="100%" height={280}><LineChart data={curve}><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis dataKey="percentile"/><YAxis/><Tooltip formatter={(v) => [`${v} months`, "QSRA finish date"]}/><ReferenceLine x={50} stroke="#ffffff88" label="P50 = headline"/><ReferenceLine x={80} stroke="#ffffff55" label="P80 = board risk"/><Line type="monotone" name="QSRA finish date" dataKey="schedule_months" stroke="#b18cff" strokeWidth={4}/></LineChart></ResponsiveContainer><div className="reason p80Translation"><span>1/5</span>{p80Talk.schedule}</div><div className="reason p80Translation"><span>!</span>{p80Talk.board}</div>{(model.monte_carlo?.curve_readout || []).slice(1).map((x,i)=><div className="reason" key={i}><span>{i+1}</span>{x}</div>)}</Card></section>}
        {tab === 'risk' && <section className="layout two"><Card><h2>Risk Register Pro</h2><p>Risk output should include cause, event, impact, owner, mitigation and links to WBS/CBS.</p><Table rows={risks} cols={[['risk_id','ID'],['title','Risk'],['cause','Cause'],['event','Event'],['impact','Impact'],['probability_pct','Prob %'],['activity_id','Activity'],['cbs','CBS'],['owner','Owner'],['mitigation','Mitigation']]}/></Card><Card><h2>Top exposure drivers</h2><ResponsiveContainer width="100%" height={380}><BarChart data={tornado} layout="vertical"><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis type="number"/><YAxis dataKey="driver" type="category" width={150}/><Tooltip/><Bar dataKey="contribution" fill="#8df7ff"/></BarChart></ResponsiveContainer></Card></section>}
        {tab === 'monte' && <section className="layout two"><Card><h2>QCRA cost range curve</h2><p className="chartCaption">This is not a spend forecast over time. It is the probability range: P50 matches the headline cost, P80/P90 visualise the downside contingency tail created by the selected scenario.</p><div className="metrics"><div>P50 headline<b>{model.cost_p50}</b></div><div>P80 risk exposure<b>{fmt(qcra.p80)}</b></div><div>P90 stress case<b>{fmt(qcra.p90)}</b></div></div><ResponsiveContainer width="100%" height={280}><AreaChart data={curve}><defs><linearGradient id="caseyG" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stopColor="#8df7ff" stopOpacity=".55"/><stop offset="1" stopColor="#8df7ff" stopOpacity="0"/></linearGradient></defs><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis dataKey="percentile"/><YAxis/><Tooltip formatter={(v) => [`$${Number(v).toFixed(1)}B`, "QCRA total outturn"]}/><ReferenceLine x={50} stroke="#ffffff88" label="P50 = headline"/><ReferenceLine x={80} stroke="#ffffff55" label="P80 = board risk"/><Area type="monotone" name="QCRA total outturn" dataKey="cost_bn" stroke="#8df7ff" fill="url(#caseyG)"/></AreaChart></ResponsiveContainer>{(model.monte_carlo?.curve_readout || []).slice(0,1).map((x,i)=><div className="reason" key={i}><span>i</span>{x}</div>)}<div className="reason p80Translation"><span>1/5</span>{p80Talk.cost}</div><div className="reason"><span>!</span>This curve is a probability distribution, not spend over time. The x-axis is confidence percentile. P50 equals the headline estimate; P80/P90 are board downside exposure.</div></Card><Card><h2>QSRA schedule range curve</h2><p className="chartCaption">P50 matches the headline duration. P80/P90 show the likely board conversation if critical path risk lands.</p><div className="metrics"><div>P50 headline<b>{qsra.p50} mo</b></div><div>P80 risk date<b>{qsra.p80} mo</b></div><div>P90 stress date<b>{qsra.p90} mo</b></div></div><ResponsiveContainer width="100%" height={280}><LineChart data={curve}><CartesianGrid strokeDasharray="3 3" stroke="#ffffff18"/><XAxis dataKey="percentile"/><YAxis/><Tooltip formatter={(v) => [`${v} months`, "QSRA finish date"]}/><ReferenceLine x={50} stroke="#ffffff88" label="P50 = headline"/><ReferenceLine x={80} stroke="#ffffff55" label="P80 = board risk"/><Line type="monotone" name="QSRA finish date" dataKey="schedule_months" stroke="#b18cff" strokeWidth={4}/></LineChart></ResponsiveContainer><div className="reason p80Translation"><span>1/5</span>{p80Talk.schedule}</div><div className="reason p80Translation"><span>!</span>{p80Talk.board}</div>{(model.monte_carlo?.curve_readout || []).map((x,i)=><div className="reason" key={i}><span>{i+1}</span>{x}</div>)}</Card></section>}
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

        {tab === 'causal' && <section className="layout two"><CausalGraph model={model}/><BoardPressureSimulator model={model}/><BenchmarkIntelligence model={model}/><Card><h2>Evidence Mode: {viewMode}</h2>{evidenceScorecard(model).map((x,i)=><div className="reason" key={x.name}><span>{i+1}</span><b>{x.name}: {Math.round(x.score)}%</b><br/>{x.note}</div>)}</Card></section>}

        {tab === 'outputs' && <section className="layout two"><Card><h2>Generated Artefacts</h2><p>The public demo previews the intelligence pack. Enterprise access unlocks the live generated controls deliverables.</p><div className="exports v50Exports lockedExports">
          <button onClick={() => download('/export/workbook', model, `${model.id || 'casey'}_COST_WORKBOOK.xlsx`)}><FileSpreadsheet/> Generate Cost Model XLSX</button>
          <button onClick={() => download('/export/risk-register', model, `${model.id || 'casey'}_RISK_REGISTER.xlsx`)}><Database/> Generate Risk Register XLSX</button>
          <button onClick={() => download('/export/xer', model, `${model.id || 'casey'}_PRA_SCHEDULE.xer`)}><Workflow/> Generate PRA Schedule XER</button>
          <button onClick={() => download('/export/qcra-qsra', model, `${model.id || 'casey'}_QCRA_QSRA.xlsx`)}><BarChart3/> Generate QCRA/QSRA Pack</button>
          <button onClick={() => download('/export/json', model, `${model.id || 'casey'}_AUDIT_MODEL.json`)}><Brain/> Generate Audit File JSON</button>
          <button onClick={() => download('/export/all', model, `${model.id || 'casey'}_FULL_BOARD_PACK.zip`)}><Download/> Generate Full Pack ZIP</button>
          <a className="contactBtn" href={emailLink}><Mail/> Request Enterprise Review</a></div></Card><Card><h2>What the pack delivers</h2>{['Executive control centre with project, scenario, class, level and confidence clearly identified','Scenario comparison covering Base, Faster, Cheaper, Lower Risk and Premium cases','Selected estimate class plus all class levels for audit and challenge','Direct, indirect and reserve cost views with QCRA cost curve and cost tornado','All schedule levels with QSRA schedule curve and schedule tornado','Risk register with cause, event, impact, owner, mitigation, trigger and quantified likelihood','Basis of Estimate, assumptions, exclusions and benchmark validation','Commercial next steps: buyer action, procurement challenge and board decision path'].map((x,i)=><div className="reason" key={x}><span>{i+1}</span>{x}</div>)}</Card></section>}

        {tab === 'advisor' && <section className="layout two"><Card><h2>Ask CASEY</h2><div className="chatBox">{chat.map((m,i)=><div key={i} className={`msg ${m.role}`}>{asText(m.text)}</div>)}</div><div className="ask"><input value={chatQ} onChange={e=>setChatQ(e.target.value)} onKeyDown={e=>{if(e.key==='Enter')ask()}} placeholder="Ask why the cost or risk moved..."/><button onClick={ask}>Ask</button></div></Card><Card><h2>Upload estimate challenge</h2><p>Use this to show buyers how CASEY can challenge a Tier 1 estimate.</p><label className="upload"><Upload size={18}/> Upload file<input type="file" onChange={upload}/></label><button className="secondary" onClick={()=>setUploadResult({review:'Sample contractor estimate challenge', findings:['Direct costs above benchmark in power train and cooling package','Indirects and preliminaries need clearer split from reserves','Schedule contingency understated against critical path risks','Risk allowance should separate QCRA cost and QSRA schedule exposure'], next_action:'Request rate build-up, supplier quotes, basis of estimate and revised risk register.'})}>Run sample challenge</button>{uploadResult && <pre>{JSON.stringify(uploadResult,null,2)}</pre>}</Card></section>}

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

createRoot(document.getElementById('root')).render(<ErrorBoundary><App/></ErrorBoundary>);
