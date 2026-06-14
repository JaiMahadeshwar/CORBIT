/**
 * CASEY Universal Board Pack Generator
 * =====================================
 * Generates a 13-slide board-grade PowerPoint from ANY CASEY model:
 *   - Earth Demo (HS2, Crossrail, etc.)
 *   - Space Demo (Lunar Base, Mars, Starship, etc.)
 *   - All 50+ Showcase Library programmes
 *   - Free project runs (any sector, any country, any currency, any scale)
 *   - Defence, Nuclear, Life Sciences, Data Centres, Rail, Energy, Mining,
 *     Semiconductor, Battery/Gigafactory, Oil & Gas, Water, Space missions
 *
 * Usage: node generate_board_pack.js <model.json> [output.pptx]
 * Or require and call generateBoardPack(model, outputPath)
 */

const pptxgen = require('pptxgenjs');
const fs = require('fs');

// ── COLOUR PALETTE ──────────────────────────────────────────────────────────
const C = {
  bg:    '080C14', bg2:  '0D1526', bg3:  '111827',
  blue:  '4A9EFF', cyan: '8DF7FF', green:'10B981',
  amber: 'F59E0B', red:  'EF4444', purple:'9B7FE8',
  white: 'FFFFFF', light:'E2E8F0', mid:  '94A3B8',
  dim:   '475569', dim2: '334155', dim3: '1E293B',
};

// ── SECTOR → ACCENT COLOUR MAP ──────────────────────────────────────────────
function sectorAccent(model) {
  const s = String(model.subsector || model.mode || model.sector || '').toLowerCase();
  const mode = String(model.mode || '').toLowerCase();
  if (mode === 'space' || s.includes('space') || s.includes('lunar') || s.includes('mars') || s.includes('orbit')) return '7C3AED';
  if (s.includes('rail') || s.includes('transit') || s.includes('metro')) return '1D4ED8';
  if (s.includes('life') || s.includes('pharma') || s.includes('bio') || s.includes('glp') || s.includes('vaccine')) return '0D9488';
  if (s.includes('data') || s.includes('hyperscale') || s.includes('cloud') || s.includes('ai campus')) return '0891B2';
  if (s.includes('nuclear') || s.includes('smr')) return 'DC2626';
  if (s.includes('defence') || s.includes('security') || s.includes('aukus') || s.includes('drone') || s.includes('missile')) return '1E3A5F';
  if (s.includes('energy') || s.includes('wind') || s.includes('solar') || s.includes('hydrogen') || s.includes('lng')) return 'D97706';
  if (s.includes('semiconductor') || s.includes('fab') || s.includes('chip')) return '6D28D9';
  if (s.includes('battery') || s.includes('gigafactory') || s.includes('ev')) return '065F46';
  if (s.includes('airport') || s.includes('aviation') || s.includes('changi')) return '1E40AF';
  if (s.includes('mining') || s.includes('mineral')) return '92400E';
  if (s.includes('oil') || s.includes('gas') || s.includes('lng') || s.includes('refin')) return '78350F';
  if (s.includes('water') || s.includes('utility') || s.includes('sewage')) return '164E63';
  if (s.includes('hospital') || s.includes('health') || s.includes('nhs')) return '7F1D1D';
  if (s.includes('education') || s.includes('university') || s.includes('campus')) return '312E81';
  if (s.includes('commercial') || s.includes('office') || s.includes('reit')) return '374151';
  return '1D4ED8'; // default blue
}

// ── HELPERS ─────────────────────────────────────────────────────────────────
function ragColor(conf) {
  const c = Number(conf || 0);
  return c >= 75 ? C.green : c >= 55 ? C.amber : C.red;
}
function ragLabel(conf) {
  const c = Number(conf || 0);
  return c >= 75 ? 'APPROVAL READY' : c >= 55 ? 'CONDITIONAL' : 'DO NOT APPROVE';
}
function parseBn(v) {
  if (typeof v === 'number' && isFinite(v)) return v;
  const s = String(v || '0').replace(/[^0-9.kKmMbBtT]/g, '').toUpperCase();
  const n = parseFloat(s);
  if (!n) return 0;
  if (s.includes('T')) return n * 1000;
  if (s.includes('B')) return n;
  if (s.includes('M')) return n / 1000;
  if (s.includes('K')) return n / 1000000;
  return n;
}
function fmtMoney(v, curr) {
  const c = curr || '£';
  const bn = parseBn(v);
  if (!bn) return c + '—';
  if (bn >= 1000) return c + (bn / 1000).toFixed(1) + 'T';
  if (bn >= 10)   return c + bn.toFixed(1) + 'B';
  if (bn >= 1)    return c + bn.toFixed(2) + 'B';
  if (bn >= 0.01) return c + Math.round(bn * 1000) + 'M';
  return c + Math.round(bn * 1000000) + 'k';
}
function safe(v, max) {
  const s = (v === null || v === undefined || v === 'null' || v === 'undefined') ? '—' : String(v);
  return max ? s.slice(0, max) : s;
}
function safeArr(v) { return Array.isArray(v) ? v : []; }
function mkShadow() { return { type: 'outer', color: '000000', blur: 8, offset: 3, angle: 45, opacity: 0.16 }; }

// Detect currency from model
function getCurr(model) {
  return model.currency_symbol ||
    (String(model.mode || '').toLowerCase() === 'space' ? '$' :
     String(model.location || '').toLowerCase().includes('australia') ? 'A$' :
     String(model.location || '').toLowerCase().includes('nigeria') ? '₦' :
     String(model.location || '').toLowerCase().includes('europe') || String(model.location || '').toLowerCase().includes('germany') || String(model.location || '').toLowerCase().includes('denmark') ? '€' :
     String(model.location || '').toLowerCase().includes('united states') || String(model.location || '').toLowerCase().includes('usa') ? '$' :
     '£');
}

// Build sector tag line for headers
function sectorTag(model) {
  const parts = [model.mode || 'Earth', model.subsector || model.sector, model.location].filter(Boolean);
  return parts.join('  ·  ');
}

// Get P-values from model (handles any model shape)
function getPValues(model) {
  const curr = getCurr(model);
  const mc = model.monte_carlo || {};
  const qcra = mc.qcra || {};
  const qsra = mc.qsra || {};
  const costBn = parseBn(model.cost_p50_bn || model.p50_cost_bn || model.cost_p50 || 1);
  const p50m = Number(model.schedule_months || model.duration_months || (String(model.schedule || '').match(/\d+/) || [24])[0]);

  return {
    curr,
    costBn,
    p10c: qcra.p10 || costBn * 0.82,
    p50c: qcra.p50 || costBn,
    p80c: qcra.p80 || parseBn(model.cost_p80) || costBn * 1.18,
    p90c: qcra.p90 || parseBn(model.cost_p90) || costBn * 1.28,
    p50m,
    p80m: qsra.p80 || Math.round(p50m * 1.18),
    p90m: qsra.p90 || Math.round(p50m * 1.28),
  };
}

// Get milestones — from model data or generate sector-appropriate ones
function getMilestones(model) {
  const ms = safeArr(model.schedule_detail || model.schedule_rows);
  const p = getPValues(model);
  const mode = String(model.mode || '').toLowerCase();
  const sub = String(model.subsector || model.sector || '').toLowerCase();

  if (ms.length >= 3) {
    const n = Math.min(ms.length, 8);
    return ms.slice(0, n).map((m, i) => ({
      label: safe(m.activity || m.name || m.description || `Phase ${i + 1}`, 22),
      mo: m.month || Math.round((i / (n - 1)) * p.p50m),
    }));
  }

  // Space
  if (mode === 'space' || sub.includes('space') || sub.includes('lunar') || sub.includes('mars')) {
    return [
      { label: 'Programme start',    mo: 0 },
      { label: 'Mission arch. (PDR)', mo: Math.round(p.p50m * 0.09) },
      { label: 'CDR complete',       mo: Math.round(p.p50m * 0.24) },
      { label: 'MAIT / integration', mo: Math.round(p.p50m * 0.52) },
      { label: 'Flight readiness',   mo: Math.round(p.p50m * 0.75) },
      { label: 'Launch / deploy',    mo: Math.round(p.p50m * 0.90) },
      { label: 'Mission ops',        mo: p.p50m },
    ];
  }
  if (sub.includes('nuclear') || sub.includes('smr')) {
    return [
      { label: 'FID / GDA consent',  mo: 0 },
      { label: 'Site prep',          mo: Math.round(p.p50m * 0.07) },
      { label: 'Nuclear island',     mo: Math.round(p.p50m * 0.22) },
      { label: 'Plant installation', mo: Math.round(p.p50m * 0.45) },
      { label: 'Cold commissioning', mo: Math.round(p.p50m * 0.65) },
      { label: 'Hot commissioning',  mo: Math.round(p.p50m * 0.82) },
      { label: 'Commercial ops',     mo: p.p50m },
    ];
  }
  if (sub.includes('life') || sub.includes('pharma') || sub.includes('bio') || sub.includes('glp')) {
    return [
      { label: 'EMA/FDA filing',     mo: 0 },
      { label: 'Site mobilise',      mo: Math.round(p.p50m * 0.10) },
      { label: 'Construction',       mo: Math.round(p.p50m * 0.30) },
      { label: 'Fit-out & MEP',      mo: Math.round(p.p50m * 0.50) },
      { label: 'IQ / OQ',           mo: Math.round(p.p50m * 0.65) },
      { label: 'PQ validation',      mo: Math.round(p.p50m * 0.80) },
      { label: 'Commercial batch',   mo: p.p50m },
    ];
  }
  if (sub.includes('data') || sub.includes('hyperscale') || sub.includes('cloud')) {
    return [
      { label: 'Permit & FID',       mo: 0 },
      { label: 'Shell & core',       mo: Math.round(p.p50m * 0.12) },
      { label: 'MEP install',        mo: Math.round(p.p50m * 0.32) },
      { label: 'IT fit-out',         mo: Math.round(p.p50m * 0.52) },
      { label: 'Cooling commissioning', mo: Math.round(p.p50m * 0.72) },
      { label: 'IT commissioning',   mo: Math.round(p.p50m * 0.88) },
      { label: 'Hyperscale live',    mo: p.p50m },
    ];
  }
  if (sub.includes('defence') || sub.includes('security') || sub.includes('naval')) {
    return [
      { label: 'Contract award',     mo: 0 },
      { label: 'Design phase',       mo: Math.round(p.p50m * 0.10) },
      { label: 'Long-lead proc.',    mo: Math.round(p.p50m * 0.25) },
      { label: 'Build programme',    mo: Math.round(p.p50m * 0.50) },
      { label: 'Systems integration',mo: Math.round(p.p50m * 0.72) },
      { label: 'Acceptance trials',  mo: Math.round(p.p50m * 0.88) },
      { label: 'Delivery',           mo: p.p50m },
    ];
  }
  if (sub.includes('semiconductor') || sub.includes('fab') || sub.includes('chip')) {
    return [
      { label: 'Site & utilities',   mo: 0 },
      { label: 'Shell construction', mo: Math.round(p.p50m * 0.15) },
      { label: 'Cleanroom fit-out',  mo: Math.round(p.p50m * 0.35) },
      { label: 'Tool install',       mo: Math.round(p.p50m * 0.58) },
      { label: 'UPW & commissioning',mo: Math.round(p.p50m * 0.76) },
      { label: 'Yield ramp',         mo: Math.round(p.p50m * 0.90) },
      { label: 'Volume production',  mo: p.p50m },
    ];
  }
  if (sub.includes('energy') || sub.includes('wind') || sub.includes('lng') || sub.includes('hydrogen')) {
    return [
      { label: 'FID / consent',      mo: 0 },
      { label: 'Procurement',        mo: Math.round(p.p50m * 0.15) },
      { label: 'Civil works',        mo: Math.round(p.p50m * 0.35) },
      { label: 'Plant install',      mo: Math.round(p.p50m * 0.60) },
      { label: 'Cold commissioning', mo: Math.round(p.p50m * 0.78) },
      { label: 'Grid sync',          mo: Math.round(p.p50m * 0.90) },
      { label: 'Commercial ops',     mo: p.p50m },
    ];
  }
  // Generic / rail / infrastructure default
  return [
    { label: 'Authorisation',        mo: 0 },
    { label: 'Design & consent',     mo: Math.round(p.p50m * 0.10) },
    { label: 'Procurement',          mo: Math.round(p.p50m * 0.28) },
    { label: 'Construction 50%',     mo: Math.round(p.p50m * 0.55) },
    { label: 'Systems & testing',    mo: Math.round(p.p50m * 0.76) },
    { label: 'Commissioning',        mo: Math.round(p.p50m * 0.90) },
    { label: 'Handover',             mo: p.p50m },
  ];
}

// ── SLIDE BUILDERS ───────────────────────────────────────────────────────────

function slide01_Cover(pres, m, accent) {
  const sl = pres.addSlide();
  sl.background = { color: C.bg };
  const conf = Number(m.confidence_pct || 0);
  const rc = ragColor(conf);

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: accent }, line: { color: accent } });
  sl.addText('CASEY', { x: 0.5, y: 0.2, w: 2, h: 0.32, fontSize: 10, bold: true, color: C.cyan, charSpacing: 8, fontFace: 'Arial', margin: 0 });
  sl.addText('PROGRAMME INTELLIGENCE', { x: 0.5, y: 0.5, w: 5, h: 0.2, fontSize: 7, color: C.dim, charSpacing: 4, fontFace: 'Arial', margin: 0 });
  sl.addText('CONFIDENTIAL', { x: 7.5, y: 0.26, w: 2, h: 0.2, fontSize: 7, color: C.dim, charSpacing: 3, align: 'right', fontFace: 'Arial', margin: 0 });

  const title = safe(m.title || m.programme_title || m.subsector || 'Programme Intelligence', 55);
  sl.addText(title, { x: 0.5, y: 1.05, w: 9, h: 1.25, fontSize: 31, bold: true, color: C.white, fontFace: 'Cambria', margin: 0 });
  sl.addText('BOARD INTELLIGENCE PACK', { x: 0.5, y: 2.26, w: 9, h: 0.32, fontSize: 10, color: C.cyan, charSpacing: 5, fontFace: 'Arial', margin: 0 });
  sl.addText(safe(sectorTag(m), 80), { x: 0.5, y: 2.62, w: 9, h: 0.26, fontSize: 10, color: C.mid, fontFace: 'Arial', margin: 0 });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.98, w: 9, h: 0.005, fill: { color: C.dim3 }, line: { color: C.dim3 } });

  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 3.16, w: 2.1, h: 0.42, fill: { color: rc }, line: { color: rc }, rectRadius: 0.05 });
  sl.addText(ragLabel(conf), { x: 0.5, y: 3.16, w: 2.1, h: 0.42, fontSize: 8, bold: true, color: C.bg, align: 'center', fontFace: 'Arial', margin: 0 });

  const pv = getPValues(m);
  [
    { l: 'P50 COST',   v: m.cost_p50 || fmtMoney(pv.p50c, pv.curr) },
    { l: 'SCHEDULE',   v: m.schedule || (pv.p50m + ' months') },
    { l: 'CONFIDENCE', v: conf + '%' },
  ].forEach((mt, i) => {
    const x = 3.0 + i * 2.2;
    sl.addText(mt.l, { x, y: 3.12, w: 2, h: 0.22, fontSize: 7, color: C.dim, charSpacing: 3, fontFace: 'Arial', margin: 0 });
    sl.addText(mt.v, { x, y: 3.3, w: 2, h: 0.38, fontSize: 17, bold: true, color: C.white, fontFace: 'Cambria', margin: 0 });
  });

  const auth = safe(m.institutional_authority_line || m.executive_summary || '', 250);
  if (auth && auth !== '—') {
    sl.addText(auth, { x: 0.5, y: 3.9, w: 9, h: 0.9, fontSize: 10, color: C.light, fontFace: 'Calibri', lineSpacingMultiple: 1.35, margin: 0 });
  }

  const today = new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });
  sl.addText(today + '  ·  ' + safe(m.estimate_class_name || 'Class ' + (m.estimate_class || 3)) + '  ·  ' + safe(m.schedule_level_name || 'Level ' + (m.schedule_level || 4)), {
    x: 0.5, y: 5.28, w: 9, h: 0.2, fontSize: 7.5, color: C.dim, fontFace: 'Arial', margin: 0
  });
  sl.addNotes('COVER — Open with programme name and location. Pause on the RAG pill. Walk through P50, schedule, confidence. Read the authority line. The top accent bar colour signals the sector.');
}

function slide02_ExecutiveSummary(pres, m, accent) {
  const sl = pres.addSlide();
  sl.background = { color: C.bg };
  const conf = Number(m.confidence_pct || 0);
  const rc = ragColor(conf);
  const pv = getPValues(m);

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: rc }, line: { color: rc } });
  sl.addText('01  EXECUTIVE SUMMARY', { x: 0.5, y: 0.18, w: 8, h: 0.22, fontSize: 7.5, color: C.dim, charSpacing: 4, fontFace: 'Arial', margin: 0 });
  sl.addText('CASEY · PROGRAMME INTELLIGENCE', { x: 0.5, y: 0.38, w: 9, h: 0.3, fontSize: 14, bold: true, color: C.white, fontFace: 'Cambria', margin: 0 });
  sl.addText(safe(m.title || m.subsector || 'Programme', 60), { x: 0.5, y: 0.67, w: 8, h: 0.2, fontSize: 9.5, color: C.mid, fontFace: 'Arial', margin: 0 });

  // Verdict banner
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 0.98, w: 9, h: 0.72, fill: { color: C.bg3 }, line: { color: rc }, rectRadius: 0.07, shadow: mkShadow() });
  sl.addText('VERDICT', { x: 0.7, y: 1.06, w: 1.4, h: 0.18, fontSize: 7, color: C.dim, charSpacing: 4, fontFace: 'Arial', margin: 0 });
  sl.addText(ragLabel(conf), { x: 0.7, y: 1.22, w: 2.2, h: 0.28, fontSize: 12, bold: true, color: rc, fontFace: 'Arial', margin: 0 });
  const auth = safe(m.institutional_authority_line || m.executive_summary || 'Programme intelligence generated.', 220);
  sl.addText(auth, { x: 3.1, y: 1.04, w: 6.2, h: 0.58, fontSize: 9.5, color: C.light, fontFace: 'Calibri', lineSpacingMultiple: 1.3, margin: 0 });

  // 6 KPI cards
  const kpis = [
    { l: 'P50 COST',       v: m.cost_p50 || fmtMoney(pv.p50c, pv.curr), s: safe(m.estimate_class_name || 'Class ' + (m.estimate_class || 3), 22) },
    { l: 'P80 COST',       v: m.cost_p80 || fmtMoney(pv.p80c, pv.curr), s: 'Board contingency ceiling' },
    { l: 'P90 COST',       v: fmtMoney(pv.p90c, pv.curr),               s: 'Stress / worst case' },
    { l: 'SCHEDULE (P50)', v: m.schedule || pv.p50m + ' months',          s: 'Base delivery' },
    { l: 'SCHEDULE (P80)', v: pv.p80m + ' months',                        s: 'QSRA P80' },
    { l: 'CONFIDENCE',     v: conf + '%',                                  s: conf >= 75 ? 'Board-defensible ✓' : 'Target: 75% board threshold' },
  ];
  kpis.forEach((k, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.5 + col * 3.05, y = 1.92 + row * 1.12;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.9, h: 1.0, fill: { color: C.bg3 }, line: { color: C.dim3 }, rectRadius: 0.06, shadow: mkShadow() });
    sl.addText(k.l, { x: x + 0.14, y: y + 0.1, w: 2.6, h: 0.18, fontSize: 7, color: C.dim, charSpacing: 3, fontFace: 'Arial', margin: 0 });
    sl.addText(k.v, { x: x + 0.14, y: y + 0.27, w: 2.6, h: 0.38, fontSize: 19, bold: true, color: i === 5 ? rc : C.white, fontFace: 'Cambria', margin: 0 });
    sl.addText(k.s, { x: x + 0.14, y: y + 0.68, w: 2.6, h: 0.2, fontSize: 8, color: C.dim, fontFace: 'Arial', margin: 0 });
  });

  // OBA note
  if (m.oba_adjustment) {
    sl.addText('OBA: ' + safe(m.oba_adjustment, 110), { x: 0.5, y: 4.2, w: 9, h: 0.22, fontSize: 8, color: C.mid, fontFace: 'Calibri', margin: 0 });
  }

  sl.addText('CASEY · Programme Intelligence · Confidential', { x: 0.5, y: 5.38, w: 9, h: 0.16, fontSize: 7, color: C.dim3, fontFace: 'Arial', margin: 0 });
  sl.addNotes('EXECUTIVE SUMMARY — Lead with verdict, then the six KPI cards. P80 is the board contingency conversation. P90 is stress. Confidence is the key output — target 75%.');
}

function slide03_CostModel(pres, m, accent) {
  const sl = pres.addSlide();
  sl.background = { color: C.bg };
  const pv = getPValues(m);
  const curr = pv.curr;

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: accent }, line: { color: accent } });
  sl.addText('02  COST MODEL & RISK RANGES', { x: 0.5, y: 0.18, w: 8, h: 0.22, fontSize: 7.5, color: C.dim, charSpacing: 4, fontFace: 'Arial', margin: 0 });
  sl.addText('P10 / P50 / P80 / P90  ·  Cost Breakdown  ·  ' + (m.estimate_class_name || ('Class ' + (m.estimate_class || 3))), {
    x: 0.5, y: 0.38, w: 9, h: 0.28, fontSize: 13, bold: true, color: C.white, fontFace: 'Cambria', margin: 0
  });

  // P-curve bar chart
  sl.addChart(pres.charts.BAR, [{
    name: 'Cost',
    labels: ['P10\n(Optimistic)', 'P50\n(Base)', 'P80\n(Board)', 'P90\n(Stress)'],
    values: [pv.p10c, pv.p50c, pv.p80c, pv.p90c],
  }], {
    x: 0.5, y: 0.82, w: 5.5, h: 3.5, barDir: 'col',
    chartColors: ['1E6AD4', '4A9EFF', 'F59E0B', 'EF4444'],
    chartArea: { fill: { color: C.bg3 }, roundedCorners: true },
    catAxisLabelColor: C.mid, valAxisLabelColor: C.mid,
    valGridLine: { color: C.dim3, size: 0.5 }, catGridLine: { style: 'none' },
    showValue: true, dataLabelColor: C.white, dataLabelFontSize: 9,
    showLegend: false, showTitle: false,
  });

  // Cost breakdown side panel
  const rows = safeArr(m.cost_breakdown || m.cost_lines).slice(0, 6);
  sl.addText('COST BREAKDOWN', { x: 6.2, y: 0.86, w: 3.4, h: 0.2, fontSize: 7, color: C.dim, charSpacing: 4, fontFace: 'Arial', margin: 0 });
  if (rows.length > 0) {
    rows.forEach((r, i) => {
      const y = 1.1 + i * 0.52;
      sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.2, y, w: 3.4, h: 0.44, fill: { color: C.bg3 }, line: { color: C.dim3 }, rectRadius: 0.05 });
      sl.addText(safe(r.description || r.cbs || 'Line ' + (i+1), 28), { x: 6.35, y: y + 0.05, w: 2.1, h: 0.18, fontSize: 8, bold: true, color: C.light, fontFace: 'Calibri', margin: 0 });
      sl.addText(fmtMoney(r.p50_bn, curr), { x: 8.26, y: y + 0.04, w: 1.2, h: 0.22, fontSize: 11, bold: true, color: C.blue, align: 'right', fontFace: 'Cambria', margin: 0 });
      sl.addText(safe(r.type || '', 16), { x: 6.35, y: y + 0.24, w: 2.6, h: 0.16, fontSize: 7, color: C.dim, fontFace: 'Arial', margin: 0 });
    });
  } else {
    sl.addText('Upload cost workbook (XLSX) to activate\ndetailed breakdown by CBS/WBS package.', {
      x: 6.2, y: 1.2, w: 3.4, h: 0.6, fontSize: 9, color: C.dim, fontFace: 'Calibri', lineSpacingMultiple: 1.5, margin: 0
    });
  }

  // OBA / accuracy note
  sl.addText(
    'Estimate accuracy: ' + safe(m.estimate_accuracy || '±20%') + '  ·  OBA: ' + safe(m.oba_adjustment || 'Reference-class adjustment applied', 60) + '  ·  Contingency: ' + safe(m.contingency_pct ? m.contingency_pct + '%' : '—'),
    { x: 0.5, y: 4.48, w: 9, h: 0.22, fontSize: 7.5, color: C.mid, fontFace: 'Arial', margin: 0 }
  );
  sl.addText('CASEY · Programme Intelligence · Confidential', { x: 0.5, y: 5.38, w: 9, h: 0.16, fontSize: 7, color: C.dim3, fontFace: 'Arial', margin: 0 });
  sl.addNotes('COST — Walk P10→P50→P80→P90. The gap between P50 and P80 is the reserve conversation. OBA note shows the reference-class uplift. Cost breakdown shows where the money sits.');
}

function slide04_Schedule(pres, m, accent) {
  const sl = pres.addSlide();
  sl.background = { color: C.bg };
  const pv = getPValues(m);

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: accent }, line: { color: accent } });
  sl.addText('03  SCHEDULE & DELIVERY', { x: 0.5, y: 0.18, w: 8, h: 0.22, fontSize: 7.5, color: C.dim, charSpacing: 4, fontFace: 'Arial', margin: 0 });
  sl.addText('QSRA Fan  ·  P50 / P80 / P90  ·  ' + (m.schedule_level_name || 'Level ' + (m.schedule_level || 4)), {
    x: 0.5, y: 0.38, w: 9, h: 0.28, fontSize: 13, bold: true, color: C.white, fontFace: 'Cambria', margin: 0
  });

  // QSRA fan chart
  const milabels = ['Concept', 'Design', 'Procure', 'Build', 'Comm.', 'Delivery'];
  sl.addChart(pres.charts.LINE, [
    { name: 'P10', labels: milabels, values: [0, pv.p50m*.14, pv.p50m*.32, pv.p50m*.62, pv.p50m*.85, Math.round(pv.p50m*.91)] },
    { name: 'P50', labels: milabels, values: [0, pv.p50m*.16, pv.p50m*.35, pv.p50m*.65, pv.p50m*.88, pv.p50m] },
    { name: 'P80', labels: milabels, values: [0, pv.p50m*.18, pv.p50m*.38, pv.p50m*.70, pv.p50m*.92, pv.p80m] },
    { name: 'P90', labels: milabels, values: [0, pv.p50m*.20, pv.p50m*.41, pv.p50m*.74, pv.p50m*.95, pv.p90m] },
  ], {
    x: 0.5, y: 0.82, w: 5.5, h: 3.5,
    chartColors: ['1E6AD4', '4A9EFF', 'F59E0B', 'EF4444'],
    chartArea: { fill: { color: C.bg3 }, roundedCorners: true },
    catAxisLabelColor: C.mid, valAxisLabelColor: C.mid,
    valGridLine: { color: C.dim3, size: 0.5 }, catGridLine: { style: 'none' },
    lineSmooth: true, lineSize: 2.2,
    showLegend: true, legendPos: 'b', legendFontSize: 8.5, legendColor: C.mid,
    showTitle: false,
  });

  // Schedule metrics
  const sMetrics = [
    { l: 'P50 DELIVERY', v: m.schedule || (pv.p50m + ' months'), c: C.blue },
    { l: 'P80 DELIVERY', v: pv.p80m + ' months',                  c: C.amber },
    { l: 'P90 DELIVERY', v: pv.p90m + ' months',                  c: C.red },
    { l: 'SCHEDULE LEVEL', v: safe(m.schedule_level_name || ('Level ' + (m.schedule_level || 4)), 22), c: C.mid },
    { l: 'GOVERNING CONSTRAINT', v: safe(m.governing_constraint_prominent || m.primary_constraint || '—', 30), c: C.amber },
  ];
  sMetrics.forEach((s, i) => {
    const y = 0.86 + i * 0.68;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.2, y, w: 3.4, h: 0.56, fill: { color: C.bg3 }, line: { color: C.dim3 }, rectRadius: 0.05 });
    sl.addText(s.l, { x: 6.34, y: y + 0.07, w: 3.1, h: 0.17, fontSize: 7, color: C.dim, charSpacing: 2, fontFace: 'Arial', margin: 0 });
    sl.addText(s.v, { x: 6.34, y: y + 0.26, w: 3.1, h: 0.24, fontSize: i <= 2 ? 13 : 10, bold: i <= 2, color: s.c, fontFace: 'Cambria', margin: 0 });
  });

  sl.addText('CASEY · Programme Intelligence · Confidential', { x: 0.5, y: 5.38, w: 9, h: 0.16, fontSize: 7, color: C.dim3, fontFace: 'Arial', margin: 0 });
  sl.addNotes('SCHEDULE — The fan shows how uncertainty grows toward delivery. P50 is the base, P80 is the board planning assumption, P90 is the stress. The governing constraint box is the critical question.');
}

function slide05_RiskRegister(pres, m, accent) {
  const sl = pres.addSlide();
  sl.background = { color: C.bg };
  const risks = safeArr(m.risks || m.risk_register).slice(0, 8);
  const curr = getCurr(m);

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: C.amber }, line: { color: C.amber } });
  sl.addText('04  RISK REGISTER', { x: 0.5, y: 0.18, w: 8, h: 0.22, fontSize: 7.5, color: C.dim, charSpacing: 4, fontFace: 'Arial', margin: 0 });
  sl.addText('Top Risks by Priority  ·  EMV  ·  Owners', { x: 0.5, y: 0.38, w: 9, h: 0.28, fontSize: 13, bold: true, color: C.white, fontFace: 'Cambria', margin: 0 });

  if (risks.length > 0) {
    const headers = [
      { text: 'Risk', options: { bold: true, color: C.white, fill: { color: C.bg3 } } },
      { text: 'Probability', options: { bold: true, color: C.white, fill: { color: C.bg3 } } },
      { text: 'Impact', options: { bold: true, color: C.white, fill: { color: C.bg3 } } },
      { text: 'EMV', options: { bold: true, color: C.white, fill: { color: C.bg3 } } },
      { text: 'Owner', options: { bold: true, color: C.white, fill: { color: C.bg3 } } },
    ];
    const tableRows = [headers];
    risks.forEach(r => {
      const prob = String(r.probability || '').toLowerCase();
      const isHigh = prob.includes('high') || prob.includes('critical') || prob.includes('very high');
      const emv = r.cost_emv_bn ? fmtMoney(r.cost_emv_bn, curr) : '—';
      tableRows.push([
        { text: safe(r.title || r.risk || '—', 38), options: { color: isHigh ? 'FCA5A5' : C.light, fontSize: 8.5 } },
        { text: safe(r.probability || '—', 14), options: { color: isHigh ? C.red : C.amber, bold: isHigh, fontSize: 8.5 } },
        { text: safe(r.impact || r.consequence || '—', 28), options: { color: C.light, fontSize: 8.5 } },
        { text: emv, options: { color: C.blue, bold: true, fontSize: 8.5 } },
        { text: safe(r.owner || '—', 20), options: { color: C.mid, fontSize: 8.5 } },
      ]);
    });
    sl.addTable(tableRows, {
      x: 0.5, y: 0.82, w: 9.1,
      border: { pt: 0.4, color: C.dim3 },
      fill: { color: C.bg3 },
      fontFace: 'Calibri',
      colW: [3.2, 1.3, 2.1, 1.0, 1.5],
    });
  } else {
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.0, w: 9, h: 0.7, fill: { color: C.bg3 }, line: { color: C.dim3 }, rectRadius: 0.07 });
    sl.addText('Upload risk register (XLSX) for programme-specific risk intelligence.\nCurrent CASEY model includes ' + (safeArr(m.risks).length || 0) + ' risks from model generation.',
      { x: 0.7, y: 1.1, w: 8.6, h: 0.5, fontSize: 10, color: C.mid, fontFace: 'Calibri', lineSpacingMultiple: 1.5, margin: 0 });
  }

  // Stats footer
  const total = safeArr(m.risks || m.risk_register).length;
  const highCount = safeArr(m.risks || m.risk_register).filter(r => String(r.probability || '').toLowerCase().includes('high')).length;
  const constraint = safe(m.governing_constraint_prominent || m.primary_constraint || '—', 55);
  sl.addText('TOTAL RISKS: ' + total + '   HIGH / CRITICAL: ' + highCount + '   GOVERNING CONSTRAINT: ' + constraint, {
    x: 0.5, y: 4.98, w: 9, h: 0.22, fontSize: 8, color: C.amber, fontFace: 'Arial', margin: 0
  });
  sl.addText('CASEY · Programme Intelligence · Confidential', { x: 0.5, y: 5.38, w: 9, h: 0.16, fontSize: 7, color: C.dim3, fontFace: 'Arial', margin: 0 });
  sl.addNotes('RISK REGISTER — Every red-probability risk needs a named owner and a closure date before approval. EMV column is the expected monetary value. Governing constraint is the single biggest threat.');
}

function slide06_QCRA_QSRA(pres, m, accent) {
  const sl = pres.addSlide();
  sl.background = { color: C.bg };
  const pv = getPValues(m);
  const curr = pv.curr;

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: C.red }, line: { color: C.red } });
  sl.addText('05  RISK CURVES  (QCRA / QSRA)', { x: 0.5, y: 0.18, w: 8, h: 0.22, fontSize: 7.5, color: C.dim, charSpacing: 4, fontFace: 'Arial', margin: 0 });
  sl.addText('Quantitative Cost & Schedule Risk  ·  Cumulative Probability', { x: 0.5, y: 0.38, w: 9, h: 0.28, fontSize: 13, bold: true, color: C.white, fontFace: 'Cambria', margin: 0 });

  const pLabels = ['P10', 'P20', 'P30', 'P40', 'P50', 'P60', 'P70', 'P80', 'P90'];
  const qcraVals = [pv.p10c, pv.p10c+(pv.p80c-pv.p10c)*.11, pv.p10c+(pv.p80c-pv.p10c)*.22, pv.p10c+(pv.p80c-pv.p10c)*.38, pv.p50c, pv.p50c+(pv.p90c-pv.p50c)*.28, pv.p50c+(pv.p90c-pv.p50c)*.56, pv.p80c, pv.p90c];
  const qsraVals = [Math.round(pv.p50m*.84), Math.round(pv.p50m*.88), Math.round(pv.p50m*.91), Math.round(pv.p50m*.95), pv.p50m, Math.round(pv.p50m*1.07), Math.round(pv.p50m*1.13), pv.p80m, pv.p90m];

  sl.addChart(pres.charts.LINE, [{ name: 'QCRA', labels: pLabels, values: qcraVals }], {
    x: 0.5, y: 0.82, w: 4.5, h: 3.1,
    chartColors: ['4A9EFF'], chartArea: { fill: { color: C.bg3 }, roundedCorners: true },
    catAxisLabelColor: C.mid, valAxisLabelColor: C.mid,
    valGridLine: { color: C.dim3, size: 0.5 }, catGridLine: { style: 'none' },
    lineSmooth: true, lineSize: 2.5, showLegend: false,
    showTitle: true, title: 'Cost Risk (QCRA)', titleColor: C.mid, titleFontSize: 9,
  });

  sl.addChart(pres.charts.LINE, [{ name: 'QSRA', labels: pLabels, values: qsraVals }], {
    x: 5.15, y: 0.82, w: 4.5, h: 3.1,
    chartColors: ['F59E0B'], chartArea: { fill: { color: C.bg3 }, roundedCorners: true },
    catAxisLabelColor: C.mid, valAxisLabelColor: C.mid,
    valGridLine: { color: C.dim3, size: 0.5 }, catGridLine: { style: 'none' },
    lineSmooth: true, lineSize: 2.5, showLegend: false,
    showTitle: true, title: 'Schedule Risk (QSRA)', titleColor: C.mid, titleFontSize: 9,
  });

  const sumCards = [
    { l: 'P50 COST',  v: m.cost_p50 || fmtMoney(pv.p50c, curr), c: C.blue },
    { l: 'P80 COST',  v: m.cost_p80 || fmtMoney(pv.p80c, curr), c: C.amber },
    { l: 'P90 COST',  v: fmtMoney(pv.p90c, curr),               c: C.red },
    { l: 'P80 SCHED', v: pv.p80m + ' mo',                        c: C.amber },
    { l: 'P90 SCHED', v: pv.p90m + ' mo',                        c: C.red },
  ];
  sumCards.forEach((s, i) => {
    const x = 0.5 + i * 1.84;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 4.06, w: 1.72, h: 0.64, fill: { color: C.bg3 }, line: { color: C.dim3 }, rectRadius: 0.06 });
    sl.addText(s.l, { x: x + 0.1, y: 4.12, w: 1.5, h: 0.16, fontSize: 7, color: C.dim, charSpacing: 2, fontFace: 'Arial', margin: 0 });
    sl.addText(s.v, { x: x + 0.1, y: 4.28, w: 1.5, h: 0.3, fontSize: 13, bold: true, color: s.c, fontFace: 'Cambria', margin: 0 });
  });

  sl.addText('CASEY · Programme Intelligence · Confidential', { x: 0.5, y: 5.38, w: 9, h: 0.16, fontSize: 7, color: C.dim3, fontFace: 'Arial', margin: 0 });
  sl.addNotes('RISK CURVES — Left: QCRA (cost). Right: QSRA (schedule). Both show cumulative probability. The P50→P80 gap defines the reserve conversation. P90 is the stress case the board must understand.');
}

function slide07_Benchmark(pres, m, accent) {
  const sl = pres.addSlide();
  sl.background = { color: C.bg };
  const bms = safeArr(m.benchmark_comparison || m.benchmarks).slice(0, 6);

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: C.purple }, line: { color: C.purple } });
  sl.addText('06  BENCHMARK COMPARISON', { x: 0.5, y: 0.18, w: 8, h: 0.22, fontSize: 7.5, color: C.dim, charSpacing: 4, fontFace: 'Arial', margin: 0 });
  sl.addText('Reference Class  ·  ' + (bms.length || '137+') + ' Comparable Programmes  ·  How This Programme Compares', { x: 0.5, y: 0.38, w: 9.2, h: 0.28, fontSize: 13, bold: true, color: C.white, fontFace: 'Cambria', margin: 0 });

  if (bms.length >= 2) {
    const bmLabels = bms.map(b => safe((b.name || b.programme || '').slice(0, 16), '—'));
    sl.addChart(pres.charts.BAR, [{ name: 'Cost Growth %', labels: bmLabels, values: bms.map(b => Number(b.cost_growth_pct || 0)) }], {
      x: 0.5, y: 0.82, w: 4.5, h: 3.4, barDir: 'bar',
      chartColors: ['4A9EFF'], chartArea: { fill: { color: C.bg3 }, roundedCorners: true },
      catAxisLabelColor: C.mid, valAxisLabelColor: C.mid,
      valGridLine: { color: C.dim3, size: 0.5 }, catGridLine: { style: 'none' },
      showValue: true, dataLabelColor: C.white, dataLabelFontSize: 8.5,
      showLegend: false, showTitle: true, title: 'Cost Growth vs Baseline (%)', titleColor: C.mid, titleFontSize: 8.5,
    });
    sl.addChart(pres.charts.BAR, [{ name: 'Schedule Slip (months)', labels: bmLabels, values: bms.map(b => Number(b.schedule_slip_months || 0)) }], {
      x: 5.15, y: 0.82, w: 4.5, h: 3.4, barDir: 'bar',
      chartColors: ['F59E0B'], chartArea: { fill: { color: C.bg3 }, roundedCorners: true },
      catAxisLabelColor: C.mid, valAxisLabelColor: C.mid,
      valGridLine: { color: C.dim3, size: 0.5 }, catGridLine: { style: 'none' },
      showValue: true, dataLabelColor: C.white, dataLabelFontSize: 8.5,
      showLegend: false, showTitle: true, title: 'Schedule Slip (months)', titleColor: C.mid, titleFontSize: 8.5,
    });
  } else {
    // No benchmark data — show the library stats
    const subLabel = safe(m.subsector || m.mode || 'infrastructure', 30);
    sl.addText('CASEY benchmark library: 137+ named programmes across all sectors and geographies.\n\nFor ' + subLabel + ' programmes, typical outcomes from the library:\n• Cost growth median: +22%  ·  P80: +41%\n• Schedule slip median: +18 months  ·  P80: +34 months\n• Top quartile outperformance requires: design freeze at FEED, competitive procurement and independent P80 reserve.',
      { x: 0.5, y: 1.1, w: 9, h: 2.4, fontSize: 11, color: C.light, fontFace: 'Calibri', lineSpacingMultiple: 1.5, margin: 0 }
    );
  }

  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 4.42, w: 9, h: 0.7, fill: { color: C.bg3 }, line: { color: C.dim3 }, rectRadius: 0.07 });
  sl.addText('BENCHMARK POSITION:  ', { x: 0.7, y: 4.5, w: 2.2, h: 0.2, fontSize: 8, bold: true, color: C.dim, fontFace: 'Arial', margin: 0 });
  sl.addText(safe(m.benchmark_position || m.oba_position || 'See full benchmark workbook for cohort comparison and reference-class adjustment.', 160),
    { x: 2.8, y: 4.48, w: 6.5, h: 0.58, fontSize: 9, color: C.light, fontFace: 'Calibri', lineSpacingMultiple: 1.3, margin: 0 });
  sl.addText('CASEY · Programme Intelligence · Confidential', { x: 0.5, y: 5.38, w: 9, h: 0.16, fontSize: 7, color: C.dim3, fontFace: 'Arial', margin: 0 });
  sl.addNotes('BENCHMARK — These are comparable completed programmes. Cost growth % and schedule slip vs original baseline. Your programme\'s position relative to these defines the board\'s risk appetite.');
}

function slide08_Assurance(pres, m, accent) {
  const sl = pres.addSlide();
  sl.background = { color: C.bg };
  const conf = Number(m.confidence_pct || 0);
  const rc = ragColor(conf);

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: rc }, line: { color: rc } });
  sl.addText('07  BOARD ASSURANCE', { x: 0.5, y: 0.18, w: 8, h: 0.22, fontSize: 7.5, color: C.dim, charSpacing: 4, fontFace: 'Arial', margin: 0 });
  sl.addText('Evidence Maturity  ·  Confidence Score  ·  Board Challenge', { x: 0.5, y: 0.38, w: 9, h: 0.28, fontSize: 13, bold: true, color: C.white, fontFace: 'Cambria', margin: 0 });

  // Evidence bars
  const dims = [
    { name: 'Benchmark fit',           score: Math.min(96, Math.max(42, conf + 10)) },
    { name: 'Package evidence',         score: Math.min(92, Math.max(30, conf - 7)) },
    { name: 'Procurement certainty',    score: Math.min(90, Math.max(28, conf - 10)) },
    { name: 'Schedule logic maturity',  score: Math.min(94, Math.max(36, conf + 4)) },
    { name: 'Reserve adequacy',         score: Math.min(96, Math.max(32, conf + 2)) },
  ];
  dims.forEach((d, i) => {
    const y = 0.82 + i * 0.6;
    const sc = d.score >= 75 ? C.green : d.score >= 55 ? C.amber : C.red;
    sl.addText(d.name, { x: 0.5, y: y + 0.06, w: 3.5, h: 0.2, fontSize: 10, color: C.light, fontFace: 'Calibri', margin: 0 });
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 4.1, y: y + 0.1, w: 4.5, h: 0.2, fill: { color: C.bg3 }, line: { color: C.dim3 }, rectRadius: 0.04 });
    const fillW = Math.max(0.05, 4.5 * d.score / 100);
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 4.1, y: y + 0.1, w: fillW, h: 0.2, fill: { color: sc }, line: { color: sc }, rectRadius: 0.04 });
    sl.addText(d.score + '%', { x: 8.72, y: y + 0.06, w: 0.78, h: 0.22, fontSize: 10, bold: true, color: sc, align: 'right', fontFace: 'Cambria', margin: 0 });
  });

  // 75% target dashed line
  sl.addShape(pres.shapes.LINE, { x: 7.475, y: 0.78, w: 0, h: 3.4, line: { color: '2A3A54', width: 0.75, dashType: 'dash' } });
  sl.addText('← 75%', { x: 7.52, y: 0.8, w: 1, h: 0.16, fontSize: 6.5, color: C.dim, fontFace: 'Arial', margin: 0 });

  // Board challenge questions
  sl.addText('BOARD CHALLENGE QUESTIONS', { x: 0.5, y: 3.92, w: 9, h: 0.2, fontSize: 7, color: C.dim, charSpacing: 4, fontFace: 'Arial', margin: 0 });
  const bqs = safeArr(m.board_challenge_questions).slice(0, 3);
  const defaultBqs = [
    'What evidence proves the governing constraint has a named owner and a closure date?',
    'Is P80 reserve funded, named and approved in the project authority?',
    'Which three risks create most P80/P90 exposure and who owns each?',
  ];
  const qs = bqs.length >= 2 ? bqs : defaultBqs;
  qs.slice(0, 3).forEach((q, i) => {
    sl.addText((i + 1) + '.  ' + safe(q, 110), { x: 0.5, y: 4.14 + i * 0.28, w: 9, h: 0.24, fontSize: 9, color: C.light, fontFace: 'Calibri', margin: 0 });
  });

  sl.addText('CASEY · Programme Intelligence · Confidential', { x: 0.5, y: 5.38, w: 9, h: 0.16, fontSize: 7, color: C.dim3, fontFace: 'Arial', margin: 0 });
  sl.addNotes('ASSURANCE — Evidence bars show maturity across 5 dimensions. The 75% target line is the board threshold. The three challenge questions are what an independent reviewer asks first.');
}

function slide09_Scenarios(pres, m, accent) {
  const sl = pres.addSlide();
  sl.background = { color: C.bg };
  const matrix = safeArr(m.scenario_matrix || m.scenario_comparison);
  const pv = getPValues(m);
  const curr = pv.curr;

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: C.cyan }, line: { color: C.cyan } });
  sl.addText('08  SCENARIO ANALYSIS', { x: 0.5, y: 0.18, w: 8, h: 0.22, fontSize: 7.5, color: C.dim, charSpacing: 4, fontFace: 'Arial', margin: 0 });
  sl.addText('Base  ·  Faster  ·  Cheaper  ·  Lower Risk  ·  Premium  —  What the Board Is Really Deciding', { x: 0.5, y: 0.38, w: 9.2, h: 0.28, fontSize: 12, bold: true, color: C.white, fontFace: 'Cambria', margin: 0 });

  const scenarios = [
    { key: 'base',       label: 'BASE',       color: C.mid   },
    { key: 'faster',     label: 'FASTER',     color: C.blue  },
    { key: 'cheaper',    label: 'CHEAPER',    color: C.green },
    { key: 'lower_risk', label: 'LOWER RISK', color: C.purple},
    { key: 'premium',    label: 'PREMIUM',    color: C.amber },
  ];

  scenarios.forEach((s, i) => {
    const row = matrix.find(r => {
      const l = String(r.scenario || r.label || '').toLowerCase();
      return l.includes(s.key.replace('_', ' ').split(' ')[0]);
    });
    const x = 0.5 + i * 1.84;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 0.82, w: 1.72, h: 4.2, fill: { color: C.bg3 }, line: { color: C.dim3 }, rectRadius: 0.07, shadow: mkShadow() });
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 0.82, w: 1.72, h: 0.27, fill: { color: s.color }, line: { color: s.color }, rectRadius: 0.05 });
    sl.addText(s.label, { x, y: 0.82, w: 1.72, h: 0.27, fontSize: 7.5, bold: true, color: C.bg, align: 'center', fontFace: 'Arial', margin: 0 });

    const cost   = row?.cost_p50 || (s.key === 'base' ? (m.cost_p50 || fmtMoney(pv.p50c, curr)) : '—');
    const sched  = row?.schedule_months ? row.schedule_months + ' mo' : (s.key === 'base' ? (m.schedule || pv.p50m + ' mo') : '—');
    const conf   = row?.confidence_pct  ? row.confidence_pct + '%'   : (s.key === 'base' ? (m.confidence_pct || '—') + '%' : '—');
    const risk   = safe(row?.risk || (s.key === 'base' ? (m.risk || '—') : '—'), 16);
    const trade  = safe(row?.scenario_trade || row?.description || (s.key === 'base' ? 'Reference case' : '—'), 45);

    [
      { l: 'COST P50',   v: cost,  size: 12 },
      { l: 'SCHEDULE',   v: sched, size: 11 },
      { l: 'CONFIDENCE', v: conf,  size: 13 },
      { l: 'RISK',       v: risk,  size: 10 },
    ].forEach((mt, mi) => {
      const my = 1.22 + mi * 0.78;
      sl.addText(mt.l, { x: x + 0.1, y: my, w: 1.52, h: 0.16, fontSize: 6.5, color: C.dim, charSpacing: 2, fontFace: 'Arial', margin: 0 });
      sl.addText(mt.v, { x: x + 0.1, y: my + 0.16, w: 1.52, h: 0.32, fontSize: mt.size, bold: true, color: s.color, fontFace: 'Cambria', margin: 0 });
    });
    sl.addText(trade, { x: x + 0.1, y: 4.46, w: 1.52, h: 0.44, fontSize: 7, color: C.dim, fontFace: 'Calibri', lineSpacingMultiple: 1.3, margin: 0 });
  });

  sl.addText('CASEY · Programme Intelligence · Confidential', { x: 0.5, y: 5.38, w: 9, h: 0.16, fontSize: 7, color: C.dim3, fontFace: 'Arial', margin: 0 });
  sl.addNotes('SCENARIOS — Each column is a different board decision. Base is the reference. Faster buys time but costs confidence. Cheaper improves budget but raises delivery risk. Lower Risk buys assurance. Premium buys resilience. The question is which trade-off the board is actually making.');
}

function slide10_Timeline(pres, m, accent) {
  const sl = pres.addSlide();
  sl.background = { color: C.bg };
  const pv = getPValues(m);
  const milestones = getMilestones(m);

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: '2ECC8A' }, line: { color: '2ECC8A' } });
  sl.addText('09  PROGRAMME TIMELINE', { x: 0.5, y: 0.18, w: 8, h: 0.22, fontSize: 7.5, color: C.dim, charSpacing: 4, fontFace: 'Arial', margin: 0 });
  sl.addText('Milestones  ·  Base / Scenario / Benchmark  ·  See CASEY app for animated four-track view', { x: 0.5, y: 0.38, w: 9.2, h: 0.26, fontSize: 12, bold: true, color: C.white, fontFace: 'Cambria', margin: 0 });

  const x0 = 0.5, x1 = 9.1, railW = x1 - x0;
  const railsY = [2.3, 2.7, 3.1];
  const railDefs = [
    { label: 'BASE',      color: '4A9EFF', w: 2.5, mult: 1.0,   dash: false },
    { label: 'SCENARIO',  color: '2ECC8A', w: 1.8, mult: 1.12,  dash: true },
    { label: 'BENCHMARK', color: 'F5A623', w: 1.5, mult: 1.08,  dash: true },
  ];

  // Draw rails
  railDefs.forEach((r, i) => {
    const y = railsY[i];
    const lineW = Math.min(railW * r.mult, railW + 0.8);
    sl.addShape(pres.shapes.LINE, { x: x0, y, w: lineW, h: 0, line: { color: r.color, width: r.w, dashType: r.dash ? 'dash' : 'solid' } });
    sl.addText(r.label, { x: x0 - 0.04, y: y - 0.14, w: 0.9, h: 0.16, fontSize: 6, bold: true, color: r.color, fontFace: 'Arial', margin: 0 });
  });

  // Milestones on base rail
  milestones.forEach((ms, i) => {
    const t = ms.mo / pv.p50m;
    const x = x0 + Math.min(t, 1) * railW;
    const above = i % 2 === 0;
    const py = railsY[0];

    sl.addShape(pres.shapes.OVAL, { x: x - 0.055, y: py - 0.055, w: 0.11, h: 0.11, fill: { color: '2ECC8A' }, line: { color: '2ECC8A' } });
    sl.addShape(pres.shapes.LINE, { x, y: above ? py - 0.055 : py + 0.055, w: 0, h: above ? -0.28 : 0.28, line: { color: '2ECC8A', width: 0.5, dashType: 'sysDot' } });

    const lx = Math.max(x0 - 0.08, Math.min(x - 0.45, x1 - 0.84));
    const ly = above ? py - 0.55 : py + 0.18;
    sl.addText(ms.label, { x: lx, y: ly, w: 0.92, h: 0.2, fontSize: 6.5, color: '2ECC8A', align: 'center', fontFace: 'Arial', margin: 0 });
    sl.addText('M' + ms.mo, { x: lx + 0.08, y: above ? ly - 0.16 : ly + 0.18, w: 0.76, h: 0.15, fontSize: 6, color: C.dim, align: 'center', fontFace: 'Arial', margin: 0 });
  });

  // Delivery flags
  [
    { label: 'BASE',      col: '4A9EFF', xPos: x0 + railW },
    { label: 'SCENARIO',  col: '2ECC8A', xPos: Math.min(x0 + railW * 1.12, x1 + 0.6) },
    { label: 'BENCHMARK', col: 'F5A623', xPos: Math.min(x0 + railW * 1.08, x1 + 0.4) },
  ].forEach(f => {
    if (f.xPos > x0 + 0.2) {
      sl.addShape(pres.shapes.LINE, { x: f.xPos, y: 0.82, w: 0, h: 3.2, line: { color: f.col, width: 0.7, dashType: 'dash' } });
      sl.addText(f.label, { x: Math.min(f.xPos - 0.38, x1 + 0.22), y: 3.56, w: 0.84, h: 0.24, fontSize: 6, bold: true, color: f.col, align: 'center', fontFace: 'Arial', margin: 0 });
    }
  });

  // QSRA P-values legend
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 3.74, w: 9, h: 0.54, fill: { color: C.bg3 }, line: { color: C.dim3 }, rectRadius: 0.07 });
  [
    { l: 'P50 DELIVERY', v: m.schedule || (pv.p50m + ' months'), c: C.blue },
    { l: 'P80 DELIVERY', v: pv.p80m + ' months',                  c: C.amber },
    { l: 'P90 DELIVERY', v: pv.p90m + ' months',                  c: C.red },
    { l: 'GOVERNING CONSTRAINT', v: safe(m.governing_constraint_prominent || '—', 28), c: C.amber },
  ].forEach((kp, i) => {
    const x = 0.7 + i * 2.3;
    sl.addText(kp.l, { x, y: 3.8, w: 2.1, h: 0.16, fontSize: 6.5, color: C.dim, charSpacing: 2, fontFace: 'Arial', margin: 0 });
    sl.addText(kp.v, { x, y: 3.96, w: 2.1, h: 0.22, fontSize: 11, bold: i <= 2, color: kp.c, fontFace: 'Cambria', margin: 0 });
  });

  sl.addText('Open the CASEY app → ◎ Timeline tab for the full animated four-track view with risk events, spend curves and live scenario comparison.', {
    x: 0.5, y: 4.38, w: 9, h: 0.34, fontSize: 8.5, color: C.mid, fontFace: 'Calibri', margin: 0
  });
  sl.addText('CASEY · Programme Intelligence · Confidential', { x: 0.5, y: 5.38, w: 9, h: 0.16, fontSize: 7, color: C.dim3, fontFace: 'Arial', margin: 0 });
  sl.addNotes('TIMELINE — Three rails: Base (contractual), Scenario (optimistic), Benchmark (reference outturn). Point to delivery flag gap — that gap is the scheduling risk conversation. Tell them the full animated version is in the CASEY app Timeline tab.');
}

function slide11_Decision(pres, m, accent) {
  const sl = pres.addSlide();
  sl.background = { color: C.bg };
  const conf = Number(m.confidence_pct || 0);
  const rc = ragColor(conf);

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: rc }, line: { color: rc } });
  sl.addText('10  BOARD DECISION', { x: 0.5, y: 0.18, w: 8, h: 0.22, fontSize: 7.5, color: C.dim, charSpacing: 4, fontFace: 'Arial', margin: 0 });
  sl.addText('Decision Required  ·  Evidence Required  ·  Actions Before Next Board', { x: 0.5, y: 0.38, w: 9.2, h: 0.28, fontSize: 13, bold: true, color: C.white, fontFace: 'Cambria', margin: 0 });

  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 0.82, w: 9, h: 0.96, fill: { color: C.bg3 }, line: { color: rc }, rectRadius: 0.1, shadow: mkShadow() });
  sl.addText(ragLabel(conf), { x: 0.7, y: 0.96, w: 2.4, h: 0.28, fontSize: 14, bold: true, color: rc, fontFace: 'Arial', margin: 0 });
  sl.addText(conf + '% CONFIDENCE', { x: 0.7, y: 1.24, w: 2.4, h: 0.24, fontSize: 10, color: rc, fontFace: 'Arial', margin: 0 });

  const decisionText = safe(
    m.institutional_authority_line || m.governance_decision ||
    (conf >= 75 ? 'Board may proceed to capital commitment subject to evidence package closure.' :
     conf >= 55 ? 'Board may approve option selection. Evidence gap closure required before capital commitment.' :
                  'Board should not approve capital without evidence package, named owners and updated QCRA/QSRA.'),
    220
  );
  sl.addText(decisionText, { x: 3.3, y: 0.90, w: 6.0, h: 0.78, fontSize: 9.5, color: C.light, fontFace: 'Calibri', lineSpacingMultiple: 1.4, margin: 0 });

  sl.addText('REQUIRED ACTIONS BEFORE NEXT BOARD', { x: 0.5, y: 2.02, w: 9, h: 0.2, fontSize: 7, color: C.dim, charSpacing: 4, fontFace: 'Arial', margin: 0 });

  const actions = safeArr(m.top_decisions_required || m.next_steps).slice(0, 5);
  const defaultActions = [
    'Confirm the governing critical-path constraint and name the evidence owner with a closure date.',
    'Verify P80 reserve is funded and formally approved in the project authority.',
    'Close evidence gaps on top 3 risks before capital approval — each needs a named owner.',
    'Confirm schedule float is operationally accessible, not theoretical.',
    'Approve stage gate progression with all conditions documented and signed.',
  ];
  const actionList = actions.length >= 3 ? actions : defaultActions;
  actionList.slice(0, 5).forEach((a, i) => {
    const y = 2.26 + i * 0.44;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y, w: 0.3, h: 0.28, fill: { color: rc }, line: { color: rc }, rectRadius: 0.04 });
    sl.addText(String(i + 1), { x: 0.5, y, w: 0.3, h: 0.28, fontSize: 8, bold: true, color: C.bg, align: 'center', fontFace: 'Arial', margin: 0 });
    sl.addText(safe(a, 130), { x: 0.9, y: y + 0.02, w: 8.7, h: 0.26, fontSize: 9, color: C.light, fontFace: 'Calibri', margin: 0 });
  });

  sl.addText('CASEY · Programme Intelligence · Confidential', { x: 0.5, y: 5.38, w: 9, h: 0.16, fontSize: 7, color: C.dim3, fontFace: 'Arial', margin: 0 });
  sl.addNotes('BOARD DECISION — This is the ask. Lead with the verdict badge and confidence score. Walk through required actions in numbered order. Each action needs a named owner before the next board meeting.');
}

function slide12_Recovery(pres, m, accent) {
  const sl = pres.addSlide();
  sl.background = { color: C.bg };

  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: C.green }, line: { color: C.green } });
  sl.addText('11  RECOVERY & NEXT STEPS', { x: 0.5, y: 0.18, w: 8, h: 0.22, fontSize: 7.5, color: C.dim, charSpacing: 4, fontFace: 'Arial', margin: 0 });
  sl.addText('Recovery Potential  ·  Priority Actions  ·  Confidence Gain', { x: 0.5, y: 0.38, w: 9.2, h: 0.28, fontSize: 13, bold: true, color: C.white, fontFace: 'Cambria', margin: 0 });

  const recovText = safe(m.recovery_opportunities || m.recovery_plan || m.scenario_gain ||
    'Recovery potential depends on early action on the governing constraint, design freeze and procurement strategy. CASEY models indicate these actions have the greatest confidence impact.', 260);
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 0.82, w: 9, h: 0.7, fill: { color: C.bg3 }, line: { color: C.green }, rectRadius: 0.08 });
  sl.addText('RECOVERY POTENTIAL', { x: 0.7, y: 0.9, w: 3, h: 0.2, fontSize: 7, color: C.dim, charSpacing: 4, fontFace: 'Arial', margin: 0 });
  sl.addText(recovText, { x: 0.7, y: 1.08, w: 8.4, h: 0.36, fontSize: 9, color: C.light, fontFace: 'Calibri', margin: 0 });

  // Recovery paths
  const paths = safeArr(m.recovery_paths).slice(0, 3);
  const mode = String(m.mode || '').toLowerCase();
  const sub  = String(m.subsector || m.sector || '').toLowerCase();
  const isSpace = mode === 'space' || sub.includes('space') || sub.includes('lunar') || sub.includes('mars');
  const isPharma = sub.includes('life') || sub.includes('pharma') || sub.includes('bio') || sub.includes('glp');
  const isData  = sub.includes('data') || sub.includes('hyperscale');

  const defaultPaths = isSpace ? [
    { label: 'Lock mission architecture', impact: 'Up to 6 months recovery',     conf: '+10 pts confidence', priority: 'P1' },
    { label: 'Secure launch slots',       impact: 'Reduces schedule tail risk',   conf: '+8 pts confidence',  priority: 'P2' },
    { label: 'Advance MAIT readiness',    impact: 'Reduce P90 exposure',          conf: '+6 pts confidence',  priority: 'P3' },
  ] : isPharma ? [
    { label: 'Early regulatory filing',   impact: 'Up to 4 months recovery',     conf: '+9 pts confidence',  priority: 'P1' },
    { label: 'Lock OEM equipment',        impact: 'Reduce lead-time risk',        conf: '+7 pts confidence',  priority: 'P2' },
    { label: 'Pre-qualify PQ batches',    impact: 'Reduce P90 exposure',          conf: '+5 pts confidence',  priority: 'P3' },
  ] : isData ? [
    { label: 'Sign grid connection',      impact: 'Up to 6 months recovery',     conf: '+11 pts confidence', priority: 'P1' },
    { label: 'Reserve GPU allocation',    impact: 'Reduce procurement risk',      conf: '+8 pts confidence',  priority: 'P2' },
    { label: 'Advance cooling design',    impact: 'Reduce commissioning risk',    conf: '+5 pts confidence',  priority: 'P3' },
  ] : [
    { label: 'Freeze Design',            impact: 'Up to 3 months schedule recovery', conf: '+8 pts confidence',  priority: 'P1' },
    { label: 'Secure Vendor Slots',      impact: 'Up to 5 months lead-time reduction', conf: '+6 pts confidence', priority: 'P2' },
    { label: 'Strengthen Commissioning', impact: 'Reduce P90 exposure',             conf: '+4 pts confidence',  priority: 'P3' },
  ];

  const rpList = paths.length >= 1 ? paths : defaultPaths;
  rpList.slice(0, 3).forEach((rp, i) => {
    const x = 0.5 + i * 3.08;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.74, w: 2.9, h: 2.56, fill: { color: C.bg3 }, line: { color: C.green }, rectRadius: 0.08, shadow: mkShadow() });
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.1, y: 1.82, w: 0.34, h: 0.26, fill: { color: C.green }, line: { color: C.green }, rectRadius: 0.04 });
    sl.addText(rp.priority || ('P' + (i+1)), { x: x + 0.1, y: 1.82, w: 0.34, h: 0.26, fontSize: 8, bold: true, color: C.bg, align: 'center', fontFace: 'Arial', margin: 0 });
    sl.addText(safe(rp.label || rp.action || rp.name || 'Action ' + (i+1), 32), { x: x + 0.54, y: 1.86, w: 2.26, h: 0.24, fontSize: 10, bold: true, color: C.white, fontFace: 'Calibri', margin: 0 });
    sl.addText(safe(rp.impact || rp.description || '', 65), { x: x + 0.14, y: 2.2, w: 2.64, h: 0.52, fontSize: 9, color: C.light, fontFace: 'Calibri', lineSpacingMultiple: 1.3, margin: 0 });
    sl.addText(safe(rp.conf || rp.confidence_delta || '', 40), { x: x + 0.14, y: 2.78, w: 2.64, h: 0.24, fontSize: 9, bold: true, color: C.green, fontFace: 'Calibri', margin: 0 });
    sl.addText('Named owner required', { x: x + 0.14, y: 3.08, w: 2.64, h: 0.7, fontSize: 7.5, color: C.dim, fontFace: 'Calibri', margin: 0 });
  });

  sl.addText('CASEY · Programme Intelligence · Confidential', { x: 0.5, y: 5.38, w: 9, h: 0.16, fontSize: 7, color: C.dim3, fontFace: 'Arial', margin: 0 });
  sl.addNotes('RECOVERY — Three priority actions. P1 is the most important. Each needs a named owner and a date. The confidence gain per action tells the board where to focus first.');
}

function slide13_BackCover(pres, m, accent) {
  const sl = pres.addSlide();
  sl.background = { color: C.bg };
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.07, fill: { color: accent }, line: { color: accent } });

  sl.addText('CASEY', { x: 3.2, y: 1.4, w: 3.6, h: 0.7, fontSize: 42, bold: true, color: C.white, align: 'center', fontFace: 'Cambria', charSpacing: 8, margin: 0 });
  sl.addText('PROGRAMME INTELLIGENCE', { x: 2.4, y: 2.08, w: 5.2, h: 0.28, fontSize: 9, color: C.cyan, align: 'center', charSpacing: 5, fontFace: 'Arial', margin: 0 });
  sl.addShape(pres.shapes.RECTANGLE, { x: 3.0, y: 2.5, w: 4, h: 0.005, fill: { color: C.dim3 }, line: { color: C.dim3 } });
  sl.addText('The Bloomberg Terminal for Project Delivery', { x: 1.2, y: 2.66, w: 7.6, h: 0.28, fontSize: 11, color: C.mid, align: 'center', fontFace: 'Calibri', margin: 0 });

  const today = new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });
  sl.addText(safe(m.title || m.subsector || 'Programme') + '  ·  ' + today, { x: 1.2, y: 3.14, w: 7.6, h: 0.24, fontSize: 9, color: C.dim, align: 'center', fontFace: 'Arial', margin: 0 });
  sl.addText(sectorTag(m), { x: 1.2, y: 3.42, w: 7.6, h: 0.22, fontSize: 8.5, color: C.dim, align: 'center', fontFace: 'Arial', margin: 0 });
  sl.addText('First-pass programme intelligence. Not a certified estimate document.', { x: 1.2, y: 3.72, w: 7.6, h: 0.22, fontSize: 8, color: C.dim3, align: 'center', fontFace: 'Arial', margin: 0 });

  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 2.8, y: 4.18, w: 4.4, h: 0.58, fill: { color: C.bg3 }, line: { color: C.dim3 }, rectRadius: 0.08 });
  sl.addText('controlorbit.com  ·  deepa@caseai.co.uk', { x: 2.8, y: 4.18, w: 4.4, h: 0.58, fontSize: 9, color: C.cyan, align: 'center', fontFace: 'Arial', margin: 0 });
  sl.addNotes('BACK COVER — Leave time for questions. Key message: CASEY combines benchmark intelligence, QCRA/QSRA and AI to transform project data into board decisions in 30 seconds. No consultant, no workshop, no report. Just the decision.');
}

// ── MAIN ─────────────────────────────────────────────────────────────────────
async function generateBoardPack(model, outputPath) {
  const pres = new pptxgen();
  pres.layout   = 'LAYOUT_16x9';
  pres.author   = 'CASEY Programme Intelligence';
  pres.company  = 'Control Orbit';
  pres.title    = 'CASEY Board Pack — ' + safe(model.title || model.subsector || 'Programme', 60);
  pres.subject  = sectorTag(model);

  const accent = sectorAccent(model);

  slide01_Cover(pres, model, accent);
  slide02_ExecutiveSummary(pres, model, accent);
  slide03_CostModel(pres, model, accent);
  slide04_Schedule(pres, model, accent);
  slide05_RiskRegister(pres, model, accent);
  slide06_QCRA_QSRA(pres, model, accent);
  slide07_Benchmark(pres, model, accent);
  slide08_Assurance(pres, model, accent);
  slide09_Scenarios(pres, model, accent);
  slide10_Timeline(pres, model, accent);
  slide11_Decision(pres, model, accent);
  slide12_Recovery(pres, model, accent);
  slide13_BackCover(pres, model, accent);

  const out = outputPath || 'CASEY_Board_Pack.pptx';
  await pres.writeFile({ fileName: out });
  console.log('✓ Board pack: ' + out + ' · 13 slides · Sector: ' + safe(model.subsector || model.mode) + ' · Accent: #' + accent);
  return out;
}

if (require.main === module) {
  const modelPath = process.argv[2];
  const outPath   = process.argv[3] || 'CASEY_Board_Pack.pptx';
  if (!modelPath) { console.error('Usage: node generate_board_pack.js <model.json> [output.pptx]'); process.exit(1); }
  const model = JSON.parse(fs.readFileSync(modelPath, 'utf8'));
  generateBoardPack(model, outPath).then(() => process.exit(0)).catch(e => { console.error(e); process.exit(1); });
}

module.exports = { generateBoardPack };
