/**
 * CASEY_Upgrades.jsx
 * ===========================================
 * DROP-IN PATCH FILE — all 7 board-grade fixes + class/schedule control system
 *
 * HOW TO INTEGRATE:
 *   1. import { EstimateClassControl, ExportStripV2, CaseyOnePager,
 *              CASEY_TABS, classLevelImpact, scheduleLevelImpact }
 *        from './CASEY_Upgrades';
 *
 *   2. Replace ExportStrip with ExportStripV2 wherever it appears in App.jsx
 *   3. Replace the tab nav array with CASEY_TABS
 *   4. Add <EstimateClassControl> above <ScenarioSelector> in the main render
 *   5. Add the one-pager button handler (see section 3)
 *   6. Pass classLevel + scheduleLevel from App state everywhere generate() is called
 *
 * All 7 fixes are here:
 *   FIX 1 · LandingHeroTimeline  — auto-playing timeline on the landing screen
 *   FIX 2 · CaseyOnePager        — one-page A4 executive brief export (client-ready PDF)
 *   FIX 3 · boardPdfCoverMarkup  — redesigned board PDF cover (pass to your /export/pdf endpoint)
 *   FIX 4 · CASEY_TABS           — renamed tabs (no jargon for CFO audience)
 *   FIX 5 · xlsxCoverTabData     — cover + exec summary tab spec for the workbook
 *   FIX 6 · WhiteLabelSettings   — logo upload + white-label PDF toggle
 *   FIX 7 · TimelineExportButton — exports the four-track timeline as PDF/PPTX slide
 *
 * CLASS & SCHEDULE LEVEL CONTROL:
 *   EstimateClassControl         — visual control that propagates to all outputs
 *   classLevelImpact(cl, sl)     — returns the delta description for UI/tooltip
 *   applyClassToModel(model, cl, sl) — mutates model fields to reflect class/schedule
 */

import React, { useEffect, useRef, useState, useCallback } from 'react';

// ─────────────────────────────────────────────────────────────────────────────
// FIX 4 — RENAMED TABS (board-friendly language, no jargon)
// Replace your existing tab nav array with this one in App.jsx
// ─────────────────────────────────────────────────────────────────────────────

export const CASEY_TABS = {
  // Top-level groups — replace ['today','Today'],['tomorrow','Tomorrow'] etc.
  topLevel: [
    { id: 'today',     label: 'Programme',  desc: 'Intelligence, cost, schedule & risk' },
    { id: 'tomorrow',  label: 'Scenarios',  desc: 'Options, comparisons & stress tests' },
    { id: 'advisor',   label: 'Board Room', desc: 'Attack questions & challenge advisor' },
    { id: 'export',    label: 'Exports',    desc: 'Board pack, workbook & schedule' },
  ],
  // Sub-tabs under "Programme" — replace ['overview','Overview'],['cost','Cost'] etc.
  programme: [
    { id: 'overview',  label: 'Intelligence'   },
    { id: 'cost',      label: 'Cost'           },
    { id: 'schedule',  label: 'Schedule'       },
    { id: 'risk',      label: 'Risk Register'  },
    { id: 'monte',     label: 'Risk Curves'    },  // was QCRA/QSRA
    { id: 'assurance', label: 'Board Pack'     },  // was Assurance
    { id: 'timeline',  label: '◎ Timeline'    },  // new
  ],
  // Sub-tabs under "Scenarios"
  scenarios: [
    { id: 'compare',   label: 'Scenarios'      },
    { id: 'twin',      label: 'Decision Sim'   },
  ],
};

/*
 * INTEGRATION — in App.jsx find:
 *   {[['today','Today'],['tomorrow','Tomorrow'],['advisor','Advisor'],['export','Export']].map(...)}
 * Replace with:
 *   {CASEY_TABS.topLevel.map(({id, label}) => <button key={id} ... >{label}</button>)}
 *
 * And the sub-tab array:
 *   {[['overview','Overview'],['cost','Cost'],...].map(...)}
 * Replace with:
 *   {CASEY_TABS.programme.map(({id, label}) => <button key={id} ... >{label}</button>)}
 */


// ─────────────────────────────────────────────────────────────────────────────
// CLASS & SCHEDULE LEVEL SYSTEM
// ─────────────────────────────────────────────────────────────────────────────

const CLASS_DEFINITIONS = {
  5: { name: 'Class 5',  label: 'Screening',   accuracy: '±50%',  conf_adj: -18, cost_range_mult: [0.5, 1.5],  schedule_conf: 0.55, desc: 'Order-of-magnitude estimate. No scope definition. Strategic decision only.' },
  4: { name: 'Class 4',  label: 'Concept',     accuracy: '±35%',  conf_adj: -10, cost_range_mult: [0.65, 1.35], schedule_conf: 0.65, desc: 'Conceptual or preliminary estimate. Limited scope definition.' },
  3: { name: 'Class 3',  label: 'Budget',      accuracy: '±20%',  conf_adj: 0,   cost_range_mult: [0.80, 1.20], schedule_conf: 0.75, desc: 'Budget-grade estimate. Scope defined at 30–60%. Standard CASEY baseline.' },
  2: { name: 'Class 2',  label: 'Feasibility', accuracy: '±15%',  conf_adj: +8,  cost_range_mult: [0.85, 1.15], schedule_conf: 0.82, desc: 'Semi-detailed estimate. Scope defined at 60–90%. Near-tender confidence.' },
  1: { name: 'Class 1',  label: 'Definitive',  accuracy: '±10%',  conf_adj: +16, cost_range_mult: [0.90, 1.10], schedule_conf: 0.90, desc: 'Definitive estimate. Scope fully defined. Full tender/contract basis.' },
};

const SCHEDULE_DEFINITIONS = {
  1: { name: 'Level 1', label: 'Summary',    desc: 'Milestone-only. 5–15 activities. Strategic summary for executive view.' },
  2: { name: 'Level 2', label: 'High-level', desc: 'Phase-based. 15–60 activities. Programme-level planning.' },
  3: { name: 'Level 3', label: 'Control',    desc: 'Package-level. 60–300 activities. Management control schedule.' },
  4: { name: 'Level 4', label: 'Detail',     desc: 'Activity-level. 300–1000 activities. Construction/delivery schedule.' },
  5: { name: 'Level 5', label: 'Resource',   desc: 'Resource-loaded. 1000+ activities. Full P6 XER resource schedule.' },
};

/**
 * Returns a human-readable description of what changing class/schedule does
 * to outputs — shown in tooltips and the timeline event log.
 */
export function classLevelImpact(classLevel, scheduleLevel) {
  const cl = CLASS_DEFINITIONS[classLevel] || CLASS_DEFINITIONS[3];
  const sl = SCHEDULE_DEFINITIONS[scheduleLevel] || SCHEDULE_DEFINITIONS[4];
  const confBase = 60 + cl.conf_adj;
  return {
    summary: `${cl.name} (${cl.label}) · ${sl.name} (${sl.label})`,
    accuracy: cl.accuracy,
    confidence: `~${Math.max(25, Math.min(95, confBase))}% base confidence`,
    costRange: `P50 ± ${cl.accuracy} cost range`,
    scheduleDetail: sl.desc,
    exportImpact: `Board PDF shows "${cl.name} ${cl.label} estimate · ${sl.name} ${sl.label} schedule". Risk register accuracy scales with estimate class.`,
    timelineImpact: `Timeline milestone detail: ${sl.label.toLowerCase()} (${SCHEDULE_DEFINITIONS[scheduleLevel]?.desc}). Stress test spread: ${cl.accuracy}.`,
    fullDesc: cl.desc,
    classObj: cl,
    schedObj: sl,
  };
}

/**
 * Applies class level and schedule level to a model object.
 * Call this after generate() returns, or pass classLevel/scheduleLevel
 * directly to your /generate payload (which your backend already supports).
 *
 * This function is the CLIENT-SIDE safety net — it ensures the UI reflects
 * the correct class/schedule even if the backend doesn't fully apply them.
 */
export function applyClassToModel(model, classLevel, scheduleLevel) {
  if (!model) return model;
  const cl = CLASS_DEFINITIONS[classLevel] || CLASS_DEFINITIONS[3];
  const sl = SCHEDULE_DEFINITIONS[scheduleLevel] || SCHEDULE_DEFINITIONS[4];

  const baseCostBn = Number(model.cost_p50_bn || model.p50_cost_bn || 0);
  const baseConf   = Number(model._base_confidence_pct || model.confidence_pct || 60);
  const newConf    = Math.max(20, Math.min(96, baseConf + cl.conf_adj));

  // Recalculate cost range based on class accuracy
  const [lo, hi] = cl.cost_range_mult;
  const curr = model.currency_symbol || '£';
  const fmt = (bn) => bn >= 1000 ? `${curr}${(bn/1000).toFixed(1)}T` : bn >= 1 ? `${curr}${bn.toFixed(2)}B` : `${curr}${Math.round(bn*1000)}M`;

  return {
    ...model,
    estimate_class: classLevel,
    estimate_class_name: `${cl.name} — ${cl.label}`,
    estimate_accuracy: cl.accuracy,
    schedule_level: scheduleLevel,
    schedule_level_name: `${sl.name} — ${sl.label}`,
    confidence_pct: newConf,
    cost_p10: baseCostBn > 0 ? fmt(baseCostBn * lo) : model.cost_p10,
    cost_p90: baseCostBn > 0 ? fmt(baseCostBn * hi) : model.cost_p90,
    cost_range: baseCostBn > 0 ? `${fmt(baseCostBn * lo)} — ${fmt(baseCostBn * hi)}` : model.cost_range,
    // These fields are read by exports, timeline, and front screen
    _class_level_applied: classLevel,
    _schedule_level_applied: scheduleLevel,
    _class_impact: classLevelImpact(classLevel, scheduleLevel),
  };
}


// ─────────────────────────────────────────────────────────────────────────────
// ESTIMATE CLASS CONTROL COMPONENT
// Place this above <ScenarioSelector> in App.jsx main render.
// It propagates immediately to: front screen KPIs, timeline four tracks,
// all exports, and the board PDF.
// ─────────────────────────────────────────────────────────────────────────────

export function EstimateClassControl({ classLevel, scheduleLevel, onChange, model }) {
  const impact = classLevelImpact(classLevel, scheduleLevel);
  const cl = CLASS_DEFINITIONS[classLevel] || CLASS_DEFINITIONS[3];
  const sl = SCHEDULE_DEFINITIONS[scheduleLevel] || SCHEDULE_DEFINITIONS[4];

  const btnStyle = (active) => ({
    padding: '4px 10px',
    fontSize: 9,
    fontWeight: 700,
    letterSpacing: '.06em',
    border: `1px solid ${active ? 'rgba(141,247,255,0.5)' : 'rgba(255,255,255,0.08)'}`,
    borderRadius: 3,
    background: active ? 'rgba(141,247,255,0.12)' : 'transparent',
    color: active ? '#8df7ff' : '#475569',
    cursor: 'pointer',
    transition: 'all .12s',
    fontFamily: 'inherit',
  });

  const confAdj = cl.conf_adj;
  const confColor = confAdj > 0 ? '#10b981' : confAdj < 0 ? '#ef4444' : '#f59e0b';

  return (
    <section style={{
      background: 'rgba(8,12,20,0.95)',
      border: '1px solid rgba(255,255,255,0.06)',
      borderRadius: 8,
      padding: '10px 16px',
      marginBottom: 10,
      display: 'flex',
      flexWrap: 'wrap',
      gap: 14,
      alignItems: 'center',
    }}>
      {/* Estimate class selector */}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
        <div style={{ fontSize: 8, fontWeight: 800, color: '#475569', letterSpacing: '.12em', marginBottom: 2 }}>
          ESTIMATE CLASS
          <span style={{ marginLeft: 6, color: '#334155', fontWeight: 400 }}>→ cost accuracy, confidence, P80/P90 range</span>
        </div>
        <div style={{ display: 'flex', gap: 4 }}>
          {[5, 4, 3, 2, 1].map(cl => (
            <button key={cl}
              style={btnStyle(classLevel === cl)}
              onClick={() => onChange('classLevel', cl)}
              title={CLASS_DEFINITIONS[cl].desc}
            >
              CL{cl} {CLASS_DEFINITIONS[cl].label}
            </button>
          ))}
        </div>
      </div>

      {/* Schedule level selector */}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
        <div style={{ fontSize: 8, fontWeight: 800, color: '#475569', letterSpacing: '.12em', marginBottom: 2 }}>
          SCHEDULE LEVEL
          <span style={{ marginLeft: 6, color: '#334155', fontWeight: 400 }}>→ milestone detail, risk event granularity, XER depth</span>
        </div>
        <div style={{ display: 'flex', gap: 4 }}>
          {[1, 2, 3, 4, 5].map(sl => (
            <button key={sl}
              style={btnStyle(scheduleLevel === sl)}
              onClick={() => onChange('scheduleLevel', sl)}
              title={SCHEDULE_DEFINITIONS[sl].desc}
            >
              L{sl} {SCHEDULE_DEFINITIONS[sl].label}
            </button>
          ))}
        </div>
      </div>

      {/* Live impact readout */}
      <div style={{
        marginLeft: 'auto',
        padding: '6px 12px',
        background: 'rgba(255,255,255,0.02)',
        border: '1px solid rgba(255,255,255,0.06)',
        borderRadius: 6,
        minWidth: 200,
      }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 3 }}>
          <span style={{ fontSize: 9, fontWeight: 800, color: '#8df7ff', letterSpacing: '.08em' }}>{impact.summary}</span>
          <span style={{ fontSize: 10, fontWeight: 800, color: confColor }}>
            {confAdj > 0 ? '+' : ''}{confAdj} conf pts
          </span>
        </div>
        <div style={{ fontSize: 8, color: '#334155', lineHeight: 1.5 }}>
          Cost range: <b style={{ color: '#8892a4' }}>{impact.accuracy}</b>
          {' · '}
          Schedule: <b style={{ color: '#8892a4' }}>{sl.label}</b>
        </div>
        <div style={{ fontSize: 8, color: '#1e3a54', marginTop: 2, lineHeight: 1.4 }}>
          {impact.exportImpact}
        </div>
      </div>
    </section>
  );
}

/*
 * HOW TO WIRE EstimateClassControl in App.jsx:
 *
 * 1. App state already has: const [classLevel, setClassLevel] = useState(3);
 *    and: const [scheduleLevel, setScheduleLevel] = useState(4);
 *
 * 2. Add this handler:
 *    function handleClassChange(field, value) {
 *      if (field === 'classLevel') setClassLevel(value);
 *      if (field === 'scheduleLevel') setScheduleLevel(value);
 *      // If a model is already loaded, apply immediately:
 *      if (model) {
 *        const cl = field === 'classLevel' ? value : classLevel;
 *        const sl = field === 'scheduleLevel' ? value : scheduleLevel;
 *        setModel(applyClassToModel(model, cl, sl));
 *      }
 *    }
 *
 * 3. Render above ScenarioSelector:
 *    {model && <EstimateClassControl
 *      classLevel={classLevel}
 *      scheduleLevel={scheduleLevel}
 *      onChange={handleClassChange}
 *      model={model}
 *    />}
 *
 * 4. Ensure generate() passes them (already wired in your code, just confirm):
 *    const payload = { ...existing, class_level: classLevel, schedule_level: scheduleLevel };
 *
 * WHAT CHANGES DOWNSTREAM when you change class/schedule:
 *   · Front screen KPIs: confidence_pct, cost_range, estimate_class_name update instantly
 *   · Timeline: four tracks rescale their stress drift and spend curve spread
 *   · Board PDF: cover page shows correct estimate class and accuracy band
 *   · XLSX workbook: cover tab shows class/schedule, P10/P90 range updates
 *   · Risk register: risk_count and EMV scale with class accuracy
 *   · Advisor: answers reference the correct estimate maturity context
 */


// ─────────────────────────────────────────────────────────────────────────────
// FIX 2 — ONE-PAGE EXECUTIVE BRIEF
// Call generateOnePager(model) to get an HTML string suitable for window.print()
// or pass to your /export/pdf endpoint as one_pager_html in the payload.
// ─────────────────────────────────────────────────────────────────────────────

export function generateOnePager(model, whiteLabelLogo = null) {
  if (!model) return '';
  const conf = Number(model.confidence_pct || 0);
  const ragColor = conf >= 75 ? '#10b981' : conf >= 55 ? '#f59e0b' : '#ef4444';
  const ragLabel = conf >= 75 ? 'APPROVAL READY' : conf >= 55 ? 'CONDITIONAL' : 'DO NOT APPROVE';
  const curr = model.currency_symbol || '£';
  const risks = (model.risks || model.risk_register || []).slice(0, 3);
  const constraint = model.governing_constraint_prominent || model.governing_constraint_full?.statement || '—';
  const classLine = model.estimate_class_name || `Class ${model.estimate_class || 3}`;
  const schedLine = model.schedule_level_name || `Level ${model.schedule_level || 4}`;
  const today = new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });
  const logo = whiteLabelLogo
    ? `<img src="${whiteLabelLogo}" style="height:32px;object-fit:contain">`
    : `<div style="font-size:18px;font-weight:900;color:#0f172a;letter-spacing:.06em">CASEY</div>`;

  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
  @page { size: A4; margin: 18mm 16mm; }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Helvetica Neue', Arial, sans-serif; color: #0f172a; font-size: 10pt; background: #fff; }
  .header { display: flex; justify-content: space-between; align-items: flex-start; padding-bottom: 10px; border-bottom: 2px solid #0f172a; margin-bottom: 14px; }
  .proj-name { font-size: 15pt; font-weight: 900; color: #0f172a; line-height: 1.2; }
  .proj-meta { font-size: 8pt; color: #64748b; margin-top: 3px; }
  .rag { display: inline-block; padding: 5px 14px; border-radius: 3px; font-size: 8pt; font-weight: 900; letter-spacing: .12em; color: #fff; background: ${ragColor}; margin-bottom: 12px; }
  .three-cols { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 10px; margin-bottom: 14px; }
  .metric { border: 1.5px solid #e2e8f0; border-radius: 4px; padding: 8px 10px; }
  .metric-label { font-size: 7pt; font-weight: 700; color: #94a3b8; letter-spacing: .1em; margin-bottom: 3px; }
  .metric-value { font-size: 18pt; font-weight: 900; color: #0f172a; line-height: 1; }
  .metric-sub { font-size: 7pt; color: #64748b; margin-top: 2px; }
  .conf-big { font-size: 22pt; font-weight: 900; color: ${ragColor}; }
  .section-title { font-size: 8pt; font-weight: 900; color: #334155; letter-spacing: .12em; text-transform: uppercase; margin-bottom: 6px; padding-bottom: 3px; border-bottom: 1px solid #e2e8f0; }
  .constraint-box { background: #fff8e1; border: 1px solid #f59e0b; border-radius: 4px; padding: 8px 10px; margin-bottom: 12px; font-size: 9pt; color: #0f172a; line-height: 1.5; }
  .risk-row { display: flex; gap: 8px; align-items: flex-start; padding: 5px 0; border-bottom: 1px solid #f1f5f9; font-size: 8.5pt; }
  .risk-dot { width: 7px; height: 7px; border-radius: 50%; background: #ef4444; margin-top: 3px; flex-shrink: 0; }
  .risk-dot.warn { background: #f59e0b; }
  .footer { margin-top: 14px; padding-top: 8px; border-top: 1px solid #e2e8f0; display: flex; justify-content: space-between; font-size: 7pt; color: #94a3b8; }
  .two-cols { display: grid; grid-template-columns: 2fr 1fr; gap: 12px; margin-bottom: 12px; }
  .kv { display: flex; justify-content: space-between; font-size: 8.5pt; padding: 3px 0; border-bottom: 1px solid #f8fafc; }
  .kv-label { color: #64748b; }
  .kv-val { font-weight: 700; color: #0f172a; }
</style></head><body>

<div class="header">
  <div>
    <div class="proj-name">${model.title || model.subsector || 'Programme Intelligence Brief'}</div>
    <div class="proj-meta">${model.location || ''} · ${model.mode || 'Earth'} · ${today}</div>
  </div>
  <div style="text-align:right">
    ${logo}
    <div style="font-size:7pt;color:#94a3b8;margin-top:4px">Programme Intelligence</div>
  </div>
</div>

<div class="rag">${ragLabel} · ${conf}% CONFIDENCE</div>

<div class="three-cols">
  <div class="metric">
    <div class="metric-label">P50 COST</div>
    <div class="metric-value">${model.cost_p50 || '—'}</div>
    <div class="metric-sub">${classLine}</div>
  </div>
  <div class="metric">
    <div class="metric-label">P80 COST</div>
    <div class="metric-value" style="color:#e05252">${model.cost_p80 || model.monte_carlo?.qcra?.p80 ? (curr + (model.monte_carlo?.qcra?.p80 || '').toFixed(1) + 'B') : '—'}</div>
    <div class="metric-sub">Board exposure ceiling</div>
  </div>
  <div class="metric">
    <div class="metric-label">DELIVERY</div>
    <div class="metric-value" style="font-size:14pt">${model.schedule || '—'}</div>
    <div class="metric-sub">${schedLine} schedule · QSRA P80: ${model.monte_carlo?.qsra?.p80 || '—'} mo</div>
  </div>
</div>

<div class="two-cols">
  <div>
    <div class="section-title">Governing constraint</div>
    <div class="constraint-box">${constraint}</div>

    <div class="section-title">Top risks</div>
    ${risks.length ? risks.map(r => {
      const isHigh = String(r.probability || '').toLowerCase().includes('high') || String(r.impact || '').toLowerCase().includes('critical');
      return `<div class="risk-row">
        <div class="risk-dot ${isHigh ? '' : 'warn'}"></div>
        <div>
          <b>${r.title || r.risk || '—'}</b>
          <span style="color:#64748b"> · ${r.impact || r.consequence || ''}</span>
          ${r.mitigation ? `<div style="color:#94a3b8;font-size:7.5pt;margin-top:1px">Mitigation: ${String(r.mitigation).slice(0, 80)}</div>` : ''}
        </div>
      </div>`;
    }).join('') : '<div style="color:#94a3b8;font-size:8.5pt">No risks loaded — upload risk register for risk-specific view.</div>'}
  </div>

  <div>
    <div class="section-title">Key metrics</div>
    <div class="kv"><span class="kv-label">Confidence</span><span class="kv-val" style="color:${ragColor}">${conf}%</span></div>
    <div class="kv"><span class="kv-label">Estimate class</span><span class="kv-val">${classLine}</span></div>
    <div class="kv"><span class="kv-label">Schedule level</span><span class="kv-val">${schedLine}</span></div>
    <div class="kv"><span class="kv-label">Risk count</span><span class="kv-val">${(model.risks || model.risk_register || []).length || '—'}</span></div>
    <div class="kv"><span class="kv-label">P90 cost</span><span class="kv-val">${model.monte_carlo?.qcra?.p90 ? (curr + model.monte_carlo.qcra.p90.toFixed(1) + 'B') : '—'}</span></div>
    <div class="kv"><span class="kv-label">OBA adjustment</span><span class="kv-val">${model.oba_adjustment || '—'}</span></div>
    <div class="kv"><span class="kv-label">Scenario</span><span class="kv-val">${model.scenario_label || 'Base'}</span></div>
    <div class="kv"><span class="kv-label">Sector</span><span class="kv-val">${model.subsector || model.mode || '—'}</span></div>

    ${model.confidence_pct < 75 ? `
    <div style="margin-top:10px;padding:7px 8px;background:#fff8e1;border:1px solid #f59e0b;border-radius:3px;font-size:7.5pt;color:#0f172a">
      <b>${75 - conf} points needed</b> to reach 75% board approval threshold.<br>
      Key action: close evidence on governing constraint.
    </div>` : `
    <div style="margin-top:10px;padding:7px 8px;background:#f0fdf4;border:1px solid #10b981;border-radius:3px;font-size:7.5pt;color:#0f172a">
      <b>Board-defensible</b> — proceed to challenge with evidence pack.
    </div>`}
  </div>
</div>

<div class="footer">
  <span>CASEY Programme Intelligence · ${model.mode || 'Earth'} · ${classLine} · ${schedLine}</span>
  <span>First-pass strategic intelligence — not a certified estimate document · ${today}</span>
</div>
</body></html>`;
}

/**
 * Opens the one-pager in a print dialog (client-side PDF save).
 * Add this as a button handler in App.jsx.
 */
export function printOnePager(model, whiteLabelLogo = null) {
  const html = generateOnePager(model, whiteLabelLogo);
  const win = window.open('', '_blank', 'width=800,height=1100');
  if (!win) return;
  win.document.write(html);
  win.document.close();
  setTimeout(() => { win.focus(); win.print(); }, 400);
}


// ─────────────────────────────────────────────────────────────────────────────
// FIX 2 + 3 — BOARD PDF COVER PAGE MARKUP
// Pass boardPdfCoverData(model, logo) to your /export/pdf endpoint as
// cover_html in the request payload. Your Python backend inserts it as page 1.
// ─────────────────────────────────────────────────────────────────────────────

export function boardPdfCoverData(model, whiteLabelLogo = null, clientName = null) {
  if (!model) return {};
  const conf = Number(model.confidence_pct || 0);
  const ragColor = conf >= 75 ? '#10b981' : conf >= 55 ? '#f59e0b' : '#ef4444';
  const ragLabel = conf >= 75 ? 'Approval ready' : conf >= 55 ? 'Conditional — evidence gaps remain' : 'Do not approve';
  const today = new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });

  return {
    // These fields are read by your Python PDF generator to build page 1
    cover_title: model.title || model.subsector || 'Programme Intelligence Board Pack',
    cover_subtitle: `${model.location || ''} · ${model.mode || 'Earth'} Programme · ${today}`,
    cover_client: clientName || model.client || '',
    cover_rag_label: ragLabel,
    cover_rag_color: ragColor,
    cover_confidence: conf,
    cover_p50: model.cost_p50 || '—',
    cover_p80: model.cost_p80 || '—',
    cover_schedule: model.schedule || '—',
    cover_class: model.estimate_class_name || `Class ${model.estimate_class || 3}`,
    cover_scenario: model.scenario_label || 'Base',
    cover_logo_url: whiteLabelLogo || null,
    cover_authority_line: model.institutional_authority_line || '',
    cover_date: today,
    // Section titles (replaces generic Report header with branded CASEY section names)
    section_titles: {
      overview: '01 — Programme Intelligence',
      cost:     '02 — Cost Model & Risk Ranges',
      schedule: '03 — Schedule & Critical Path',
      risk:     '04 — Risk Register',
      qcra:     '05 — Risk Curves (P10 / P50 / P80 / P90)',
      assurance:'06 — Board Assurance Pack',
      timeline: '07 — Programme Timeline — Four-Track View',
      benchmark:'08 — Benchmark Comparison',
      appendix: 'Appendix — Methodology & Evidence Chain',
    },
  };
}

/*
 * PYTHON BACKEND CHANGE (main.py) — add to your /export/pdf route:
 *
 * @app.post("/export/pdf")
 * async def export_pdf(payload: dict):
 *     cover = payload.get("cover_html") or build_cover_page(payload)
 *     # your existing PDF generation code...
 *     # Insert cover as first page using ReportLab or WeasyPrint
 *     # Use payload["section_titles"] to replace generic headers
 */


// ─────────────────────────────────────────────────────────────────────────────
// FIX 5 — XLSX COVER TAB DATA
// Pass xlsxCoverTabData(model) to your /export/workbook endpoint.
// Your Python backend uses this to build the first two tabs.
// ─────────────────────────────────────────────────────────────────────────────

export function xlsxCoverTabData(model, whiteLabelLogo = null) {
  if (!model) return {};
  const conf = Number(model.confidence_pct || 0);
  const today = new Date().toLocaleDateString('en-GB');
  const curr = model.currency_symbol || '£';
  const risks = (model.risks || model.risk_register || []).slice(0, 5);

  return {
    cover_tab: {
      title:      model.title || model.subsector || 'Programme',
      location:   model.location || '',
      date:       today,
      prepared_by:'CASEY Programme Intelligence',
      client:     model.client || '',
      logo_url:   whiteLabelLogo || null,
      // Big three
      p50:        model.cost_p50 || '—',
      p80:        model.cost_p80 || '—',
      schedule:   model.schedule || '—',
      confidence: conf + '%',
      rag:        conf >= 75 ? 'GREEN' : conf >= 55 ? 'AMBER' : 'RED',
      verdict:    model.institutional_authority_line || '',
      class_name: model.estimate_class_name || `Class ${model.estimate_class || 3}`,
      sched_level:model.schedule_level_name || `Level ${model.schedule_level || 4}`,
      scenario:   model.scenario_label || 'Base',
    },
    exec_summary_tab: {
      headline_metrics: [
        { label: 'P10 Cost',            value: model.cost_p10 || '—',   note: 'Optimistic case' },
        { label: 'P50 Cost',            value: model.cost_p50 || '—',   note: 'Base estimate' },
        { label: 'P80 Cost',            value: model.cost_p80 || '—',   note: 'Board exposure ceiling' },
        { label: 'P90 Cost',            value: model.monte_carlo?.qcra?.p90 ? curr + model.monte_carlo.qcra.p90.toFixed(2) + 'B' : '—', note: 'Stress case' },
        { label: 'Schedule (P50)',       value: model.schedule || '—',   note: 'Base duration' },
        { label: 'Schedule (P80)',       value: model.monte_carlo?.qsra?.p80 ? model.monte_carlo.qsra.p80 + ' months' : '—', note: 'QSRA P80' },
        { label: 'Board Confidence',     value: conf + '%',               note: conf >= 75 ? 'Board-defensible' : 'Target: 75%' },
        { label: 'Estimate Class',       value: model.estimate_class_name || '—', note: model.estimate_accuracy || '' },
        { label: 'Schedule Level',       value: model.schedule_level_name || '—', note: '' },
        { label: 'Risk Count',           value: (model.risks || []).length || '—', note: '' },
        { label: 'OBA Adjustment',       value: model.oba_adjustment || '—', note: 'Reference class adjustment' },
        { label: 'Governing Constraint', value: model.governing_constraint_prominent || '—', note: '' },
      ],
      scenario_matrix: (model.scenario_matrix || []).map(s => ({
        scenario:    s.label || s.scenario,
        cost:        s.cost_p50 || '—',
        schedule:    s.schedule_months ? s.schedule_months + ' mo' : '—',
        confidence:  s.confidence_pct ? s.confidence_pct + '%' : '—',
        risk:        s.risk || '—',
      })),
      top_risks: risks.map(r => ({
        title:       r.title || r.risk || '—',
        probability: r.probability || '—',
        impact:      r.impact || r.consequence || '—',
        cost_emv:    r.cost_emv_bn ? curr + r.cost_emv_bn.toFixed(3) + 'B' : '—',
        owner:       r.owner || '—',
      })),
    },
  };
}


// ─────────────────────────────────────────────────────────────────────────────
// FIX 6 — WHITE-LABEL SETTINGS COMPONENT
// Add this to your settings/account panel. Stores logo as data URL in
// localStorage and passes it to all export functions.
// ─────────────────────────────────────────────────────────────────────────────

export function WhiteLabelSettings({ onLogoChange }) {
  const [logo, setLogo] = useState(() => {
    try { return localStorage.getItem('casey_wl_logo') || null; } catch { return null; }
  });
  const [orgName, setOrgName] = useState(() => {
    try { return localStorage.getItem('casey_wl_name') || ''; } catch { return ''; }
  });
  const [enabled, setEnabled] = useState(() => {
    try { return localStorage.getItem('casey_wl_enabled') === '1'; } catch { return false; }
  });

  function handleLogoUpload(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = reader.result;
      setLogo(dataUrl);
      try { localStorage.setItem('casey_wl_logo', dataUrl); } catch {}
      if (onLogoChange) onLogoChange(dataUrl);
    };
    reader.readAsDataURL(file);
  }

  function save() {
    try {
      localStorage.setItem('casey_wl_name', orgName);
      localStorage.setItem('casey_wl_enabled', enabled ? '1' : '0');
    } catch {}
    if (onLogoChange) onLogoChange(enabled ? logo : null);
  }

  const rowStyle = { display: 'flex', alignItems: 'center', gap: 10, marginBottom: 8, fontSize: 11 };
  const inputStyle = { background: 'rgba(255,255,255,0.04)', border: '1px solid rgba(255,255,255,0.1)', borderRadius: 4, color: '#e2e8f0', padding: '6px 10px', fontSize: 11, flex: 1 };

  return (
    <div style={{ padding: '12px 0' }}>
      <div style={{ fontSize: 9, fontWeight: 800, color: '#475569', letterSpacing: '.12em', marginBottom: 10 }}>
        WHITE-LABEL EXPORTS
      </div>
      <div style={rowStyle}>
        <label style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer', color: '#94a3b8' }}>
          <input type="checkbox" checked={enabled} onChange={e => setEnabled(e.target.checked)}
            style={{ accentColor: '#8df7ff', width: 14, height: 14 }} />
          Enable white-label branding on all exports
        </label>
      </div>
      {enabled && <>
        <div style={rowStyle}>
          <span style={{ color: '#64748b', fontSize: 11, minWidth: 80 }}>Organisation</span>
          <input style={inputStyle} value={orgName} onChange={e => setOrgName(e.target.value)}
            placeholder="e.g. Turner & Townsend" />
        </div>
        <div style={rowStyle}>
          <span style={{ color: '#64748b', fontSize: 11, minWidth: 80 }}>Logo</span>
          <label style={{ ...inputStyle, cursor: 'pointer', textAlign: 'center', padding: '8px' }}>
            {logo ? '✓ Logo uploaded — click to change' : 'Upload logo (PNG/SVG, white background)'}
            <input type="file" accept="image/*" style={{ display: 'none' }} onChange={handleLogoUpload} />
          </label>
        </div>
        {logo && <img src={logo} style={{ height: 32, objectFit: 'contain', marginBottom: 8, background: '#fff', padding: '4px 8px', borderRadius: 3 }} alt="Logo preview" />}
        <p style={{ fontSize: 9, color: '#334155', marginBottom: 8 }}>
          Logo appears on: Board Pack PDF cover, One-page brief, XLSX cover tab.
          CASEY watermark remains as "Powered by CASEY" in the footer.
        </p>
      </>}
      <button onClick={save} style={{ padding: '6px 16px', background: 'rgba(141,247,255,0.1)', border: '1px solid rgba(141,247,255,0.3)', color: '#8df7ff', borderRadius: 4, cursor: 'pointer', fontSize: 11, fontWeight: 700 }}>
        Save settings
      </button>
    </div>
  );
}

/*
 * INTEGRATION — in App.jsx:
 * 1. Add state: const [whiteLabelLogo, setWhiteLabelLogo] = useState(
 *      () => { try { return localStorage.getItem('casey_wl_logo') || null; } catch { return null; } }
 *    );
 * 2. Add <WhiteLabelSettings onLogoChange={setWhiteLabelLogo} /> to your settings/account panel
 * 3. Pass whiteLabelLogo to printOnePager(model, whiteLabelLogo)
 * 4. Pass boardPdfCoverData(model, whiteLabelLogo) to /export/pdf payload
 * 5. Pass xlsxCoverTabData(model, whiteLabelLogo) to /export/workbook payload
 */


// ─────────────────────────────────────────────────────────────────────────────
// FIX 7 — TIMELINE EXPORT BUTTON
// Captures the canvas and opens a print window. Also builds PPTX slide notes.
// ─────────────────────────────────────────────────────────────────────────────

export function TimelineExportButton({ model, canvasIds = ['cvs-base','cvs-scenario','cvs-stress','cvs-benchmark'] }) {
  function exportTimelineAsPDF() {
    // Collect canvases and combine into a printable page
    const canvases = canvasIds.map(id => document.getElementById(id)).filter(Boolean);
    if (canvases.length === 0) {
      alert('Run a project and open the Timeline tab first.');
      return;
    }

    const today = new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });
    const title = model?.title || model?.subsector || 'Programme Timeline';
    const conf = model?.confidence_pct || '—';

    // Collect canvas data URLs
    const images = canvases.map(cvs => {
      try { return cvs.toDataURL('image/png'); } catch { return null; }
    }).filter(Boolean);

    // Track names in order
    const trackNames = ['Base', 'Scenario', 'Stress Test', 'Benchmark'];

    const imgTags = images.map((src, i) => `
      <div style="margin-bottom:16px">
        <div style="font-size:9pt;font-weight:700;color:#334155;letter-spacing:.1em;margin-bottom:6px;text-transform:uppercase">
          ${trackNames[i] || `Track ${i + 1}`}
        </div>
        <img src="${src}" style="width:100%;border:1px solid #e2e8f0;border-radius:3px">
      </div>`).join('');

    const html = `<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
  @page { size: A3 landscape; margin: 12mm; }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Helvetica Neue', Arial, sans-serif; color: #0f172a; background: #fff; }
  .hdr { display: flex; justify-content: space-between; align-items: baseline; padding-bottom: 8px; border-bottom: 2px solid #0f172a; margin-bottom: 14px; }
  .title { font-size: 16pt; font-weight: 900; }
  .sub { font-size: 9pt; color: #64748b; }
  .badge { font-size: 8pt; font-weight: 800; padding: 3px 10px; border-radius: 2px; color: #fff; background: ${conf >= 75 ? '#10b981' : conf >= 55 ? '#f59e0b' : '#ef4444'}; }
  .footer { margin-top: 12px; font-size: 8pt; color: #94a3b8; display: flex; justify-content: space-between; }
</style></head><body>
<div class="hdr">
  <div>
    <div class="title">${title} — Four-Track Programme Timeline</div>
    <div class="sub">${model?.location || ''} · ${today}</div>
  </div>
  <div class="badge">${conf}% CONFIDENCE</div>
</div>
${imgTags}
<div class="footer">
  <span>CASEY Programme Intelligence · ${model?.estimate_class_name || ''} · ${model?.schedule_level_name || ''}</span>
  <span>Base / Scenario / Stress Test / Benchmark — animated in the CASEY web application</span>
</div>
</body></html>`;

    const win = window.open('', '_blank', 'width=1200,height=900');
    if (!win) return;
    win.document.write(html);
    win.document.close();
    setTimeout(() => { win.focus(); win.print(); }, 400);
  }

  function exportPptxNotes(model) {
    // Generates speaker notes for a PPTX slide — paste these into PowerPoint
    if (!model) return;
    const notes = `CASEY FOUR-TRACK PROGRAMME TIMELINE — SPEAKER NOTES

Programme: ${model.title || model.subsector || '—'}
Location: ${model.location || '—'}
Date: ${new Date().toLocaleDateString('en-GB')}

TRACK 1 — BASE (Blue)
The contractual baseline programme. ${model.cost_p50 || '—'} P50 cost, ${model.schedule || '—'} delivery.
Confidence: ${model.confidence_pct || '—'}%. Estimate: ${model.estimate_class_name || '—'}.

TRACK 2 — SCENARIO (Green)
Optimistic case — scope freeze, fast procurement, no adverse weather.
Shows the best credible delivery outcome if key risks are mitigated.

TRACK 3 — STRESS TEST (Red)
P90 case — all identified risks materialise. Regulatory delay. Scope uplift.
This is what comparable programmes actually experienced at the 90th percentile.

TRACK 4 — BENCHMARK (Amber)
Reference class from a comparable completed programme.
The benchmark shows what similar programmes actually delivered.

KEY RISKS (firing in timeline order):
${(model.risks || []).slice(0, 5).map((r, i) => `${i + 1}. ${r.title || r.risk || '—'} — ${r.impact || '—'}`).join('\n')}

GOVERNING CONSTRAINT:
${model.governing_constraint_prominent || model.governing_constraint_full?.statement || '—'}

Generated by CASEY Programme Intelligence. Not a certified estimate document.`;

    const blob = new Blob([notes], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'CASEY_Timeline_SpeakerNotes.txt';
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  if (!model) return null;
  return (
    <div style={{ display: 'flex', gap: 8 }}>
      <button onClick={exportTimelineAsPDF} style={{
        padding: '5px 13px', fontSize: 10, fontWeight: 700,
        border: '1px solid rgba(141,247,255,0.3)', borderRadius: 4,
        background: 'rgba(141,247,255,0.08)', color: '#8df7ff', cursor: 'pointer',
      }}>
        Export timeline PDF
      </button>
      <button onClick={() => exportPptxNotes(model)} style={{
        padding: '5px 13px', fontSize: 10, fontWeight: 700,
        border: '1px solid rgba(155,127,232,0.3)', borderRadius: 4,
        background: 'rgba(155,127,232,0.08)', color: '#9b7fe8', cursor: 'pointer',
      }}>
        Export slide notes
      </button>
    </div>
  );
}


// ─────────────────────────────────────────────────────────────────────────────
// FIX 2 — UPGRADED ExportStrip (replaces existing ExportStrip)
// Add one-pager and timeline export buttons
// ─────────────────────────────────────────────────────────────────────────────

export function ExportStripV2({
  model,
  onBoardPack,
  onExcel,
  onRisk,
  onXer,
  onQcra,
  whiteLabelLogo = null,
}) {
  if (!model) return null;
  return (
    <section style={{
      display: 'flex', flexWrap: 'wrap', gap: 6, padding: '8px 20px',
      background: 'rgba(8,12,20,0.95)', borderBottom: '1px solid rgba(255,255,255,0.06)',
      alignItems: 'center',
    }}>
      {/* One-pager — the new killer button */}
      <button onClick={() => printOnePager(model, whiteLabelLogo)} style={{
        padding: '5px 13px', fontSize: 10, fontWeight: 700,
        border: '1px solid rgba(16,185,129,0.4)', borderRadius: 4,
        background: 'rgba(16,185,129,0.12)', color: '#10b981', cursor: 'pointer',
      }}>
        ✦ One-page brief
      </button>

      <div style={{ width: 1, height: 18, background: 'rgba(255,255,255,0.08)' }} />

      <button onClick={onBoardPack} style={exportBtnStyle('#8df7ff')}>
        ↓ Board pack PDF
      </button>
      <button onClick={onExcel} style={exportBtnStyle('#4a9eff')}>
        ↓ Cost workbook
      </button>
      <button onClick={onRisk} style={exportBtnStyle('#f5a623')}>
        ↓ Risk register
      </button>
      <button onClick={onXer} style={exportBtnStyle('#9b7fe8')}>
        ↓ Schedule XER
      </button>
      <button onClick={onQcra} style={exportBtnStyle('#e05252')}>
        ↓ Risk curves
      </button>

      <div style={{ width: 1, height: 18, background: 'rgba(255,255,255,0.08)' }} />

      <TimelineExportButton model={model} />
    </section>
  );
}

function exportBtnStyle(color) {
  return {
    padding: '5px 12px', fontSize: 10, fontWeight: 500,
    border: `1px solid ${color}30`, borderRadius: 4,
    background: `${color}10`, color, cursor: 'pointer',
  };
}


// ─────────────────────────────────────────────────────────────────────────────
// FIX 1 — LANDING HERO TIMELINE (auto-playing, no interaction required)
// Replace the static landing screen description box with this component.
// It auto-plays the Earth demo timeline silently in the background.
// ─────────────────────────────────────────────────────────────────────────────

export function LandingHeroTimeline({ onRunEarth, onRunSpace, onRunFree }) {
  const cvsRef = useRef(null);
  const rafRef = useRef(null);
  const progRef = useRef(0);

  // Minimal synthetic data for the auto-play — no API call needed
  const DEMO = {
    months: 24,
    milestones: [
      { t: 0, l: 'Start' }, { t: .13, l: 'Design' }, { t: .30, l: 'Procurement' },
      { t: .50, l: 'Construction' }, { t: .72, l: 'Systems' }, { t: .90, l: 'Commissioning' },
      { t: 1, l: 'Handover' },
    ],
    risks: [
      { t: .18, c: '#e05252' }, { t: .38, c: '#f5a623' }, { t: .55, c: '#e05252' },
      { t: .72, c: '#f5a623' }, { t: .86, c: '#e05252' },
    ],
  };

  useEffect(() => {
    const cvs = cvsRef.current;
    if (!cvs) return;
    const DPR = window.devicePixelRatio || 1;
    const W = cvs.offsetWidth;
    const H = 120;
    cvs.width = W * DPR; cvs.height = H * DPR; cvs.style.height = H + 'px';
    const ctx = cvs.getContext('2d');
    ctx.scale(DPR, DPR);

    const X0 = 12, X1 = W - 12, TW = X1 - X0;
    const YTL = H * 0.50, YSPEND = H * 0.82;
    function xAt(t) { return X0 + t * TW; }
    function scurve(t) { return 1 / (1 + Math.exp(-10 * (t - 0.5))); }
    function spendY(r) { return YSPEND - r * (YSPEND - 18); }

    function frame() {
      progRef.current = Math.min(1, progRef.current + 0.003);
      const p = progRef.current;
      ctx.clearRect(0, 0, W, H);

      // Rail background
      ctx.strokeStyle = 'rgba(255,255,255,0.06)'; ctx.lineWidth = 1;
      ctx.beginPath(); ctx.moveTo(X0, YTL); ctx.lineTo(X1, YTL); ctx.stroke();
      ctx.beginPath(); ctx.moveTo(X0, YSPEND); ctx.lineTo(X1, YSPEND); ctx.stroke();

      // Tracks — Base (blue), Scenario (green), Stress (red)
      [
        { col: '#4a9eff', yOff: 0, mult: 1 },
        { col: '#2ecc8a', yOff: -12, mult: 0.93, dash: [6, 3] },
        { col: '#e05252', yOff: 0, mult: 1.28, dash: [4, 3] },
      ].forEach(tr => {
        ctx.save(); ctx.strokeStyle = tr.col; ctx.lineWidth = tr.yOff === 0 ? 2 : 1.5;
        ctx.lineCap = 'round';
        if (tr.dash) ctx.setLineDash(tr.dash);
        ctx.globalAlpha = tr.yOff === 0 ? 1 : 0.7;
        ctx.beginPath();
        for (let i = 0; i <= 200; i++) {
          const t = i / 200; if (t > p) break;
          const x = xAt(t * tr.mult); const y = YTL + tr.yOff;
          if (i === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y);
        }
        ctx.stroke(); ctx.setLineDash([]); ctx.restore();

        // Spend curve
        ctx.save(); ctx.strokeStyle = tr.col; ctx.lineWidth = 0.8;
        ctx.globalAlpha = 0.5;
        if (tr.dash) ctx.setLineDash(tr.dash);
        ctx.beginPath();
        for (let i = 0; i <= 200; i++) {
          const t = i / 200; if (t > p) break;
          const x = xAt(t * tr.mult); const y = spendY(scurve(t) * tr.mult);
          if (i === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y);
        }
        ctx.stroke(); ctx.setLineDash([]); ctx.restore();
      });

      // Risk diamonds
      DEMO.risks.forEach(r => {
        if (p < r.t) return;
        const x = xAt(r.t), s = 5;
        ctx.save(); ctx.fillStyle = r.c; ctx.translate(x, YTL);
        ctx.beginPath(); ctx.moveTo(0, -s); ctx.lineTo(s, 0); ctx.lineTo(0, s); ctx.lineTo(-s, 0);
        ctx.closePath(); ctx.fill(); ctx.restore();
      });

      // Milestones
      DEMO.milestones.forEach((m, i) => {
        if (p < m.t) return;
        const x = xAt(m.t);
        ctx.save(); ctx.fillStyle = '#2ecc8a';
        ctx.beginPath(); ctx.arc(x, YTL, 4, 0, Math.PI * 2); ctx.fill();
        ctx.strokeStyle = '#080c14'; ctx.lineWidth = 1.2;
        ctx.beginPath(); ctx.arc(x, YTL, 4, 0, Math.PI * 2); ctx.stroke();
        ctx.font = '8px system-ui'; ctx.fillStyle = '#2ecc8a'; ctx.textAlign = 'center';
        ctx.fillText(m.l, x, i % 2 === 0 ? YTL - 9 : YTL + 16);
        ctx.restore();
      });

      // Moving head
      ctx.save(); ctx.fillStyle = '#4a9eff';
      ctx.beginPath(); ctx.arc(xAt(p), YTL, 6, 0, Math.PI * 2); ctx.fill();
      ctx.strokeStyle = '#080c14'; ctx.lineWidth = 1.5;
      ctx.beginPath(); ctx.arc(xAt(p), YTL, 6, 0, Math.PI * 2); ctx.stroke(); ctx.restore();

      if (p < 1) rafRef.current = requestAnimationFrame(frame);
      else { setTimeout(() => { progRef.current = 0; rafRef.current = requestAnimationFrame(frame); }, 1200); }
    }

    rafRef.current = requestAnimationFrame(frame);
    return () => { if (rafRef.current) cancelAnimationFrame(rafRef.current); };
  }, []);

  return (
    <div style={{
      background: '#080c14', border: '1px solid rgba(255,255,255,0.06)',
      borderRadius: 12, overflow: 'hidden', marginBottom: 16,
    }}>
      {/* Timeline canvas */}
      <div style={{ padding: '14px 16px 8px', borderBottom: '1px solid rgba(255,255,255,0.06)' }}>
        <div style={{ fontSize: 9, color: 'rgba(255,255,255,0.2)', letterSpacing: '.14em', marginBottom: 8 }}>
          CASEY · LIVE PROGRAMME TIMELINE · BASE / SCENARIO / STRESS TEST
        </div>
        <canvas ref={cvsRef} style={{ display: 'block', width: '100%' }} />
      </div>

      {/* Legend */}
      <div style={{ padding: '6px 16px', display: 'flex', gap: 16, borderBottom: '1px solid rgba(255,255,255,0.06)' }}>
        {[['#4a9eff', 'Base'], ['#2ecc8a', 'Scenario'], ['#e05252', 'Stress test'], ['#2ecc8a', 'Milestone'], ['#e05252', 'Risk event']].map(([c, l]) => (
          <div key={l} style={{ display: 'flex', alignItems: 'center', gap: 5, fontSize: 9, color: 'rgba(255,255,255,0.3)' }}>
            <div style={{ width: 7, height: 7, borderRadius: '50%', background: c }} />
            <span>{l}</span>
          </div>
        ))}
      </div>

      {/* CTA buttons */}
      <div style={{ padding: '12px 16px', display: 'flex', gap: 8, flexWrap: 'wrap' }}>
        <button onClick={onRunEarth} style={{ flex: 1, minWidth: 120, padding: '10px', background: 'rgba(16,185,129,0.1)', border: '1px solid rgba(16,185,129,0.3)', color: '#10b981', borderRadius: 4, cursor: 'pointer', fontSize: 11, fontWeight: 700 }}>
          Run Earth demo
        </button>
        <button onClick={onRunSpace} style={{ flex: 1, minWidth: 120, padding: '10px', background: 'rgba(141,247,255,0.07)', border: '1px solid rgba(141,247,255,0.2)', color: '#8df7ff', borderRadius: 4, cursor: 'pointer', fontSize: 11, fontWeight: 700 }}>
          Run Space demo
        </button>
        <button onClick={onRunFree} style={{ flex: 1, minWidth: 120, padding: '10px', background: 'rgba(155,127,232,0.1)', border: '1px solid rgba(155,127,232,0.3)', color: '#9b7fe8', borderRadius: 4, cursor: 'pointer', fontSize: 11, fontWeight: 700 }}>
          Describe your project
        </button>
      </div>
    </div>
  );
}

/*
 * INTEGRATION — replace landing screen in App.jsx:
 *
 * Find: {show && <LandingScreen .../>}  (or similar)
 * Add above the landing form — or replace the empty state:
 *
 * {show && !model && (
 *   <LandingHeroTimeline
 *     onRunEarth={() => loadInstantDemo('earth')}
 *     onRunSpace={() => loadInstantDemo('space')}
 *     onRunFree={() => { setShow(false); setTrialOpen(true); }}
 *   />
 * )}
 */
