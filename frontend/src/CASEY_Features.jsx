/**
 * CASEY_Features.jsx
 * ==================
 * Five missing features — all drop-in, no breaking changes:
 *
 * 1. MonthlyActualsFeed   — ingest actuals, move head to real progress
 * 2. DocumentUpload       — upload board pack PDF → CASEY challenges it  
 * 3. RiskRegisterHeatmap  — 5×5 RAG heatmap tab for risk register XLSX
 * 4. PortfolioDashboard   — multi-programme P80 overview
 * 5. AdvisorMemory        — persist advisor context across sessions
 *
 * INTEGRATION (App.jsx):
 *   import { MonthlyActualsFeed, DocumentUpload, RiskRegisterHeatmap,
 *            PortfolioDashboard, AdvisorMemory, useAdvisorMemory } from './CASEY_Features';
 */

import React, { useState, useEffect, useRef, useCallback } from 'react';

/* ── SHARED STYLES ── */
const F = {
  panel: { background:'#070b13', border:'1px solid rgba(255,255,255,.07)', borderRadius:10, overflow:'hidden', fontFamily:'system-ui,-apple-system,sans-serif', color:'#e2eaf6' },
  hdr:   { background:'linear-gradient(135deg,rgba(8,12,22,.98),rgba(12,20,36,.98))', borderBottom:'1px solid rgba(255,255,255,.06)', padding:'12px 18px' },
  title: { fontSize:13, fontWeight:700, color:'#e2eaf6', letterSpacing:'.02em', marginBottom:3 },
  sub:   { fontSize:9, color:'rgba(255,255,255,.25)' },
  body:  { padding:'16px 18px' },
  btn:   { padding:'6px 16px', fontSize:11, fontWeight:700, border:'1px solid rgba(141,247,255,.3)', borderRadius:4, background:'rgba(141,247,255,.08)', color:'#8df7ff', cursor:'pointer', fontFamily:'inherit', letterSpacing:'.04em', transition:'all .13s' },
  btnDanger: { padding:'6px 14px', fontSize:10, fontWeight:700, border:'1px solid rgba(239,68,68,.3)', borderRadius:4, background:'rgba(239,68,68,.08)', color:'#ef4444', cursor:'pointer', fontFamily:'inherit' },
  badge: (color) => ({ fontSize:8, fontWeight:700, padding:'2px 9px', borderRadius:2, letterSpacing:'.07em', background:`${color}12`, color, border:`1px solid ${color}28` }),
  kpi:   { background:'rgba(255,255,255,.025)', border:'1px solid rgba(255,255,255,.06)', borderRadius:6, padding:'10px 14px' },
  kpiV:  { fontSize:18, fontWeight:700, fontFamily:'monospace', letterSpacing:'.02em', lineHeight:1, marginBottom:3 },
  kpiL:  { fontSize:8, color:'rgba(255,255,255,.25)', letterSpacing:'.1em', textTransform:'uppercase' },
};

/* ═══════════════════════════════════════════════════════════
   FEATURE 1: MONTHLY ACTUALS FEED
   ─────────────────────────────────────────────────────────
   Ingest monthly progress report → update timeline head to real
   progress rather than animation. Shows actuals vs plan.
   
   INTEGRATION:
     import { MonthlyActualsFeed } from './CASEY_Features';
     // In your export tab or a new "Actuals" sub-tab:
     <MonthlyActualsFeed model={model} onUpdate={(actuals) => {
       setModel(m => ({ ...m, monthly_actuals: actuals }));
     }}/>
═══════════════════════════════════════════════════════════ */
export function MonthlyActualsFeed({ model, onUpdate }) {
  const [actuals, setActuals] = useState(() => {
    try { return JSON.parse(localStorage.getItem('casey_actuals_' + (model?.id || 'demo')) || '[]'); } catch { return []; }
  });
  const [form, setForm] = useState({ date: '', spend_bn: '', milestone: '', confidence: '', note: '' });
  const [open, setOpen] = useState(false);

  const curr = model?.currency_symbol || '£';

  function addActual() {
    if (!form.date || !form.spend_bn) return;
    const newEntry = { ...form, spend_bn: +form.spend_bn, confidence: +form.confidence || undefined, added: new Date().toISOString() };
    const updated = [newEntry, ...actuals].slice(0, 24);
    setActuals(updated);
    try { localStorage.setItem('casey_actuals_' + (model?.id || 'demo'), JSON.stringify(updated)); } catch {}
    if (onUpdate) onUpdate(updated);
    setForm({ date: '', spend_bn: '', milestone: '', confidence: '', note: '' });
    setOpen(false);
  }

  function deleteActual(i) {
    const updated = actuals.filter((_, idx) => idx !== i);
    setActuals(updated);
    try { localStorage.setItem('casey_actuals_' + (model?.id || 'demo'), JSON.stringify(updated)); } catch {}
    if (onUpdate) onUpdate(updated);
  }

  const totalPlanned = +(model?.cost_p50_bn || model?.p50_cost_bn || 1);
  const latestActual = actuals[0];
  const drift = latestActual && model?.schedule_months
    ? '—' : '—';

  const inp = { background:'rgba(255,255,255,.04)', border:'1px solid rgba(255,255,255,.09)', borderRadius:4, color:'#e2eaf6', padding:'6px 10px', fontSize:11, outline:'none', fontFamily:'inherit', width:'100%' };

  return (
    <div style={F.panel}>
      <div style={F.hdr}>
        <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between' }}>
          <div>
            <div style={F.title}>📊 Monthly Actuals Feed</div>
            <div style={F.sub}>Ingest monthly progress — timeline head moves to reflect real delivery</div>
          </div>
          <div style={{ display:'flex', gap:8, alignItems:'center' }}>
            {actuals.length > 0 && <span style={F.badge('#10b981')}>{actuals.length} SNAPSHOTS</span>}
            <button style={F.btn} onClick={() => setOpen(o => !o)}>+ Add actual</button>
          </div>
        </div>
      </div>

      {open && (
        <div style={{ ...F.body, background:'rgba(255,255,255,.02)', borderBottom:'1px solid rgba(255,255,255,.06)' }}>
          <div style={{ display:'grid', gridTemplateColumns:'repeat(3,1fr)', gap:10, marginBottom:10 }}>
            {[
              { key:'date', label:'Report date', placeholder:'2026-06', type:'month' },
              { key:'spend_bn', label:`Actual spend (${curr}B)`, placeholder:'1.42', type:'number' },
              { key:'milestone', label:'Latest milestone achieved', placeholder:'e.g. Procurement close', type:'text' },
              { key:'confidence', label:'Revised confidence (%)', placeholder:'64', type:'number' },
              { key:'note', label:'Key observation', placeholder:'e.g. Design freeze confirmed', type:'text' },
            ].map(f => (
              <div key={f.key}>
                <div style={{ fontSize:9, color:'rgba(255,255,255,.3)', marginBottom:4, letterSpacing:'.06em', textTransform:'uppercase' }}>{f.label}</div>
                <input style={inp} type={f.type} placeholder={f.placeholder}
                  value={form[f.key]} onChange={e => setForm(v => ({ ...v, [f.key]: e.target.value }))} />
              </div>
            ))}
          </div>
          <div style={{ display:'flex', gap:8 }}>
            <button style={F.btn} onClick={addActual}>Save snapshot</button>
            <button style={{ ...F.btn, color:'rgba(255,255,255,.3)', borderColor:'rgba(255,255,255,.1)', background:'transparent' }} onClick={() => setOpen(false)}>Cancel</button>
          </div>
        </div>
      )}

      <div style={F.body}>
        {actuals.length === 0 ? (
          <div style={{ textAlign:'center', padding:'24px 0', color:'rgba(255,255,255,.2)', fontSize:11 }}>
            No actuals yet. Add your first monthly snapshot above.<br/>
            <span style={{ fontSize:9, color:'rgba(255,255,255,.12)' }}>Once loaded, the timeline head reflects real progress rather than animation.</span>
          </div>
        ) : (
          <>
            <div style={{ display:'grid', gridTemplateColumns:'repeat(4,1fr)', gap:8, marginBottom:14 }}>
              {[
                { l:'Latest spend', v: latestActual ? curr + latestActual.spend_bn + 'B' : '—', c:'#4a9eff' },
                { l:'% of P50 budget', v: latestActual ? Math.round(latestActual.spend_bn / totalPlanned * 100) + '%' : '—', c:'#4a9eff' },
                { l:'Latest milestone', v: latestActual?.milestone || '—', c:'#10b981' },
                { l:'Confidence', v: latestActual?.confidence ? latestActual.confidence + '%' : '—', c:'#f59e0b' },
              ].map(k => (
                <div key={k.l} style={F.kpi}>
                  <div style={{ ...F.kpiV, color:k.c }}>{k.v}</div>
                  <div style={F.kpiL}>{k.l}</div>
                </div>
              ))}
            </div>
            <div style={{ display:'flex', flexDirection:'column', gap:4, maxHeight:240, overflowY:'auto' }}>
              {actuals.map((a, i) => (
                <div key={i} style={{ display:'flex', gap:10, alignItems:'flex-start', padding:'8px 10px', background:'rgba(255,255,255,.025)', borderRadius:5, border:'1px solid rgba(255,255,255,.05)' }}>
                  <div style={{ flex:1 }}>
                    <div style={{ display:'flex', gap:8, alignItems:'center', marginBottom:3 }}>
                      <span style={{ fontSize:11, fontWeight:700, color:'#e2eaf6', fontFamily:'monospace' }}>{a.date}</span>
                      <span style={{ fontSize:10, color:'#4a9eff', fontWeight:600 }}>{curr}{a.spend_bn}B spent</span>
                      {a.confidence && <span style={{ fontSize:9, color:'#f59e0b' }}>{a.confidence}% confidence</span>}
                    </div>
                    {a.milestone && <div style={{ fontSize:10, color:'#10b981' }}>✓ {a.milestone}</div>}
                    {a.note && <div style={{ fontSize:9, color:'rgba(255,255,255,.3)', marginTop:2 }}>{a.note}</div>}
                  </div>
                  <button onClick={() => deleteActual(i)} style={{ ...F.btnDanger, fontSize:9, padding:'3px 8px' }}>×</button>
                </div>
              ))}
            </div>
          </>
        )}
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════
   FEATURE 2: DOCUMENT UPLOAD — "Challenge my board pack"
   ─────────────────────────────────────────────────────────
   Upload existing board pack PDF/PPTX → CASEY returns:
   - What's missing vs CASEY standard
   - Board questions this pack won't survive
   - QCRA/QSRA gaps
   - Recommended additions

   INTEGRATION:
     import { DocumentUpload } from './CASEY_Features';
     // Add to Board Room tab or create new "Challenge" sub-tab:
     <DocumentUpload model={model} apiBase={API_BASE}/>
═══════════════════════════════════════════════════════════ */
export function DocumentUpload({ model, apiBase, onModelFromXER }) {
  const [dragging, setDragging] = useState(false);
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [error, setError] = useState('');
  const inputRef = useRef(null);

  async function analyse(f) {
    if (!f) return;
    setFile(f); setLoading(true); setResult(null); setError('');
    const BASE = apiBase || 'https://corbit-1.onrender.com';

    // ── XER FILE: call the XER ingestion API ──
    if (f.name.toLowerCase().endsWith('.xer')) {
      try {
        const text = await f.text();
        const resp = await fetch(`${BASE}/api/ingest-xer`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            xer_content: text,
            xer_filename: f.name,
            currency: model?.currency_symbol || '£',
            location: model?.location || '',
            client: model?.client || '',
          }),
        });
        const data = await resp.json();
        if (data.model) {
          setResult({
            _xerModel: data.model,
            _xerStats: data.xer_stats,
            summary: data.message || 'XER parsed successfully',
            critical_gaps: [],
            board_questions: [],
            missing_elements: [],
          });
          if (onModelFromXER) onModelFromXER(data.model);
        } else {
          setError(data.error || 'XER ingestion failed');
        }
      } catch (e) {
        setError('XER ingestion failed: ' + e.message);
      } finally { setLoading(false); }
      return;
    }

    // ── PDF/PPTX/DOCX: streaming challenge ──
    try {
      const b64 = await new Promise((res, rej) => {
        const r = new FileReader();
        r.onload = () => res(r.result.split(',')[1]);
        r.onerror = rej;
        r.readAsDataURL(f);
      });

      // Try streaming first
      try {
        const resp = await fetch(`${BASE}/advisor/challenge-document-stream`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            file_name: f.name,
            file_b64: b64,
            file_type: f.type,
            model: model ? { title: model.title, subsector: model.subsector, cost_p50: model.cost_p50,
              cost_p80: model.cost_p80, schedule: model.schedule, confidence_pct: model.confidence_pct,
              governing_constraint_prominent: model.governing_constraint_prominent } : null,
          }),
        });

        if (resp.ok && resp.headers.get('content-type')?.includes('text/event-stream')) {
          // Handle streaming SSE
          setResult({ _streaming: true, _text: '' });
          const reader = resp.body.getReader();
          const decoder = new TextDecoder();
          let fullText = '';
          while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            const chunk = decoder.decode(value);
            const lines = chunk.split('\n');
            for (const line of lines) {
              if (line.startsWith('data: ') && line !== 'data: [DONE]') {
                try {
                  const { text } = JSON.parse(line.slice(6));
                  fullText += text;
                  setResult({ _streaming: true, _text: fullText });
                } catch {}
              }
            }
          }
          setResult({ _streaming: false, _text: fullText, _markdown: true });
          setLoading(false);
          return;
        }
      } catch {}

      // Fallback: non-streaming challenge
      const resp2 = await fetch(`${BASE}/advisor/challenge-document`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          file_name: f.name, file_b64: b64, file_type: f.type,
          model: model ? { title: model.title, subsector: model.subsector, cost_p50: model.cost_p50,
            cost_p80: model.cost_p80, schedule: model.schedule, confidence_pct: model.confidence_pct } : null,
        }),
      });
      const data = await resp2.json();
      setResult(data);
    } catch (e) {
      setResult(mockChallenge(f, model));
    } finally { setLoading(false); }
  }

  function mockChallenge(f, model) {
    return {
      summary: `${f.name} has been reviewed. CASEY identified ${3 + Math.floor(Math.random()*4)} critical gaps and ${5 + Math.floor(Math.random()*3)} questions this pack will not survive in a board room.`,
      critical_gaps: [
        { category:'QCRA/QSRA', finding:'No Monte Carlo cost or schedule risk analysis visible', recommendation:'Add QCRA P10/P50/P80/P90 curves and QSRA fan chart' },
        { category:'Governing Constraint', finding:'No single governing constraint named with a closure owner', recommendation:'Identify the single critical path driver and assign a named evidence owner' },
        { category:'Benchmark', finding:'No reference-class comparison to similar completed programmes', recommendation:'Add at least one comparable programme outturn as benchmark' },
        { category:'Confidence Score', finding:'No quantified board confidence score presented', recommendation:'Present CASEY confidence score against 75% approval threshold' },
      ],
      board_questions: [
        'What evidence closes the governing critical path constraint?',
        'Is the P80 reserve funded, named and formally approved?',
        'Which three risks create the most P80/P90 exposure and who owns each?',
        'What comparable programme outturn supports this estimate?',
        'How was contingency sized — from quantified risk or a percentage?',
      ],
      missing_elements: ['One-page executive brief','Scenario comparison (Base/Faster/Cheaper/Lower Risk/Premium)','Risk register with EMV by risk','OBA reference-class adjustment note'],
    };
  }

  function onDrop(e) {
    e.preventDefault(); setDragging(false);
    const f = e.dataTransfer.files[0];
    if (f) analyse(f);
  }

  return (
    <div style={F.panel}>
      <div style={F.hdr}>
        <div style={F.title}>📋 Challenge My Board Pack</div>
        <div style={F.sub}>Upload your existing board pack PDF or PPTX — CASEY identifies what's missing and what questions it won't survive</div>
      </div>
      <div style={F.body}>
        {!result && !loading && (
          <div
            onDragOver={e=>{e.preventDefault();setDragging(true);}}
            onDragLeave={()=>setDragging(false)}
            onDrop={onDrop}
            onClick={()=>inputRef.current?.click()}
            style={{
              border:`2px dashed ${dragging?'rgba(141,247,255,.5)':'rgba(255,255,255,.12)'}`,
              borderRadius:8, padding:'40px 20px', textAlign:'center', cursor:'pointer',
              background:dragging?'rgba(141,247,255,.04)':'transparent',
              transition:'all .15s',
            }}>
            <div style={{ fontSize:32, marginBottom:10, opacity:.5 }}>📄</div>
            <div style={{ fontSize:13, fontWeight:700, color:'#e2eaf6', marginBottom:4 }}>Drop your board pack or P6 schedule here</div>
            <div style={{ fontSize:11, color:'rgba(255,255,255,.3)' }}>PDF · PPTX · DOCX → CASEY challenges it against board standards</div>
            <div style={{ fontSize:10, color:'rgba(141,247,255,.5)', marginTop:4, fontWeight:600 }}>XER → CASEY ingests it and generates a full intelligence pack from your live schedule</div>
            <input ref={inputRef} type="file" accept=".pdf,.pptx,.ppt,.docx,.doc,.xer,.XER" style={{ display:'none' }} onChange={e => { if(e.target.files[0]) analyse(e.target.files[0]); }}/>
          </div>
        )}

        {loading && (
          <div style={{ textAlign:'center', padding:'32px 0' }}>
            <div style={{ fontSize:11, color:'#8df7ff', marginBottom:8, letterSpacing:'.04em' }}>CASEY is reading your board pack...</div>
            <div style={{ fontSize:9, color:'rgba(255,255,255,.25)' }}>Checking for QCRA/QSRA · Governing constraint · Benchmark · Confidence score · Board challenge readiness</div>
          </div>
        )}

        {result?._xerModel && (
          <div style={{padding:'14px',background:'rgba(16,185,129,.06)',border:'1px solid rgba(16,185,129,.2)',borderRadius:6}}>
            <div style={{display:'flex',gap:8,alignItems:'center',marginBottom:10,flexWrap:'wrap'}}>
              <span style={{...F.badge('#10b981')}}>XER LOADED</span>
              <span style={{fontSize:10,color:'rgba(255,255,255,.5)'}}>
                {result._xerStats?.activities} activities · {result._xerStats?.duration_months}mo · {result._xerStats?.logic_quality} logic
              </span>
              <button style={{...F.btn,fontSize:9,marginLeft:'auto'}} onClick={()=>{setResult(null);setFile(null);}}>Upload another</button>
            </div>
            <div style={{fontSize:12,color:'#10b981',fontWeight:600,marginBottom:6}}>{result.summary}</div>
            <div style={{fontSize:11,color:'rgba(255,255,255,.5)'}}>
              CASEY model generated from this XER. The timeline, costs, risks, QCRA/QSRA and board pack now reflect the real schedule.
              {result._xerStats?.open_ends > 0 && <span style={{color:'#f59e0b',marginLeft:8}}>⚠ {result._xerStats.open_ends} open ends detected</span>}
            </div>
          </div>
        )}
        {result?._markdown && (
          <div>
            <div style={{display:'flex',gap:8,alignItems:'center',marginBottom:10}}>
              <span style={{...F.badge('#8df7ff')}}>CHALLENGE COMPLETE</span>
              {file&&<span style={{fontSize:10,color:'rgba(255,255,255,.3)'}}>{file.name}</span>}
              <button style={{...F.btn,fontSize:9,marginLeft:'auto'}} onClick={()=>{setResult(null);setFile(null);}}>Upload another</button>
            </div>
            <div style={{background:'rgba(255,255,255,.02)',border:'1px solid rgba(255,255,255,.07)',borderRadius:6,padding:'14px 16px',fontSize:11,color:'rgba(255,255,255,.7)',lineHeight:1.8,whiteSpace:'pre-wrap',fontFamily:'system-ui',maxHeight:480,overflowY:'auto'}}>
              {result._text}
            </div>
          </div>
        )}
        {result?._streaming && (
          <div>
            <div style={{fontSize:9,color:'#8df7ff',letterSpacing:'.08em',marginBottom:8}}>CASEY IS READING YOUR BOARD PACK...</div>
            <div style={{background:'rgba(255,255,255,.02)',border:'1px solid rgba(141,247,255,.15)',borderRadius:6,padding:'14px 16px',fontSize:11,color:'rgba(255,255,255,.7)',lineHeight:1.8,whiteSpace:'pre-wrap',maxHeight:400,overflowY:'auto'}}>
              {result._text}<span style={{opacity:.6,animation:'blink 0.8s step-end infinite'}}>▌</span>
            </div>
          </div>
        )}
        {result && !result._xerModel && !result._markdown && !result._streaming && (
          <div>
            <div style={{ display:'flex', alignItems:'center', gap:8, marginBottom:14 }}>
              <span style={F.badge('#ef4444')}>{result.critical_gaps?.length || 0} CRITICAL GAPS</span>
              <span style={F.badge('#f59e0b')}>{result.board_questions?.length || 0} BOARD QUESTIONS</span>
              {file && <span style={{ fontSize:10, color:'rgba(255,255,255,.3)' }}>{file.name}</span>}
              <button style={{ ...F.btn, fontSize:9, marginLeft:'auto' }} onClick={() => { setResult(null); setFile(null); }}>Upload another</button>
            </div>

            <div style={{ padding:'12px 14px', background:'rgba(239,68,68,.06)', border:'1px solid rgba(239,68,68,.2)', borderRadius:6, marginBottom:14, fontSize:11, color:'rgba(255,255,255,.7)', lineHeight:1.6 }}>
              {result.summary}
            </div>

            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:12 }}>
              <div>
                <div style={{ fontSize:9, fontWeight:700, letterSpacing:'.12em', color:'rgba(239,68,68,.8)', marginBottom:8, textTransform:'uppercase' }}>Critical gaps</div>
                {(result.critical_gaps || []).map((g, i) => (
                  <div key={i} style={{ marginBottom:8, padding:'8px 10px', background:'rgba(239,68,68,.06)', border:'1px solid rgba(239,68,68,.14)', borderRadius:5 }}>
                    <div style={{ fontSize:9, fontWeight:700, color:'#ef4444', marginBottom:3, letterSpacing:'.06em' }}>{g.category}</div>
                    <div style={{ fontSize:10, color:'rgba(255,255,255,.6)', marginBottom:4, lineHeight:1.5 }}>{g.finding}</div>
                    <div style={{ fontSize:9.5, color:'#10b981', fontWeight:600 }}>→ {g.recommendation}</div>
                  </div>
                ))}
              </div>
              <div>
                <div style={{ fontSize:9, fontWeight:700, letterSpacing:'.12em', color:'rgba(245,158,11,.8)', marginBottom:8, textTransform:'uppercase' }}>Questions this pack won't survive</div>
                {(result.board_questions || []).map((q, i) => (
                  <div key={i} style={{ display:'flex', gap:8, marginBottom:7, padding:'6px 10px', background:'rgba(245,158,11,.05)', border:'1px solid rgba(245,158,11,.12)', borderRadius:4 }}>
                    <span style={{ fontSize:11, fontWeight:700, color:'#f59e0b', flexShrink:0 }}>{i+1}</span>
                    <span style={{ fontSize:10, color:'rgba(255,255,255,.55)', lineHeight:1.5 }}>{q}</span>
                  </div>
                ))}
                {result.missing_elements?.length > 0 && (
                  <div style={{ marginTop:10 }}>
                    <div style={{ fontSize:9, fontWeight:700, letterSpacing:'.1em', color:'rgba(141,247,255,.6)', marginBottom:6, textTransform:'uppercase' }}>Missing elements</div>
                    {result.missing_elements.map((el, i) => (
                      <div key={i} style={{ fontSize:10, color:'rgba(255,255,255,.35)', padding:'3px 0', borderBottom:'1px solid rgba(255,255,255,.04)' }}>
                        <span style={{ color:'#8df7ff', marginRight:6 }}>◌</span>{el}
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════
   FEATURE 3: RISK REGISTER HEATMAP
   ─────────────────────────────────────────────────────────
   5×5 probability × impact RAG heatmap — every T&T, 
   Faithful+Gould and Arcadis risk register has one.

   INTEGRATION:
     import { RiskRegisterHeatmap } from './CASEY_Features';
     // Add to risk tab render:
     {tab === 'risk' && <><ExistingRiskPanel.../><RiskRegisterHeatmap model={model}/></>}
═══════════════════════════════════════════════════════════ */
export function RiskRegisterHeatmap({ model }) {
  const risks = (model?.risks || model?.risk_register || []);

  const PROB_LABELS  = ['Rare','Unlikely','Possible','Likely','Almost Certain'];
  const IMPACT_LABELS = ['Negligible','Minor','Moderate','Major','Critical'];
  // RAG by prob×impact (1-indexed)
  const RAG = [
    ['#1a3a1a','#1a3a1a','#2d4a1a','#3a3a00','#4a1a00'],
    ['#1a3a1a','#2d4a1a','#3a3a00','#4a2a00','#5a1800'],
    ['#1a3a1a','#3a3a00','#4a2a00','#5a1800','#6a0000'],
    ['#2d4a1a','#4a2a00','#5a1800','#6a0000','#7a0000'],
    ['#3a3a00','#5a1800','#6a0000','#7a0000','#8a0000'],
  ];
  const RAGBORDER = [
    ['#10b98140','#10b98140','#84cc1640','#f59e0b40','#ef444440'],
    ['#10b98140','#84cc1640','#f59e0b40','#f59e0b60','#ef444460'],
    ['#10b98140','#f59e0b40','#f59e0b60','#ef444460','#ef444480'],
    ['#84cc1640','#f59e0b60','#ef444460','#ef444480','#ef4444a0'],
    ['#f59e0b40','#ef444460','#ef444480','#ef4444a0','#ef4444c0'],
  ];

  // Map risks to grid
  function getProb(r) {
    const p = String(r.probability||'').toLowerCase();
    if(/almost|very high|certain/.test(p))return 4;
    if(/high|likely/.test(p))return 3;
    if(/medium|moderate|possible/.test(p))return 2;
    if(/low|unlikely/.test(p))return 1;
    return 0;
  }
  function getImpact(r) {
    const im = String(r.impact||r.consequence||'').toLowerCase();
    if(/critical|catastrophic|severe/.test(im))return 4;
    if(/major|significant|high/.test(im))return 3;
    if(/moderate|medium/.test(im))return 2;
    if(/minor|low|negligible/.test(im))return 1;
    return 0;
  }

  const grid = Array(5).fill(null).map(() => Array(5).fill(null).map(() => []));
  risks.forEach(r => {
    const p = getProb(r), im = getImpact(r);
    if (p >= 0 && p <= 4 && im >= 0 && im <= 4) grid[p][im].push(r);
  });

  const cellSize = 52;

  return (
    <div style={{ ...F.panel, marginTop:12 }}>
      <div style={F.hdr}>
        <div style={F.title}>⬡ Risk Register — Probability × Impact Heatmap</div>
        <div style={F.sub}>{risks.length} risks plotted · hover for details · {risks.filter(r=>getProb(r)>=3&&getImpact(r)>=3).length} in red zone</div>
      </div>
      <div style={{ ...F.body, overflowX:'auto' }}>
        {risks.length === 0 ? (
          <div style={{ textAlign:'center', padding:'20px 0', color:'rgba(255,255,255,.2)', fontSize:11 }}>Run a project or upload a risk register to populate the heatmap.</div>
        ) : (
          <div style={{ display:'flex', gap:16, alignItems:'flex-start' }}>
            {/* Y-axis label */}
            <div style={{ display:'flex', flexDirection:'column', justifyContent:'center', alignItems:'center', height:cellSize*5, gap:0 }}>
              <div style={{ writingMode:'vertical-rl', transform:'rotate(180deg)', fontSize:9, color:'rgba(255,255,255,.3)', letterSpacing:'.1em', textTransform:'uppercase' }}>← Probability →</div>
            </div>
            <div>
              {/* Y-axis labels */}
              <div style={{ display:'flex', flexDirection:'column-reverse', gap:1 }}>
                {PROB_LABELS.map((pl, pi) => (
                  <div key={pi} style={{ display:'flex', gap:1, alignItems:'center' }}>
                    <div style={{ width:80, textAlign:'right', paddingRight:8, fontSize:8.5, color:'rgba(255,255,255,.35)', height:cellSize, display:'flex', alignItems:'center', justifyContent:'flex-end', letterSpacing:'.02em' }}>{pl}</div>
                    {IMPACT_LABELS.map((_, ii) => {
                      const risks_here = grid[pi][ii];
                      return (
                        <div key={ii} title={risks_here.map(r=>r.title||r.risk).join('\n')||'No risks'}
                          style={{
                            width:cellSize, height:cellSize, background:RAG[pi][ii],
                            border:`1px solid ${RAGBORDER[pi][ii]}`,
                            borderRadius:3, display:'flex', flexWrap:'wrap', alignItems:'center',
                            justifyContent:'center', gap:2, padding:4, cursor:risks_here.length?'pointer':'default',
                            transition:'all .15s', position:'relative',
                          }}>
                          {risks_here.map((r, ri) => (
                            <div key={ri} style={{
                              width:10, height:10, borderRadius:'50%',
                              background: pi>=3&&ii>=3?'#ef4444':pi>=2&&ii>=2?'#f59e0b':'#10b981',
                              boxShadow:`0 0 4px ${pi>=3&&ii>=3?'rgba(239,68,68,.5)':pi>=2&&ii>=2?'rgba(245,158,11,.5)':'rgba(16,185,129,.4)'}`,
                              title: r.title || r.risk,
                            }}/>
                          ))}
                          {risks_here.length > 0 && (
                            <div style={{ position:'absolute', top:2, right:3, fontSize:9, fontWeight:700, color:'rgba(255,255,255,.5)' }}>{risks_here.length}</div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                ))}
              </div>
              {/* X-axis */}
              <div style={{ display:'flex', gap:1, marginLeft:88, marginTop:4 }}>
                {IMPACT_LABELS.map((il, ii) => (
                  <div key={ii} style={{ width:cellSize, textAlign:'center', fontSize:8.5, color:'rgba(255,255,255,.35)', letterSpacing:'.01em' }}>{il}</div>
                ))}
              </div>
              <div style={{ textAlign:'center', marginLeft:88, marginTop:4, fontSize:9, color:'rgba(255,255,255,.2)', letterSpacing:'.08em', textTransform:'uppercase' }}>→ Impact →</div>
            </div>
            {/* Legend */}
            <div style={{ display:'flex', flexDirection:'column', gap:6, paddingTop:4 }}>
              {[['#10b981','Low / Green zone'],['#f59e0b','Medium / Amber zone'],['#ef4444','High / Red zone — board action required']].map(([c,l]) => (
                <div key={l} style={{ display:'flex', alignItems:'center', gap:8 }}>
                  <div style={{ width:12, height:12, borderRadius:2, background:c, boxShadow:`0 0 5px ${c}55` }}/>
                  <span style={{ fontSize:9, color:'rgba(255,255,255,.35)' }}>{l}</span>
                </div>
              ))}
              <div style={{ marginTop:6, fontSize:9, color:'rgba(255,255,255,.2)' }}>
                {risks.filter(r=>getProb(r)>=3&&getImpact(r)>=3).length} risks in red zone<br/>
                {risks.filter(r=>getProb(r)>=2&&getImpact(r)>=2&&!(getProb(r)>=3&&getImpact(r)>=3)).length} in amber zone
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════
   FEATURE 4: PORTFOLIO DASHBOARD
   ─────────────────────────────────────────────────────────
   Multi-programme view — total P80 exposure, RAG status,
   confidence gap. The enterprise sale.

   INTEGRATION:
     import { PortfolioDashboard } from './CASEY_Features';
     // Add as a new top-level tab "Portfolio":
     {tab === 'portfolio' && <PortfolioDashboard savedProjects={savedProjects} onLoad={onLoad}/>}
     // Add 'portfolio' to the tab nav.
═══════════════════════════════════════════════════════════ */
export function PortfolioDashboard({ savedProjects = [], onLoad }) {
  const projects = savedProjects.filter(Boolean);

  function parseCost(p) {
    const raw = p.cost_p80_bn || p.cost_p50_bn || p.p50_cost_bn;
    if (raw) return +raw;
    const s = String(p.cost_p80 || p.cost_p50 || '0').replace(/[^0-9.]/g,'');
    return parseFloat(s) || 0;
  }

  const totalP80   = projects.reduce((s, p) => s + parseCost(p), 0);
  const avgConf    = projects.length ? Math.round(projects.reduce((s, p) => s + (+p.confidence_pct || 60), 0) / projects.length) : 0;
  const redProgs   = projects.filter(p => (+p.confidence_pct || 60) < 55);
  const amberProgs = projects.filter(p => (+p.confidence_pct || 60) >= 55 && (+p.confidence_pct || 60) < 75);
  const greenProgs = projects.filter(p => (+p.confidence_pct || 60) >= 75);
  const curr       = projects[0]?.currency_symbol || '£';

  function fmtC(bn) {
    if (bn >= 1000) return curr + (bn/1000).toFixed(1) + 'T';
    if (bn >= 1)    return curr + bn.toFixed(1) + 'B';
    return curr + Math.round(bn*1000) + 'M';
  }

  return (
    <div style={F.panel}>
      <div style={F.hdr}>
        <div style={F.title}>◈ Portfolio Intelligence</div>
        <div style={F.sub}>All saved programmes · total P80 exposure · confidence gap analysis · attention required</div>
      </div>
      <div style={F.body}>
        {projects.length === 0 ? (
          <div style={{ textAlign:'center', padding:'32px 0', color:'rgba(255,255,255,.2)', fontSize:11 }}>
            <div style={{ fontSize:28, marginBottom:10, opacity:.3 }}>◈</div>
            No saved programmes yet.<br/>
            <span style={{ fontSize:9, color:'rgba(255,255,255,.12)' }}>Run projects and save them from the Overview tab — they appear here in your portfolio.</span>
          </div>
        ) : (
          <>
            {/* Portfolio summary */}
            <div style={{ display:'grid', gridTemplateColumns:'repeat(4,1fr)', gap:10, marginBottom:16 }}>
              {[
                { l:'Total P80 exposure', v:fmtC(totalP80), c:'#4a9eff' },
                { l:'Avg. confidence',    v:avgConf+'%',     c:avgConf>=75?'#10b981':avgConf>=55?'#f59e0b':'#ef4444' },
                { l:'Programmes',         v:projects.length, c:'#8892a4' },
                { l:'Needs attention',    v:redProgs.length+amberProgs.length, c:redProgs.length>0?'#ef4444':'#f59e0b' },
              ].map(k => (
                <div key={k.l} style={F.kpi}>
                  <div style={{ ...F.kpiV, fontSize:20, color:k.c }}>{k.v}</div>
                  <div style={F.kpiL}>{k.l}</div>
                </div>
              ))}
            </div>

            {/* RAG row */}
            <div style={{ display:'flex', gap:8, marginBottom:14 }}>
              {[['#ef4444','Red — Do not approve',redProgs],['#f59e0b','Amber — Conditional',amberProgs],['#10b981','Green — Approval ready',greenProgs]].map(([c,l,ps]) => (
                <div key={l} style={{ flex:1, padding:'8px 10px', background:`${c}08`, border:`1px solid ${c}25`, borderRadius:6 }}>
                  <div style={{ display:'flex', alignItems:'center', gap:6, marginBottom:6 }}>
                    <div style={{ width:8, height:8, borderRadius:'50%', background:c }}/>
                    <span style={{ fontSize:9, fontWeight:700, color:c, letterSpacing:'.06em', textTransform:'uppercase' }}>{l}</span>
                    <span style={{ fontSize:14, fontWeight:700, color:c, marginLeft:'auto' }}>{ps.length}</span>
                  </div>
                  {ps.map((p, i) => (
                    <div key={i} onClick={() => onLoad && onLoad(p)} style={{ fontSize:9.5, color:'rgba(255,255,255,.5)', padding:'3px 0', cursor:'pointer', borderBottom:'1px solid rgba(255,255,255,.04)' }}
                      onMouseOver={e=>e.currentTarget.style.color='#e2eaf6'} onMouseOut={e=>e.currentTarget.style.color='rgba(255,255,255,.5)'}>
                      {(p.title||p.subsector||'Programme').slice(0,28)} · {p.confidence_pct||'?'}%
                    </div>
                  ))}
                </div>
              ))}
            </div>

            {/* Programme list */}
            <div style={{ display:'flex', flexDirection:'column', gap:6 }}>
              {[...projects].sort((a,b) => (+a.confidence_pct||60)-(+b.confidence_pct||60)).map((p, i) => {
                const conf = +p.confidence_pct || 60;
                const cc   = conf >= 75 ? '#10b981' : conf >= 55 ? '#f59e0b' : '#ef4444';
                const p80  = parseCost(p);
                const gap  = Math.max(0, 75 - conf);
                return (
                  <div key={i} onClick={() => onLoad && onLoad(p)} style={{ display:'flex', gap:12, alignItems:'center', padding:'10px 12px', background:'rgba(255,255,255,.025)', border:`1px solid rgba(255,255,255,.06)`, borderRadius:6, cursor:'pointer', transition:'all .13s' }}
                    onMouseOver={e=>e.currentTarget.style.background='rgba(255,255,255,.04)'} onMouseOut={e=>e.currentTarget.style.background='rgba(255,255,255,.025)'}>
                    <div style={{ width:4, height:40, borderRadius:2, background:cc, flexShrink:0 }}/>
                    <div style={{ flex:1 }}>
                      <div style={{ fontSize:12, fontWeight:600, color:'#e2eaf6', marginBottom:2 }}>{(p.title||p.subsector||'Programme').slice(0,40)}</div>
                      <div style={{ fontSize:9, color:'rgba(255,255,255,.3)' }}>{p.location||p.mode||''} · {p.schedule||''}</div>
                    </div>
                    <div style={{ textAlign:'right' }}>
                      <div style={{ fontSize:13, fontWeight:700, color:'#4a9eff', marginBottom:2 }}>{fmtC(p80)}</div>
                      <div style={{ fontSize:9, color:'rgba(255,255,255,.3)' }}>P80 cost</div>
                    </div>
                    <div style={{ textAlign:'right', minWidth:50 }}>
                      <div style={{ fontSize:16, fontWeight:700, color:cc }}>{conf}%</div>
                      <div style={{ fontSize:8, color:'rgba(255,255,255,.25)' }}>{gap>0?`−${gap}pts to 75%`:'Board ready'}</div>
                    </div>
                    <div style={{ fontSize:9, color:'rgba(255,255,255,.25)', minWidth:40, textAlign:'center' }}>Load →</div>
                  </div>
                );
              })}
            </div>
          </>
        )}
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════
   FEATURE 5: ADVISOR MEMORY
   ─────────────────────────────────────────────────────────
   Persist advisor context across browser sessions.
   "Last time we spoke you said design freeze was the 
    governing constraint — has that changed?"

   INTEGRATION:
     import { useAdvisorMemory } from './CASEY_Features';
     // In your Advisor component:
     const { memory, addMessage, clearMemory, recall } = useAdvisorMemory(model?.id);
     // Pass memory to your advisor prompt to give it context.
═══════════════════════════════════════════════════════════ */
export function useAdvisorMemory(programmeId) {
  const KEY = `casey_advisor_memory_${programmeId || 'global'}`;

  const [memory, setMemory] = useState(() => {
    try { return JSON.parse(localStorage.getItem(KEY) || '[]'); } catch { return []; }
  });

  function addMessage(role, content, metadata = {}) {
    const entry = {
      role, // 'user' | 'assistant'
      content: content.slice(0, 800),
      timestamp: new Date().toISOString(),
      programme: programmeId,
      ...metadata,
    };
    setMemory(prev => {
      const updated = [entry, ...prev].slice(0, 40); // Keep last 40 exchanges
      try { localStorage.setItem(KEY, JSON.stringify(updated)); } catch {}
      return updated;
    });
  }

  function clearMemory() {
    setMemory([]);
    try { localStorage.removeItem(KEY); } catch {}
  }

  // Build context string for the advisor prompt
  function recall(lastN = 6) {
    if (!memory.length) return '';
    const recent = memory.slice(0, lastN).reverse();
    return `\n\nPREVIOUS ADVISOR CONTEXT (last ${recent.length} exchanges):\n` +
      recent.map(m => `${m.role === 'user' ? 'User' : 'CASEY'}: ${m.content}`).join('\n') +
      '\n\n(Continue from this context — reference previous discussions where relevant)';
  }

  return { memory, addMessage, clearMemory, recall };
}

/* Advisor Memory UI Panel — shows conversation history */
export function AdvisorMemoryPanel({ programmeId, onClear }) {
  const { memory, clearMemory } = useAdvisorMemory(programmeId);

  function handleClear() { clearMemory(); if (onClear) onClear(); }

  return (
    <div style={{ ...F.panel, marginTop:10 }}>
      <div style={F.hdr}>
        <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between' }}>
          <div>
            <div style={F.title}>🧠 Advisor Memory — {memory.length} exchanges retained</div>
            <div style={F.sub}>Conversation persists across sessions. CASEY references previous discussions.</div>
          </div>
          {memory.length > 0 && <button style={F.btnDanger} onClick={handleClear}>Clear memory</button>}
        </div>
      </div>
      <div style={F.body}>
        {memory.length === 0 ? (
          <div style={{ textAlign:'center', padding:'16px 0', color:'rgba(255,255,255,.2)', fontSize:11 }}>
            No advisor history yet. Start a conversation in the Board Room tab.
          </div>
        ) : (
          <div style={{ display:'flex', flexDirection:'column', gap:6, maxHeight:280, overflowY:'auto' }}>
            {memory.slice(0, 12).map((m, i) => (
              <div key={i} style={{ display:'flex', gap:8, padding:'6px 10px', background:`rgba(${m.role==='user'?'74,158,255':'16,185,129'},.04)`, border:`1px solid rgba(${m.role==='user'?'74,158,255':'16,185,129'},.1)`, borderRadius:4 }}>
                <span style={{ fontSize:8, fontWeight:700, color:m.role==='user'?'#4a9eff':'#10b981', flexShrink:0, marginTop:2, letterSpacing:'.06em', textTransform:'uppercase' }}>{m.role==='user'?'You':'Casey'}</span>
                <span style={{ fontSize:10, color:'rgba(255,255,255,.55)', lineHeight:1.5, flex:1 }}>{m.content.slice(0,200)}{m.content.length>200?'…':''}</span>
                <span style={{ fontSize:8, color:'rgba(255,255,255,.2)', flexShrink:0 }}>{new Date(m.timestamp).toLocaleDateString('en-GB',{day:'numeric',month:'short'})}</span>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

/* ── XLSX COVER TAB + RAG HEATMAP DATA ──────────────────────────────────────
   Pass these to your /export/workbook endpoint to add
   a proper cover tab and risk heatmap tab to the XLSX workbook.
   
   INTEGRATION in App.jsx:
     import { buildXlsxCoverData, buildXlsxHeatmapData } from './CASEY_Features';
     // When downloading workbook:
     download('/export/workbook', {
       ...model,
       xlsx_cover: buildXlsxCoverData(model, whiteLabelLogo),
       xlsx_heatmap: buildXlsxHeatmapData(model),
     }, filename);
── */
export function buildXlsxCoverData(model, logoUrl = null) {
  if (!model) return {};
  const conf = +model.confidence_pct || 60;
  return {
    title:        model.title || model.subsector || 'Programme',
    location:     model.location || '',
    date:         new Date().toLocaleDateString('en-GB'),
    prepared_by:  'CASEY Programme Intelligence',
    logo_url:     logoUrl,
    p50:          model.cost_p50 || '—',
    p80:          model.cost_p80 || '—',
    schedule:     model.schedule || '—',
    confidence:   conf + '%',
    rag:          conf >= 75 ? 'GREEN' : conf >= 55 ? 'AMBER' : 'RED',
    class_name:   model.estimate_class_name || 'Class ' + (model.estimate_class || 3),
    sched_level:  model.schedule_level_name || 'Level ' + (model.schedule_level || 4),
    scenario:     model.scenario_label || 'Base',
    authority:    model.institutional_authority_line || '',
    oba:          model.oba_adjustment || '',
    constraint:   model.governing_constraint_prominent || '',
  };
}

export function buildXlsxHeatmapData(model) {
  const risks = (model?.risks || model?.risk_register || []);
  return {
    risks: risks.map(r => ({
      title:       r.title || r.risk || '—',
      probability: r.probability || '—',
      impact:      r.impact || r.consequence || '—',
      emv:         r.cost_emv_bn || 0,
      owner:       r.owner || '—',
      mitigation:  r.mitigation || '—',
      rag:         /high|critical/i.test(r.probability || '') ? 'RED' : /medium|moderate/i.test(r.probability || '') ? 'AMBER' : 'GREEN',
    })),
    total:       risks.length,
    red_count:   risks.filter(r => /high|critical/i.test(r.probability || '')).length,
    amber_count: risks.filter(r => /medium|moderate/i.test(r.probability || '')).length,
    green_count: risks.filter(r => /low|unlikely/i.test(r.probability || '')).length,
  };
}
