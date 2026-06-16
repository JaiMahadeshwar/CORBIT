/**
 * ProjectTimeline.jsx — CASEY Programme Intelligence Timeline
 * THE DEFINITIVE VERSION
 * ================================================================
 * ✅ Particle comet trails behind moving heads
 * ✅ Risk event flash — canvas brightens when risk fires
 * ✅ Milestone pop — dots scale in when first revealed  
 * ✅ Live P80 callout box — updates with every frame
 * ✅ "NOW" label on scan line — Bloomberg feel
 * ✅ Dual-layer glow on active track
 * ✅ Spend area fill — rich gradient under each curve
 * ✅ Sector-coloured nebula — Space=purple, Pharma=teal, Rail=blue
 * ✅ Mobile & iPad: ResizeObserver + touch scrub + adaptive height
 * ✅ actualProgress prop — monthly actuals move the head
 * ✅ Four tracks · per-track risks · spend curves fully labelled
 */

import React,{useEffect,useRef,useState,useCallback,useMemo}from'react';

/* ── HELPERS ── */
function parseMonths(m){if(+m.schedule_months>0)return+m.schedule_months;if(+m.duration_months>0)return+m.duration_months;const n=String(m.schedule||'').match(/(\d+)/);return n?+n[1]:24;}
function parseCostBn(m){const raw=m.cost_p50_bn||m.p50_cost_bn;if(+raw>0)return+raw;const s=String(m.cost_p50||m.p50||'1').replace(/[^0-9.TBMKtbmk]/g,'').toUpperCase();const n=parseFloat(s)||1;if(s.includes('T'))return n*1000;if(s.includes('B'))return n;if(s.includes('M'))return n/1000;return n;}
function parseStart(m){if(m.start_date){const d=new Date(m.start_date);if(!isNaN(d))return d;}const n=new Date();return new Date(n.getFullYear(),n.getMonth(),1);}
function addM(d,n){const r=new Date(d);r.setMonth(r.getMonth()+Math.round(n));return r;}
function fmtD(d){return d.toLocaleDateString('en-GB',{month:'short',year:'numeric'});}
function fmtC(bn,c='£'){if(bn>=1000)return`${c}${(bn/1000).toFixed(1)}T`;if(bn>=10)return`${c}${bn.toFixed(1)}B`;if(bn>=1)return`${c}${bn.toFixed(2)}B`;return`${c}${Math.round(bn*1000)}M`;}
function scurve(t){return 1/(1+Math.exp(-10*(t-0.5)));}
function getSector(m){const s=String(m.subsector||m.mode||m.sector||'').toLowerCase();if(String(m.mode||'').toLowerCase()==='space')return'space';if(/rail|transit|metro/.test(s))return'rail';if(/nuclear|smr/.test(s))return'nuclear';if(/pharma|bio|life|glp/.test(s))return'pharma';if(/data.cent|hyperscale|ai.campus/.test(s))return'data';if(/defence|naval|submarine/.test(s))return'defence';if(/energy|wind|solar|lng/.test(s))return'energy';if(/airport|aviation/.test(s))return'airport';return'generic';}

function canvasH(){const w=window.innerWidth;if(w<=480)return 280;if(w<=768)return 340;if(w<=1024)return 380;return 440;}

/* Sector nebula colours */
const NEBULA={
  space:'rgba(139,92,246,',rail:'rgba(74,158,255,',nuclear:'rgba(239,68,68,',
  pharma:'rgba(16,185,129,',data:'rgba(6,182,212,',defence:'rgba(30,58,95,',
  energy:'rgba(245,158,11,',airport:'rgba(99,102,241,',generic:'rgba(74,158,255,',
};

const MS={
  space:[{t:0,l:'Start'},{t:.09,l:'PDR'},{t:.24,l:'CDR'},{t:.42,l:'MAIT'},{t:.60,l:'FRR'},{t:.78,l:'Launch ready'},{t:.92,l:'Launch'},{t:1,l:'Mission ops'}],
  rail:[{t:0,l:'Auth'},{t:.09,l:'Design'},{t:.24,l:'Procurement'},{t:.42,l:'Civil mobilise'},{t:.60,l:'Systems install'},{t:.76,l:'Testing'},{t:.91,l:'Commissioning'},{t:1,l:'Revenue service'}],
  nuclear:[{t:0,l:'FID'},{t:.07,l:'Site prep'},{t:.22,l:'Civil works'},{t:.44,l:'Plant install'},{t:.63,l:'Cold comm.'},{t:.79,l:'Hot comm.'},{t:.93,l:'Grid sync'},{t:1,l:'Commercial ops'}],
  pharma:[{t:0,l:'Filing'},{t:.11,l:'Mobilise'},{t:.30,l:'Construction'},{t:.52,l:'Fit-out'},{t:.66,l:'IQ/OQ'},{t:.79,l:'PQ valid.'},{t:.91,l:'Reg approval'},{t:1,l:'Batch 1'}],
  data:[{t:0,l:'FID'},{t:.12,l:'Shell'},{t:.32,l:'MEP install'},{t:.52,l:'IT fit-out'},{t:.70,l:'Cooling comm.'},{t:.84,l:'IT comm.'},{t:.93,l:'Ops ready'},{t:1,l:'Hyperscale live'}],
  defence:[{t:0,l:'Award'},{t:.09,l:'Design'},{t:.24,l:'Long-lead'},{t:.42,l:'Build'},{t:.62,l:'Integration'},{t:.78,l:'Acceptance'},{t:.91,l:'Trials'},{t:1,l:'In-service'}],
  energy:[{t:0,l:'FID'},{t:.13,l:'Procurement'},{t:.32,l:'Civil'},{t:.57,l:'Install'},{t:.74,l:'Cold comm.'},{t:.88,l:'Grid sync'},{t:.96,l:'Ops'},{t:1,l:'Full capacity'}],
  generic:[{t:0,l:'Start'},{t:.11,l:'Mobilise'},{t:.30,l:'Design'},{t:.47,l:'Procurement'},{t:.64,l:'Build'},{t:.80,l:'Commission'},{t:.93,l:'Testing'},{t:1,l:'Handover'}],
};

const SR={
  rail:[
    {t:.07,l:'Systems integration',imp:'+6mo · +5% cost',drv:'Signalling locked to staged possessions — each missed slot unrecoverable',sev:'high'},
    {t:.19,l:'Rolling stock delay',imp:'+4mo schedule',drv:'Fleet not available at line opening — phantom timetable risk',sev:'med'},
    {t:.33,l:'Possessions conflict',imp:'+8mo schedule',drv:'NR window availability limits parallel workfronts',sev:'high'},
    {t:.47,l:'Cost escalation',imp:'+8% cost',drv:'Inflation and productivity loss consume P80 reserve',sev:'high'},
    {t:.62,l:'Interface complexity',imp:'+4mo · +3% cost',drv:'Multi-contractor rework under live railway operations',sev:'high'},
    {t:.76,l:'Station scope growth',imp:'+6mo · +6% cost',drv:'Each addition adds £200-500M and 6-12 months',sev:'high'},
    {t:.88,l:'Workforce availability',imp:'+8mo schedule',drv:'Specialist rail trades scarce across simultaneous contracts',sev:'high'},
  ],
  space:[
    {t:.09,l:'TRL readiness gap',imp:'+12mo · +12% cost',drv:'Technology not mature at mission architecture PDR',sev:'high'},
    {t:.24,l:'Launch manifest conflict',imp:'+8mo schedule',drv:'Launch vehicle availability constrained by manifest priority',sev:'high'},
    {t:.40,l:'MAIT complexity',imp:'+10mo · +8% cost',drv:'Integration reveals interface incompatibilities',sev:'high'},
    {t:.56,l:'FOAK supply chain',imp:'+6% cost',drv:'First-of-a-kind component lead times exceed forecast',sev:'med'},
    {t:.72,l:'Ground systems delay',imp:'+6mo schedule',drv:'Operations centre readiness behind spacecraft',sev:'med'},
    {t:.87,l:'Regulatory / spectrum',imp:'+12mo schedule',drv:'Launch approval and spectrum licensing on critical path',sev:'high'},
  ],
  pharma:[
    {t:.11,l:'FDA / EMA submission',imp:'+8mo schedule',drv:'Regulatory filing timeline drives commercial readiness date',sev:'high'},
    {t:.30,l:'OEM equipment lead time',imp:'+6mo · +5% cost',drv:'Single-source fill-finish equipment — no alternative supplier',sev:'high'},
    {t:.49,l:'PQ batch failure',imp:'+12mo · +10% cost',drv:'First commercial batch fails validation — repeat cycle required',sev:'high'},
    {t:.65,l:'Clean utility qualification',imp:'+4mo schedule',drv:'WFI/PW/clean steam commissioning sequence',sev:'med'},
    {t:.80,l:'Cold chain logistics',imp:'+3% cost',drv:'-80°C storage and transport capacity constrained',sev:'med'},
  ],
  data:[
    {t:.09,l:'Grid connection delay',imp:'+12mo · +3% cost',drv:'DNO queue position is not an energisation date',sev:'high'},
    {t:.27,l:'GPU / compute allocation',imp:'+8% cost',drv:'Hyperscaler GPU allocation not confirmed at FID',sev:'high'},
    {t:.45,l:'Cooling system design',imp:'+6mo schedule',drv:'Liquid cooling density requirements changed late',sev:'high'},
    {t:.63,l:'MEP interface clash',imp:'+4mo schedule',drv:'Power and cooling routes conflict in BIM model',sev:'med'},
    {t:.80,l:'IT commissioning',imp:'+6mo schedule',drv:'Integrated systems test reveals rack-level dependencies',sev:'med'},
  ],
  defence:[
    {t:.10,l:'Scope creep',imp:'+8mo · +10% cost',drv:'Capability requirement changes post-award',sev:'high'},
    {t:.28,l:'Long-lead supply',imp:'+6mo · +6% cost',drv:'Critical components from single-source suppliers',sev:'high'},
    {t:.48,l:'System integration',imp:'+10mo · +8% cost',drv:'Sub-system compatibility issues in integration phase',sev:'high'},
    {t:.68,l:'Acceptance trials',imp:'+6mo schedule',drv:'Performance shortfalls require remediation',sev:'med'},
    {t:.85,l:'Export controls',imp:'+4mo schedule',drv:'Regulatory compliance delays delivery',sev:'med'},
  ],
  generic:[
    {t:.11,l:'Scope definition',imp:'+4mo · +4% cost',drv:'Design freeze not confirmed at procurement gateway',sev:'high'},
    {t:.28,l:'Procurement delay',imp:'+6mo schedule',drv:'Market capacity constraints delay contract award',sev:'med'},
    {t:.46,l:'Interface management',imp:'+3mo · +3% cost',drv:'Multi-contractor coordination gap',sev:'high'},
    {t:.64,l:'Commissioning risk',imp:'+8mo schedule',drv:'System integration complexity exceeds plan',sev:'high'},
    {t:.82,l:'Regulatory approval',imp:'+5mo schedule',drv:'Submission timeline on critical path',sev:'med'},
  ],
};
SR.nuclear=SR.energy=SR.airport=SR.generic;

function buildTrackRisks(model,sector){
  const raw=(model.risks||[]).slice(0,7);const bn=parseCostBn(model);
  const baseR=raw.length>=3?raw.map((r,i)=>{
    const hi=/high|critical/i.test(r.probability||r.impact||'');
    const sw=+r.schedule_impact_weeks||(hi?6:3);const cb=+r.cost_impact_bn||0;
    const cp=cb>0?Math.round(cb/bn*100):(hi?4:2);
    let imp='';if(cp>0)imp+=`+${cp}% cost`;if(sw>0)imp+=`${imp?' · ':''} +${sw}mo`;
    return{t:0.08+(i/Math.max(raw.length-1,1))*0.82,l:(r.title||r.risk||`Risk ${i+1}`).slice(0,32),imp:imp||r.impact||'Review required',drv:(r.cause||r.owner||'').slice(0,70)||'Programme intelligence',sev:hi?'high':'med'};
  }):(SR[sector]||SR.generic);
  return{
    base:baseR,
    scenario:baseR.filter((_,i)=>i%2===0).map(r=>({...r,t:r.t*0.88,imp:r.imp+' (mitigated)',sev:'med'})),
    benchmark:baseR.map(r=>({...r,imp:r.imp+' (outturn basis)',sev:'high'})),
    stress:baseR.map(r=>({...r,imp:r.imp+' (P90 materialisation)',sev:'high'})),
  };
}
function buildScenario(m){const b=parseMonths(m),c=parseCostBn(m);const mx=(m.scenario_matrix||[]).find(s=>/faster|cheaper|optimis|acceler/i.test(s.scenario||s.label||''))||(m.scenario_matrix||[])[1];return{label:mx?.label||'Faster scenario',schedMult:mx?(+mx.schedule_months||b*0.88)/b:0.88,costMult:mx?(+mx.cost_p50_bn||c*1.05)/c:1.05,confPct:+(mx?.confidence_pct||Math.max(40,+m.confidence_pct-8))};}
function buildBenchmark(m){const b=parseMonths(m);const bp=(m.benchmark_programmes||m.benchmarks||[])[0];return{label:bp?(bp.name||bp.programme||'Benchmark').slice(0,22):'Sector benchmark',schedMult:bp?(+bp.schedule_months||b*1.16)/b:1.16,costMult:bp?1+(+bp.cost_growth_pct||28)/100:1.28,confPct:+(bp?.confidence_pct||71)};}
function buildStress(m){const b=parseMonths(m),c=parseCostBn(m);const q=m.monte_carlo;return{label:'Stress test P90',schedMult:q?.qsra?.p90?+q.qsra.p90/b:1.38,costMult:q?.qcra?.p90?+q.qcra.p90/c:1.44,confPct:Math.max(22,+m.confidence_pct-30)};}

const TC={base:'#4a9eff',scenario:'#10b981',benchmark:'#f59e0b',stress:'#ef4444'};

/* ── PRIMITIVE DRAW HELPERS ── */
function dot(ctx,x,y,c,r,glow=0){
  if(glow>0){ctx.save();ctx.shadowColor=c;ctx.shadowBlur=glow;}
  ctx.save();ctx.fillStyle=c;ctx.beginPath();ctx.arc(x,y,r,0,Math.PI*2);ctx.fill();
  ctx.strokeStyle='rgba(0,0,0,.6)';ctx.lineWidth=Math.max(.6,r*.25);ctx.beginPath();ctx.arc(x,y,r,0,Math.PI*2);ctx.stroke();
  ctx.restore();if(glow>0)ctx.restore();
}

function delivFlag(ctx,tr,dx,W,PT,delivDate,mini){
  if(dx<46||dx>W+80)return;
  ctx.save();
  ctx.strokeStyle=tr.color+'55';ctx.lineWidth=.8;ctx.setLineDash([3,4]);
  ctx.beginPath();ctx.moveTo(dx,PT+2);ctx.lineTo(dx,PT+36);ctx.stroke();ctx.setLineDash([]);
  const fw=mini?70:90,fh=28,fx=Math.max(44,Math.min(dx-fw/2,W-10-fw));
  ctx.fillStyle='rgba(3,5,14,.95)';ctx.strokeStyle=tr.color;ctx.lineWidth=.9;
  ctx.beginPath();ctx.roundRect(fx,PT+3,fw,fh,3);ctx.fill();ctx.stroke();
  ctx.fillStyle=tr.color;ctx.beginPath();ctx.roundRect(fx,PT+3,fw,3,[3,3,0,0]);ctx.fill();
  ctx.font=`700 ${mini?6.5:7.5}px 'SF Mono',monospace`;ctx.fillStyle=tr.color;ctx.textAlign='center';
  ctx.fillText(tr.label.slice(0,mini?8:11),dx,PT+14);
  ctx.font=`${mini?6:7.5}px 'SF Mono',monospace`;ctx.fillStyle='#c0d4ee';
  ctx.fillText(fmtD(delivDate),dx,PT+25);ctx.restore();
}

/* ═══════════════════════════════════════════════════════════
   THE PAINT FUNCTION — everything visual lives here
═══════════════════════════════════════════════════════════ */
function paintCanvas(ctx,W,H,state){
  const{prog,activeMode,milestones,trackRisks,scenario,benchmark,stress,
        startDate,totalMonths,baseCostBn,currency,revealedMap,glowPhase,
        sector,recentFlash,milestoneReveal,p80CostBn}=state;

  ctx.clearRect(0,0,W,H);

  const mob=W<500,tab=W<800;
  const PL=mob?42:60,PR=mob?8:18,PT=mob?46:58,PB=20;
  const SPEND_H=H*(mob?.23:.30),SPEND_BOT=H*(mob?.27:.34);
  const TL_MID=H*(mob?.62:.65);
  const MAX_COST=baseCostBn*(activeMode==='stress'?stress.costMult*1.08:activeMode==='benchmark'?benchmark.costMult*1.06:1.55);

  const acts=[
    {id:'base',    label:'BASE',      shortL:'Base',  color:TC.base,     lw:mob?2.2:3.0, dash:[],   schedMult:1,                 costMult:1,                  risks:trackRisks.base},
    ...((activeMode==='scenario'||activeMode==='benchmark')?[{id:'scenario', label:(scenario.label||'FASTER').slice(0,12).toUpperCase(),shortL:'Faster',color:TC.scenario,lw:mob?1.6:2.0,dash:[8,4], schedMult:scenario.schedMult,  costMult:scenario.costMult,  risks:trackRisks.scenario}]:[]),
    ...(activeMode==='benchmark'?[{id:'benchmark',label:'BENCHMARK',shortL:'Bench', color:TC.benchmark,lw:mob?1.3:1.7,dash:[5,4],schedMult:benchmark.schedMult, costMult:benchmark.costMult, risks:trackRisks.benchmark}]:[]),
    ...(activeMode==='stress'?   [{id:'stress',   label:'STRESS P90',shortL:'P90',  color:TC.stress,   lw:mob?1.6:2.0,dash:[5,3], schedMult:stress.schedMult,   costMult:stress.costMult,    risks:trackRisks.stress}]:[]),
  ];

  const maxDM=Math.max(...acts.map(t=>t.schedMult*totalMonths))*1.07;
  const TW=W-PL-PR;
  const RSEP=acts.length<=2?(mob?20:30):acts.length===3?(mob?15:23):(mob?11:17);

  const xAt=(f,sm)=>PL+Math.min((f*(sm||1)*totalMonths)/maxDM,1.24)*TW;
  const spY=(r)=>SPEND_BOT-Math.min(r,1.75)*SPEND_H;
  const ry=(i)=>TL_MID+(i-(acts.length-1)/2)*RSEP;

  /* ─ 1. DEEP SPACE BACKGROUND with sector nebula ─ */
  const bg=ctx.createLinearGradient(0,0,0,H);
  bg.addColorStop(0,'#02050b');bg.addColorStop(.6,'#030710');bg.addColorStop(1,'#040811');
  ctx.fillStyle=bg;ctx.fillRect(0,0,W,H);

  // Sector-coloured nebula glow
  const nebCol=NEBULA[sector]||NEBULA.generic;
  const neb=ctx.createRadialGradient(W*.72,H*.28,0,W*.72,H*.28,W*.65);
  neb.addColorStop(0,nebCol+'0.028)');neb.addColorStop(.5,nebCol+'0.012)');neb.addColorStop(1,'rgba(0,0,0,0)');
  ctx.fillStyle=neb;ctx.fillRect(0,0,W,H);

  // Secondary nebula (bottom-left for depth)
  if(!mob){
    const neb2=ctx.createRadialGradient(W*.15,H*.8,0,W*.15,H*.8,W*.4);
    neb2.addColorStop(0,nebCol+'0.015)');neb2.addColorStop(1,'rgba(0,0,0,0)');
    ctx.fillStyle=neb2;ctx.fillRect(0,0,W,H);
  }

  /* ─ 2. STARFIELD — twinkles with glowPhase ─ */
  const SC=mob?28:60;
  for(let i=0;i<SC;i++){
    const sx=((i*7919+13)%997)/997*W;
    const sy=((i*6271+7)%997)/997*(H*.9)+4;
    const sr=((i*2111)%80)/80*1.3+0.1;
    const base=0.06+((i*3571)%100)/100*0.16;
    const twinkle=base*(0.6+0.4*Math.sin(glowPhase*0.4+i*0.55));
    ctx.save();ctx.globalAlpha=twinkle;
    // Some stars get the sector colour tint
    ctx.fillStyle=i%8===0?acts[0].color+'cc':'#ffffff';
    ctx.beginPath();ctx.arc(sx,sy,sr,0,Math.PI*2);ctx.fill();ctx.restore();
  }

  /* ─ 3. RISK FLASH — brief brightening when risk fires ─ */
  if(recentFlash&&recentFlash.t>0){
    const fx=xAt(recentFlash.riskT,1);
    const fg=ctx.createRadialGradient(fx,TL_MID,0,fx,TL_MID,80);
    fg.addColorStop(0,recentFlash.color+`${Math.round(recentFlash.t*40).toString(16).padStart(2,'0')}`);
    fg.addColorStop(1,'rgba(0,0,0,0)');
    ctx.fillStyle=fg;ctx.fillRect(0,0,W,H);
  }

  /* ─ 4. VERTICAL GRID ─ */
  const gs=totalMonths<=36?6:totalMonths<=84?12:24;
  for(let m=0;m<=maxDM;m+=gs){
    const x=PL+(m/maxDM)*TW;if(x>W+20)break;
    const gg=ctx.createLinearGradient(0,PT,0,H-PB);
    gg.addColorStop(0,'rgba(255,255,255,0)');
    gg.addColorStop(.35,'rgba(255,255,255,.025)');
    gg.addColorStop(1,'rgba(255,255,255,0)');
    ctx.save();ctx.strokeStyle=gg;ctx.lineWidth=.35;
    ctx.beginPath();ctx.moveTo(x,PT);ctx.lineTo(x,H-PB);ctx.stroke();ctx.restore();
  }

  /* ─ 5. DATE AXIS ─ */
  ctx.font=`${mob?6.5:7.5}px 'SF Mono','Fira Code',monospace`;
  ctx.textAlign='center';ctx.fillStyle='rgba(42,66,100,.9)';
  const dStep=mob?gs*2:gs;
  for(let m=0;m<=maxDM;m+=dStep){
    const x=PL+(m/maxDM)*TW;if(x>W+20)break;
    ctx.fillText(fmtD(addM(startDate,m)),x,H-4);
  }

  /* ─ 6. SPEND ZONE DIVIDER ─ */
  const sdg=ctx.createLinearGradient(PL,0,W-PR,0);
  sdg.addColorStop(0,'rgba(255,255,255,0)');
  sdg.addColorStop(.25,'rgba(255,255,255,.06)');
  sdg.addColorStop(.75,'rgba(255,255,255,.06)');
  sdg.addColorStop(1,'rgba(255,255,255,0)');
  ctx.save();ctx.strokeStyle=sdg;ctx.lineWidth=.7;
  ctx.beginPath();ctx.moveTo(PL,SPEND_BOT);ctx.lineTo(W-PR,SPEND_BOT);ctx.stroke();ctx.restore();
  // Zone label on the divider
  if(!mob){
    ctx.save();ctx.font=`7px 'SF Mono',monospace`;ctx.fillStyle='rgba(38,58,90,.5)';ctx.textAlign='left';
    ctx.fillText('▲ SPEND',PL,PT-3);
    ctx.fillText('▼ TIMELINE',PL,TL_MID-(acts.length-1)/2*RSEP-18);
    ctx.restore();
  }

  /* ─ 7. SPEND AXIS ─ */
  ctx.textAlign='right';ctx.font=`${mob?6:7}px 'SF Mono',monospace`;
  const spT=mob?[0,.5,1]:[0,.25,.5,.75,1];
  spT.forEach(r=>{
    const y=spY(r);if(y<PT-4)return;
    ctx.fillStyle='rgba(38,58,90,.75)';ctx.fillText(fmtC(MAX_COST*r,currency),PL-5,y+3);
    ctx.save();ctx.strokeStyle='rgba(255,255,255,.012)';ctx.lineWidth=.3;
    ctx.beginPath();ctx.moveTo(PL,y);ctx.lineTo(W-PR,y);ctx.stroke();ctx.restore();
  });

  /* ─ 8. GHOST PREDICTED FUTURE ─ */
  if(prog>0.01&&prog<0.98){
    acts.forEach((tr,i)=>{
      const ryi=ry(i),hx=xAt(prog,tr.schedMult),ex=xAt(1,tr.schedMult);
      // Ghost rail
      ctx.save();ctx.globalAlpha=.20;ctx.strokeStyle=tr.color;
      ctx.lineWidth=tr.lw*.55;ctx.setLineDash([3,8]);
      ctx.beginPath();ctx.moveTo(hx,ryi);ctx.lineTo(ex,ryi);ctx.stroke();ctx.setLineDash([]);
      // Ghost spend curve
      if(!mob){
        ctx.globalAlpha=.12;ctx.strokeStyle=tr.color;ctx.lineWidth=.9;ctx.setLineDash([2,9]);
        ctx.beginPath();let gs2=false;
        for(let k=0;k<=90;k++){
          const t=prog+(k/90)*(1-prog);
          const gx=xAt(t,tr.schedMult),gy=spY(scurve(t)*tr.costMult*baseCostBn/MAX_COST);
          gs2?ctx.lineTo(gx,gy):(ctx.moveTo(gx,gy),gs2=true);
        }
        ctx.stroke();ctx.setLineDash([]);
      }
      ctx.restore();
    });
  }

  /* ─ 9. SPEND CURVES with rich area fills ─ */
  acts.forEach((tr,i)=>{
    // Area fill
    if(prog>0.04){
      const af=ctx.createLinearGradient(PL,SPEND_BOT,PL,SPEND_BOT-SPEND_H);
      af.addColorStop(0,tr.color+'00');
      af.addColorStop(.4,tr.color+(i===0?'0a':'06'));
      af.addColorStop(1,tr.color+(i===0?'18':'0c'));
      ctx.save();ctx.globalAlpha=.85;ctx.fillStyle=af;
      ctx.beginPath();ctx.moveTo(PL,SPEND_BOT);
      for(let k=0;k<=180;k++){
        const t=k/180;if(t>prog)break;
        ctx.lineTo(xAt(t,tr.schedMult),spY(scurve(t)*tr.costMult*baseCostBn/MAX_COST));
      }
      ctx.lineTo(xAt(prog,tr.schedMult),SPEND_BOT);ctx.closePath();ctx.fill();ctx.restore();
    }
    // Curve line
    ctx.save();ctx.strokeStyle=tr.color;ctx.lineWidth=i===0?1.6:1.0;
    ctx.lineCap='round';ctx.globalAlpha=.85;ctx.setLineDash(tr.dash);
    ctx.beginPath();let s=false;
    for(let k=0;k<=220;k++){
      const t=k/220;if(t>prog)break;
      const sx2=xAt(t,tr.schedMult),sy2=spY(scurve(t)*tr.costMult*baseCostBn/MAX_COST);
      s?ctx.lineTo(sx2,sy2):(ctx.moveTo(sx2,sy2),s=true);
    }
    ctx.stroke();ctx.setLineDash([]);ctx.restore();

    // START anchor
    if(i===0){
      dot(ctx,PL,SPEND_BOT,tr.color,mob?3.5:4.5,6);
      if(!mob){
        ctx.save();ctx.font=`700 7px 'SF Mono',monospace`;ctx.fillStyle='rgba(38,58,90,.9)';
        ctx.textAlign='left';ctx.fillText(`${fmtC(0,currency)}  ·  ${fmtD(startDate)}`,PL+7,SPEND_BOT+11);ctx.restore();
      }
    }

    // END anchor — delivery + total spend
    if(prog>=0.95){
      const ex=xAt(1,tr.schedMult),ey=spY(tr.costMult*baseCostBn/MAX_COST);
      dot(ctx,ex,ey,tr.color,mob?4.5:6,12);
      // Drop line to rail
      ctx.save();ctx.strokeStyle=tr.color+'45';ctx.lineWidth=.6;ctx.setLineDash([2,4]);
      ctx.beginPath();ctx.moveTo(ex,ey+6);ctx.lineTo(ex,ry(i)-7);ctx.stroke();ctx.setLineDash([]);ctx.restore();
      // End label
      if(!mob){
        const ll=ex>W-160,lx=ll?ex-8:ex+8,la=ll?'right':'left';
        ctx.save();ctx.textAlign=la;ctx.shadowColor=tr.color;ctx.shadowBlur=5;
        ctx.font=`700 9px 'SF Mono',monospace`;ctx.fillStyle=tr.color;
        ctx.fillText(fmtC(tr.costMult*baseCostBn,currency)+' TOTAL',lx,ey-16);
        ctx.font=`7.5px 'SF Mono',monospace`;ctx.fillStyle=tr.color+'90';
        ctx.fillText('Delivery: '+fmtD(addM(startDate,totalMonths*tr.schedMult)),lx,ey-4);
        ctx.fillText(tr.label,lx,ey+9);ctx.restore();
      }
    }

    // LIVE HEAD on spend curve + spend callout
    if(prog>0.01){
      const hx=xAt(prog,tr.schedMult);
      const hy=spY(scurve(prog)*tr.costMult*baseCostBn/MAX_COST);
      dot(ctx,hx,hy,tr.color,mob?3:4,8);
      if(i===0&&!mob){
        const spent=scurve(prog)*baseCostBn;
        const pct=Math.round(spent/baseCostBn*100);
        // P80 callout box
        const p80Bn=p80CostBn||baseCostBn*1.18;
        const p80y=spY(p80Bn/MAX_COST);
        // Dashed P80 line
        ctx.save();ctx.strokeStyle='rgba(245,158,11,.22)';ctx.lineWidth=.6;ctx.setLineDash([3,5]);
        ctx.beginPath();ctx.moveTo(PL,p80y);ctx.lineTo(W-PR,p80y);ctx.stroke();ctx.setLineDash([]);
        ctx.font=`700 7px 'SF Mono',monospace`;ctx.fillStyle='rgba(245,158,11,.45)';ctx.textAlign='right';
        ctx.fillText('P80 '+fmtC(p80Bn,currency),W-PR-4,p80y-3);ctx.restore();
        // Live spend callout
        const cx2=Math.min(hx+8,W-PR-130);
        ctx.save();
        ctx.fillStyle='rgba(3,5,14,.92)';ctx.strokeStyle=tr.color+'60';ctx.lineWidth=.8;
        ctx.beginPath();ctx.roundRect(cx2,hy-32,128,30,3);ctx.fill();ctx.stroke();
        ctx.font=`700 8.5px 'SF Mono',monospace`;ctx.fillStyle=tr.color;ctx.textAlign='left';
        ctx.fillText(fmtC(spent,currency)+' spent ('+pct+'%)',cx2+5,hy-19);
        ctx.font=`7.5px 'SF Mono',monospace`;ctx.fillStyle=tr.color+'80';
        ctx.fillText(fmtD(addM(startDate,Math.round(prog*totalMonths))),cx2+5,hy-7);
        ctx.restore();
      }
    }
  });

  /* ─ 10. TIMELINE RAILS ─ */
  acts.forEach((tr,i)=>{
    const ryi=ry(i);
    // Ambient rail (full width, faded track colour)
    const rg=ctx.createLinearGradient(PL,0,W-PR,0);
    rg.addColorStop(0,tr.color+'04');rg.addColorStop(.45,tr.color+'16');rg.addColorStop(1,tr.color+'02');
    ctx.save();ctx.strokeStyle=rg;ctx.lineWidth=.6;
    ctx.beginPath();ctx.moveTo(PL,ryi);ctx.lineTo(W+20,ryi);ctx.stroke();ctx.restore();

    if(prog>0.005){
      // GLOW TRAIL (two layers for depth)
      if(!mob){
        ctx.save();ctx.globalAlpha=.1;ctx.strokeStyle=tr.color;ctx.lineWidth=tr.lw+14;
        ctx.lineCap='round';ctx.shadowColor=tr.color;ctx.shadowBlur=20;
        ctx.beginPath();ctx.moveTo(PL,ryi);ctx.lineTo(xAt(prog,tr.schedMult),ryi);ctx.stroke();ctx.restore();
        ctx.save();ctx.globalAlpha=.18;ctx.strokeStyle=tr.color;ctx.lineWidth=tr.lw+5;
        ctx.lineCap='round';ctx.shadowColor=tr.color;ctx.shadowBlur=10;
        ctx.beginPath();ctx.moveTo(PL,ryi);ctx.lineTo(xAt(prog,tr.schedMult),ryi);ctx.stroke();ctx.restore();
      }
      // CORE PROGRESS LINE
      ctx.save();ctx.strokeStyle=tr.color;ctx.lineWidth=tr.lw;ctx.lineCap='round';ctx.setLineDash(tr.dash);
      ctx.beginPath();let s2=false;
      for(let k=0;k<=300;k++){const t=k/300;if(t>prog)break;const x=xAt(t,tr.schedMult);s2?ctx.lineTo(x,ryi):(ctx.moveTo(x,ryi),s2=true);}
      ctx.stroke();ctx.setLineDash([]);ctx.restore();

      // PARTICLE COMET TRAIL — 8 fading dots behind the head
      const hx=xAt(prog,tr.schedMult);
      if(!mob){
        for(let p=1;p<=8;p++){
          const pt=Math.max(0,prog-p*0.008);
          const px=xAt(pt,tr.schedMult);
          const pa=0.5*(1-p/9);const pr=Math.max(.5,(acts.length===1?7:5)*(1-p/9)*.5);
          ctx.save();ctx.globalAlpha=pa;ctx.fillStyle=tr.color;
          ctx.beginPath();ctx.arc(px,ryi,pr,0,Math.PI*2);ctx.fill();ctx.restore();
        }
      }

      // PULSING HEAD — dual ring
      const pulse=0.55+0.45*Math.sin(glowPhase*3+i*1.3);
      const HR=acts.length===1?(mob?6:8):(mob?4:6);
      ctx.save();ctx.shadowColor=tr.color;ctx.shadowBlur=22*pulse;
      ctx.fillStyle=tr.color+'40';ctx.beginPath();ctx.arc(hx,ryi,HR*2*pulse,0,Math.PI*2);ctx.fill();
      ctx.fillStyle=tr.color+'20';ctx.beginPath();ctx.arc(hx,ryi,HR*3.2*pulse,0,Math.PI*2);ctx.fill();
      ctx.restore();
      dot(ctx,hx,ryi,tr.color,HR,14*pulse);

      // TRACK LABEL on head
      ctx.save();ctx.font=`700 ${mob?7.5:9}px 'SF Mono',monospace`;ctx.fillStyle=tr.color;
      ctx.shadowColor=tr.color;ctx.shadowBlur=6;ctx.textAlign='left';
      ctx.fillText(mob?tr.shortL:tr.label,Math.min(hx+(mob?8:11),W-PR-70),ryi-(mob?9:13));ctx.restore();
    }

    delivFlag(ctx,tr,xAt(1,tr.schedMult),W,PT,addM(startDate,totalMonths*tr.schedMult),mob||tab);

    /* ── MILESTONES on BASE rail only ── */
    if(tr.id==='base'){
      milestones.forEach((ms,mi)=>{
        if(prog<ms.t&&ms.t>0)return;
        const mx2=xAt(ms.t,1),above=mi%2===0;
        // Pop scale based on how recently this milestone was revealed
        const revAge=milestoneReveal[mi]||0;
        const popScale=revAge>0?Math.min(1,1+(1-revAge)*0.6):1;
        const msR=mob?3.5:5;
        ctx.save();ctx.translate(mx2,ryi);ctx.scale(popScale,popScale);
        ctx.shadowColor='#10b981';ctx.shadowBlur=8;
        ctx.fillStyle='#10b981';ctx.beginPath();ctx.arc(0,0,msR,0,Math.PI*2);ctx.fill();
        ctx.strokeStyle='rgba(0,0,0,.5)';ctx.lineWidth=1;ctx.beginPath();ctx.arc(0,0,msR,0,Math.PI*2);ctx.stroke();
        ctx.restore();
        // Labels — alternate above/below, skip some on mobile
        if(!mob||mi%2===0){
          ctx.save();ctx.font=`700 ${mob?6.5:8}px 'SF Mono',monospace`;ctx.fillStyle='#10b981';ctx.textAlign='center';
          ctx.shadowColor='rgba(16,185,129,.3)';ctx.shadowBlur=4;
          ctx.fillText(ms.l,mx2,above?ryi-14:ryi+19);
          if(!mob){
            ctx.font=`6.5px 'SF Mono',monospace`;ctx.fillStyle='rgba(16,185,129,.42)';ctx.shadowBlur=0;
            ctx.fillText(fmtD(addM(startDate,Math.round(ms.t*totalMonths))),mx2,above?ryi-4:ryi+30);
          }
          ctx.restore();
        }
      });
    }

    /* ── PER-TRACK RISK DIAMONDS ── */
    const revSet=revealedMap.get(tr.id)||new Set();
    revSet.forEach(ri=>{
      const r=tr.risks[ri];if(!r)return;
      const rx2=xAt(r.t,tr.schedMult),above=ri%2===0;
      const sz=r.sev==='high'?(mob?5.5:7):(mob?3.5:5);
      // Diamond
      // Risk diamonds: HIGH=orange, CRITICAL=red, MED=track colour
      const riskFill=r.sev==='high'?'#f59e0b':r.sev==='critical'?'#ef4444':tr.color+'bb';
      const riskGlow=r.sev==='high'?'rgba(245,158,11,.6)':r.sev==='critical'?'rgba(239,68,68,.6)':'transparent';
      ctx.save();ctx.shadowColor=riskGlow;ctx.shadowBlur=r.sev==='high'||r.sev==='critical'?12:0;
      ctx.fillStyle=riskFill;
      ctx.translate(rx2,ryi);ctx.beginPath();
      ctx.moveTo(0,-sz);ctx.lineTo(sz,0);ctx.lineTo(0,sz);ctx.lineTo(-sz,0);ctx.closePath();ctx.fill();
      if(r.sev==='high'||r.sev==='critical'){
        ctx.strokeStyle=r.sev==='critical'?'#ff6b6b':'#fbbf24';ctx.lineWidth=1.2;ctx.beginPath();
        ctx.moveTo(0,-sz);ctx.lineTo(sz,0);ctx.lineTo(0,sz);ctx.lineTo(-sz,0);ctx.closePath();ctx.stroke();
      }
      ctx.restore();
      // Connector + label box (tablet and desktop)
      if(!mob){
        const cLen=36,cY=above?ryi-cLen:ryi+cLen;
        ctx.save();ctx.strokeStyle=tr.color+'50';ctx.lineWidth=.5;ctx.setLineDash([2,4]);
        ctx.beginPath();ctx.moveTo(rx2,above?ryi-sz:ryi+sz);ctx.lineTo(rx2,cY);ctx.stroke();ctx.setLineDash([]);
        const bw=tab?112:134,bh=38,bx=Math.max(PL+2,Math.min(rx2-bw/2,W-PR-bw-2));
        const by=above?cY-bh-2:cY+2;
        ctx.fillStyle='rgba(3,5,14,.93)';ctx.strokeStyle=tr.color+'55';ctx.lineWidth=.7;
        ctx.beginPath();ctx.roundRect(bx,by,bw,bh,4);ctx.fill();ctx.stroke();
        // Left colour strip
        ctx.fillStyle=r.sev==='high'?'#f59e0b':r.sev==='critical'?'#ef4444':tr.color;ctx.beginPath();ctx.roundRect(bx,by,3,bh,[3,0,0,3]);ctx.fill();
        ctx.font=`700 6px 'SF Mono',monospace`;ctx.fillStyle=tr.color+'cc';ctx.textAlign='left';
        ctx.fillText(tr.shortL.slice(0,9).toUpperCase(),bx+6,by+9);
        ctx.font=`700 ${tab?7.5:8}px 'SF Mono',monospace`;ctx.fillStyle='#d8e8f4';
        ctx.fillText(r.l.length>16?r.l.slice(0,15)+'…':r.l,bx+6,by+20);
        ctx.font=`7px 'SF Mono',monospace`;ctx.fillStyle='rgba(180,210,240,.52)';
        ctx.fillText(r.imp.slice(0,30),bx+6,by+31);
        ctx.restore();
      }
    });
  });

  /* ─ 11. NOW SCAN LINE + "NOW" LABEL ─ */
  if(prog>0.01&&prog<0.995){
    const sx=xAt(prog,1);
    const sg=ctx.createLinearGradient(sx-36,0,sx+36,0);
    sg.addColorStop(0,'rgba(141,247,255,0)');
    sg.addColorStop(.5,'rgba(141,247,255,.055)');
    sg.addColorStop(1,'rgba(141,247,255,0)');
    ctx.save();ctx.fillStyle=sg;ctx.fillRect(sx-36,PT,72,H-PT-PB);
    ctx.strokeStyle='rgba(141,247,255,.22)';ctx.lineWidth=.6;
    ctx.beginPath();ctx.moveTo(sx,PT);ctx.lineTo(sx,H-PB);ctx.stroke();
    // "NOW" label on the scan line
    ctx.font=`700 7px 'SF Mono',monospace`;ctx.fillStyle='rgba(141,247,255,.5)';ctx.textAlign='center';
    ctx.fillText('NOW',sx,H-PB-6);
    ctx.restore();
  }
}

/* ═══════════════════════════════════════════════════
   REACT COMPONENT
═══════════════════════════════════════════════════ */
export default function ProjectTimeline({model,actualProgress,initialMode}){
  const cvsRef=useRef(null),rafRef=useRef(null),lastTs=useRef(null),glowRef=useRef(0);
  const[prog,setProg]=useState(0),[playing,setPlaying]=useState(false);
  const[mode,setMode]=useState(initialMode||'base');
  useEffect(()=>{if(initialMode&&initialMode!==mode)changeMode(initialMode);},[initialMode]);
  const[speed,setSpeed]=useState(2);
  const[revMap,setRevMap]=useState(()=>new Map([['base',new Set()],['scenario',new Set()],['benchmark',new Set()],['stress',new Set()]]));
  const[log,setLog]=useState([]),[advisor,setAdvisor]=useState('');
  const[mob,setMob]=useState(()=>window.innerWidth<500);
  const[tab2,setTab2]=useState(()=>window.innerWidth<800);
  // Flash state for risk events
  const flashRef=useRef(null);
  // Milestone reveal ages (for pop animation)
  const msRevealRef=useRef({});

  useEffect(()=>{
    function onR(){setMob(window.innerWidth<500);setTab2(window.innerWidth<800);}
    window.addEventListener('resize',onR);return()=>window.removeEventListener('resize',onR);
  },[]);

  useEffect(()=>{
    if(actualProgress>0&&actualProgress<=1&&!playing){
      setProg(actualProgress);repaint(actualProgress,mode,revMap,glowRef.current);
    }
  },[actualProgress]);

  const D=useMemo(()=>{
    if(!model)return null;
    const sector=getSector(model),currency=model.currency_symbol||(getSector(model)==='space'?'$':'£');
    const totalMonths=parseMonths(model),baseCostBn=parseCostBn(model),startDate=parseStart(model);
    const p80CostBn=model.monte_carlo?.qcra?.p80||baseCostBn*1.18;
    const milestones=(()=>{if(model.schedule_detail?.length>=3)return model.schedule_detail.slice(0,8).map((s,i,a)=>({t:i===0?0:i===a.length-1?1:i/(a.length-1),l:(s.activity||s.name||s.description||`Phase ${i+1}`).slice(0,mob?12:20)}));return MS[sector]||MS.generic;})();
    const trackRisks=buildTrackRisks(model,sector),scenario=buildScenario(model),benchmark=buildBenchmark(model),stress=buildStress(model);
    return{sector,currency,totalMonths,baseCostBn,p80CostBn,startDate,milestones,trackRisks,scenario,benchmark,stress,confPct:+model.confidence_pct||60,title:model.programme_title||model.title||model.subsector||'Programme',location:model.location||''};
  },[model,mob]);

  const fresh=useCallback(()=>new Map([['base',new Set()],['scenario',new Set()],['benchmark',new Set()],['stress',new Set()]]),[]);
  const CH=useCallback(()=>canvasH(),[]);

  const initCvs=useCallback(()=>{
    const cvs=cvsRef.current;if(!cvs)return;
    const dpr=window.devicePixelRatio||1,W=cvs.offsetWidth,H=CH();
    cvs.width=W*dpr;cvs.height=H*dpr;cvs.style.height=H+'px';
    cvs.getContext('2d').scale(dpr,dpr);
  },[CH]);

  const repaint=useCallback((p,m,rv,gp=0)=>{
    const cvs=cvsRef.current;if(!cvs||!D)return;
    paintCanvas(cvs.getContext('2d'),cvs.offsetWidth,CH(),{
      prog:p,activeMode:m,...D,revealedMap:rv,glowPhase:gp,
      recentFlash:flashRef.current,milestoneReveal:msRevealRef.current,
    });
  },[D,CH]);

  useEffect(()=>{initCvs();repaint(0,mode,fresh());},[D,mode]);
  useEffect(()=>{
    const ob=new ResizeObserver(()=>{initCvs();repaint(prog,mode,revMap,glowRef.current);});
    if(cvsRef.current?.parentElement)ob.observe(cvsRef.current.parentElement);
    return()=>ob.disconnect();
  },[repaint,prog,mode,revMap]);

  // Touch scrub
  useEffect(()=>{
    const cvs=cvsRef.current;if(!cvs)return;
    function onTouch(e){
      e.preventDefault();if(!D)return;
      const rect=cvs.getBoundingClientRect();
      const tx=e.touches[0].clientX-rect.left;
      const PL=window.innerWidth<500?42:60,PR=window.innerWidth<500?8:18;
      const TW=cvs.offsetWidth-PL-PR;
      onScrub(Math.round(Math.max(0,Math.min(1,(tx-PL)/TW))*1000));
    }
    cvs.addEventListener('touchmove',onTouch,{passive:false});
    return()=>cvs.removeEventListener('touchmove',onTouch);
  },[D,mode]);

  // Auto-play
  useEffect(()=>{if(D&&!playing&&prog===0){const t=setTimeout(()=>setPlaying(true),900);return()=>clearTimeout(t);}},[D]);

  const step=useCallback(ts=>{
    if(!lastTs.current)lastTs.current=ts;
    const dt=(ts-lastTs.current)*speed;lastTs.current=ts;
    glowRef.current=(glowRef.current+dt/1000)%(Math.PI*2);

    // Decay flash
    if(flashRef.current&&flashRef.current.t>0){
      flashRef.current={...flashRef.current,t:flashRef.current.t-dt/400};
      if(flashRef.current.t<=0)flashRef.current=null;
    }
    // Decay milestone pops
    Object.keys(msRevealRef.current).forEach(k=>{
      if(msRevealRef.current[k]>0)msRevealRef.current[k]=Math.max(0,msRevealRef.current[k]-dt/600);
    });

    setProg(p=>{
      const np=Math.min(1,p+dt/14000);
      const newMap=new Map(revMap);const newLogs=[];
      const tids=['base',...(mode==='scenario'||mode==='benchmark'?['scenario']:[]),...(mode==='benchmark'?['benchmark']:[]),...(mode==='stress'?['stress']:[])];

      tids.forEach(tid=>{
        const tr=D?.trackRisks?.[tid]||[];const rs=new Set(newMap.get(tid)||[]);
        tr.forEach((r,i)=>{
          if(np>=r.t&&!rs.has(i)){
            rs.add(i);newLogs.push({...r,tid,tcol:TC[tid],tlbl:tid.toUpperCase(),idx:i});
            // Trigger flash
            flashRef.current={riskT:r.t,color:TC[tid],t:1.0};
          }
        });
        newMap.set(tid,rs);
      });

      // Milestone reveals
      D?.milestones?.forEach((ms,i)=>{
        if(np>=ms.t&&ms.t>0&&!msRevealRef.current[i]){
          msRevealRef.current[i]=1.0;// pop starts at 1.0 and decays to 0
        }
      });

      if(newLogs.length){setRevMap(newMap);setLog(prev=>[...newLogs,...prev].slice(0,28));}
      repaint(np,mode,newMap,glowRef.current);
      if(np>=1){setPlaying(false);return 1;}return np;
    });
    rafRef.current=requestAnimationFrame(step);
  },[speed,mode,D,revMap,repaint]);

  useEffect(()=>{
    if(playing){lastTs.current=null;rafRef.current=requestAnimationFrame(step);}
    else if(rafRef.current)cancelAnimationFrame(rafRef.current);
    return()=>{if(rafRef.current)cancelAnimationFrame(rafRef.current);};
  },[playing,step]);

  function reset(){setPlaying(false);setProg(0);flashRef.current=null;msRevealRef.current={};const rv=fresh();setRevMap(rv);setLog([]);repaint(0,mode,rv);}
  function changeMode(m){setMode(m);setPlaying(false);setProg(0);flashRef.current=null;msRevealRef.current={};const rv=fresh();setRevMap(rv);setLog([]);setTimeout(()=>repaint(0,m,rv),40);}
  function onScrub(v){
    setPlaying(false);const p=v/1000;const nm=new Map();const nl=[];
    const tids=['base',...(mode==='scenario'||mode==='benchmark'?['scenario']:[]),...(mode==='benchmark'?['benchmark']:[]),...(mode==='stress'?['stress']:[])];
    tids.forEach(tid=>{const tr=D?.trackRisks?.[tid]||[];const rs=new Set();tr.forEach((r,i)=>{if(r.t<=p){rs.add(i);nl.push({...r,tid,tcol:TC[tid],tlbl:tid.toUpperCase(),idx:i});}});nm.set(tid,rs);});
    setRevMap(nm);setLog(nl.sort((a,b)=>b.t-a.t).slice(0,28));setProg(p);repaint(p,mode,nm);
  }
  function fireAdvisor(){if(!advisor.trim()||!D)return;setLog(prev=>[{l:`ADVISOR: ${advisor.slice(0,40)}`,imp:'What-if scenario — see Board Room tab for full AI analysis',drv:advisor,tid:'base',tcol:'#8df7ff',tlbl:'WHAT-IF',idx:999,t:prog},...prev].slice(0,28));setAdvisor('');}

  if(!model||!D)return(
    <div style={{padding:'52px 0',textAlign:'center',color:'rgba(255,255,255,.16)',fontFamily:'system-ui',background:'#03060c',borderRadius:12}}>
      <div style={{fontSize:38,marginBottom:12,opacity:.25}}>◎</div>
      <div style={{fontSize:13,letterSpacing:'.04em'}}>Run a project to activate the programme timeline</div>
      <div style={{fontSize:10,marginTop:5,color:'rgba(255,255,255,.09)'}}>Earth demo · Space demo · Showcase library · Free project run</div>
    </div>
  );

  const{currency,totalMonths,baseCostBn,startDate,scenario,benchmark,stress,confPct,title,location}=D;
  const sm={base:1,scenario:scenario.schedMult,benchmark:benchmark.schedMult,stress:stress.schedMult}[mode]||1;
  const forecastDel=fmtD(addM(startDate,totalMonths*sm));
  const confC=confPct>=75?TC.scenario:confPct>=55?TC.benchmark:TC.stress;

  const MODES=[
    {id:'base',     lbl:'Base',        shortLbl:'Base',   col:TC.base,      del:fmtD(addM(startDate,totalMonths)),                         cost:fmtC(baseCostBn,currency),                  conf:confPct,          desc:'Contractual baseline'},
    {id:'scenario', lbl:scenario.label,shortLbl:'Faster', col:TC.scenario,  del:fmtD(addM(startDate,totalMonths*scenario.schedMult)),      cost:fmtC(baseCostBn*scenario.costMult,currency), conf:scenario.confPct, desc:'Best credible outcome'},
    {id:'benchmark',lbl:benchmark.label||'Benchmark', shortLbl:(benchmark.label||'Bench').slice(0,8), col:TC.benchmark, del:fmtD(addM(startDate,totalMonths*benchmark.schedMult)),     cost:fmtC(baseCostBn*benchmark.costMult,currency),conf:benchmark.confPct,desc:benchmark.label},
    {id:'stress',   lbl:'Stress P90',  shortLbl:'P90',    col:TC.stress,    del:fmtD(addM(startDate,totalMonths*stress.schedMult)),        cost:fmtC(baseCostBn*stress.costMult,currency),   conf:stress.confPct,   desc:'All risks materialise'},
  ];

  const S={
    wrap:{background:'#02050b',borderRadius:mob?8:12,overflow:'hidden',fontFamily:'system-ui,-apple-system,sans-serif',color:'#e2eaf6',border:'1px solid rgba(255,255,255,.06)',boxShadow:'0 28px 90px rgba(0,0,0,.75)'},
    hdr:{background:'linear-gradient(135deg,rgba(4,8,22,.97),rgba(8,18,38,.97))',borderBottom:'1px solid rgba(255,255,255,.055)',padding:mob?'10px 12px':'14px 22px',display:'flex',alignItems:'flex-start',justifyContent:'space-between',flexWrap:'wrap',gap:mob?6:12},
    cards:{display:'grid',gridTemplateColumns:mob?'1fr 1fr':tab2?'1fr 1fr':'repeat(4,1fr)',background:'rgba(2,4,12,.75)',borderBottom:'1px solid rgba(255,255,255,.05)'},
    card:{position:'relative',padding:mob?'9px 10px 9px 13px':'13px 15px 13px 18px',borderRight:'1px solid rgba(255,255,255,.05)',background:'transparent',cursor:'pointer',textAlign:'left',fontFamily:'inherit',color:'inherit',transition:'background .18s,box-shadow .18s',overflow:'hidden',border:'none'},
    ctrl:{background:'rgba(3,6,15,.97)',borderBottom:'1px solid rgba(255,255,255,.04)',padding:mob?'7px 10px':'8px 18px',display:'flex',alignItems:'center',gap:mob?5:9,flexWrap:'wrap'},
    btn:{padding:mob?'5px 10px':'5px 15px',fontSize:mob?9:10,fontWeight:700,letterSpacing:'.05em',border:'1px solid',borderRadius:4,background:'transparent',cursor:'pointer',fontFamily:'inherit',transition:'all .13s'},
    legend:{background:'rgba(2,4,12,.88)',borderTop:'1px solid rgba(255,255,255,.04)',padding:mob?'6px 10px':'8px 18px',display:'flex',gap:mob?7:15,flexWrap:'wrap',fontSize:8.5,color:'rgba(255,255,255,.3)',alignItems:'center'},
    aRow:{background:'rgba(3,6,15,.97)',borderTop:'1px solid rgba(255,255,255,.05)',padding:mob?'6px 10px':'8px 18px',display:'flex',gap:7,alignItems:'center'},
    aIn:{flex:1,background:'rgba(255,255,255,.03)',border:'1px solid rgba(255,255,255,.07)',borderRadius:4,color:'#e2eaf6',padding:'6px 11px',fontSize:mob?10:11,outline:'none',fontFamily:'inherit',transition:'border-color .15s'},
    aBtn:{padding:'6px 14px',fontSize:mob?9:10,fontWeight:700,border:'1px solid rgba(141,247,255,.22)',borderRadius:4,background:'rgba(141,247,255,.07)',color:'#8df7ff',cursor:'pointer',fontFamily:'inherit'},
    log:{borderTop:'1px solid rgba(255,255,255,.05)',background:'rgba(3,6,15,.97)'},
    logH:{padding:mob?'5px 10px':'7px 18px',fontSize:mob?7:8,color:'rgba(255,255,255,.22)',letterSpacing:'.1em',borderBottom:'1px solid rgba(255,255,255,.05)',display:'flex',alignItems:'center',gap:6,flexWrap:'wrap'},
    logL:{padding:mob?'6px 10px':'8px 18px',display:'flex',flexDirection:'column',gap:3,minHeight:54,maxHeight:mob?150:240,overflowY:'auto'},
  };

  return(
    <div style={S.wrap}>
      {/* HEADER */}
      <div style={S.hdr}>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:mob?7:8,letterSpacing:'.16em',color:'rgba(141,247,255,.38)',marginBottom:3,fontWeight:700}}>{D.sector.toUpperCase()} · {String(model.mode||'EARTH').toUpperCase()} · CASEY PROGRAMME INTELLIGENCE</div>
          <div style={{fontSize:mob?11.5:14,fontWeight:700,color:'#e2eaf6',letterSpacing:'.02em',textShadow:'0 0 20px rgba(141,247,255,.1)'}}>◎  {title}{location&&!mob?` — ${location}`:''}</div>
          <div style={{fontSize:mob?8:9,color:'rgba(255,255,255,.18)',marginTop:3}}>{fmtD(startDate)} · {totalMonths}mo · {fmtC(baseCostBn,currency)} · {confPct}% confidence</div>
        </div>
        <div style={{display:'flex',gap:mob?10:22,alignItems:'center',flexWrap:'wrap'}}>
          {(mob?[{l:'FORECAST',v:forecastDel,c:TC[mode]||TC.base},{l:'CONFIDENCE',v:confPct+'%',c:confC}]:[{l:'START',v:fmtD(startDate),c:'#8892a4'},{l:'FORECAST',v:forecastDel,c:TC[mode]||TC.base},{l:'SPEND NOW',v:fmtC(baseCostBn*scurve(prog),currency),c:TC.base},{l:'CONFIDENCE',v:confPct+'%',c:confC}]).map(k=>(
            <div key={k.l} style={{textAlign:'right'}}>
              <div style={{fontSize:mob?11:14,fontWeight:700,fontFamily:'monospace',color:k.c,letterSpacing:'.02em',textShadow:`0 0 12px ${k.c}55`}}>{k.v}</div>
              <div style={{fontSize:mob?6.5:7,color:'rgba(255,255,255,.2)',letterSpacing:'.1em',marginTop:2}}>{k.l}</div>
            </div>
          ))}
        </div>
      </div>

      {/* TRACK CARDS */}
      <div style={S.cards}>
        {MODES.map(m=>(
          <button key={m.id} onClick={()=>changeMode(m.id)}
            style={{...S.card,borderColor:mode===m.id?m.col:'rgba(255,255,255,.04)',background:mode===m.id?`${m.col}0e`:'transparent',boxShadow:mode===m.id?`inset 0 0 0 1px ${m.col}30,0 0 20px ${m.col}08`:'none'}}>
            <div style={{position:'absolute',left:0,top:0,bottom:0,width:3,borderRadius:'4px 0 0 4px',background:mode===m.id?m.col:'transparent',boxShadow:mode===m.id?`0 0 10px ${m.col}`:''}}/>
            <div style={{display:'flex',alignItems:'center',gap:5,marginBottom:4}}>
              <div style={{width:7,height:7,borderRadius:'50%',background:mode===m.id?m.col:m.col+'30',boxShadow:mode===m.id?`0 0 8px ${m.col}`:'',transition:'all .2s'}}/>
              <span style={{fontSize:mob?7:8,fontWeight:700,letterSpacing:'.09em',color:mode===m.id?m.col:'rgba(255,255,255,.2)'}}>{mob?m.shortLbl.toUpperCase():m.lbl.toUpperCase()}</span>
            </div>
            <div style={{fontSize:mob?12:15,fontWeight:700,color:mode===m.id?m.col:'rgba(255,255,255,.18)',marginBottom:2,letterSpacing:'.01em'}}>{m.del}</div>
            <div style={{fontSize:mob?9:10,fontWeight:600,color:mode===m.id?m.col+'cc':'rgba(255,255,255,.12)',marginBottom:2}}>{m.cost}</div>
            <div style={{fontSize:mob?7:8,color:m.conf>=75?TC.scenario:m.conf>=55?TC.benchmark:TC.stress,opacity:mode===m.id?1:.4}}>{m.conf}% conf.</div>
            {!mob&&<div style={{fontSize:7.5,color:'rgba(255,255,255,.15)',marginTop:5,lineHeight:1.45,borderTop:'1px solid rgba(255,255,255,.05)',paddingTop:5}}>{m.desc}</div>}
          </button>
        ))}
      </div>

      {/* CONTROLS */}
      <div style={S.ctrl}>
        <button style={{...S.btn,color:playing?'#8df7ff':'#e2eaf6',borderColor:playing?'rgba(141,247,255,.4)':'rgba(255,255,255,.12)',boxShadow:playing?'0 0 12px rgba(141,247,255,.22)':''}} onClick={()=>setPlaying(p=>!p)}>{playing?'⏸  Pause':'▶  Play'}</button>
        <button style={{...S.btn,color:'rgba(255,255,255,.2)',borderColor:'rgba(255,255,255,.06)',fontSize:8.5}} onClick={reset}>Reset</button>
        <input type="range" min={0} max={1000} step={1} value={Math.round(prog*1000)} onChange={e=>onScrub(+e.target.value)} style={{flex:1,accentColor:'#8df7ff',cursor:'pointer',minWidth:mob?50:80}}/>
        <span style={{fontSize:mob?9.5:11,fontFamily:'monospace',color:'#8df7ff',minWidth:mob?64:82,textAlign:'right',letterSpacing:'.03em',textShadow:'0 0 10px rgba(141,247,255,.45)'}}>{fmtD(addM(startDate,Math.round(prog*totalMonths)))}</span>
        {!mob&&<div style={{display:'flex',alignItems:'center',gap:4,fontSize:8.5,color:'rgba(255,255,255,.2)',marginLeft:4}}>SPD<input type="range" min={1} max={5} step={1} value={speed} onChange={e=>setSpeed(+e.target.value)} style={{width:48,accentColor:'#8df7ff'}}/><span style={{minWidth:14}}>{speed}×</span></div>}
      </div>

      {/* CANVAS */}
      <div style={{background:'#02050b',borderTop:'1px solid rgba(255,255,255,.03)',borderBottom:'1px solid rgba(255,255,255,.03)'}}>
        <canvas ref={cvsRef} style={{display:'block',width:'100%',touchAction:'none'}}/>
      </div>

      {/* LEGEND */}
      <div style={S.legend}>
        {[{c:TC.base,l:'Base',solid:true},...(mode==='scenario'||mode==='benchmark'?[{c:TC.scenario,l:mob?'Faster':scenario.label,dash:true}]:[]),...(mode==='benchmark'?[{c:TC.benchmark,l:'Benchmark',dot:true}]:[]),...(mode==='stress'?[{c:TC.stress,l:'P90',dash:true}]:[])].map((li,i)=>(
          <div key={i} style={{display:'flex',alignItems:'center',gap:4}}>
            {li.dot?<div style={{width:18,height:0,borderTop:`1.5px dotted ${li.c}`}}/>:li.dash?<div style={{width:18,height:0,borderTop:`1.5px dashed ${li.c}`}}/>:<div style={{width:18,height:2,background:li.c,borderRadius:1,boxShadow:`0 0 4px ${li.c}55`}}/>}
            <span>{li.l}</span>
          </div>
        ))}
        <div style={{width:1,height:10,background:'rgba(255,255,255,.07)',margin:'0 2px'}}/>
        <div style={{display:'flex',alignItems:'center',gap:3}}><div style={{width:7,height:7,borderRadius:'50%',background:'#10b981',boxShadow:'0 0 5px rgba(16,185,129,.5)'}}/><span>Milestone</span></div>
        <div style={{display:'flex',alignItems:'center',gap:6}}>
          <div style={{display:'flex',alignItems:'center',gap:3}}><div style={{width:7,height:7,background:'#f59e0b',transform:'rotate(45deg)',boxShadow:'0 0 5px rgba(245,158,11,.5)'}}/><span>Risk HIGH</span></div>
          <div style={{display:'flex',alignItems:'center',gap:3}}><div style={{width:7,height:7,background:'#ef4444',transform:'rotate(45deg)',boxShadow:'0 0 5px rgba(239,68,68,.5)'}}/><span>CRITICAL</span></div>
          <div style={{display:'flex',alignItems:'center',gap:3}}><div style={{width:7,height:7,background:'rgba(255,255,255,.35)',transform:'rotate(45deg)'}}/><span>MED</span></div>
        </div>
        <div style={{display:'flex',alignItems:'center',gap:3}}><div style={{width:18,height:0,borderTop:'1px dashed rgba(141,247,255,.3)'}}/><span>Predicted future</span></div>
        {!mob&&<div style={{marginLeft:'auto',fontSize:7,color:'rgba(255,255,255,.1)',letterSpacing:'.04em'}}>CASEY · PROGRAMME INTELLIGENCE · {D.sector.toUpperCase()}</div>}
      </div>

      {/* ADVISOR */}
      <div style={S.aRow}>
        {!mob&&<span style={{fontSize:7.5,fontWeight:700,letterSpacing:'.12em',color:'rgba(141,247,255,.22)',flexShrink:0}}>ADVISOR WHAT-IF</span>}
        <input value={advisor} onChange={e=>setAdvisor(e.target.value)} onKeyDown={e=>e.key==='Enter'&&fireAdvisor()}
          placeholder={mob?"Ask CASEY a what-if…":"e.g. 'What if steel costs rise 20%?' — fires a what-if event on the log"}
          style={S.aIn}/>
        <button onClick={fireAdvisor} style={S.aBtn}>{mob?'→':'Run →'}</button>
      </div>

      {/* LOG */}
      <div style={S.log}>
        <div style={S.logH}>
          <span>RISK &amp; MILESTONE INTELLIGENCE LOG</span>
          {!mob&&<span style={{fontSize:7.5,color:'rgba(255,255,255,.1)',marginLeft:5}}>Track · severity · impact · calendar date</span>}
          <span style={{marginLeft:'auto',fontSize:7,padding:'2px 8px',borderRadius:2,background:'rgba(16,185,129,.07)',color:'#10b981',border:'1px solid rgba(16,185,129,.15)',fontWeight:700,letterSpacing:'.07em'}}>CASEY MODEL</span>
        </div>
        <div style={S.logL}>
          {log.length===0
            ?<div style={{fontSize:10,color:'rgba(255,255,255,.15)',padding:'6px 0'}}>Press Play — risk and milestone events appear here, track-attributed in real time.</div>
            :log.map((r,i)=>(
              <div key={i} style={{display:'flex',gap:6,alignItems:'flex-start',padding:'5px 0',borderBottom:'1px solid rgba(255,255,255,.03)'}}>
                <div style={{width:6,height:6,borderRadius:'50%',background:r.tcol||TC.base,marginTop:4,flexShrink:0,boxShadow:`0 0 5px ${r.tcol||TC.base}`}}/>
                <div style={{fontSize:7,fontWeight:700,padding:'2px 7px',borderRadius:2,flexShrink:0,letterSpacing:'.07em',background:`${r.tcol||TC.base}12`,color:r.tcol||TC.base,border:`1px solid ${r.tcol||TC.base}28`,marginTop:1,whiteSpace:'nowrap'}}>{r.tlbl||'BASE'}</div>
                {r.sev==='high'&&!mob&&<div style={{fontSize:7,fontWeight:700,padding:'2px 6px',borderRadius:2,background:'rgba(239,68,68,.08)',color:'#ef4444',border:'1px solid rgba(239,68,68,.2)',marginTop:1,flexShrink:0}}>CRITICAL</div>}
                <div style={{flex:1,minWidth:0}}>
                  <div style={{display:'flex',gap:6,alignItems:'baseline',flexWrap:'wrap'}}>
                    <span style={{fontSize:mob?9.5:11,fontWeight:600,color:'#e2eaf6'}}>{r.l}</span>
                    {!mob&&<span style={{fontSize:8,fontFamily:'monospace',color:'rgba(255,255,255,.2)'}}>{r.t?fmtD(addM(startDate,Math.round(r.t*totalMonths))):''}</span>}
                  </div>
                  <div style={{fontSize:mob?8:9,fontFamily:'monospace',color:r.tcol||TC.base,marginTop:1,fontWeight:600}}>{r.imp}</div>
                  {!mob&&r.drv&&<div style={{fontSize:8.5,color:'rgba(255,255,255,.2)',marginTop:2}}>{r.drv}</div>}
                </div>
              </div>
            ))
          }
        </div>
      </div>
    </div>
  );
}
