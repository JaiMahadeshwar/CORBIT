"""
CASEY main.py — COMPLETE ADDITIONS
====================================
Add all of these to your existing main.py.
Place AFTER your existing /export/pdf route.

STEP 1: Add these imports near the top of main.py:

    import subprocess, tempfile, json, os, base64
    from fastapi import Request, UploadFile, File, Form
    from fastapi.responses import Response as FastAPIResponse

STEP 2: Install Python dependencies:
    pip install anthropic pdfplumber python-pptx python-docx openpyxl --break-system-packages

STEP 3: npm install in backend folder (for pptxgenjs):
    npm install pptxgenjs

STEP 4: Set Render environment variable:
    ANTHROPIC_API_KEY = sk-ant-api03-...
"""

import subprocess, tempfile, json, os, base64
from fastapi import Request, UploadFile, File, Form
from fastapi.responses import Response as FastAPIResponse


# ════════════════════════════════════════════════════════════════
# ROUTE 1: BOARD PACK PPTX
# 13-slide world-class PowerPoint — any sector, country, currency
# ════════════════════════════════════════════════════════════════

@app.post("/export/board-pack-pptx")
async def export_board_pack_pptx(request: Request):
    try:
        payload = await request.json()
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False, encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False, default=str)
            model_path = f.name
        output_path = model_path.replace('.json', '.pptx')
        script_path = os.path.join(os.path.dirname(__file__), 'generate_board_pack.js')
        if not os.path.exists(script_path):
            return {"error": "generate_board_pack.js not found in backend folder"}
        result = subprocess.run(
            ['node', script_path, model_path, output_path],
            capture_output=True, text=True, timeout=60,
            cwd=os.path.dirname(script_path)
        )
        if result.returncode != 0:
            return {"error": f"PPTX generation failed: {result.stderr or result.stdout}"}
        if not os.path.exists(output_path):
            return {"error": "PPTX file was not created"}
        with open(output_path, 'rb') as f:
            content = f.read()
        try:
            os.unlink(model_path)
            os.unlink(output_path)
        except Exception:
            pass
        title = str(payload.get('title') or payload.get('subsector') or 'CASEY')
        safe = ''.join(c if c.isalnum() or c in '-_ ' else '_' for c in title)[:50]
        return FastAPIResponse(
            content=content,
            media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            headers={
                'Content-Disposition': f'attachment; filename="CASEY_{safe}_Board_Pack.pptx"',
                'Content-Length': str(len(content)),
            }
        )
    except subprocess.TimeoutExpired:
        return {"error": "Board pack timed out — check Node.js is available on Render"}
    except Exception as e:
        return {"error": f"Board pack failed: {str(e)}"}


# ════════════════════════════════════════════════════════════════
# ROUTE 2: DOCUMENT CHALLENGE
# Upload any board pack PDF/PPTX/DOCX — CASEY reads it and
# returns what's missing, what the board will challenge,
# and a rebuilt confidence score
# ════════════════════════════════════════════════════════════════

@app.post("/advisor/challenge-document")
async def challenge_document(request: Request):
    """
    Accepts JSON with:
      - file_b64: base64-encoded file content
      - file_name: original filename
      - file_type: MIME type
      - model: optional CASEY model context (title, cost, confidence etc)
    Returns:
      - summary: one paragraph of findings
      - critical_gaps: list of {category, finding, recommendation}
      - board_questions: list of strings
      - missing_elements: list of strings
    """
    try:
        import anthropic
        payload = await request.json()
        file_b64 = payload.get('file_b64', '')
        file_name = payload.get('file_name', 'document')
        file_type = payload.get('file_type', 'application/pdf')
        model_ctx = payload.get('model', {})

        # Extract text from document
        doc_text = ""
        if file_b64:
            file_bytes = base64.b64decode(file_b64)
            fname_lower = file_name.lower()

            if fname_lower.endswith('.pdf'):
                try:
                    import pdfplumber, io
                    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                        pages = [p.extract_text() or '' for p in pdf.pages[:20]]
                        doc_text = '\n'.join(pages)[:12000]
                except ImportError:
                    doc_text = f"[PDF: {file_name} — pdfplumber not installed]"

            elif fname_lower.endswith(('.pptx', '.ppt')):
                try:
                    from pptx import Presentation
                    import io
                    prs = Presentation(io.BytesIO(file_bytes))
                    texts = []
                    for slide in prs.slides:
                        for shape in slide.shapes:
                            if hasattr(shape, 'text'):
                                texts.append(shape.text)
                    doc_text = '\n'.join(texts)[:12000]
                except ImportError:
                    doc_text = f"[PPTX: {file_name} — python-pptx not installed]"

            elif fname_lower.endswith(('.docx', '.doc')):
                try:
                    from docx import Document
                    import io
                    doc = Document(io.BytesIO(file_bytes))
                    doc_text = '\n'.join(p.text for p in doc.paragraphs)[:12000]
                except ImportError:
                    doc_text = f"[DOCX: {file_name} — python-docx not installed]"

        # Build context string from model
        model_context = ""
        if model_ctx:
            model_context = f"""
The programme being reviewed:
- Title: {model_ctx.get('title') or model_ctx.get('subsector', 'Unknown')}
- CASEY P50 estimate: {model_ctx.get('cost_p50', 'Unknown')}
- CASEY confidence score: {model_ctx.get('confidence_pct', 'Unknown')}%
- Schedule: {model_ctx.get('schedule', 'Unknown')}
"""

        # Challenge prompt — CASEY acting as hostile board examiner
        prompt = f"""You are CASEY, the world's most rigorous programme intelligence system. 
A user has uploaded their existing board pack for you to challenge.
{model_context}

DOCUMENT CONTENT:
{doc_text if doc_text else '[No text extracted — responding based on document metadata only]'}

Your job: act as a hostile, expert investment committee examiner. 
Review this document and identify exactly what will fail in a board room.

Return ONLY valid JSON in this exact structure (no markdown, no preamble):
{{
  "summary": "Two sentences summarising the document and the overall board-readiness verdict.",
  "critical_gaps": [
    {{
      "category": "Category name (e.g. QCRA/QSRA, Governing Constraint, Benchmark, Confidence Score, Evidence Gate)",
      "finding": "What is missing or wrong in one sentence.",
      "recommendation": "What CASEY would require to fix it."
    }}
  ],
  "board_questions": [
    "Question the board will ask that this pack cannot answer"
  ],
  "missing_elements": [
    "Element missing from this pack"
  ],
  "confidence_gap": "One sentence on how far this pack is from 75% board confidence threshold"
}}

Generate 3-5 critical gaps, 5 board questions, 3-4 missing elements.
Be specific, technical and unsparing. Do not be polite."""

        client = anthropic.Anthropic()
        msg = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=2000,
            messages=[{"role": "user", "content": prompt}]
        )
        raw = msg.content[0].text.strip()
        # Strip markdown fences if present
        if raw.startswith('```'):
            raw = raw.split('```')[1]
            if raw.startswith('json'):
                raw = raw[4:]
        result = json.loads(raw)
        return result

    except ImportError:
        return {"error": "anthropic not installed — pip install anthropic"}
    except json.JSONDecodeError as e:
        return {"error": f"Failed to parse Claude response: {str(e)}", "raw": raw[:500] if 'raw' in dir() else ""}
    except Exception as e:
        return {"error": f"Challenge failed: {str(e)}"}


# ════════════════════════════════════════════════════════════════
# ROUTE 3: MONTHLY ACTUALS INGESTION
# Accepts monthly progress data → returns updated model with
# actual_progress_t (0-1) so the timeline head moves to real position
# ════════════════════════════════════════════════════════════════

@app.post("/actuals/ingest")
async def ingest_actuals(request: Request):
    """
    Accepts JSON with:
      - model: the CASEY model
      - actuals: list of {date, spend_bn, milestone, confidence, note}
    Returns:
      - updated_model: model with actual_progress_t, actual_spend_bn, actual_confidence
      - timeline_head: float 0-1 for where to place the timeline head
      - narrative: one sentence describing the actual vs planned position
    """
    try:
        payload = await request.json()
        model = payload.get('model', {})
        actuals = payload.get('actuals', [])

        if not actuals:
            return {"error": "No actuals provided", "timeline_head": 0}

        # Get the latest actual
        latest = actuals[0]  # Already sorted newest-first from frontend

        # Parse the total programme
        total_months_str = str(model.get('schedule', '') or model.get('schedule_months', '') or '24')
        total_months_match = __import__('re').search(r'\d+', total_months_str)
        total_months = int(total_months_match.group()) if total_months_match else 24

        # Get start date
        start_str = model.get('start_date', '')
        from datetime import datetime
        try:
            if start_str:
                start = datetime.fromisoformat(start_str.replace('Z', ''))
            else:
                start = datetime.now().replace(day=1)
        except Exception:
            start = datetime.now().replace(day=1)

        # Calculate months elapsed to the actual report date
        try:
            report_date_str = latest.get('date', '')
            if report_date_str:
                # Handle YYYY-MM format from month input
                if len(report_date_str) == 7:
                    report_date_str += '-01'
                report_date = datetime.fromisoformat(report_date_str)
                months_elapsed = (report_date.year - start.year) * 12 + (report_date.month - start.month)
                months_elapsed = max(0, min(months_elapsed, total_months))
            else:
                months_elapsed = 0
        except Exception:
            months_elapsed = 0

        # Calculate progress as fraction of total programme
        timeline_head = months_elapsed / max(total_months, 1)

        # Build updated model
        updated = dict(model)
        updated['actual_progress_t'] = round(timeline_head, 4)
        updated['actual_spend_bn'] = float(latest.get('spend_bn', 0))
        updated['actual_confidence'] = int(latest.get('confidence', model.get('confidence_pct', 60)))
        updated['actual_milestone'] = latest.get('milestone', '')
        updated['actual_report_date'] = latest.get('date', '')
        updated['_actualsLoaded'] = True

        # Simple narrative
        planned_pct = round(timeline_head * 100)
        spend_bn = float(latest.get('spend_bn', 0))
        p50_bn = float(model.get('cost_p50_bn') or 1)
        spend_pct = round(spend_bn / p50_bn * 100) if p50_bn > 0 else 0

        if spend_pct > planned_pct + 10:
            narrative = f"Programme is {months_elapsed} months in ({planned_pct}% of baseline) but has consumed {spend_pct}% of P50 budget — cost overrun signal."
        elif spend_pct < planned_pct - 10:
            narrative = f"Programme is {months_elapsed} months in ({planned_pct}% of baseline) with {spend_pct}% budget consumed — under-spend may indicate programme delay, not efficiency."
        else:
            narrative = f"Programme is {months_elapsed} months in ({planned_pct}% of baseline). Spend and schedule appear broadly aligned at {spend_pct}% of P50 budget consumed."

        return {
            "updated_model": updated,
            "timeline_head": timeline_head,
            "months_elapsed": months_elapsed,
            "narrative": narrative,
            "latest_actual": latest,
        }

    except Exception as e:
        return {"error": f"Actuals ingestion failed: {str(e)}", "timeline_head": 0}


# ════════════════════════════════════════════════════════════════
# ROUTE 4: WORKBOOK WITH COVER TAB + RISK HEATMAP
# Returns XLSX with a proper cover tab and risk heatmap tab
# ════════════════════════════════════════════════════════════════

@app.post("/export/workbook-with-cover")
async def export_workbook_with_cover(request: Request):
    """
    Generates an XLSX workbook with:
      Tab 1 — Cover (programme name, P50/P80/P90, confidence, RAG, date, CASEY branding)
      Tab 2 — Risk Heatmap (5×5 grid with risks plotted by probability × impact)
      Tab 3+ — Existing cost model data
    """
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        payload = await request.json()
        model = payload
        curr = model.get('currency_symbol', '£')

        wb = openpyxl.Workbook()

        # ── COVER TAB ──────────────────────────────────────────────────
        ws_cover = wb.active
        ws_cover.title = "Cover"

        # Colours
        BG_DARK  = "03060C"  # CASEY dark background
        CYAN     = "8DF7FF"
        AMBER    = "F59E0B"
        GREEN    = "10B981"
        RED      = "EF4444"
        WHITE    = "FFFFFF"
        GREY     = "94A3B8"
        MIDGREY  = "1E293B"

        conf = int(model.get('confidence_pct', 60))
        rag = "GREEN" if conf >= 75 else "AMBER" if conf >= 55 else "RED"
        rag_color = GREEN if conf >= 75 else AMBER if conf >= 55 else RED
        rag_verdict = "Ready for capital commitment" if conf >= 75 else f"Conditional — {75-conf}pts below 75% board threshold" if conf >= 55 else "Do not approve without recovery plan"

        def cell(ws, row, col, value, bold=False, size=11, color=WHITE, bg=None, align='left', wrap=False):
            c = ws.cell(row=row, column=col, value=value)
            c.font = Font(name='Calibri', bold=bold, size=size, color=color)
            c.alignment = Alignment(horizontal=align, vertical='center', wrap_text=wrap)
            if bg:
                c.fill = PatternFill("solid", fgColor=bg)
            return c

        # Set column widths
        ws_cover.column_dimensions['A'].width = 28
        ws_cover.column_dimensions['B'].width = 36
        ws_cover.column_dimensions['C'].width = 22
        ws_cover.column_dimensions['D'].width = 18
        for row in range(1, 40):
            ws_cover.row_dimensions[row].height = 16

        # Header
        ws_cover.row_dimensions[1].height = 48
        cell(ws_cover, 1, 1, "CASEY PROGRAMME INTELLIGENCE", bold=True, size=16, color=CYAN, bg=BG_DARK, align='center')
        ws_cover.merge_cells('A1:D1')

        ws_cover.row_dimensions[2].height = 28
        title = str(model.get('title') or model.get('subsector') or 'Programme')
        cell(ws_cover, 2, 1, title, bold=True, size=13, color=WHITE, bg=MIDGREY, align='center')
        ws_cover.merge_cells('A2:D2')

        # RAG verdict
        ws_cover.row_dimensions[3].height = 32
        cell(ws_cover, 3, 1, rag + " — " + rag_verdict, bold=True, size=12, color=BG_DARK, bg=rag_color, align='center')
        ws_cover.merge_cells('A3:D3')

        # Key metrics
        row = 5
        metrics = [
            ("P50 Cost (Baseline)", str(model.get('cost_p50', '—'))),
            ("P80 Cost (Board approve at)", str(model.get('cost_p80', '—'))),
            ("P90 Cost (Stress case)", str(model.get('cost_p90', '—'))),
            ("Schedule", str(model.get('schedule', '—'))),
            ("Confidence Score", f"{conf}%"),
            ("Board Threshold", "75%"),
            ("Gap to Threshold", f"{max(0, 75-conf)} pts" if conf < 75 else "None — board ready"),
            ("Estimate Class", str(model.get('estimate_class_name', f"Class {model.get('estimate_class', 3)}"))),
            ("Schedule Level", str(model.get('schedule_level_name', f"Level {model.get('schedule_level', 4)}"))),
            ("Scenario Active", str(model.get('scenario_label', 'Base'))),
            ("Location", str(model.get('location', '—'))),
            ("Sector", str(model.get('subsector', model.get('mode', '—')))),
            ("Report Date", __import__('datetime').datetime.now().strftime('%d %b %Y')),
            ("Prepared by", "CASEY Programme Intelligence"),
            ("Governing Constraint", str(model.get('governing_constraint_prominent', '—'))[:80]),
        ]
        for label, value in metrics:
            cell(ws_cover, row, 1, label, bold=True, size=10, color=GREY, bg=BG_DARK)
            cell(ws_cover, row, 2, value, size=11, color=WHITE, bg=MIDGREY, wrap=True)
            row += 1

        # ── RISK HEATMAP TAB ──────────────────────────────────────────
        ws_heat = wb.create_sheet("Risk Heatmap")
        ws_heat.column_dimensions['A'].width = 6
        for col in ['B','C','D','E','F']:
            ws_heat.column_dimensions[col].width = 20
        for r in range(1, 30):
            ws_heat.row_dimensions[r].height = 18

        PROB_LABELS  = ['Almost Certain', 'Likely', 'Possible', 'Unlikely', 'Rare']
        IMPACT_LABELS= ['Negligible', 'Minor', 'Moderate', 'Major', 'Critical']
        CELL_COLORS  = [  # row=prob (high to low), col=impact (low to high)
            ["C6EFCE","C6EFCE","FFEB9C","FFC7CE","FFC7CE"],  # Almost certain
            ["C6EFCE","C6EFCE","FFEB9C","FFC7CE","FFC7CE"],  # Likely
            ["C6EFCE","FFEB9C","FFEB9C","FFC7CE","FFC7CE"],  # Possible
            ["C6EFCE","C6EFCE","FFEB9C","FFEB9C","FFC7CE"],  # Unlikely
            ["C6EFCE","C6EFCE","C6EFCE","FFEB9C","FFEB9C"],  # Rare
        ]

        # Header row
        cell(ws_heat, 1, 1, "Risk Heatmap", bold=True, size=12, color=WHITE, bg=BG_DARK)
        ws_heat.merge_cells('A1:F1')
        for ci, il in enumerate(IMPACT_LABELS):
            cell(ws_heat, 2, ci+2, il, bold=True, size=9, color=WHITE, bg=MIDGREY, align='center')

        # Grid
        for ri, pl in enumerate(PROB_LABELS):
            cell(ws_heat, ri+3, 1, pl, bold=True, size=8, color=WHITE, bg=MIDGREY, align='right')
            for ci in range(5):
                bg_hex = CELL_COLORS[ri][ci]
                ws_heat.cell(row=ri+3, column=ci+2).fill = PatternFill("solid", fgColor=bg_hex)

        # Plot risks on heatmap
        def get_prob_idx(prob_str):
            p = str(prob_str).lower()
            if any(x in p for x in ['almost','very high','certain']): return 0
            if any(x in p for x in ['high','likely']): return 1
            if any(x in p for x in ['medium','moderate','possible']): return 2
            if any(x in p for x in ['low','unlikely']): return 3
            return 4

        def get_impact_idx(imp_str):
            im = str(imp_str).lower()
            if any(x in im for x in ['critical','catastrophic']): return 4
            if any(x in im for x in ['major','significant','high']): return 3
            if any(x in im for x in ['moderate','medium']): return 2
            if any(x in im for x in ['minor','low']): return 1
            return 0

        risks = model.get('risks', model.get('risk_register', []))
        cell_texts = [[[] for _ in range(5)] for _ in range(5)]
        for risk in risks:
            pi = get_prob_idx(risk.get('probability', ''))
            ii = get_impact_idx(risk.get('impact', risk.get('consequence', '')))
            title_r = str(risk.get('title', risk.get('risk', '?')))[:20]
            cell_texts[pi][ii].append(title_r)

        for ri in range(5):
            for ci in range(5):
                if cell_texts[ri][ci]:
                    existing = ws_heat.cell(row=ri+3, column=ci+2).value or ''
                    ws_heat.cell(row=ri+3, column=ci+2).value = '\n'.join(cell_texts[ri][ci])
                    ws_heat.cell(row=ri+3, column=ci+2).font = Font(size=7, bold=True)
                    ws_heat.cell(row=ri+3, column=ci+2).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    ws_heat.row_dimensions[ri+3].height = max(18, 14*len(cell_texts[ri][ci]))

        # ── RISK REGISTER TAB ──────────────────────────────────────────
        ws_risk = wb.create_sheet("Risk Register")
        headers = ['#','Risk Title','Probability','Impact','Schedule (wk)','Cost EMV (Bn)','Owner','Mitigation','Trigger','RAG']
        for ci, h in enumerate(headers, 1):
            c = ws_risk.cell(row=1, column=ci, value=h)
            c.font = Font(bold=True, size=9, color=WHITE)
            c.fill = PatternFill("solid", fgColor=BG_DARK)
            c.alignment = Alignment(horizontal='center')

        for ri, risk in enumerate(risks, 2):
            conf_r = str(risk.get('probability','')).lower()
            rag_r = "RED" if any(x in conf_r for x in ['high','likely','almost','critical']) else "AMBER" if any(x in conf_r for x in ['medium','moderate','possible']) else "GREEN"
            rag_c_r = FFC7CE if rag_r=="RED" else FFEB9C if rag_r=="AMBER" else C6EFCE
            row_vals = [
                ri-1,
                str(risk.get('title',risk.get('risk','?')))[:50],
                str(risk.get('probability','?')),
                str(risk.get('impact','?')),
                str(risk.get('schedule_impact_weeks','?')),
                str(risk.get('cost_emv_bn',risk.get('cost_impact_bn','?'))),
                str(risk.get('owner','?')),
                str(risk.get('mitigation','?'))[:80],
                str(risk.get('trigger','?'))[:60],
                rag_r,
            ]
            for ci, val in enumerate(row_vals, 1):
                c = ws_risk.cell(row=ri, column=ci, value=val)
                c.font = Font(size=9)
                c.alignment = Alignment(wrap_text=True, vertical='top')
                if ci == len(row_vals):
                    c.fill = PatternFill("solid", fgColor=rag_c_r if rag_r else "FFFFFF")

        for col in range(1, len(headers)+1):
            ws_risk.column_dimensions[get_column_letter(col)].width = 16

        # Save
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
            xlsx_path = f.name
        wb.save(xlsx_path)
        with open(xlsx_path, 'rb') as f:
            content = f.read()
        try:
            os.unlink(xlsx_path)
        except Exception:
            pass

        title_s = str(model.get('title') or model.get('subsector') or 'CASEY')
        safe = ''.join(c if c.isalnum() or c in '-_ ' else '_' for c in title_s)[:50]

        return FastAPIResponse(
            content=content,
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers={
                'Content-Disposition': f'attachment; filename="CASEY_{safe}_Workbook.xlsx"',
                'Content-Length': str(len(content)),
            }
        )

    except ImportError:
        return {"error": "openpyxl not installed — pip install openpyxl"}
    except Exception as e:
        import traceback
        return {"error": f"Workbook failed: {str(e)}", "trace": traceback.format_exc()[-500:]}


# ════════════════════════════════════════════════════════════════
# ROUTE 5: AI TIMELINE NARRATIVE (optional)
# Called by the advisor what-if bar on the timeline
# ════════════════════════════════════════════════════════════════

@app.post("/ai/timeline-narrative")
async def timeline_narrative(request: Request):
    try:
        import anthropic
        payload = await request.json()
        prompt_text = payload.get('prompt', '')
        if not prompt_text:
            return {"text": ""}
        client = anthropic.Anthropic()
        msg = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=150,
            messages=[{"role": "user", "content": prompt_text}]
        )
        return {"text": msg.content[0].text if msg.content else ""}
    except Exception as e:
        return {"text": "", "error": str(e)}


# ════════════════════════════════════════════════════════════════
# ROUTE 6: TEST NODE (verify pptxgenjs available — run once then remove)
# ════════════════════════════════════════════════════════════════

# @app.get("/test-node")
# async def test_node():
#     r = subprocess.run(['node', '--version'], capture_output=True, text=True)
#     r2 = subprocess.run(['node', '-e', 'require("pptxgenjs"); console.log("pptxgenjs OK")'], capture_output=True, text=True, cwd=os.path.dirname(__file__))
#     return {"node": r.stdout.strip(), "pptxgenjs": r2.stdout.strip() or r2.stderr.strip(), "ok": r.returncode == 0}


# ════════════════════════════════════════════════════════════════
# ROUTE 7: ADVISOR MEMORY — server-side storage
# Stores advisor conversations server-side, keyed by email + programme
# Called by the CASEY_Features useAdvisorMemory hook
# Requires: pip install anthropic (already done)
# ════════════════════════════════════════════════════════════════

import hashlib
from pathlib import Path

MEMORY_DIR = Path("/tmp/casey_memories")
MEMORY_DIR.mkdir(exist_ok=True)

def memory_path(email: str, programme_id: str) -> Path:
    key = hashlib.sha256(f"{email}:{programme_id}".encode()).hexdigest()[:16]
    return MEMORY_DIR / f"mem_{key}.json"

@app.post("/advisor/memory/save")
async def save_advisor_memory(request: Request):
    try:
        payload = await request.json()
        email = payload.get('email', 'anon')
        programme_id = payload.get('programme_id', 'global')
        messages = payload.get('messages', [])
        path = memory_path(email, programme_id)
        data = {"email": email, "programme_id": programme_id, "messages": messages[:40], "updated": __import__('datetime').datetime.now().isoformat()}
        path.write_text(json.dumps(data))
        return {"saved": True, "count": len(messages)}
    except Exception as e:
        return {"saved": False, "error": str(e)}

@app.post("/advisor/memory/load")
async def load_advisor_memory(request: Request):
    try:
        payload = await request.json()
        email = payload.get('email', 'anon')
        programme_id = payload.get('programme_id', 'global')
        path = memory_path(email, programme_id)
        if not path.exists():
            return {"messages": [], "found": False}
        data = json.loads(path.read_text())
        return {"messages": data.get("messages", []), "found": True, "updated": data.get("updated")}
    except Exception as e:
        return {"messages": [], "found": False, "error": str(e)}


# ════════════════════════════════════════════════════════════════
# ROUTE 8: VERSION HISTORY — server-side audit trail
# Stores every model version server-side, keyed by programme id
# This closes T&T's counter-argument about audit trails
# ════════════════════════════════════════════════════════════════

VERSIONS_DIR = Path("/tmp/casey_versions")
VERSIONS_DIR.mkdir(exist_ok=True)

@app.post("/versions/save")
async def save_version(request: Request):
    try:
        payload = await request.json()
        programme_id = str(payload.get('programme_id', 'unknown'))
        event = payload.get('event', 'Model updated')
        model_snapshot = payload.get('model', {})
        safe_id = ''.join(c if c.isalnum() or c in '-_' else '_' for c in programme_id)[:40]
        path = VERSIONS_DIR / f"v_{safe_id}.json"
        existing = []
        if path.exists():
            try: existing = json.loads(path.read_text())
            except: existing = []
        entry = {
            "id": __import__('time').time_ns(),
            "timestamp": __import__('datetime').datetime.now().isoformat(),
            "event": event,
            "confidence_pct": model_snapshot.get('confidence_pct'),
            "cost_p50": model_snapshot.get('cost_p50'),
            "cost_p80": model_snapshot.get('cost_p80'),
            "schedule": model_snapshot.get('schedule'),
            "scenario": model_snapshot.get('scenario_label') or model_snapshot.get('scenario', 'Base'),
            "estimate_class": model_snapshot.get('estimate_class_name') or f"Class {model_snapshot.get('estimate_class', 3)}",
            "programme_id": programme_id,
        }
        versions = [entry, *existing][:40]
        path.write_text(json.dumps(versions))
        return {"saved": True, "total": len(versions)}
    except Exception as e:
        return {"saved": False, "error": str(e)}

@app.get("/versions/{programme_id}")
async def get_versions(programme_id: str):
    try:
        safe_id = ''.join(c if c.isalnum() or c in '-_' else '_' for c in programme_id)[:40]
        path = VERSIONS_DIR / f"v_{safe_id}.json"
        if not path.exists():
            return {"versions": [], "found": False}
        versions = json.loads(path.read_text())
        return {"versions": versions, "found": True, "count": len(versions)}
    except Exception as e:
        return {"versions": [], "found": False, "error": str(e)}


# ════════════════════════════════════════════════════════════════
# ROUTE 9: PORTFOLIO — server-side saved programmes
# Backs up saved projects to server so they survive browser clears
# ════════════════════════════════════════════════════════════════

PORTFOLIO_DIR = Path("/tmp/casey_portfolios")
PORTFOLIO_DIR.mkdir(exist_ok=True)

@app.post("/portfolio/save")
async def save_portfolio(request: Request):
    try:
        payload = await request.json()
        email = str(payload.get('email', 'anon'))
        projects = payload.get('projects', [])
        safe = hashlib.sha256(email.encode()).hexdigest()[:16]
        path = PORTFOLIO_DIR / f"p_{safe}.json"
        data = {"email": email, "projects": projects[:30], "updated": __import__('datetime').datetime.now().isoformat()}
        path.write_text(json.dumps(data, default=str))
        return {"saved": True, "count": len(projects)}
    except Exception as e:
        return {"saved": False, "error": str(e)}

@app.post("/portfolio/load")
async def load_portfolio(request: Request):
    try:
        payload = await request.json()
        email = str(payload.get('email', 'anon'))
        safe = hashlib.sha256(email.encode()).hexdigest()[:16]
        path = PORTFOLIO_DIR / f"p_{safe}.json"
        if not path.exists():
            return {"projects": [], "found": False}
        data = json.loads(path.read_text())
        return {"projects": data.get("projects", []), "found": True, "updated": data.get("updated")}
    except Exception as e:
        return {"projects": [], "found": False, "error": str(e)}


"""
════════════════════════════════════════════════════════════════
FINAL RENDER DEPLOYMENT CHECKLIST
════════════════════════════════════════════════════════════════

You have the £20/month Render plan — backend stays awake 24/7.
No more cold start. Every demo opens instantly.

1. COPY FILES INTO backend/ FOLDER:
   main_py_COMPLETE.py  → add all routes to main.py
   generate_board_pack.js  → stays in backend/
   package.json (from backend_package.json)  → stays in backend/

2. INSTALL DEPENDENCIES:
   pip install anthropic pdfplumber python-pptx python-docx openpyxl --break-system-packages
   npm install pptxgenjs

3. RENDER ENVIRONMENT VARIABLES (Dashboard → Service → Environment):
   ANTHROPIC_API_KEY = sk-ant-api03-...
   VITE_ANTHROPIC_KEY = same key
   NODE_VERSION = 18

4. GIT PUSH:
   cd C:/Users/jaima/897/backend
   git add main.py generate_board_pack.js package.json
   git commit -m "V290 — complete platform: board pack, challenge, actuals, workbook, memory, versions, portfolio"
   git push

   cd C:/Users/jaima/897/frontend
   git add src/App.jsx src/ProjectTimeline.jsx src/CASEY_Upgrades.jsx src/CASEY_Features.jsx src/useTimelineAI.js
   git commit -m "V290 — typewriter verdict, count-up confidence, rotating challenge, version history, portfolio strip, actualProgress wired"
   git push

════════════════════════════════════════════════════════════════
"""
