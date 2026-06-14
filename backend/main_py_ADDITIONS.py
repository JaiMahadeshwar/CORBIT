"""
CASEY main.py — ADDITIONS ONLY
================================
Add these routes to your existing main.py.
Place them after your existing /export/pdf route.

STEP 1: Add these imports near the top of main.py (after existing imports):
"""

# ── ADD TO TOP OF main.py (after existing imports) ──────────────────────────
import subprocess
import tempfile
import json
import os

# ── ADD THESE ROUTES TO main.py (after /export/pdf route) ───────────────────


# ── BOARD PACK PPTX ─────────────────────────────────────────────────────────
# This generates the 13-slide world-class PowerPoint board pack.
# Calls Node.js generate_board_pack.js which must be in the same folder as main.py.
# Run: npm install pptxgenjs   (in your backend folder on Render)

from fastapi import Request
from fastapi.responses import Response as FastAPIResponse

@app.post("/export/board-pack-pptx")
async def export_board_pack_pptx(request: Request):
    """
    Generate a 13-slide CASEY board pack PowerPoint from any model.
    Works for all demos, showcase library, free project runs.
    Any sector · any country · any currency · Earth + Space.
    """
    try:
        payload = await request.json()

        # Write model to a temp file
        with tempfile.NamedTemporaryFile(
            mode='w', suffix='.json', delete=False, encoding='utf-8'
        ) as f:
            json.dump(payload, f, ensure_ascii=False, default=str)
            model_path = f.name

        output_path = model_path.replace('.json', '.pptx')

        # Find generate_board_pack.js relative to this file
        script_path = os.path.join(os.path.dirname(__file__), 'generate_board_pack.js')
        if not os.path.exists(script_path):
            raise FileNotFoundError(
                f'generate_board_pack.js not found at {script_path}. '
                'Copy it into the backend folder alongside main.py.'
            )

        # Run Node.js generator
        result = subprocess.run(
            ['node', script_path, model_path, output_path],
            capture_output=True, text=True, timeout=60,
            cwd=os.path.dirname(script_path)
        )

        if result.returncode != 0:
            raise RuntimeError(
                f'Board pack generation failed: {result.stderr or result.stdout}'
            )

        if not os.path.exists(output_path):
            raise RuntimeError('Board pack file was not created')

        # Read and return
        with open(output_path, 'rb') as f:
            content = f.read()

        # Cleanup
        try:
            os.unlink(model_path)
            os.unlink(output_path)
        except Exception:
            pass

        title = str(payload.get('title') or payload.get('subsector') or 'CASEY')
        safe_title = ''.join(c if c.isalnum() or c in '-_ ' else '_' for c in title)[:50]

        return FastAPIResponse(
            content=content,
            media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            headers={
                'Content-Disposition': f'attachment; filename="CASEY_{safe_title}_Board_Pack.pptx"',
                'Content-Length': str(len(content)),
            }
        )

    except FileNotFoundError as e:
        return {"error": str(e), "fix": "Run: npm install pptxgenjs in your backend folder"}
    except subprocess.TimeoutExpired:
        return {"error": "Board pack generation timed out (>60s). Check Node.js is available on Render."}
    except Exception as e:
        return {"error": f"Board pack failed: {str(e)}"}


# ── WHAT-IF TIMELINE NARRATIVE (optional — requires ANTHROPIC_API_KEY) ───────
@app.post("/ai/timeline-narrative")
async def timeline_narrative(request: Request):
    """
    Returns a one-sentence board narrative for a risk event on the timeline.
    Called by useTimelineAI.js when a risk fires during animation.
    Requires ANTHROPIC_API_KEY in Render environment variables.
    """
    try:
        import anthropic
        payload = await request.json()
        prompt_text = payload.get('prompt', '')
        max_tokens = int(payload.get('max_tokens', 120))

        if not prompt_text:
            return {"text": ""}

        client = anthropic.Anthropic()  # reads ANTHROPIC_API_KEY from env
        msg = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=max_tokens,
            messages=[{"role": "user", "content": prompt_text}]
        )
        return {"text": msg.content[0].text if msg.content else ""}

    except ImportError:
        return {"text": "", "error": "anthropic package not installed. Run: pip install anthropic"}
    except Exception as e:
        return {"text": "", "error": str(e)}


"""
═══════════════════════════════════════════════════════════════════
RENDER DEPLOYMENT CHECKLIST
═══════════════════════════════════════════════════════════════════

1. BACKEND FOLDER STRUCTURE (what must be in C:/Users/jaima/897/backend/):
   
   main.py                    ← your existing backend with routes above added
   generate_board_pack.js     ← the board pack generator (from this package)
   package.json               ← so Render installs Node + pptxgenjs (see below)
   requirements.txt           ← your existing Python deps

2. CREATE package.json in backend/ folder:
   
   {
     "name": "casey-backend",
     "version": "1.0.0",
     "dependencies": {
       "pptxgenjs": "^4.0.1"
     }
   }

3. RENDER ENVIRONMENT VARIABLES (Dashboard → Service → Environment):
   
   ANTHROPIC_API_KEY = sk-ant-api03-...    (for live AI narrative — optional)
   VITE_ANTHROPIC_KEY = sk-ant-api03-...   (for frontend direct calls — optional)
   
   Node.js is available on Render Python services — no extra config needed.
   pptxgenjs installs automatically when Render sees package.json.

4. VERIFY Node.js works on Render:
   Add this temporary test route to main.py, call it once, then remove:
   
   @app.get("/test-node")
   async def test_node():
       r = subprocess.run(['node', '--version'], capture_output=True, text=True)
       return {"node": r.stdout.strip(), "ok": r.returncode == 0}

5. PUSH TO GITHUB:
   
   cd C:/Users/jaima/897/backend
   git add main.py generate_board_pack.js package.json
   git commit -m "V260 - Universal board pack PPTX + AI timeline narrative routes"
   git push
   
   → Render auto-deploys on push.

═══════════════════════════════════════════════════════════════════
TESTING LOCALLY
═══════════════════════════════════════════════════════════════════

Terminal 1 — Backend:
   cd C:/Users/jaima/897/backend
   pip install anthropic --break-system-packages    (if not already installed)
   npm install pptxgenjs
   uvicorn main:app --reload

Terminal 2 — Frontend:
   cd C:/Users/jaima/897/frontend
   npm run dev

Then open http://localhost:5173
Run Earth Demo → click Export Board Pack → 13-slide PPTX downloads.

═══════════════════════════════════════════════════════════════════
"""
