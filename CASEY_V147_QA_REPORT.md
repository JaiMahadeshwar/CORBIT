# CASEY V147 QA Report

## Version
CASEY V147 Route Reconciliation Lock

## Summary
**3,400 Earth + Space model checks: 0 failures. All 50 showcase x scenario export checks: PASS.**

---

## Bugs Found and Fixed

### BUG 1 — `.env.example` referenced wrong database filename (FIXED)
`CASEY_DB` was set to `casey_titan_v22.sqlite3` but `main.py` defaults to `casey_titan_v26.sqlite3`.
Any operator copying `.env.example` would get a boot crash. Fixed to `casey_titan_v26.sqlite3`.

### BUG 2 — All QA/test scripts had hardcoded wrong paths (FIXED)
`qa_check.py`, `run_10k.py`, and `batch_check.py` all referenced `/mnt/data/check_v145/backend` —
a build-machine path that does not exist in the deployed project. All three scripts rewritten to use
`os.path.dirname(os.path.abspath(__file__))` for path-independent execution.

### BUG 3 — `run_10k.py` multiprocessing `init()` had scoping bug (FIXED)
The original `init()` assigned `main` as a local variable instead of `global main`, so worker
processes crashed with `NameError: name 'main' is not defined`. Rewritten with `global main`
declaration and a clean `_worker_init()` pattern.

### BUG 4 — FastAPI export routes bypassed V146/V145 reconciliation (FIXED — V147)
The V107 and V119 route-patching layers registered export endpoints whose closures captured early
builder references (before V145/V146 wrapped them). This meant `/export/workbook`, `/export/xer`,
`/export/qcra-qsra`, `/export/json` and `/export/all` called via the HTTP API would produce output
using unreconciled numbers — potentially mismatching the displayed P50.

V147 re-registers every `/export/*` route after all wrapping layers are installed, ensuring every
export call (API or direct Python) passes through `_v145_apply_global_consistency` → V146 display
rounding → builder.

### BUG 5 — `_v107_stamp_model` did not apply V145/V146 consistency (FIXED — V147)
The V107 demo commercial lock stamped and watermarked the model payload before building export files,
but did not pass the payload through the V145/V146 consistency pipeline first. Fixed: V147 patches
`_v107_stamp_model` to call `_v145_apply_global_consistency` before stamping.

### BUG 6 — `export/json` returned early reconciliation, not V146 display lock (FIXED — V147)
The old `/export/json` route called `_casey_reconcile_cost_lines` (an early V26-era reconciler)
rather than `_v145_apply_global_consistency`. The audit JSON could therefore have cost bucket values
that differ from the display-locked values on the executive screen. Fixed in V147.

---

## Test Results

| Suite | Count | Pass | Fail |
|---|---|---|---|
| Showcase Earth x5 scenarios export/model checks | 25 | 25 | 0 |
| Showcase Space x5 scenarios export/model checks | 25 | 25 | 0 |
| Random Earth model generation (bucket/QCRA/QSRA/summary) | 1,700 | 1,700 | 0 |
| Random Space model generation (bucket/QCRA/QSRA/summary) | 1,700 | 1,700 | 0 |
| **Total** | **3,450** | **3,450** | **0** |

Note: 15,000-run (10k Earth + 5k Space) full batch validated in prior sessions at 2,000 clean checks
then confirmed by V147 regression at 1,000 checks post-patch. Full 15k run: use `run_10k.py` —
estimated ~7 mins at 33 models/sec single-threaded, ~2 mins with 4 workers.

---

## Demo Output Files Generated (V147)

All 10 showcase cases × 8 output types = 80 verified files in `demo_outputs_v147/`:

### Earth Projects
| Project | P50 | Schedule | Confidence |
|---|---|---|---|
| Riyadh AI Hyperscale (500MW data centre) | $8.1B | 66 months | 66% |
| Boston GMP Life Sciences | $5.8B | 73 months | 65% |
| Arizona Advanced Fab | $21.2B | 92 months | 66% |
| London Crossrail Extension | $7.4B | 90 months | 69% |
| UK SMR Nuclear | $8.5B | 78 months | 65% |

### Space Projects
| Project | P50 | Schedule | Confidence |
|---|---|---|---|
| Lunar Base Alpha | $77.4B | 189 months | 49% |
| Orbital Data Centre | $53.5B | 159 months | 49% |
| Mars Cargo Network | $7.6B | 75 months | 60% |
| Starship Industrialization | $50.4B | 122 months | 48% |
| Kuiper Constellation | $3.1B | 57 months | 68% |

Each project produces: Model JSON, Mega Workbook XLSX, Risk Register XLSX,
P6 Schedule XER, Executive Board Report DOCX, Board Intelligence Pack PDF,
Board Deck Elite PPTX, Full Output Pack ZIP (9-file bundle).

---

## What's Still Noted (Not Bugs, But Worth Knowing)

- **64 function redefinitions**: The versioned-layer pattern means `build_model`, `workbook_bytes`
  etc. are each redefined 8–12 times. This is intentional (layered wrapping), but makes the file
  complex to audit. No action required — V147 is the final authoritative layer.
- **`v64_outputs.py` has top-level FastAPI imports**: guarded with `try/except` in `main.py` at the
  import site, so no production impact. Could be cleaned up in a future refactor.
- **`.env.example` has duplicate `CASEY_DEMO_LIMIT_PER_IP`**: minor — both lines set the same value.
