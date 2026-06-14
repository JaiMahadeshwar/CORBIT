/**
 * useTimelineAI.js — Claude API hook for ProjectTimeline
 *
 * Adds two AI-powered capabilities:
 *   1. getRiskNarrative(risk) — when a risk fires on the timeline, Claude writes
 *      a one-sentence board-ready explanation of what it means RIGHT NOW
 *      for this specific programme and location.
 *
 *   2. runWhatIf(scenarioLabel, model) — when a what-if button is pressed
 *      (or triggered from the Advisor tab), Claude returns a 3-sentence
 *      board impact statement: delivery date change, cost impact, critical risk.
 *
 * Requirements:
 *   - ANTHROPIC_API_KEY must be set in your Render environment variables.
 *   - The fetch calls go directly to api.anthropic.com from the browser.
 *     If you want to keep the key server-side, proxy through your backend:
 *     replace the fetch URL with `${PROD_URL}/ai/narrative` and add a
 *     FastAPI route that forwards to Anthropic.
 *
 * Usage inside ProjectTimeline.jsx:
 *   import useTimelineAI from './useTimelineAI';
 *   const { narrative, whatIfText, whatIfLoading, getRiskNarrative, runWhatIf } = useTimelineAI(model);
 */

import { useState, useCallback, useRef } from 'react';

const CLAUDE_MODEL = 'claude-sonnet-4-6';
const ANTHROPIC_URL = 'https://api.anthropic.com/v1/messages';

// ─── IMPORTANT: API key handling ─────────────────────────────────────────────
// Option A (simplest, dev only): inject via Vite env
//   VITE_ANTHROPIC_KEY=sk-ant-... in your .env.local
//   const API_KEY = import.meta.env.VITE_ANTHROPIC_KEY;
//
// Option B (production, recommended): proxy through your backend
//   Replace ANTHROPIC_URL with `${window._CASEY_API}/ai/timeline`
//   and add a FastAPI route that forwards with server-side key.
//   This keeps the key off the client entirely.
//
// The hook works without a key — it just returns empty strings silently.
// ─────────────────────────────────────────────────────────────────────────────

const API_KEY = import.meta.env.VITE_ANTHROPIC_KEY || '';

async function callClaude(prompt, maxTokens = 120) {
  if (!API_KEY) return '';
  try {
    const res = await fetch(ANTHROPIC_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': API_KEY,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true',
      },
      body: JSON.stringify({
        model: CLAUDE_MODEL,
        max_tokens: maxTokens,
        messages: [{ role: 'user', content: prompt }],
      }),
    });
    if (!res.ok) return '';
    const d = await res.json();
    return d.content?.[0]?.text?.trim() || '';
  } catch {
    return '';
  }
}

export default function useTimelineAI(model) {
  // narrative: { [riskTitle]: string } — populated as risks fire
  const [narrative, setNarrative]       = useState({});
  // whatIfText: plain-English board impact of most recent what-if
  const [whatIfText, setWhatIfText]     = useState('');
  const [whatIfLoading, setWhatIfLoading] = useState(false);
  // prevent duplicate calls for the same risk
  const inflight = useRef(new Set());

  const sector   = model?.subsector || model?.mode || 'infrastructure';
  const location = model?.location  || model?.region || 'this location';
  const cost     = model?.cost_p50  || model?.cost_p50_bn + 'B' || 'undisclosed';
  const months   = model?.schedule_months || model?.duration_months || '';
  const conf     = model?.confidence_pct || '';

  /**
   * getRiskNarrative(risk)
   * Call this when the animation reaches a risk event.
   * risk must have: { label, impact, driver? }
   * Result appears in narrative[risk.label] — React re-renders automatically.
   */
  const getRiskNarrative = useCallback(async (risk) => {
    const key = risk.label || risk.title || '';
    if (!key || narrative[key] || inflight.current.has(key)) return;
    inflight.current.add(key);

    const prompt = `You are a programme controls expert advising a board.
In ONE sentence (max 28 words), explain what the risk "${key}" means right now
for a ${sector} programme in ${location}.
Risk impact stated: ${risk.impact || risk.imp || 'not specified'}.
${risk.driver || risk.drv ? `Root cause: ${risk.driver || risk.drv}.` : ''}
Be specific, quantified where possible, no hedging. Start with the consequence.`;

    const text = await callClaude(prompt, 80);
    if (text) setNarrative(n => ({ ...n, [key]: text }));
    inflight.current.delete(key);
  }, [narrative, sector, location]);

  /**
   * runWhatIf(scenarioLabel)
   * Call this when a what-if button is pressed or when the Advisor tab
   * sends a what-if signal. scenarioLabel is a human-readable string like
   * "+15% scope increase" or "governing constraint +12 months".
   * Result appears in whatIfText.
   */
  const runWhatIf = useCallback(async (scenarioLabel) => {
    if (!scenarioLabel) return;
    setWhatIfLoading(true);
    setWhatIfText('');

    const prompt = `You are a programme controls advisor. Be direct — board audience.
Programme: ${sector} in ${location}.
Baseline: cost ${cost}, ${months ? months + ' months, ' : ''}${conf ? conf + '% confidence.' : ''}
What-if scenario applied: "${scenarioLabel}".

In exactly 3 sentences:
1. How does the forecast delivery date change and by how much?
2. What is the cost impact in absolute terms?
3. Which single risk becomes most critical as a result, and what must the board do about it?

No preamble, no hedging, no "it depends". Numbers and actions only.`;

    const text = await callClaude(prompt, 200);
    setWhatIfText(text);
    setWhatIfLoading(false);
  }, [sector, location, cost, months, conf]);

  /**
   * clearWhatIf()
   * Call when resetting the animation or switching projects.
   */
  const clearWhatIf = useCallback(() => {
    setWhatIfText('');
  }, []);

  return {
    narrative,       // { [riskTitle]: string }
    whatIfText,      // string — 3-sentence board impact
    whatIfLoading,   // boolean
    getRiskNarrative,
    runWhatIf,
    clearWhatIf,
  };
}

// ─── BACKEND PROXY ROUTE (FastAPI) ───────────────────────────────────────────
// If you prefer to keep the API key server-side, add this to your main.py
// and replace callClaude's URL with `${window._CASEY_API}/ai/timeline-narrative`
//
// @app.post("/ai/timeline-narrative")
// async def timeline_narrative(req: dict):
//     import anthropic
//     client = anthropic.Anthropic()  # reads ANTHROPIC_API_KEY from env
//     msg = client.messages.create(
//         model="claude-sonnet-4-6",
//         max_tokens=req.get("max_tokens", 120),
//         messages=[{"role": "user", "content": req["prompt"]}]
//     )
//     return {"text": msg.content[0].text}
// ─────────────────────────────────────────────────────────────────────────────
