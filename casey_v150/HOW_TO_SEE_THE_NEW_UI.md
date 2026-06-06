# How to see the new UI

The screenshot you sent was from an OLD version still running.
This zip has all the new features. To see them:

## Step 1 — Stop the old server
In the terminal where `npm run dev` is running:
Press Ctrl+C to stop it.

## Step 2 — Extract THIS zip
Unzip CASEY_V166_FINAL.zip to a NEW folder
(don't overwrite the old folder — extract fresh)

## Step 3 — Hard refresh the browser
After restarting, press Ctrl+Shift+R (Windows) or Cmd+Shift+R (Mac)
to clear the browser cache.

## Step 4 — Restart the dev server
cd into the new folder:
  cd casey_v150/frontend
  npm run dev

Then open localhost:5173

---

## What you will see that's new:

### Advisor tab → RIGHT PANEL (Live Client Challenge Room)
Click any of the 3 buttons (Challenge messy cost workbook / Challenge XER schedule / Challenge risk register)
You will see:
- CASEY CHALLENGE VERDICT (colour coded: red/amber/yellow)
- Live model metrics (P50, P80, confidence %, risks mapped)
- Challenge findings (click any to expand)
- RED FLAGS — Board will challenge these
- BOARD ATTACK KILL-CHAIN (sector-specific questions)
- IF THIS PROGRAMME FAILS (only when a model is loaded)
- REQUIRED NEXT ACTIONS

### Advisor tab → LEFT PANEL (Board Attack Console)
Click any button. CASEY answers IMMEDIATELY with formatted text.
Bold headings in cyan. Numbered lists. No raw JSON.
The answers come from the live model — they know what sector you ran.

### Holy Grail tab (in the tab bar)
6 buttons that mutate the live model:
- Signalling slips 4 months
- Procurement evidence gap
- Contingency cut 10%
- Operator acceptance late
- Scope growth 8%
- Political/funding pressure
Click one → watch cost, schedule, confidence and board language change.

---

IMPORTANT: You MUST stop the old server and restart with this new zip.
The old server still running will show the old UI no matter what.
