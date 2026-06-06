CONTROLORBIT RENDER FIX

What was fixed:
- React production crash caused by rendering an object with keys {decision_rule, primary_constraint, plain_english}.
- Scenario vs Base now uses safeRender(baseVs.plain_english).
- safeRender now converts CASEY narrative objects into readable prose instead of raw JSON/code.
- npm run build passes successfully.

Deploy steps from Windows PowerShell:
1. Extract this zip.
2. Copy/replace the contents into C:\Users\jaima\897\frontend.
3. Run:
   npm run build
   git add .
   git commit -m "Fix ControlOrbit production render crash"
   git push
4. Wait for Vercel deployment to show Ready.
5. Open https://www.controlorbit.com and press Ctrl+F5.
