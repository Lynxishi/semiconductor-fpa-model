# Validation summary - Texas Instruments (TXN)

Run: 2026-09-23T01:53:17+00:00 UTC  |  Division: Power, Signal & Embedded Division  |  Forecast year: FY2026

**Overall: ALL PASS**  |  Excel Checks tab: 147 pass / 0 fail

| Step | Result | Time | Detail |
|---|---|---|---|
| Reported data (SEC filings) internal consistency | PASS | 0s | TXN: 88 checks: {'PASS': 88} |
| Build Excel model | PASS | 2s | Saved outputs\TXN\model\TXN_FPA_Model.xlsx; Hist_Data rows=53, Hist_Analysis metrics=43, checks=147 |
| Recalculate Excel model (formula errors) | PASS | 12s | 6817 formulas, 0 errors |
| Phase 1 - historical analysis: Python vs Excel | PASS | 1s | Checks tab: 147 PASS, 0 FAIL -> ALL CHECKS PASS |
| Phase 2 - operating model, all 3 scenarios | PASS | 43s | {'scenario': 'Upside', 'excel_revenue': 1259.553, 'python_revenue': 1259.553, 'excel_gm': 0.6138, 'cells_compared': 180, 'max_abs_diff': 4.973799150320701e-14,  |
| Phase 3 - revenue, P&L, FCF, all 3 scenarios | PASS | 36s | {'scenario': 'Upside', 'revenue': 1259.6, 'operating_profit': 531.7, 'op_margin': 0.4221, 'net_income': 467.9, 'fcf': 427.9, 'monthly_values_compared': 72, 'max |
| Phase 4 - budget vs actual & variance (Base and Downside selector) | PASS | 20s | Budget/actual/variance identical with selector on Base and Downside: True |
| Phase 5 - capex NPV / IRR / payback / break-even | PASS | 1s | Recommendation rule re-derived in Python matches: True |
| Phase 6 - dashboard export (3 scenarios) + JSON vs Excel | PASS | 27s |   Upside   revenue    1259.6  op   531.7  fcf   427.9  outlook op vs budget  +61.6 |
| Phase 7 - build dashboard page | PASS | 0s | Wrote outputs\TXN\dashboard\index.html (178,027 bytes); browser test embedded: False |
| Phase 7 - rendered dashboard vs Excel (headless browser) | PASS | 14s | Dashboard validation: 67 PASS, 0 FAIL; displayed values compared with Excel: 38 |
| Phase 7 - rebuild dashboard with test result embedded | PASS | 0s | Wrote outputs\TXN\dashboard\index.html (178,137 bytes); browser test embedded: True |
| Phase 8 - collect presentation figures and wording from the validated export | PASS | 1s | Wrote outputs\TXN\presentation\deck_data.json; checks 147/147, bottleneck POWER, verdict CHECKPOINT, overrides 0, dashboard test matches export: True |
| Phase 8 - build executive presentation (.pptx) | PASS | 4s | Wrote outputs\TXN\presentation\Power_Signal_Embedded_FY26_Executive_Review.pptx |
| Phase 8 - extract deck text for review | PASS | 2s | wrote deck_text.md |
| Phase 9 - regenerate sources & assumptions document | PASS | 0s | Wrote outputs\TXN\docs\sources_and_assumptions.md (5 sources, 24 assumptions, 272 reported values) |

Dashboard totals vs Excel: the JSON export is checked against Excel (Phase 6) and the rendered page's displayed totals are checked against the Excel control totals in a headless browser (Phase 7).

Outputs: `outputs\TXN\dashboard\index.html`, `outputs\TXN\model\TXN_FPA_Model.xlsx`, `outputs\TXN\presentation\Power_Signal_Embedded_FY26_Executive_Review.pptx`
