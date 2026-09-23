# Validation summary - Microchip Technology (MCHP)

Run: 2026-09-23T00:59:09+00:00 UTC  |  Division: Embedded Control Division  |  Forecast year: FY2027

**Overall: FAILURES**  |  Excel Checks tab: 149 pass / 0 fail

| Step | Result | Time | Detail |
|---|---|---|---|
| Reported data (SEC filings) internal consistency | PASS | 0s | MCHP: 88 checks: {'PASS': 88} |
| Build Excel model | PASS | 1s | Saved outputs\MCHP\model\MCHP_FPA_Model.xlsx; Hist_Data rows=55, Hist_Analysis metrics=43, checks=149 |
| Recalculate Excel model (formula errors) | PASS | 9s | 6821 formulas, 0 errors |
| Phase 1 - historical analysis: Python vs Excel | PASS | 1s | Checks tab: 149 PASS, 0 FAIL -> ALL CHECKS PASS |
| Phase 2 - operating model, all 3 scenarios | PASS | 23s | {'scenario': 'Upside', 'excel_revenue': 1309.782, 'python_revenue': 1309.782, 'excel_gm': 0.6414, 'cells_compared': 180, 'max_abs_diff': 5.1514348342607263e-14, |
| Phase 3 - revenue, P&L, FCF, all 3 scenarios | PASS | 21s | {'scenario': 'Upside', 'revenue': 1309.8, 'operating_profit': 513.3, 'op_margin': 0.3919, 'net_income': 431.2, 'fcf': 377.4, 'monthly_values_compared': 72, 'max |
| Phase 4 - budget vs actual & variance (Base and Downside selector) | PASS | 18s | Budget/actual/variance identical with selector on Base and Downside: True |
| Phase 5 - capex NPV / IRR / payback / break-even | PASS | 1s | Recommendation rule re-derived in Python matches: True |
| Phase 6 - dashboard export (3 scenarios) + JSON vs Excel | PASS | 26s |   Upside   revenue    1309.8  op   513.3  fcf   377.4  outlook op vs budget  +71.9 |
| Phase 7 - build dashboard page | PASS | 0s | Wrote outputs\MCHP\dashboard\index.html (177,669 bytes); browser test embedded: False |
| Phase 7 - rendered dashboard vs Excel (headless browser) | PASS | 5s | Dashboard validation: 67 PASS, 0 FAIL; displayed values compared with Excel: 38 |
| Phase 7 - rebuild dashboard with test result embedded | PASS | 0s | Wrote outputs\MCHP\dashboard\index.html (177,779 bytes); browser test embedded: True |
| Phase 8 - collect presentation figures and wording from the validated export | PASS | 1s | Wrote outputs\MCHP\presentation\deck_data.json; checks 149/149, bottleneck FPGA, verdict CHECKPOINT, overrides 5, dashboard test matches export: True |
| Phase 8 - build executive presentation (.pptx) | PASS | 0s | Wrote outputs\MCHP\presentation\Embedded_Control_FY27_Executive_Review.pptx |
| Phase 8 - extract deck text for review | PASS | 1s | wrote deck_text.md |
| Phase 9 - regenerate sources & assumptions document | PASS | 0s | Wrote outputs\MCHP\docs\sources_and_assumptions.md (5 sources, 37 assumptions, 263 reported values) |
| Phase 10 - regression: Microchip outputs unchanged vs pre-engine baseline | FAIL | 0s | Regression vs baseline: 2 files compared, 1 differences |

Dashboard totals vs Excel: the JSON export is checked against Excel (Phase 6) and the rendered page's displayed totals are checked against the Excel control totals in a headless browser (Phase 7).

Outputs: `outputs\MCHP\dashboard\index.html`, `outputs\MCHP\model\MCHP_FPA_Model.xlsx`, `outputs\MCHP\presentation\Embedded_Control_FY27_Executive_Review.pptx`
