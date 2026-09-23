# Methodology

How the model turns engineering and manufacturing drivers into revenue, profit, cash and a capital decision.
Every step below is a visible row in `outputs/<TICKER>/model/<TICKER>_FPA_Model.xlsx` (Microchip: `outputs/MCHP/model/MCHP_FPA_Model.xlsx`); nothing important sits inside one long formula.

> Simulated data created for portfolio demonstration. It does not represent nonpublic Microchip Technology information.

## 1. Scope and data rules

- **Company reference:** Microchip Technology Inc. (NASDAQ: MCHP). Its SEC filings supply the five-year history only
  (fiscal years end March 31; FY2026 = Apr-2025 to Mar-2026).
- **The division is fictional.** The "Embedded Control Division" has three product lines: mixed-signal microcontrollers,
  analog & power management, and FPGAs. Its base-year (FY2026) revenue is set to $1,000m and split using Microchip's
  reported FY2026 product-line mix (FPGAs use Microchip's whole "Other" line as a stand-in, because FPGAs are not
  reported separately).
- **Every value has one of four classifications:** reported public data, calculated value, user-provided assumption,
  or simulated operational data (see `outputs/<TICKER>/docs/sources_and_assumptions.md`). Missing public data is never filled in.
- **Forecast year:** FY2027 (Apr-2026 to Mar-2027), monthly, calendar months. April–August 2026 are simulated actuals;
  September–March are forecast.

## 2. Workbook architecture

| Tab | Role |
|---|---|
| Cover | Status (live link to the Checks tab), tab guide, color legend |
| Sources_Log | Source register and assumptions log |
| Assumptions | The one scenario selector (`C6`), 9 drivers × 3 scenarios, product inputs, seasonality, P&L and cash inputs |
| Hist_Data / Hist_Analysis | Reported Microchip financials FY2022–FY2026 and the ratios calculated from them |
| Ops_Model | Product × month operating model for the active scenario |
| Revenue / PnL / FCF | Monthly forecast statements for the active scenario (months, quarters, FY, FY2026 comparison) |
| Budget | FY2027 budget = Base scenario, read from the Base column, so it does not move with the selector |
| Actuals | Simulated April–August 2026 results, same operating logic, different inputs |
| Variance | Budget vs actual with price / volume / mix and an operating-profit bridge |
| Capex | Hypothetical FPGA test-platform proposal evaluated in all three scenarios |
| Dash_Export | Tidy tables read by the dashboard export script |
| Checks | 147 independent checks and hand-calculable unit tests |

Color convention: blue = hardcoded input, black = formula, green = link to another tab, yellow fill = key input.

## 3. Historical analysis (Hist_Analysis)

Growth, gross / operating / net margin, R&D and SG&A as a share of sales, effective tax rate (shown as n.m. when pre-tax
income is under 2% of sales), incremental gross margin, product mix and growth, days sales / inventory / payables
outstanding, cash conversion cycle, and working capital as a share of sales.

Free cash flow uses the project definition — **net income + depreciation − capex − increase in working capital** — and is
reconciled to the market definition (operating cash flow − capex) through a line for the other non-cash items.

## 4. Operating model (Ops_Model), per product and month

| Step | Formula |
|---|---|
| Customer demand | base-year units × (1 + demand growth + product differential) × seasonality ÷ 12 |
| Units produced | internal final-test capacity × utilization |
| Good units | units produced × final-test yield |
| Available units | opening finished-goods inventory + good units |
| Units sold | MIN(customer demand, available units) |
| Closing inventory / unmet demand | available − sold / demand − sold |
| Revenue | units sold × average selling price |
| Variable manufacturing cost | units produced × variable cost per unit |
| Total manufacturing cost | variable + fixed manufacturing cost |
| Cost of revenue | units sold × weighted-average cost of available units (unsold production stays in inventory) |
| Contribution margin | revenue − units sold × (variable cost ÷ yield) |
| Gross profit | revenue − cost of revenue |
| Risk ratio | demand ÷ (capacity × yield) — the heat map's input |

"Capacity" means internal final-test capacity, since Microchip performs about 69% of its test internally (10-K FY2026).

## 5. P&L and free cash flow (PnL, FCF)

- Operating expenses: base-year R&D and SG&A × (1 + scenario growth), spread evenly across months.
- Operating profit = gross profit − operating expenses. No division-level interest; tax = scenario rate × pre-tax income.
- Capex = revenue × scenario capex %. Depreciation on new capex is straight-line over 60 months, starting the month after spend.
  Depreciation already inside factory fixed cost = 30% of fixed manufacturing cost.
- Working capital = receivables (monthly revenue × 12 × DSO ÷ 365) + inventory (from Ops_Model) − payables
  (monthly variable cost × 12 × DPO ÷ 365).
- **Free cash flow = net income + depreciation − capex − increase in working capital.**

## 6. Scenarios

Nine drivers, each with Downside / Base / Upside values and a written rationale: demand growth, selling-price change,
yield change, capacity utilization, material-cost inflation, R&D growth, SG&A growth, capex % of revenue, tax rate.
`Assumptions!D6` converts the selector to an index and every live driver is `INDEX(Downside:Upside, index)`, so there
is one model, not three. The dashboard export switches the selector three times and recalculates.

## 7. Budget vs actual (Budget, Actuals, Variance)

Year to date April–August, favorable = positive. U = units sold, P = units produced, G = good units, b = budget, a = actual.

| Component | Formula |
|---|---|
| Price | Σ Ua × (ASPa − ASPb) |
| Volume (revenue) | (total Ua − total Ub) × budget average price |
| Product mix (revenue) | Σ (actual mix − budget mix) × total Ua × ASPb |
| Sales volume / mix (profit) | same as above, valued at budget gross profit per unit |
| Material cost | −Pa × (actual − budget variable cost per unit) |
| Manufacturing yield | −budget variable cost × (Pa − Ga ÷ budget yield) |
| Fixed manufacturing spending | −(actual − budget fixed cost), plus the new-capex depreciation difference |
| Production volume & inventory | −[budget variable cost × (Ga ÷ budget yield − Pb) − Δ closing inventory − (Ua − Ub) × budget cost per unit sold] |
| Operating expenses | −(actual − budget R&D and SG&A) |

Each component is calculated separately; the Variance tab and the Checks tab confirm that the components sum exactly to
the total revenue variance and to actual operating profit, and show "UNRECONCILED" if they ever do not.
Note: volume and mix add up units across very different product lines (a $0.80 analog part counts like a $25 FPGA); this
is the standard textbook method, and its effect is explained on the dashboard.

## 8. Capital-investment decision (Capex)

Hypothetical FPGA automated final-test platform: $24.0m equipment + $3.0m installation, +0.30m units per month of
capacity, +3 points of FPGA yield, $1.5m a year of savings, $1.2m a year of maintenance, 6-year life, no salvage.
Evaluated with and without the project over FY2028–FY2033 in every scenario:

- Units sold with / without = MIN(demand, capacity with / without the project at 90% maximum utilization).
- Incremental EBITDA = incremental units × (price − variable cost ÷ new yield) + scrap reduction + savings − maintenance.
- Incremental FCF = after-tax EBIT + depreciation − increase in receivables (released in the final year) − investment.
- NPV at 9%, IRR, payback (interpolated), and **break-even utilization** = the average share of the added capacity that must
  be sold for NPV = 0 (NPV is linear in incremental units, so it is solved exactly).
- Decision rule: a scenario passes if NPV > 0, IRR ≥ 11% and payback ≤ 4 years. Reject if Base fails; approve outright only
  if Downside also passes; otherwise approve with a demand checkpoint.

## 9. Validation

| Layer | What it proves |
|---|---|
| `tests/validate_reported_data.py` | 87 checks: statements add up, product lines sum to sales, cross-filing agreement, no duplicates, units |
| Excel `Checks` tab | 147 checks and unit tests, live in the workbook; the Cover shows the overall status |
| `tests/validate_phase1_model.py` … `validate_phase5_capex.py` | Independent Python rebuilds of each phase from raw inputs, compared with Excel in every scenario |
| `scripts/export_dashboard_data.py` | 88 checks that the exported data equals Excel for all three scenarios |
| `tests/validate_dashboard.py` | Headless browser reads the numbers displayed on the dashboard and compares them with Excel |
| `scripts/run_all_validations.py --company TICKER` | Runs everything above in order and writes `outputs/<TICKER>/tests/VALIDATION_SUMMARY.md` |

## 10. Known limitations

- The division, its operating data and the capital proposal are simulated; only the history is reported.
- The outlook combines actuals with the forecast without re-basing forecast months on actual August inventory.
- Fixed factory cost is held flat across scenarios, which widens the margin range between scenarios.
- No interest or tax timing at division level; working capital covers receivables, inventory and payables only.
