# Multi-company engine

The model, the checks, the dashboard and the deck are generic. Everything about one company lives in
`companies/<TICKER>/`, and one command builds, tests and publishes it:

```bash
python scripts/run_all_validations.py --company MCHP
```

Outputs go to `outputs/<TICKER>/`:

| Output | Path |
|---|---|
| Excel model | `outputs/<TICKER>/model/<TICKER>_FPA_Model.xlsx` |
| Dashboard (self-contained, open in any browser) | `outputs/<TICKER>/dashboard/index.html` |
| Dashboard data (validated export) | `outputs/<TICKER>/dashboard/data.json` |
| Executive deck + extracted text | `outputs/<TICKER>/presentation/*.pptx`, `deck_text.md` |
| Sources & assumptions log | `outputs/<TICKER>/docs/sources_and_assumptions.md` |
| Every test result + screenshots | `outputs/<TICKER>/tests/` (`VALIDATION_SUMMARY.md` first) |

Each script also runs on its own with `--company TICKER` (or the `FPA_COMPANY` environment variable; the default is MCHP).

## What a company folder contains

| File | Required | What it is |
|---|---|---|
| `config.toml` | yes | Every company-specific value: names, fiscal calendar, the three product lines, simulated operating inputs, scenario drivers, actuals, the capex proposal, optional wording |
| `reported_financials.csv` | yes | The reported SEC data, one row per value, in the standard format below |
| `source_register.csv` | yes | One row per filing used (feeds the sources log and the Sources tab) |
| `assumptions_log.csv` | yes | One row per decision / assumption, with who decided it |
| `cross_filing_checks.csv` | no | The same values transcribed from a second filing, to catch transcription errors |
| `collect_reported_data.py` | no | The script that produced `reported_financials.csv` (kept for audit) |

`config.toml` is validated before anything is built (`scripts/engine.py`); errors say which key is wrong.

## Reported data format

`reported_financials.csv` columns:
`line_id, period, period_type, statement, line_item, value_usd_m, classification, source_document, accession_no, source_url, label_as_printed, note`.

Values are USD millions as printed (cash outflows negative, as in the cash-flow statement). `classification` is always
`Reported public data`. A value that is not printed in a filing is **left out**, never estimated. The engine treats a
missing optional line as "not presented" and says so in the notes.

`line_id` is what the model reads; `line_item` / `label_as_printed` are what the filing says.

| Group | line_ids | Required |
|---|---|---|
| Income statement | `net_sales, cost_of_sales, gross_profit, rd, sga, operating_income, pretax_income, income_tax, net_income` | yes |
| | `total_opex` (if printed; otherwise the checks test the components directly); `other_opex:<name>` (amortization, restructuring ...); `nonop:<name>` (interest, other) | as printed |
| Product lines | `segment:<name>` - the reported revenue lines; each product in the config points at one | at least one per product |
| Segment info | `segment_info:<name>` - other segment disclosures (used only by optional checks) | no |
| Cash flow | `cf:net_income, cf:da, cf:cfo, cf:capex` | yes |
| | `cf:sbc, cf:cfi, cf:cff, cf:dividends, cf:cash_end`; working-capital lines `wc:<name>` | as printed |
| Balance sheet | `bs:cash, bs:receivables, bs:inventories, bs:payables, bs:accrued, bs:total_current_assets, bs:total_current_liabilities, bs:ppe, bs:total_assets, bs:equity` | yes |
| | `bs:st_investments, bs:other_current_assets, bs:current_debt, bs:long_term_debt, bs:ca_other:<name>, bs:cl_other:<name>` | as printed |
| Other | `inv:<name>` (inventory detail), `memo:<name>`, `fact:<name>` (non-financial facts) | no |

Periods are the labels in `company.history_years` (exactly five fiscal years) plus, optionally, the two quarters in
`[company.quarter]` (the same line_ids, e.g. `Q1 FY2026` and `Q1 FY2027`).

`tests/validate_reported_data.py` checks, by line_id: the income statement adds up, product lines sum to net sales,
the cash flow ties to the balance sheet and income statement, current assets and liabilities equal their components,
the optional identities and cross-filing values, and record integrity (no duplicates, sources on every row, valid periods, units).

## config.toml sections

| Section | Contents |
|---|---|
| `[company]` | ticker, name, legal_name, exchange, cik, fiscal_year_end, history_years (5), base_year (= last history year), filings_summary. Optional `short_name`, `[company.quarter]`, `[company.checks]` |
| `[forecast]` | year, start (first month, ISO date), actual_months (must be 5), actuals_label, as_of |
| `[division]` | the fictional division: name, short, base_revenue, rd0, sga0 and notes |
| `[[products]]` ×3 | key, id, name, short, optional abbr / tag labels, segment (a `segment:` line_id), optional segment_share (see below), base-year inputs asp0, yld0, vc0, fix0, cap0, gdiff, inv_m, and actual inputs a_dem, a_asp, a_yld, a_util, a_vc, a_fix. All **simulated** |
| `[drivers.<id>]` ×9 | `values = [Downside, Base, Upside]` and `why`, for dem_g, asp_chg, yld_adj, util, mat_inf, rd_g, sga_g, capex_pct, tax |
| `[seasonality]` | 12 monthly factors that sum to 12 |
| `[pnl]` | dso, dpo, dep_share, life_months |
| `[actuals]` | month-by-month demand noise (5 values), a_rd, a_sga, a_capex, notes |
| `[capex]` | target (a product key), project name, price, install, add_cap, yld_gain, savings, maint, life, salvage, umax, erosion, rate, hurdle, max_pb, growth_after [3 scenarios], wording keys |
| `[history.*]` | optional labels, notes and short source labels for Hist_Data rows |
| `[narrative]` | optional deck wording overrides (below) |

The base-year product mix is **not** an input: it is each product's reported `segment:` line divided by reported net sales.

If a company does not report a product separately, two products can share one reported segment with `segment_share`
(a user-provided assumption, e.g. 0.5 / 0.5; shares of one segment may not exceed 100%). When shares are used, or the
three products do not cover every reported segment, the Assumptions tab shows the reported share, the assigned split,
and the mix renormalized to 100% over the three products. Texas Instruments uses this: power management and signal
chain are an assumed 50/50 split of TI's Analog segment, embedded processors = Embedded Processing, and TI's Other
segment is not modeled.

## Deck wording

`scripts/prepare_deck_data.py` writes every sentence in the deck from the model's numbers by fixed rules:

- ahead / behind: sign of the YTD operating-profit variance
- bottleneck product: highest Base-case peak demand ÷ good-unit capacity
- yield problem: largest YTD yield shortfall vs budget
- soft product: largest YTD utilization cut vs budget
- largest variance driver: biggest absolute step in the profit bridge ("driven by" if it has the same sign as the total, else "despite")
- history: peak and trough years of reported net sales and gross margin
- capex headline: which scenarios meet every criterion, and the decision rule's verdict (approve / approve with checkpoint / reject)

Any sentence can be replaced in `[narrative]`. An override is a template: `{placeholders}` are filled from the model
(`{ytd_op_var}`, `{capex}`, `{tl}`, `{be_util}`, `{q_growth0}`, `{dem_base}` ...; the full list is the `ctx` dictionary
in `prepare_deck_data.py`). Numbers must come from placeholders, not be typed. An unknown key stops the build.

## Adding a company

1. `companies/<TICKER>/reported_financials.csv`: transcribe five years of statements (and optionally two comparable
   quarters) from the 10-K / 10-Q, one row per value, with the accession number and the statement page URL. Record each
   filing in `source_register.csv`.
2. Copy `companies/MCHP/config.toml`, then change the company block, fiscal calendar, product lines (pick three reported
   revenue lines and point each product's `segment` at one), the simulated inputs, drivers and capex target.
3. Log every decision in `assumptions_log.csv`.
4. Run `python scripts/run_all_validations.py --company <TICKER>` and read `outputs/<TICKER>/tests/VALIDATION_SUMMARY.md`.
5. Open the dashboard and the deck and review the wording before sharing.

## Regression test

`tests/regress_baseline.py` compares the Microchip outputs with `tests/baseline_mchp/` (saved before the engine
refactor). Every number in `data.json` must match to 1e-9 and every string exactly; the deck text must match except the
three documented changes listed in the script. It runs automatically as the last step for MCHP.

## Known limits

- Exactly three product lines, five history years and five actual months (the sheet layouts are built for these).
- The shell sandbox used to build this project cannot reach sec.gov, so the reported data is transcribed from filings
  (with citations) rather than downloaded automatically. An automatic SEC download (planned Phase 10b) was not built.

## Companies built so far

| Ticker | Division | Products | Forecast year | Status |
|---|---|---|---|---|
| MCHP | Embedded Control Division | Mixed-signal MCUs, Analog & power, FPGAs | FY2027 (Apr-Mar) | All pass; regression vs pre-engine baseline: 0 differences |
| TXN | Power, Signal & Embedded Division | Power management, Signal chain (assumed 50/50 split of Analog), Embedded processors | FY2026 (calendar) | All pass |
