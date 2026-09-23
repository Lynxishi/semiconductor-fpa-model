# Dashboard data dictionary

`outputs/<TICKER>/dashboard/data.json` is written by `scripts/export_dashboard_data.py` from the `Dash_Export` tab of the workbook.
Money is in USD millions, units in millions, ratios as fractions (0.25 = 25%), months as ISO dates (first of month).

```
{
  "generated_utc", "source_workbook", "source_sha256", "disclaimer",
  "company":   { ticker, company, legal_name, exchange, division, forecast_year, base_year, filings, capex_project, products[] },
  "scenarios": { "Downside" | "Base" | "Upside": { meta, monthly, product_monthly, product_fy, kpis, control_totals } },
  "static":    { variance_pnl, revenue_variance, op_bridge, variance_required, variance_kpis,
                 capex_summary, capex_cashflows, capex_inputs, history, scenario_drivers }
}
```

Scenario tables change with the scenario; static tables are identical in every scenario (the export checks this).

## Scenario tables

| Table | Grain | Key fields |
|---|---|---|
| `meta` | key / value | scenario, checks_status, actuals_through, outlook_note |
| `monthly` | 12 months | revenue, gross_profit, gross_margin, opex, operating_profit, net_income, ebitda, fcf, cumulative_fcf, capex, budget_revenue, budget_operating_profit, actual_revenue (Apr–Aug only, else null), actual_operating_profit, outlook_revenue, outlook_operating_profit, period_type (Actual / Forecast) |
| `product_monthly` | 3 products × 12 months | product_id (the three `[[products]]` ids in the config; Microchip: MCU, ANALOG, FPGA), demand_units, good_unit_capacity, demand_capacity_ratio, units_sold, unmet_demand, revenue, asp, gross_profit, gross_margin, yield, utilization, cost_per_good_unit, closing_inventory_units |
| `product_fy` | 3 products | revenue, mix, units_sold, asp, gross_profit, gross_margin, contribution_margin(_pct), yield, utilization, cost_per_good_unit, peak_demand_capacity_ratio, unmet_demand, base_year_revenue |
| `kpis` | 25 metrics | metric_id, label, value, unit, basis |
| `control_totals` | 10 totals | revenue_fy, operating_profit_fy, net_income_fy, fcf_fy, gross_profit_fy, budget_revenue_fy, ytd_actual_revenue, ytd_actual_op, ytd_op_variance, capex_npv_base — the dashboard must reproduce these |

## Static tables

| Table | Grain | Key fields |
|---|---|---|
| `variance_pnl` | P&L line | budget, actual, variance, unit, assessment |
| `revenue_variance` | component | price / volume / mix / total by product id and division |
| `op_bridge` | waterfall step | order, step_id, label, value, kind (start / delta / end) |
| `variance_required` | the 8 required variances + 2 reconciliation flags | value, assessment |
| `variance_kpis` | product × KPI | budget, actual (units sold, ASP, yield, utilization, cost per good unit, unmet demand) |
| `capex_summary` | scenario | npv, irr, payback_years, break_even_utilization, implied_utilization, npv_test, irr_test, payback_test, meets_all |
| `capex_cashflows` | scenario × year 0–6 | fiscal_year, fcf, cumulative_fcf, pv, incremental_units, scrap_savings, ebitda |
| `capex_inputs` | key / value | project inputs, recommendation, rule, checkpoint |
| `history` | 5 reported fiscal years | Reported net sales and the segment line behind each product (`<product key>_segment_sales`); calculated margins, FCF, capex %, days inventory |
| `scenario_drivers` | 9 drivers | Downside, Base, Upside values |
