# Appendix: rebuilding the dashboard in Power BI

The project's dashboard is a custom web page (`outputs/<TICKER>/dashboard/index.html`). This appendix shows how the same five pages map
to Power BI for teams that standardize on it. It reads the same export (`outputs/<TICKER>/dashboard/data.json`, produced by
`scripts/export_dashboard_data.py` from the Excel model), so Excel stays the single source of truth.

> Not verified in Power BI. These steps and measures were written against the export schema but have not been run in
> Power BI Desktop. The web dashboard is the verified deliverable (see `outputs/<TICKER>/tests/phase7_dashboard_validation.json`).

## 1. Power Query (M)

Create a parameter `DataPath` pointing to `outputs/<TICKER>/dashboard/data.json`, then one query per table.

```m
// Source (reference query, disable load)
let
    Source = Json.Document(File.Contents(DataPath))
in
    Source

// Monthly - one row per scenario x month
let
    Src = Source[scenarios],
    Scen = Record.ToTable(Src),                                 // Name = Downside/Base/Upside
    AddRows = Table.AddColumn(Scen, "rows", each Table.FromRecords([Value][monthly])),
    Keep = Table.SelectColumns(AddRows, {"Name", "rows"}),
    Expanded = Table.ExpandTableColumn(Keep, "rows", Table.ColumnNames(Keep{0}[rows])),
    Renamed = Table.RenameColumns(Expanded, {{"Name", "scenario"}}),
    Typed = Table.TransformColumnTypes(Renamed, {{"month", type date}, {"revenue", type number},
        {"gross_profit", type number}, {"operating_profit", type number}, {"net_income", type number},
        {"fcf", type number}, {"budget_revenue", type number}, {"budget_operating_profit", type number},
        {"actual_revenue", type number}, {"actual_operating_profit", type number},
        {"outlook_revenue", type number}, {"outlook_operating_profit", type number}})
in
    Typed
```

Repeat the same pattern for `product_monthly`, `product_fy`, `kpis` and `control_totals` (scenario tables), and use
`Table.FromRecords(Source[static][<name>])` for `variance_pnl`, `revenue_variance`, `op_bridge`, `variance_required`,
`variance_kpis`, `capex_summary`, `capex_cashflows`, `capex_inputs`, `scenario_drivers` and `history`.

Supporting tables:

```m
// Scenario (disconnected slicer table)
#table(type table [scenario = text, order = Int64.Type], {{"Downside", 1}, {"Base", 2}, {"Upside", 3}})
// Date: CALENDAR over Apr-2026..Mar-2027, month grain, fiscal quarter = Q1 Apr-Jun
```

Relationships: `Date[Date]` 1-* `Monthly[month]` and `ProductMonthly[month]`. Scenario is applied by measures
(below), not a relationship, so budget and actual (scenario-independent) are not filtered by it.

## 2. DAX measures

```dax
Selected Scenario = SELECTEDVALUE ( Scenario[scenario], "Base" )

Revenue =
CALCULATE ( SUM ( Monthly[revenue] ), Monthly[scenario] = [Selected Scenario] )
Gross Profit =
CALCULATE ( SUM ( Monthly[gross_profit] ), Monthly[scenario] = [Selected Scenario] )
Gross Margin = DIVIDE ( [Gross Profit], [Revenue] )
Operating Profit =
CALCULATE ( SUM ( Monthly[operating_profit] ), Monthly[scenario] = [Selected Scenario] )
Operating Margin = DIVIDE ( [Operating Profit], [Revenue] )
Net Income =
CALCULATE ( SUM ( Monthly[net_income] ), Monthly[scenario] = [Selected Scenario] )
Free Cash Flow =
CALCULATE ( SUM ( Monthly[fcf] ), Monthly[scenario] = [Selected Scenario] )

-- budget and actual are locked: always read one copy (Base rows) regardless of the slicer
Budget Revenue = CALCULATE ( SUM ( Monthly[budget_revenue] ), Monthly[scenario] = "Base" )
Budget Operating Profit = CALCULATE ( SUM ( Monthly[budget_operating_profit] ), Monthly[scenario] = "Base" )
Actual Revenue = CALCULATE ( SUM ( Monthly[actual_revenue] ), Monthly[scenario] = "Base" )
Actual Operating Profit = CALCULATE ( SUM ( Monthly[actual_operating_profit] ), Monthly[scenario] = "Base" )
YTD Budget Operating Profit =
CALCULATE ( [Budget Operating Profit], Monthly[period_type] = "Actual" )
YTD OP Variance = [Actual Operating Profit] - [YTD Budget Operating Profit]
Outlook Operating Profit =
CALCULATE ( SUM ( Monthly[outlook_operating_profit] ), Monthly[scenario] = [Selected Scenario] )
Outlook vs Budget = [Outlook Operating Profit] - [Budget Operating Profit]

-- operations (product grain)
Yield (weighted) =
DIVIDE (
    SUMX ( FILTER ( ProductMonthly, ProductMonthly[scenario] = [Selected Scenario] ),
           ProductMonthly[yield] * ProductMonthly[good_unit_capacity] / ProductMonthly[yield] * ProductMonthly[utilization] ),
    SUMX ( FILTER ( ProductMonthly, ProductMonthly[scenario] = [Selected Scenario] ),
           ProductMonthly[good_unit_capacity] / ProductMonthly[yield] * ProductMonthly[utilization] )
)
Demand Capacity Ratio =
CALCULATE ( MAX ( ProductMonthly[demand_capacity_ratio] ), ProductMonthly[scenario] = [Selected Scenario] )
Risk Bin =
SWITCH ( TRUE (), [Demand Capacity Ratio] >= 1, "SHORT", [Demand Capacity Ratio] >= 0.85, "TIGHT", "OK" )
Cost per Good Unit =
CALCULATE ( AVERAGE ( ProductFY[cost_per_good_unit] ), ProductFY[scenario] = [Selected Scenario] )

-- capital (scenario-independent table, highlight the selected column)
NPV = CALCULATE ( MAX ( CapexSummary[npv] ), CapexSummary[scenario] = [Selected Scenario] )
IRR = CALCULATE ( MAX ( CapexSummary[irr] ), CapexSummary[scenario] = [Selected Scenario] )
NPV (rebuilt) =
VAR r = LOOKUPVALUE ( CapexInputs[value], CapexInputs[key], "rate" )
RETURN SUMX ( FILTER ( CapexCashflows, CapexCashflows[scenario] = [Selected Scenario] ),
              CapexCashflows[fcf] / ( 1 + r ) ^ CapexCashflows[year] )

-- validation: must equal the Excel control totals
Revenue Control = CALCULATE ( MAX ( ControlTotals[value] ), ControlTotals[control_id] = "revenue_fy",
                              ControlTotals[scenario] = [Selected Scenario] )
Revenue Check = IF ( ABS ( [Revenue] - [Revenue Control] ) < 0.000001, "PASS", "FAIL" )
```

Note: `Yield (weighted)` above reconstructs units tested from capacity x utilization; the simpler equivalent is to add
`units_produced` and `good_units` columns to `product_monthly` in `phase6_export.py` and take
`DIVIDE ( SUM ( good ), SUM ( produced ) )`.

## 3. Page map

| Page | Visuals | Measures / tables |
|---|---|---|
| Executive Overview | 6 cards; line (budget vs outlook by month); column (outlook OP) with budget line; product table | Revenue, Gross Margin, Operating Profit, Free Cash Flow, YTD OP Variance, Outlook vs Budget; ProductFY |
| Semiconductor Operations | product cards; matrix heat map (product x month) with conditional formatting on the ratio; GM line by product | Demand Capacity Ratio, Risk Bin, Cost per Good Unit, ProductMonthly |
| Budget vs Actual | waterfall chart on `op_bridge` (sort by `order`); table of `variance_required`; PVM matrix | op_bridge, revenue_variance, variance_pnl, variance_kpis |
| Scenario Planning | clustered bar by scenario (no slicer filter); driver table; line by scenario | Monthly (all scenarios), scenario_drivers |
| Capital Allocation | cards NPV / IRR / payback; line of cumulative FCF by scenario; criteria matrix | capex_summary, capex_cashflows, capex_inputs |

## 4. Verification before use

Before calling a Power BI build verified, compare every card with `control_totals` for all three scenarios
(e.g. a hidden validation page using the `... Check` measures), the same way `tests/validate_dashboard.py`
verifies the web dashboard.
