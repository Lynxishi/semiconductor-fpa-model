# Phase 6 - Dash_Export tab: tidy tables for the web dashboard (one header row per table, one record per row).
# Executed inside build_model.py's namespace after phase5_capex.py. Not a standalone script.
#
# Tables that depend on the scenario show the LIVE scenario; scripts/export_dashboard_data.py switches the selector
# to Downside / Base / Upside, recalculates, and reads these tables three times.

DX = {}  # table id -> dict(header=row, first=row, last=row, cols=[names])
QLAB = [f"Q{q} FY{FYL[-2:]}" for q in (1, 2, 3, 4) for _ in range(3)]
PRODUCT_IDS = {p["key"]: p["id"] for p in EN.products}
_LAST_ACT = _dt.datetime(FSTART.year + (FSTART.month + 3) // 12, (FSTART.month + 3) % 12 + 1, 1)
_FC1 = _dt.datetime(_LAST_ACT.year + _LAST_ACT.month // 12, _LAST_ACT.month % 12 + 1, 1)
AMON = f"{FSTART:%b}-{_LAST_ACT:%b}"    # actual months, e.g. Apr-Aug
FMON = f"{_FC1:%b}-{FEND:%b}"           # forecast months, e.g. Sep-Mar


def build_phase6_tab():
    global r
    dx = wb.create_sheet("Dash_Export")
    title_block(dx, "Dashboard Export Tables",
                "Read by scripts/export_dashboard_data.py. Do not insert rows inside a table. Every cell is a link or formula.")
    setw(dx, {get_column_letter(i): 15 for i in range(1, 22)})
    dx.column_dimensions["A"].width = 24
    r = 5

    def table(tid, desc, cols, rows):
        """cols: list of (name, number_format). rows: list of lists (formula strings with '=' or literals/None)."""
        global r
        dx.cell(row=r, column=1, value=f"TABLE: {tid}").font = Font(name=F, size=11, bold=True, color=NAVY)
        dx.cell(row=r, column=3, value=desc).font = f_sub
        r += 1
        header_row(dx, r, [c for c, _ in cols])
        DX[tid] = {"header": r, "first": r + 1, "last": r + len(rows), "cols": [c for c, _ in cols]}
        r += 1
        for row in rows:
            for j, ((cname, fmt), v) in enumerate(zip(cols, row)):
                c = dx.cell(row=r, column=1 + j, value=v)
                c.font = f_link if isinstance(v, str) and v.startswith("=") else f_base
                if fmt:
                    c.number_format = fmt
            r += 1
        r += 1

    def col_of(tid, name):
        return get_column_letter(1 + DX[tid]["cols"].index(name))

    PL = lambda k, c: f"=PnL!{c}${PNL[k]}"  # noqa: E731
    D4 = '#,##0.0000'

    # ---------- meta ----------
    table("meta", "Model metadata (live)", [("key", None), ("value", None)], [
        ["scenario", "=Assumptions!$C$6"],
        ["checks_status", None],  # patched in build_model.py once the Checks status cell exists
        ["model_phase", "Phase 7 of 9"],
        ["disclaimer", DISCLAIMER],
        ["actuals_through", f"{_mlab(_LAST_ACT)} (simulated)"],
        ["fiscal_year", f"{FYL} = {_mlab(FSTART)} to {_mlab(FEND)}"],
        ["company_reference", f"{EN.company['legal_name']} ({EN.company['exchange']}: {EN.ticker}) - reported data only in history tables"],
        ["units", "USD millions unless noted; units in millions"],
        ["outlook_note", f"Outlook = simulated actuals {AMON} + active-scenario forecast {FMON} (forecast months are not re-based on actual inventory)"],
    ])

    # ---------- monthly (live) ----------
    mrows = []
    for i, c in enumerate(MONTH_COLS):
        act = i < 5
        mrows.append([
            f"=Ops_Model!{c}$6", QLAB[i], "Actual" if act else "Forecast",
            PL("rev", c), PL("gp", c), PL("gm", c), PL("opex", c), PL("op", c), PL("om", c), PL("ni", c), PL("ebitda", c),
            f"=FCF!{c}${FCFT['fcf']}", f"=FCF!{c}${FCFT['fcf_cum']}", f"=FCF!{c}${FCFT['capex']}",
            f"=Budget!{c}${BUD[('div', 'rev')]}", f"=Budget!{c}${BUD[('div', 'gp')]}", f"=Budget!{c}${BUD[('div', 'op')]}",
            f"=Actuals!{c}${ACT[('div', 'rev')]}" if act else None,
            f"=Actuals!{c}${ACT[('div', 'gp')]}" if act else None,
            f"=Actuals!{c}${ACT[('div', 'op')]}" if act else None,
            (f"=Actuals!{c}${ACT[('div', 'rev')]}" if act else PL("rev", c)),
            (f"=Actuals!{c}${ACT[('div', 'op')]}" if act else PL("op", c)),
        ])
    table("monthly", f"One row per month; forecast columns follow the live scenario; budget is locked; actuals {AMON} only",
          [("month", "yyyy-mm-dd"), ("fiscal_quarter", None), ("period_type", None), ("revenue", D4),
           ("gross_profit", D4), ("gross_margin", D4), ("opex", D4), ("operating_profit", D4), ("operating_margin", D4),
           ("net_income", D4), ("ebitda", D4), ("fcf", D4), ("cumulative_fcf", D4), ("capex", D4),
           ("budget_revenue", D4), ("budget_gross_profit", D4), ("budget_operating_profit", D4),
           ("actual_revenue", D4), ("actual_gross_profit", D4), ("actual_operating_profit", D4),
           ("outlook_revenue", D4), ("outlook_operating_profit", D4)], mrows)

    # ---------- product x month (live) ----------
    prow_ = []
    for p, pname in PRODUCTS:
        O = lambda k, c: f"Ops_Model!{c}${OPS[(p, k)]}"  # noqa: E731
        for i, c in enumerate(MONTH_COLS):
            prow_.append([
                f"=Ops_Model!{c}$6", PRODUCT_IDS[p], pname, "=" + O("dem", c), f"={O('cap', c)}*{O('yld', c)}",
                "=" + O("dcr", c), "=" + O("sold", c), "=" + O("unmet", c), "=" + O("rev", c), "=" + O("asp", c),
                "=" + O("gp", c), "=" + O("gm", c), "=" + O("yld", c), "=" + O("util", c), "=" + O("cpgu", c),
                "=" + O("inv_close", c)])
    table("product_monthly", "3 products x 12 months (live scenario). demand_capacity_ratio feeds the risk heat map",
          [("month", "yyyy-mm-dd"), ("product_id", None), ("product", None), ("demand_units", D4),
           ("good_unit_capacity", D4), ("demand_capacity_ratio", D4), ("units_sold", D4), ("unmet_demand", D4),
           ("revenue", D4), ("asp", D4), ("gross_profit", D4), ("gross_margin", D4), ("yield", D4),
           ("utilization", D4), ("cost_per_good_unit", D4), ("closing_inventory_units", D4)], prow_)

    # ---------- product FY (live) ----------
    pfy = []
    for p, pname in PRODUCTS:
        O = lambda k: f"Ops_Model!${FYC}${OPS[(p, k)]}"  # noqa: E731
        pfy.append([PRODUCT_IDS[p], pname, "=" + O("rev"), f"={O('rev')}/Ops_Model!${FYC}${OPS[('div', 'rev')]}",
                    "=" + O("sold"), "=" + O("asp"), "=" + O("gp"), "=" + O("gm"), "=" + O("cm"), "=" + O("cm_pct"),
                    "=" + O("yld"), "=" + O("util"), "=" + O("cpgu"),
                    f"=MAX(Ops_Model!$C${OPS[(p, 'dcr')]}:$N${OPS[(p, 'dcr')]})", "=" + O("unmet"),
                    f"=Assumptions!{PCOL[p]}${ASM['rev0']}"])
    table("product_fy", f"{FYL} product profitability (live scenario)",
          [("product_id", None), ("product", None), ("revenue", D4), ("mix", D4), ("units_sold", D4), ("asp", D4),
           ("gross_profit", D4), ("gross_margin", D4), ("contribution_margin", D4), ("contribution_margin_pct", D4),
           ("yield", D4), ("utilization", D4), ("cost_per_good_unit", D4), ("peak_demand_capacity_ratio", D4),
           ("unmet_demand", D4), ("base_year_revenue", D4)], pfy)

    # ---------- KPIs (live) ----------
    V_ = lambda key, col: f"=Variance!${col}${VAR[key]}"  # noqa: E731
    mrow_ = lambda name: f"Dash_Export!${col_of('monthly', name)}${DX['monthly']['first']}:${col_of('monthly', name)}${DX['monthly']['last']}"  # noqa: E731
    kpis = [
        ["revenue", f"Revenue {FYL}", f"=PnL!${FYC}${PNL['rev']}", "$m", "Forecast (live scenario)"],
        ["revenue_growth", f"Revenue growth vs {BY} base year", f"=PnL!$U${PNL['rev']}", "%", "Forecast"],
        ["gross_profit", f"Gross profit {FYL}", f"=PnL!${FYC}${PNL['gp']}", "$m", "Forecast"],
        ["gross_margin", f"Gross margin {FYL}", f"=PnL!${FYC}${PNL['gm']}", "%", "Forecast"],
        ["operating_profit", f"Operating profit {FYL}", f"=PnL!${FYC}${PNL['op']}", "$m", "Forecast"],
        ["operating_margin", f"Operating margin {FYL}", f"=PnL!${FYC}${PNL['om']}", "%", "Forecast"],
        ["net_income", f"Net income {FYL}", f"=PnL!${FYC}${PNL['ni']}", "$m", "Forecast"],
        ["ebitda", f"EBITDA {FYL}", f"=PnL!${FYC}${PNL['ebitda']}", "$m", "Forecast"],
        ["fcf", f"Free cash flow {FYL}", f"=FCF!${FYC}${FCFT['fcf']}", "$m", "Forecast"],
        ["fcf_margin", f"FCF margin {FYL}", f"=FCF!${FYC}${FCFT['fcf_m']}", "%", "Forecast"],
        ["capex", f"Capital expenditures {FYL}", f"=FCF!${FYC}${FCFT['capex']}", "$m", "Forecast"],
        ["yield_division", "Division yield (good / tested units)",
         "=(" + "+".join(f"Ops_Model!${FYC}${OPS[(p, 'good')]}" for p, _ in PRODUCTS) + ")/(" +
         "+".join(f"Ops_Model!${FYC}${OPS[(p, 'prod')]}" for p, _ in PRODUCTS) + ")", "%", "Forecast"],
        ["utilization", "Capacity utilization", f"=Assumptions!$F${ASM['util']}", "%", "Forecast"],
        ["revenue_at_risk", "Revenue at risk from unmet demand", f"=Ops_Model!${FYC}${OPS[('div', 'lost_rev')]}", "$m", "Forecast"],
        ["budget_revenue_fy", f"Budget revenue {FYL}", f"=Budget!$O${BUD[('div', 'rev')]}", "$m", "Budget (locked)"],
        ["budget_op_fy", f"Budget operating profit {FYL}", f"=Budget!$O${BUD[('div', 'op')]}", "$m", "Budget (locked)"],
        ["ytd_budget_revenue", f"YTD budget revenue ({AMON})", V_("rev", "C"), "$m", "Budget (locked)"],
        ["ytd_actual_revenue", f"YTD actual revenue ({AMON})", V_("rev", "D"), "$m", "Actual (simulated)"],
        ["ytd_revenue_variance", "YTD revenue variance", V_("rev", "E"), "$m", "Variance"],
        ["ytd_budget_op", "YTD budget operating profit", V_("op", "C"), "$m", "Budget (locked)"],
        ["ytd_actual_op", "YTD actual operating profit", V_("op", "D"), "$m", "Actual (simulated)"],
        ["ytd_op_variance", "YTD operating-profit variance", V_("op", "E"), "$m", "Variance"],
        ["outlook_revenue", f"{FYL} outlook revenue (actual + forecast)", f"=SUM({mrow_('outlook_revenue')})", "$m", "Outlook"],
        ["outlook_op", f"{FYL} outlook operating profit", f"=SUM({mrow_('outlook_operating_profit')})", "$m", "Outlook"],
        ["outlook_op_vs_budget", "Outlook operating profit vs budget", f"=SUM({mrow_('outlook_operating_profit')})-Budget!$O${BUD[('div', 'op')]}", "$m", "Outlook"],
    ]
    table("kpis", "Headline KPIs (live scenario unless noted)",
          [("metric_id", None), ("label", None), ("value", D4), ("unit", None), ("basis", None)], kpis)

    # ---------- variance (scenario-independent) ----------
    vs = []
    for key, lab in [*((f"rev_{p['key']}", f"Revenue - {p['abbr']}") for p in EN.products),
                     ("rev", "Total revenue"), ("cor", "Cost of revenue"), ("gp", "Gross profit"), ("gm", "Gross margin"),
                     ("rd", "R&D"), ("sga", "SG&A"), ("opex", "Total operating expenses"), ("op", "Operating profit"),
                     ("om", "Operating margin")]:
        vs.append([key, lab, V_(key, "C"), V_(key, "D"), V_(key, "E"), "%" if key in ("gm", "om") else "$m",
                   None if key in ("gm", "om") else V_(key, "G")])
    table("variance_pnl", f"YTD {AMON} budget vs actual P&L (favorable variance = positive)",
          [("line_id", None), ("label", None), ("budget", D4), ("actual", D4), ("variance", D4), ("unit", None),
           ("assessment", None)], vs)
    comp = []
    for key, lab in [("r_price", "Price"), ("r_vol", "Volume"), ("r_mix", "Product mix"), ("r_tot", "Total revenue variance")]:
        comp.append([key, lab] + [f"=Variance!{PCOL[p]}${VAR[key]}" for p, _ in PRODUCTS] + [f"=Variance!$F${VAR[key]}"])
    table("revenue_variance", "Revenue variance by component and product (YTD)",
          [("component_id", None), ("label", None), *((p["id"], D4) for p in EN.products), ("division", D4)], comp)
    br = [[0, "start", "Budget operating profit", f"=Variance!$F${VAR['b_start']}", "start"]]
    for i, (key, lab) in enumerate([("b_price", "Price"), ("b_vol", "Sales volume"), ("b_mix", "Product mix"),
                                    ("b_mat", "Material cost"), ("b_yld", "Manufacturing yield"),
                                    ("b_fix", "Fixed manufacturing spending"), ("b_abs", "Production volume & inventory"),
                                    ("b_rd", "R&D"), ("b_sga", "SG&A")], start=1):
        br.append([i, key, lab, f"=Variance!$F${VAR[key]}", "delta"])
    br.append([10, "end", "Actual operating profit", f"=Variance!$F${VAR['b_act']}", "end"])
    table("op_bridge", "Operating-profit waterfall, YTD (division)",
          [("order", '0'), ("step_id", None), ("label", None), ("value", D4), ("kind", None)], br)
    req = [[k, va_l, f"=Variance!$F${VAR[k]}", f"=Variance!$G${VAR[k]}"] for k, va_l in [
        ("s_rev", "Total revenue variance"), ("s_price", "Price variance"), ("s_vol", "Volume variance"),
        ("s_mix", "Product-mix variance"), ("s_mat", "Material-cost variance"), ("s_yld", "Manufacturing-yield variance"),
        ("s_opex", "Operating-expense variance"), ("s_op", "Operating-profit variance")]]
    req.append(["reconciliation_revenue", "Revenue decomposition status", f"=Variance!$F${VAR['r_unrec']}", f"=Variance!$G${VAR['r_unrec']}"])
    req.append(["reconciliation_bridge", "Operating-profit bridge status", f"=Variance!$F${VAR['b_unrec']}", f"=Variance!$G${VAR['b_unrec']}"])
    table("variance_required", "The eight required variances + reconciliation flags",
          [("variance_id", None), ("label", None), ("value", D4), ("assessment", None)], req)
    vk = []
    for key, lab in [("sold", "Units sold"), ("asp", "ASP"), ("yld", "Yield"), ("util", "Utilization"),
                     ("cpgu", "Cost per good unit"), ("unmet", "Unmet demand")]:
        for p, _ in PRODUCTS:
            rr = VAR[(f"kpi_{key}", p)]
            vk.append([PRODUCT_IDS[p], key, lab, f"=Variance!$C${rr}", f"=Variance!$D${rr}"])
    table("variance_kpis", "YTD operating KPIs by product, budget vs actual",
          [("product_id", None), ("kpi_id", None), ("label", None), ("budget", D4), ("actual", D4)], vk)

    # ---------- capex (all scenarios) ----------
    cs = []
    for si, (sname, _) in enumerate(SCEN):
        colL = get_column_letter(3 + si)
        cs.append([sname, f"=Capex!$F${CPX[(sname, 'npv')]}", f"=Capex!$F${CPX[(sname, 'irr')]}",
                   f"=Capex!$F${CPX[(sname, 'pb')]}", f"=Capex!$F${CPX[(sname, 'be_util')]}",
                   f"=Capex!$F${CPX[(sname, 'impl_util')]}", f"=Capex!${colL}${CPX['c_npv']}",
                   f"=Capex!${colL}${CPX['c_irr']}", f"=Capex!${colL}${CPX['c_pb']}", f"=Capex!${colL}${CPX['c_all']}"])
    table("capex_summary", "Capital decision by scenario (all three evaluated at once)",
          [("scenario", None), ("npv", D4), ("irr", D4), ("payback_years", D4), ("break_even_utilization", D4),
           ("implied_utilization", D4), ("npv_test", None), ("irr_test", None), ("payback_test", None),
           ("meets_all", None)], cs)
    ccf = []
    for sname, _ in SCEN:
        for yi, col in enumerate(YCOLS):
            ccf.append([sname, yi, f"=Capex!{col}${CPX[(sname, 'fy')]}", f"=Capex!{col}${CPX[(sname, 'fcf')]}",
                        f"=Capex!{col}${CPX[(sname, 'cum')]}", f"=Capex!{col}${CPX[(sname, 'pv')]}",
                        f"=Capex!{col}${CPX[(sname, 'd_units')]}", f"=Capex!{col}${CPX[(sname, 'scrap')]}",
                        f"=Capex!{col}${CPX[(sname, 'ebitda')]}"])
    table("capex_cashflows", "Incremental cash flows by scenario and year",
          [("scenario", None), ("year", '0'), ("fiscal_year", None), ("fcf", D4), ("cumulative_fcf", D4), ("pv", D4),
           ("incremental_units", D4), ("scrap_savings", D4), ("ebitda", D4)], ccf)
    cin = [[k, cx_lab, f"=Capex!$F${CPX[k]}"] for k, cx_lab in [
        ("price", "Equipment purchase price ($m)"), ("install", "Installation ($m)"), ("capex", "Total investment ($m)"),
        ("add_cap", "Added capacity (m units/month)"), ("yld_gain", "Yield improvement (pts)"),
        ("savings", "Operating savings ($m/yr)"), ("maint", "Maintenance ($m/yr)"), ("life", "Useful life (years)"),
        ("rate", "Discount rate"), ("hurdle", "IRR hurdle"), ("max_pb", "Maximum payback (years)")]]
    cin += [["recommendation", "Recommendation", f"=Capex!$C${CPX['reco']}"],
            ["rule", "Decision rule", f"=Capex!$C${CPX['reco_why']}"],
            ["checkpoint", "Checkpoint", f"=Capex!$C${CPX['reco_cond']}"]]
    table("capex_inputs", "Project inputs and recommendation text",
          [("key", None), ("label", None), ("value", None)], cin)

    # ---------- scenario drivers (all three columns, scenario-independent) ----------
    drv = []
    for key in ["dem_g", "asp_chg", "yld_adj", "util", "mat_inf", "rd_g", "sga_g", "capex_pct", "tax"]:
        drv.append([key, f"=Assumptions!$A${ASM[key]}", f"=Assumptions!$B${ASM[key]}", f"=Assumptions!$C${ASM[key]}",
                    f"=Assumptions!$D${ASM[key]}", f"=Assumptions!$E${ASM[key]}"])
    table("scenario_drivers", "The nine scenario drivers, all three cases (User-provided assumptions)",
          [("driver_id", None), ("label", None), ("unit", None), ("Downside", D4), ("Base", D4), ("Upside", D4)], drv)

    # ---------- history (reported / calculated) ----------
    hist = []
    for y in YEARS:
        hist.append([y, f"={H('net_sales', y)}", f"={AA('gm', y)}", f"={AA('om', y)}", f"={AA('om_core', y)}",
                     f"={AA('fcf', y)}", f"={AA('capex_pct', y)}", f"={AA('dio', y)}", *(f"={H(p['segment'], y)}" for p in EN.products)])
    table("history", f"{EN.company['name']} reported history (Reported public data + Calculated values)",
          [("fiscal_year", None), ("net_sales", D4), ("gross_margin", D4), ("operating_margin", D4),
           ("margin_before_amortization", D4), ("fcf_project_definition", D4), ("capex_pct_sales", D4), ("dio_days", D4),
           *((f"{p['key']}_segment_sales", D4) for p in EN.products)], hist)

    # ---------- control totals ----------
    ctl = [
        ["revenue_fy", f"=PnL!${FYC}${PNL['rev']}"], ["operating_profit_fy", f"=PnL!${FYC}${PNL['op']}"],
        ["net_income_fy", f"=PnL!${FYC}${PNL['ni']}"], ["fcf_fy", f"=FCF!${FYC}${FCFT['fcf']}"],
        ["gross_profit_fy", f"=PnL!${FYC}${PNL['gp']}"], ["budget_revenue_fy", f"=Budget!$O${BUD[('div', 'rev')]}"],
        ["ytd_actual_revenue", f"=Actuals!$H${ACT[('div', 'rev')]}"], ["ytd_actual_op", f"=Actuals!$H${ACT[('div', 'op')]}"],
        ["ytd_op_variance", f"=Variance!$F${VAR['s_op']}"],
        ["capex_npv_base", f"=Capex!$F${CPX[('Base', 'npv')]}"],
    ]
    table("control_totals", "Excel control totals - the dashboard must reproduce these exactly",
          [("control_id", None), ("value", D4)], ctl)
    dx.sheet_properties.tabColor = "00B0F0"


def add_phase6_checks():
    global r
    cksec("Phase 6 - dashboard export tables")
    header_row(ck, r, ["Check", "Area", "Result", "", "", "", "", "Tolerance", "Status"])
    r += 1

    def one(label, area, formula, tol=0.0001):
        global r
        ck.cell(row=r, column=1, value=label).font = f_base
        ck.cell(row=r, column=2, value=area).font = f_base
        c = ck[f"C{r}"]
        c.value, c.font, c.number_format = "=" + formula, f_calc, '0.0000;(0.0000);"-"'
        t = ck[f"H{r}"]
        t.value, t.font = tol, f_input
        s = ck[f"I{r}"]
        s.value, s.font = f'=IF(ABS(C{r})<=H{r},"PASS","FAIL")', f_bold
        CK_ROWS.append(r)
        r += 1

    def rng(tid, name):
        colL = get_column_letter(1 + DX[tid]["cols"].index(name))
        return f"Dash_Export!${colL}${DX[tid]['first']}:${colL}${DX[tid]['last']}"

    one(f"Export monthly: 12 records, unique months, first {_mlab(FSTART)}, last {_mlab(FEND)}", "Missing / duplicate records",
        f"(ROWS({rng('monthly', 'month')})<>12)+(SUMPRODUCT(1/COUNTIF({rng('monthly', 'month')},{rng('monthly', 'month')}))<>12)"
        f"+(INDEX({rng('monthly', 'month')},1)<>DATE({FSTART.year},{FSTART.month},1))+(INDEX({rng('monthly', 'month')},12)<>DATE({FEND.year},{FEND.month},1))", 0)
    one("Export product x month: 36 records, 36 unique (month, product) keys, 12 per product", "Missing / duplicate records",
        f"(ROWS({rng('product_monthly', 'month')})<>36)"
        f"+(SUMPRODUCT(1/COUNTIFS({rng('product_monthly', 'month')},{rng('product_monthly', 'month')},"
        f"{rng('product_monthly', 'product_id')},{rng('product_monthly', 'product_id')}))<>36)"
        + "".join(f"+(COUNTIF({rng('product_monthly', 'product_id')},\"{pid}\")<>12)" for pid in PRODUCT_IDS.values()), 0)
    one("Export: no blank forecast values (monthly revenue/OP/FCF, product revenue)", "Missing / duplicate records",
        f"COUNTBLANK({rng('monthly', 'revenue')})+COUNTBLANK({rng('monthly', 'operating_profit')})"
        f"+COUNTBLANK({rng('monthly', 'fcf')})+COUNTBLANK({rng('product_monthly', 'revenue')})", 0)
    one(f"Export: actuals present for exactly 5 months ({AMON}) and labeled 'Actual'", "Dates & units",
        f"(COUNT({rng('monthly', 'actual_revenue')})<>5)+(COUNTIF({rng('monthly', 'period_type')},\"Actual\")<>5)", 0)
    one("Export monthly revenue / OP / FCF sum = PnL and FCF FY totals", "Dashboard totals vs Excel",
        f"ABS(SUM({rng('monthly', 'revenue')})-PnL!${FYC}${PNL['rev']})+ABS(SUM({rng('monthly', 'operating_profit')})-PnL!${FYC}${PNL['op']})"
        f"+ABS(SUM({rng('monthly', 'fcf')})-FCF!${FYC}${FCFT['fcf']})")
    one("Export product x month revenue sum = division FY revenue; product FY sums match", "Dashboard totals vs Excel",
        f"ABS(SUM({rng('product_monthly', 'revenue')})-PnL!${FYC}${PNL['rev']})+ABS(SUM({rng('product_fy', 'revenue')})-PnL!${FYC}${PNL['rev']})"
        f"+ABS(SUM({rng('product_fy', 'gross_profit')})-Ops_Model!${FYC}${OPS[('div', 'gp')]})")
    one("Export budget monthly sums = Budget FY; actual monthly sums = Actuals YTD", "Dashboard totals vs Excel",
        f"ABS(SUM({rng('monthly', 'budget_revenue')})-Budget!$O${BUD[('div', 'rev')]})"
        f"+ABS(SUM({rng('monthly', 'budget_operating_profit')})-Budget!$O${BUD[('div', 'op')]})"
        f"+ABS(SUM({rng('monthly', 'actual_revenue')})-Actuals!$H${ACT[('div', 'rev')]})"
        f"+ABS(SUM({rng('monthly', 'actual_operating_profit')})-Actuals!$H${ACT[('div', 'op')]})")
    one(f"Outlook = actuals ({AMON}) + forecast ({FMON})", "Dashboard totals vs Excel",
        f"ABS(SUM({rng('monthly', 'outlook_revenue')})-(Actuals!$H${ACT[('div', 'rev')]}+SUM(PnL!$H${PNL['rev']}:$N${PNL['rev']})))"
        f"+ABS(SUM({rng('monthly', 'outlook_operating_profit')})-(Actuals!$H${ACT[('div', 'op')]}+SUM(PnL!$H${PNL['op']}:$N${PNL['op']})))")
    first_b = DX["op_bridge"]["first"]
    one("Export bridge: start + deltas = end, and end = Actuals operating profit", "Variance reconciliation",
        f"ABS(Dash_Export!$D${first_b}+SUM(Dash_Export!$D${first_b + 1}:$D${first_b + 9})-Dash_Export!$D${first_b + 10})"
        f"+ABS(Dash_Export!$D${first_b + 10}-Actuals!$H${ACT[('div', 'op')]})")
    one("Export capex summary = Capex tab NPV for all scenarios", "Dashboard totals vs Excel",
        "+".join(f"ABS(INDEX({rng('capex_summary', 'npv')},{i + 1})-Capex!$F${CPX[(s, 'npv')]})" for i, (s, _) in enumerate(SCEN)))
    one("Export capex cash flows: 21 records; NPV rebuilt from exported FCF = exported NPV", "Missing / duplicate records",
        f"(ROWS({rng('capex_cashflows', 'fcf')})<>21)+"
        f"SUMPRODUCT(({rng('capex_cashflows', 'scenario')}=\"Base\")*{rng('capex_cashflows', 'fcf')}"
        f"/(1+Capex!$F${CPX['rate']})^{rng('capex_cashflows', 'year')})-INDEX({rng('capex_summary', 'npv')},2)", 0.0001)
    one("Export value ranges: margins, yields, utilization within 0-100% (violations)", "Dates & units",
        f"SUMPRODUCT(({rng('product_monthly', 'yield')}<0)+({rng('product_monthly', 'yield')}>1)"
        f"+({rng('product_monthly', 'utilization')}<0)+({rng('product_monthly', 'utilization')}>1)"
        f"+({rng('product_monthly', 'gross_margin')}<-1)+({rng('product_monthly', 'gross_margin')}>1))", 0)
    one(f"Export history: 5 fiscal years {YEARS[0]}-{YEARS[-1]}, net sales = Hist_Data", "Dates & units",
        f"(ROWS({rng('history', 'fiscal_year')})<>5)+ABS(SUM({rng('history', 'net_sales')})-SUM(Hist_Data!$C${HD['net_sales']}:$G${HD['net_sales']}))", 0.0001)
    one("Export scenario label = selector", "Scenario",
        f"IF(INDEX(Dash_Export!$B${DX['meta']['first']}:$B${DX['meta']['last']},1)=Assumptions!$C$6,0,1)", 0)
