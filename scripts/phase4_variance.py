# Phase 4 - Budget (locked Base scenario), simulated Actuals (first forecast.actual_months months) and Variance tabs.
ACT_LBL = CFG["forecast"]["actuals_label"]                       # e.g. "Apr-Aug 2026"
_AEND = _dt.datetime(FSTART.year + (FSTART.month + 3) // 12, (FSTART.month + 3) % 12 + 1, 1)  # last actual month
YTD_LBL = "YTD " + FSTART.strftime("%b") + "-" + _AEND.strftime("%b")
PSH = [p["abbr"] for p in EN.products]
ACFG = CFG["actuals"]

# Executed inside build_model.py's namespace after phase3_pnl.py. Not a standalone script.

BUD, ACT, VAR = {}, {}, {}
BLOCK_ORDER = [
    # key, label, units, kind, fmt
    ("dem", "Customer demand", "m units", "flow", '#,##0.00'),
    ("cap", "Internal final-test capacity", "m units", "flow", '#,##0.00'),
    ("util", "Capacity utilization", "%", "rate:prod/cap", PCT),
    ("prod", "Units produced (tested)", "m units", "flow", '#,##0.00'),
    ("yld", "Manufacturing yield (final test)", "%", "rate:good/prod", PCT),
    ("good", "Good units produced", "m units", "flow", '#,##0.00'),
    ("inv_open", "Opening finished-goods inventory", "m units", "open", '#,##0.00'),
    ("avail", "Available good units", "m units", "none", '#,##0.00'),
    ("sold", "Units sold", "m units", "flow", '#,##0.00'),
    ("inv_close", "Closing finished-goods inventory", "m units", "close", '#,##0.00'),
    ("unmet", "Unmet demand", "m units", "flow", '#,##0.00'),
    ("asp", "Average selling price", "$/unit", "rate:rev/sold", '$#,##0.000'),
    ("rev", "Revenue", "$m", "flow", USD),
    ("vcu", "Variable cost per unit produced", "$/unit", "rate:vcost/prod", '$#,##0.000'),
    ("vcost", "Variable manufacturing cost", "$m", "flow", USD),
    ("fix", "Fixed manufacturing cost", "$m", "flow", USD),
    ("mfg", "Total manufacturing cost", "$m", "flow", USD),
    ("cpgu", "Cost per good unit", "$/unit", "rate:mfg/good", '$#,##0.000'),
    ("invv_open", "Opening inventory value", "$m", "open", USD),
    ("avgc", "Average cost per available unit", "$/unit", "none", '$#,##0.000'),
    ("cor", "Cost of revenue (manufacturing)", "$m", "flow", USD),
    ("invv_close", "Closing inventory value", "$m", "close", USD),
    ("gp", "Gross profit (before new-capex depreciation)", "$m", "flow", USD),
    ("gm", "Gross margin", "%", "rate:gp/rev", PCT),
]


def _sheet_hdr(ws, month_cols, tot_specs, cls_col, note_col):
    header_row(ws, 6, ["Line item", "Units"] + [""] * len(month_cols) + [t[1] for t in tot_specs] + ["Classification", "Formula"])
    for col in month_cols:
        c = ws[f"{col}6"]
        c.value = f"=Ops_Model!{col}$6"
        c.number_format, c.font, c.fill = "mmm-yy", f_hdr, fill_hdr
        c.alignment = Alignment(horizontal="center")


def build_block(ws, reg, p, month_cols, tot_specs, refs, cls, cls_col, note_col):
    """Operating block for one product. refs: dict of col->formula builders for dem, cap, util, yld, asp, vcu, fix,
    and strings inv_u0, inv_v0. tot_specs: list of (col, header, first, last)."""
    global r
    start = r
    for i, (key, *_rest) in enumerate(BLOCK_ORDER):
        reg[(p, key)] = start + i
    R = lambda key, c: f"{c}{reg[(p, key)]}"  # noqa: E731
    formulas = {
        "dem": lambda c, w: refs["dem"](c),
        "cap": lambda c, w: refs["cap"](c),
        "util": lambda c, w: refs["util"](c),
        "prod": lambda c, w: f"{R('cap', c)}*{R('util', c)}",
        "yld": lambda c, w: refs["yld"](c),
        "good": lambda c, w: f"{R('prod', c)}*{R('yld', c)}",
        "inv_open": lambda c, w: refs["inv_u0"] if w is None else R("inv_close", w),
        "avail": lambda c, w: f"{R('inv_open', c)}+{R('good', c)}",
        "sold": lambda c, w: f"MIN({R('dem', c)},{R('avail', c)})",
        "inv_close": lambda c, w: f"{R('avail', c)}-{R('sold', c)}",
        "unmet": lambda c, w: f"{R('dem', c)}-{R('sold', c)}",
        "asp": lambda c, w: refs["asp"](c),
        "rev": lambda c, w: f"{R('sold', c)}*{R('asp', c)}",
        "vcu": lambda c, w: refs["vcu"](c),
        "vcost": lambda c, w: f"{R('prod', c)}*{R('vcu', c)}",
        "fix": lambda c, w: refs["fix"](c),
        "mfg": lambda c, w: f"{R('vcost', c)}+{R('fix', c)}",
        "cpgu": lambda c, w: f"IFERROR({R('mfg', c)}/{R('good', c)},0)",
        "invv_open": lambda c, w: refs["inv_v0"] if w is None else R("invv_close", w),
        "avgc": lambda c, w: f"IFERROR(({R('invv_open', c)}+{R('mfg', c)})/{R('avail', c)},0)",
        "cor": lambda c, w: f"{R('sold', c)}*{R('avgc', c)}",
        "invv_close": lambda c, w: f"{R('invv_open', c)}+{R('mfg', c)}-{R('cor', c)}",
        "gp": lambda c, w: f"{R('rev', c)}-{R('cor', c)}",
        "gm": lambda c, w: f"IFERROR({R('gp', c)}/{R('rev', c)},0)",
    }
    notes = {"prod": "Capacity x utilization", "good": "Units produced x yield", "sold": "MIN(demand, available)",
             "inv_close": "Available - sold", "rev": "Units sold x ASP", "vcost": "Units produced x variable cost",
             "mfg": "Variable + fixed", "avgc": "Weighted-average cost", "cor": "Units sold x average cost",
             "invv_close": "Opening + manufacturing cost - cost of revenue", "gp": "Revenue - cost of revenue"}
    for key, label, units, kind, fmt in BLOCK_ORDER:
        rr = reg[(p, key)]
        ws.cell(row=rr, column=1, value=label).font = f_bold if key in ("sold", "rev", "gp") else f_base
        ws.cell(row=rr, column=2, value=units).font = f_base
        for i, col in enumerate(month_cols):
            c = ws[f"{col}{rr}"]
            c.value = "=" + formulas[key](col, month_cols[i - 1] if i else None)
            c.font, c.number_format = f_calc, fmt
        for tcol, _h, first, last in tot_specs:
            if kind == "flow":
                f = f"SUM({first}{rr}:{last}{rr})"
            elif kind == "open":
                f = f"{first}{rr}"
            elif kind == "close":
                f = f"{last}{rr}"
            elif kind.startswith("rate:"):
                num, den = kind[5:].split("/")
                f = f"IFERROR({tcol}{reg[(p, num)]}/{tcol}{reg[(p, den)]},0)"
            else:
                continue
            c = ws[f"{tcol}{rr}"]
            c.value, c.font, c.number_format = "=" + f, f_bold, fmt
        ws[f"{cls_col}{rr}"] = cls
        ws[f"{cls_col}{rr}"].font = f_base
        ws[f"{note_col}{rr}"] = notes.get(key, "")
        ws[f"{note_col}{rr}"].font = f_base
    r = start + len(BLOCK_ORDER)


def build_division(ws, reg, month_cols, tot_specs, capex_ref, life_ref, rd_ref, sga_ref, cls, cls_col, note_col):
    global r
    rows = [("rev", "Revenue", "flow"), ("vcost", "Variable manufacturing cost", "flow"),
            ("fix", "Fixed manufacturing cost", "flow"), ("cor_mfg", "Cost of revenue - manufacturing", "flow"),
            ("capex", "Capital expenditures", "flow"), ("dep_new", f"Depreciation on {FYL} capex", "flow"),
            ("cor", "Total cost of revenue", "flow"), ("gp", "Gross profit", "flow"), ("gm", "Gross margin", "rate:gp/rev"),
            ("rd", "Research and development", "flow"), ("sga", "Selling, general and administrative", "flow"),
            ("opex", "Total operating expenses", "flow"), ("op", "Operating profit", "flow"),
            ("om", "Operating margin", "rate:op/rev"), ("invv_close", "Closing inventory value", "close")]
    start = r
    for i, (k, *_x) in enumerate(rows):
        reg[("div", k)] = start + i
    D = lambda k, c: f"{c}{reg[('div', k)]}"  # noqa: E731
    PS = lambda k, c: "+".join(f"{c}{reg[(p, k)]}" for p, _ in PRODUCTS)  # noqa: E731
    first = month_cols[0]
    fm = {
        "rev": lambda c, w: PS("rev", c), "vcost": lambda c, w: PS("vcost", c), "fix": lambda c, w: PS("fix", c),
        "cor_mfg": lambda c, w: PS("cor", c), "capex": lambda c, w: f"{D('rev', c)}*{capex_ref}",
        "dep_new": lambda c, w: "0" if w is None else f"SUM(${first}${reg[('div', 'capex')]}:{w}{reg[('div', 'capex')]})/{life_ref}",
        "cor": lambda c, w: f"{D('cor_mfg', c)}+{D('dep_new', c)}", "gp": lambda c, w: f"{D('rev', c)}-{D('cor', c)}",
        "gm": lambda c, w: f"IFERROR({D('gp', c)}/{D('rev', c)},0)", "rd": lambda c, w: rd_ref,
        "sga": lambda c, w: sga_ref, "opex": lambda c, w: f"{D('rd', c)}+{D('sga', c)}",
        "op": lambda c, w: f"{D('gp', c)}-{D('opex', c)}", "om": lambda c, w: f"IFERROR({D('op', c)}/{D('rev', c)},0)",
        "invv_close": lambda c, w: PS("invv_close", c),
    }
    for k, label, kind in rows:
        rr = reg[("div", k)]
        ws.cell(row=rr, column=1, value=label).font = f_bold if k in ("rev", "gp", "op") else f_base
        ws.cell(row=rr, column=2, value="%" if kind.startswith("rate") else "$m").font = f_base
        fmt = PCT if kind.startswith("rate") else USD
        for i, col in enumerate(month_cols):
            c = ws[f"{col}{rr}"]
            c.value = "=" + fm[k](col, month_cols[i - 1] if i else None)
            c.font, c.number_format = f_calc, fmt
        for tcol, _h, a, b in tot_specs:
            if kind == "flow":
                f = f"SUM({a}{rr}:{b}{rr})"
            elif kind == "close":
                f = f"{b}{rr}"
            else:
                num, den = kind[5:].split("/")
                f = f"IFERROR({tcol}{reg[('div', num)]}/{tcol}{reg[('div', den)]},0)"
            c = ws[f"{tcol}{rr}"]
            c.value, c.font, c.number_format = "=" + f, f_bold, fmt
        ws[f"{cls_col}{rr}"] = cls
        ws[f"{cls_col}{rr}"].font = f_base
        if k in ("gp", "op"):
            for col in range(1, 3 + len(month_cols) + len(tot_specs)):
                ws.cell(row=rr, column=col).border = b_top
    r = start + len(rows)


def build_phase4_tabs():
    global r
    AP = lambda key, p: f"Assumptions!${PCOL[p]}${ASM[key]}"  # noqa: E731
    AF = lambda key: f"Assumptions!$F${ASM[key]}"             # noqa: E731

    # ================= Budget =================
    bu = wb.create_sheet("Budget")
    title_block(bu, f"{FYL} Budget - Base scenario, LOCKED",
                "Reads the Base column (Assumptions column D) directly, so it does NOT change with the scenario selector.")
    setw(bu, {"A": 46, "B": 10, **{c: 9.5 for c in MONTH_COLS}, "O": 11, "P": 12, "Q": 30, "R": 50})
    bu.freeze_panes = "C7"
    B_TOT = [("O", FYL, "C", "N"), ("P", YTD_LBL, "C", "G")]
    _sheet_hdr(bu, MONTH_COLS, B_TOT, "Q", "R")
    r = 8
    section(bu, r, "Budget drivers (links to the Base column of Assumptions section A)", 18)
    r += 1
    for key in ["dem_g", "asp_chg", "yld_adj", "util", "mat_inf", "rd_g", "sga_g", "capex_pct"]:
        BUD[("drv", key)] = r
        bu.cell(row=r, column=1, value=wb["Assumptions"].cell(row=ASM[key], column=1).value).font = f_base
        c = bu[f"C{r}"]
        c.value, c.font, c.number_format = f"=Assumptions!$D${ASM[key]}", f_link, PCT
        bu[f"Q{r}"] = "User-provided assumption (Base)"
        bu[f"Q{r}"].font = f_base
        r += 1
    BD = lambda key: f"$C${BUD[('drv', key)]}"  # noqa: E731
    r += 1
    for p, pname in PRODUCTS:
        section(bu, r, f"{pname} - budget  ({DISCLAIMER})", 18)
        r += 1
        refs = {
            "dem": lambda c, p=p: f"{AP('sold0', p)}*(1+{BD('dem_g')}+{AP('gdiff', p)})*Assumptions!{c}${ASM['season']}/12",
            "cap": lambda c, p=p: AP("cap0", p), "util": lambda c: BD("util"),
            "yld": lambda c, p=p: f"MIN(1,MAX(0,{AP('yld0', p)}+{BD('yld_adj')}))",
            "asp": lambda c, p=p: f"{AP('asp0', p)}*(1+{BD('asp_chg')})",
            "vcu": lambda c, p=p: f"{AP('vc0', p)}*(1+{BD('mat_inf')})", "fix": lambda c, p=p: AP("fix0", p),
            "inv_u0": AP("inv_u0", p), "inv_v0": AP("inv_v0", p),
        }
        build_block(bu, BUD, p, MONTH_COLS, B_TOT, refs, "Calculated value (budget)", "Q", "R")
        r += 1
    section(bu, r, f"Division budget P&L to operating profit  ({DISCLAIMER})", 18)
    r += 1
    build_division(bu, BUD, MONTH_COLS, B_TOT, BD("capex_pct"), AF("life"),
                   f"{AF('rd0')}*(1+{BD('rd_g')})/12", f"{AF('sga0')}*(1+{BD('sga_g')})/12",
                   "Calculated value (budget)", "Q", "R")

    # ================= Actuals =================
    ac = wb.create_sheet("Actuals")
    A_MONTHS = MONTH_COLS[:5]  # C..G = the five actual months (engine requires actual_months = 5)
    title_block(ac, f"Actual Results - {ACT_LBL} (SIMULATED)",
                "Simulated actuals for the budget-vs-actual exercise. Same operating logic as the budget, different inputs.")
    ac["A3"] = DISCLAIMER + "  ALL VALUES ON THIS TAB ARE SIMULATED."
    ac["A3"].font = Font(name=F, size=10, bold=True, color="C00000")
    setw(ac, {"A": 46, "B": 12, "C": 12, "D": 12, "E": 12, "F": 12, "G": 12, "H": 12, "I": 30, "J": 60})
    ac.freeze_panes = "C7"
    A_TOT = [("H", YTD_LBL, "C", "G")]
    _sheet_hdr(ac, A_MONTHS, A_TOT, "I", "J")
    r = 8
    section(ac, r, "Simulated actual inputs", 10)
    r += 1
    header_row(ac, r, ["Input", "Units", *PSH, "", "", "", "Classification", "Note"])
    r += 1
    AN = ACFG.get("notes", {})
    A_IN = [(k, lab, u, {p["key"]: p[k] for p in EN.products}, fmt, AN.get(k, "")) for k, lab, u, fmt in [
        ("a_dem", "Demand vs budget", "%", PCT),
        ("a_asp", f"ASP change vs {EN.base_year} base-year ASP", "%", PCT),
        ("a_yld", "Actual final-test yield", "%", PCT),
        ("a_util", "Actual capacity utilization", "%", PCT),
        ("a_vc", f"Variable cost change vs {EN.base_year} base-year cost", "%", PCT),
        ("a_fix", "Actual fixed manufacturing cost per month", "$m", USD)]]
    for key, lab, units, vals, fmt, why in A_IN:
        ACT[("in", key)] = r
        ac.cell(row=r, column=1, value=lab).font = f_base
        ac.cell(row=r, column=2, value=units).font = f_base
        for p, _ in PRODUCTS:
            c = ac[f"{PCOL[p]}{r}"]
            c.value, c.font, c.number_format = vals[p], f_sim, fmt
        ac[f"I{r}"] = SIM
        ac[f"J{r}"] = why
        r += 1
    for key, lab, units, val, fmt, why in [
            ("a_rd", "Actual R&D expense per month", "$m", ACFG["a_rd"], USD, AN.get("a_rd", "")),
            ("a_sga", "Actual SG&A expense per month", "$m", ACFG["a_sga"], USD, AN.get("a_sga", "")),
            ("a_capex", "Actual capex % of revenue", "%", ACFG["a_capex"], PCT, AN.get("a_capex", ""))]:
        ACT[("in", key)] = r
        ac.cell(row=r, column=1, value=lab).font = f_base
        ac.cell(row=r, column=2, value=units).font = f_base
        c = ac[f"C{r}"]
        c.value, c.font, c.number_format = val, f_sim, fmt
        ac[f"I{r}"] = SIM
        ac[f"J{r}"] = why
        r += 1
    ACT[("in", "noise")] = r
    ac.cell(row=r, column=1, value="Monthly demand noise (applies to all products)").font = f_base
    ac.cell(row=r, column=2, value="x").font = f_base
    for col, v in zip(A_MONTHS, ACFG["noise"]):
        c = ac[f"{col}{r}"]
        c.value, c.font, c.number_format = v, f_sim, '0.00'
    ac[f"I{r}"] = SIM
    ac[f"J{r}"] = f"Row is by month ({YTD_LBL[4:]}), not by product"
    r += 2
    AIN = lambda key, p: f"${PCOL[p]}${ACT[('in', key)]}"  # noqa: E731
    for p, pname in PRODUCTS:
        section(ac, r, f"{pname} - actual  ({DISCLAIMER})", 10)
        r += 1
        refs = {
            "dem": lambda c, p=p: f"Budget!{c}{BUD[(p, 'dem')]}*(1+{AIN('a_dem', p)})*{c}${ACT[('in', 'noise')]}",
            "cap": lambda c, p=p: AP("cap0", p), "util": lambda c, p=p: AIN("a_util", p),
            "yld": lambda c, p=p: AIN("a_yld", p), "asp": lambda c, p=p: f"{AP('asp0', p)}*(1+{AIN('a_asp', p)})",
            "vcu": lambda c, p=p: f"{AP('vc0', p)}*(1+{AIN('a_vc', p)})", "fix": lambda c, p=p: AIN("a_fix", p),
            "inv_u0": AP("inv_u0", p), "inv_v0": AP("inv_v0", p),
        }
        build_block(ac, ACT, p, A_MONTHS, A_TOT, refs, SIM, "I", "J")
        r += 1
    section(ac, r, f"Division actual P&L to operating profit  ({DISCLAIMER})", 10)
    r += 1
    build_division(ac, ACT, A_MONTHS, A_TOT, f"$C${ACT[('in', 'a_capex')]}", AF("life"),
                   f"$C${ACT[('in', 'a_rd')]}", f"$C${ACT[('in', 'a_sga')]}", SIM, "I", "J")

    # ================= Variance =================
    va = wb.create_sheet("Variance")
    title_block(va, f"Budget vs Actual - YTD {ACT_LBL}",
                "Favorable = positive (adds to profit). Components are computed separately and must reconcile to the totals.")
    va["A3"] = DISCLAIMER + "  (Actuals are simulated.)"
    va["A3"].font = Font(name=F, size=9, italic=True, color="C00000")
    setw(va, {"A": 56, "B": 10, "C": 13, "D": 13, "E": 13, "F": 13, "G": 13, "H": 80})
    va.freeze_panes = "C6"
    BY = lambda key, p: f"Budget!$P${BUD[(p, key)]}"   # noqa: E731  budget YTD
    AY = lambda key, p: f"Actuals!$H${ACT[(p, key)]}"  # noqa: E731  actual YTD
    r = 5

    # --- A. YTD summary
    section(va, r, "A. YTD P&L summary", 8)
    r += 1
    header_row(va, r, ["Line", "Units", "Budget", "Actual", "Variance", "Variance %", "Fav / Unfav", "Definition"])
    r += 1

    def srow(key, label, bud, act, cost=False, fmt=USD, pct=False, bold=False):
        global r
        VAR[key] = r
        va.cell(row=r, column=1, value=label).font = f_bold if bold else f_base
        va.cell(row=r, column=2, value="%" if pct else "$m").font = f_base
        va[f"C{r}"], va[f"D{r}"] = "=" + bud, "=" + act
        va[f"C{r}"].font = va[f"D{r}"].font = f_link
        va[f"E{r}"] = f"=C{r}-D{r}" if cost else f"=D{r}-C{r}"
        for col in "CDE":
            va[f"{col}{r}"].number_format = PCT if pct else fmt
        va[f"E{r}"].font = f_bold if bold else f_calc
        if not pct:
            va[f"F{r}"] = f"=IFERROR(E{r}/ABS(C{r}),0)"
            va[f"F{r}"].number_format = PCT
            va[f"G{r}"] = f'=IF(ABS(E{r})<0.0005,"On budget",IF(E{r}>0,"Favorable","Unfavorable"))'
        va[f"H{r}"] = ("Budget - actual (cost: lower actual = favorable)" if cost else "Actual - budget")
        r += 1

    for p, pname in PRODUCTS:
        srow(f"rev_{p}", f"Revenue - {pname}", BY("rev", p), AY("rev", p))
    srow("rev", "Total revenue", f"Budget!$P${BUD[('div', 'rev')]}", f"Actuals!$H${ACT[('div', 'rev')]}", bold=True)
    srow("cor", "Cost of revenue", f"Budget!$P${BUD[('div', 'cor')]}", f"Actuals!$H${ACT[('div', 'cor')]}", cost=True)
    srow("gp", "Gross profit", f"Budget!$P${BUD[('div', 'gp')]}", f"Actuals!$H${ACT[('div', 'gp')]}", bold=True)
    srow("gm", "Gross margin", f"Budget!$P${BUD[('div', 'gm')]}", f"Actuals!$H${ACT[('div', 'gm')]}", pct=True)
    srow("rd", "Research and development", f"Budget!$P${BUD[('div', 'rd')]}", f"Actuals!$H${ACT[('div', 'rd')]}", cost=True)
    srow("sga", "Selling, general and administrative", f"Budget!$P${BUD[('div', 'sga')]}",
         f"Actuals!$H${ACT[('div', 'sga')]}", cost=True)
    srow("opex", "Total operating expenses", f"Budget!$P${BUD[('div', 'opex')]}", f"Actuals!$H${ACT[('div', 'opex')]}",
         cost=True)
    srow("op", "Operating profit", f"Budget!$P${BUD[('div', 'op')]}", f"Actuals!$H${ACT[('div', 'op')]}", bold=True)
    srow("om", "Operating margin", f"Budget!$P${BUD[('div', 'om')]}", f"Actuals!$H${ACT[('div', 'om')]}", pct=True)

    # --- KPIs
    r += 1
    section(va, r, "B. Operating KPIs, YTD (budget vs actual)", 8)
    r += 1
    header_row(va, r, ["KPI", "Units", "Budget", "Actual", "Difference", "", "", "Definition"])
    r += 1
    KPI = [("sold", "Units sold", "m units", '#,##0.00'), ("asp", "Average selling price", "$/unit", '$#,##0.000'),
           ("yld", "Manufacturing yield", "%", PCT), ("util", "Capacity utilization", "%", PCT),
           ("cpgu", "Cost per good unit", "$/unit", '$#,##0.000'), ("unmet", "Unmet demand", "m units", '#,##0.00')]
    for key, lab, units, fmt in KPI:
        for p, pname in PRODUCTS:
            VAR[(f"kpi_{key}", p)] = r
            va.cell(row=r, column=1, value=f"{lab} - {pname}").font = f_base
            va.cell(row=r, column=2, value=units).font = f_base
            va[f"C{r}"], va[f"D{r}"] = "=" + BY(key, p), "=" + AY(key, p)
            va[f"E{r}"] = f"=D{r}-C{r}"
            for col in "CDE":
                va[f"{col}{r}"].number_format = fmt
            va[f"C{r}"].font = va[f"D{r}"].font = f_link
            va[f"H{r}"] = "Actual - budget"
            r += 1

    # --- C. driver table per product
    r += 1
    section(va, r, "C. Variance inputs by product (YTD)", 8)
    r += 1
    header_row(va, r, ["Input", "Units", *PSH, "Division", "", "Definition"])
    r += 1

    def drow(key, label, units, fn, div=None, fmt=USD, note=""):
        global r
        VAR[key] = r
        va.cell(row=r, column=1, value=label).font = f_base
        va.cell(row=r, column=2, value=units).font = f_base
        for p, _ in PRODUCTS:
            c = va[f"{PCOL[p]}{r}"]
            c.value, c.number_format = "=" + fn(p), fmt
            c.font = f_link if fn(p).startswith(("Budget!", "Actuals!")) else f_calc
        if div:
            c = va[f"F{r}"]
            c.value, c.font, c.number_format = "=" + div, f_calc, fmt
        va[f"H{r}"] = note
        r += 1

    V = lambda key, p: f"{PCOL[p]}{VAR[key]}"  # noqa: E731
    U3 = '#,##0.000'
    drow("Ub", "Units sold - budget (Ub)", "m units", lambda p: BY("sold", p), f"SUM(C{r}:E{r})", U3)
    drow("Ua", "Units sold - actual (Ua)", "m units", lambda p: AY("sold", p), f"SUM(C{r}:E{r})", U3)
    drow("mixb", "Mix - budget (Ub / total Ub)", "%", lambda p: f"{V('Ub', p)}/$F${VAR['Ub']}", f"SUM(C{r}:E{r})", PCT)
    drow("mixa", "Mix - actual (Ua / total Ua)", "%", lambda p: f"{V('Ua', p)}/$F${VAR['Ua']}", f"SUM(C{r}:E{r})", PCT)
    drow("ASPb", "ASP - budget", "$/unit", lambda p: BY("asp", p), f"Budget!$P${BUD[('div', 'rev')]}/F{VAR['Ub']}",
         '$#,##0.0000', "Division = budget revenue / total budget units (average price at budget mix)")
    drow("ASPa", "ASP - actual", "$/unit", lambda p: AY("asp", p), None, '$#,##0.0000')
    drow("cb", "Cost of revenue per unit sold - budget (cb)", "$/unit",
         lambda p: f"{BY('cor', p)}/{V('Ub', p)}", None, '$#,##0.0000', "Budget manufacturing cost of revenue / budget units")
    drow("ca", "Cost of revenue per unit sold - actual (ca)", "$/unit",
         lambda p: f"{AY('cor', p)}/{V('Ua', p)}", None, '$#,##0.0000')
    drow("GPub", "Gross profit per unit - budget (ASPb - cb)", "$/unit", lambda p: f"{V('ASPb', p)}-{V('cb', p)}",
         None, '$#,##0.0000')
    drow("Pb", "Units produced - budget (Pb)", "m units", lambda p: BY("prod", p), None, U3)
    drow("Pa", "Units produced - actual (Pa)", "m units", lambda p: AY("prod", p), None, U3)
    drow("Ga", "Good units - actual (Ga)", "m units", lambda p: AY("good", p), None, U3)
    drow("yb", "Yield - budget (yb)", "%", lambda p: BY("yld", p), None, PCT)
    drow("vcb", "Variable cost per unit produced - budget", "$/unit", lambda p: BY("vcu", p), None, '$#,##0.0000')
    drow("vca", "Variable cost per unit produced - actual", "$/unit", lambda p: AY("vcu", p), None, '$#,##0.0000')
    drow("Fb", "Fixed manufacturing cost - budget", "$m", lambda p: BY("fix", p), f"SUM(C{r}:E{r})")
    drow("Fa", "Fixed manufacturing cost - actual", "$m", lambda p: AY("fix", p), f"SUM(C{r}:E{r})")
    drow("CIb", "Closing inventory value (Aug) - budget", "$m", lambda p: BY("invv_close", p), f"SUM(C{r}:E{r})")
    drow("CIa", "Closing inventory value (Aug) - actual", "$m", lambda p: AY("invv_close", p), f"SUM(C{r}:E{r})")

    # --- D. revenue decomposition
    r += 1
    section(va, r, "D. Revenue variance: price, volume, mix", 8)
    r += 1
    header_row(va, r, ["Component", "Units", *PSH, "Division", "", "Formula"])
    r += 1
    TUa, TUb = f"$F${VAR['Ua']}", f"$F${VAR['Ub']}"
    drow("r_price", "Price variance", "$m", lambda p: f"{V('Ua', p)}*({V('ASPa', p)}-{V('ASPb', p)})",
         f"SUM(C{r}:E{r})", note="Ua x (ASPa - ASPb)")
    drow("r_vol", "Volume variance", "$m", lambda p: f"{V('mixb', p)}*({TUa}-{TUb})*{V('ASPb', p)}",
         f"SUM(C{r}:E{r})", note="Budget mix x (total Ua - total Ub) x ASPb.  Division = change in total units x budget average price")
    drow("r_mix", "Product-mix variance", "$m", lambda p: f"({V('mixa', p)}-{V('mixb', p)})*{TUa}*{V('ASPb', p)}",
         f"SUM(C{r}:E{r})", note="(actual mix - budget mix) x total Ua x ASPb")
    drow("r_expl", "Sum of components", "$m", lambda p: f"{V('r_price', p)}+{V('r_vol', p)}+{V('r_mix', p)}",
         f"SUM(C{r}:E{r})")
    drow("r_tot", "Total revenue variance (actual - budget, computed directly)", "$m",
         lambda p: f"{AY('rev', p)}-{BY('rev', p)}", f"Actuals!$H${ACT[('div', 'rev')]}-Budget!$P${BUD[('div', 'rev')]}")
    VAR["r_unrec"] = r
    va.cell(row=r, column=1, value="Unreconciled difference (division)").font = f_bold
    va[f"F{r}"] = f"=F{VAR['r_tot']}-F{VAR['r_expl']}"
    va[f"F{r}"].number_format = '0.000000'
    va[f"G{r}"] = f'=IF(ABS(F{r})>0.001,"UNRECONCILED","Reconciled")'
    va[f"G{r}"].font = f_bold
    r += 1
    va.cell(row=r, column=1, value="Note: the product columns of volume and mix are an allocation; only their sum per product "
                                   "((Ua - Ub) x ASPb) and the division totals are unique.").font = f_sub
    r += 2

    # --- E. operating-profit bridge
    section(va, r, "E. Operating-profit bridge (waterfall): budget to actual", 8)
    r += 1
    header_row(va, r, ["Step", "Units", *PSH, "Division", "", "Formula (favorable = positive)"])
    r += 1
    VAR["b_start"] = r
    va.cell(row=r, column=1, value="Budget operating profit (YTD)").font = f_bold
    va[f"F{r}"] = f"=Budget!$P${BUD[('div', 'op')]}"
    va[f"F{r}"].font, va[f"F{r}"].number_format = f_link, USD
    r += 1
    VAR["b_first"] = r
    drow("b_price", "Price", "$m", lambda p: V("r_price", p), f"SUM(C{r}:E{r})", note="Same as revenue price variance")
    drow("b_vol", "Sales volume", "$m", lambda p: f"{V('mixb', p)}*({TUa}-{TUb})*{V('GPub', p)}", f"SUM(C{r}:E{r})",
         note="Budget mix x (total Ua - total Ub) x budget gross profit per unit")
    drow("b_mix", "Product mix", "$m", lambda p: f"({V('mixa', p)}-{V('mixb', p)})*{TUa}*{V('GPub', p)}",
         f"SUM(C{r}:E{r})", note="(actual mix - budget mix) x total Ua x budget gross profit per unit")
    drow("b_mat", "Material cost", "$m", lambda p: f"-{V('Pa', p)}*({V('vca', p)}-{V('vcb', p)})", f"SUM(C{r}:E{r})",
         note="-Pa x (actual - budget variable cost per unit)")
    drow("b_yld", "Manufacturing yield", "$m", lambda p: f"-{V('vcb', p)}*({V('Pa', p)}-{V('Ga', p)}/{V('yb', p)})",
         f"SUM(C{r}:E{r})", note="-budget variable cost x (units started - units needed at budget yield for actual good output)")
    drow("b_fix", "Fixed manufacturing spending (incl. new-capex depreciation)", "$m",
         lambda p: f"-({V('Fa', p)}-{V('Fb', p)})",
         f"SUM(C{r}:E{r})-(Actuals!$H${ACT[('div', 'dep_new')]}-Budget!$P${BUD[('div', 'dep_new')]})",
         note="-(actual - budget fixed cost); division also includes the new-capex depreciation difference")
    drow("b_abs", "Production volume, absorption & inventory", "$m",
         lambda p: f"-({V('vcb', p)}*({V('Ga', p)}/{V('yb', p)}-{V('Pb', p)})-({V('CIa', p)}-{V('CIb', p)})"
                   f"-({V('Ua', p)}-{V('Ub', p)})*{V('cb', p)})",
         f"SUM(C{r}:E{r})",
         note="-[budget var. cost x (standard starts - Pb) - (closing inventory a - b) - (Ua - Ub) x cb]: "
              "cost of producing more/less than sold, held in inventory")
    drow("b_rd", "R&D", "$m", lambda p: "0", f"-(Actuals!$H${ACT[('div', 'rd')]}-Budget!$P${BUD[('div', 'rd')]})",
         note="-(actual - budget R&D)")
    drow("b_sga", "SG&A", "$m", lambda p: "0", f"-(Actuals!$H${ACT[('div', 'sga')]}-Budget!$P${BUD[('div', 'sga')]})",
         note="-(actual - budget SG&A)")
    VAR["b_last"] = r - 1
    VAR["b_end"] = r
    va.cell(row=r, column=1, value="Actual operating profit - from bridge").font = f_bold
    va[f"F{r}"] = f"=F{VAR['b_start']}+SUM(F{VAR['b_first']}:F{VAR['b_last']})"
    va[f"F{r}"].font, va[f"F{r}"].number_format = f_bold, USD
    r += 1
    VAR["b_act"] = r
    va.cell(row=r, column=1, value="Actual operating profit - from Actuals tab").font = f_bold
    va[f"F{r}"] = f"=Actuals!$H${ACT[('div', 'op')]}"
    va[f"F{r}"].font, va[f"F{r}"].number_format = f_link, USD
    r += 1
    VAR["b_unrec"] = r
    va.cell(row=r, column=1, value="Unreconciled difference").font = f_bold
    va[f"F{r}"] = f"=F{VAR['b_act']}-F{VAR['b_end']}"
    va[f"F{r}"].number_format = '0.000000'
    va[f"G{r}"] = f'=IF(ABS(F{r})>0.001,"UNRECONCILED","Reconciled")'
    va[f"G{r}"].font = f_bold
    r += 2
    section(va, r, "Required variance summary (division, YTD, favorable = positive)", 8)
    r += 1
    VAR["sum_first"] = r
    for key, lab, fml in [
            ("s_rev", "Total revenue variance", f"F{VAR['r_tot']}"), ("s_price", "Price variance", f"F{VAR['r_price']}"),
            ("s_vol", "Volume variance (revenue)", f"F{VAR['r_vol']}"), ("s_mix", "Product-mix variance (revenue)", f"F{VAR['r_mix']}"),
            ("s_mat", "Material-cost variance", f"F{VAR['b_mat']}"), ("s_yld", "Manufacturing-yield variance", f"F{VAR['b_yld']}"),
            ("s_opex", "Operating-expense variance", f"F{VAR['b_rd']}+F{VAR['b_sga']}"),
            ("s_op", "Operating-profit variance", f"F{VAR['b_act']}-F{VAR['b_start']}")]:
        VAR[key] = r
        va.cell(row=r, column=1, value=lab).font = f_bold if key in ("s_rev", "s_op") else f_base
        va[f"F{r}"] = "=" + fml
        va[f"F{r}"].number_format = USD
        va[f"G{r}"] = f'=IF(ABS(F{r})<0.0005,"On budget",IF(F{r}>0,"Favorable","Unfavorable"))'
        r += 1

    # --- F. monthly trend
    r += 1
    section(va, r, f"F. Monthly budget vs actual ({ACT_LBL})", 8)
    r += 1
    header_row(va, r, ["Line", "Units", "", "", "", "", "", "Formula"])
    for col in A_MONTHS:
        va[f"{col}{r}"] = f"=Ops_Model!{col}$6"
        va[f"{col}{r}"].number_format, va[f"{col}{r}"].font, va[f"{col}{r}"].fill = "mmm-yy", f_hdr, fill_hdr
    VAR["m_hdr"] = r
    r += 1
    for key, lab, bref, aref in [("m_rev_b", "Revenue - budget", "rev", None), ("m_rev_a", "Revenue - actual", None, "rev"),
                                 ("m_op_b", "Operating profit - budget", "op", None),
                                 ("m_op_a", "Operating profit - actual", None, "op")]:
        VAR[key] = r
        va.cell(row=r, column=1, value=lab).font = f_base
        va.cell(row=r, column=2, value="$m").font = f_base
        for col in A_MONTHS:
            va[f"{col}{r}"] = f"=Budget!{col}{BUD[('div', bref)]}" if bref else f"=Actuals!{col}{ACT[('div', aref)]}"
            va[f"{col}{r}"].number_format, va[f"{col}{r}"].font = USD, f_link
        r += 1
    for key, lab, a, b in [("m_rev_v", "Revenue variance", "m_rev_a", "m_rev_b"),
                           ("m_op_v", "Operating-profit variance", "m_op_a", "m_op_b")]:
        VAR[key] = r
        va.cell(row=r, column=1, value=lab).font = f_bold
        va.cell(row=r, column=2, value="$m").font = f_base
        for col in A_MONTHS:
            va[f"{col}{r}"] = f"={col}{VAR[a]}-{col}{VAR[b]}"
            va[f"{col}{r}"].number_format, va[f"{col}{r}"].font = USD, f_bold
        r += 1


def add_phase4_checks():
    global r
    cksec("Phase 4 - budget vs actual and variance decomposition")
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

    bd0, bd1 = BUD[("drv", "dem_g")], BUD[("drv", "capex_pct")]
    one("Budget drivers = Base column of Assumptions (independent of selector)", "Scenario",
        "+".join(f"ABS(Budget!$C${BUD[('drv', k)]}-Assumptions!$D${ASM[k]})"
                 for k in ["dem_g", "asp_chg", "yld_adj", "util", "mat_inf", "rd_g", "sga_g", "capex_pct"]))
    one("When Base is selected: Budget = Ops_Model (revenue and cost of revenue, every product and month)", "Scenario",
        'IF(Assumptions!$C$6="Base",' + "+".join(
            f"SUMPRODUCT(ABS(Budget!$C${BUD[(p, k)]}:$N${BUD[(p, k)]}-Ops_Model!$C${OPS[(p, k)]}:$N${OPS[(p, k)]}))"
            for p, _ in PRODUCTS for k in ("rev", "cor")) + ",0)")
    one("When Base is selected: budget operating profit = PnL operating profit (FY)", "Scenario",
        f'IF(Assumptions!$C$6="Base",Budget!$O${BUD[("div", "op")]}-PnL!${FYC}${PNL["op"]},0)')
    for tab, reg, cols in [("Budget", BUD, ("C", "N")), ("Actuals", ACT, ("C", "G"))]:
        a, b = cols
        M = lambda p, k: f"{tab}!${a}${reg[(p, k)]}:${b}${reg[(p, k)]}"  # noqa: E731
        one(f"{tab}: units sold = MIN(demand, available); produced = capacity x utilization; good = produced x yield",
            "Operations", "+".join(
                f"SUMPRODUCT(ABS({M(p, 'sold')}-(({M(p, 'dem')}<{M(p, 'avail')})*{M(p, 'dem')}+({M(p, 'dem')}>={M(p, 'avail')})*{M(p, 'avail')})))"
                f"+SUMPRODUCT(ABS({M(p, 'prod')}-{M(p, 'cap')}*{M(p, 'util')}))+SUMPRODUCT(ABS({M(p, 'good')}-{M(p, 'prod')}*{M(p, 'yld')}))"
                for p, _ in PRODUCTS))
        one(f"{tab}: inventory value rolls forward (opening + mfg cost - cost of revenue = closing)", "Operations",
            "+".join(f"ABS({tab}!${a}${reg[(p, 'invv_open')]}+SUM({M(p, 'mfg')})-SUM({M(p, 'cor')})-{tab}!${b}${reg[(p, 'invv_close')]})"
                     for p, _ in PRODUCTS))
        one(f"{tab}: no negative inventory (violations)", "Operations",
            "+".join(f"SUMPRODUCT(--({M(p, 'inv_close')}<-0.000001))" for p, _ in PRODUCTS), 0)
    one(f"Budget YTD revenue (Variance tab) = SUM of Budget {YTD_LBL[4:]} product revenue", "Budget vs actual",
        f"Variance!$C${VAR['rev']}-(" + "+".join(f"SUM(Budget!$C${BUD[(p, 'rev')]}:$G${BUD[(p, 'rev')]})" for p, _ in PRODUCTS) + ")")
    one(f"Actual YTD revenue (Variance tab) = SUM of Actuals {YTD_LBL[4:]} product revenue", "Budget vs actual",
        f"Variance!$D${VAR['rev']}-(" + "+".join(f"SUM(Actuals!$C${ACT[(p, 'rev')]}:$G${ACT[(p, 'rev')]})" for p, _ in PRODUCTS) + ")")
    one("Budget and actual YTD operating profit = SUM of monthly operating profit", "Budget vs actual",
        f"ABS(Variance!$C${VAR['op']}-SUM(Budget!$C${BUD[('div', 'op')]}:$G${BUD[('div', 'op')]}))"
        f"+ABS(Variance!$D${VAR['op']}-SUM(Actuals!$C${ACT[('div', 'op')]}:$G${ACT[('div', 'op')]}))")
    one("Revenue: price + volume + mix = total revenue variance (division)", "Variance reconciliation",
        f"Variance!$F${VAR['r_unrec']}")
    one("Revenue: volume + mix per product = (Ua - Ub) x ASPb", "Variance reconciliation",
        "+".join(f"ABS(Variance!{PCOL[p]}{VAR['r_vol']}+Variance!{PCOL[p]}{VAR['r_mix']}-(Variance!{PCOL[p]}{VAR['Ua']}"
                 f"-Variance!{PCOL[p]}{VAR['Ub']})*Variance!{PCOL[p]}{VAR['ASPb']})" for p, _ in PRODUCTS))
    one("Operating-profit bridge components reconcile to actual operating profit", "Variance reconciliation",
        f"Variance!$F${VAR['b_unrec']}")
    one("Unit-cost components (material + yield + fixed + absorption) = -Ua x (ca - cb), per product", "Variance reconciliation",
        "+".join(f"ABS(Variance!{PCOL[p]}{VAR['b_mat']}+Variance!{PCOL[p]}{VAR['b_yld']}+Variance!{PCOL[p]}{VAR['b_abs']}"
                 f"+Variance!{PCOL[p]}{VAR['b_fix']}+Variance!{PCOL[p]}{VAR['Ua']}*(Variance!{PCOL[p]}{VAR['ca']}-Variance!{PCOL[p]}{VAR['cb']}))"
                 for p, _ in PRODUCTS))
    one("Operating-profit variance = revenue + cost-of-revenue + opex variances (favorable-positive signs)", "Variance reconciliation",
        f"Variance!$F${VAR['s_op']}-(Variance!$E${VAR['rev']}+Variance!$E${VAR['cor']}+Variance!$E${VAR['opex']})")
    one("Monthly variances sum to YTD (revenue and operating profit)", "Monthly-to-annual",
        f"ABS(SUM(Variance!$C${VAR['m_rev_v']}:$G${VAR['m_rev_v']})-Variance!$E${VAR['rev']})"
        f"+ABS(SUM(Variance!$C${VAR['m_op_v']}:$G${VAR['m_op_v']})-Variance!$E${VAR['op']})")
    one(f"Actuals months = first five budget months ({ACT_LBL})", "Dates & units",
        f"SUMPRODUCT(--(Actuals!$C$6:$G$6<>Budget!$C$6:$G$6))+(Actuals!$C$6<>DATE({FSTART.year},{FSTART.month},1))+(Actuals!$G$6<>DATE({_AEND.year},{_AEND.month},1))", 0)
    one("Budget and actual start from the same opening inventory", "Operations",
        "+".join(f"ABS(Budget!$C${BUD[(p, 'invv_open')]}-Actuals!$C${ACT[(p, 'invv_open')]})" for p, _ in PRODUCTS))

    header_row(ck, r, ["Unit test (hand-calculable)", "", "Result", "Expected", "Difference", "", "", "Tolerance", "Status"])
    r += 1
    # Textbook: budget A 100u @ $10, B 100u @ $5 (rev 1,500); actual A 120u @ $9, B 90u @ $5 (rev 1,530)
    tests4 = [
        ("PVM example - price: 120 x ($9 - $10) + 90 x ($5 - $5) = -120", "=120*(9-10)+90*(5-5)", -120),
        ("PVM example - volume: (210 - 200) x ($1,500 / 200) = +75", "=(210-200)*(1500/200)", 75),
        ("PVM example - mix: (120/210 - 0.5) x 210 x $10 + (90/210 - 0.5) x 210 x $5 = +75",
         "=(120/210-0.5)*210*10+(90/210-0.5)*210*5", 75),
        ("PVM example - components sum to total variance: 1,530 - 1,500 = 30", "=-120+75+75", 30),
        ("Yield: var. cost $1.00, budget yield 90%, 90 good from 105 started -> -$1 x (105 - 100) = -5",
         "=-1*(105-90/0.9)", -5),
        ("Material: 105 produced x ($1.10 - $1.00) -> -10.5", "=-105*(1.1-1)", -10.5),
        ("Opex: actual 90 vs budget 87.5 -> -2.5 unfavorable", "=-(90-87.5)", -2.5),
    ]
    for lab, formula, exp in tests4:
        ck.cell(row=r, column=1, value=lab).font = f_base
        ck[f"C{r}"] = formula
        ck[f"D{r}"] = exp
        ck[f"D{r}"].font = f_input
        ck[f"E{r}"] = f"=C{r}-D{r}"
        ck[f"H{r}"] = 0.000001
        ck[f"H{r}"].font = f_input
        ck[f"I{r}"] = f'=IF(ABS(E{r})<=H{r},"PASS","FAIL")'
        ck[f"I{r}"].font = f_bold
        CK_ROWS.append(r)
        r += 1
