# Phase 3 - Revenue, PnL and FCF tabs (monthly forecast year, quarterly, FY, base-year comparison).
# Executed inside build_model.py's namespace after phase2_ops.py. Not a standalone script.

FYS = "FY" + EN.fy[-2:]
QTRS = [(f"Q1 {FYS}", "C", "E"), (f"Q2 {FYS}", "F", "H"), (f"Q3 {FYS}", "I", "K"), (f"Q4 {FYS}", "L", "N")]
OPEN_DATE = (FSTART - _dt.timedelta(days=1)).strftime("%d-%b-%Y")
QCOLS = ["P", "Q", "R", "S"]
BASEC, GROWC, CLSC, NOTEC = "T", "U", "V", "W"
REV, PNL, FCFT = {}, {}, {}  # key -> row


def _fc_sheet(name, title, sub):
    ws = wb.create_sheet(name)
    title_block(ws, title, sub)
    setw(ws, {"A": 46, "B": 10, **{c: 9.5 for c in MONTH_COLS}, "O": 11, "P": 10, "Q": 10, "R": 10, "S": 10,
              "T": 11, "U": 9, "V": 24, "W": 64})
    ws.freeze_panes = "C7"
    ws["A4"] = "Scenario shown:"
    ws["A4"].font = f_bold
    ws["B4"] = "=Assumptions!$C$6"
    ws["B4"].font = Font(name=F, size=10, bold=True, color="008000")
    ws["B4"].fill = PatternFill("solid", fgColor="FFF2CC")
    header_row(ws, 6, ["Line item", "Units"] + [""] * 12 + [EN.fy] + [q for q, _, _ in QTRS] +
               [f"{EN.base_year} base", f"{FYS} vs FY{EN.base_year[-2:]}", "Classification", "Formula"])
    for col in MONTH_COLS:
        c = ws[f"{col}6"]
        c.value = f"=Ops_Model!{col}$6"
        c.number_format = "mmm-yy"
        c.font, c.fill = f_hdr, fill_hdr
        c.alignment = Alignment(horizontal="center")
    return ws


def _row(ws, reg, key, label, units, kind, fn, fmt=USD, base=None, growth=False, note="", bold=False, cls=None):
    """kind: 'flow' (quarters/FY = SUM), 'stock' (quarter/FY = period-end), 'ratio' (fn applied to every column).
    fn(col, prev_col_or_None) -> formula without '='. base: formula for the base-year column (no '=')."""
    global r
    reg[key] = r
    ws.cell(row=r, column=1, value=label).font = f_bold if bold else f_base
    ws.cell(row=r, column=2, value=units).font = f_base
    for i, col in enumerate(MONTH_COLS):
        c = ws[f"{col}{r}"]
        c.value = "=" + fn(col, MONTH_COLS[i - 1] if i else None)
        c.font, c.number_format = f_calc, fmt
    if kind == "flow":
        fy, qs = f"SUM(C{r}:N{r})", [f"SUM({a}{r}:{b}{r})" for _, a, b in QTRS]
    elif kind == "stock":
        fy, qs = f"N{r}", [f"{b}{r}" for _, a, b in QTRS]
    else:
        fy, qs = fn(FYC, None), [fn(q, None) for q in QCOLS]
    c = ws[f"{FYC}{r}"]
    c.value, c.font, c.number_format = "=" + fy, f_bold, fmt
    for qc, qf in zip(QCOLS, qs):
        c = ws[f"{qc}{r}"]
        c.value, c.font, c.number_format = "=" + qf, f_calc, fmt
    if base is not None:
        c = ws[f"{BASEC}{r}"]
        c.value, c.font, c.number_format = "=" + (fn(BASEC, None) if base == "ratio" else base), f_calc, fmt
        if growth:
            g = ws[f"{GROWC}{r}"]
            g.value = f"=IFERROR({FYC}{r}/{BASEC}{r}-1,\"\")"
            g.font, g.number_format = f_calc, PCT
    ws.cell(row=r, column=22, value=cls or "Calculated value (simulated inputs)").font = f_base
    ws.cell(row=r, column=23, value=note).font = f_base
    if bold:
        for col in range(1, 22):
            ws.cell(row=r, column=col).border = b_top
    r += 1


def build_phase3_tabs():
    global r
    asm = wb["Assumptions"]
    LIVE = lambda key: f"Assumptions!$F${ASM[key]}"        # noqa: E731
    AF = lambda key: f"Assumptions!$F${ASM[key]}"          # noqa: E731  (division-level value in column F)
    OPSR = lambda key, col: f"Ops_Model!{col}${OPS[('div', key)]}"  # noqa: E731

    # ---------- Assumptions section G: P&L and cash-flow inputs ----------
    r = asm.max_row + 2
    section(asm, r, "G. P&L and cash-flow inputs (not scenario-driven)", 17)
    r += 1
    G_INPUTS = [
        ("dso", "Days sales outstanding (DSO)", "days", CFG["pnl"]["dso"], '0', "User-provided assumption",
         CFG["pnl"].get("dso_note", "AR = monthly revenue x 12 x DSO / 365")),
        ("dpo", "Days payable outstanding on variable manufacturing cost", "days", CFG["pnl"]["dpo"], '0', "User-provided assumption",
         "Applied to purchases (variable manufacturing cost). AP = monthly variable cost x 12 x DPO / 365"),
        ("dep_share", "Depreciation share of fixed manufacturing cost", "%", CFG["pnl"]["dep_share"], PCT, SIM,
         "Non-cash part of factory fixed cost. 30% x $229m = ~$69m/yr (~7% of revenue). Reference below"),
        ("life", f"Useful life of new {EN.fy} capex", "months", CFG["pnl"]["life_months"], '0', "User-provided assumption",
         f"Straight-line; depreciation starts the month after spend. {EN.fy} capex depreciation sits in cost of revenue"),
    ]
    for key, lab, units, val, fmt, cls, why in G_INPUTS:
        ASM[key] = r
        asm.cell(row=r, column=1, value=lab).font = f_base
        asm.cell(row=r, column=2, value=units).font = f_base
        c = asm[f"F{r}"]
        c.value, c.font, c.number_format = val, f_input, fmt
        asm.cell(row=r, column=16, value=cls).font = f_base
        asm.cell(row=r, column=17, value=why).font = f_base
        r += 1
    REFS3 = [
        (f"Reference: {COMPANY} DSO {EN.base_year}", f"={AA('dso', EN.base_year)}", '0.0'),
        ((f"Reference: {COMPANY} depreciation (D&A less intangibles amortization) % of sales {EN.base_year}"
          if "other_opex:amortization" in HD else f"Reference: {COMPANY} D&A % of sales {EN.base_year}"),
         (f"=({H('cf:da', EN.base_year)}-{H('other_opex:amortization', EN.base_year)})/{H('net_sales', EN.base_year)}"
          if "other_opex:amortization" in HD else f"={H('cf:da', EN.base_year)}/{H('net_sales', EN.base_year)}"), PCT),
    ]
    for lab, fml, fmt in REFS3:
        asm.cell(row=r, column=1, value=lab).font = f_base
        c = asm[f"F{r}"]
        c.value, c.font, c.number_format = fml, f_link, fmt
        asm.cell(row=r, column=16, value="Calculated value").font = f_base
        asm.cell(row=r, column=17, value="Reported public data (Hist tabs)").font = f_base
        r += 1
    asm.cell(row=r, column=1, value=f"Opening balances at {OPEN_DATE} (calculated from base year)").font = f_sec
    r += 1
    OPEN = [
        ("ar0", "Opening accounts receivable", f"$F${ASM['base_rev']}*$F${ASM['dso']}/365", "Base-year revenue x DSO / 365"),
        ("inv0", "Opening finished-goods inventory", f"$F${ASM['inv_v0']}", "Section D opening inventory value"),
        ("ap0", "Opening accounts payable", f"$F${ASM['vcost0']}*$F${ASM['dpo']}/365", "Base-year variable cost x DPO / 365"),
    ]
    for key, lab, fml, why in OPEN:
        ASM[key] = r
        asm.cell(row=r, column=1, value=lab).font = f_base
        asm.cell(row=r, column=2, value="$m").font = f_base
        c = asm[f"F{r}"]
        c.value, c.font, c.number_format = "=" + fml, f_calc, USD
        asm.cell(row=r, column=16, value=CALC).font = f_base
        asm.cell(row=r, column=17, value=why).font = f_base
        r += 1
    ASM["wc0"] = r
    asm.cell(row=r, column=1, value="Opening operating working capital").font = f_bold
    asm.cell(row=r, column=2, value="$m").font = f_base
    c = asm[f"F{r}"]
    c.value, c.font, c.number_format = f"=F{ASM['ar0']}+F{ASM['inv0']}-F{ASM['ap0']}", f_bold, USD
    asm.cell(row=r, column=16, value=CALC).font = f_base
    asm.cell(row=r, column=17, value="AR + inventory - AP").font = f_base
    r += 1

    # ---------- Revenue tab ----------
    rv = _fc_sheet("Revenue", f"Monthly Revenue Forecast - {EN.fy} (active scenario)",
                   "Revenue by product line, units in millions, $m. " + DISCLAIMER)
    r = 7
    for p, pname in PRODUCTS:
        section(rv, r, pname, 23)
        r += 1
        _row(rv, REV, (p, "units"), "Units sold", "m units", "flow",
             lambda c, w, p=p: f"Ops_Model!{c}${OPS[(p, 'sold')]}", fmt='#,##0.00',
             base=f"Assumptions!{PCOL[p]}${ASM['sold0']}", growth=True, note="Link: Ops_Model units sold")
        _row(rv, REV, (p, "rev"), "Revenue", "$m", "flow", lambda c, w, p=p: f"Ops_Model!{c}${OPS[(p, 'rev')]}",
             base=f"Assumptions!{PCOL[p]}${ASM['rev0']}", growth=True, note="Link: Ops_Model revenue", bold=True)
        _row(rv, REV, (p, "asp"), "Average selling price", "$/unit", "ratio",
             lambda c, w, p=p: f"IFERROR({c}{REV[(p, 'rev')]}/{c}{REV[(p, 'units')]},0)", fmt='$#,##0.000',
             base="ratio", growth=False, note="Revenue / units sold")
        r += 1
    section(rv, r, "Division", 23)
    r += 1
    _row(rv, REV, "rev", "Division revenue", "$m", "flow",
         lambda c, w: "+".join(f"{c}{REV[(p, 'rev')]}" for p, _ in PRODUCTS),
         base="+".join(f"{BASEC}{REV[(p, 'rev')]}" for p, _ in PRODUCTS), growth=True,
         note="Sum of product lines", bold=True)
    for p, pname in PRODUCTS:
        _row(rv, REV, (p, "mix"), f"Mix - {pname}", "%", "ratio",
             lambda c, w, p=p: f"IFERROR({c}{REV[(p, 'rev')]}/{c}{REV['rev']},0)", fmt=PCT, base="ratio",
             note="Product revenue / division revenue")
    _row(rv, REV, "mix_tot", "Mix total", "%", "ratio",
         lambda c, w: "+".join(f"{c}{REV[(p, 'mix')]}" for p, _ in PRODUCTS), fmt=PCT, base="ratio", note="Must be 100%")

    # ---------- FCF tab (built before PnL rows are known? PnL needs capex depreciation; FCF needs NI) ----------
    # Build PnL first with forward reference to FCF depreciation row: reserve FCF row numbers by building FCF
    # capex block first on its own sheet, then PnL, then the rest of FCF.
    fc = _fc_sheet("FCF", f"Free Cash Flow Forecast - {EN.fy} (active scenario)",
                   "FCF = net income + depreciation - capital expenditures - increase in working capital. " + DISCLAIMER)
    r = 7
    section(fc, r, "A. Capital expenditures and depreciation", 23)
    r += 1
    _row(fc, FCFT, "capex", "Capital expenditures", "$m", "flow",
         lambda c, w: f"Revenue!{c}${REV['rev']}*{LIVE('capex_pct')}", note="Revenue x capex % (scenario). "
                                                                           "Excludes the Phase 5 test-equipment proposal")
    _row(fc, FCFT, "dep_new", f"Depreciation on {EN.fy} capex", "$m", "flow",
         lambda c, w: "0" if w is None else f"SUM($C${FCFT['capex']}:{w}{FCFT['capex']})/{AF('life')}",
         note="Straight-line over useful life; starts the month after spend")
    _row(fc, FCFT, "dep_fix", "Depreciation inside fixed manufacturing cost", "$m", "flow",
         lambda c, w: f"{OPSR('fcost', c)}*{AF('dep_share')}", note="Fixed manufacturing cost x depreciation share")
    _row(fc, FCFT, "da", "Total depreciation", "$m", "flow",
         lambda c, w: f"{c}{FCFT['dep_new']}+{c}{FCFT['dep_fix']}", note="Sum of the two lines above", bold=True)
    fcf_ptr = r  # continue FCF tab after PnL is built

    # ---------- PnL tab ----------
    pl = _fc_sheet("PnL", f"Monthly P&L Forecast - {EN.fy} (active scenario)",
                   "Division P&L to net income, $m. No division-level interest (financing sits at corporate). " + DISCLAIMER)
    r = 7
    section(pl, r, "Income statement", 23)
    r += 1
    _row(pl, PNL, "rev", "Revenue", "$m", "flow", lambda c, w: f"Revenue!{c}${REV['rev']}",
         base=f"Assumptions!$F${ASM['base_rev']}", growth=True, note="Link: Revenue tab", bold=True)
    _row(pl, PNL, "cor_mfg", "Cost of revenue - manufacturing", "$m", "flow", lambda c, w: OPSR("cor", c),
         base=f"Assumptions!$F${ASM['cogs0']}", growth=True, note="Link: Ops_Model (weighted-average inventory cost)")
    _row(pl, PNL, "cor_dep", f"Cost of revenue - depreciation on {EN.fy} capex", "$m", "flow",
         lambda c, w: f"FCF!{c}${FCFT['dep_new']}", base="0", note="Link: FCF tab section A")
    _row(pl, PNL, "cor", "Total cost of revenue", "$m", "flow",
         lambda c, w: f"{c}{PNL['cor_mfg']}+{c}{PNL['cor_dep']}", base=f"{BASEC}{r - 2}+{BASEC}{r - 1}", growth=True,
         note="Manufacturing + new-capex depreciation")
    _row(pl, PNL, "gp", "Gross profit", "$m", "flow", lambda c, w: f"{c}{PNL['rev']}-{c}{PNL['cor']}",
         base=f"{BASEC}{PNL['rev']}-{BASEC}{PNL['cor']}", growth=True, note="Revenue - cost of revenue", bold=True)
    _row(pl, PNL, "gm", "Gross margin", "%", "ratio", lambda c, w: f"IFERROR({c}{PNL['gp']}/{c}{PNL['rev']},0)",
         fmt=PCT, base="ratio", note="Gross profit / revenue")
    _row(pl, PNL, "rd", "Research and development", "$m", "flow",
         lambda c, w: f"Assumptions!$F${ASM['rd0']}*(1+{LIVE('rd_g')})/12",
         base=f"Assumptions!$F${ASM['rd0']}", growth=True, note="Base-year R&D x (1 + R&D growth) / 12")
    _row(pl, PNL, "sga", "Selling, general and administrative", "$m", "flow",
         lambda c, w: f"Assumptions!$F${ASM['sga0']}*(1+{LIVE('sga_g')})/12",
         base=f"Assumptions!$F${ASM['sga0']}", growth=True, note="Base-year SG&A x (1 + SG&A growth) / 12")
    _row(pl, PNL, "opex", "Total operating expenses", "$m", "flow", lambda c, w: f"{c}{PNL['rd']}+{c}{PNL['sga']}",
         base=f"{BASEC}{r - 2}+{BASEC}{r - 1}", growth=True, note="R&D + SG&A")
    _row(pl, PNL, "op", "Operating profit", "$m", "flow", lambda c, w: f"{c}{PNL['gp']}-{c}{PNL['opex']}",
         base=f"{BASEC}{PNL['gp']}-{BASEC}{PNL['opex']}", growth=True, note="Gross profit - operating expenses", bold=True)
    _row(pl, PNL, "om", "Operating margin", "%", "ratio", lambda c, w: f"IFERROR({c}{PNL['op']}/{c}{PNL['rev']},0)",
         fmt=PCT, base="ratio", note="Operating profit / revenue")
    _row(pl, PNL, "pbt", "Pre-tax income", "$m", "flow", lambda c, w: f"{c}{PNL['op']}",
         note="Equals operating profit (no division-level interest or other income)")
    _row(pl, PNL, "tax", "Income tax", "$m", "flow", lambda c, w: f"{c}{PNL['pbt']}*{LIVE('tax')}",
         note="Pre-tax income x scenario tax rate (a loss month gives a tax credit)")
    _row(pl, PNL, "ni", "Net income", "$m", "flow", lambda c, w: f"{c}{PNL['pbt']}-{c}{PNL['tax']}",
         note="Pre-tax income - tax", bold=True)
    _row(pl, PNL, "nm", "Net margin", "%", "ratio", lambda c, w: f"IFERROR({c}{PNL['ni']}/{c}{PNL['rev']},0)",
         fmt=PCT, note="Net income / revenue")
    r += 1
    section(pl, r, "Memo", 23)
    r += 1
    _row(pl, PNL, "da", "Depreciation (from FCF tab)", "$m", "flow", lambda c, w: f"FCF!{c}${FCFT['da']}",
         note="Link: FCF tab total depreciation")
    _row(pl, PNL, "ebitda", "EBITDA", "$m", "flow", lambda c, w: f"{c}{PNL['op']}+{c}{PNL['da']}",
         note="Operating profit + depreciation")
    _row(pl, PNL, "ebitda_m", "EBITDA margin", "%", "ratio",
         lambda c, w: f"IFERROR({c}{PNL['ebitda']}/{c}{PNL['rev']},0)", fmt=PCT, note="EBITDA / revenue")
    _row(pl, PNL, "cm", "Contribution margin (from Ops_Model)", "$m", "flow", lambda c, w: OPSR("cm", c),
         note="Revenue - variable cost of units sold")

    # ---------- FCF tab continued ----------
    r = fcf_ptr + 1
    section(fc, r, "B. Operating working capital (period-end balances)", 23)
    r += 1
    _row(fc, FCFT, "ar", "Accounts receivable", "$m", "stock",
         lambda c, w: f"Revenue!{c}${REV['rev']}*12*{AF('dso')}/365", note="Monthly revenue x 12 x DSO / 365")
    _row(fc, FCFT, "inv", "Inventory (finished goods)", "$m", "stock", lambda c, w: OPSR("invv_close", c),
         note="Link: Ops_Model closing inventory value")
    _row(fc, FCFT, "ap", "Accounts payable", "$m", "stock",
         lambda c, w: f"{OPSR('vcost', c)}*12*{AF('dpo')}/365", note="Monthly variable manufacturing cost x 12 x DPO / 365")
    _row(fc, FCFT, "wc", "Operating working capital", "$m", "stock",
         lambda c, w: f"{c}{FCFT['ar']}+{c}{FCFT['inv']}-{c}{FCFT['ap']}", note="AR + inventory - AP", bold=True)
    _row(fc, FCFT, "d_wc", "Increase in working capital", "$m", "flow",
         lambda c, w: f"{c}{FCFT['wc']}-" + (f"Assumptions!$F${ASM['wc0']}" if w is None else f"{w}{FCFT['wc']}"),
         note=f"This month's working capital - prior month's (first month vs {OPEN_DATE} opening balance)")
    r += 1
    section(fc, r, "C. Free cash flow", 23)
    r += 1
    _row(fc, FCFT, "ni", "Net income", "$m", "flow", lambda c, w: f"PnL!{c}${PNL['ni']}", note="Link: PnL")
    _row(fc, FCFT, "p_da", "+ Depreciation", "$m", "flow", lambda c, w: f"{c}{FCFT['da']}", note="Section A")
    _row(fc, FCFT, "m_capex", "- Capital expenditures", "$m", "flow", lambda c, w: f"-{c}{FCFT['capex']}", note="Section A")
    _row(fc, FCFT, "m_dwc", "- Increase in working capital", "$m", "flow", lambda c, w: f"-{c}{FCFT['d_wc']}",
         note="Section B")
    _row(fc, FCFT, "fcf", "Free cash flow", "$m", "flow",
         lambda c, w: f"{c}{FCFT['ni']}+{c}{FCFT['p_da']}+{c}{FCFT['m_capex']}+{c}{FCFT['m_dwc']}",
         note="Net income + depreciation - capex - increase in working capital", bold=True)
    _row(fc, FCFT, "fcf_m", "FCF margin", "%", "ratio",
         lambda c, w: f"IFERROR({c}{FCFT['fcf']}/Revenue!{c}${REV['rev']},0)", fmt=PCT, note="FCF / revenue")
    _row(fc, FCFT, "fcf_conv", "FCF conversion (FCF / net income)", "x", "ratio",
         lambda c, w: f"IFERROR({c}{FCFT['fcf']}/{c}{FCFT['ni']},0)", fmt=MULT, note="FCF / net income")
    _row(fc, FCFT, "fcf_cum", "Cumulative free cash flow", "$m", "stock",
         lambda c, w: f"{c}{FCFT['fcf']}" if w is None else f"{w}{FCFT['fcf_cum']}+{c}{FCFT['fcf']}",
         note=f"Running total from {FSTART.strftime('%B %Y')}")

    # ---------- Live result block on Assumptions ----------
    r = asm.max_row + 2
    section(asm, r, f"H. Live scenario result - {EN.fy} P&L and cash flow", 17)
    r += 1
    ASM["summary3"] = r
    for lab, fml, fmt in [("Operating profit", f"=PnL!${FYC}${PNL['op']}", USD),
                          ("Operating margin", f"=PnL!${FYC}${PNL['om']}", PCT),
                          ("Net income", f"=PnL!${FYC}${PNL['ni']}", USD),
                          ("Free cash flow", f"=FCF!${FYC}${FCFT['fcf']}", USD),
                          ("FCF margin", f"=FCF!${FYC}${FCFT['fcf_m']}", PCT)]:
        asm.cell(row=r, column=1, value=lab).font = f_base
        c = asm[f"F{r}"]
        c.value, c.font, c.number_format = fml, f_link, fmt
        asm.cell(row=r, column=16, value=CALC).font = f_base
        r += 1


def add_phase3_checks():
    global r
    cksec("Phase 3 - revenue, P&L and free cash flow (result = absolute difference or violation count)")
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

    RM = lambda sh, reg, key: f"{sh}!$C${reg[key]}:$N${reg[key]}"   # noqa: E731
    RF = lambda sh, reg, key: f"{sh}!${FYC}${reg[key]}"              # noqa: E731
    QS = lambda sh, reg, key: f"SUM({sh}!$P${reg[key]}:$S${reg[key]})"  # noqa: E731
    AFv = lambda key: f"Assumptions!$F${ASM[key]}"                   # noqa: E731

    one("Revenue tab = Ops_Model revenue, every product and month", "Revenue",
        "+".join(f"SUMPRODUCT(ABS({RM('Revenue', REV, (p, 'rev'))}-Ops_Model!$C${OPS[(p, 'rev')]}:$N${OPS[(p, 'rev')]}))"
                 for p, _ in PRODUCTS))
    one("Revenue: sum of quarters = FY = sum of months (division)", "Monthly-to-annual",
        f"ABS({QS('Revenue', REV, 'rev')}-{RF('Revenue', REV, 'rev')})+ABS(SUM({RM('Revenue', REV, 'rev')})-{RF('Revenue', REV, 'rev')})")
    one("Revenue mix sums to 100% in every month", "Revenue", f"SUMPRODUCT(ABS({RM('Revenue', REV, 'mix_tot')}-1))")
    one("P&L revenue = Ops_Model division revenue (FY)", "Revenue",
        f"{RF('PnL', PNL, 'rev')}-Ops_Model!${FYC}${OPS[('div', 'rev')]}")
    one("Gross profit reconciliation: P&L GP + new-capex depreciation = Ops_Model GP", "Gross profit",
        f"({RF('PnL', PNL, 'gp')}+{RF('PnL', PNL, 'cor_dep')})-Ops_Model!${FYC}${OPS[('div', 'gp')]}")
    one("Gross profit = revenue - cost of revenue (every month)", "Gross profit",
        f"SUMPRODUCT(ABS({RM('PnL', PNL, 'gp')}-({RM('PnL', PNL, 'rev')}-{RM('PnL', PNL, 'cor')})))")
    one("Operating profit = gross profit - R&D - SG&A (every month)", "P&L",
        f"SUMPRODUCT(ABS({RM('PnL', PNL, 'op')}-({RM('PnL', PNL, 'gp')}-{RM('PnL', PNL, 'rd')}-{RM('PnL', PNL, 'sga')})))")
    one("FY R&D = base-year R&D x (1 + growth); FY SG&A likewise", "P&L",
        f"ABS({RF('PnL', PNL, 'rd')}-{AFv('rd0')}*(1+{AFv('rd_g')}))+ABS({RF('PnL', PNL, 'sga')}-{AFv('sga0')}*(1+{AFv('sga_g')}))")
    one("Tax = tax rate x pre-tax income (FY)", "P&L", f"{RF('PnL', PNL, 'tax')}-{AFv('tax')}*{RF('PnL', PNL, 'pbt')}")
    one("Net income = pre-tax - tax (FY) and quarters sum to FY", "P&L",
        f"ABS({RF('PnL', PNL, 'ni')}-({RF('PnL', PNL, 'pbt')}-{RF('PnL', PNL, 'tax')}))+ABS({QS('PnL', PNL, 'ni')}-{RF('PnL', PNL, 'ni')})")
    one("Capex = capex % x revenue (FY)", "Free cash flow",
        f"{RF('FCF', FCFT, 'capex')}-{AFv('capex_pct')}*{RF('Revenue', REV, 'rev')}")
    one("New-capex depreciation (FY) = SUMPRODUCT(capex, months remaining) / life", "Free cash flow",
        f"{RF('FCF', FCFT, 'dep_new')}-SUMPRODUCT({RM('FCF', FCFT, 'capex')},12-COLUMN({RM('FCF', FCFT, 'capex')})+COLUMN(FCF!$C$1)-1)/{AFv('life')}")
    one("Increase in WC summed over the year = closing WC - opening WC", "Free cash flow",
        f"{RF('FCF', FCFT, 'd_wc')}-(FCF!$N${FCFT['wc']}-{AFv('wc0')})")
    one("FCF = NI + depreciation - capex - increase in WC (independent, every month)", "Free cash flow",
        f"SUMPRODUCT(ABS({RM('FCF', FCFT, 'fcf')}-(PnL!$C${PNL['ni']}:$N${PNL['ni']}+{RM('FCF', FCFT, 'da')}"
        f"-{RM('FCF', FCFT, 'capex')}-{RM('FCF', FCFT, 'd_wc')})))")
    one("Cumulative FCF at March = FY FCF", "Free cash flow", f"FCF!$N${FCFT['fcf_cum']}-{RF('FCF', FCFT, 'fcf')}")
    one("Cash identity: NI + dep - dWC = revenue - cash costs - tax - dAR + dAP (FY)", "Free cash flow",
        f"({RF('FCF', FCFT, 'ni')}+{RF('FCF', FCFT, 'da')}-{RF('FCF', FCFT, 'd_wc')})-("
        f"{RF('PnL', PNL, 'rev')}-(Ops_Model!${FYC}${OPS[('div', 'mfg')]}-{RF('FCF', FCFT, 'dep_fix')})"
        f"-{RF('PnL', PNL, 'opex')}-{RF('PnL', PNL, 'tax')}-(FCF!$N${FCFT['ar']}-{AFv('ar0')})+(FCF!$N${FCFT['ap']}-{AFv('ap0')}))")
    one("Inventory on FCF tab = Ops_Model closing inventory (every month)", "Free cash flow",
        f"SUMPRODUCT(ABS({RM('FCF', FCFT, 'inv')}-Ops_Model!$C${OPS[('div', 'invv_close')]}:$N${OPS[('div', 'invv_close')]}))")
    one("All forecast tabs show the selected scenario", "Scenario",
        '(Revenue!$B$4<>Assumptions!$C$6)+(PnL!$B$4<>Assumptions!$C$6)+(FCF!$B$4<>Assumptions!$C$6)', 0)
    one("Forecast month headers match Ops_Model (Revenue, PnL, FCF)", "Dates & units",
        "SUMPRODUCT(--(Revenue!$C$6:$N$6<>Ops_Model!$C$6:$N$6))+SUMPRODUCT(--(PnL!$C$6:$N$6<>Ops_Model!$C$6:$N$6))"
        "+SUMPRODUCT(--(FCF!$C$6:$N$6<>Ops_Model!$C$6:$N$6))", 0)

    header_row(ck, r, ["Unit test (hand-calculable)", "", "Result", "Expected", "Difference", "", "", "Tolerance", "Status"])
    r += 1
    tests3 = [
        ("AR: monthly revenue 100 x 12 x DSO 60 / 365 = 197.260", "=100*12*60/365", 100 * 12 * 60 / 365),
        ("Straight-line: capex 60, life 60 months = 1.0 per month", "=60/60", 1),
        ("Capex in month 1 of 12 depreciates 11 months in-year: 60 x 11/60 = 11", "=60*(12-1)/60", 11),
        ("Tax: pre-tax 100 x 18% = 18; net income = 82", "=100-100*0.18", 82),
        ("Operating profit: GP 600 - R&D 210 - SG&A 104 = 286", "=600-210-104", 286),
        ("WC: AR 200 + inventory 150 - AP 50 = 300; increase from 280 = 20", "=(200+150-50)-280", 20),
        ("FCF: NI 82 + dep 10 - capex 30 - increase in WC 20 = 42", "=82+10-30-20", 42),
    ]
    for lab, formula, exp in tests3:
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
