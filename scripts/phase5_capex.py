# Phase 5 - Capital-investment analysis: hypothetical equipment project on the capex.target product line.
# Executed inside build_model.py's namespace after phase4_variance.py. Not a standalone script.

CPX = {}
SCEN = [("Downside", "C"), ("Base", "D"), ("Upside", "E")]  # scenario -> Assumptions column
YCOLS = ["C", "D", "E", "F", "G", "H", "I"]  # Year 0..6
HYP = "Hypothetical project assumption"
CX = CFG["capex"]
TP = EN.target                                  # target product (dict)
TL = CX.get("target_label", TP["id"])          # how the target product is named in labels, e.g. "FPGA"
TCOL = PCOL[TP["key"]]                          # its column on Assumptions
FY0 = int(FYL[2:])


def build_phase5_tab():
    global r
    cx = wb.create_sheet("Capex")
    title_block(cx, f"Capital Investment Decision - {CX.get('project_title', CX['project'])} (hypothetical)",
                "Incremental with-vs-without-project cash flows, evaluated in all three scenarios. $m unless noted.")
    cx["A3"] = DISCLAIMER + "  The equipment proposal is hypothetical."
    cx["A3"].font = Font(name=F, size=9, italic=True, color="C00000")
    setw(cx, {"A": 58, "B": 12, "C": 12, "D": 12, "E": 12, "F": 12, "G": 12, "H": 12, "I": 12, "J": 26, "K": 70})
    cx.freeze_panes = "B5"
    AP = lambda key, col: f"Assumptions!${col}${ASM[key]}"   # noqa: E731
    FP = lambda key: f"Assumptions!${TCOL}${ASM[key]}"       # target product column on Assumptions  # noqa: E731

    r = 5
    section(cx, r, "A. Project inputs (same in every scenario)", 11)
    r += 1
    _addp = CX["add_cap"] / TP["cap0"]
    INPUTS = [
        ("price", "Equipment purchase price", "$m", CX["price"], USD, CX.get("price_note", "")),
        ("install", "Installation, qualification and software", "$m", CX["install"], USD, "Capitalized with the equipment"),
        ("add_cap", "Additional final-test capacity", "m units/month", CX["add_cap"], '#,##0.00',
         f"+{_addp:.0%} on the {TL} line's {TP['cap0']:.2f}m units/month (Assumptions D)"),
        ("yld_gain", f"Expected {TL} yield improvement", "pts", CX["yld_gain"], PCT,
         f"Better contacting and thermal control; applies to ALL {TL} units tested"),
        ("savings", "Operating-cost savings", "$m/yr", CX["savings"], USD, CX.get("savings_note", "")),
        ("maint", "Maintenance cost of new equipment", "$m/yr", CX["maint"], USD, "Service contract and spares"),
        ("life", "Useful life (and tax depreciation period)", "years", CX["life"], '0', "Straight-line"),
        ("salvage", "Salvage value at end of life", "$m", CX["salvage"], USD, "Conservative: none" if not CX["salvage"] else ""),
        ("umax", "Maximum practical utilization", "%", CX["umax"], PCT, "Test floors cannot run at 100% (changeovers, maintenance)"),
        ("erosion", f"Annual {TL} ASP erosion after {FYL}", "%", CX["erosion"], PCT, "Price declines each year; unit cost held flat"),
        ("rate", "Discount rate (division cost of capital)", "%", CX["rate"], PCT, CX.get("rate_note", "")),
        ("hurdle", "Decision criterion: minimum IRR (hurdle)", "%", CX["hurdle"], PCT,
         f"Discount rate + {round((CX['hurdle'] - CX['rate']) * 100):g} pts risk premium"),
        ("max_pb", "Decision criterion: maximum payback", "years", CX["max_pb"], '0.0', "Equipment-payback norm; shorter than useful life"),
    ]
    for key, lab, units, val, fmt, why in INPUTS:
        CPX[key] = r
        cx.cell(row=r, column=1, value=lab).font = f_base
        cx.cell(row=r, column=2, value=units).font = f_base
        c = cx[f"F{r}"]
        c.value, c.font, c.number_format = val, f_input, fmt
        if key in ("rate", "hurdle", "max_pb"):
            c.fill = fill_key
        cx[f"J{r}"] = HYP if key not in ("rate", "hurdle", "max_pb") else "User-provided assumption"
        cx[f"K{r}"] = why
        r += 1
    CPX["capex"] = r
    cx.cell(row=r, column=1, value="Total capital investment (Year 0)").font = f_bold
    cx.cell(row=r, column=2, value="$m").font = f_base
    cx[f"F{r}"] = f"=F{CPX['price']}+F{CPX['install']}"
    cx[f"F{r}"].font, cx[f"F{r}"].number_format = f_bold, USD
    cx[f"J{r}"] = "Calculated value"
    r += 1
    CPX["dso"] = r
    cx.cell(row=r, column=1, value="DSO for incremental receivables (link)").font = f_base
    cx.cell(row=r, column=2, value="days").font = f_base
    cx[f"F{r}"] = f"=Assumptions!$F${ASM['dso']}"
    cx[f"F{r}"].font, cx[f"F{r}"].number_format = f_link, '0'
    cx[f"J{r}"] = "Link: Assumptions G"
    r += 2

    section(cx, r, "B. Scenario inputs (read from the Downside / Base / Upside columns - all three are evaluated)", 11)
    r += 1
    header_row(cx, r, ["Input", "Units", "Downside", "Base", "Upside", "", "", "", "", "Classification", "Source"])
    r += 1

    def srow(key, label, units, fn, fmt, cls, note, vals=None):
        global r
        CPX[key] = r
        cx.cell(row=r, column=1, value=label).font = f_base
        cx.cell(row=r, column=2, value=units).font = f_base
        for i, (sname, acol) in enumerate(SCEN):
            c = cx.cell(row=r, column=3 + i)
            if vals:
                c.value, c.font = vals[i], f_input
            else:
                c.value, c.font = "=" + fn(acol, get_column_letter(3 + i)), f_link
            c.number_format = fmt
        cx[f"J{r}"], cx[f"K{r}"] = cls, note
        r += 1

    srow("lg", f"{TL} demand growth per year after {FYL}", "%", None, PCT, HYP,
         CX.get("growth_note", ""), vals=CX["growth_after"])
    srow("dem27", f"{TL} demand {FYL}", "m units",
         lambda a, s: f"{FP('sold0')}*(1+{AP('dem_g', a)}+{FP('gdiff')})", '#,##0.000', CALC,
         f"Base-year {TL} units x (1 + scenario demand growth + {TL} differential)")
    srow("asp27", f"{TL} ASP {FYL}", "$/unit", lambda a, s: f"{FP('asp0')}*(1+{AP('asp_chg', a)})", '$#,##0.00', CALC,
         "Base-year ASP x (1 + scenario ASP change)")
    srow("vc", f"{TL} variable cost per unit tested", "$/unit", lambda a, s: f"{FP('vc0')}*(1+{AP('mat_inf', a)})",
         '$#,##0.000', CALC, "Base-year cost x (1 + scenario material inflation)")
    srow("y_old", f"{TL} yield without project", "%", lambda a, s: f"MIN(1,MAX(0,{FP('yld0')}+{AP('yld_adj', a)}))", PCT,
         CALC, "Base-year yield + scenario yield change")
    srow("y_new", f"{TL} yield with project", "%", lambda a, s: f"MIN(1,{s}{CPX['y_old']}+$F${CPX['yld_gain']})", PCT,
         CALC, "Yield without project + improvement")
    srow("tax", "Tax rate", "%", lambda a, s: AP("tax", a), PCT, "User-provided assumption", "Assumptions A (scenario)")
    srow("cap_wo", "Good-unit capacity without project", "m units/yr",
         lambda a, s: f"{FP('cap0')}*12*$F${CPX['umax']}*{s}{CPX['y_old']}", '#,##0.000', CALC,
         "Existing capacity x 12 x max utilization x yield")
    srow("cap_wi", "Good-unit capacity with project", "m units/yr",
         lambda a, s: f"({FP('cap0')}+$F${CPX['add_cap']})*12*$F${CPX['umax']}*{s}{CPX['y_new']}", '#,##0.000', CALC,
         "(Existing + added capacity) x 12 x max utilization x new yield")
    srow("full_units", "Good units from the added capacity at 100% use", "m units/yr",
         lambda a, s: f"$F${CPX['add_cap']}*12*{s}{CPX['y_new']}", '#,##0.000', CALC,
         "Added capacity x 12 x new yield - denominator for break-even utilization")
    r += 1

    # ---------- C. annual cash-flow blocks ----------
    CF_ROWS = [
        ("t", "Year", "#", '0'),
        ("fy", "Fiscal year", "", '@'),
        ("dem", f"{TL} customer demand", "m units", '#,##0.000'),
        ("asp", f"{TL} ASP", "$/unit", '$#,##0.00'),
        ("sold_wo", "Units sold without project", "m units", '#,##0.000'),
        ("sold_wi", "Units sold with project", "m units", '#,##0.000'),
        ("d_units", "Incremental units sold", "m units", '#,##0.000'),
        ("d_rev", "Incremental revenue", "$m", USD),
        ("contrib", "Contribution from incremental units", "$m", USD),
        ("scrap_u", "Scrap reduction (units no longer scrapped)", "m units", '#,##0.000'),
        ("scrap", "Scrap reduction ($)", "$m", USD),
        ("sav", "Operating-cost savings", "$m", USD),
        ("mnt", "Maintenance cost", "$m", USD),
        ("ebitda", "Incremental EBITDA", "$m", USD),
        ("dep", "Tax depreciation", "$m", USD),
        ("ebit", "Incremental operating profit (EBIT)", "$m", USD),
        ("tax", "Tax on incremental EBIT", "$m", USD),
        ("nopat", "Incremental after-tax operating profit", "$m", USD),
        ("wc", "Incremental working capital (receivables)", "$m", USD),
        ("dwc", "Increase in working capital (recovered in final year)", "$m", USD),
        ("inv", "Capital investment / salvage", "$m", USD),
        ("fcf", "Incremental free cash flow", "$m", USD),
        ("df", "Discount factor", "x", '0.0000'),
        ("pv", "Present value of free cash flow", "$m", USD),
        ("cum", "Cumulative free cash flow", "$m", USD),
        ("pb_help", "Payback helper (year in which cumulative turns positive)", "years", '0.00'),
        ("vol_cf", "Volume-related after-tax cash flow (for break-even)", "$m", USD),
        ("full_cf", "After-tax cash flow if added capacity fully used (for break-even)", "$m", USD),
    ]
    NOTES = {
        "sold_wo": "MIN(demand, capacity without)", "sold_wi": "MIN(demand, capacity with)",
        "d_rev": "Incremental units x ASP", "contrib": "Incremental units x (ASP - variable cost / new yield)",
        "scrap_u": "Units sold without x (1/old yield - 1/new yield)", "scrap": "Scrap units x variable cost",
        "ebitda": "Contribution + scrap reduction + savings - maintenance", "dep": "(Investment - salvage) / life",
        "ebit": "EBITDA - depreciation", "tax": "EBIT x tax rate (a loss offsets other division profit)",
        "nopat": "EBIT - tax", "wc": "Incremental revenue x DSO / 365",
        "dwc": "Change in incremental WC; final year releases the balance",
        "inv": "Year 0: -investment; final year: + salvage", "fcf": "After-tax profit + depreciation - increase in WC + investment",
        "df": "1 / (1 + discount rate) ^ year", "pv": "FCF x discount factor", "cum": "Running total of FCF",
        "pb_help": "(year - 1) + unrecovered amount / that year's FCF, only in the crossover year",
        "vol_cf": "Contribution x (1 - tax) - WC effect of incremental revenue",
        "full_cf": "Full added units x (ASP - var. cost / new yield) x (1 - tax) - WC effect",
    }
    for si, (sname, acol) in enumerate(SCEN):
        sc = get_column_letter(3 + si)  # this scenario's column in section B
        S = lambda key: f"${sc}${CPX[key]}"  # noqa: E731
        IN = lambda key: f"$F${CPX[key]}"   # noqa: E731
        section(cx, r, f"C{si + 1}. Incremental cash flows - {sname.upper()} scenario", 11)
        r += 1
        header_row(cx, r, ["Line", "Units", "Year 0", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5", "Year 6",
                           "Classification", "Formula"])
        r += 1
        start = r
        for i, (k, *_x) in enumerate(CF_ROWS):
            CPX[(sname, k)] = start + i
        R = lambda k, c: f"{c}{CPX[(sname, k)]}"  # noqa: E731
        LASTC = YCOLS[-1]
        for k, lab, units, fmt in CF_ROWS:
            rr = CPX[(sname, k)]
            cx.cell(row=rr, column=1, value=lab).font = f_bold if k in ("fcf", "ebitda") else f_base
            cx.cell(row=rr, column=2, value=units).font = f_base
            for yi, col in enumerate(YCOLS):
                prev = YCOLS[yi - 1] if yi else None
                y0 = yi == 0
                if k == "t":
                    f = str(yi)
                elif k == "fy":
                    f = f'"FY"&({FY0}+{col}{CPX[(sname, "t")]})'
                elif k in ("dem", "asp", "sold_wo", "sold_wi", "d_units", "d_rev", "contrib", "scrap_u", "scrap",
                           "sav", "mnt", "ebitda", "dep", "ebit", "tax", "nopat", "wc", "vol_cf", "full_cf") and y0:
                    f = "0"
                elif k == "dem":
                    f = f"{S('dem27')}*(1+{S('lg')})^{R('t', col)}"
                elif k == "asp":
                    f = f"{S('asp27')}*(1-{IN('erosion')})^{R('t', col)}"
                elif k == "sold_wo":
                    f = f"MIN({R('dem', col)},{S('cap_wo')})"
                elif k == "sold_wi":
                    f = f"MIN({R('dem', col)},{S('cap_wi')})"
                elif k == "d_units":
                    f = f"{R('sold_wi', col)}-{R('sold_wo', col)}"
                elif k == "d_rev":
                    f = f"{R('d_units', col)}*{R('asp', col)}"
                elif k == "contrib":
                    f = f"{R('d_units', col)}*({R('asp', col)}-{S('vc')}/{S('y_new')})"
                elif k == "scrap_u":
                    f = f"{R('sold_wo', col)}*(1/{S('y_old')}-1/{S('y_new')})"
                elif k == "scrap":
                    f = f"{R('scrap_u', col)}*{S('vc')}"
                elif k == "sav":
                    f = IN("savings")
                elif k == "mnt":
                    f = f"-{IN('maint')}"
                elif k == "ebitda":
                    f = f"{R('contrib', col)}+{R('scrap', col)}+{R('sav', col)}+{R('mnt', col)}"
                elif k == "dep":
                    f = f"IF({R('t', col)}<={IN('life')},({IN('capex')}-{IN('salvage')})/{IN('life')},0)"
                elif k == "ebit":
                    f = f"{R('ebitda', col)}-{R('dep', col)}"
                elif k == "tax":
                    f = f"{R('ebit', col)}*{S('tax')}"
                elif k == "nopat":
                    f = f"{R('ebit', col)}-{R('tax', col)}"
                elif k == "wc":
                    f = f"IF({R('t', col)}={IN('life')},0,{R('d_rev', col)}*{IN('dso')}/365)"
                elif k == "dwc":
                    f = "0" if y0 else f"{R('wc', col)}-{R('wc', prev)}"
                elif k == "inv":
                    f = f"-{IN('capex')}" if y0 else f"IF({R('t', col)}={IN('life')},{IN('salvage')},0)"
                elif k == "fcf":
                    f = f"{R('nopat', col)}+{R('dep', col)}-{R('dwc', col)}+{R('inv', col)}"
                elif k == "df":
                    f = f"1/(1+{IN('rate')})^{R('t', col)}"
                elif k == "pv":
                    f = f"{R('fcf', col)}*{R('df', col)}"
                elif k == "cum":
                    f = R("fcf", col) if y0 else f"{R('cum', prev)}+{R('fcf', col)}"
                elif k == "pb_help":
                    f = "0" if y0 else (f"IF(AND({R('cum', prev)}<0,{R('cum', col)}>=0),"
                                        f"{R('t', prev)}+(-{R('cum', prev)})/{R('fcf', col)},0)")
                elif k == "vol_cf":
                    # contribution after tax minus WC effect of the incremental revenue (same WC logic as above)
                    wc_now = f"IF({R('t', col)}={IN('life')},0,{R('d_rev', col)}*{IN('dso')}/365)"
                    wc_prev = "0" if yi == 1 else f"IF({R('t', prev)}={IN('life')},0,{R('d_rev', prev)}*{IN('dso')}/365)"
                    f = f"{R('contrib', col)}*(1-{S('tax')})-(({wc_now})-({wc_prev}))"
                elif k == "full_cf":
                    fr = lambda c: f"{S('full_units')}*{R('asp', c)}"  # noqa: E731
                    wc_now = f"IF({R('t', col)}={IN('life')},0,{fr(col)}*{IN('dso')}/365)"
                    wc_prev = "0" if yi == 1 else f"IF({R('t', prev)}={IN('life')},0,{fr(prev)}*{IN('dso')}/365)"
                    f = (f"{S('full_units')}*({R('asp', col)}-{S('vc')}/{S('y_new')})*(1-{S('tax')})"
                         f"-(({wc_now})-({wc_prev}))")
                c = cx[f"{col}{rr}"]
                c.value = "=" + f
                c.font, c.number_format = (f_bold if k == "fcf" else f_calc), fmt
            cx[f"J{rr}"] = "Calculated value (hypothetical inputs)"
            cx[f"K{rr}"] = NOTES.get(k, "")
            if k == "fcf":
                for col in range(1, 10):
                    cx.cell(row=rr, column=col).border = b_top
        r = start + len(CF_ROWS)
        # results for this scenario (row numbers assigned first so formulas can reference each other)
        RES_KEYS = ["npv", "irr", "pb", "npv_novol", "pv_full", "be_util", "impl_util"]
        for i, k in enumerate(RES_KEYS):
            CPX[(sname, k)] = r + i
        FR = lambda k: CPX[(sname, k)]  # noqa: E731
        res = [
            ("npv", "NPV at discount rate", f"C{FR('fcf')}+NPV({IN('rate')},D{FR('fcf')}:{LASTC}{FR('fcf')})", USD,
             "Year-0 cash flow + NPV(rate, years 1-6)"),
            ("irr", "IRR", f"IFERROR(IRR(C{FR('fcf')}:{LASTC}{FR('fcf')}),\"n.m.\")", PCT,
             "Rate at which NPV = 0 (n.m. if cash flows never recover the investment)"),
            ("pb", "Payback period", f"IF(SUM(D{FR('pb_help')}:{LASTC}{FR('pb_help')})=0,"
                                     f"\"Not within life\",SUM(D{FR('pb_help')}:{LASTC}{FR('pb_help')}))",
             '0.00 "yrs"', "Years until cumulative FCF turns positive (interpolated)"),
            ("npv_novol", "NPV with no incremental volume",
             f"F{FR('npv')}-SUMPRODUCT(D{FR('vol_cf')}:{LASTC}{FR('vol_cf')},D{FR('df')}:{LASTC}{FR('df')})", USD,
             "NPV minus PV of volume-related cash flow: what yield, scrap, savings and maintenance alone are worth"),
            ("pv_full", "PV of cash flow if added capacity fully used",
             f"SUMPRODUCT(D{FR('full_cf')}:{LASTC}{FR('full_cf')},D{FR('df')}:{LASTC}{FR('df')})", USD,
             "PV of selling every good unit the added capacity can produce"),
            ("be_util", "Break-even utilization of the added capacity",
             f"IF(F{FR('npv_novol')}>=0,0,IFERROR(-F{FR('npv_novol')}/F{FR('pv_full')},\"n.m.\"))", PCT,
             "Average share of the added capacity that must be sold for NPV = 0 (0% = pays back on yield and savings alone)"),
            ("impl_util", "Implied average use of the added capacity in this scenario",
             f"IFERROR(SUM(D{FR('d_units')}:{LASTC}{FR('d_units')})/({S('full_units')}*{IN('life')}),0)", PCT,
             "Incremental units / (full added units x years). Can exceed 100% because the yield gain also lifts old-tester output"),
        ]
        for k, lab, fml, fmt, note in res:
            rr = CPX[(sname, k)]
            cx.cell(row=rr, column=1, value=lab).font = f_bold
            c = cx[f"F{rr}"]
            c.value, c.font, c.number_format = "=" + fml, f_bold, fmt
            cx[f"J{rr}"] = "Calculated value"
            cx[f"K{rr}"] = note
        r += len(res) + 1

    # ---------- D. decision ----------
    section(cx, r, "D. Decision criteria and recommendation", 11)
    r += 1
    header_row(cx, r, ["Criterion", "Threshold", "Downside", "Base", "Upside", "LIVE", "", "", "", "", "Rule"])
    r += 1
    RES = lambda k, s: f"$F${CPX[(s, k)]}"  # noqa: E731

    def crow(key, label, thresh, fn, fmt, rule):
        global r
        CPX[key] = r
        cx.cell(row=r, column=1, value=label).font = f_base
        if thresh:
            cx[f"B{r}"] = "=" + thresh
            cx[f"B{r}"].number_format, cx[f"B{r}"].font = fmt, f_link
        for i, (sname, _) in enumerate(SCEN):
            c = cx.cell(row=r, column=3 + i, value="=" + fn(sname))
            c.font = f_calc
            c.number_format = fmt if key.startswith("v_") else "General"
        cx[f"F{r}"] = f"=INDEX(C{r}:E{r},1,Assumptions!$D$6)"
        cx[f"F{r}"].font, cx[f"F{r}"].fill = f_bold, PatternFill("solid", fgColor="FFF2CC")
        cx[f"F{r}"].number_format = fmt if key.startswith("v_") else "General"
        cx[f"K{r}"] = rule
        r += 1

    crow("v_npv", "NPV ($m)", None, lambda s: RES("npv", s), USD, "")
    crow("v_irr", "IRR", None, lambda s: RES("irr", s), PCT, "")
    crow("v_pb", "Payback (years)", None, lambda s: RES("pb", s), '0.00', "")
    crow("v_be", "Break-even utilization of added capacity", None, lambda s: RES("be_util", s), PCT, "")
    crow("c_npv", "1. NPV > 0 at the discount rate", f"$F${CPX['rate']}",
         lambda s: f'IF({RES("npv", s)}>0,"PASS","FAIL")', PCT, "Creates value at the division's cost of capital")
    crow("c_irr", "2. IRR >= hurdle rate", f"$F${CPX['hurdle']}",
         lambda s: f'IF(ISNUMBER({RES("irr", s)}),IF({RES("irr", s)}>={"$F$" + str(CPX["hurdle"])},"PASS","FAIL"),"FAIL")',
         PCT, "Return clears the risk-adjusted hurdle")
    crow("c_pb", "3. Payback <= maximum", f"$F${CPX['max_pb']}",
         lambda s: f'IF(ISNUMBER({RES("pb", s)}),IF({RES("pb", s)}<={"$F$" + str(CPX["max_pb"])},"PASS","FAIL"),"FAIL")',
         '0.0', "Capital is recovered well within the equipment's life")
    crow("c_all", "Meets all three criteria?", None,
         lambda s: f'IF(AND({get_column_letter(3 + [x for x, _ in SCEN].index(s))}{CPX["c_npv"]}="PASS",'
                   f'{get_column_letter(3 + [x for x, _ in SCEN].index(s))}{CPX["c_irr"]}="PASS",'
                   f'{get_column_letter(3 + [x for x, _ in SCEN].index(s))}{CPX["c_pb"]}="PASS"),"YES","NO")',
         "General", "")
    r += 1
    CPX["reco"] = r
    cx.cell(row=r, column=1, value="RECOMMENDATION").font = Font(name=F, size=11, bold=True, color=NAVY)
    base_ok, down_ok, up_ok = f"D{CPX['c_all']}", f"C{CPX['c_all']}", f"E{CPX['c_all']}"
    cx[f"C{r}"] = (f'=IF({base_ok}="NO","REJECT - the Base case does not meet the criteria",'
                   f'IF({down_ok}="YES","APPROVE - meets the criteria in all three scenarios",'
                   f'"APPROVE WITH A DEMAND CHECKPOINT - meets the criteria in Base"&IF({up_ok}="YES"," and Upside","")'
                   f'&", fails in Downside"))')
    cx[f"C{r}"].font = Font(name=F, size=11, bold=True)
    cx[f"C{r}"].fill = fill_key
    r += 1
    CPX["reco_why"] = r
    cx[f"C{r}"] = (f'="Rule: reject if Base fails; approve outright only if Downside also passes; otherwise approve with a checkpoint. '
                   f'Downside NPV "&TEXT(C{CPX["v_npv"]},"$#,##0.0")&"m: with {TL} demand growth of "&TEXT($C${CPX["lg"]},"0%")'
                   f'&" a year, demand never exceeds existing capacity, so only yield, scrap and savings benefits remain."')
    cx[f"C{r}"].font = f_base
    r += 1
    CPX["reco_cond"] = r
    cx[f"C{r}"] = (f'="Checkpoint: release the purchase order once trailing {TL} demand exceeds "'
                   f'&TEXT($D${CPX["cap_wo"]}/12*0.95,"0.00")&"m good units a month (95% of current good-unit capacity at max utilization)."')
    cx[f"C{r}"].font = f_base
    r += 1


def add_phase5_checks():
    global r
    cksec("Phase 5 - capital investment analysis")
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

    rng = lambda s, k: f"Capex!$C${CPX[(s, k)]}:$I${CPX[(s, k)]}"  # noqa: E731
    rate = f"Capex!$F${CPX['rate']}"
    for s, _ in SCEN:
        one(f"{s}: NPV = SUMPRODUCT(FCF, 1/(1+r)^t) (independent of NPV())", "NPV / IRR",
            f"Capex!$F${CPX[(s, 'npv')]}-SUMPRODUCT({rng(s, 'fcf')},1/(1+{rate})^{rng(s, 't')})", 0.0001)
        one(f"{s}: NPV at the IRR = 0 (skipped if IRR n.m.)", "NPV / IRR",
            f"IF(ISNUMBER(Capex!$F${CPX[(s, 'irr')]}),SUMPRODUCT({rng(s, 'fcf')},1/(1+Capex!$F${CPX[(s, 'irr')]})^{rng(s, 't')}),0)",
            0.001)
        one(f"{s}: incremental EBITDA = (rev with - rev without) - (var. cost with - without) + savings - maintenance",
            "Cash flows",
            f"SUMPRODUCT(ABS(Capex!$D${CPX[(s, 'ebitda')]}:$I${CPX[(s, 'ebitda')]}-("
            f"(Capex!$D${CPX[(s, 'sold_wi')]}:$I${CPX[(s, 'sold_wi')]}-Capex!$D${CPX[(s, 'sold_wo')]}:$I${CPX[(s, 'sold_wo')]})"
            f"*Capex!$D${CPX[(s, 'asp')]}:$I${CPX[(s, 'asp')]}"
            f"-Capex!${get_column_letter(3 + [x for x, _ in SCEN].index(s))}${CPX['vc']}*("
            f"Capex!$D${CPX[(s, 'sold_wi')]}:$I${CPX[(s, 'sold_wi')]}/Capex!${get_column_letter(3 + [x for x, _ in SCEN].index(s))}${CPX['y_new']}"
            f"-Capex!$D${CPX[(s, 'sold_wo')]}:$I${CPX[(s, 'sold_wo')]}/Capex!${get_column_letter(3 + [x for x, _ in SCEN].index(s))}${CPX['y_old']})"
            f"+Capex!$F${CPX['savings']}-Capex!$F${CPX['maint']})))")
        one(f"{s}: units sold with >= without, and <= demand (violations)", "Cash flows",
            f"SUMPRODUCT(--(Capex!$D${CPX[(s, 'sold_wi')]}:$I${CPX[(s, 'sold_wi')]}<Capex!$D${CPX[(s, 'sold_wo')]}:$I${CPX[(s, 'sold_wo')]}-0.000001))"
            f"+SUMPRODUCT(--(Capex!$D${CPX[(s, 'sold_wi')]}:$I${CPX[(s, 'sold_wi')]}>Capex!$D${CPX[(s, 'dem')]}:$I${CPX[(s, 'dem')]}+0.000001))", 0)
        one(f"{s}: depreciation over life = investment - salvage; working capital fully recovered", "Cash flows",
            f"ABS(SUM({rng(s, 'dep')})-(Capex!$F${CPX['capex']}-Capex!$F${CPX['salvage']}))+ABS(SUM({rng(s, 'dwc')}))")
        one(f"{s}: cumulative FCF in Year 6 = sum of FCF", "Cash flows",
            f"Capex!$I${CPX[(s, 'cum')]}-SUM({rng(s, 'fcf')})")
    one(f"Capex scenario inputs read the matching Assumptions columns ({FYL} {TL} demand, Base = Ops_Model when Base selected)",
        "Scenario", f'IF(Assumptions!$C$6="Base",Capex!$D${CPX["dem27"]}-Assumptions!${TCOL}${ASM["dem27"]},0)')
    one("Tax rates in Capex = Assumptions scenario tax rates", "Scenario",
        f"ABS(Capex!$C${CPX['tax']}-Assumptions!$C${ASM['tax']})+ABS(Capex!$D${CPX['tax']}-Assumptions!$D${ASM['tax']})"
        f"+ABS(Capex!$E${CPX['tax']}-Assumptions!$E${ASM['tax']})")
    one("Live decision column = selected scenario", "Scenario",
        f"Capex!$F${CPX['v_npv']}-INDEX(Capex!$C${CPX['v_npv']}:$E${CPX['v_npv']},1,Assumptions!$D$6)")

    header_row(ck, r, ["Unit test (hand-calculable)", "", "Result", "Expected", "Difference", "", "", "Tolerance", "Status"])
    r += 1
    tests5 = [
        ("NPV: -100 now, +60 in years 1 and 2 at 10% = 4.132", "=-100+NPV(0.1,60,60)", -100 + 60 / 1.1 + 60 / 1.21),
        ("IRR of -100, +60, +60 = 13.066%", "=IRR({-100,60,60})", 0.130662386),
        ("Payback: -100, +60, +60 -> 1 + 40/60 = 1.667 years", "=1+40/60", 1 + 40 / 60),
        ("Discount factor year 3 at 9% = 0.7722", "=1/1.09^3", 1 / 1.09 ** 3),
        ("Break-even: NPV without volume -20, PV at full use 80 -> 25%", "=-(-20)/80", 0.25),
        ("Scrap reduction: 10m units, yield 85% -> 88%: 10 x (1/0.85 - 1/0.88) = 0.4011m", "=10*(1/0.85-1/0.88)",
         10 * (1 / 0.85 - 1 / 0.88)),
        ("Tax shield: $27m / 6 years x 18% = $0.81m per year", "=27/6*0.18", 0.81),
    ]
    for lab, formula, exp in tests5:
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
