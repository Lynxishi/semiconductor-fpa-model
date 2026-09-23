"""
Build the reported historical dataset for Texas Instruments (TXN) -> companies/TXN/reported_financials.csv.

Every value was transcribed from the SEC XBRL-rendered statement pages (R-pages) of TI's Form 10-K / 10-Q filings,
fetched 2026-09-22. Nothing is estimated or filled; a line TI does not print is simply absent.
Units: USD millions. Fiscal year = calendar year (FY2025 = Jan-Dec 2025).

Sign convention: values as printed, with two documented exceptions stored with the sign of their effect on pre-tax
income so the engine's identities hold: "Interest and debt expense" (printed positive, stored negative) and
"Restructuring charges/other" (printed as a charge; a printed credit such as (124) in 2024 is stored as -124).
Cash-flow items are stored as printed (outflows negative).
"""
import csv
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT = HERE / "reported_financials.csv"
EDGAR = "https://www.sec.gov/Archives/edgar/data/97476/"
SRC = {
    "K25": ("Form 10-K FY2025 (period ended 2025-12-31), filed 2026-02-06", "0000097476-26-000059", EDGAR + "000009747626000059/"),
    "K24": ("Form 10-K FY2024 (period ended 2024-12-31), filed 2025-02-14", "0000097476-25-000007", EDGAR + "000009747625000007/"),
    "K23": ("Form 10-K FY2023 (period ended 2023-12-31), filed 2024-02-02", "0000097476-24-000007", EDGAR + "000009747624000007/"),
    "K22": ("Form 10-K FY2022 (period ended 2022-12-31), filed 2023-02-03", "0000097476-23-000007", EDGAR + "000009747623000007/"),
    "Q226": ("Form 10-Q Q2 FY2026 (period ended 2026-06-30), filed 2026-07-24", "0000097476-26-000152", EDGAR + "000009747626000152/"),
}
rows = []


def add(lid, statement, label, vals, src, rpage, note="", ptype="Fiscal year", printed=None):
    doc, acc, base = SRC[src]
    for p, v in vals.items():
        rows.append({"line_id": lid, "period": p, "period_type": ptype, "statement": statement, "line_item": label,
                     "value_usd_m": v, "classification": "Reported public data", "source_document": doc,
                     "accession_no": acc, "source_url": base + rpage, "label_as_printed": printed or label, "note": note})


IS, BS, CF, SEG, SEGI, INV = ("Income statement", "Balance sheet", "Cash flow statement", "Revenue by segment",
                              "Segment information", "Inventory detail")
Y = ["FY2025", "FY2024", "FY2023"]          # columns of the FY2025 10-K
Y2 = ["FY2022", "FY2021"]                   # extra years from the FY2023 10-K
z = lambda ys, vs: dict(zip(ys, vs))       # noqa: E731

# ---------------- income statement: FY2023-FY2025 from the FY2025 10-K (R3), FY2021-FY2022 from the FY2023 10-K (R3)
IS_LINES = [
    ("net_sales", "Revenue", [17682, 15641, 17519], [20028, 18344], ""),
    ("cost_of_sales", "Cost of revenue (COR)", [7599, 6547, 6500], [6257, 5968], ""),
    ("gross_profit", "Gross profit", [10083, 9094, 11019], [13771, 12376], ""),
    ("rd", "Research and development (R&D)", [2083, 1959, 1863], [1670, 1554], ""),
    ("sga", "Selling, general and administrative (SG&A)", [1860, 1794, 1825], [1704, 1666], ""),
    ("other_opex:restructuring", "Restructuring charges/other", [117, -124, 0], [257, 54],
     "2024 printed as a credit (124)."),
    ("operating_income", "Operating profit", [6023, 5465, 7331], [10140, 8960], ""),
    ("nonop:other_income", "Other income (expense), net (OI&E)", [230, 496, 440], [106, 143], ""),
    ("nonop:interest", "Interest and debt expense", [-543, -508, -353], [-214, -184],
     "Printed as a positive expense; stored negative (its effect on pre-tax income)."),
    ("pretax_income", "Income before income taxes", [5710, 5453, 7418], [10032, 8919], ""),
    ("income_tax", "Provision for income taxes", [709, 654, 908], [1283, 1150], ""),
    ("net_income", "Net income", [5001, 4799, 6510], [8749, 7769], ""),
]
for lid, lab, v3, v2, note in IS_LINES:
    add(lid, IS, lab, z(Y, v3), "K25", "R3.htm", note)
    add(lid, IS, lab, z(Y2, v2), "K23", "R3.htm", note)
add("other_opex:acquisition", IS, "Acquisition charges", z(Y2, [0, 142]), "K23", "R3.htm",
    "Printed only in the FY2023 10-K income statement (FY2021 142, FY2022 0); not presented for FY2023-FY2025.")

# ---------------- revenue by segment: FY2023-25 (FY2025 10-K R37), FY2021-22 (FY2023 10-K R34)
for lid, lab, v3, v2 in [("segment:analog", "Analog", [14006, 12161, 13040], [15359, 14050]),
                         ("segment:embedded_processing", "Embedded Processing", [2697, 2533, 3368], [3261, 3049]),
                         ("segment:other", "Other", [979, 947, 1111], [1408, 1245])]:
    add(lid, SEG, lab, z(Y, v3), "K25", "R37.htm", printed=f"{lab} - Total revenue")
    add(lid, SEG, lab, z(Y2, v2), "K23", "R34.htm", printed=f"{lab} - Total revenue")
for lid, lab, v3 in [("segment_info:analog_gross_profit", "Analog - gross profit", [8242, 7292, 8425]),
                     ("segment_info:embedded_gross_profit", "Embedded Processing - gross profit", [1226, 1218, 1875]),
                     ("segment_info:other_gross_profit", "Other - gross profit", [615, 584, 719])]:
    add(lid, SEGI, lab, z(Y, v3), "K25", "R37.htm", "Segment gross profit first shown in this form in the FY2025 10-K.")

# ---------------- balance sheet: FY2025/24 (FY2025 10-K R5), FY2023/22 (FY2023 10-K R6), FY2021 (FY2022 10-K R6)
BS_Y = [("K25", "R5.htm", ["FY2025", "FY2024"]), ("K23", "R6.htm", ["FY2023", "FY2022"]), ("K22", "R6.htm", ["FY2021"])]
BS_LINES = {  # line_id: (label, {period: value})
    "bs:cash": ("Cash and cash equivalents", [3225, 3200, 2964, 3050, 4631]),
    "bs:st_investments": ("Short-term investments", [1656, 4380, 5611, 6017, 5108]),
    "bs:receivables": ("Accounts receivable, net", [1963, 1719, 1787, 1895, 1701]),
    "bs:inventories": ("Inventories", [4804, 4527, 3999, 2757, 1910]),
    "bs:other_current_assets": ("Prepaid expenses and other current assets", [2102, 1200, 761, 302, 335]),
    "bs:total_current_assets": ("Total current assets", [13750, 15026, 15122, 14021, 13685]),
    "bs:ppe": ("Property, plant and equipment (net)", [12320, 11347, 9999, 6876, 5141]),
    "bs:total_assets": ("Total assets", [34585, 35509, 32348, 27207, 24676]),
    "bs:current_debt": ("Current portion of long-term debt", [500, 750, 599, 500, 500]),
    "bs:payables": ("Accounts payable", [756, 820, 802, 851, 571]),
    "bs:cl_other:accrued_compensation": ("Accrued compensation", [829, 839, 836, 799, 775]),
    "bs:cl_other:income_taxes_payable": ("Income taxes payable", [67, 159, 172, 189, 121]),
    "bs:accrued": ("Accrued expenses and other liabilities", [1007, 1075, 911, 646, 602]),
    "bs:total_current_liabilities": ("Total current liabilities", [3159, 3643, 3320, 2985, 2569]),
    "bs:long_term_debt": ("Long-term debt", [13548, 12846, 10624, 8235, 7241]),
    "bs:equity": ("Total stockholders' equity", [16273, 16903, 16897, 14577, 13333]),
    "inv:raw_materials": ("Raw materials", [465, 395, 420, 353, 245]),
    "inv:work_in_process": ("Work in process", [2372, 2214, 2109, 1546, 1067]),
    "inv:finished_goods": ("Finished goods", [1967, 1918, 1470, 858, 598]),
}
ALLY = ["FY2025", "FY2024", "FY2023", "FY2022", "FY2021"]
for lid, (lab, vals) in BS_LINES.items():
    v = z(ALLY, vals)
    for src, rp, ys in BS_Y:
        add(lid, INV if lid.startswith("inv:") else BS, lab, {y: v[y] for y in ys}, src, rp,
            "Printed on the face of the balance sheet." if lid.startswith("inv:") else "")

# ---------------- cash flow: FY2023-25 (FY2025 10-K R7), FY2021-22 (FY2023 10-K R8)
CF_LINES = [
    ("cf:net_income", "Net income", [5001, 4799, 6510], [8749, 7769], ""),
    ("cf:da", "Depreciation", [1918, 1508, 1175], [925, 755],
     "TI prints depreciation, amortization of capitalized software and (FY2021) acquisition amortization separately; "
     "cf:da = depreciation only. The amortization lines fall into other non-cash items."),
    ("cf:sbc", "Stock compensation", [419, 387, 362], [289, 230], ""),
    ("wc:receivables", "Accounts receivable", [-244, 68, 108], [-194, -287], ""),
    ("wc:inventories", "Inventories", [-277, -528, -1242], [-847, 45], ""),
    ("wc:prepaid", "Prepaid expenses and other current assets", [10, 7, 46], [6, 57], ""),
    ("wc:payables_accrued", "Accounts payable and accrued expenses", [77, 125, -33], [106, 33], ""),
    ("wc:accrued_compensation", "Accrued compensation", [-28, -12, 29], [22, 7], ""),
    ("wc:income_tax", "Income taxes payable", [191, 597, -7], [94, -20], ""),
    ("cf:cfo", "Cash flows from operating activities", [7153, 6318, 6420], [8720, 8756], ""),
    ("cf:capex", "Capital expenditures", [-4550, -4820, -5071], [-2797, -2462],
     "Gross capital expenditures. FY2025 also shows 335 of CHIPS Act incentive proceeds as a separate investing line (not netted)."),
    ("cf:cfi", "Cash flows from investing activities", [-1439, -3202, -4362], [-3583, -4095], ""),
    ("cf:cff", "Cash flows from financing activities", [-5689, -2880, -2144], [-6718, -3137], ""),
    ("cf:dividends", "Dividends paid", [-4999, -4795, -4557], [-4297, -3886], ""),
    ("cf:cash_end", "Cash and cash equivalents at end of period", [3225, 3200, 2964], [3050, 4631], ""),
]
for lid, lab, v3, v2, note in CF_LINES:
    add(lid, CF, lab, z(Y, v3), "K25", "R7.htm", note)
    add(lid, CF, lab, z(Y2, v2), "K23", "R8.htm", note)

# ---------------- quarterly context: Q2 FY2026 10-Q (income statement R2, segments R27); three months ended June 30
Q = ["Q2 FY2026", "Q2 FY2025"]
for lid, st, lab, vals, rp in [
        ("net_sales", IS, "Revenue", [5463, 4448], "R2.htm"),
        ("cost_of_sales", IS, "Cost of revenue (COR)", [2111, 1873], "R2.htm"),
        ("gross_profit", IS, "Gross profit", [3352, 2575], "R2.htm"),
        ("operating_income", IS, "Operating profit", [2310, 1563], "R2.htm"),
        ("net_income", IS, "Net income", [1980, 1295], "R2.htm"),
        ("segment:analog", SEG, "Analog", [4365, 3452], "R27.htm"),
        ("segment:embedded_processing", SEG, "Embedded Processing", [788, 679], "R27.htm"),
        ("segment:other", SEG, "Other", [310, 317], "R27.htm")]:
    add(lid, st, lab, z(Q, vals), "Q226", rp, "Three months ended June 30.", ptype="Fiscal quarter")


def main():
    with OUT.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    print(f"Wrote {len(rows)} reported values -> {OUT.name}")


if __name__ == "__main__":
    main()
