"""
Phase 0 validation - independent checks on a company's reported dataset (companies/<TICKER>/reported_financials.csv).
Generic: every check is written in terms of line_ids, so it works for any company in the standard format.

  1. Income statement adds up: GP = sales - COGS; opex = R&D + SG&A + other_opex:*; OI = GP - opex;
     pre-tax = OI + nonop:*; NI = pre-tax - tax
  2. Product/segment lines (segment:*) sum to net sales, in every period where they are present
  3. Cash flow ties to the balance sheet and income statement (ending cash, net income)
  4. Balance-sheet subtotals: total current assets / liabilities = their components
  5. Optional: [company.checks] segments_sum_to_sales and identities from the config; inventory detail (inv:*)
  6. Optional: cross-filing checks - the same value transcribed from a second, independent filing
     (companies/<TICKER>/cross_filing_checks.csv)
  7. Informational: inventory change on the balance sheet vs the cash-flow working-capital line
  8. Record integrity: no duplicates, no missing core values, every value sourced and classified, valid periods, units

Run: python tests/validate_reported_data.py --company TICKER   (default MCHP)
Writes outputs/<TICKER>/tests/phase0_validation_results.csv. Exit code 1 if any FAIL.
"""
import csv
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import engine  # noqa: E402

EN = engine.get()
OUT = EN.tests / "phase0_validation_results.csv"
TOL = 0.05  # USD m; filings print to 0.1, so a rounding tolerance of 0.05 is applied

rows = list(csv.DictReader(EN.reported_csv.open(encoding="utf-8")))
V = {(r["period"], r["line_id"]): float(r["value_usd_m"]) for r in rows if r["value_usd_m"] not in ("", None)}
LIDS = {r["line_id"] for r in rows}
YEARS = EN.years
QCFG = EN.company.get("quarter", {})
QUARTERS = [QCFG[k] for k in ("prior", "latest") if k in QCFG]
results = []


def check(name, expected, actual, tol=TOL, kind="FAIL"):
    diff = round(actual - expected, 4)
    status = "PASS" if abs(diff) <= tol else kind
    results.append({"check": name, "expected": round(expected, 4), "actual": round(actual, 4),
                    "difference": diff, "status": status})


def g(p, lid, default=None):
    if (p, lid) in V:
        return V[(p, lid)]
    if default is not None:
        return default
    raise KeyError(f"{lid} missing for {p}")


def has(p, *lids):
    return all((p, x) in V for x in lids)


def comps(prefix):
    return sorted(x for x in LIDS if x.startswith(prefix))


def S(p, prefix):
    return sum(V.get((p, x), 0.0) for x in comps(prefix))


SEGS = comps("segment:")
for y in YEARS:
    ns = g(y, "net_sales")
    # 1. income statement
    check(f"{y} gross profit = net sales - cost of sales", ns - g(y, "cost_of_sales"), g(y, "gross_profit"))
    if has(y, "total_opex"):
        check(f"{y} operating expenses = R&D + SG&A + other operating items",
              g(y, "rd") + g(y, "sga") + S(y, "other_opex:"), g(y, "total_opex"))
        check(f"{y} operating income = gross profit - opex", g(y, "gross_profit") - g(y, "total_opex"), g(y, "operating_income"))
    else:
        check(f"{y} operating income = GP - R&D - SG&A - other operating items",
              g(y, "gross_profit") - g(y, "rd") - g(y, "sga") - S(y, "other_opex:"), g(y, "operating_income"))
    check(f"{y} pre-tax income = operating income + non-operating items", g(y, "operating_income") + S(y, "nonop:"),
          g(y, "pretax_income"))
    check(f"{y} net income = pre-tax - tax", g(y, "pretax_income") - g(y, "income_tax"), g(y, "net_income"))
    # 3. cash flow ties
    if has(y, "bs:cash", "cf:cash_end"):
        check(f"{y} CF ending cash = BS cash", g(y, "bs:cash"), g(y, "cf:cash_end"))
    if has(y, "cf:net_income"):
        check(f"{y} CF net income = IS net income", g(y, "net_income"), g(y, "cf:net_income"))
    # 4. balance-sheet subtotals
    if has(y, "bs:total_current_liabilities"):
        cl = g(y, "bs:payables") + g(y, "bs:accrued") + g(y, "bs:current_debt", 0.0) + S(y, "bs:cl_other:")
        check(f"{y} total current liabilities = AP + accrued + current debt + other", cl, g(y, "bs:total_current_liabilities"))
    if has(y, "bs:total_current_assets"):
        ca = (g(y, "bs:cash") + g(y, "bs:st_investments", 0.0) + g(y, "bs:receivables") + g(y, "bs:inventories")
              + g(y, "bs:other_current_assets", 0.0) + S(y, "bs:ca_other:"))
        check(f"{y} total current assets = cash + ST investments + AR + inventory + other", ca, g(y, "bs:total_current_assets"))

# 2. segments / product lines sum to net sales wherever all are present
for p in YEARS + QUARTERS:
    if SEGS and has(p, "net_sales", *SEGS):
        check(f"{p} product lines sum to net sales", g(p, "net_sales"), sum(g(p, x) for x in SEGS))

# 5. optional config checks
ck = EN.company.get("checks", {})
seg_tot = ck.get("segments_sum_to_sales", [])
for p in YEARS:
    if seg_tot and has(p, *seg_tot):
        check(f"{p} reportable segments sum to net sales", g(p, "net_sales"), sum(g(p, x) for x in seg_tot))
    for idn in ck.get("identity", []):
        if has(p, idn["total"], *idn["parts"]):
            check(f"{p} {idn['name']}", g(p, idn["total"]), sum(g(p, x) for x in idn["parts"]))
    inv = comps("inv:")
    if inv and has(p, "bs:inventories", *inv):
        check(f"{p} inventory components sum to total", g(p, "bs:inventories"), sum(g(p, x) for x in inv))

# 6. cross-filing checks
xf = EN.cdir / "cross_filing_checks.csv"
if xf.exists():
    for r in csv.DictReader(xf.open(encoding="utf-8")):
        check(f"{r['period']} {r['line_id']}: stored value vs {r['second_source']}", float(r["expected_usd_m"]),
              g(r["period"], r["line_id"]))

# 7. informational
if "wc:inventories" in LIDS:
    for y0, y1 in zip(YEARS[:-1], YEARS[1:]):
        if has(y0, "bs:inventories") and has(y1, "bs:inventories", "wc:inventories"):
            check(f"{y1} inventory change: BS delta vs CF line (INFO)", g(y1, "bs:inventories") - g(y0, "bs:inventories"),
                  -g(y1, "wc:inventories"), tol=25.0, kind="INFO")

# 8. record integrity
keys = [(r["period"], r["line_id"]) for r in rows]
dups = [k for k, c in Counter(keys).items() if c > 1]
results.append({"check": "No duplicated (period, line_id) records", "expected": 0,
                "actual": len(dups), "difference": len(dups), "status": "PASS" if not dups else "FAIL"})
core = ["net_sales", "cost_of_sales", "gross_profit", "operating_income", "net_income"]
missing = [(y, c) for y in YEARS for c in core if (y, c) not in V]
results.append({"check": f"No missing core income-statement values {YEARS[0]}-{YEARS[-1]}", "expected": 0,
                "actual": len(missing), "difference": len(missing), "status": "PASS" if not missing else "FAIL"})
bad_src = [r for r in rows if not r["source_url"].startswith("https://") or not r["accession_no"]]
results.append({"check": "Every value has a source URL and accession number", "expected": 0,
                "actual": len(bad_src), "difference": len(bad_src), "status": "PASS" if not bad_src else "FAIL"})
bad_cls = [r for r in rows if r["classification"] != "Reported public data"]
results.append({"check": "Every value in this file classified as Reported public data", "expected": 0,
                "actual": len(bad_cls), "difference": len(bad_cls), "status": "PASS" if not bad_cls else "FAIL"})
valid_periods = set(YEARS) | set(QUARTERS)
bad_p = [r for r in rows if r["period"] not in valid_periods]
results.append({"check": "All periods are valid fiscal labels", "expected": 0, "actual": len(bad_p),
                "difference": len(bad_p), "status": "PASS" if not bad_p else "FAIL"})
# Unit sanity: USD-million statement values must be < 1,000,000 (catches values entered in $ or $000)
big = [r for r in rows if not r["line_id"].startswith("fact:") and abs(float(r["value_usd_m"] or 0)) >= 1_000_000]
results.append({"check": "Unit sanity: no value >= 1,000,000 (USD m)", "expected": 0, "actual": len(big),
                "difference": len(big), "status": "PASS" if not big else "FAIL"})
blank_id = [r for r in rows if not r["line_id"]]
results.append({"check": "Every record has a line_id", "expected": 0, "actual": len(blank_id),
                "difference": len(blank_id), "status": "PASS" if not blank_id else "FAIL"})

with OUT.open("w", newline="") as f:
    w = csv.DictWriter(f, fieldnames=["check", "expected", "actual", "difference", "status"])
    w.writeheader()
    w.writerows(results)

c = Counter(r["status"] for r in results)
print(f"{EN.ticker}: {len(results)} checks: {dict(c)}")
for r in results:
    if r["status"] != "PASS":
        print(f"  {r['status']}: {r['check']}  expected={r['expected']} actual={r['actual']} diff={r['difference']}")
sys.exit(1 if c.get("FAIL") else 0)
