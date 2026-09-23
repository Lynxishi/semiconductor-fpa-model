"""
Export the dashboard data from the Excel model.

For each scenario (Downside, Base, Upside):
  1. copy the workbook, set Assumptions!C6, recalculate with LibreOffice
  2. read every table on the Dash_Export tab (by its 'TABLE: <id>' marker)
Then write outputs/<TICKER>/dashboard/data.json and VALIDATE the file against Excel:
  * JSON control totals = Excel control totals (per scenario)
  * sums rebuilt from the JSON arrays = Excel totals (monthly -> FY, products -> division)
  * scenario-independent tables (budget, actuals, variance, capex, history) identical in all three runs
  * Checks tab = ALL CHECKS PASS in all three runs, 0 formula errors
Writes outputs/<TICKER>/tests/phase6_export_validation.json. Exit code 1 on any failure.

Run: python scripts/export_dashboard_data.py --company TICKER   (default MCHP)
"""
import datetime as dt
import hashlib
import json
import shutil
import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
import engine  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
EN = engine.get()
SRC = EN.model
OUT = EN.data_json
RECALC = str(Path(__file__).resolve().parents[1] / "scripts" / "recalc.py")
TMP = EN.tests / "_export_runs"
SCENARIOS = ["Downside", "Base", "Upside"]
LIVE_TABLES = ["meta", "monthly", "product_monthly", "product_fy", "kpis", "control_totals"]
STATIC_TABLES = ["variance_pnl", "revenue_variance", "op_bridge", "variance_required", "variance_kpis",
                 "capex_summary", "capex_cashflows", "capex_inputs", "history", "scenario_drivers"]
TOL = 1e-9
FSTART = dt.date.fromisoformat(EN.cfg["forecast"]["start"])
FEND = dt.date(FSTART.year + (FSTART.month + 10) // 12, (FSTART.month + 10) % 12 + 1, 1)


def clean(v):
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    return v


def read_tables(ws):
    tables, r = {}, 1
    while r <= ws.max_row:
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and v.startswith("TABLE: "):
            tid = v[7:]
            hdr = r + 1
            cols = []
            c = 1
            while ws.cell(row=hdr, column=c).value not in (None, ""):
                cols.append(ws.cell(row=hdr, column=c).value)
                c += 1
            recs, rr = [], hdr + 1
            while ws.cell(row=rr, column=1).value not in (None, ""):
                recs.append({cols[j]: clean(ws.cell(row=rr, column=1 + j).value) for j in range(len(cols))})
                rr += 1
            tables[tid] = recs
            r = rr
        r += 1
    return tables


def run_scenario(name):
    f = TMP / f"export_{name}.xlsx"
    shutil.copy(SRC, f)
    wb = load_workbook(f)
    wb["Assumptions"]["C6"] = name
    wb.save(f)
    rc = json.loads(subprocess.run([sys.executable, RECALC, str(f), "180"], capture_output=True, text=True).stdout)
    v = load_workbook(f, data_only=True)
    t = read_tables(v["Dash_Export"])
    direct = {  # values read straight from the model tabs (not through Dash_Export) for the cross-check
        "revenue_fy": v["PnL"]["O" + str(find(v["PnL"], "Revenue"))].value,
        "operating_profit_fy": v["PnL"]["O" + str(find(v["PnL"], "Operating profit"))].value,
        "fcf_fy": v["FCF"]["O" + str(find(v["FCF"], "Free cash flow"))].value,
    }
    return t, rc, direct


def find(ws, label):
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=1).value == label:
            return r
    raise KeyError(label)


def main():
    TMP.mkdir(exist_ok=True)
    runs, results = {}, []
    for s in SCENARIOS:
        runs[s] = run_scenario(s)

    data = {
        "generated_utc": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
        "source_workbook": str(SRC.relative_to(ROOT)),
        "company": EN.meta(),
        "source_sha256": hashlib.sha256(SRC.read_bytes()).hexdigest(),
        "disclaimer": runs["Base"][0]["meta"][3]["value"],
        "scenarios": {s: {k: runs[s][0][k] for k in LIVE_TABLES} for s in SCENARIOS},
        "static": {k: runs["Base"][0][k] for k in STATIC_TABLES},
    }
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(data, indent=1, default=str))

    # ------------------- validation of the written file -------------------
    j = json.loads(OUT.read_text())

    def add(check, expected, actual, ok=None):
        if ok is None:
            ok = abs(expected - actual) <= 1e-6 * max(1.0, abs(expected))
        results.append({"check": check, "expected": expected, "actual": actual, "status": "PASS" if ok else "FAIL"})

    for s in SCENARIOS:
        t, rc, direct = runs[s]
        sj = j["scenarios"][s]
        ctl = {x["control_id"]: x["value"] for x in sj["control_totals"]}
        meta = {x["key"]: x["value"] for x in sj["meta"]}
        add(f"{s}: scenario label in export", 1, 1, meta["scenario"] == s)
        add(f"{s}: Checks tab status", 1, 1, meta["checks_status"] == "ALL CHECKS PASS")
        add(f"{s}: formula errors after recalculation", 0, rc.get("total_errors", -1))
        add(f"{s}: JSON revenue control total = PnL tab", direct["revenue_fy"], ctl["revenue_fy"])
        add(f"{s}: JSON operating-profit control total = PnL tab", direct["operating_profit_fy"], ctl["operating_profit_fy"])
        add(f"{s}: JSON FCF control total = FCF tab", direct["fcf_fy"], ctl["fcf_fy"])
        m = sj["monthly"]
        add(f"{s}: sum of 12 monthly revenue records = FY revenue", ctl["revenue_fy"], sum(x["revenue"] for x in m))
        add(f"{s}: sum of monthly operating profit = FY", ctl["operating_profit_fy"], sum(x["operating_profit"] for x in m))
        add(f"{s}: sum of monthly FCF = FY", ctl["fcf_fy"], sum(x["fcf"] for x in m))
        add(f"{s}: 12 monthly records, unique months", 12, len({x["month"] for x in m}), len(m) == 12)
        pm = sj["product_monthly"]
        add(f"{s}: 36 product-month records, unique keys", 36, len({(x['month'], x['product_id']) for x in pm}), len(pm) == 36)
        add(f"{s}: product-month revenue sums to FY revenue", ctl["revenue_fy"], sum(x["revenue"] for x in pm))
        add(f"{s}: product FY revenue sums to FY revenue", ctl["revenue_fy"], sum(x["revenue"] for x in sj["product_fy"]))
        add(f"{s}: actual months = first five months only", 5, sum(1 for x in m if x["actual_revenue"] is not None),
            [x["month"] for x in m if x["actual_revenue"] is not None] == [x["month"] for x in m[:5]])
        add(f"{s}: YTD actual revenue from monthly records = control total", ctl["ytd_actual_revenue"],
            sum(x["actual_revenue"] for x in m if x["actual_revenue"] is not None))
        add(f"{s}: budget FY from monthly records = control total", ctl["budget_revenue_fy"], sum(x["budget_revenue"] for x in m))
        months = [dt.date.fromisoformat(x["month"]) for x in m]
        add(f"{s}: months consecutive {FSTART:%b-%Y}..{FEND:%b-%Y}", 1, 1,
            months[0] == FSTART and months[-1] == FEND and
            all((b.year - a.year) * 12 + b.month - a.month == 1 for a, b in zip(months, months[1:])))
        bad = [x for x in pm if not (0 <= x["yield"] <= 1 and 0 <= x["utilization"] <= 1)]
        add(f"{s}: yields and utilization within 0-100%", 0, len(bad))
        # static tables must not depend on the scenario
        for k in STATIC_TABLES:
            add(f"{s}: '{k}' identical to Base run (scenario-independent)", 1, 1, t[k] == runs["Base"][0][k])
    br = j["static"]["op_bridge"]
    add("Bridge: start + deltas = end", br[-1]["value"], br[0]["value"] + sum(x["value"] for x in br[1:-1]))
    base_ctl = {x["control_id"]: x["value"] for x in j["scenarios"]["Base"]["control_totals"]}
    add("Bridge end = YTD actual operating profit control total", base_ctl["ytd_actual_op"], br[-1]["value"])
    cs = {x["scenario"]: x for x in j["static"]["capex_summary"]}
    add("Capex Base NPV = control total", base_ctl["capex_npv_base"], cs["Base"]["npv"])
    cf = [x for x in j["static"]["capex_cashflows"] if x["scenario"] == "Base"]
    rate = next(x["value"] for x in j["static"]["capex_inputs"] if x["key"] == "rate")
    add("Capex Base NPV rebuilt from exported cash flows", cs["Base"]["npv"], sum(x["fcf"] / (1 + rate) ** x["year"] for x in cf))

    shutil.rmtree(TMP, ignore_errors=True)
    n_fail = sum(r["status"] == "FAIL" for r in results)
    (EN.tests / "phase6_export_validation.json").write_text(json.dumps(results, indent=2, default=str))
    print(f"Wrote {OUT.relative_to(ROOT)} ({OUT.stat().st_size:,} bytes)")
    print(f"Export validation: {len(results) - n_fail} PASS, {n_fail} FAIL")
    for r in results:
        if r["status"] == "FAIL":
            print("  FAIL:", r)
    for s in SCENARIOS:
        k = {x["metric_id"]: x["value"] for x in j["scenarios"][s]["kpis"]}
        print(f"  {s:8} revenue {k['revenue']:9.1f}  op {k['operating_profit']:7.1f}  fcf {k['fcf']:7.1f}  "
              f"outlook op vs budget {k['outlook_op_vs_budget']:+6.1f}")
    sys.exit(1 if n_fail else 0)


if __name__ == "__main__":
    main()
