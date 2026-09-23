"""
Build outputs/<TICKER>/dashboard/index.html: fill the company tokens ({{DIVISION}}, {{FY}}, ...) in
dashboard/template.html and inject data.json (and the latest browser-test result, if it was produced against this
exact export) so the page is fully self-contained (no external data fetch).

Run: python scripts/build_dashboard.py --company TICKER   (default MCHP)
"""
import datetime as dt
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import engine  # noqa: E402

ROOT = engine.ROOT


def tokens(EN):
    c, f, cx = EN.company, EN.cfg["forecast"], EN.cfg["capex"]
    start = dt.date.fromisoformat(f["start"])
    last = dt.date(start.year + (start.month + 3) // 12, (start.month + 3) % 12 + 1, 1)   # 5th actual month
    div = EN.division["name"]
    return {
        "CONSOLE": (div[:-len(" Division")] if div.endswith(" Division") else div) + " Console",
        "DIVISION": div, "COMPANY": c["name"], "COMPANY_SHORT": c.get("short_name", c["name"].split()[0]),
        "LEGAL": c["legal_name"], "FILINGS": c.get("filings_summary", "SEC filings"),
        "FY": EN.fy, "FYS": "FY" + EN.fy[-2:], "BY": EN.base_year, "BYS": "FY" + EN.base_year[-2:],
        "ACT_THROUGH": f"{last:%b %Y}", "ACT_YEAR": f"{last:%Y}" if last.year == start.year else f"{start:%Y}-{last:%y}",
        "ACT_RANGE": f"{start:%b}&ndash;{last:%b}", "ACT_RANGE_TXT": f"{start:%b}–{last:%b}",
        "ACT_LONG": (f"{start:%B}&ndash;{last:%B} {last:%Y}" if last.year == start.year
                     else f"{start:%B %Y}&ndash;{last:%B %Y}"),
        "CAPEX_PROJECT": cx["project"], "CAPEX_LABEL": cx.get("target_label", EN.target["id"]),
        "MODEL_FILE": EN.model.name,
        "LAST_MONTH": f"{dt.date(start.year + (start.month + 10) // 12, (start.month + 10) % 12 + 1, 1):%b-%y}",
        "COMPANY_POSS": c["name"] + ("'" if c["name"].endswith("s") else "'s"),
        "CAPEX_PROJECT_CAP": cx["project"][:1].upper() + cx["project"][1:],
    }


def main():
    EN = engine.get()
    tpl = (ROOT / "dashboard" / "template.html").read_text()
    for k, v in tokens(EN).items():
        tpl = tpl.replace("{{" + k + "}}", v)
    left = sorted(set(re.findall(r"\{\{[A-Z_]+\}\}", tpl)))
    if left:
        raise SystemExit(f"Unfilled template tokens: {left}")
    data = EN.data_json.read_text()
    bt_path = EN.tests / "phase7_dashboard_validation.json"
    bt = "null"
    if bt_path.exists():
        res = json.loads(bt_path.read_text())
        # only embed a result that was produced against this exact data export
        if res.get("data_sha256") == json.loads(data)["source_sha256"]:
            bt = json.dumps({k: res[k] for k in ("compared", "failed", "scenarios", "run_utc")})
    safe = lambda s: s.replace("</", "<\\/")  # noqa: E731
    html = tpl.replace("/*__DATA__*/null", safe(json.dumps(json.loads(data), separators=(",", ":")))) \
              .replace("/*__BROWSER_TEST__*/null", safe(bt))
    EN.index_html.write_text(html)
    print(f"Wrote {EN.index_html.relative_to(ROOT)} ({len(html):,} bytes); browser test embedded: {bt != 'null'}")


if __name__ == "__main__":
    main()
