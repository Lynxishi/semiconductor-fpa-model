"""
Shared engine context: which company is being built, its config, and every input/output path.

The company is chosen by `--company TICKER` on the command line or the FPA_COMPANY environment variable
(default MCHP). Inputs live in companies/<TICKER>/, outputs are written to outputs/<TICKER>/.
"""
import os
import re
import sys
import tomllib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DRIVER_KEYS = ["dem_g", "asp_chg", "yld_adj", "util", "mat_inf", "rd_g", "sga_g", "capex_pct", "tax"]
PRODUCT_FIELDS = ["key", "id", "name", "short", "segment", "asp0", "yld0", "vc0", "fix0", "cap0", "gdiff", "inv_m",
                  "a_dem", "a_asp", "a_yld", "a_util", "a_vc", "a_fix"]


def ticker():
    for i, a in enumerate(sys.argv):
        if a == "--company" and i + 1 < len(sys.argv):
            return sys.argv[i + 1].upper()
        if a.startswith("--company="):
            return a.split("=", 1)[1].upper()
    return os.environ.get("FPA_COMPANY", "MCHP").upper()


class Engine:
    def __init__(self, tk=None):
        self.ticker = (tk or ticker()).upper()
        self.cdir = ROOT / "companies" / self.ticker
        cfg_path = self.cdir / "config.toml"
        if not cfg_path.exists():
            raise SystemExit(f"No config for {self.ticker}: expected {cfg_path.relative_to(ROOT)}")
        self.cfg = tomllib.loads(cfg_path.read_text(encoding="utf-8"))
        self.validate()
        c, d = self.cfg["company"], self.cfg["division"]
        self.company, self.division = c, d
        self.products = self.cfg["products"]
        for p in self.products:
            p.setdefault("abbr", p["id"])
            p.setdefault("tag", p["abbr"])
        self.pkeys = [p["key"] for p in self.products]
        self.years = c["history_years"]
        self.base_year = c["base_year"]
        self.fy = self.cfg["forecast"]["year"]
        self.disclaimer = (f"Simulated data created for portfolio demonstration. "
                           f"It does not represent nonpublic {c['name']} information.")
        slug = re.sub(r"[^A-Za-z0-9]+", "_", d["name"][:-len(" Division")] if d["name"].endswith(" Division") else d["name"]).strip("_")
        self.out = ROOT / "outputs" / self.ticker
        self.model = self.out / "model" / f"{self.ticker}_FPA_Model.xlsx"
        self.data_json = self.out / "dashboard" / "data.json"
        self.index_html = self.out / "dashboard" / "index.html"
        self.deck = self.out / "presentation" / f"{slug}_FY{self.fy[-2:]}_Executive_Review.pptx"
        self.deck_data = self.out / "presentation" / "deck_data.json"
        self.tests = self.out / "tests"
        self.screens = self.tests / "screens"
        self.docs = self.out / "docs"
        self.reported_csv = self.cdir / c.get("reported_csv", "reported_financials.csv")
        self.source_register = self.cdir / "source_register.csv"
        self.assumptions_log = self.cdir / "assumptions_log.csv"
        for p in (self.model.parent, self.data_json.parent, self.deck.parent, self.tests, self.screens, self.docs):
            p.mkdir(parents=True, exist_ok=True)

    # ---- config sanity: fail early with a message a person can act on ----
    def validate(self):
        c = self.cfg
        errs = []
        for sec in ("company", "forecast", "division", "products", "drivers", "seasonality", "pnl", "actuals", "capex"):
            if sec not in c:
                errs.append(f"missing section [{sec}]")
        if errs:
            raise SystemExit("Config errors: " + "; ".join(errs))
        if len(c["products"]) != 3:
            errs.append(f"exactly 3 [[products]] required, found {len(c['products'])}")
        for p in c["products"]:
            miss = [f for f in PRODUCT_FIELDS if f not in p]
            if miss:
                errs.append(f"product {p.get('key', '?')}: missing {miss}")
            for f in ("yld0", "a_yld", "a_util"):
                if f in p and not 0 < p[f] <= 1:
                    errs.append(f"product {p['key']}: {f} must be in (0, 1]")
        for k in DRIVER_KEYS:
            v = c["drivers"].get(k, {}).get("values")
            if not v or len(v) != 3:
                errs.append(f"drivers.{k}.values must have 3 numbers [Downside, Base, Upside]")
        s = c["seasonality"]["factors"]
        if len(s) != 12 or abs(sum(s) - 12) > 1e-9:
            errs.append(f"seasonality.factors must be 12 numbers summing to 12 (sum = {sum(s):.4f})")
        if c["forecast"]["actual_months"] != 5:
            errs.append("forecast.actual_months must be 5 (the Actuals/Variance layout is built for five actual months)")
        if len(c["actuals"]["noise"]) != c["forecast"]["actual_months"]:
            errs.append("actuals.noise must have one value per actual month")
        seg_sum = {}
        for p in c["products"]:
            sh = p.get("segment_share", 1.0)
            if not 0 < sh <= 1:
                errs.append(f"product {p.get('key')}: segment_share must be in (0, 1]")
            seg_sum[p.get("segment")] = seg_sum.get(p.get("segment"), 0) + sh
        for sg, tot in seg_sum.items():
            if tot > 1 + 1e-9:
                errs.append(f"segment_share values for {sg} add up to {tot:.3f} (> 100%)")
        if c["capex"]["target"] not in [p.get("key") for p in c["products"]]:
            errs.append("capex.target must be one of the product keys")
        if c["company"]["base_year"] != c["company"]["history_years"][-1]:
            errs.append("company.base_year must be the last of history_years")
        if errs:
            raise SystemExit("Config errors:\n  - " + "\n  - ".join(errs))

    def product(self, key):
        return next(p for p in self.products if p["key"] == key)

    @property
    def target(self):
        return self.product(self.cfg["capex"]["target"])

    def meta(self):
        """Names the dashboard and deck need (exported into data.json)."""
        c, d, cx = self.company, self.division, self.cfg["capex"]
        return {"ticker": self.ticker, "company": c["name"], "legal_name": c["legal_name"], "exchange": c["exchange"],
                "division": d["name"], "division_short": d["short"], "forecast_year": self.fy, "base_year": self.base_year,
                "filings": c.get("filings_summary", ""), "capex_project": cx["project"],
                "capex_target": self.target["id"], "as_of": self.cfg["forecast"].get("as_of", ""),
                "actuals_label": self.cfg["forecast"]["actuals_label"],
                "products": [{"id": p["id"], "key": p["key"], "name": p["name"], "short": p["short"], "abbr": p["abbr"], "tag": p["tag"]}
                             for p in self.products]}


E = None


def get():
    global E
    if E is None:
        E = Engine()
    return E
