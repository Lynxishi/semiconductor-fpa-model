"""Shared helpers for the validation scripts: engine context, row lookup by label or by line_id."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import engine  # noqa: E402

EN = engine.get()
NY = len(EN.years)
C_LID = 6 + NY          # Hist_Data column holding each row's line_id
BASE_COL = 2 + NY       # Hist_Data column of the base (last reported) year
PRODUCTS = EN.pkeys
PCOL = {k: 3 + i for i, k in enumerate(PRODUCTS)}   # product columns C, D, E on Assumptions
FY, BY = EN.fy, EN.base_year


def find_row(ws, label, col=1, start=1):
    for r in range(start, ws.max_row + 1):
        if ws.cell(row=r, column=col).value == label:
            return r
    raise KeyError(label)


def lid_row(hd, lid, start=1):
    """First Hist_Data row carrying this line_id (annual block)."""
    return find_row(hd, lid, col=C_LID, start=start)


def base_mix(hd):
    """Base-year product mix straight from the reported segment rows on Hist_Data (independent of Hist_Analysis)."""
    ns = hd.cell(row=lid_row(hd, "net_sales"), column=BASE_COL).value
    raw = {p["key"]: hd.cell(row=lid_row(hd, p["segment"]), column=BASE_COL).value / ns * p.get("segment_share", 1.0)
           for p in EN.products}
    segs = {p["segment"] for p in EN.products}
    all_segs = {hd.cell(row=r, column=C_LID).value for r in range(1, hd.max_row + 1)
                if str(hd.cell(row=r, column=C_LID).value or "").startswith("segment:")}
    if any(p.get("segment_share", 1.0) != 1.0 for p in EN.products) or segs != all_segs:
        tot = sum(raw.values())
        return {k: v / tot for k, v in raw.items()}      # renormalized over the modeled products
    return raw
