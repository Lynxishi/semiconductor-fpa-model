"""
Regression test for the engine refactor: the Microchip outputs must be unchanged.

Compares outputs/MCHP/dashboard/data.json with tests/baseline_mchp/data.json (saved before the refactor), and the
deck text with tests/baseline_mchp/deck_text.md when both exist. Every number must match to 1e-9 (relative); every
string must match exactly, except the documented, intentional changes listed in RENAMED / IGNORED below.

Run: python tests/regress_baseline.py            (only meaningful for MCHP)
Exit code 1 on any unexplained difference.
"""
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT / "tests" / "baseline_mchp"
NEW = ROOT / "outputs" / "MCHP"
IGNORED = {"generated_utc", "source_workbook", "source_sha256", "company"}   # run metadata, new engine meta block
RENAMED = {"mcu_sales": "mcu_segment_sales", "analog_sales": "ana_segment_sales", "other_sales": "fpga_segment_sales"}
# Intentional, explained differences in the deck text: (old fragment, new fragment)
TEXT_OK = [
    ("147/147 checks pass", "149/149 checks pass"),            # 2 new generic Phase-1 checks (segment totals, inventory detail)
    ("147 Excel checks", "149 Excel checks"),
    ("(utilization 65% → 80%)", "(utilization 67% → 80%)"),   # was typed by hand; now computed from Assumptions (units-weighted)
]


def walk(a, b, path, out):
    if isinstance(a, dict):
        if not isinstance(b, dict):
            out.append((path, "type", a, b))
            return
        for k in a:
            if path == "" and k in IGNORED:
                continue
            k2 = RENAMED.get(k, k)
            if k2 not in b:
                out.append((f"{path}/{k}", "missing in new", a[k], None))
                continue
            walk(a[k], b[k2], f"{path}/{k}", out)
        for k in b:
            if path == "" and k in IGNORED:
                continue
            if k not in a and k not in RENAMED.values():
                out.append((f"{path}/{k}", "new field", None, b[k]))
    elif isinstance(a, list):
        if not isinstance(b, list) or len(a) != len(b):
            out.append((path, "length", len(a), len(b) if isinstance(b, list) else b))
            return
        for i, (x, y) in enumerate(zip(a, b)):
            walk(x, y, f"{path}[{i}]", out)
    elif isinstance(a, (int, float)) and not isinstance(a, bool) and isinstance(b, (int, float)):
        if abs(a - b) > 1e-9 * max(1.0, abs(a)):
            out.append((path, "number", a, b))
    elif a != b:
        out.append((path, "value", a, b))


def main():
    diffs, n = [], 0
    a = json.loads((BASE / "data.json").read_text(encoding="utf-8"))
    b = json.loads((NEW / "dashboard" / "data.json").read_text(encoding="utf-8"))
    walk(a, b, "", diffs)
    n += 1
    deck_old, deck_new = BASE / "deck_text.md", NEW / "presentation" / "deck_text.md"
    if deck_new.exists():
        n += 1
        ol, nl = deck_old.read_text(encoding="utf-8").splitlines(), deck_new.read_text(encoding="utf-8").splitlines()
        if len(ol) != len(nl):
            diffs.append(("deck_text.md", "line count", len(ol), len(nl)))
        for i, (x, y) in enumerate(zip(ol, nl)):
            if x != y and not any(x.replace(o, n_) == y for o, n_ in TEXT_OK):
                diffs.append((f"deck_text.md:{i + 1}", "text", x, y))
    for d in diffs:
        print("DIFF", d)
    print(f"Regression vs baseline: {n} files compared, {len(diffs)} differences")
    sys.exit(1 if diffs else 0)


if __name__ == "__main__":
    main()
