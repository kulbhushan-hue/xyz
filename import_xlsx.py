"""
Import corrections from an Excel workbook back into data/*_compat.csv.

Typical use: you download a workbook from the app, correct the status / PTF /
usermod / EOS / latest-version cells in Excel, then run

    python tools/import_xlsx.py my_corrected.xlsx            # apply
    python tools/import_xlsx.py my_corrected.xlsx --dry-run  # only show what would change

The z/OS version is read from the header "R3.2 Compatible (Yes/No)" (3.2, 3.1 or 2.5).
Rows are matched on Vendor + Product + VRM. Only these CSV columns are touched:
status (zos32/zos31/zos25), the matching PTF column, eos_date, latest_version.
Cells that are empty are never used to overwrite existing data.
"""
from __future__ import annotations

import argparse
import glob
import os
import re
import sys
from datetime import date

import pandas as pd
from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from compatibility_db import VER, norm, norm_vrm  # noqa: E402

PLACEHOLDERS = {"", "none", "nan"}
NON_MATRIX_SHEETS = ("Sources", "Notes", "PTF Worklist")


def read_workbook(path: str, sheet: str | None = None):
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet] if sheet else next(s for s in wb if s.title not in NON_MATRIX_SHEETS)
    head = [str(c.value or "").strip() for c in ws[1]]
    version = None
    for h in head:
        m = re.match(r"R(\d\.\d)\s+Compatible", h)
        if m:
            version = m.group(1)
    if version not in VER:
        raise SystemExit(f"Cannot find a header like 'R3.2 Compatible (Yes/No)' on sheet '{ws.title}'.")
    try:
        idx = {"vendor": head.index("Vendor"), "product": head.index("Product"), "vrm": head.index("VRM"),
               "status": head.index(f"R{version} Compatible (Yes/No)"), "ptf": head.index("If Yes, PTF Required?"),
               "eos": head.index("EOS Date"), "latest": head.index("Latest Version")}
    except ValueError as e:
        raise SystemExit(f"Sheet '{ws.title}' is missing an expected column header: {e}")
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or not r[idx["vendor"]] or not r[idx["product"]]:
            continue
        rows.append({k: ("" if r[i] is None else str(r[i]).strip()) for k, i in idx.items()})
    return version, rows


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("xlsx")
    ap.add_argument("--data-dir", default=os.path.join(ROOT, "data"))
    ap.add_argument("--sheet", default=None, help="sheet name (default: first matrix sheet)")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args(argv)

    version, rows = read_workbook(a.xlsx, a.sheet)
    scol, pcol, _, _ = VER[version]
    book = {(r["vendor"].lower(), norm(r["product"]), norm_vrm(r["vrm"])): r for r in rows}
    files = sorted(glob.glob(os.path.join(a.data_dir, "*_compat.csv")))
    if not files:
        raise SystemExit(f"No *_compat.csv files in {a.data_dir}")

    changed_total, used = 0, set()
    for f in files:
        df = pd.read_csv(f, dtype=str, keep_default_na=False)
        dirty = False
        for i, r in df.iterrows():
            key = (r["vendor"].lower(), norm(r["product"]), norm_vrm(r["vrm"]))
            src = book.get(key)
            if not src:
                continue
            used.add(key)
            new = {scol: src["status"], pcol: src["ptf"], "eos_date": src["eos"], "latest_version": src["latest"]}
            diffs = {c: v for c, v in new.items() if v.lower() not in PLACEHOLDERS and v != r[c]}
            if not diffs:
                continue
            changed_total += 1
            print(f"{os.path.basename(f)}: {r['product']} {r['vrm']}")
            for c, v in diffs.items():
                print(f"    {c}: {r[c]!r}  ->  {v!r}")
                df.at[i, c] = v
            df.at[i, "last_checked"] = date.today().isoformat()
            dirty = True
        if dirty and not a.dry_run:
            df.to_csv(f, index=False, lineterminator="\r\n")
    for k in book:
        if k not in used:
            print("NOT FOUND in data (ignored):", " | ".join(k))
    print(f"\nz/OS {version}: {changed_total} row(s) {'would change' if a.dry_run else 'updated'}; "
          f"{len(used)} of {len(book)} workbook rows matched.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
