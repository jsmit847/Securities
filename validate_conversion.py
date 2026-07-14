"""
validate_conversion.py — regression check for the CREFC converter.

Run it whenever you onboard a new securitization (or after changing the
header workbook). It converts the raw files and, if you have a known-good
"gold" workbook to compare against, diffs them cell-by-cell — tolerating the
two *intended* differences (YYYYMMDD ints promoted to real dates, and
placeholder text like 'N/A' blanked out).

Usage:
    # just convert + sanity-check shapes:
    python validate_conversion.py --raw path/to/CVAF_20182_*.txt

    # convert AND diff against gold workbooks:
    python validate_conversion.py \\
        --raw  CVAF_20182_LPER.txt --gold CVAF2018-02-Periodic.xlsx \\
        --raw  CVAF_20182_PROP.txt --gold CVAF2018-02-Property.xlsx
"""

from __future__ import annotations

import argparse
import datetime
import sys

import pandas as pd
import openpyxl

import crefc_convert as cc


def _gold_df(path: str) -> pd.DataFrame:
    wb = openpyxl.load_workbook(path, data_only=True)
    rows = list(wb.active.values)
    return pd.DataFrame(rows[1:])


def _canon(v):
    """Canonical comparable token: ('', ''), ('D', 'YYYYMMDD'), ('N', num), ('S', str)."""
    try:
        if v is None or pd.isna(v):
            return ("", "")
    except (TypeError, ValueError):
        pass
    if isinstance(v, (datetime.date, datetime.datetime)):
        return ("D", f"{v.year:04d}{v.month:02d}{v.day:02d}")
    s = str(v).strip()
    if s.lower() in ("nan", "none", "", "<na>", "nat", "n/a", "n'a", "--"):
        return ("", "")
    try:
        return ("N", f"{float(s.replace(',', '')):.4f}")
    except ValueError:
        return ("S", s)


def diff(mine: pd.DataFrame, gold: pd.DataFrame, name: str) -> int:
    mine = mine.reset_index(drop=True).copy(); mine.columns = range(mine.shape[1])
    gold = gold.reset_index(drop=True).copy(); gold.columns = range(gold.shape[1])
    rows, cols = min(len(mine), len(gold)), min(mine.shape[1], gold.shape[1])
    same = upgrades = real = 0
    examples = []
    for r in range(rows):
        for c in range(cols):
            ta, va = _canon(mine.iat[r, c])
            tb, vb = _canon(gold.iat[r, c])
            if (ta, va) == (tb, vb):
                same += 1
                continue
            if ta == "D" and tb == "N":              # date promoted from raw int
                try:
                    if int(float(vb)) == int(va):
                        upgrades += 1
                        continue
                except ValueError:
                    pass
            if {ta, tb} <= {"N", "S", ""}:            # numeric fuzz / blanked placeholder
                try:
                    if abs(float(va) - float(vb)) < 0.005:
                        same += 1
                        continue
                except ValueError:
                    pass
                if "" in (ta, tb):                    # one side blank placeholder
                    same += 1
                    continue
            real += 1
            if len(examples) < 10:
                examples.append((r, c, mine.iat[r, c], gold.iat[r, c]))
    flag = "OK" if real == 0 else "REVIEW"
    print(f"[{flag}] {name}: {rows}x{cols} cells | identical={same} "
          f"date-upgrades={upgrades} real-diffs={real}")
    for r, c, a, b in examples:
        print(f"        row {r} col {c}: converter={a!r}  gold={b!r}")
    return real


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--raw", action="append", default=[], help="raw CREFC .txt file")
    ap.add_argument("--gold", action="append", default=[], help="matching gold .xlsx (optional)")
    ap.add_argument("--no-dates", action="store_true", help="keep raw YYYYMMDD ints")
    args = ap.parse_args(argv)

    golds = args.gold + [None] * (len(args.raw) - len(args.gold))
    total_real = 0
    for raw, gold in zip(args.raw, golds):
        code = cc.detect_type(raw)
        if code is None:
            print(f"[SKIP] {raw}: unrecognized type")
            continue
        res = cc.convert(open(raw, "rb").read(), code, normalize_dates=not args.no_dates)
        print(f"------ {raw} -> {code} ({res.n_rows} rows x {res.n_cols} cols, "
              f"period={cc.infer_period(res.df)}) :: {', '.join(res.messages)}")
        if gold:
            total_real += diff(res.df, _gold_df(gold), code)

    if total_real:
        print(f"\n{total_real} cell(s) need review.")
        sys.exit(1)
    print("\nAll checks passed.")


if __name__ == "__main__":
    main()
