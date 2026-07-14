"""
crefc_convert.py  —  CREFC IRP v8.4 raw-file conversion engine (optimized).

Drop-in replacement for the module app.py imports:
    from crefc_convert import convert, detect_type, infer_period, FILE_TYPES, CREFC_HEADERS

What changed vs. a naive "Text-to-Columns" conversion
-----------------------------------------------------
1. Real CSV parsing (csv module), so quoted fields containing commas
   (e.g. property addresses "5177 Salmon Drive SE, Unit B") stay in one
   column instead of shifting every column to their right.
2. The 8.4 header workbook is read ONCE and cached (functools.lru_cache),
   instead of being re-opened for every file and every deal.
3. Header layout is 100% data-driven from 8_4_CREFC_Header_File.xlsx —
   add/rename a column there and conversion follows, no code change.
4. EVERY date column (any header containing "date", plus the FINF
   "YYYYMMDD" columns) is normalized from YYYYMMDD -> real Excel date,
   not just Maturity Date. Blanks / placeholders (' ', '', '--', 'N/A',
   "N'A", '0') become empty cells.
5. Numeric columns are coerced to real numbers (vectorized) so downstream
   formulas never see text-that-looks-like-a-number. Identifier-style
   columns (CUSIP, Zip) are preserved as text to keep leading zeros.
"""

from __future__ import annotations

import csv
import io
import os
import re
import datetime as _dt
from dataclasses import dataclass, field
from functools import lru_cache
from typing import Optional

import pandas as pd
import openpyxl


# --------------------------------------------------------------------------- #
# Where the standardized 8.4 header workbook lives. Override with the
# CREFC_HEADER_FILE env var, or drop the file next to this module.
# --------------------------------------------------------------------------- #
HEADER_WORKBOOK = os.environ.get("CREFC_HEADER_FILE") or os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "8_4_CREFC_Header_File.xlsx"
)

# type-code -> display / output metadata. `sheet` matches the header-workbook
# tab name AND the on-screen tab; `out` is used for output filenames.
FILE_TYPES: dict[str, dict] = {
    "CBND": {"label": "Bond / Certificate", "sheet": "Bond",      "out": "Bond"},
    "CCOL": {"label": "Collateral Summary", "sheet": "Coll Summ", "out": "CollSUM"},
    "LPER": {"label": "Loan Periodic",      "sheet": "Periodic",  "out": "Periodic"},
    "PROP": {"label": "Property",           "sheet": "Property",  "out": "Property"},
    "FINF": {"label": "Financial",          "sheet": "Financial", "out": "Financial"},
    "TLOAN": {"label": "Total Loan",        "sheet": "Total Loan", "out": "TotalLoan"},
    "CLTL":  {"label": "Total Loan",        "sheet": "Total Loan", "out": "TotalLoan"},
}

# filename suffix -> type code (case-insensitive). Aliases welcome.
_SUFFIX_ALIASES = {
    "CBND": "CBND", "BOND": "CBND",
    "CCOL": "CCOL", "COLLSUM": "CCOL", "COLL": "CCOL",
    "LPER": "LPER", "PERIODIC": "LPER",
    "PROP": "PROP", "PROPERTY": "PROP",
    "FINF": "FINF", "FIN": "FINF", "FINANCIAL": "FINF",
    "TLOAN": "TLOAN", "CLTL": "CLTL", "TOTALLOAN": "TLOAN",
}

# tokens that mean "this cell is blank"
_BLANKS = {"", " ", "--", "-", "n/a", "na", "n'a", "nan", "none", "null"}

# columns we never want auto-typed to a number (leading zeros / codes)
_KEEP_TEXT = ("cusip", "zip")


# --------------------------------------------------------------------------- #
# Header loading (cached)
# --------------------------------------------------------------------------- #
@lru_cache(maxsize=1)
def _load_all_headers(path: str) -> dict[str, list[str]]:
    """Read row 1 of every tab in the header workbook, once. sheet -> headers."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    headers: dict[str, list[str]] = {}
    for ws in wb.worksheets:
        first = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
        headers[ws.title] = [("" if v is None else str(v)) for v in first]
    wb.close()
    return headers


def _headers_for(code: str) -> list[str]:
    sheet = FILE_TYPES[code]["sheet"]
    all_h = _load_all_headers(HEADER_WORKBOOK)
    if sheet not in all_h:
        raise KeyError(f"Header sheet {sheet!r} not found in {HEADER_WORKBOOK}")
    return all_h[sheet]


# CREFC_HEADERS: {type_code: [header, header, ...]} — exported for compatibility.
class _HeaderView(dict):
    """Lazy dict so importing the module never forces a workbook read."""
    def __missing__(self, code):
        self[code] = _headers_for(code)
        return self[code]


CREFC_HEADERS = _HeaderView()


# --------------------------------------------------------------------------- #
# Type detection & period inference
# --------------------------------------------------------------------------- #
def detect_type(name: str) -> Optional[str]:
    """Return a FILE_TYPES code from a filename, or None if unrecognized."""
    base = os.path.basename(name)
    stem = os.path.splitext(base)[0]
    token = re.split(r"[_\-\s]", stem)[-1].upper()
    return _SUFFIX_ALIASES.get(token)


def _yyyymmdd_to_date(s: str) -> Optional[_dt.datetime]:
    if re.fullmatch(r"\d{8}", s):
        try:
            return _dt.datetime(int(s[:4]), int(s[4:6]), int(s[6:8]))
        except ValueError:
            return None
    return None


def infer_period(df: pd.DataFrame) -> Optional[str]:
    """Best-effort 'YYYY_MM' reporting period from the Distribution Date col."""
    for col in df.columns:
        if str(col).strip().lower() == "distribution date":
            for v in df[col]:
                if isinstance(v, (_dt.date, _dt.datetime)):
                    return f"{v.year:04d}_{v.month:02d}"
                s = str(v).strip().replace(".0", "")
                d = _yyyymmdd_to_date(s)
                if d:
                    return f"{d.year:04d}_{d.month:02d}"
    return None


# --------------------------------------------------------------------------- #
# Cell / column helpers (vectorized)
# --------------------------------------------------------------------------- #
def _clean_str(v) -> str:
    return "" if v is None else str(v).strip()


def _is_date_header(h: str) -> bool:
    low = h.strip().lower()
    return ("date" in low) or (low == "yyyymmdd")


def _normalize_date_series(s: pd.Series) -> pd.Series:
    def conv(v):
        t = _clean_str(v).replace(".0", "")
        if t.lower() in _BLANKS or t == "0":
            return None
        d = _yyyymmdd_to_date(t)
        return d if d is not None else (v if t else None)
    return s.map(conv)


def _coerce_numeric_series(s: pd.Series) -> Optional[pd.Series]:
    """Return a numeric series if EVERY non-blank value parses; else None."""
    stripped = s.map(_clean_str)
    is_blank = stripped.str.lower().isin(_BLANKS)
    nonblank = stripped[~is_blank]
    if nonblank.empty:
        return None  # leave all-blank columns as-is
    parsed = pd.to_numeric(nonblank.str.replace(",", "", regex=False), errors="coerce")
    if not parsed.notna().all():
        return None  # not a clean numeric column -> keep text
    out = pd.to_numeric(stripped.str.replace(",", "", regex=False), errors="coerce")
    # keep integers integer-looking (dates already handled elsewhere)
    if (out.dropna() % 1 == 0).all():
        out = out.astype("Int64")
    return out


# --------------------------------------------------------------------------- #
# Result container + main entry point
# --------------------------------------------------------------------------- #
@dataclass
class ConversionResult:
    df: pd.DataFrame
    ok: bool = True
    best_effort: bool = False
    messages: list[str] = field(default_factory=list)

    @property
    def n_rows(self) -> int:
        return len(self.df)

    @property
    def n_cols(self) -> int:
        return self.df.shape[1]


def convert(data: bytes, code: str, normalize_dates: bool = True) -> ConversionResult:
    """Convert one raw CREFC text file (as bytes) into a labeled DataFrame."""
    if code not in FILE_TYPES:
        raise ValueError(f"Unknown file type code: {code!r}")

    headers = _headers_for(code)
    n = len(headers)
    messages: list[str] = []

    text = data.decode("utf-8-sig", "replace")
    rows = [r for r in csv.reader(io.StringIO(text)) if r and any(c.strip() for c in r)]

    # width guard — pad short rows, flag long ones (real feeds match exactly)
    fixed, ragged = [], 0
    for r in rows:
        if len(r) < n:
            r = r + [""] * (n - len(r))
        elif len(r) > n:
            ragged += 1
            r = r[:n]
        fixed.append(r)
    if ragged:
        messages.append(f"{ragged} row(s) had extra fields (truncated to {n} cols)")

    df = pd.DataFrame(fixed, columns=headers)

    # de-duplicate header labels for pandas while keeping order (Excel is fine
    # with duplicates; pandas needs uniqueness to address columns by position)
    df.columns = pd.RangeIndex(len(headers))  # address by position, safe & fast

    for i, h in enumerate(headers):
        if normalize_dates and _is_date_header(h):
            df[i] = _normalize_date_series(df[i])
            continue
        low = h.strip().lower()
        if any(k in low for k in _KEEP_TEXT):
            df[i] = df[i].map(lambda v: None if _clean_str(v).lower() in _BLANKS else _clean_str(v))
            continue
        num = _coerce_numeric_series(df[i])
        if num is not None:
            df[i] = num
        else:
            df[i] = df[i].map(lambda v: None if _clean_str(v).lower() in _BLANKS else _clean_str(v))

    df.columns = headers  # restore real labels for output

    ok = ragged == 0
    if not messages:
        messages.append("clean")
    return ConversionResult(df=df, ok=ok, best_effort=(code == "FINF"), messages=messages)
