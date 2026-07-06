"""
CREFC IRP v8.4 conversion engine.

Turns raw, headerless Trustee (Computershare) investor-reporting files into
labeled, cleaned tables ready for the monthly reporting template.

Supported file types (detected from the filename token or chosen manually):

    CBND  ->  Bond level file          (Certificate / Bond detail)
    CCOL  ->  Collateral summary file  (deal-level rollup)
    LPER  ->  Loan periodic update     (loan-level periodic)
    PROP  ->  Property file            (property-level)
    FINF  ->  Financial file           (operating-statement line items)

The transformation for every type is the same:
    1. Parse the comma-delimited text with a real CSV reader (PROP contains
       quoted fields with embedded commas, so a naive split breaks it).
    2. Prepend the standardized CREFC 8.4 header row for that file type.
    3. Coerce numeric-looking cells to int/float; blank/whitespace -> empty.
    4. Optionally normalize YYYYMMDD date fields to real dates (YYYY-MM-DD).
"""

from __future__ import annotations

import csv
import io
import re
import datetime as _dt
from dataclasses import dataclass, field

import pandas as pd

from crefc_headers import CREFC_HEADERS

# ----------------------------------------------------------------------------- 
# File-type metadata
# ----------------------------------------------------------------------------- 
FILE_TYPES = {
    "CBND": {"label": "Bond",     "sheet": "Bond",     "out": "Bond"},
    "CCOL": {"label": "CollSUM",  "sheet": "CollSUM",  "out": "CollSUM"},
    "LPER": {"label": "Periodic", "sheet": "Periodic", "out": "Periodic"},
    "PROP": {"label": "Property", "sheet": "Property", "out": "Property"},
    "FINF": {"label": "Financial","sheet": "Financial","out": "Financial"},
}

# FINF has no validated reference file in the sample set; flag it as best-effort.
BEST_EFFORT = {"FINF"}

_INT_RE = re.compile(r"^-?\d+$")
_FLOAT_RE = re.compile(r"^-?(\d+\.\d*|\.\d+|\d+)([eE][-+]?\d+)?$")


def detect_type(filename: str) -> str | None:
    """Detect the CREFC file type from a filename token (e.g. CVAF_20181_LPER.txt)."""
    stem = filename.upper()
    for key in CREFC_HEADERS:
        if key in stem:
            return key
    return None


def _coerce_scalar(s: str):
    """Coerce a raw CSV cell string to int / float / cleaned string / None."""
    if s is None:
        return None
    if s == "" or s.strip() == "":
        return None                      # blank or whitespace-only -> empty cell
    v = s.strip()
    if _INT_RE.match(v):
        # keep as int unless it has a leading zero (identifier like a zip/id)
        if len(v.lstrip("-")) > 1 and v.lstrip("-").startswith("0"):
            return v
        try:
            return int(v)
        except ValueError:
            return v
    if _FLOAT_RE.match(v):
        try:
            return float(v)
        except ValueError:
            return v
    return v


def _as_date(value):
    """Parse a YYYYMMDD integer/string into a date; return None if not a valid date."""
    if value is None:
        return None
    s = str(value).strip()
    if not s or s in {"0", "00000000"}:
        return None
    s = s.split(".")[0]                   # tolerate '20271209.0'
    if len(s) != 8 or not s.isdigit():
        return None
    try:
        return _dt.date(int(s[:4]), int(s[4:6]), int(s[6:8]))
    except ValueError:
        return None


def date_columns(file_type: str) -> list[str]:
    """CREFC date fields: any header whose name contains 'date' (case-insensitive)."""
    return [h for h in CREFC_HEADERS[file_type] if "date" in h.lower()]


def date_column_indices(file_type: str) -> list[int]:
    """Positional indices of date fields (robust to duplicate header labels)."""
    return [i for i, h in enumerate(CREFC_HEADERS[file_type]) if "date" in h.lower()]


@dataclass
class ConversionResult:
    file_type: str
    label: str
    df: pd.DataFrame
    n_rows: int
    n_cols: int
    expected_cols: int
    ok: bool
    messages: list = field(default_factory=list)
    best_effort: bool = False


def convert(raw_bytes: bytes, file_type: str, normalize_dates: bool = True) -> ConversionResult:
    """Convert one raw CREFC text file's bytes into a labeled, cleaned DataFrame."""
    if file_type not in CREFC_HEADERS:
        raise ValueError(f"Unknown CREFC file type: {file_type}")

    headers = CREFC_HEADERS[file_type]
    expected = len(headers)
    messages: list[str] = []

    text = raw_bytes.decode("utf-8-sig", errors="replace")
    rows = [r for r in csv.reader(io.StringIO(text)) if r and any(c.strip() for c in r)]

    # Normalize row width to the expected column count.
    fixed = []
    ragged = 0
    for r in rows:
        if len(r) != expected:
            ragged += 1
            r = (r + [""] * expected)[:expected]
        fixed.append(r)
    if ragged:
        messages.append(f"{ragged} row(s) had an unexpected field count and were padded/trimmed to {expected}.")

    df = pd.DataFrame(fixed, columns=headers)

    # Type coercion + optional date normalization, done positionally so that
    # duplicate header labels (e.g. "Not Used") don't collapse columns.
    date_idx = set(date_column_indices(file_type)) if normalize_dates else set()
    cols = []
    for j in range(len(headers)):
        raw_col = df.iloc[:, j]
        coerced = raw_col.map(_coerce_scalar)
        if j in date_idx:
            coerced = coerced.map(_as_date)
        cols.append(coerced)
    df = pd.concat(cols, axis=1)
    df.columns = headers

    result = ConversionResult(
        file_type=file_type,
        label=FILE_TYPES.get(file_type, {}).get("label", file_type),
        df=df,
        n_rows=len(df),
        n_cols=len(df.columns),
        expected_cols=expected,
        ok=(len(df.columns) == expected and ragged == 0),
        messages=messages,
        best_effort=file_type in BEST_EFFORT,
    )
    return result


def infer_period(df: pd.DataFrame) -> str | None:
    """Return YYYY_MM from a Distribution Date column, if present."""
    for col in df.columns:
        if col.strip().lower() == "distribution date":
            for v in df[col].dropna():
                if isinstance(v, _dt.date):
                    return f"{v.year}_{v.month:02d}"
                s = str(v).split(".")[0]
                if len(s) == 8 and s.isdigit():
                    return f"{s[:4]}_{s[4:6]}"
    return None
