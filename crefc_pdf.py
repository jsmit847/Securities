"""
Parse the Computershare remittance statement (the DDST .pdf) into clean tables,
so the statement no longer has to be passed through untouched.

Uses pdfplumber to pull the high-value sections (Certificate Distribution,
Delinquency, Specially Serviced, Historical, etc.), joins the multi-line headers,
drops empty columns and the footnote/legend blocks, and converts formatted
numbers ("148,691,000.00", "(413.54)") into real numbers.
"""

from __future__ import annotations

import io
import re

import pandas as pd
import pdfplumber

# Section title (as printed on the statement) -> short sheet name (<=31 chars)
SECTIONS = {
    "Certificate Distribution Detail": "Stmt-Cert Distribution",
    "Certificate Interest Reconciliation Detail": "Stmt-Cert Interest Recon",
    "Delinquency Loan Detail": "Stmt-Delinquency",
    "Specially Serviced Loan Detail": "Stmt-Specially Serviced",
    "Historical Detail": "Stmt-Historical",
    "Principal Prepayment Detail": "Stmt-Prepayments",
    "Modified Loan Detail": "Stmt-Modified Loans",
    "Historical Liquidated Loan Detail": "Stmt-Liquidated Loans",
}

_NUM_RE = re.compile(r"^\(?-?[\d,]+(\.\d+)?\)?%?$")
_LEGEND_STARTS = ("1 ", "2 ", "* ", "*D", "(1)", "(2)", "(3)", "(4)", "NR ", "NR-",
                  "N/A -", "X -", "Note:", "HC -", "A - ", "B - ", "0 - ")


def _clean_cell(v):
    if v is None:
        return None
    s = str(v).replace("\n", " ").strip()
    s = re.sub(r"\s+", " ", s)
    return s or None


def _to_number(s):
    """Turn a displayed statement value into a number where sensible."""
    if s is None:
        return None
    t = s.strip()
    if not _NUM_RE.match(t):
        return s
    neg = t.startswith("(") and t.endswith(")")
    t = t.strip("()").rstrip("%").replace(",", "")
    if t in ("", "-"):
        return s
    try:
        val = float(t)
    except ValueError:
        return s
    if neg:
        val = -val
    return int(val) if val.is_integer() else val


def _looks_like_legend(row) -> bool:
    first = (row[0] or "") if row else ""
    return any(first.startswith(p) for p in _LEGEND_STARTS)


def _clean_table(rows):
    """rows: raw pdfplumber table (list of lists). Return a tidy DataFrame or None."""
    rows = [[_clean_cell(c) for c in r] for r in rows]
    # Drop the title row if the first row is a single populated cell (section name).
    if rows and sum(c is not None for c in rows[0]) <= 1:
        rows = rows[1:]
    # Find the header: first row with >=3 populated cells.
    hidx = next((i for i, r in enumerate(rows) if sum(c is not None for c in r) >= 3), None)
    if hidx is None:
        return None
    header = rows[hidx]
    body = rows[hidx + 1:]
    # Stop at the legend / footnote block.
    out = []
    for r in body:
        if _looks_like_legend(r) or all(c is None for c in r):
            if _looks_like_legend(r):
                break
            continue
        out.append(r)
    if not out:
        return None
    # Drop columns that are empty across header+body.
    ncol = max(len(header), max(len(r) for r in out))
    keep = [j for j in range(ncol)
            if (j < len(header) and header[j]) or any(j < len(r) and r[j] for r in out)]
    header = [(header[j] if j < len(header) and header[j] else f"Col{j+1}") for j in keep]
    # De-duplicate header labels.
    seen = {}
    for i, h in enumerate(header):
        if h in seen:
            seen[h] += 1
            header[i] = f"{h} ({seen[h]})"
        else:
            seen[h] = 0
    data = [[_to_number(r[j]) if j < len(r) else None for j in keep] for r in out]
    return pd.DataFrame(data, columns=header)


def parse_statement(pdf_bytes: bytes) -> dict[str, pd.DataFrame]:
    """Return {sheet_name: DataFrame} for the statement sections we recognise."""
    out: dict[str, pd.DataFrame] = {}
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            for tbl in page.extract_tables():
                if not tbl or not tbl[0]:
                    continue
                title = _clean_cell(tbl[0][0]) or ""
                match = next((full for full in SECTIONS if title.startswith(full)), None)
                if not match:
                    continue
                sheet = SECTIONS[match]
                if sheet in out:  # multi-page section (e.g. SS Part 1/2) -> keep the first
                    continue
                df = _clean_table(tbl)
                if df is not None and len(df):
                    out[sheet] = df
    return out


def statement_meta(pdf_bytes: bytes) -> dict:
    """Light metadata (distribution date, trust name) from page 1."""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        text = pdf.pages[0].extract_text() or ""
    dist = re.search(r"Distribution Date:\s*([\d/]+)", text)
    return {"distribution_date": dist.group(1) if dist else None}
