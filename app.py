"""
CREFC Investor-Reporting Converter
==================================

Drop in the raw, headerless files from the Trustee site (Computershare) and get
back clean, labeled CREFC IRP v8.4 tables ready for the monthly reporting
template - no more manual Text-to-Columns and date fixing.

Run locally:
    pip install -r requirements.txt
    streamlit run app.py
"""

from __future__ import annotations

import io
import zipfile
import datetime as _dt

import pandas as pd
import streamlit as st

from crefc_convert import convert, detect_type, infer_period, FILE_TYPES, CREFC_HEADERS
from crefc_io import single_workbook_bytes, combined_workbook_bytes

# ----------------------------------------------------------------------------- 
# Page + theme
# ----------------------------------------------------------------------------- 
st.set_page_config(
    page_title="CREFC Investor-Reporting Converter",
    page_icon="▤",
    layout="wide",
)

st.markdown(
    """
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;700&family=IBM+Plex+Mono:wght@400;500;600&family=Inter:wght@400;500;600&display=swap');

      :root{
        --ink:#14202B; --paper:#F5F6F4; --panel:#FFFFFF;
        --gold:#B8862F; --ok:#2E6E4E; --alert:#B03A2E;
        --line:#D7DBD5; --muted:#5C6A72;
      }
      html, body, [class*="css"]  { font-family:'Inter',system-ui,sans-serif; color:var(--ink); }
      .stApp { background:var(--paper); }
      .block-container{ padding-top:2.2rem; max-width:1200px; }

      /* Masthead */
      .mast{ border-top:3px solid var(--ink); border-bottom:1px solid var(--line);
             padding:14px 0 12px; margin-bottom:6px; }
      .mast .eyebrow{ font-family:'IBM Plex Mono',monospace; font-size:11px; letter-spacing:.28em;
             text-transform:uppercase; color:var(--gold); font-weight:600; }
      .mast h1{ font-family:'Space Grotesk',sans-serif; font-weight:700; font-size:30px;
             letter-spacing:-.01em; margin:.15rem 0 .1rem; }
      .mast .sub{ color:var(--muted); font-size:13.5px; }

      /* Blotter */
      .blotter{ background:var(--panel); border:1px solid var(--line); border-radius:2px;
                padding:6px 0; font-family:'IBM Plex Mono',monospace; }
      .blot-row{ display:flex; align-items:center; gap:12px; padding:9px 16px;
                 border-bottom:1px dotted var(--line); font-size:13px; }
      .blot-row:last-child{ border-bottom:none; }
      .blot-tkr{ color:var(--gold); font-weight:600; min-width:150px; }
      .blot-name{ flex:1; color:var(--ink); }
      .blot-dim{ color:var(--muted); min-width:120px; text-align:right; }
      .blot-badge{ min-width:92px; text-align:right; font-weight:600; }
      .ok{ color:var(--ok); } .warn{ color:var(--gold); } .err{ color:var(--alert); }

      .stat{ font-family:'IBM Plex Mono',monospace; }
      .stat b{ font-size:22px; } .stat span{ color:var(--muted); font-size:11px;
               letter-spacing:.14em; text-transform:uppercase; display:block; }

      h2, h3 { font-family:'Space Grotesk',sans-serif; letter-spacing:-.01em; }
      .stDownloadButton button, .stButton button{ border-radius:2px; border:1px solid var(--ink);
             background:var(--ink); color:#fff; font-weight:600; }
      .stDownloadButton button:hover, .stButton button:hover{ background:var(--gold); border-color:var(--gold); }
      section[data-testid="stSidebar"]{ background:var(--panel); border-right:1px solid var(--line); }
      .note{ background:#fff; border-left:3px solid var(--gold); padding:12px 16px;
             font-size:13px; color:var(--ink); border-radius:0 2px 2px 0; }
      code{ font-family:'IBM Plex Mono',monospace; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="mast">
      <div class="eyebrow">Investor Reporting · CREFC IRP v8.4</div>
      <h1>Raw → Converted File Converter</h1>
      <div class="sub">Drop the Trustee-site files (CBND · CCOL · LPER · PROP · FINF). Get labeled,
      date-clean tables for the monthly reporting template — the Text-to-Columns step, automated.</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------------- 
# Sidebar controls
# ----------------------------------------------------------------------------- 
with st.sidebar:
    st.markdown("### Settings")
    deal_label = st.text_input("Deal label (output prefix)", value="CVAF2018-01",
                               help="Used to name the downloads, e.g. CVAF2018-01-Periodic_2026_05.xlsx")
    date_mode = st.radio(
        "Date fields",
        ["Normalize to YYYY-MM-DD (recommended)", "Keep raw (YYYYMMDD)"],
        help="Trustee files store dates as YYYYMMDD integers. Normalizing writes real Excel dates.",
    )
    normalize_dates = date_mode.startswith("Normalize")
    st.divider()
    st.markdown(
        "<div style='font-family:IBM Plex Mono,monospace;font-size:11px;color:#5C6A72;line-height:1.7'>"
        "TYPE&nbsp;&nbsp;→ CONTENTS<br>"
        "CBND → Bond / certificate<br>"
        "CCOL → Collateral summary<br>"
        "LPER → Loan periodic<br>"
        "PROP → Property<br>"
        "FINF → Financial (best-effort)</div>",
        unsafe_allow_html=True,
    )

# ----------------------------------------------------------------------------- 
# Upload
# ----------------------------------------------------------------------------- 
uploaded = st.file_uploader(
    "Upload raw Trustee files",
    type=["txt", "csv", "xls", "xlsx", "pdf"],
    accept_multiple_files=True,
    help="File type is detected from the name (…_LPER.txt etc.). You can override anything below.",
)

if not uploaded:
    st.markdown(
        "<div class='note'>Waiting for files. Drop the raw <code>.txt</code> exports "
        "(e.g. <code>CVAF_20181_LPER.txt</code>). The <code>DDST.pdf</code> statement and the "
        "<code>RSRV.xls</code> reserve file are already human-readable and are passed through, not reparsed.</div>",
        unsafe_allow_html=True,
    )
    st.stop()

# ----------------------------------------------------------------------------- 
# Classify + let user override each file's type
# ----------------------------------------------------------------------------- 
TYPE_CHOICES = ["(skip)"] + list(FILE_TYPES.keys())
plan = []  # (uploaded_file, chosen_type, note)
with st.expander("Review detected file types", expanded=False):
    for uf in uploaded:
        low = uf.name.lower()
        if low.endswith(".pdf"):
            st.write(f"📄 `{uf.name}` — statement PDF, passed through (no conversion).")
            plan.append((uf, "PDF", "passthrough"))
            continue
        if low.endswith((".xls", ".xlsx")) and detect_type(uf.name) is None:
            st.write(f"📊 `{uf.name}` — already a workbook, passed through.")
            plan.append((uf, "XLS", "passthrough"))
            continue
        guess = detect_type(uf.name) or "(skip)"
        choice = st.selectbox(
            f"`{uf.name}`", TYPE_CHOICES,
            index=TYPE_CHOICES.index(guess) if guess in TYPE_CHOICES else 0,
            key=f"type_{uf.name}",
        )
        plan.append((uf, choice, "convert" if choice != "(skip)" else "skip"))

# ----------------------------------------------------------------------------- 
# Convert
# ----------------------------------------------------------------------------- 
results = {}       # sheet_name -> DataFrame
meta = []          # blotter rows
passthrough = []   # (name, bytes)
period_guess = None

for uf, choice, note in plan:
    data = uf.getvalue()
    if choice in ("PDF", "XLS"):
        passthrough.append((uf.name, data))
        meta.append((deal_label, uf.name, note, "—", "PASS", "passthrough"))
        continue
    if choice == "(skip)":
        meta.append((deal_label, uf.name, "skipped", "—", "SKIP", "skip"))
        continue
    try:
        res = convert(data, choice, normalize_dates=normalize_dates)
    except Exception as exc:  # noqa: BLE001
        meta.append((deal_label, uf.name, choice, "—", "ERR", str(exc)))
        continue
    period_guess = period_guess or infer_period(res.df)
    sheet = FILE_TYPES[choice]["sheet"]
    results[sheet] = res.df
    badge = "OK" if res.ok else "CHK"
    detail = "; ".join(res.messages) if res.messages else (
        "best-effort headers" if res.best_effort else "clean")
    meta.append((deal_label, uf.name, FILE_TYPES[choice]["label"],
                 f"{res.n_rows:,} × {res.n_cols}", badge, detail))

period = period_guess or _dt.date.today().strftime("%Y_%m")

# ----------------------------------------------------------------------------- 
# Summary stats + Blotter
# ----------------------------------------------------------------------------- 
c1, c2, c3, c4 = st.columns(4)
converted_n = len(results)
row_total = sum(len(df) for df in results.values())
c1.markdown(f"<div class='stat'><b>{converted_n}</b><span>Tables converted</span></div>", unsafe_allow_html=True)
c2.markdown(f"<div class='stat'><b>{row_total:,}</b><span>Data rows</span></div>", unsafe_allow_html=True)
c3.markdown(f"<div class='stat'><b>{period}</b><span>Reporting period</span></div>", unsafe_allow_html=True)
c4.markdown(f"<div class='stat'><b>v8.4</b><span>CREFC IRP</span></div>", unsafe_allow_html=True)

st.markdown("### Conversion blotter")
rows_html = []
for tkr, name, label, dim, badge, detail in meta:
    cls = {"OK": "ok", "PASS": "ok", "CHK": "warn", "SKIP": "warn", "ERR": "err"}.get(badge, "warn")
    rows_html.append(
        f"<div class='blot-row'><span class='blot-tkr'>{tkr}</span>"
        f"<span class='blot-name'>{name}<br>"
        f"<span style='color:#5C6A72;font-size:11px'>{label} · {detail}</span></span>"
        f"<span class='blot-dim'>{dim}</span>"
        f"<span class='blot-badge {cls}'>{badge}</span></div>"
    )
st.markdown(f"<div class='blotter'>{''.join(rows_html)}</div>", unsafe_allow_html=True)

if not results:
    st.stop()

# ----------------------------------------------------------------------------- 
# Downloads
# ----------------------------------------------------------------------------- 
st.markdown("### Download")
dl1, dl2 = st.columns(2)

combined_bytes = combined_workbook_bytes(results)
dl1.download_button(
    "Combined workbook (one tab per table)",
    data=combined_bytes,
    file_name=f"{deal_label}-CREFC-Converted_{period}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

zip_buf = io.BytesIO()
with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
    for sheet, df in results.items():
        out = FILE_TYPES[[k for k, v in FILE_TYPES.items() if v["sheet"] == sheet][0]]["out"]
        zf.writestr(f"{deal_label}-{out}_{period}.xlsx", single_workbook_bytes(df, sheet))
    for name, data in passthrough:
        zf.writestr(f"_passthrough/{name}", data)
dl2.download_button(
    "All files (.zip — individual workbooks)",
    data=zip_buf.getvalue(),
    file_name=f"{deal_label}-CREFC-Converted_{period}.zip",
    mime="application/zip",
    use_container_width=True,
)

# ----------------------------------------------------------------------------- 
# Per-table preview + individual download
# ----------------------------------------------------------------------------- 
st.markdown("### Tables")
for sheet, df in results.items():
    out = FILE_TYPES[[k for k, v in FILE_TYPES.items() if v["sheet"] == sheet][0]]["out"]
    with st.expander(f"{sheet}  ·  {len(df):,} rows × {len(df.columns)} cols", expanded=False):
        st.dataframe(df.head(200), use_container_width=True, height=340)
        st.caption(f"Showing up to 200 rows of {len(df):,}.")
        st.download_button(
            f"Download {sheet}",
            data=single_workbook_bytes(df, sheet),
            file_name=f"{deal_label}-{out}_{period}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_{sheet}",
        )

# ----------------------------------------------------------------------------- 
# Derived analytics: Maturity ladder + Watchlist extract (from Periodic)
# ----------------------------------------------------------------------------- 
def _is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    return str(v).strip() in ("", "nan", "None")


def _num(series):
    return pd.to_numeric(series, errors="coerce")


periodic = results.get("Periodic")
if periodic is not None:
    st.markdown("### Surveillance extracts")
    st.markdown(
        "<div class='note'>Starter views pulled from the Loan Periodic file to feed the "
        "<b>WL Maturity Report</b> and <b>CMBS Watchlist</b> tabs. Scoped to <b>active loans</b> "
        "(ending balance &gt; 0). Wire them into your reporting template as needed.</div>",
        unsafe_allow_html=True,
    )

    bal_col = "Current Ending Scheduled  Balance"
    active = periodic[_num(periodic[bal_col]) > 0].copy() if bal_col in periodic.columns else periodic.copy()

    tab_mat, tab_wl = st.tabs([f"Maturity ladder ({len(active)})", "Watchlist extract"])

    with tab_mat:
        keep = [c for c in [
            "Prospectus Loan ID", "Loan ID", "Maturity Date", bal_col, "Current Note Rate",
            "Most Recent DSCR (NOI)", "Payment Status of Loan   (fka Status of Loan)",
            "Number of Properties",
        ] if c in active.columns]
        mat = active[keep].copy()
        mcol = "Maturity Date"
        if mcol in mat.columns:
            today = _dt.date.today()
            mat["Months to Maturity"] = mat[mcol].map(
                lambda d: (d.year - today.year) * 12 + (d.month - today.month)
                if isinstance(d, _dt.date) else None)
            mat = mat.sort_values(mcol, na_position="last")
        st.dataframe(mat, use_container_width=True, height=360)
        st.download_button(
            "Download Maturity ladder",
            data=single_workbook_bytes(mat, "Maturity"),
            file_name=f"{deal_label}-WL-Maturity_{period}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_mat",
        )

    with tab_wl:
        status_col = "Payment Status of Loan   (fka Status of Loan)"
        wl_col = "Date Added to Servicer Watchlist"
        ss_col = "Reason for SS Transfer"
        dscr_col = "Most Recent DSCR (NOI)"

        mask = pd.Series(False, index=active.index)
        reasons = pd.Series("", index=active.index)

        if status_col in active.columns:
            noncur = active[status_col].map(lambda v: (not _is_blank(v)) and str(v).strip() != "0")
            mask |= noncur
            reasons = reasons.where(~noncur, reasons + "non-current; ")
        if wl_col in active.columns:
            onwl = active[wl_col].map(lambda v: isinstance(v, _dt.date))
            mask |= onwl
            reasons = reasons.where(~onwl, reasons + "on watchlist; ")
        if ss_col in active.columns:
            inss = active[ss_col].map(lambda v: not _is_blank(v))
            mask |= inss
            reasons = reasons.where(~inss, reasons + "special servicing; ")
        if dscr_col in active.columns:
            low = _num(active[dscr_col]) < 1.10
            low = low.fillna(False)
            mask |= low
            reasons = reasons.where(~low, reasons + "DSCR<1.10x; ")

        keep = [c for c in [
            "Prospectus Loan ID", "Loan ID", status_col, dscr_col, wl_col, ss_col,
            "Workout Strategy", "Maturity Date", bal_col,
        ] if c in active.columns]
        wl = active.loc[mask, keep].copy()
        wl.insert(0, "Flag Reason", reasons[mask].str.rstrip("; "))
        st.caption(f"{len(wl)} of {len(active)} active loans flagged "
                   f"(non-current status, on watchlist, in special servicing, or DSCR < 1.10x).")
        st.dataframe(wl, use_container_width=True, height=360)
        st.download_button(
            "Download Watchlist extract",
            data=single_workbook_bytes(wl if len(wl) else active.head(0), "Watchlist"),
            file_name=f"{deal_label}-CMBS-Watchlist_{period}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_wl",
        )

st.markdown(
    "<div style='margin-top:26px;color:#5C6A72;font-size:12px;font-family:IBM Plex Mono,monospace'>"
    "Structural fidelity validated against the sample converted files: headers, row counts, and all "
    "financial values match. MBS/agency deals may still need manual plug-ins per the reporting process."
    "</div>",
    unsafe_allow_html=True,
)
