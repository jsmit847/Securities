"""
CREFC Investor-Reporting Converter & Monthly Template Builder
=============================================================

Two modes:
  1. Convert raw files  - turn the headerless Trustee (Computershare) exports into
     clean, labeled CREFC IRP v8.4 tables (the Text-to-Columns + date fixing step).
  2. Build reporting template - drop the raw files for one or more deals plus your
     monthly reporting template, and get the template back with the data tabs
     refreshed and the dashboard formulas ready to recalc.

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
from crefc_template import build_template, CORE_TABS


# ----------------------------------------------------------------------------- 
# Helpers
# ----------------------------------------------------------------------------- 
def _dedup_columns(df):
    """Unique column labels for on-screen display (Excel keeps the real duplicates)."""
    seen: dict = {}
    new = []
    for c in df.columns:
        if c in seen:
            seen[c] += 1
            new.append(f"{c} ({seen[c]})")
        else:
            seen[c] = 0
            new.append(c)
    out = df.copy()
    out.columns = new
    return out


def _is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    return str(v).strip() in ("", "nan", "None")


def _num(series):
    return pd.to_numeric(series, errors="coerce")


def _suggest_deal(files) -> str:
    import re as _re
    for f in files or []:
        base = f.name.rsplit("/", 1)[-1]
        m = _re.search(r"([A-Za-z]{2,}[_-]?\d{3,})_(?:CBND|CCOL|LPER|PROP|FINF|DDST|RSRV)", base, _re.I)
        if m:
            return m.group(1).replace("_", "")
    return "DEAL"


def classify(files):
    """Return a plan: list of (uploaded_file, chosen_type_or_flag)."""
    plan = []
    for uf in files:
        low = uf.name.lower()
        if low.endswith(".pdf"):
            plan.append((uf, "PDF"))
        elif low.endswith((".xls", ".xlsx")) and detect_type(uf.name) is None:
            plan.append((uf, "XLS"))
        else:
            plan.append((uf, detect_type(uf.name) or "(skip)"))
    return plan


def run_conversions(plan, normalize_dates):
    """Convert the recognised raw files. Returns results, by_type, passthrough, meta, period."""
    results = {}       # sheet_name -> DataFrame (last one wins; used in convert mode)
    by_type = {}       # file_type -> [DataFrame, ...] across deals (template mode)
    passthrough = []   # (name, bytes)
    meta = []          # blotter rows: (txid, name, label, dim, badge, detail)
    period = None

    for uf, choice in plan:
        data = uf.getvalue()
        if choice in ("PDF", "XLS"):
            passthrough.append((uf.name, data))
            meta.append(("—", uf.name, "passthrough", "—", "PASS", "not reparsed"))
            continue
        if choice == "(skip)":
            meta.append(("—", uf.name, "unrecognized", "—", "SKIP", "choose a type below"))
            continue
        try:
            res = convert(data, choice, normalize_dates=normalize_dates)
        except Exception as exc:  # noqa: BLE001
            meta.append(("—", uf.name, choice, "—", "ERR", str(exc)))
            continue
        period = period or infer_period(res.df)
        sheet = FILE_TYPES[choice]["sheet"]
        results[sheet] = res.df
        by_type.setdefault(choice, []).append(res.df)
        txid = str(res.df.iloc[0, 0]).strip() if len(res.df) else "—"
        badge = "OK" if res.ok else "CHK"
        detail = "; ".join(res.messages) if res.messages else (
            "best-effort headers" if res.best_effort else "clean")
        meta.append((txid, uf.name, FILE_TYPES[choice]["label"], f"{res.n_rows:,} × {res.n_cols}",
                     badge, detail))
    return results, by_type, passthrough, meta, (period or _dt.date.today().strftime("%Y_%m"))


def blotter_html(meta):
    rows = []
    for tkr, name, label, dim, badge, detail in meta:
        cls = {"OK": "ok", "PASS": "ok", "CHK": "warn", "SKIP": "warn", "ERR": "err"}.get(badge, "warn")
        rows.append(
            f"<div class='blot-row'><span class='blot-tkr'>{tkr}</span>"
            f"<span class='blot-name'>{name}<br>"
            f"<span style='color:#5C6A72;font-size:11px'>{label} · {detail}</span></span>"
            f"<span class='blot-dim'>{dim}</span>"
            f"<span class='blot-badge {cls}'>{badge}</span></div>")
    return f"<div class='blotter'>{''.join(rows)}</div>"


# ----------------------------------------------------------------------------- 
# Page + theme
# ----------------------------------------------------------------------------- 
st.set_page_config(page_title="CREFC Reporting Toolkit", page_icon="▤", layout="wide")

st.markdown(
    """
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;700&family=IBM+Plex+Mono:wght@400;500;600&family=Inter:wght@400;500;600&display=swap');
      :root{ --ink:#14202B; --paper:#F5F6F4; --panel:#FFFFFF; --gold:#B8862F;
             --ok:#2E6E4E; --alert:#B03A2E; --line:#D7DBD5; --muted:#5C6A72; }
      html, body, [class*="css"]{ font-family:'Inter',system-ui,sans-serif; color:var(--ink); }
      .stApp{ background:var(--paper); }
      .block-container{ padding-top:2.2rem; max-width:1200px; }
      .mast{ border-top:3px solid var(--ink); border-bottom:1px solid var(--line);
             padding:14px 0 12px; margin-bottom:6px; }
      .mast .eyebrow{ font-family:'IBM Plex Mono',monospace; font-size:11px; letter-spacing:.28em;
             text-transform:uppercase; color:var(--gold); font-weight:600; }
      .mast h1{ font-family:'Space Grotesk',sans-serif; font-weight:700; font-size:30px;
             letter-spacing:-.01em; margin:.15rem 0 .1rem; }
      .mast .sub{ color:var(--muted); font-size:13.5px; }
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
      h2, h3{ font-family:'Space Grotesk',sans-serif; letter-spacing:-.01em; }
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

# ----------------------------------------------------------------------------- 
# Sidebar
# ----------------------------------------------------------------------------- 
with st.sidebar:
    st.markdown("### Mode")
    mode = st.radio(
        "What do you want to do?",
        ["Convert raw files", "Build reporting template"],
        label_visibility="collapsed",
    )
    st.divider()
    st.markdown("### Settings")
    date_mode = st.radio(
        "Date fields",
        ["Normalize to YYYY-MM-DD (recommended)", "Keep raw (YYYYMMDD)"],
        help="Trustee files store dates as YYYYMMDD integers. Normalizing writes real Excel dates, "
             "which the template's maturity formulas need.",
    )
    normalize_dates = date_mode.startswith("Normalize")
    st.divider()
    st.markdown(
        "<div style='font-family:IBM Plex Mono,monospace;font-size:11px;color:#5C6A72;line-height:1.7'>"
        "TYPE&nbsp;&nbsp;→ CONTENTS<br>CBND → Bond / certificate<br>CCOL → Collateral summary<br>"
        "LPER → Loan periodic<br>PROP → Property<br>FINF → Financial (best-effort)</div>",
        unsafe_allow_html=True,
    )

st.markdown(
    f"""
    <div class="mast">
      <div class="eyebrow">Investor Reporting · CREFC IRP v8.4</div>
      <h1>{'Raw → Converted File Converter' if mode == 'Convert raw files' else 'Monthly Reporting Template Builder'}</h1>
      <div class="sub">{'Drop the Trustee-site files (CBND · CCOL · LPER · PROP · FINF). Get labeled, date-clean tables — the Text-to-Columns step, automated.' if mode == 'Convert raw files' else 'Drop your reporting template plus the raw files for the deals. The data tabs are refreshed and the dashboard recomputes on open.'}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ============================================================================= 
# MODE 1 — CONVERT RAW FILES
# ============================================================================= 
if mode == "Convert raw files":
    uploaded = st.file_uploader(
        "Upload raw Trustee files",
        type=["txt", "csv", "xls", "xlsx", "pdf"], accept_multiple_files=True,
        help="Type is detected from the name (…_LPER.txt etc.). Override anything below.",
    )
    deal_label = st.sidebar.text_input("Deal label (output prefix)", value=_suggest_deal(uploaded),
                                       help="Auto-filled from the file names, e.g. CVAF20182.")

    if not uploaded:
        st.markdown("<div class='note'>Waiting for files. Drop the raw <code>.txt</code> exports "
                    "(e.g. <code>CVAF_20182_LPER.txt</code>). The statement PDF and the "
                    "<code>RSRV.xls</code> reserve file are passed through, not reparsed.</div>",
                    unsafe_allow_html=True)
        st.stop()

    plan = classify(uploaded)
    with st.expander("Review detected file types", expanded=False):
        fixed = []
        TYPE_CHOICES = ["(skip)"] + list(FILE_TYPES.keys())
        for uf, choice in plan:
            if choice in ("PDF", "XLS"):
                st.write(f"{'📄' if choice=='PDF' else '📊'} `{uf.name}` — passed through.")
                fixed.append((uf, choice))
            else:
                sel = st.selectbox(f"`{uf.name}`", TYPE_CHOICES,
                                   index=TYPE_CHOICES.index(choice) if choice in TYPE_CHOICES else 0,
                                   key=f"c_{uf.name}")
                fixed.append((uf, sel))
        plan = fixed

    results, by_type, passthrough, meta, period = run_conversions(plan, normalize_dates)

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f"<div class='stat'><b>{len(results)}</b><span>Tables converted</span></div>", unsafe_allow_html=True)
    c2.markdown(f"<div class='stat'><b>{sum(len(d) for d in results.values()):,}</b><span>Data rows</span></div>", unsafe_allow_html=True)
    c3.markdown(f"<div class='stat'><b>{period}</b><span>Reporting period</span></div>", unsafe_allow_html=True)
    c4.markdown(f"<div class='stat'><b>v8.4</b><span>CREFC IRP</span></div>", unsafe_allow_html=True)

    st.markdown("### Conversion blotter")
    st.markdown(blotter_html(meta), unsafe_allow_html=True)
    if not results:
        st.stop()

    st.markdown("### Download")
    d1, d2 = st.columns(2)
    d1.download_button("Combined workbook (one tab per table)", data=combined_workbook_bytes(results),
                       file_name=f"{deal_label}-CREFC-Converted_{period}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)
    zbuf = io.BytesIO()
    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
        for sheet, df in results.items():
            out = FILE_TYPES[[k for k, v in FILE_TYPES.items() if v["sheet"] == sheet][0]]["out"]
            zf.writestr(f"{deal_label}-{out}_{period}.xlsx", single_workbook_bytes(df, sheet))
        for name, data in passthrough:
            zf.writestr(f"_passthrough/{name}", data)
    d2.download_button("All files (.zip — individual workbooks)", data=zbuf.getvalue(),
                       file_name=f"{deal_label}-CREFC-Converted_{period}.zip", mime="application/zip",
                       use_container_width=True)

    st.markdown("### Tables")
    for sheet, df in results.items():
        out = FILE_TYPES[[k for k, v in FILE_TYPES.items() if v["sheet"] == sheet][0]]["out"]
        with st.expander(f"{sheet}  ·  {len(df):,} rows × {len(df.columns)} cols", expanded=False):
            st.dataframe(_dedup_columns(df.head(200)), use_container_width=True, height=340)
            st.caption(f"Showing up to 200 rows of {len(df):,}.")
            st.download_button(f"Download {sheet}", data=single_workbook_bytes(df, sheet),
                               file_name=f"{deal_label}-{out}_{period}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key=f"dl_{sheet}")

    periodic = results.get("Periodic")
    if periodic is not None:
        st.markdown("### Surveillance extracts")
        st.markdown("<div class='note'>Starter Maturity and Watchlist views from the Loan Periodic file, "
                    "scoped to active loans (ending balance &gt; 0).</div>", unsafe_allow_html=True)
        bal_col = "Current Ending Scheduled  Balance"
        active = periodic[_num(periodic[bal_col]) > 0].copy() if bal_col in periodic.columns else periodic.copy()
        tmat, twl = st.tabs([f"Maturity ladder ({len(active)})", "Watchlist extract"])
        with tmat:
            keep = [c for c in ["Prospectus Loan ID", "Loan ID", "Maturity Date", bal_col,
                                "Current Note Rate", "Most Recent DSCR (NOI)",
                                "Payment Status of Loan   (fka Status of Loan)", "Number of Properties"]
                    if c in active.columns]
            mat = active[keep].copy()
            if "Maturity Date" in mat.columns:
                today = _dt.date.today()
                mat["Months to Maturity"] = mat["Maturity Date"].map(
                    lambda d: (d.year - today.year) * 12 + (d.month - today.month)
                    if isinstance(d, _dt.date) else None)
                mat = mat.sort_values("Maturity Date", na_position="last")
            st.dataframe(_dedup_columns(mat), use_container_width=True, height=360)
            st.download_button("Download Maturity ladder", data=single_workbook_bytes(mat, "Maturity"),
                               file_name=f"{deal_label}-WL-Maturity_{period}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dlmat")
        with twl:
            sc = "Payment Status of Loan   (fka Status of Loan)"
            wc = "Date Added to Servicer Watchlist"
            ssc = "Reason for SS Transfer"
            dc = "Most Recent DSCR (NOI)"
            mask = pd.Series(False, index=active.index)
            reasons = pd.Series("", index=active.index)
            if sc in active.columns:
                m = active[sc].map(lambda v: (not _is_blank(v)) and str(v).strip() != "0")
                mask |= m; reasons = reasons.where(~m, reasons + "non-current; ")
            if wc in active.columns:
                m = active[wc].map(lambda v: isinstance(v, _dt.date))
                mask |= m; reasons = reasons.where(~m, reasons + "on watchlist; ")
            if ssc in active.columns:
                m = active[ssc].map(lambda v: not _is_blank(v))
                mask |= m; reasons = reasons.where(~m, reasons + "special servicing; ")
            if dc in active.columns:
                m = (_num(active[dc]) < 1.10).fillna(False)
                mask |= m; reasons = reasons.where(~m, reasons + "DSCR<1.10x; ")
            keep = [c for c in ["Prospectus Loan ID", "Loan ID", sc, dc, wc, ssc, "Workout Strategy",
                                "Maturity Date", bal_col] if c in active.columns]
            wl = active.loc[mask, keep].copy()
            wl.insert(0, "Flag Reason", reasons[mask].str.rstrip("; "))
            st.caption(f"{len(wl)} of {len(active)} active loans flagged.")
            st.dataframe(_dedup_columns(wl), use_container_width=True, height=360)
            st.download_button("Download Watchlist extract",
                               data=single_workbook_bytes(wl if len(wl) else active.head(0), "Watchlist"),
                               file_name=f"{deal_label}-CMBS-Watchlist_{period}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dlwl")

# ============================================================================= 
# MODE 2 — BUILD REPORTING TEMPLATE
# ============================================================================= 
else:
    st.markdown("<div class='note'>1) Upload your reporting template. 2) Upload the raw files for the "
                "deals you downloaded this month. The four core tabs (Coll Summ, Periodic, Property, Bond) "
                "are refreshed by transaction ID; the Dlq and REO surveillance tabs are separate CTS "
                "downloads and are left untouched.</div>", unsafe_allow_html=True)

    tpl = st.file_uploader("1 · Reporting template (.xlsx)", type=["xlsx"], accept_multiple_files=False)
    uploaded = st.file_uploader("2 · Raw Trustee files (one or many deals)",
                                type=["txt", "csv", "xls", "xlsx", "pdf"], accept_multiple_files=True)

    if not tpl or not uploaded:
        st.stop()

    plan = classify(uploaded)
    results, by_type, passthrough, meta, period = run_conversions(plan, normalize_dates)

    # deals detected (transaction IDs) from the four core types
    deals = sorted({str(df.iloc[0, 0]).strip() for ft in CORE_TABS for df in by_type.get(ft, []) if len(df)})
    core_present = [ft for ft in CORE_TABS if ft in by_type]

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f"<div class='stat'><b>{len(deals)}</b><span>Deals detected</span></div>", unsafe_allow_html=True)
    c2.markdown(f"<div class='stat'><b>{len(core_present)}/4</b><span>Core tabs fed</span></div>", unsafe_allow_html=True)
    c3.markdown(f"<div class='stat'><b>{period}</b><span>Reporting period</span></div>", unsafe_allow_html=True)
    c4.markdown(f"<div class='stat'><b>{sum(len(d) for v in by_type.values() for d in v):,}</b><span>Rows loaded</span></div>", unsafe_allow_html=True)

    st.markdown("### Conversion blotter")
    st.markdown(blotter_html(meta), unsafe_allow_html=True)
    if deals:
        st.caption("Deals to refresh: " + ", ".join(deals))

    missing = [ft for ft in CORE_TABS if ft not in by_type]
    if missing:
        st.markdown(f"<div class='note'>Heads up — no {', '.join(missing)} file(s) uploaded, so those "
                    "tabs won't be refreshed for these deals.</div>", unsafe_allow_html=True)

    if not by_type:
        st.stop()

    if st.button("Build reporting template", type="primary"):
        with st.spinner("Refreshing data tabs and wiring up the dashboard…"):
            out_bytes, log = build_template(tpl.getvalue(), by_type)
        st.success("Template built. It will recalculate automatically when you open it in Excel.")
        for line in log:
            st.write("· " + line)
        out_name = tpl.name.replace(".xlsx", "") + f"_refreshed_{period}.xlsx"
        st.download_button("Download refreshed template", data=out_bytes, file_name=out_name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
        st.markdown("<div class='note'><b>Note on Dlq / REO:</b> the Delinquent and REO Report tabs come "
                    "from separate Computershare surveillance downloads, so they're preserved as-is. Paste "
                    "the new month's Dlq/REO reports into those tabs to refresh the advance, delinquency, "
                    "and REO sections of the dashboard.</div>", unsafe_allow_html=True)

st.markdown("<div style='margin-top:26px;color:#5C6A72;font-size:12px;font-family:IBM Plex Mono,monospace'>"
            "Validated against sample converted files for two deals: headers, row counts, and financial "
            "values match; the build introduces no new formula errors. MBS/agency deals may still need "
            "manual plug-ins.</div>", unsafe_allow_html=True)
