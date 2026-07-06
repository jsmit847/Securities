# CREFC Reporting Toolkit

A Streamlit app for the CoreVest CMBS monthly reporting process. It replaces the
manual "download → Text-to-Columns → fix dates → paste into the template" steps
with two modes:

1. **Convert raw files** — turn the headerless Trustee (Computershare) exports
   into clean, labeled **CREFC IRP v8.4** tables.
2. **Build reporting template** — drop your monthly reporting template plus the
   raw files for the deals, and get the template back with the data tabs
   refreshed and the dashboard set to recalculate on open.

---

## Run it

```bash
pip install -r requirements.txt
streamlit run app.py
```

Open the local URL Streamlit prints (usually http://localhost:8501).

---

## Mode 1 — Convert raw files

Drop the raw `.txt` exports. The app:

1. Parses each with a real CSV reader (the Property file has quoted fields with
   embedded commas, so a naive split corrupts it).
2. Applies the standardized CREFC 8.4 header row for that file type.
3. Types each column the way Excel's Text-to-Columns does — an all-numeric column
   becomes numbers (zero-padded IDs like `030299963` -> `30299963`), a column with
   any letter/code stays text.
4. Normalizes dates from `YYYYMMDD` to real Excel dates (`YYYY-MM-DD`).
5. Offers per-file downloads, a combined workbook, a zip, and starter
   **Maturity** and **Watchlist** extracts.

| Raw token | Contents            | Output tab |
|-----------|---------------------|------------|
| `CBND`    | Bond / certificate  | Bond       |
| `CCOL`    | Collateral summary  | CollSUM    |
| `LPER`    | Loan periodic       | Periodic   |
| `PROP`    | Property            | Property   |
| `FINF`    | Financial (op stmt) | Financial  |
| `DDST`.pdf| Statement           | passthrough|
| `RSRV`.xls| Reserve / LOC       | passthrough|

## Mode 2 — Build reporting template

1. Upload your reporting template (`.xlsx`).
2. Upload the raw files for the deals you downloaded this month (one or many
   deals at once — the app groups them by transaction ID).
3. Click **Build reporting template**.

The app refreshes the four core data tabs — `CMBS - Coll Summ Data`,
`CMBS - Periodic Data`, `CMBS - Property Data`, `CMBS - Bond Data` — keyed by
transaction ID (column A), which the Dashboard's `SUMIF/SUMIFS` formulas look up.
Everything else is preserved: the Dashboard and all its formulas, the Transaction
List, the hidden header-file tabs, styles, and merged cells. The workbook is set
to recalculate every formula when you open it in Excel.

**Merge behaviour:** rows for the uploaded deals are replaced; rows for deals you
didn't upload this run are kept. So you can refresh the whole book at once or just
the deals you re-downloaded.

**Dlq / REO tabs:** the Delinquent Loan and REO reports are separate Computershare
surveillance downloads, not CREFC files, so those two tabs are left untouched.
Paste the new month's Dlq/REO reports into them to refresh the advance,
delinquency, and REO sections of the dashboard.

---

## Files

| File                | Purpose                                                    |
|---------------------|------------------------------------------------------------|
| `app.py`            | Streamlit UI (both modes)                                  |
| `crefc_convert.py`  | Parsing / column-typing / date engine                      |
| `crefc_io.py`       | Excel writers (single sheet + combined workbook)           |
| `crefc_template.py` | Merges converted data into the reporting template          |
| `crefc_headers.py`  | CREFC 8.4 header definitions per file type                 |
| `requirements.txt`  | Dependencies                                               |

---

## Validation

Checked against sample converted files for two deals (CoreVest 2018-1 and 2018-2):

- **Headers, row counts, financial values** match; Bond, Coll Summ, and Periodic
  are 0-diff. The engine also fixes two reference bugs (a `0.01` value the manual
  process dropped to `0`, and a `Â` encoding artifact).

Template build, verified after a LibreOffice recalc:

- Dashboard balances match the remittance statements exactly — e.g. COREVEST18-1
  beginning 34,019,241.93 -> ending 32,052,000.59; CAF2018-2 46,381,295.34 ->
  46,307,218.20, with the correct H-tranche bond balances and specially-serviced
  counts (1 and 5).
- **Zero new formula errors.** The template's ~650 `#N/A` cells live in the
  delinquency-detail spill formulas and are present in the original file too.

---

## Notes

- **MBS / agency deals** don't provide the full CREFC file set and still need the
  manual plug-ins described in the process doc.
- **`FINF` headers** are best-effort (no validated reference in the sample set).
- Building the template loads a large workbook (the Property tab can be tens of
  thousands of rows), so the build step takes a little time — that's expected.
