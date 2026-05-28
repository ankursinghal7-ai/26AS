# 26AS Reconciliation Assistant

> Automated reconciliation of **Form 26AS** (TRACES) against the **TDS Receivable ledger** from books of account, with mismatch classification and CA-authentic action notes.

Built by a practising Indian Chartered Accountant for fellow finance professionals tired of manually matching hundreds of TDS entries every quarter. Ships in two flavours: a polished **Streamlit web app** for the demo / synthetic template, and a **production-grade CLI** that handles the messy, multi-sheet TRACES PDF-to-Excel exports and Tally TDS ledgers that you see in real practice.

![Status](https://img.shields.io/badge/status-portfolio-blue) ![Python](https://img.shields.io/badge/python-3.10%2B-blue) ![License](https://img.shields.io/badge/license-MIT-green) ![Streamlit](https://img.shields.io/badge/streamlit-cloud--ready-FF4B4B)

---

## Why this exists

Every quarter-end and year-end, Indian assessees reconcile:

- **Form 26AS** — what their customers reported to the income-tax department as TDS deducted
- **Books of account** — what they themselves recorded as TDS receivable

For a mid-sized B2B company this can be hundreds to thousands of line items, spread across 200+ vendors, with mismatches caused by Q4 filing lag, TAN changes, vendor name spelling variants, branch suffixes, section misclassification, and amount tolerances. Doing this by hand is a 2-3 day exercise. This tool does it in under a minute and tells you exactly which entries need human eyes.

---

## What it does

The matching engine works through five mismatch categories:

| Category | Trigger | Typical Cause |
|---|---|---|
| **Matched** | Vendor + amount agree | Fully reconciled — no action |
| **26AS Only — Not in Books** | TDS in 26AS, absent in books | March invoice booked in April |
| **Books Only — Not in 26AS** | TDS in books, absent in 26AS | Deductor yet to file Q4 return |
| **Amount Mismatch** | Vendor agrees, amount differs (±2%) | Part-payment, debit/credit note, rate error |
| **Section Mismatch*** | Section in 26AS differs from books | 194C vs 194J misclassification |
| **TAN Mismatch*** | Same vendor, different TAN | Group restructuring, demerger |

*\* Section/TAN mismatch detection requires those columns in the books ledger — see [docs/DATA_FORMATS.md](docs/DATA_FORMATS.md).*

Output is a colour-coded `.xlsx` with three sheets:

- **Detailed Reconciliation** — every entry with action note
- **Executive Summary** — KPI block + category breakdown
- **Vendor Summary** — one row per vendor, sorted by absolute gap (the sheet you actually work from)

---

## Two ways to use it

### 1. Streamlit web app (`app.py`)
For demos, single-CA use, and when your books export already has TAN + Section columns.

![Streamlit screenshot placeholder](docs/screenshots/01_streamlit_home.png)

### 2. Real-world CLI (`scripts/reconcile_real.py`)
For the actual TRACES multi-sheet Excel + Tally TDS ledger you have on your desk. Vendor name fuzzy matching with location/branch suffix stripping, abbreviation expansion (L&T → Larsen Toubro), and multi-ratio scoring.

```bash
python scripts/reconcile_real.py \
    --file26  ./my_26AS_FY24-25.xlsx \
    --books   ./my_Books_TDS_Ledger.xlsx \
    --company "Acme Manufacturing Pvt Ltd" \
    --fy      2024-25 \
    --output  ./out/
```

### 3. Synthetic template CLI (`scripts/reconcile.py`)
The original "happy path" CLI from the underlying Claude skill — works when both files follow the canonical template (TAN + Section on both sides).

---

## Installation

Tested on Python 3.10+. Recommend a virtual environment.

```bash
git clone https://github.com/<your-username>/26as-reconciliation.git
cd 26as-reconciliation

python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate

pip install -r requirements.txt
```

---

## Run locally

**Streamlit web app**
```bash
streamlit run app.py
# Opens at http://localhost:8501
```

**Real-world CLI**
```bash
python scripts/reconcile_real.py --file26 26AS.xlsx --books BOOKS_TDS.xlsx \
    --company "Your Co Pvt Ltd" --fy 2024-25 --output ./out/
```

**Synthetic template CLI**
```bash
python scripts/reconcile.py --file26 26AS.xlsx --books BOOKS_TDS.xlsx \
    --company "Your Co Pvt Ltd" --fy 2024-25 --output ./out/
```

---

## Deploy to Streamlit Community Cloud (free)

Step-by-step in [docs/DEPLOYMENT.md](docs/DEPLOYMENT.md). Short version: push to GitHub, connect to share.streamlit.io, pick `app.py`, click Deploy. Done in 3 minutes.

---

## Input file formats

Detailed column-by-column spec in [docs/DATA_FORMATS.md](docs/DATA_FORMATS.md).

Quick view:

- **26AS:** TRACES Annual Tax Statement (Part I), exported to Excel. Either the canonical synthetic template (single sheet, named columns) or the real multi-sheet PDF-to-Excel export (handled by `reconcile_real.py`).
- **Books:** TDS Receivable ledger from Tally, Zoho, SAP, or custom Excel. For best results include TAN and Section columns. The real-world script also handles Tally exports with just Date + Vendor + Debit (TDS amount).

---

## Caveats — read before using on live data

1. **Indicative only.** Use this to drastically narrow what needs human review, not to skip the review.
2. **Vendor-only fuzzy matching loses precision.** If your books ledger has no TAN column, expect ~70-90% auto-match by value depending on vendor naming hygiene. The Vendor Summary sheet shows you exactly what's left over.
3. **No PDF parsing.** Export your 26AS from TRACES as Excel; PDF-only support is on the roadmap.
4. **±2% amount tolerance is configurable but defaulted.** Adjust `AMT_TOL_PCT` in the scripts if your business needs tighter or looser bounds.
5. **Not professional tax advice.** Always verify against the original TRACES PDF and audited books before filing returns or responding to notices.

---

## Roadmap

- [ ] PDF parsing for the original TRACES 26AS download (currently Excel-only)
- [ ] Vendor master file support — let users register canonical vendor names + aliases
- [ ] GSTR-2A reconciliation (same engine, different inputs)
- [ ] Multi-company batch mode
- [ ] Streamlit version of the real-world parser

---

## License

MIT — see [LICENSE](LICENSE). Free to use, modify, and share. Attribution appreciated.

---

## Author

**Ankur Singhal** · Indian Chartered Accountant · 24 years in finance and operations · learning AI to multiply CA productivity.

Built with [Claude](https://claude.ai) as a coding pair. If you're a CA exploring how LLMs can change the workflow, find me on LinkedIn — happy to compare notes.
