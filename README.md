# Data Cleaning and Business Analysis using Python and Excel

**CodeOrbit Tech · Data Analyst Internship (1-Month Track) — Task 1**
Intern: Sanjana

A reproducible Python/pandas pipeline that turns a messy 32-row sales CSV into a
validated, analysis-ready dataset and a professional 8-sheet Excel workbook —
with automated data-quality checks, native Excel charts, and a small business
analysis. All results below are computed by the code; nothing is hardcoded.

## Problem statement

Raw business data is rarely analysis-ready. The source file
(`raw_dataset.csv`, synthetic sales orders) contains the classic problems:
**duplicate rows, missing values, four mixed date formats, and inconsistent
whitespace/capitalisation** in text fields. Analysing it as-is would produce
wrong totals, split categories (e.g. `EAST` vs `East `), and unparseable dates.

## Objectives

1. Profile the raw data before touching it (rows, columns, missing, duplicates).
2. Remove exact duplicates; **flag conflicting Order_IDs for investigation**
   instead of silently deleting potentially valid transactions.
3. Handle missing values with documented, column-appropriate rules —
   never delete rows just for missing values, never invent data.
4. Standardise dates, text, and data types.
5. Validate the cleaned data automatically (PASS / WARNING / FAIL) and block
   all outputs if any check fails.
6. Derive `Revenue = Quantity × Unit_Price` and produce a simple, honest
   business analysis (revenue ≠ profit; no unsupported predictions).
7. Deliver a polished Excel workbook and a GitHub-ready repository.

## Technologies

- **Python 3.8+** · **pandas** (cleaning, validation, analysis) ·
  **NumPy** (numeric handling) · **openpyxl** (styled workbook + native charts)
- **Microsoft Excel** (or LibreOffice) to review the workbook
- matplotlib — only needed to regenerate the PNGs in `screenshots/`

## Features

- Safe CSV loading with clear error messages (missing/corrupt file, wrong columns)
- Initial data-profile report printed to the console
- 8 documented cleaning steps, each with issue → action → records affected → reason
- 10 automated validation checks with PASS / WARNING / FAIL
  (WARNING is never presented as a pass)
- Independent test suite: `python tests/test_cleaning.py` (17 tests, no pytest needed)
- 8-sheet Excel workbook with filters, freeze panes, ₹ formatting, conditional
  formatting, and **native Excel bar charts**
- `screenshots/` with exported charts for the README and LinkedIn post
- 100% relative paths — runs on any machine, no cloud account or credentials

## Data-cleaning methodology

| Step | Issue | Action | Records | Reason |
|---|---|---|---|---|
| 0 | Conflicting duplicate Order_IDs | Flag for manual investigation (none found) | 0 | Same ID + different data may be two real transactions; deleting could destroy valid sales |
| 1 | 2 exact duplicate rows (`ORD-1004`, `ORD-1018`) | `drop_duplicates(keep='first')` | 32 → 30 rows | Every field identical = same order recorded twice; first is the original |
| 2 | 1 missing `Quantity` (`ORD-1013`) | Filled with column median (6), as integer | 1 | Median resists outliers; preserves a real order instead of deleting it |
| 3 | 1 missing `Category` (`ORD-1006`), 1 missing `Region` (`ORD-1022`) | Filled with `'Unknown'` | 2 | A label can't be derived; `'Unknown'` flags the gap honestly |
| 4 | 23 dates in 4 mixed formats | Parsed each format → `YYYY-MM-DD` | 23 | One ISO format is unambiguous and sorts correctly |
| 5 | Whitespace / random casing in text fields | Trimmed + canonical casing via lookup | 71 cells | `EAST` vs `East ` would split every group-by and chart |
| 6 | `Quantity`/`Unit_Price` stored as text | `Quantity` → int, `Unit_Price` → float (2 dp) | all rows | Correct types prevent silent errors and enable arithmetic |
| 7 | No revenue figure | `Revenue = Quantity × Unit_Price`, added **after** cleaning | all rows | Derived metrics must rest on cleaned values |

## Validation methodology

`validate_cleaned()` runs 10 checks after cleaning. Any **FAIL** stops the run
with exit code 1 and **no output files are written**. **WARNING**s (e.g. the two
documented `'Unknown'` placeholders) are reported honestly and never shown as
passes. The independent suite in `tests/` re-verifies the pipeline end to end.

| Check | Result |
|---|---|
| Missing values (all columns) | ✅ PASS — 0 |
| Exact duplicate rows | ✅ PASS — 0 |
| Unique Order_IDs | ✅ PASS — 30 of 30 |
| Conflicting duplicate Order_IDs | ✅ PASS — 0 |
| Invalid quantities (≤ 0) | ✅ PASS — 0 |
| Invalid unit prices (≤ 0 / missing) | ✅ PASS — 0 |
| Invalid / unparseable dates | ✅ PASS — 0 |
| Revenue = Quantity × Unit_Price (every row) | ✅ PASS — consistent |
| Row reconciliation (32 − 2 = 30) | ✅ PASS |
| Placeholder `'Unknown'` values | ⚠️ WARNING — Category: 1, Region: 1 (documented, acceptable) |

**Data-quality summary:** 32 → 30 rows · 2 duplicates removed · 3 missing values
resolved · 23 dates standardised · 71 text cells standardised.

## Verified results (computed by the pipeline)

- **Total revenue:** Rs.2,021,479.80 across 30 orders
- **Average order value:** Rs.67,382.66
- **Top region:** North — Rs.795,828.88 (39.4%), followed by South Rs.634,083.37,
  East Rs.329,671.55, West Rs.257,437.28
- **Top product by revenue:** Laptop — Rs.1,141,479.30 (high unit price × modest
  volume beats low price × high volume here)
- **Top products by units:** Desk Lamp (32), Standing Desk (32), Gel Pen Pack (26)
- **Interpretation:** revenue concentration is high — the top 2 products drive
  ~84% of sales value. Note this is *revenue*, not profit: costs aren't in the
  dataset, so no margin claims are made.

### Workbook screenshots

![Revenue by region](screenshots/revenue_by_region.png)
![Top 5 products by revenue](screenshots/revenue_by_product.png)
![Data quality before vs after](screenshots/data_quality_before_after.png)

### Workbook sheets

| Sheet | Contents |
|---|---|
| `README_Summary` | Overview, methodology, key results, reproduce steps |
| `Raw_Data` | Original 32-row dataset, untouched |
| `Cleaned_Data` | Final 30-row dataset + `Revenue`, ₹ formatting |
| `Cleaning_Log` | Issue, action, records affected, reason, example |
| `Data_Quality` | Before → after quality table |
| `Revenue_by_Region` | Summary table + native Excel bar chart |
| `Revenue_by_Product` | Top-5 table + native Excel bar chart |
| `Validation_Summary` | Automated checks with PASS/WARNING/FAIL + conditional formatting |

## Installation and execution

### Windows

```bat
:: 1. Open a terminal in the project folder
cd CodeOrbit_DataCleaningExcel

:: 2. Create and activate a virtual environment
py -m venv venv
venv\Scripts\activate

:: 3. Install dependencies
pip install -r requirements.txt

:: 4. Run the pipeline
python build_task1.py

:: 5. Run the independent test suite
python tests/test_cleaning.py
```

### macOS / Linux

```bash
cd CodeOrbit_DataCleaningExcel
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python build_task1.py
python tests/test_cleaning.py
```

### Expected outputs

- Console: initial data profile → 8 cleaning steps → 10 validation checks →
  output summary
- `cleaned_dataset.csv` — 30 rows × 8 columns (written only if validation has no FAIL)
- `Task1_Data_Cleaning_Excel.xlsx` — 8 sheets (written only if validation has no FAIL)
- `raw_dataset.csv` — never modified

### Troubleshooting

| Problem | Fix |
|---|---|
| `Raw dataset not found` error | Run the commands from the project folder (the one containing `raw_dataset.csv`) |
| `pip` not recognised (Windows) | Use `py -m pip install -r requirements.txt` |
| Charts don't render | Open the `.xlsx` in Excel/LibreOffice — charts are native Excel objects |
| `python` vs `python3` | On Windows use `python`/`py`; on macOS/Linux use `python3` |
| Want fresh screenshots | `pip install matplotlib`, then run the generator in `screenshots/` |

## Project structure

```
CodeOrbit_DataCleaningExcel/
├── README.md                      # this file
├── build_task1.py                 # pipeline: profile → clean → validate → analyse → workbook
├── requirements.txt               # pandas, numpy, openpyxl
├── raw_dataset.csv                # source data — 32 rows, NEVER overwritten
├── cleaned_dataset.csv            # cleaned data — 30 rows + Revenue (generated)
├── Task1_Data_Cleaning_Excel.xlsx # 8-sheet workbook (generated)
├── tests/
│   └── test_cleaning.py           # 17 independent tests, zero extra dependencies
├── screenshots/                   # exported charts for README / LinkedIn
│   ├── revenue_by_region.png
│   ├── revenue_by_product.png
│   └── data_quality_before_after.png
└── .gitignore
```

## Limitations and future improvements

- **Synthetic data** (seed 20261010): patterns illustrate technique, not a real business.
- **Small sample** (30 orders): medians and shares are illustrative; a larger
  dataset would support sturdier statistics.
- **Median imputation** assumes the missing quantity is typical; with real data,
  check *why* values are missing first.
- **Canonical text mapping** uses a lookup table — new unseen products fall back
  to title case and would need review.
- **No cost data**, so profit/margin analysis isn't possible.
- Future: monthly trend analysis, a Power BI dashboard build, and exporting the
  validation report as PDF.

## GitHub repository

Repository URL: `[your-github-repo-url]` ← replace with your repo link after upload.

Upload the extracted project files (not the zip) to a public repo named
`CodeOrbit_DataCleaningExcel`, then paste the repo link into the internship
submission form shared in the WhatsApp group.
