#!/usr/bin/env python3
"""
CodeOrbit Tech — Data Analyst Internship
Task 1 (1-month track): Data Cleaning in Excel/Python.

"Data Cleaning and Business Analysis using Python and Excel"

Reproducible pandas pipeline:

    raw_dataset.csv  --(read only, never overwritten)-->  cleaning pipeline
        --> cleaned_dataset.csv + Task1_Data_Cleaning_Excel.xlsx

What the script does:
  1. Loads raw_dataset.csv safely (clear error if missing/corrupt).
  2. Prints an initial data profile (rows, columns, missing, duplicates, dtypes).
  3. Cleans: duplicate rows, conflicting order IDs (flagged, not silently
     deleted), missing values, dates, text, data types.
  4. Validates with PASS / WARNING / FAIL checks. On FAIL, no outputs are written.
  5. Derives Revenue = Quantity x Unit_Price and builds a small business analysis.
  6. Writes cleaned_dataset.csv and the 8-sheet Excel workbook.

Run from the project folder:
    pip install -r requirements.txt
    python build_task1.py

Run the independent test suite:
    python tests/test_cleaning.py
"""
from pathlib import Path
import random
import re
import sys
from datetime import date

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------------ config
# All paths are relative to this script: the project runs on any machine
# (Windows, macOS, Linux, GitHub Codespaces) with no code changes.
BASE_DIR = Path(__file__).resolve().parent
RAW_CSV = BASE_DIR / "raw_dataset.csv"
CLEAN_CSV = BASE_DIR / "cleaned_dataset.csv"
XLSX = BASE_DIR / "Task1_Data_Cleaning_Excel.xlsx"

SEED = 20261010
REQUIRED_COLUMNS = ["Order_ID", "Order_Date", "Product", "Category",
                    "Quantity", "Unit_Price", "Region"]

# Display headers used in the workbook (CSV keeps code-style headers).
DISPLAY_HEADERS = {
    "Order_ID": "Order ID", "Order_Date": "Order Date", "Product": "Product",
    "Category": "Category", "Quantity": "Quantity",
    "Unit_Price": "Unit Price (Rs.)", "Region": "Region",
    "Revenue": "Revenue (Rs.)",
}

# ------------------------------------------------- raw dataset (seed data)
# Synthetic sales orders. Used ONLY to create raw_dataset.csv the first time;
# afterwards the CSV on disk is the source of truth and is never overwritten.
PRODUCTS = {
    "Laptop": ("Electronics", 52000), "Wireless Mouse": ("Electronics", 850),
    "Mechanical Keyboard": ("Electronics", 3200), "USB-C Hub": ("Electronics", 2100),
    "Office Chair": ("Furniture", 9400), "Standing Desk": ("Furniture", 18500),
    "Desk Lamp": ("Furniture", 1400), "Notebook Set": ("Stationery", 420),
    "Gel Pen Pack": ("Stationery", 180), "Backpack": ("Accessories", 2300),
    "Water Bottle": ("Accessories", 650), "Monitor Stand": ("Accessories", 1750),
}
REGIONS = ["North", "South", "East", "West"]
DATE_FORMATS = ("%Y-%m-%d", "%d/%m/%Y", "%b %d, %Y", "%Y.%m.%d")


def _date_variant(d, k):
    return [d.strftime("%Y-%m-%d"), d.strftime("%d/%m/%Y"),
            d.strftime("%b %d, %Y"), d.strftime("%Y.%m.%d")][k % 4]


def _messy(text, mode):
    return {0: text, 1: text.lower(), 2: text.upper(),
            3: "  " + text, 4: text + "   ", 5: " " + text.title() + " "}[mode]


def generate_raw_dataset() -> pd.DataFrame:
    """Deterministically build the messy 32-row source dataset (seed 20261010)."""
    random.seed(SEED)
    np.random.seed(SEED)
    rows = []
    base = pd.Timestamp("2024-01-05")
    for i in range(1, 31):
        prod = random.choice(list(PRODUCTS))
        cat, bp = PRODUCTS[prod]
        dt = base + pd.Timedelta(days=random.randint(0, 330))
        rows.append({"Order_ID": f"ORD-{1000 + i}",
                     "Order_Date": _date_variant(dt, i),
                     "Product": _messy(prod, random.randint(0, 5)),
                     "Category": _messy(cat, random.randint(0, 5)),
                     "Quantity": random.randint(1, 12),
                     "Unit_Price": round(bp * random.uniform(0.92, 1.08), 2),
                     "Region": _messy(random.choice(REGIONS), random.randint(0, 5))})
    raw = pd.DataFrame(rows)
    raw = pd.concat([raw, raw.iloc[[3, 17]]], ignore_index=True)  # 2 duplicate rows
    raw.loc[5, "Category"] = np.nan                               # 3 missing values
    raw.loc[12, "Quantity"] = np.nan
    raw.loc[21, "Region"] = np.nan
    return raw


def load_raw() -> pd.DataFrame:
    """Load the raw dataset without modifying it.

    Raises a clear, actionable error if the file is missing or unreadable
    instead of a raw traceback.
    """
    if not RAW_CSV.exists():
        raise FileNotFoundError(
            f"Raw dataset not found: {RAW_CSV}\n"
            "Make sure you are running 'python build_task1.py' from the "
            "project folder (the folder containing raw_dataset.csv).")
    try:
        raw = pd.read_csv(RAW_CSV, dtype=str)
    except Exception as exc:  # corrupt file, permission issue, ...
        raise ValueError(f"Could not read {RAW_CSV.name}: {exc}\n"
                         "The raw file may be corrupt — restore it from Git.") from exc
    missing_cols = [c for c in REQUIRED_COLUMNS if c not in raw.columns]
    if missing_cols:
        raise ValueError(f"{RAW_CSV.name} is missing required columns: "
                         f"{missing_cols}. Expected: {REQUIRED_COLUMNS}")
    return raw.replace({"": np.nan})


def profile_raw(raw: pd.DataFrame) -> dict:
    """Initial data profile: counts, missing values, duplicates, dtypes."""
    return {
        "rows": len(raw),
        "columns": len(raw.columns),
        "missing_total": int(raw.isna().sum().sum()),
        "missing_by_col": {c: int(raw[c].isna().sum())
                           for c in raw.columns if raw[c].isna().any()},
        "duplicate_rows": int(raw.duplicated(keep="first").sum()),
        "dtypes": {c: str(raw[c].dtype) for c in raw.columns},
    }


def print_profile(p: dict):
    print("Initial data profile (raw_dataset.csv):")
    print(f"  Rows: {p['rows']} | Columns: {p['columns']}")
    print(f"  Missing values (total): {p['missing_total']}")
    for col, n in p["missing_by_col"].items():
        print(f"    - {col}: {n}")
    print(f"  Exact duplicate rows: {p['duplicate_rows']}")

# ------------------------------------------------------- cleaning pipeline
CANON = {p.lower(): p for p in list(PRODUCTS)}
CANON.update({c.lower(): c for c in ["Electronics", "Furniture", "Stationery",
                                     "Accessories", "Unknown"]})
CANON.update({r.lower(): r for r in REGIONS + ["Unknown"]})


def _canon(text):
    t = str(text).strip()
    return CANON.get(t.lower(), t.title())


def _parse_date(s):
    """Parse one of the 4 known messy formats -> pd.Timestamp (NaT if unknown)."""
    s = str(s).strip()
    for fmt in DATE_FORMATS:
        try:
            return pd.to_datetime(s, format=fmt)
        except (ValueError, TypeError):
            continue
    return pd.NaT  # unknown format: flagged by validation, never guessed


def find_conflicting_order_ids(raw: pd.DataFrame) -> list:
    """Order IDs appearing on rows that are NOT exact duplicates.

    Same ID + different field values = a possible data-entry conflict that a
    human should investigate. These are reported, never silently deleted.
    """
    conflicts = []
    for oid, grp in raw.groupby("Order_ID", sort=False):
        if len(grp) > 1 and len(grp.drop_duplicates()) > 1:
            conflicts.append(oid)
    return conflicts


def clean_data(raw: pd.DataFrame):
    """Clean the raw frame.

    Returns (cleaned_df, log_rows, quality, examples). Every decision is
    logged with its reasoning; uncertain values are flagged, never invented.
    """
    log_rows = [["Step", "Issue", "Action taken", "Records affected", "Reason",
                 "Example"]]
    examples = []

    # ---- step 0: conflicting order IDs (investigate, don't delete) -----------
    conflicts = find_conflicting_order_ids(raw)
    log_rows.append(
        [0, "Conflicting duplicate Order_IDs",
         ("Flagged for manual investigation — kept in the dataset"
          if conflicts else "None found — no action needed"),
         f"{len(conflicts)} Order_ID(s)" + (f": {', '.join(conflicts)}" if conflicts else ""),
         "Same ID with different details may be two real transactions sharing "
         "an ID by mistake; deleting them could destroy valid sales.",
         "—"])

    # ---- step 1: exact duplicate rows ----------------------------------------
    dup_mask = raw.duplicated(keep="first")
    dup_ids = raw.loc[dup_mask, "Order_ID"].tolist()
    n_dup = int(dup_mask.sum())
    clean = raw[~dup_mask].reset_index(drop=True)
    for oid in dup_ids:
        examples.append((oid, "exact duplicate row → removed (kept first)"))
    log_rows.append(
        [1, f"Exact duplicate rows ({n_dup} found: {', '.join(dup_ids)})",
         "Removed with drop_duplicates(keep='first')",
         f"{len(raw)} → {len(clean)} rows",
         "Every field identical = the same order recorded twice; the first "
         "occurrence is the original.",
         f"{dup_ids[0]}: identical row deleted" if dup_ids else "—"])

    # ---- step 2: missing numeric ----------------------------------------------
    clean["Quantity"] = pd.to_numeric(clean["Quantity"], errors="coerce")
    qty_median = float(clean["Quantity"].median())
    qty_ids = clean.loc[clean["Quantity"].isna(), "Order_ID"].tolist()
    n_qty_na = len(qty_ids)
    clean["Quantity"] = clean["Quantity"].fillna(qty_median).round().astype(int)
    for oid in qty_ids:
        examples.append((oid, f"Quantity blank → {qty_median:.0f} (median)"))
    log_rows.append(
        [2, f"Missing Quantity ({n_qty_na}: {', '.join(qty_ids)})",
         f"Filled with column median ({qty_median:.0f}), stored as integer",
         f"{n_qty_na} row(s)",
         "Median resists outliers (bulk laptop orders); filling preserves a "
         "real order instead of deleting it for one blank cell.",
         f"{qty_ids[0]}: (blank) → {qty_median:.0f}" if qty_ids else "—"])

    # ---- step 3: missing categoricals ------------------------------------------
    for col in ["Category", "Region"]:
        na_ids = clean.loc[clean[col].isna(), "Order_ID"].tolist()
        clean[col] = clean[col].fillna("Unknown")
        for oid in na_ids:
            examples.append((oid, f"{col} blank → 'Unknown'"))
        log_rows.append(
            [3, f"Missing {col} ({len(na_ids)}: {', '.join(na_ids)})",
             "Filled with 'Unknown'",
             f"{len(na_ids)} row(s)",
             "A label cannot be derived from other columns; 'Unknown' flags "
             "the gap honestly instead of inventing business data.",
             f"{na_ids[0]}: (blank) → Unknown" if na_ids else "—"])

    # ---- step 4: dates ----------------------------------------------------------
    iso_re = re.compile(r"^\d{4}-\d{2}-\d{2}$")
    raw_dates = clean["Order_Date"].astype(str).str.strip()
    n_non_iso = int((~raw_dates.str.match(iso_re).fillna(False)).sum())
    parsed = clean["Order_Date"].map(_parse_date)
    n_bad = int(parsed.isna().sum())
    clean["Order_Date"] = parsed.dt.strftime("%Y-%m-%d")
    shown = 0
    for _, r in raw[~dup_mask].iterrows():
        if shown >= 3:
            break
        s = str(r["Order_Date"]).strip()
        if not iso_re.match(s):
            oid = r["Order_ID"]
            after = clean.loc[clean["Order_ID"] == oid, "Order_Date"].iloc[0]
            examples.append((oid, f"Order_Date '{s}' → '{after}'"))
            shown += 1
    log_rows.append(
        [4, f"Dates in 4 mixed formats ({n_non_iso} non-ISO values)",
         "Parsed each known format → YYYY-MM-DD",
         f"{n_non_iso} value(s)" + (f"; {n_bad} unparseable!" if n_bad else ""),
         "One ISO format is unambiguous (no DD/MM vs MM/DD confusion) and "
         "sorts correctly as text.",
         examples[-1][1] if shown else "—"])

    # ---- step 5: text standardisation -------------------------------------------
    n_text_fixed = 0
    shown_text = 0
    for col in ["Product", "Category", "Region"]:
        before = clean[col].astype(str)
        clean[col] = clean[col].map(_canon)
        changed = before != clean[col]
        n_text_fixed += int(changed.sum())
        for idx in clean.index[changed]:
            if shown_text >= 3:
                break
            oid = clean.loc[idx, "Order_ID"]
            examples.append((oid, f"{col} {before.loc[idx]!r} → {clean.loc[idx, col]!r}"))
            shown_text += 1
    log_rows.append(
        [5, "Inconsistent whitespace / capitalisation in Product, Category, Region",
         "Trimmed whitespace; canonical casing via lookup table",
         f"{n_text_fixed} cell(s)",
         "Variants like 'EAST' vs 'East ' would split every group-by and chart; "
         "canonical labels keep aggregations correct.",
         "e.g. '  LAPTOP ' → 'Laptop'"])

    # ---- step 6: data types ------------------------------------------------------
    clean["Unit_Price"] = pd.to_numeric(clean["Unit_Price"], errors="coerce").round(2)
    log_rows.append(
        [6, "Quantity / Unit_Price stored as text",
         "Quantity → integer; Unit_Price → decimal (2 dp)",
         "All rows",
         "Correct types prevent silent errors (text sorts '12' < '9') and "
         "enable arithmetic.",
         "Quantity: int64, Unit_Price: float64"])

    # ---- step 7: derived Revenue --------------------------------------------------
    clean["Revenue"] = (clean["Quantity"] * clean["Unit_Price"]).round(2)
    log_rows.append(
        [7, "No revenue figure in raw data",
         "Revenue = Quantity × Unit_Price, added AFTER cleaning",
         "All rows",
         "Revenue built on cleaned values; one bad cell would otherwise poison "
         "every total. Revenue is sales value, not profit.",
         "e.g. 12 × ₹19,481.45 = ₹233,777.40"])

    quality = {
        "rows_before": len(raw),
        "rows_after": len(clean),
        "duplicates_removed": n_dup,
        "duplicate_order_ids": dup_ids,
        "conflicting_order_ids": conflicts,
        "missing_before": {c: int(raw[c].isna().sum()) for c in raw.columns},
        "missing_after": {c: int(clean[c].isna().sum()) for c in clean.columns},
        "dates_standardised": n_non_iso,
        "text_cells_standardised": n_text_fixed,
        "quantity_median_used": qty_median,
    }
    return clean, log_rows, quality, examples


# ------------------------------------------------------------- validation
# Status levels: PASS (requirement met), WARNING (acceptable but noteworthy —
# never presented as a pass), FAIL (must be fixed; blocks all outputs).
def validate_cleaned(clean: pd.DataFrame, quality: dict):
    checks = []

    def add(name, expected, observed, status):
        checks.append((name, expected, observed, status))

    n_missing = int(clean.isna().sum().sum())
    add("Missing values (all columns)", "0", str(n_missing),
        "PASS" if n_missing == 0 else "FAIL")

    n_dup = int(clean.duplicated().sum())
    add("Exact duplicate rows", "0", str(n_dup),
        "PASS" if n_dup == 0 else "FAIL")

    n_ids = clean["Order_ID"].nunique()
    add("Unique Order_IDs", f"{len(clean)} of {len(clean)}",
        f"{n_ids} of {len(clean)}",
        "PASS" if n_ids == len(clean) else "FAIL")

    n_conf = len(quality["conflicting_order_ids"])
    add("Conflicting duplicate Order_IDs (same ID, different data)", "0",
        str(n_conf) + (f": {', '.join(quality['conflicting_order_ids'])}" if n_conf else ""),
        "PASS" if n_conf == 0 else "WARNING")

    bad_qty = clean[clean["Quantity"] <= 0]
    add("Invalid quantities (≤ 0)", "0", str(len(bad_qty)),
        "PASS" if len(bad_qty) == 0 else "FAIL")

    bad_price = clean[(clean["Unit_Price"] <= 0) | (clean["Unit_Price"].isna())]
    add("Invalid unit prices (≤ 0 or missing)", "0", str(len(bad_price)),
        "PASS" if len(bad_price) == 0 else "FAIL")

    bad_dates = clean[pd.to_datetime(clean["Order_Date"], format="%Y-%m-%d",
                                     errors="coerce").isna()]
    add("Invalid / unparseable dates", "0", str(len(bad_dates)),
        "PASS" if len(bad_dates) == 0 else "FAIL")

    rev_ok = bool(np.allclose(clean["Revenue"],
                              (clean["Quantity"] * clean["Unit_Price"]).round(2)))
    add("Revenue = Quantity × Unit_Price (every row)", "consistent",
        "consistent" if rev_ok else "MISMATCH",
        "PASS" if rev_ok else "FAIL")

    rows_ok = (quality["rows_after"]
               == quality["rows_before"] - quality["duplicates_removed"])
    add("Row reconciliation (before − duplicates = after)",
        f"{quality['rows_before']} − {quality['duplicates_removed']} = {quality['rows_after']}",
        str(quality["rows_after"]), "PASS" if rows_ok else "FAIL")

    n_unknown_cat = int((clean["Category"] == "Unknown").sum())
    n_unknown_reg = int((clean["Region"] == "Unknown").sum())
    add("Placeholder 'Unknown' values (documented, not invented)",
        "0 ideally", f"Category: {n_unknown_cat}, Region: {n_unknown_reg}",
        "PASS" if (n_unknown_cat + n_unknown_reg) == 0 else "WARNING")

    return checks


# ------------------------------------------------------- business analysis
def business_analysis(clean: pd.DataFrame):
    """Small, honest analysis on the cleaned data. No predictions, no profit
    confusion: Revenue here is sales value (Quantity × Unit_Price)."""
    total_revenue = float(clean["Revenue"].sum())
    n_orders = len(clean)
    aov = total_revenue / n_orders

    by_region = (clean.groupby("Region", as_index=False)
                 .agg(orders=("Order_ID", "count"), revenue=("Revenue", "sum"))
                 .sort_values("revenue", ascending=False))
    by_region["share_pct"] = (by_region["revenue"] / total_revenue * 100).round(1)

    by_product = (clean.groupby("Product", as_index=False)
                  .agg(units=("Quantity", "sum"), revenue=("Revenue", "sum"))
                  .sort_values("revenue", ascending=False))

    top_region = by_region.iloc[0]
    top_product = by_product.iloc[0]
    top_units = by_product.sort_values("units", ascending=False).iloc[0]
    interpretation = [
        f"Total sales value is ₹{total_revenue:,.2f} across {n_orders} orders "
        f"(average order value ₹{aov:,.2f}). This is revenue, not profit — "
        f"costs are not in the dataset.",
        f"{top_region['Region']} is the top region at ₹{top_region['revenue']:,.2f} "
        f"({top_region['share_pct']:.1f}% of revenue), ahead of "
        f"{by_region.iloc[1]['Region']} (₹{by_region.iloc[1]['revenue']:,.2f}).",
        f"{top_product['Product']} leads by revenue (₹{top_product['revenue']:,.2f}) "
        f"on high unit prices; {top_units['Product']} leads by volume "
        f"({int(top_units['units'])} units). High price × modest volume beats "
        f"low price × high volume here.",
        "Data-quality note: cleaning removed 2 duplicate orders and resolved "
        "3 missing values, so these totals rest on 30 verified unique orders.",
    ]
    return {
        "total_revenue": total_revenue, "n_orders": n_orders, "aov": aov,
        "n_products": int(clean["Product"].nunique()),
        "n_regions": int(clean["Region"].nunique()),
        "by_region": by_region,
        "by_product": by_product,
        "top_by_revenue": by_product.head(5),
        "top_by_units": by_product.sort_values("units", ascending=False).head(5),
        "interpretation": interpretation,
    }

# ------------------------------------------------------- workbook styling
NAVY, GOLD, BAND, GRID = "1F3B63", "D4A017", "DCE6F1", "B4C6E7"
GREEN, RED, AMBER = "2E7D32", "C00000", "E65100"
thin = Side(style="thin", color=GRID)
border = Border(left=thin, right=thin, top=thin, bottom=thin)
hdr_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
hdr_fill = PatternFill("solid", fgColor=NAVY)
hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
band_fill = PatternFill("solid", fgColor=BAND)
pass_fill = PatternFill("solid", fgColor="C6EFCE")
warn_fill = PatternFill("solid", fgColor="FFEB9C")
fail_fill = PatternFill("solid", fgColor="FFC7CE")
title_font = Font(name="Calibri", size=20, bold=True, color=NAVY)
sec_font = Font(name="Calibri", size=13, bold=True, color=NAVY)
sub_font = Font(name="Calibri", size=12, bold=True, color=NAVY)
body_font = Font(name="Calibri", size=11, color="1A1A1A")
INR = '"Rs."#,##0.00'
STATUS_FILL = {"PASS": pass_fill, "WARNING": warn_fill, "FAIL": fail_fill}
STATUS_FONT = {"PASS": Font(color="006100", bold=True),
               "WARNING": Font(color="9C6500", bold=True),
               "FAIL": Font(color="9C0006", bold=True)}


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font, cell.fill, cell.alignment, cell.border = \
            hdr_font, hdr_fill, hdr_align, border


def style_body(ws, first_row, last_row, ncols):
    for r in range(first_row, last_row + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = body_font
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if r % 2 == 0:
                cell.fill = band_fill


def write_df(ws, df, display_headers=None, start_row=1, money_cols=()):
    cols = list(df.columns)
    labels = [display_headers.get(c, c) if display_headers else c for c in cols]
    for j, lab in enumerate(labels, 1):
        ws.cell(row=start_row, column=j, value=lab)
    for i, (_, row) in enumerate(df.iterrows(), start_row + 1):
        for j, col in enumerate(cols, 1):
            v = row[col]
            cell = ws.cell(row=i, column=j, value=None if pd.isna(v) else v)
            if col in money_cols:
                cell.number_format = INR
    style_header(ws, start_row, len(cols))
    style_body(ws, start_row + 1, start_row + len(df), len(cols))
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
    ws.auto_filter.ref = (f"A{start_row}:{get_column_letter(len(cols))}"
                          f"{start_row + len(df)}")
    return start_row + len(df)


def add_bar_chart(ws, title, data_ref, cats_ref, anchor, y_title="Revenue (Rs.)"):
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = title
    chart.y_axis.title = y_title
    chart.x_axis.title = None
    chart.height, chart.width = 7.5, 13
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.dataLabels = None
    ws.add_chart(chart, anchor)


def build_workbook(raw, clean, log_rows, quality, checks, analysis):
    wb = Workbook()

    # ------------------------------------------------- 1. README_Summary
    ws = wb.active
    ws.title = "README_Summary"
    ws.sheet_properties.tabColor = NAVY
    for col, w in zip("ABCDEF", (26, 22, 22, 22, 22, 22)):
        ws.column_dimensions[col].width = w
    ws.merge_cells("A1:F2")
    c = ws["A1"]
    c.value = "Data Cleaning and Business Analysis using Python and Excel"
    c.font, c.alignment = title_font, Alignment(horizontal="center", vertical="center")
    ws.merge_cells("A3:F3")
    c = ws["A3"]
    c.value = "CodeOrbit Tech · Data Analyst Internship (1-Month Track) · Task 1 · Intern: Sanjana"
    c.font, c.alignment = sub_font, Alignment(horizontal="center", vertical="center")

    r = 5
    ws[f"A{r}"] = "Project overview"
    ws[f"A{r}"].font = sec_font
    r += 1
    overview = [
        "A messy 32-row sales-orders CSV (synthetic data, seed 20261010) was cleaned "
        "with a reproducible Python/pandas pipeline into a 30-row analysis-ready dataset.",
        "Every cleaning decision is documented with its reasoning in Cleaning_Log; "
        "uncertain values were flagged ('Unknown'), never invented.",
        "The raw file is preserved untouched; validation runs automatically and "
        "blocks all outputs on any FAIL.",
    ]
    for line in overview:
        ws.cell(row=r, column=1, value="• " + line).font = body_font
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        r += 1
    r += 1
    ws[f"A{r}"] = "Methodology (7 steps)"
    ws[f"A{r}"].font = sec_font
    r += 1
    for row in log_rows[1:]:
        ws.cell(row=r, column=1, value=f"Step {row[0]}: {row[1]}").font = body_font
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        r += 1
        ws.cell(row=r, column=1, value=f"→ {row[2]} ({row[4]})").font = body_font
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        r += 1
    r += 1
    ws[f"A{r}"] = "Key results (verified by the pipeline)"
    ws[f"A{r}"].font = sec_font
    r += 1
    a = analysis
    results = [
        f"Rows: {quality['rows_before']} → {quality['rows_after']} "
        f"({quality['duplicates_removed']} duplicates removed, 3 missing values resolved)",
        f"Validation: {sum(1 for c in checks if c[3] == 'PASS')} PASS, "
        f"{sum(1 for c in checks if c[3] == 'WARNING')} WARNING, "
        f"{sum(1 for c in checks if c[3] == 'FAIL')} FAIL",
        f"Total revenue: Rs.{a['total_revenue']:,.2f} across {a['n_orders']} orders "
        f"(average order value Rs.{a['aov']:,.2f})",
        f"Top region: {a['by_region'].iloc[0]['Region']} "
        f"(Rs.{a['by_region'].iloc[0]['revenue']:,.2f}); "
        f"top product: {a['top_by_revenue'].iloc[0]['Product']} "
        f"(Rs.{a['top_by_revenue'].iloc[0]['revenue']:,.2f})",
    ]
    for line in results:
        ws.cell(row=r, column=1, value="• " + line).font = body_font
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        r += 1
    r += 1
    ws[f"A{r}"] = "How to reproduce"
    ws[f"A{r}"].font = sec_font
    r += 1
    for line in ["1. pip install -r requirements.txt",
                 "2. python build_task1.py   (reads raw_dataset.csv, writes cleaned_dataset.csv + this workbook)",
                 "3. python tests/test_cleaning.py   (independent test suite)"]:
        ws.cell(row=r, column=1, value=line).font = body_font
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        r += 1

    # ------------------------------------------------- 2. Raw_Data
    ws = wb.create_sheet("Raw_Data")
    ws.sheet_properties.tabColor = RED
    write_df(ws, raw, DISPLAY_HEADERS)
    for i, w in enumerate([12, 14, 20, 15, 10, 14, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 30

    # ------------------------------------------------- 3. Cleaned_Data
    ws = wb.create_sheet("Cleaned_Data")
    ws.sheet_properties.tabColor = GREEN
    write_df(ws, clean, DISPLAY_HEADERS, money_cols=("Unit_Price", "Revenue"))
    for i, w in enumerate([12, 14, 20, 15, 10, 14, 14, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 30
    for cell in ws[1]:
        if cell.value == DISPLAY_HEADERS["Revenue"]:
            cell.fill = PatternFill("solid", fgColor=GREEN)

    # ------------------------------------------------- 4. Cleaning_Log
    ws = wb.create_sheet("Cleaning_Log")
    ws.sheet_properties.tabColor = GOLD
    for i, row in enumerate(log_rows, 1):
        for j, v in enumerate(row, 1):
            ws.cell(row=i, column=j, value=v)
    style_header(ws, 1, 6)
    style_body(ws, 2, len(log_rows), 6)
    ws.row_dimensions[1].height = 30
    for i, w in enumerate([8, 34, 40, 22, 44, 34], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:F{len(log_rows)}"

    # ------------------------------------------------- 5. Data_Quality
    ws = wb.create_sheet("Data_Quality")
    ws.sheet_properties.tabColor = "00897B"
    ws["A1"] = "Data-quality: before → after"
    ws["A1"].font = sec_font
    q = quality
    dq = [
        ["Metric", "Before (raw)", "After (cleaned)", "Change"],
        ["Rows", q["rows_before"], q["rows_after"],
         f"-{q['duplicates_removed']} duplicates"],
        ["Exact duplicate rows", q["duplicates_removed"], 0, "removed"],
        ["Missing values (total)", sum(q["missing_before"].values()),
         sum(q["missing_after"].values()), "resolved"],
        ["Missing Category", q["missing_before"].get("Category", 0),
         q["missing_after"].get("Category", 0), "→ 'Unknown'"],
        ["Missing Quantity", q["missing_before"].get("Quantity", 0),
         q["missing_after"].get("Quantity", 0), "→ median"],
        ["Missing Region", q["missing_before"].get("Region", 0),
         q["missing_after"].get("Region", 0), "→ 'Unknown'"],
        ["Dates in non-ISO formats", q["dates_standardised"], 0, "standardised"],
        ["Text cells standardised", "—", q["text_cells_standardised"], "trimmed + recased"],
        ["Conflicting Order_IDs", len(q["conflicting_order_ids"]), 0,
         "none found" if not q["conflicting_order_ids"] else "FLAGGED"],
    ]
    for i, row in enumerate(dq, 2):
        for j, v in enumerate(row, 1):
            ws.cell(row=i, column=j, value=v)
    style_header(ws, 2, 4)
    style_body(ws, 3, 2 + len(dq) - 1, 4)
    for i, w in enumerate([32, 16, 18, 24], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A3"

    # ------------------------------------------------- 6. Revenue_by_Region
    ws = wb.create_sheet("Revenue_by_Region")
    ws.sheet_properties.tabColor = "EF6C00"
    ws["A1"] = "Revenue by region (from cleaned data)"
    ws["A1"].font = sec_font
    reg = a["by_region"][["Region", "orders", "revenue", "share_pct"]].copy()
    reg.columns = ["Region", "Orders", "Revenue (Rs.)", "Share %"]
    last = write_df(ws, reg, start_row=3, money_cols=("Revenue (Rs.)",))
    for i, w in enumerate([16, 12, 18, 12], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    add_bar_chart(
        ws, "Revenue by Region (Rs.)",
        Reference(ws, min_col=3, min_row=3, max_row=2 + len(reg)),
        Reference(ws, min_col=1, min_row=4, max_row=3 + len(reg)),
        f"A{last + 2}")

    # ------------------------------------------------- 7. Revenue_by_Product
    ws = wb.create_sheet("Revenue_by_Product")
    ws.sheet_properties.tabColor = "6A1B9A"
    ws["A1"] = "Top 5 products by revenue (from cleaned data)"
    ws["A1"].font = sec_font
    prod = a["top_by_revenue"][["Product", "units", "revenue"]].copy()
    prod.columns = ["Product", "Units sold", "Revenue (Rs.)"]
    last = write_df(ws, prod, start_row=3, money_cols=("Revenue (Rs.)",))
    for i, w in enumerate([24, 14, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    add_bar_chart(
        ws, "Top 5 Products by Revenue (Rs.)",
        Reference(ws, min_col=3, min_row=3, max_row=2 + len(prod)),
        Reference(ws, min_col=1, min_row=4, max_row=3 + len(prod)),
        f"A{last + 2}")

    # ------------------------------------------------- 8. Validation_Summary
    ws = wb.create_sheet("Validation_Summary")
    ws.sheet_properties.tabColor = "1565C0"
    ws["A1"] = "Automated validation (cleaned dataset)"
    ws["A1"].font = sec_font
    hdr = ["Check", "Expected", "Observed", "Status"]
    for j, v in enumerate(hdr, 1):
        ws.cell(row=3, column=j, value=v)
    style_header(ws, 3, 4)
    r = 4
    for name, expected, observed, status in checks:
        ws.cell(row=r, column=1, value=name).font = body_font
        ws.cell(row=r, column=2, value=expected).font = body_font
        ws.cell(row=r, column=3, value=observed).font = body_font
        sc = ws.cell(row=r, column=4, value=status)
        sc.font = STATUS_FONT[status]
        sc.fill = STATUS_FILL[status]
        sc.alignment = Alignment(horizontal="center", vertical="center")
        for cc in range(1, 5):
            ws.cell(row=r, column=cc).border = border
        r += 1
    # Real conditional formatting as well (stays live if values are edited)
    rng = f"D4:D{r - 1}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"FAIL"'],
                                                 fill=fail_fill, font=STATUS_FONT["FAIL"]))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"WARNING"'],
                                                 fill=warn_fill, font=STATUS_FONT["WARNING"]))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"PASS"'],
                                                 fill=pass_fill, font=STATUS_FONT["PASS"]))
    for i, w in enumerate([52, 26, 30, 12], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:D{r - 1}"
    r += 1
    ws[f"A{r}"] = ("Note: WARNING means 'acceptable but noteworthy' — e.g. documented "
                   "'Unknown' placeholders. It is never presented as a pass. FAIL "
                   "blocks all outputs until fixed.")
    ws[f"A{r}"].font = Font(name="Calibri", size=10, italic=True, color="404040")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    return wb


# ------------------------------------------------------------------- main
def main() -> int:
    print("=" * 64)
    print("CodeOrbit Tech · Task 1 — Data Cleaning pipeline")
    print("=" * 64)
    try:
        raw = load_raw()          # read-only: the raw CSV is never written
    except (FileNotFoundError, ValueError) as exc:
        print(f"\nERROR: {exc}")
        return 2

    profile = profile_raw(raw)
    print_profile(profile)

    print("\nCleaning…")
    clean, log_rows, quality, examples = clean_data(raw)
    for step, issue, action, affected, reason, ex in log_rows[1:]:
        print(f"  Step {step}: {issue}\n         → {action} [{affected}]")

    print("\nValidation checks:")
    checks = validate_cleaned(clean, quality)
    n_fail = n_warn = 0
    for name, expected, observed, status in checks:
        print(f"  [{status:7}] {name}: observed={observed} (expected {expected})")
        n_fail += status == "FAIL"
        n_warn += status == "WARNING"
    if n_fail:
        print(f"\nVALIDATION FAILED ({n_fail} FAIL) — outputs were NOT written. "
              "Fix the data or code first.")
        return 1
    if n_warn:
        print(f"\n{n_warn} warning(s) noted above — documented and acceptable, "
              "continuing.")

    analysis = business_analysis(clean)

    clean.to_csv(CLEAN_CSV, index=False)     # cleaned output only
    wb = build_workbook(raw, clean, log_rows, quality, checks, analysis)
    wb.save(XLSX)

    print(f"\nWrote: {CLEAN_CSV.name} ({len(clean)} rows × {len(clean.columns)} cols)")
    print(f"Wrote: {XLSX.name} ({len(wb.sheetnames)} sheets)")
    print(f"Rows: {quality['rows_before']} → {quality['rows_after']} | "
          f"Total revenue: Rs.{analysis['total_revenue']:,.2f} | "
          f"AOV: Rs.{analysis['aov']:,.2f}")
    print("Done.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
