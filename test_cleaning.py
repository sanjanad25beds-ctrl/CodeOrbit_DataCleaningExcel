#!/usr/bin/env python3
"""Independent test suite for the Task 1 cleaning pipeline.

Runs with plain Python (no pytest required):

    python tests/test_cleaning.py        # from the project folder

or with pytest:

    pytest tests/ -v
"""
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import build_task1 as b

PASS, FAIL = "PASS", "FAIL"
results = []


def check(name, condition, detail=""):
    results.append((name, PASS if condition else FAIL, detail))


def main():
    raw = b.load_raw()
    clean, log_rows, quality, _ = b.clean_data(raw)
    checks = b.validate_cleaned(clean, quality)

    # --- raw dataset is preserved -------------------------------------------
    check("raw has 32 rows", len(raw) == 32, f"got {len(raw)}")
    check("raw still has missing values",
          int(raw.isna().sum().sum()) == 3, "raw must stay messy")
    check("raw still has duplicates",
          int(raw.duplicated(keep='first').sum()) == 2)

    # --- cleaning outcomes ----------------------------------------------------
    check("cleaned has 30 rows", len(clean) == 30, f"got {len(clean)}")
    check("no missing values after cleaning",
          int(clean.isna().sum().sum()) == 0)
    check("no duplicate rows after cleaning",
          int(clean.duplicated().sum()) == 0)
    check("Order_IDs unique", clean["Order_ID"].nunique() == len(clean))
    check("no conflicting Order_IDs",
          quality["conflicting_order_ids"] == [],
          str(quality["conflicting_order_ids"]))
    check("dates are ISO YYYY-MM-DD",
          clean["Order_Date"].str.match(r"^\d{4}-\d{2}-\d{2}$").all())
    check("Quantity is integer and positive",
          str(clean["Quantity"].dtype).startswith("int")
          and bool((clean["Quantity"] > 0).all()))
    check("Unit_Price positive", bool((clean["Unit_Price"] > 0).all()))
    check("Revenue = Quantity x Unit_Price",
          bool(((clean["Quantity"] * clean["Unit_Price"]).round(2)
                == clean["Revenue"]).all()))
    check("row reconciliation",
          quality["rows_after"]
          == quality["rows_before"] - quality["duplicates_removed"])

    # --- validation report agrees ----------------------------------------------
    statuses = [c[3] for c in checks]
    check("no FAIL in validation", "FAIL" not in statuses,
          f"statuses={statuses}")
    check("cleaning log documents every step", len(log_rows) - 1 == 9,
          f"{len(log_rows) - 1} steps logged (0-7)")

    # --- business analysis sanity -----------------------------------------------
    a = b.business_analysis(clean)
    check("total revenue positive", a["total_revenue"] > 0,
          f"Rs.{a['total_revenue']:,.2f}")
    check("region shares sum to ~100%",
          abs(a["by_region"]["share_pct"].sum() - 100) < 0.2)

    print("\nTest results:")
    failed = 0
    for name, status, detail in results:
        print(f"  [{status}] {name}" + (f" — {detail}" if detail else ""))
        failed += status == FAIL
    print(f"\n{len(results) - failed}/{len(results)} tests passed.")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
