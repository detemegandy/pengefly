"""
Run all four prototype tabs into the same workbook so you can browse them together.

Tab order after run:
  1. Yearly 2026          — annual budget overview; status dropdowns (✓ Closed / Open / —)
  2. Config               — accounts, categories, card import mapping
  3. Card Transactions … — Jun 2026 transaction detail + CC min payment rows
  4. Credit Card Tracker  — EIKA and Mona payment schedules + over-limit simulation

Each prototype manages its own tab and preserves the others, so they can also be
run independently when iterating on a single sheet.

Run: uv run prototypes/run_all.py
THROWAWAY — do not merge to main.
"""

import subprocess
import sys
import os

SHEET_URL = "https://docs.google.com/spreadsheets/d/1gE5wc6lXaAOsiQ_dpjg7uGBnUrtGr0-oy-56UKyTdQc/edit"
SCRIPTS   = [
    "prototypes/yearly_overview.py",       # must run first — CC tracker reads its status row
    "prototypes/config_sheet.py",
    "prototypes/transaction_entry.py",
    "prototypes/credit_card_tracker.py",
]

def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    for script in SCRIPTS:
        print(f"\n{'='*60}")
        print(f"Running {script}...")
        print('='*60)
        result = subprocess.run(
            ["uv", "run", script],
            cwd=root,
        )
        if result.returncode != 0:
            print(f"\nERROR in {script} — stopping.")
            sys.exit(1)

    print(f"\n{'='*60}")
    print("All prototypes built. Open the spreadsheet:")
    print(f"  {SHEET_URL}")
    print('='*60)


if __name__ == "__main__":
    main()
