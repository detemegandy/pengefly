"""
Run all four prototype tabs into the same workbook so you can browse them together.

Each script builds its own tab with static/self-referencing values — it does NOT write
any formula that references another tab. After all tabs exist, this script makes one
final API call to patch all cross-sheet formula references.

To add a new cross-sheet reference: append an entry to _build_cross_refs() below.

Tab order after run:
  1. Yearly 2026          — annual budget overview; status dropdowns (✓ Closed / Open / —)
  2. Config               — accounts, categories, card import mapping
  3. Card Transactions … — Jun 2026 transaction detail + CC min payment rows
  4. Credit Card Tracker  — EIKA and Mona payment schedules + over-limit simulation

Run: uv run prototypes/run_all.py
THROWAWAY — do not merge to main.
"""

import subprocess
import sys
import os

# Insert prototypes/ onto path so we can import module-level constants for cross-refs
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import credit_card_tracker as cct
import transaction_entry as te

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

CREDS_FILE = cct.CREDS_FILE
SCOPES     = cct.SCOPES
SHEET_ID   = cct.SHEET_ID
SHEET_URL  = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit"

SCRIPTS = [
    "prototypes/yearly_overview.py",
    "prototypes/config_sheet.py",
    "prototypes/transaction_entry.py",
    "prototypes/credit_card_tracker.py",
]


# ── Cross-sheet formula references ────────────────────────────────────────────
# Add entries here when a cell in one tab must reference a cell in another tab.
# Format: {"range": "'Tab Name'!A1", "values": [["=formula"]]}

def _build_cross_refs():
    # CC min-payment row numbers depend on how many regular transactions are above them
    regular_count = len([r for r in te.SAMPLES if "CC min" not in r[2]])
    eika_row = te.DATA_ROW + regular_count      # 1-indexed sheet row
    mona_row = te.DATA_ROW + regular_count + 1

    return [
        # Credit Card Tracker: "Closed through" reads Yearly 2026 status row (D3:O3)
        {"range": f"'{cct.TAB_NAME}'!B4", "values": [[cct.CLOSED_THROUGH_FX]]},
        {"range": f"'{cct.TAB_NAME}'!H4", "values": [[cct.CLOSED_THROUGH_FX]]},
        # Card Transactions: CC min-payment amounts read CC Tracker balance/rate cells
        {"range": f"'{te.TAB_NAME}'!D{eika_row}", "values": [[te._EIKA_MIN]]},
        {"range": f"'{te.TAB_NAME}'!D{mona_row}", "values": [[te._MONA_MIN]]},
    ]


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    for script in SCRIPTS:
        print(f"\n{'='*60}")
        print(f"Running {script}...")
        print('='*60)
        result = subprocess.run(["uv", "run", script], cwd=root)
        if result.returncode != 0:
            print(f"\nERROR in {script} — stopping.")
            sys.exit(1)

    print(f"\n{'='*60}")
    print("Patching cross-sheet formula references...")
    print('='*60)
    creds   = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds)
    cross   = _build_cross_refs()
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={"valueInputOption": "USER_ENTERED", "data": cross},
    ).execute()
    for ref in cross:
        print(f"  ✓  {ref['range']}")

    print(f"\n{'='*60}")
    print("All prototypes built. Open the spreadsheet:")
    print(f"  {SHEET_URL}")
    print('='*60)


if __name__ == "__main__":
    main()
