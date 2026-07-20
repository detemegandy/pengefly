"""
PROTOTYPE — budget vs actual visualization for the monthly transaction sheet.

Answers issue #8: where should the budget vs actual visual comparison live?

v3: In-cell SPARKLINE progress bars (one per category row in the budget table).
    Replaces the raw "%" column with a horizontal bar that fills proportionally
    to spent/budget — blue when on track, red when over.

    This is how YNAB and Sbanken show it: no separate chart, just each row
    has its own bar so you read the table and the picture at the same time.

Run: uv run prototypes/budget_chart_prototype.py
THROWAWAY — do not merge to main.
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import transaction_entry as te

CREDS_FILE = te.CREDS_FILE
SCOPES     = te.SCOPES
SHEET_ID   = te.SHEET_ID

# Budget table layout (1-indexed):
#   row 3  = header: Category | Budget | Spent | Remaining | % | ⚠ Open
#   rows 4-10 = 7 category data rows
_HDR_ROW  = 3
_DATA_ROW = _HDR_ROW + 1


def main():
    creds   = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds)

    info = service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
    sid  = None
    for s in info["sheets"]:
        if s["properties"]["title"] == te.TAB_NAME:
            sid = s["properties"]["sheetId"]
            existing_charts = s.get("charts", [])
            break

    if sid is None:
        print(f"Tab '{te.TAB_NAME}' not found — run run_all.py first.")
        return

    # Remove any chart left over from earlier prototype iterations
    if existing_charts:
        dels = [{"deleteEmbeddedObject": {"objectId": c["chartId"]}}
                for c in existing_charts]
        service.spreadsheets().batchUpdate(
            spreadsheetId=SHEET_ID, body={"requests": dels}).execute()
        print(f"  Removed {len(dels)} existing chart(s)")

    n_cats = len(te.CATEGORIES)
    end_row = _DATA_ROW + n_cats - 1

    # Column E holds the raw "%" value written by transaction_entry.py.
    # Replace it with a SPARKLINE formula:
    #   - fills bar from 0 → 1 using C/B (spent / budget)
    #   - capped at 1 so an over-budget bar just turns solid red, not wider
    #   - blue (#4472C4) when under budget, red (#E53935) when over
    #   - light fill (#E8F0FE) for remaining budget portion
    sparklines = []
    for i in range(n_cats):
        row = _DATA_ROW + i
        formula = (
            f'=SPARKLINE(MIN(C{row}/B{row},1),'
            f'{{"charttype","bar";'
            f'"max",1;'
            f'"color1",IF(C{row}>B{row},"#E53935","#4472C4");'
            f'"color2","#E8F0FE"}})'
        )
        sparklines.append([formula])

    tab = te.TAB_NAME
    service.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range=f"'{tab}'!E{_DATA_ROW}:E{end_row}",
        valueInputOption="USER_ENTERED",
        body={"values": sparklines},
    ).execute()

    n = n_cats
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit"
    print(f"Written {n} SPARKLINE progress bars to '{te.TAB_NAME}'!E{_DATA_ROW}:E{end_row}")
    print()
    print("Column E now shows an in-cell progress bar per category:")
    print("  Blue fill  = proportion of budget spent (under budget)")
    print("  Red fill   = proportion of budget spent (over budget)")
    print("  Light fill = remaining budget headroom")
    print(f"Open: {url}")


if __name__ == "__main__":
    main()
