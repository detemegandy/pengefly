"""
PROTOTYPE — budget vs actual chart for the monthly transaction sheet.

Answers issue #8: where should the budget vs actual visual comparison live?

This prototype places a grouped column chart (Budget blue, Spent amber) directly
on the Card Transactions sheet, anchored to the right of the transaction data area.
The chart references the live SUMPRODUCT formulas in the budget summary (rows 4-10),
so it updates automatically as transactions are added or removed.

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

    # Remove any charts already on this sheet (idempotent rerun)
    if existing_charts:
        dels = [{"deleteEmbeddedObject": {"objectId": c["chartId"]}}
                for c in existing_charts]
        service.spreadsheets().batchUpdate(
            spreadsheetId=SHEET_ID, body={"requests": dels}).execute()
        print(f"  Removed {len(dels)} existing chart(s)")

    # Budget summary layout (0-indexed rows):
    #   row 2 = header:  Category | Budget | Spent | Remaining | % | Open
    #   rows 3-9 = 7 categories
    hdr_row  = 2
    data_end = hdr_row + 1 + len(te.CATEGORIES)   # exclusive

    def src(col_s, col_e):
        return {"sourceRange": {"sources": [{
            "sheetId": sid,
            "startRowIndex": hdr_row,
            "endRowIndex": data_end,
            "startColumnIndex": col_s,
            "endColumnIndex": col_e,
        }]}}

    BLUE  = {"red": 68/255,  "green": 114/255, "blue": 196/255}
    AMBER = {"red": 255/255, "green": 192/255, "blue": 0.0}

    req = {"addChart": {"chart": {
        "spec": {
            "title": "Budget vs Actual — Jun 2026",
            "basicChart": {
                "chartType": "COLUMN",
                "legendPosition": "BOTTOM_LEGEND",
                "axis": [
                    {"position": "BOTTOM_AXIS", "title": "Category"},
                    {"position": "LEFT_AXIS",   "title": "NOK"},
                ],
                "domains": [{"domain": src(0, 1)}],
                "series": [
                    {"series": src(1, 2), "targetAxis": "LEFT_AXIS",
                     "colorStyle": {"rgbColor": BLUE}},
                    {"series": src(2, 3), "targetAxis": "LEFT_AXIS",
                     "colorStyle": {"rgbColor": AMBER}},
                ],
                "headerCount": 1,
            },
        },
        "position": {"overlayPosition": {
            # Anchor just right of the H-column transaction data, at the budget header row
            "anchorCell": {"sheetId": sid, "rowIndex": 2, "columnIndex": 9},
            "widthPixels":  480,
            "heightPixels": 300,
        }},
    }}}

    service.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID, body={"requests": [req]}).execute()

    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit"
    print(f"Chart added to '{te.TAB_NAME}'")
    print(f"Open: {url}")
    print()
    print("The chart appears to the RIGHT of the transaction columns (col J area),")
    print("aligned with the budget summary rows. Blue = Budget, Amber = Spent.")
    print("It live-updates as you add transactions (SUMPRODUCT formulas drive it).")


if __name__ == "__main__":
    main()
