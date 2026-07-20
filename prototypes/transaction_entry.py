"""
PROTOTYPE — transaction entry sheet layout v6

Architecture:
  Card Transactions [Month] — Detailed = single editable master sheet (all transactions from all cards)
  Filter views = one per card, created via API — same data, focused per-card view

Import flow (what the CLI will do):
  1. User exports from bank → pastes into a temp "Import" sheet
  2. CLI reads it, maps columns, appends rows to A with Card column pre-filled
  3. User opens per-card filter view to annotate (Category, Settlement, Notes)

No formula sheets. Everything in A is editable.

Run: uv run prototypes/transaction_entry.py
THROWAWAY — do not merge to main.
"""

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

CREDS_FILE = "/Users/andreas.saltveit/.config/gcp/google-sheets-sa.json"
SCOPES     = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
SHEET_ID   = "1gE5wc6lXaAOsiQ_dpjg7uGBnUrtGr0-oy-56UKyTdQc"

CATEGORIES = ["Groceries", "Entertainment", "Shopping", "Other", "Media", "Liana", "Pets"]
MONTH_OPTS = ["This month", "Previous month", "Next month"]
CARDS      = ["Shared Card", "Shared Credit", "Personal"]
SETTLEMENT = ["Needs discussion",
              "Settled — keep personal",
              "Settled — from General savings",
              "Settled — from SOS fund",
              "Settled — from Vacation savings",
              "Settled — next month budget",
              "Settled — credit card"]
BUDGETS    = {"Groceries": 13000, "Entertainment": 1000, "Shopping": 3000,
              "Other": 3000, "Media": 288, "Liana": 2000, "Pets": 1500}

# Col indices (0-indexed): Card | Date | Merchant | Amount | Category | Month | Settlement | Notes
CARD_COL, DATE_COL, MERCH_COL, AMT_COL = 0, 1, 2, 3
CAT_COL, MONTH_COL, SETTLE_COL, NOTES_COL = 4, 5, 6, 7
NCOLS = 8

# Sheet layout (1-indexed rows):
#  1  = title
#  2  = instruction
#  3  = budget summary header
#  4-10 = per-category rows (7 categories)
#  11 = total row
#  12 = alert banner
#  13 = transaction header  ← frozen through here
#  14+ = transactions (CLI appends here)
HDR_ROW  = 13   # 1-indexed row of transaction column header
DATA_ROW = 14   # 1-indexed first transaction row
ROWS     = 150  # max transactions rows to reserve for formatting

# All sample data in one flat list — this is what the CLI would produce after importing
# all three card exports and appending them to Card Transactions [Month] — Detailed.
SAMPLES = [
    # card              date    merchant              amount  category       month         settlement            notes
    ("Shared Card",   "01.06", "Extra supermarked",   359,  "Groceries",   "This month", "",                  ""),
    ("Shared Card",   "02.06", "CircleK",             553,  "Groceries",   "This month", "",                  ""),
    ("Shared Card",   "05.06", "Rema 1000",           981,  "Groceries",   "This month", "",                  ""),
    ("Shared Card",   "08.06", "Apotek 1",            237,  "Other",       "This month", "",                  ""),
    ("Shared Card",   "12.06", "Extra supermarked",   969,  "Groceries",   "This month", "",                  ""),
    ("Shared Card",   "14.06", "Lekebutikken",        480,  "Liana",       "This month", "",                  ""),
    ("Shared Card",   "18.06", "Bauhaus",            1678,  "Liana",       "This month", "",                  "Gate for Liana"),
    ("Shared Card",   "20.06", "Extra supermarked",   713,  "Groceries",   "This month", "",                  ""),
    ("Shared Card",   "22.06", "Europris",            293,  "Shopping",    "This month", "",                  ""),
    ("Shared Card",   "25.06", "Extra supermarked",   869,  "Groceries",   "This month", "",                  ""),
    ("Shared Credit", "03.06", "McDonalds",           299,  "Entertainment","This month","",                  ""),
    ("Shared Credit", "07.06", "Bellis restaurant",  1026,  "Entertainment","This month","",                  ""),
    ("Shared Credit", "15.06", "Dyreklinikken",       850,  "Pets",        "This month", "",                  "Annual vet check"),
    ("Shared Credit", "17.06", "Pasha restaurant",    778,  "Entertainment","This month","",                  ""),
    ("Shared Credit", "19.06", "Zouq restaurant",    1189,  "Entertainment","This month","",                  ""),
    ("Shared Credit", "21.06", "Pizza To Go",         299,  "Entertainment","This month","",                  ""),
    ("Shared Credit", "23.06", "Renseriet",           863,  "Shopping",    "This month", "",                  ""),
    ("Personal",      "04.06", "Netflix",             139,  "",            "This month", "",                  ""),
    ("Personal",      "06.06", "Spotify",              99,  "",            "This month", "",                  ""),
    ("Personal",      "09.06", "Dyreklinikken",       450,  "Pets",        "This month", "Needs discussion",  "Shared card was at limit — reimburse from Pets fund"),
    ("Personal",      "16.06", "Extra supermarked",   648,  "Groceries",   "This month", "Needs discussion",  "Forgot shared card at home"),
    ("Personal",      "24.06", "Tool Pool",           249,  "Other",       "This month", "Needs discussion",  "Shared tool rental — booked with personal card"),
]


# ── helpers ────────────────────────────────────────────────────────────────────

def rgb(r, g, b):
    return {"red": r/255, "green": g/255, "blue": b/255}

BLUE_DARK   = rgb(68, 114, 196)
BLUE_LIGHT  = rgb(180, 198, 231)
BLUE_PALE   = rgb(221, 235, 246)
AMBER       = rgb(255, 192, 0)
WHITE       = rgb(255, 255, 255)
YELLOW_SOFT = rgb(255, 217, 102)

def fmt(bg=None, bold=False, fg=None, size=None, halign=None, italic=False):
    f = {}
    if bg:     f["backgroundColor"] = bg
    tf = {}
    if bold:   tf["bold"] = True
    if italic: tf["italic"] = True
    if fg:     tf["foregroundColor"] = fg
    if size:   tf["fontSize"] = size
    if tf:     f["textFormat"] = tf
    if halign: f["horizontalAlignment"] = halign
    return f

def rpt(sid, r, c, nrows, ncols, f):
    return {"repeatCell": {
        "range": {"sheetId": sid, "startRowIndex": r, "endRowIndex": r+nrows,
                  "startColumnIndex": c, "endColumnIndex": c+ncols},
        "cell": {"userEnteredFormat": f},
        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
    }}

def mrg(sid, r1, c1, r2, c2):
    return {"mergeCells": {
        "range": {"sheetId": sid, "startRowIndex": r1, "endRowIndex": r2,
                  "startColumnIndex": c1, "endColumnIndex": c2},
        "mergeType": "MERGE_ALL"}}

def cw(sid, col, px):
    return {"updateDimensionProperties": {
        "range": {"sheetId": sid, "dimension": "COLUMNS",
                  "startIndex": col, "endIndex": col+1},
        "properties": {"pixelSize": px}, "fields": "pixelSize"}}

def frz(sid, rows=0, cols=0):
    return {"updateSheetProperties": {
        "properties": {"sheetId": sid,
                       "gridProperties": {"frozenRowCount": rows, "frozenColumnCount": cols}},
        "fields": "gridProperties(frozenRowCount,frozenColumnCount)"}}

def ddv(sid, r1, r2, col, opts):
    return {"setDataValidation": {
        "range": {"sheetId": sid, "startRowIndex": r1, "endRowIndex": r2,
                  "startColumnIndex": col, "endColumnIndex": col+1},
        "rule": {"condition": {"type": "ONE_OF_LIST",
                               "values": [{"userEnteredValue": o} for o in opts]},
                 "showCustomUi": True, "strict": False}}}

def heat(sid, r1, r2, col):
    return {"addConditionalFormatRule": {"rule": {
        "ranges": [{"sheetId": sid, "startRowIndex": r1, "endRowIndex": r2,
                    "startColumnIndex": col, "endColumnIndex": col+1}],
        "gradientRule": {
            "minpoint": {"color": rgb(198,239,206), "type": "MIN"},
            "midpoint": {"color": rgb(255,235,156), "type": "PERCENTILE", "value": "50"},
            "maxpoint": {"color": rgb(255,199,206), "type": "MAX"},
        }}, "index": 0}}

def row_cfmt(sid, r1, r2, text, color, idx=1):
    return {"addConditionalFormatRule": {"rule": {
        "ranges": [{"sheetId": sid, "startRowIndex": r1, "endRowIndex": r2,
                    "startColumnIndex": 0, "endColumnIndex": NCOLS}],
        "booleanRule": {
            "condition": {"type": "TEXT_CONTAINS",
                          "values": [{"userEnteredValue": text}]},
            "format": {"backgroundColor": color},
        }}, "index": idx}}

def wrap(sid, r1, r2, col):
    return {"repeatCell": {
        "range": {"sheetId": sid, "startRowIndex": r1, "endRowIndex": r2,
                  "startColumnIndex": col, "endColumnIndex": col+1},
        "cell": {"userEnteredFormat": {"wrapStrategy": "WRAP"}},
        "fields": "userEnteredFormat.wrapStrategy"}}


# ── build Card Transactions [Month] — Detailed ─────────────────────────────────────────────────────────

def build_all_cards(sid):
    rq, dt = [], []
    T1 = DATA_ROW
    T2 = DATA_ROW + ROWS - 1

    # Column widths: Card | Date | Merchant | Amount | Category | Month | Settlement | Notes
    for i, px in enumerate([110, 70, 200, 80, 125, 110, 215, 240]):
        rq.append(cw(sid, i, px))

    # Row 0: title
    rq.append(mrg(sid, 0, 0, 1, NCOLS))
    rq.append(rpt(sid, 0, 0, 1, NCOLS, fmt(bg=BLUE_DARK, bold=True, fg=WHITE, size=13, halign="CENTER")))
    dt.append({"range": "A1", "values": [["Card Transactions Jun 2026 — Detailed"]]})

    # Row 1: instruction
    rq.append(mrg(sid, 1, 0, 2, NCOLS))
    rq.append(rpt(sid, 1, 0, 1, NCOLS, fmt(bg=YELLOW_SOFT, italic=True)))
    dt.append({"range": "A2",
               "values": [["  CLI imports each card export here with Card pre-filled. "
                            "To switch to a per-card view: Data → Change view → [card name]. "
                            "Everything here is editable — annotate Category, Settlement, Notes in any view. "
                            "To filter manually: Data → Create a filter."]]}),

    # Row 2: budget summary header
    rq.append(rpt(sid, 2, 0, 1, NCOLS, fmt(bg=BLUE_LIGHT, bold=True, halign="CENTER")))
    dt.append({"range": "A3:H3",
               "values": [["Category", "Budget", "Spent", "Remaining", "%", "⚠ Open", "", ""]]})

    # Rows 3–9: per-category budget rows
    # Col A=Card, D=Amount, E=Category, G=Settlement in transaction rows
    cat_rows = []
    for i, cat in enumerate(CATEGORIES):
        rn = i + 4   # 1-indexed sheet row
        b  = BUDGETS[cat]
        spent = (f'=SUMPRODUCT((E${T1}:E${T2}=A{rn})'
                 f'*(A${T1}:A${T2}<>"Personal")'
                 f'*D${T1}:D${T2})')
        left  = f"=B{rn}-C{rn}"
        pct   = f"=IF(B{rn}>0,C{rn}/B{rn},0)"
        opens = (f'=COUNTIFS(E${T1}:E${T2},A{rn},'
                 f'G${T1}:G${T2},"Needs discussion")')
        cat_rows.append([cat, b, spent, left, pct, opens, "", ""])
        rq.append(rpt(sid, i+3, 0, 1, NCOLS, fmt(bg=BLUE_PALE if i%2==0 else WHITE)))
    dt.append({"range": f"A4:H{3+len(CATEGORIES)}", "values": cat_rows})

    # Row 10 (0-indexed): total
    tot = 3 + len(CATEGORIES)   # 0-indexed → sheet row tot+1
    rq.append(rpt(sid, tot, 0, 1, NCOLS, fmt(bg=AMBER, bold=True)))
    dt.append({"range": f"A{tot+1}:H{tot+1}",
               "values": [["TOTAL",
                            f"=SUM(B4:B{tot})", f"=SUM(C4:C{tot})",
                            f"=SUM(D4:D{tot})", "", f"=SUM(F4:F{tot})", "", ""]]})

    # Row 11 (0-indexed): settlement alert
    alert = tot + 1
    rq.append(mrg(sid, alert, 0, alert+1, NCOLS))
    rq.append(rpt(sid, alert, 0, 1, NCOLS, fmt(bg=rgb(255,199,206), bold=True, halign="CENTER")))
    dt.append({"range": f"A{alert+1}",
               "values": [[f'=IF(F{tot+1}>0,'
                            f'"⚠  "&F{tot+1}&" transaction(s) need discussion",'
                            f'"✓  No open settlements this month")']]})

    # Heat map on Spent col in summary
    rq.append(heat(sid, 3, tot, 2))

    # Row 12 (0-indexed = HDR_ROW-1): transaction column header
    rq.append(rpt(sid, HDR_ROW-1, 0, 1, NCOLS,
                  fmt(bg=BLUE_DARK, bold=True, fg=WHITE, halign="CENTER")))
    dt.append({"range": f"A{HDR_ROW}:H{HDR_ROW}",
               "values": [["Card", "Date", "Merchant", "Amount",
                            "Category", "Month", "Settlement", "Notes"]]})

    # Transaction rows — alternating row shading
    for i in range(ROWS):
        rq.append(rpt(sid, DATA_ROW-1+i, 0, 1, NCOLS,
                      fmt(bg=BLUE_PALE if i%2==0 else WHITE)))

    # Dropdowns on all transaction rows
    rq.append(ddv(sid, DATA_ROW-1, DATA_ROW-1+ROWS, CARD_COL,   CARDS))
    rq.append(ddv(sid, DATA_ROW-1, DATA_ROW-1+ROWS, CAT_COL,    CATEGORIES))
    rq.append(ddv(sid, DATA_ROW-1, DATA_ROW-1+ROWS, MONTH_COL,  MONTH_OPTS))
    rq.append(ddv(sid, DATA_ROW-1, DATA_ROW-1+ROWS, SETTLE_COL, SETTLEMENT))

    # Heat map on Amount col
    rq.append(heat(sid, DATA_ROW-1, DATA_ROW-1+ROWS, AMT_COL))

    # Row highlights: red = needs discussion, green = settled
    rq.append(row_cfmt(sid, DATA_ROW-1, DATA_ROW-1+ROWS, "Needs discussion", rgb(255,199,206), idx=1))
    rq.append(row_cfmt(sid, DATA_ROW-1, DATA_ROW-1+ROWS, "Settled",          rgb(198,239,206), idx=2))

    # Wrap Settlement and Notes
    rq.append(wrap(sid, DATA_ROW-1, DATA_ROW-1+ROWS, SETTLE_COL))
    rq.append(wrap(sid, DATA_ROW-1, DATA_ROW-1+ROWS, NOTES_COL))

    # Freeze through the transaction header
    rq.append(frz(sid, rows=HDR_ROW, cols=0))

    # Sample data — sorted by date, as the CLI would produce after all three card imports
    rows = sorted(SAMPLES, key=lambda r: r[1])   # sort by date string
    dt.append({"range": f"A{DATA_ROW}:H{DATA_ROW+len(rows)-1}",
               "values": [list(r) for r in rows]})

    return rq, dt


# ── filter views — one per card ────────────────────────────────────────────────

def filter_view(sid, title, card_name):
    """Creates a filter view that shows only rows where col A (Card) = card_name."""
    return {"addFilterView": {"filter": {
        "title": title,
        "range": {
            "sheetId": sid,
            "startRowIndex": HDR_ROW - 1,   # include header row
            "startColumnIndex": 0,
            "endColumnIndex": NCOLS,
        },
        "filterSpecs": [{
            "columnIndex": CARD_COL,
            "filterCriteria": {
                "condition": {
                    "type": "TEXT_EQ",
                    "values": [{"userEnteredValue": card_name}]
                }
            }
        }]
    }}}


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    creds   = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds)
    sid     = SHEET_ID

    print(f"Building: https://docs.google.com/spreadsheets/d/{sid}/edit")

    # Clear — add a temp sheet, delete everything else, add fresh "Card Transactions Jun 2026 — Detailed".
    # This ensures no leftover slicers or filter views from previous runs.
    service.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [
        {"addSheet": {"properties": {"title": "_tmp"}}}]}).execute()
    info = service.spreadsheets().get(spreadsheetId=sid).execute()
    tmp_id = next(s["properties"]["sheetId"] for s in info["sheets"]
                  if s["properties"]["title"] == "_tmp")
    to_del = [{"deleteSheet": {"sheetId": s["properties"]["sheetId"]}}
              for s in info["sheets"] if s["properties"]["sheetId"] != tmp_id]
    if to_del:
        service.spreadsheets().batchUpdate(
            spreadsheetId=sid, body={"requests": to_del}).execute()
    result = service.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [
        {"addSheet": {"properties": {"title": "Card Transactions Jun 2026 — Detailed", "index": 0}}}]}).execute()
    keep = result["replies"][0]["addSheet"]["properties"]["sheetId"]
    service.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [
        {"deleteSheet": {"sheetId": tmp_id}}]}).execute()

    # Formatting + structure
    rq, dt = build_all_cards(keep)

    # Filter views (one per card) — added in same batch as formatting
    for card in CARDS:
        rq.append(filter_view(keep, card, card))

    service.spreadsheets().batchUpdate(
        spreadsheetId=sid, body={"requests": rq}).execute()

    # Write data
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=sid,
        body={"valueInputOption": "USER_ENTERED", "data":
              [{"range": r["range"], "values": r["values"]} for r in dt]}).execute()

    # Number formats + slicers (after data write)
    nok = {"numberFormat": {"type": "NUMBER",  "pattern": "#,##0"}}
    pct = {"numberFormat": {"type": "PERCENT", "pattern": "0%"}}
    post = [
        # Amount column — number format
        {"repeatCell": {
            "range": {"sheetId": keep, "startRowIndex": DATA_ROW-1, "endRowIndex": DATA_ROW+ROWS,
                      "startColumnIndex": AMT_COL, "endColumnIndex": AMT_COL+1},
            "cell": {"userEnteredFormat": nok}, "fields": "userEnteredFormat.numberFormat"}},
        # Budget Spent column — also number format
        {"repeatCell": {
            "range": {"sheetId": keep, "startRowIndex": 3, "endRowIndex": 3+len(CATEGORIES),
                      "startColumnIndex": 2, "endColumnIndex": 4},
            "cell": {"userEnteredFormat": nok}, "fields": "userEnteredFormat.numberFormat"}},
        # % column
        {"repeatCell": {
            "range": {"sheetId": keep, "startRowIndex": 3, "endRowIndex": 3+len(CATEGORIES),
                      "startColumnIndex": 4, "endColumnIndex": 5},
            "cell": {"userEnteredFormat": pct}, "fields": "userEnteredFormat.numberFormat"}},
    ]

    # Slicers anchored right of col H
    for col_i, title, arow, acol, w, h in [
        (CAT_COL,    "Category",   1,  8, 155, 195),
        (CARD_COL,   "Card",       7,  8, 155, 130),
        (SETTLE_COL, "Settlement", 1, 10, 200, 255),
    ]:
        post.append({"addSlicer": {"slicer": {
            "spec": {
                "dataRange": {"sheetId": keep,
                              "startRowIndex": DATA_ROW-1, "endRowIndex": DATA_ROW+ROWS,
                              "startColumnIndex": 0, "endColumnIndex": NCOLS},
                "columnIndex": col_i,
                "title": title,
            },
            "position": {"overlayPosition": {
                "anchorCell": {"sheetId": keep, "rowIndex": arow, "columnIndex": acol},
                "widthPixels": w, "heightPixels": h,
            }},
        }}})

    service.spreadsheets().batchUpdate(
        spreadsheetId=sid, body={"requests": post}).execute()

    print("\nDone — single sheet 'Card Transactions [Month] — Detailed':")
    print("  - Budget summary rows 1-12")
    print("  - All transactions rows 14+ (all cards mixed, sorted by date)")
    print("  - Filter views: Data → Change view → [Shared Card / Shared Credit / Personal]")
    print("  - All cells editable — CLI pre-fills Card column on import")
    print("\nTo filter manually: Data → Create a filter")


if __name__ == "__main__":
    main()
