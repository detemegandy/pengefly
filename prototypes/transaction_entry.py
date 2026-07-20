"""
PROTOTYPE — transaction entry sheet layout v7

Changes from v6:
  - main() no longer wipes all tabs — only deletes/recreates its own tab.
    Use run_all.py to build all prototype tabs simultaneously.
  - SAMPLES now includes auto-populated credit card minimum payment rows.
    Settlement = "Needs discussion" so they appear flagged for review at month close.
    Logic: only added when the card has a non-zero balance. At month close the user
    reviews and updates the actual payment amount.

Architecture:
  Card Transactions [Month] — Detailed = single editable master sheet (all transactions from all cards)
  Filter views = one per card, created via API — same data, focused per-card view

Import flow (what the CLI will do):
  1. User exports from bank → pastes into a temp "Import" sheet
  2. CLI reads it, maps columns, appends rows to A with Card column pre-filled
  3. User opens per-card filter view to annotate (Category, Settlement, Notes)

No formula sheets. Everything in A is editable.

Run standalone:  uv run prototypes/transaction_entry.py
Run all tabs:    uv run prototypes/run_all.py
THROWAWAY — do not merge to main.
"""

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

CREDS_FILE = "/Users/andreas.saltveit/.config/gcp/google-sheets-sa.json"
SCOPES     = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
SHEET_ID   = "1gE5wc6lXaAOsiQ_dpjg7uGBnUrtGr0-oy-56UKyTdQc"
TAB_NAME   = "Card Transactions Jun 2026 — Detailed"

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
ROWS     = 150  # max transaction rows to reserve for formatting

# Credit card minimum payments for Jun 2026 — auto-populated, flagged for review.
# Min payment = MAX(ROUND(balance × min_pct), floor).
# EIKA: MAX(ROUND(45000 × 0.03), 250) = 1,350
# Mona: MAX(ROUND(28000 × 0.03), 250) = 840
_CC_PAYMENTS = [
    ("Shared Card", "30.06", "⬇ CC min. payment — EIKA",     1350, "", "This month",
     "Needs discussion",
     "Auto-added: EIKA statement 45,000 × 3% min. Review actual payment at month close."),
    ("Shared Card", "30.06", "⬇ CC min. payment — Mona DNB",  840, "", "This month",
     "Needs discussion",
     "Auto-added: Mona DNB statement 28,000 × 3% min. Review actual payment at month close."),
]

# All sample data — what the CLI produces after importing all three card exports,
# plus the auto-added CC minimum payment rows.
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
] + list(_CC_PAYMENTS)


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


# ── build Card Transactions [Month] — Detailed ─────────────────────────────────

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
                            "CC min. payment rows auto-added at month end — review and update "
                            "Settlement to 'Settled — credit card' when paid. "
                            "To filter manually: Data → Create a filter."]]}),

    # Row 2: budget summary header
    rq.append(rpt(sid, 2, 0, 1, NCOLS, fmt(bg=BLUE_LIGHT, bold=True, halign="CENTER")))
    dt.append({"range": "A3:H3",
               "values": [["Category", "Budget", "Spent", "Remaining", "%", "⚠ Open", "", ""]]})

    # Rows 3–9: per-category budget rows
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

    # Total row
    tot = 3 + len(CATEGORIES)
    rq.append(rpt(sid, tot, 0, 1, NCOLS, fmt(bg=AMBER, bold=True)))
    dt.append({"range": f"A{tot+1}:H{tot+1}",
               "values": [["TOTAL",
                            f"=SUM(B4:B{tot})", f"=SUM(C4:C{tot})",
                            f"=SUM(D4:D{tot})", "", f"=SUM(F4:F{tot})", "", ""]]})

    # Alert banner (row 11)
    alert = tot + 1
    rq.append(mrg(sid, alert, 0, alert+1, NCOLS))
    rq.append(rpt(sid, alert, 0, 1, NCOLS, fmt(bg=rgb(255,199,206), bold=True, halign="CENTER")))
    dt.append({"range": f"A{alert+1}",
               "values": [[f'=IF(F{tot+1}>0,'
                            f'"⚠  "&F{tot+1}&" transaction(s) need discussion",'
                            f'"✓  No open settlements this month")']]})

    # Heat map on Spent col in summary
    rq.append(heat(sid, 3, tot, 2))

    # Transaction column header (row 12)
    rq.append(rpt(sid, HDR_ROW-1, 0, 1, NCOLS,
                  fmt(bg=BLUE_DARK, bold=True, fg=WHITE, halign="CENTER")))
    dt.append({"range": f"A{HDR_ROW}:H{HDR_ROW}",
               "values": [["Card", "Date", "Merchant", "Amount",
                            "Category", "Month", "Settlement", "Notes"]]})

    # Transaction rows — alternating shading
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

    rq.append(frz(sid, rows=HDR_ROW, cols=0))

    # Sample data — sorted by date (as CLI would produce), CC payment rows at end
    regular = sorted([r for r in SAMPLES if "CC min" not in r[2]], key=lambda r: r[1])
    cc_rows = [r for r in SAMPLES if "CC min" in r[2]]
    rows = regular + cc_rows
    dt.append({"range": f"A{DATA_ROW}:H{DATA_ROW+len(rows)-1}",
               "values": [list(r) for r in rows]})

    return rq, dt


# ── filter views — one per card ────────────────────────────────────────────────

def filter_view(sid, title, card_name):
    """Creates a filter view showing only rows where col A (Card) = card_name."""
    return {"addFilterView": {"filter": {
        "title": title,
        "range": {
            "sheetId": sid,
            "startRowIndex": HDR_ROW - 1,
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

    # Only delete/recreate the transaction entry tab — other tabs are preserved.
    # (Previous v6 wipe-all approach would destroy Yearly 2026 and other tabs.)
    info = service.spreadsheets().get(spreadsheetId=sid).execute()
    to_del = [{"deleteSheet": {"sheetId": s["properties"]["sheetId"]}}
              for s in info["sheets"] if s["properties"]["title"] == TAB_NAME]
    if to_del:
        service.spreadsheets().batchUpdate(
            spreadsheetId=sid, body={"requests": to_del}).execute()

    result = service.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [
        {"addSheet": {"properties": {"title": TAB_NAME}}}]}).execute()
    keep = result["replies"][0]["addSheet"]["properties"]["sheetId"]

    rq, dt = build_all_cards(keep)
    for card in CARDS:
        rq.append(filter_view(keep, card, card))

    service.spreadsheets().batchUpdate(
        spreadsheetId=sid, body={"requests": rq}).execute()

    # Qualify every range with the tab name — bare ranges default to the first sheet,
    # which is "Yearly 2026" after run_all.py builds all tabs.
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=sid,
        body={"valueInputOption": "USER_ENTERED", "data":
              [{"range": f"'{TAB_NAME}'!{r['range']}", "values": r["values"]} for r in dt]}).execute()

    nok = {"numberFormat": {"type": "NUMBER",  "pattern": "#,##0"}}
    pct = {"numberFormat": {"type": "PERCENT", "pattern": "0%"}}
    post = [
        {"repeatCell": {
            "range": {"sheetId": keep, "startRowIndex": DATA_ROW-1, "endRowIndex": DATA_ROW+ROWS,
                      "startColumnIndex": AMT_COL, "endColumnIndex": AMT_COL+1},
            "cell": {"userEnteredFormat": nok}, "fields": "userEnteredFormat.numberFormat"}},
        {"repeatCell": {
            "range": {"sheetId": keep, "startRowIndex": 3, "endRowIndex": 3+len(CATEGORIES),
                      "startColumnIndex": 2, "endColumnIndex": 4},
            "cell": {"userEnteredFormat": nok}, "fields": "userEnteredFormat.numberFormat"}},
        {"repeatCell": {
            "range": {"sheetId": keep, "startRowIndex": 3, "endRowIndex": 3+len(CATEGORIES),
                      "startColumnIndex": 4, "endColumnIndex": 5},
            "cell": {"userEnteredFormat": pct}, "fields": "userEnteredFormat.numberFormat"}},
    ]
    service.spreadsheets().batchUpdate(
        spreadsheetId=sid, body={"requests": post}).execute()

    print(f"\nDone — '{TAB_NAME}' (other tabs preserved)")
    print("  CC minimum payment rows at bottom — flagged 'Needs discussion' for month-close review")
    print("  Filter views: Data → Change view → [Shared Card / Shared Credit / Personal]")


if __name__ == "__main__":
    main()
