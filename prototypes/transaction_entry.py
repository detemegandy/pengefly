"""
PROTOTYPE — transaction entry sheet layout v4

Architecture:
  B: per-card sheets = editable source (paste here, annotate here, settle here)
  A: All Cards = read-only QUERY union of all B sheets (notes flow up automatically)

Per-card columns: Date | Merchant | Amount | Category | Settlement | Notes
  (Card column hidden col A, pre-filled by CLI — used by the QUERY in A)

Per-card summary: simple category totals only (no budget comparison — that lives on A)

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

# Col layout in B sheets (0-indexed):
#  0=Card(hidden)  1=Date  2=Merchant  3=Amount  4=Category  5=Month  6=Settlement  7=Notes
CARD_COL, DATE_COL, MERCH_COL, AMT_COL = 0, 1, 2, 3
CAT_COL, MONTH_COL, SETTLE_COL, NOTES_COL = 4, 5, 6, 7

SAMPLES = {
    "Shared Card": [
        ("01.06", "Extra supermarked",  359,  "Groceries",     "This month", "",                 ""),
        ("02.06", "CircleK",            553,  "Groceries",     "This month", "",                 ""),
        ("05.06", "Rema 1000",          981,  "Groceries",     "This month", "",                 ""),
        ("08.06", "Apotek 1",           237,  "Other",         "This month", "",                 ""),
        ("12.06", "Extra supermarked",  969,  "Groceries",     "This month", "",                 ""),
        ("14.06", "Lekebutikken",       480,  "Liana",         "This month", "",                 ""),
        ("18.06", "Bauhaus",           1678,  "Liana",         "This month", "",                 "Gate for Liana — paid shared card"),
        ("20.06", "Extra supermarked",  713,  "Groceries",     "This month", "",                 ""),
        ("22.06", "Europris",           293,  "Shopping",      "This month", "",                 ""),
        ("25.06", "Extra supermarked",  869,  "Groceries",     "This month", "",                 ""),
    ],
    "Shared Credit": [
        ("03.06", "McDonalds",          299,  "Entertainment", "This month", "",                 ""),
        ("07.06", "Bellis restaurant", 1026,  "Entertainment", "This month", "",                 ""),
        ("15.06", "Dyreklinikken",      850,  "Pets",          "This month", "",                 "Annual vet check"),
        ("17.06", "Pasha restaurant",   778,  "Entertainment", "This month", "",                 ""),
        ("19.06", "Zouq restaurant",   1189,  "Entertainment", "This month", "",                 ""),
        ("21.06", "Pizza To Go",        299,  "Entertainment", "This month", "",                 ""),
        ("23.06", "Renseriet",          863,  "Shopping",      "This month", "",                 ""),
    ],
    "Personal": [
        ("04.06", "Netflix",            139,  "",              "This month", "",                 ""),
        ("06.06", "Spotify",             99,  "",              "This month", "",                 ""),
        ("09.06", "Dyreklinikken",      450,  "Pets",          "This month", "Needs discussion", "Shared card was at limit — reimburse from Pets fund"),
        ("16.06", "Extra supermarked",  648,  "Groceries",     "This month", "Needs discussion", "Forgot shared card at home"),
        ("24.06", "Tool Pool",          249,  "Other",         "This month", "Needs discussion", "Shared tool rental — Andrea booked with personal"),
    ],
}

# ── colour helpers ─────────────────────────────────────────────────────────────

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

def rpt(sid, r, c, rows, cols, f):
    return {"repeatCell": {
        "range": {"sheetId": sid, "startRowIndex": r, "endRowIndex": r+rows,
                  "startColumnIndex": c, "endColumnIndex": c+cols},
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

def hide_col(sid, col):
    return {"updateDimensionProperties": {
        "range": {"sheetId": sid, "dimension": "COLUMNS",
                  "startIndex": col, "endIndex": col+1},
        "properties": {"hiddenByUser": True}, "fields": "hiddenByUser"}}

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
                    "startColumnIndex": 0, "endColumnIndex": 9}],
        "booleanRule": {
            "condition": {"type": "TEXT_CONTAINS",
                          "values": [{"userEnteredValue": text}]},
            "format": {"backgroundColor": color},
        }}, "index": idx}}

def slicer(sid, col_idx, title, anchor_row, anchor_col, w, h):
    return {"addSlicer": {"slicer": {
        "spec": {
            "dataRange": {"sheetId": sid,
                          "startRowIndex": 11, "endRowIndex": 111,
                          "startColumnIndex": 0, "endColumnIndex": 9},
            "columnIndex": col_idx,
            "title": title,
        },
        "position": {"overlayPosition": {
            "anchorCell": {"sheetId": sid, "rowIndex": anchor_row, "columnIndex": anchor_col},
            "widthPixels": w, "heightPixels": h,
        }},
    }}}


# ── B: per-card editable sheet ─────────────────────────────────────────────────
# Hidden col A = Card name (pre-filled by CLI, used by QUERY in combined sheet)
# Visible cols B-H: Date | Merchant | Amount | Category | Month | Settlement | Notes
# Top: simple per-category totals (no budget comparison)
# Bottom: transaction table

def build_card_sheet(sid, sname, card_name, samples):
    rq, dt = [], []
    TR   = 14   # 0-indexed row where transactions start
    ROWS = 80

    # column widths (col 0 = hidden Card)
    for i, px in enumerate([0, 75, 210, 90, 130, 115, 210, 200]):
        rq.append(cw(sid, i, px))
    rq.append(hide_col(sid, 0))   # hide Card column

    # Row 0: title
    rq.append(mrg(sid, 0, 0, 1, 8))
    rq.append(rpt(sid, 0, 0, 1, 8, fmt(bg=BLUE_DARK, bold=True, fg=WHITE, size=12, halign="CENTER")))
    dt.append({"range": f"{sname}!A1", "values": [[f"June 2026 — {card_name}"]]})

    # Row 1: instruction
    rq.append(mrg(sid, 1, 0, 2, 8))
    rq.append(rpt(sid, 1, 0, 1, 8, fmt(bg=YELLOW_SOFT, italic=True)))
    dt.append({"range": f"{sname}!A2",
               "values": [["  Paste bank export (Date, Merchant, Amount) into cols B-D. "
                            "CLI adds dropdowns. Leave Category empty = personal. "
                            "To filter: Data → Create a filter."]]}),

    # Row 2: category totals header
    rq.append(rpt(sid, 2, 0, 1, 8, fmt(bg=BLUE_LIGHT, bold=True, halign="CENTER")))
    dt.append({"range": f"{sname}!B3:H3",
               "values": [["Category", "Spent on this card", "", "", "", "", ""]]})

    # Rows 3-9: per-category totals (SUMIF on transaction block below)
    T1 = TR+1; T2 = TR+ROWS
    cat_rows = []
    for i, cat in enumerate(CATEGORIES):
        rn  = i + 4
        tot = f"=SUMIF(E${T1}:E${T2},B{rn},D${T1}:D${T2})"
        cat_rows.append(["", cat, tot, "", "", "", "", ""])
        rq.append(rpt(sid, i+3, 0, 1, 8, fmt(bg=BLUE_PALE if i%2==0 else WHITE)))
    dt.append({"range": f"{sname}!A4:H{3+len(CATEGORIES)}", "values": cat_rows})

    # Total row
    tot_r = 3 + len(CATEGORIES)   # 0-indexed
    rq.append(rpt(sid, tot_r, 0, 1, 8, fmt(bg=AMBER, bold=True)))
    dt.append({"range": f"{sname}!A{tot_r+1}:H{tot_r+1}",
               "values": [["", "TOTAL SHARED",
                            f"=SUMPRODUCT((E${T1}:E${T2}<>\"\")*D${T1}:D${T2})",
                            "", "", "", "", ""]]})

    # Settlement alert row
    alert_r = tot_r + 1
    rq.append(mrg(sid, alert_r, 0, alert_r+1, 8))
    rq.append(rpt(sid, alert_r, 0, 1, 8, fmt(bg=rgb(255,199,206), bold=True, halign="CENTER")))
    dt.append({"range": f"{sname}!A{alert_r+1}",
               "values": [[f'=IF(COUNTIF(G${T1}:G${T2},"Needs discussion")>0,'
                            f'"⚠  "&COUNTIF(G${T1}:G${T2},"Needs discussion")&'
                            f'" transaction(s) need discussion — see red rows below",'
                            f'"✓  All transactions settled or personal")']]})

    # Transaction header row
    hdr_r = TR - 1   # 0-indexed
    rq.append(rpt(sid, hdr_r, 0, 1, 8, fmt(bg=BLUE_DARK, bold=True, fg=WHITE, halign="CENTER")))
    dt.append({"range": f"{sname}!A{hdr_r+1}:H{hdr_r+1}",
               "values": [["(card)", "Date", "Merchant", "Amount",
                            "Category", "Month", "Settlement", "Notes"]]})

    # Transaction rows
    for i in range(ROWS):
        rq.append(rpt(sid, TR+i, 0, 1, 8, fmt(bg=BLUE_PALE if i%2==0 else WHITE)))

    # Dropdowns
    rq.append(ddv(sid, TR, TR+ROWS, CAT_COL,    CATEGORIES))
    rq.append(ddv(sid, TR, TR+ROWS, MONTH_COL,  MONTH_OPTS))
    rq.append(ddv(sid, TR, TR+ROWS, SETTLE_COL, SETTLEMENT))

    # Heat map on Amount
    rq.append(heat(sid, TR, TR+ROWS, AMT_COL))

    # Row color for settlement status
    rq.append(row_cfmt(sid, TR, TR+ROWS, "Needs discussion", rgb(255,199,206), idx=1))
    rq.append(row_cfmt(sid, TR, TR+ROWS, "Settled",          rgb(198,239,206), idx=2))

    # Sample data — col A = card name (hidden), B-H = visible data
    rows = [[card_name, d, m, a, cat, mo, s, n]
            for d, m, a, cat, mo, s, n in samples]
    dt.append({"range": f"{sname}!A{TR+1}:H{TR+len(rows)}", "values": rows})

    rq.append(frz(sid, rows=TR, cols=0))
    return rq, dt


# ── A: All Cards (read-only aggregate QUERY view) ─────────────────────────────
# Pulls from all three B sheets via QUERY union.
# Full budget summary at top. Slicers on the right.
# You never edit here — notes/settlement changes on B sheets appear automatically.

def build_combined(sid, sname, b_names):
    rq, dt = [], []
    TR   = 12   # 0-indexed, row 13 in Sheets = where QUERY result lands
    ROWS = 120

    # column widths: Card | Date | Merchant | Amount | Category | Month | Settlement | Notes
    for i, px in enumerate([110, 75, 200, 90, 130, 115, 210, 200]):
        rq.append(cw(sid, i, px))

    # Row 0: title
    rq.append(mrg(sid, 0, 0, 1, 8))
    rq.append(rpt(sid, 0, 0, 1, 8, fmt(bg=BLUE_DARK, bold=True, fg=WHITE, size=13, halign="CENTER")))
    dt.append({"range": f"{sname}!A1", "values": [["June 2026 — All Cards (read-only)"]]})

    # Row 1: instruction
    rq.append(mrg(sid, 1, 0, 2, 8))
    rq.append(rpt(sid, 1, 0, 1, 8, fmt(bg=YELLOW_SOFT, italic=True)))
    dt.append({"range": f"{sname}!A2",
               "values": [["  Aggregated view — edit notes and settlement on the individual card tabs. "
                            "Changes appear here automatically. "
                            "To filter columns: Data → Create a filter."]]}),

    # Rows 2-8: budget summary
    T1 = TR+1; T2 = TR+ROWS   # 1-indexed sheet rows of QUERY output (formula at A{TR+1}, expands down)
    rq.append(rpt(sid, 2, 0, 1, 8, fmt(bg=BLUE_LIGHT, bold=True, halign="CENTER")))
    dt.append({"range": f"{sname}!A3:H3",
               "values": [["Category", "Budget", "Spent", "Remaining", "%", "⚠ Open", "", ""]]})

    cat_rows = []
    for i, cat in enumerate(CATEGORIES):
        rn = i + 4
        b  = BUDGETS[cat]
        # Source is the QUERY block below — col A=Card, B=Date, C=Merchant, D=Amount,
        # E=Category, F=Month, G=Settlement, H=Notes
        spent = (f'=SUMPRODUCT((E${T1}:E${T2}=A{rn})'
                 f'*(A${T1}:A${T2}<>"Personal")'
                 f'*D${T1}:D${T2})')
        left  = f"=B{rn}-C{rn}"
        pct   = f"=IF(B{rn}>0,C{rn}/B{rn},0)"
        opens = (f'=COUNTIFS(E${T1}:E${T2},A{rn},'
                 f'G${T1}:G${T2},"Needs discussion")')
        cat_rows.append([cat, b, spent, left, pct, opens, "", ""])
        rq.append(rpt(sid, i+3, 0, 1, 8, fmt(bg=BLUE_PALE if i%2==0 else WHITE)))
    dt.append({"range": f"{sname}!A4:H{3+len(CATEGORIES)}", "values": cat_rows})

    tot = 3 + len(CATEGORIES)
    rq.append(rpt(sid, tot, 0, 1, 8, fmt(bg=AMBER, bold=True)))
    dt.append({"range": f"{sname}!A{tot+1}:H{tot+1}",
               "values": [["TOTAL",
                            f"=SUM(B4:B{tot})", f"=SUM(C4:C{tot})",
                            f"=SUM(D4:D{tot})", "", f"=SUM(F4:F{tot})", "", ""]]})

    # Heat map on Spent (col C = index 2) in summary
    rq.append(heat(sid, 3, tot, 2))

    # Settlement alert
    alert = tot + 1
    rq.append(mrg(sid, alert, 0, alert+1, 8))
    rq.append(rpt(sid, alert, 0, 1, 8, fmt(bg=rgb(255,199,206), bold=True, halign="CENTER")))
    dt.append({"range": f"{sname}!A{alert+1}",
               "values": [[f'=IF(F{tot+1}>0,'
                            f'"⚠  "&F{tot+1}&" transaction(s) need discussion — see red rows below",'
                            f'"✓  No open settlements this month")']]})

    # Row TR-1 = transaction list header (written by QUERY — no separate header needed)
    # But add a label row
    rq.append(rpt(sid, TR-1, 0, 1, 8, fmt(bg=BLUE_DARK, bold=True, fg=WHITE, halign="CENTER")))
    dt.append({"range": f"{sname}!A{TR}:H{TR}",
               "values": [["Card", "Date", "Merchant", "Amount",
                            "Category", "Month", "Settlement", "Notes"]]})

    # QUERY union — pulls hidden col A (Card) + visible cols B-H from each B sheet
    # Each B sheet: col A=Card, B=Date, C=Merchant, D=Amount, E=Cat, F=Month, G=Settle, H=Notes
    # We want ALL 8 cols, rows where Date is not empty (to skip blank rows)
    parts = []
    for bn in b_names:
        parts.append(f"'{bn}'!A{TR+1}:H{TR+ROWS}")
    union = "{" + ";".join(parts) + "}"
    query_formula = (f'=IFERROR(QUERY({union},'
                     f'"SELECT * WHERE Col2 <> \'\' ORDER BY Col2 ASC", 0),'
                     f'"No transactions yet")')
    dt.append({"range": f"{sname}!A{TR+1}", "values": [[query_formula]]})

    # Conditional formatting on query result area
    rq.append(heat(sid, TR+1, TR+1+ROWS, 3))
    rq.append(row_cfmt(sid, TR+1, TR+1+ROWS, "Needs discussion", rgb(255,199,206), idx=1))
    rq.append(row_cfmt(sid, TR+1, TR+1+ROWS, "Settled",          rgb(198,239,206), idx=2))

    rq.append(frz(sid, rows=TR+1, cols=0))
    return rq, dt


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    creds   = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds)
    sid     = SHEET_ID

    print(f"Building: https://docs.google.com/spreadsheets/d/{sid}/edit")

    # clear
    info    = service.spreadsheets().get(spreadsheetId=sid).execute()
    sheets  = info["sheets"]
    keep    = sheets[0]["properties"]["sheetId"]
    deletes = [{"deleteSheet": {"sheetId": s["properties"]["sheetId"]}} for s in sheets[1:]]
    if deletes:
        service.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": deletes}).execute()
    service.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [
        {"updateSheetProperties": {"properties": {"sheetId": keep, "title": "Sheet1"},
                                   "fields": "title"}}]}).execute()

    # add sheets
    add = service.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [
        {"addSheet": {"properties": {"title": "A: All Cards",     "index": 0}}},
        {"addSheet": {"properties": {"title": "B: Shared Card",   "index": 1}}},
        {"addSheet": {"properties": {"title": "B: Shared Credit", "index": 2}}},
        {"addSheet": {"properties": {"title": "B: Personal",      "index": 3}}},
    ]}).execute()
    ids = {r["addSheet"]["properties"]["title"]: r["addSheet"]["properties"]["sheetId"]
           for r in add["replies"]}

    info2 = service.spreadsheets().get(spreadsheetId=sid).execute()
    for s in info2["sheets"]:
        if s["properties"]["title"] == "Sheet1":
            service.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [
                {"deleteSheet": {"sheetId": s["properties"]["sheetId"]}}]}).execute()

    all_rq, all_dt = [], []

    for card, sname in [("Shared Card",   "B: Shared Card"),
                        ("Shared Credit", "B: Shared Credit"),
                        ("Personal",      "B: Personal")]:
        rq, dt = build_card_sheet(ids[sname], f"'{sname}'", card, SAMPLES[card])
        all_rq += rq; all_dt += dt

    rq, dt = build_combined(ids["A: All Cards"], "'A: All Cards'",
                             ["B: Shared Card", "B: Shared Credit", "B: Personal"])
    all_rq += rq; all_dt += dt

    service.spreadsheets().batchUpdate(
        spreadsheetId=sid, body={"requests": all_rq}).execute()
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=sid, body={"valueInputOption": "USER_ENTERED", "data": all_dt}).execute()

    # number formats + slicers (separate call)
    nok = {"numberFormat": {"type": "NUMBER",  "pattern": "#,##0"}}
    pct = {"numberFormat": {"type": "PERCENT", "pattern":     "0%"}}
    post = []
    for sname, s_id, amt_col, pct_col in [
        ("A: All Cards",     ids["A: All Cards"],     3, 4),
        ("B: Shared Card",   ids["B: Shared Card"],   3, None),
        ("B: Shared Credit", ids["B: Shared Credit"], 3, None),
        ("B: Personal",      ids["B: Personal"],      3, None),
    ]:
        post.append({"repeatCell": {
            "range": {"sheetId": s_id, "startRowIndex": 14, "endRowIndex": 100,
                      "startColumnIndex": amt_col, "endColumnIndex": amt_col+1},
            "cell": {"userEnteredFormat": nok}, "fields": "userEnteredFormat.numberFormat"}})
        if pct_col:
            post.append({"repeatCell": {
                "range": {"sheetId": s_id, "startRowIndex": 3, "endRowIndex": 11,
                          "startColumnIndex": pct_col, "endColumnIndex": pct_col+1},
                "cell": {"userEnteredFormat": pct}, "fields": "userEnteredFormat.numberFormat"}})

    # Slicers on A: All Cards — compact, anchored just past col H
    a_sid = ids["A: All Cards"]
    for col_i, title, arow, acol, w, h in [
        (4, "Category",   1,  8, 155, 200),
        (0, "Card",       7,  8, 155, 130),
        (6, "Settlement", 1, 10, 200, 260),
    ]:
        post.append({"addSlicer": {"slicer": {
            "spec": {
                "dataRange": {"sheetId": a_sid,
                              "startRowIndex": 12, "endRowIndex": 132,
                              "startColumnIndex": 0, "endColumnIndex": 8},
                "columnIndex": col_i,
                "title": title,
            },
            "position": {"overlayPosition": {
                "anchorCell": {"sheetId": a_sid, "rowIndex": arow, "columnIndex": acol},
                "widthPixels": w, "heightPixels": h,
            }},
        }}})

    service.spreadsheets().batchUpdate(
        spreadsheetId=sid, body={"requests": post}).execute()

    print("\nDone:")
    print("  B: Shared Card   — editable; paste transactions, add notes, flag settlements here")
    print("  B: Shared Credit — same")
    print("  B: Personal      — 3 rows pre-flagged 'Needs discussion'")
    print("  A: All Cards     — read-only QUERY union; budget summary; slicers")
    print("\nTo enable column filters: Data → Create a filter (can't be set via API reliably)")


if __name__ == "__main__":
    main()
