"""
PROTOTYPE — config sheet layout v3

Sections:
  ACCOUNTS          — master list: account name + number (dropdown source)
  FLEX CATEGORIES   — category + account (dropdown) + account no (VLOOKUP) + budget
  FIXED EXPENSES    — expense + account (dropdown) + account no (VLOOKUP) + amount
  SAVINGS TRANSFERS — from/to accounts (dropdowns) + account nos (VLOOKUPs) + amount
  CARD MAPPING      — per-card xlsx column headers for import

Run: uv run prototypes/config_sheet.py
THROWAWAY — do not merge to main.
"""

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

CREDS_FILE = "/Users/andreas.saltveit/.config/gcp/google-sheets-sa.json"
SCOPES     = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
SHEET_ID   = "1gE5wc6lXaAOsiQ_dpjg7uGBnUrtGr0-oy-56UKyTdQc"

# ── data ──────────────────────────────────────────────────────────────────────

ACCOUNTS = [
    ("Joint Salary",    "1228.66.85291", "income",   "Salary lands here; transfers out monthly"),
    ("Shared Card",     "1228.66.84686", "spending",  "Day-to-day flex spending"),
    ("Shared Bill",     "1228.66.83612", "spending",  "Fixed bills (electricity, internet, gym)"),
    ("Shared Credit",   "—",             "spending",  "Credit card — settled monthly"),
    ("Personal",        "—",             "spending",  "Andreas personal spending"),
    ("House & Ins.",    "1226.55.75135", "loan",      "House loan + home insurance"),
    ("Car Loan",        "1228.66.84627", "loan",      "Car loan repayments"),
    ("Saving-General",  "1228.66.84902", "savings",   "General savings"),
    ("Saving-Vacation", "1228.66.85038", "savings",   "Vacation savings"),
    ("Saving-House",    "1228.66.84945", "savings",   "House savings"),
    ("Pets Fund",       "1228.66.85127", "savings",   "Pets emergency/savings"),
    ("SOS",             "1228.66.58178", "savings",   "Emergency fund"),
    ("Liana",           "1228.66.85089", "savings",   "Liana's fund"),
    ("Baby",            "1229.46.50049", "savings",   "Baby/Minis fund"),
]

# (name, account_name, monthly_budget)
FLEX = [
    ("Groceries",     "Shared Card",  13000),
    ("Entertainment", "Shared Card",   1000),
    ("Shopping",      "Shared Card",   3000),
    ("Other",         "Shared Card",   3000),
    ("Media",         "Shared Card",    288),
    ("Liana",         "Shared Card",   2000),
    ("Pets",          "Shared Card",   1500),
]

# (name, account_name, monthly_amount)
FIXED = [
    ("Loan + insurance", "House & Ins.", 35105),
    ("Electricity",      "Shared Bill",   4000),
    ("Internet",         "Shared Bill",    500),
    ("Gym",              "Shared Bill",    900),
    ("Barnehage",        "Shared Bill",   3000),
]

# (label, from_account, to_account, monthly_amount)
TRANSFERS = [
    ("General savings",    "Joint Salary", "Saving-General",  2000),
    ("Vacation savings",   "Joint Salary", "Saving-Vacation", 2000),
    ("House savings",      "Joint Salary", "Saving-House",    1000),
    ("Pets fund",          "Joint Salary", "Pets Fund",        500),
    ("SOS fund",           "Joint Salary", "SOS",              500),
    ("Liana",              "Joint Salary", "Liana",           1000),
    ("Baby",               "Joint Salary", "Baby",            1000),
    ("Maman (debt)",       "Joint Salary", "—",               2000),
    ("Andreas personal",   "Joint Salary", "—",               5483),
    ("Mona personal",      "Joint Salary", "—",               5500),
    ("Extra loan paydown", "Joint Salary", "House & Ins.",    1000),
]

# (card_name, date_col, description_col, amount_col, amount_sign, notes)
CARD_MAPPING = [
    ("Shared Card",   "Dato", "Tekst",             "Beløp", "debit",  "DNB debit format"),
    ("Shared Credit", "Dato", "Forklarende tekst", "Beløp", "credit", "DNB credit format"),
    ("Personal",      "Dato", "Tekst",             "Beløp", "debit",  "DNB debit format"),
]

# ── helpers ───────────────────────────────────────────────────────────────────

def rgb(r, g, b):
    return {"red": r/255, "green": g/255, "blue": b/255}

BLUE_DARK   = rgb(68,  114, 196)
BLUE_LIGHT  = rgb(180, 198, 231)
BLUE_PALE   = rgb(221, 235, 246)
AMBER       = rgb(255, 192,   0)
GREEN_FILL  = rgb(226, 239, 218)
WHITE       = rgb(255, 255, 255)
GREY        = rgb(217, 217, 217)
GREY_TEXT   = rgb(120, 120, 120)
YELLOW_SOFT = rgb(255, 243, 204)
PURPLE_PALE = rgb(234, 226, 245)

def fmt(bg=None, bold=False, fg=None, size=None, halign=None, italic=False, wrap=None):
    f = {}
    if bg:   f["backgroundColor"] = bg
    tf = {}
    if bold:   tf["bold"]   = True
    if italic: tf["italic"] = True
    if fg:     tf["foregroundColor"] = fg
    if size:   tf["fontSize"] = size
    if tf:   f["textFormat"] = tf
    if halign: f["horizontalAlignment"] = halign
    if wrap:   f["wrapStrategy"] = wrap
    return f

def rpt(sid, r, c, nrows, ncols, f):
    return {"repeatCell": {
        "range": {"sheetId": sid, "startRowIndex": r, "endRowIndex": r+nrows,
                  "startColumnIndex": c, "endColumnIndex": c+ncols},
        "cell": {"userEnteredFormat": f},
        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy)",
    }}

def valign_mid(sid, r, c1, c2):
    return {"repeatCell": {
        "range": {"sheetId": sid, "startRowIndex": r, "endRowIndex": r+1,
                  "startColumnIndex": c1, "endColumnIndex": c2},
        "cell": {"userEnteredFormat": {"verticalAlignment": "MIDDLE"}},
        "fields": "userEnteredFormat.verticalAlignment",
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

def rh(sid, r, px):
    return {"updateDimensionProperties": {
        "range": {"sheetId": sid, "dimension": "ROWS", "startIndex": r, "endIndex": r+1},
        "properties": {"pixelSize": px}, "fields": "pixelSize"}}

def frz(sid, rows=0, cols=0):
    return {"updateSheetProperties": {
        "properties": {"sheetId": sid,
                       "gridProperties": {"frozenRowCount": rows, "frozenColumnCount": cols}},
        "fields": "gridProperties(frozenRowCount,frozenColumnCount)"}}

def border_box(sid, r1, c1, r2, c2):
    solid = {"style": "SOLID", "width": 1, "color": rgb(180, 198, 231)}
    inner = {"style": "SOLID", "width": 1, "color": rgb(217, 217, 217)}
    return {"updateBorders": {
        "range": {"sheetId": sid, "startRowIndex": r1, "endRowIndex": r2,
                  "startColumnIndex": c1, "endColumnIndex": c2},
        "top": solid, "bottom": solid, "left": solid, "right": solid,
        "innerHorizontal": inner, "innerVertical": inner,
    }}

def dropdown(sid, r1, r2, col, acct_names_a1):
    """Data validation dropdown sourced from the accounts names column."""
    return {"setDataValidation": {
        "range": {"sheetId": sid, "startRowIndex": r1, "endRowIndex": r2,
                  "startColumnIndex": col, "endColumnIndex": col+1},
        "rule": {
            "condition": {"type": "ONE_OF_RANGE",
                          "values": [{"userEnteredValue": f"={acct_names_a1}"}]},
            "showCustomUi": True,
            "strict": False,
        }
    }}

def col_letter(idx):
    s = ""
    while True:
        s = chr(ord('A') + idx % 26) + s
        idx = idx // 26 - 1
        if idx < 0:
            break
    return s

def vlookup(lookup_col, row_1idx, lookup_range_a1):
    """VLOOKUP formula: look up account number from account name."""
    return f'=IFERROR(VLOOKUP({lookup_col}{row_1idx},{lookup_range_a1},2,FALSE),"—")'

def to_a1(sheet_name, row, col):
    return f"'{sheet_name}'!{col_letter(col)}{row+1}"

def section_header(sid, rq, dt, row, text, ncols, bg=BLUE_DARK):
    rq.append(mrg(sid, row, 0, row+1, ncols))
    rq.append(rpt(sid, row, 0, 1, ncols, fmt(bg=bg, bold=True, fg=WHITE, size=11)))
    dt.append((row, 0, f"  {text}"))
    return row + 1

def col_header_row(sid, rq, dt, row, labels, ncols):
    rq.append(rpt(sid, row, 0, 1, ncols, fmt(bg=BLUE_LIGHT, bold=True, wrap="WRAP")))
    rq.append(valign_mid(sid, row, 0, ncols))
    rq.append(rh(sid, row, 40))
    for c, label in enumerate(labels):
        dt.append((row, c, label))
    return row + 1

def note_row(sid, rq, dt, row, text, ncols):
    rq.append(mrg(sid, row, 0, row+1, ncols))
    rq.append(rpt(sid, row, 0, 1, ncols, fmt(bg=YELLOW_SOFT, italic=True, fg=GREY_TEXT, wrap="WRAP")))
    rq.append(valign_mid(sid, row, 0, ncols))
    rq.append(rh(sid, row, 52))
    dt.append((row, 0, f"  ℹ  {text}"))
    return row + 1


# ── sheet builder ─────────────────────────────────────────────────────────────

def build_config(sid, sheet_name):
    rq, dt, val = [], [], []   # val = validation requests (applied separately)

    NCOLS = 7
    rq += [
        cw(sid, 0, 190),  # A: label
        cw(sid, 1, 145),  # B: account name (dropdown)
        cw(sid, 2, 120),  # C: account number (VLOOKUP, grey)
        cw(sid, 3,  90),  # D: amount / col-header
        cw(sid, 4, 145),  # E: to-account name (dropdown, transfers only)
        cw(sid, 5, 120),  # F: to-account number (VLOOKUP, grey)
        cw(sid, 6, 240),  # G: notes
        frz(sid, rows=2),
    ]

    # ── title + banner ────────────────────────────────────────────────────────
    rq.append(mrg(sid, 0, 0, 1, NCOLS))
    rq.append(rpt(sid, 0, 0, 1, NCOLS,
                  fmt(bg=BLUE_DARK, bold=True, fg=WHITE, size=13, halign="CENTER")))
    dt.append((0, 0, "Config — pengefly 2026"))

    rq.append(mrg(sid, 1, 0, 2, NCOLS))
    rq.append(rpt(sid, 1, 0, 1, NCOLS,
                  fmt(bg=YELLOW_SOFT, italic=True, fg=GREY_TEXT, wrap="WRAP")))
    rq.append(valign_mid(sid, 1, 0, NCOLS))
    rq.append(rh(sid, 1, 52))
    dt.append((1, 0,
        "  Edit the green rows freely — changes take effect next time the CLI runs. "
        "To add a new account: add the row, then run 'pengefly account sync' to extend the dropdowns. "
        "Account number columns fill via VLOOKUP and cannot be edited directly."))

    row = 2

    # ══════════════════════════════════════════════════════════════════════════
    # ACCOUNTS
    # ══════════════════════════════════════════════════════════════════════════
    row += 1
    row = section_header(sid, rq, dt, row, "ACCOUNTS", NCOLS)
    row = note_row(sid, rq, dt, row,
        "Master account list — source for all dropdowns. "
        "Edit names or numbers here; all sections update automatically.", NCOLS)
    row = col_header_row(sid, rq, dt, row,
        ["Account Name", "Account Number", "Type", "", "", "", "Notes"], NCOLS)
    acct_data_start = row
    for i, (name, number, atype, notes) in enumerate(ACCOUNTS):
        bg = GREEN_FILL if i % 2 == 0 else WHITE
        rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=bg)))
        dt += [(row, 0, name), (row, 1, number), (row, 2, atype), (row, 6, notes)]
        row += 1
    acct_data_end = row
    rq.append(border_box(sid, acct_data_start, 0, acct_data_end, 3))

    # A1 ranges for dropdown source and VLOOKUP — absolute refs on same sheet
    # Fixed ranges — CLI updates these when accounts are added (pengefly config sync)
    acct_names_a1  = f"$A${acct_data_start+1}:$A${acct_data_end}"
    acct_lookup_a1 = f"$A${acct_data_start+1}:$B${acct_data_end}"

    # ══════════════════════════════════════════════════════════════════════════
    # FLEX CATEGORIES
    # ══════════════════════════════════════════════════════════════════════════
    row += 1
    row = section_header(sid, rq, dt, row, "FLEX CATEGORIES", NCOLS)
    row = note_row(sid, rq, dt, row,
        "Transaction-tagged spending. Pick an account from the dropdown — "
        "account number fills automatically. "
        "Monthly Budget = default for carry-forward in yearly overview.", NCOLS)
    row = col_header_row(sid, rq, dt, row,
        ["Category", "Account", "Account No", "Monthly Budget", "", "", "Notes"], NCOLS)
    flex_start = row
    for i, (name, acct, budget) in enumerate(FLEX):
        bg = GREEN_FILL if i % 2 == 0 else WHITE
        rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=bg)))
        rq.append(rpt(sid, row, 2, 1, 1, fmt(bg=GREY, fg=GREY_TEXT, italic=True)))  # VLOOKUP col
        rq.append(rpt(sid, row, 3, 1, 1, fmt(bg=bg, halign="RIGHT")))
        dt += [
            (row, 0, name),
            (row, 1, acct),
            (row, 2, vlookup("B", row+1, acct_lookup_a1)),
            (row, 3, budget),
        ]
        row += 1
    flex_end = row
    rq.append(border_box(sid, flex_start, 0, flex_end, 4))
    val.append(dropdown(sid, flex_start, flex_end, 1, acct_names_a1))

    # ══════════════════════════════════════════════════════════════════════════
    # FIXED EXPENSES
    # ══════════════════════════════════════════════════════════════════════════
    row += 1
    row = section_header(sid, rq, dt, row, "FIXED EXPENSES", NCOLS)
    row = note_row(sid, rq, dt, row,
        "Fixed monthly bills — no per-transaction tagging. "
        "Amount change mid-year: prior months keep the old amount (carry-forward); "
        "change month highlighted orange in yearly overview.", NCOLS)
    row = col_header_row(sid, rq, dt, row,
        ["Expense", "Account", "Account No", "Monthly Amount", "", "", "Notes"], NCOLS)
    fixed_start = row
    for i, (name, acct, amount) in enumerate(FIXED):
        bg = GREEN_FILL if i % 2 == 0 else WHITE
        rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=bg)))
        rq.append(rpt(sid, row, 2, 1, 1, fmt(bg=GREY, fg=GREY_TEXT, italic=True)))
        rq.append(rpt(sid, row, 3, 1, 1, fmt(bg=bg, halign="RIGHT")))
        dt += [
            (row, 0, name),
            (row, 1, acct),
            (row, 2, vlookup("B", row+1, acct_lookup_a1)),
            (row, 3, amount),
        ]
        row += 1
    fixed_end = row
    rq.append(border_box(sid, fixed_start, 0, fixed_end, 4))
    val.append(dropdown(sid, fixed_start, fixed_end, 1, acct_names_a1))

    # ══════════════════════════════════════════════════════════════════════════
    # SAVINGS TRANSFERS
    # ══════════════════════════════════════════════════════════════════════════
    row += 1
    row = section_header(sid, rq, dt, row, "SAVINGS TRANSFERS", NCOLS)
    row = note_row(sid, rq, dt, row,
        "Monthly transfers out of Joint Salary. Both From and To use dropdowns. "
        "New row mid-year: prior months show — in yearly overview (not backfilled).", NCOLS)
    row = col_header_row(sid, rq, dt, row,
        ["Transfer", "From Account", "From No",
         "Amount", "To Account", "To No", "Notes"], NCOLS)
    tx_start = row
    for i, (name, frm, to, amount) in enumerate(TRANSFERS):
        bg = GREEN_FILL if i % 2 == 0 else WHITE
        rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=bg)))
        rq.append(rpt(sid, row, 2, 1, 1, fmt(bg=GREY, fg=GREY_TEXT, italic=True)))
        rq.append(rpt(sid, row, 3, 1, 1, fmt(bg=bg, halign="RIGHT")))
        rq.append(rpt(sid, row, 5, 1, 1, fmt(bg=GREY, fg=GREY_TEXT, italic=True)))
        dt += [
            (row, 0, name),
            (row, 1, frm),
            (row, 2, vlookup("B", row+1, acct_lookup_a1)),
            (row, 3, amount),
            (row, 4, to),
            (row, 5, vlookup("E", row+1, acct_lookup_a1)),
        ]
        row += 1
    tx_end = row
    rq.append(border_box(sid, tx_start, 0, tx_end, NCOLS))
    val.append(dropdown(sid, tx_start, tx_end, 1, acct_names_a1))  # from
    val.append(dropdown(sid, tx_start, tx_end, 4, acct_names_a1))  # to

    # ══════════════════════════════════════════════════════════════════════════
    # CARD / BANK IMPORT MAPPING
    # ══════════════════════════════════════════════════════════════════════════
    row += 1
    row = section_header(sid, rq, dt, row, "CARD / BANK IMPORT MAPPING", NCOLS)
    row = note_row(sid, rq, dt, row,
        "Column headers from each xlsx export. "
        "Amount Sign: 'debit' = negative = expense; 'credit' = positive = expense. "
        "Placeholders — update once a sample file is available (issue #7).", NCOLS)
    row = col_header_row(sid, rq, dt, row,
        ["Card / Bank", "Date Col", "Description Col",
         "Amount Col", "Amount Sign", "", "Notes"], NCOLS)
    card_start = row
    for i, (card, date_col, desc_col, amt_col, sign, notes) in enumerate(CARD_MAPPING):
        bg = GREEN_FILL if i % 2 == 0 else WHITE
        rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=bg)))
        dt += [(row, 0, card), (row, 1, date_col), (row, 2, desc_col),
               (row, 3, amt_col), (row, 4, sign), (row, 6, notes)]
        row += 1
    rq.append(border_box(sid, card_start, 0, row, NCOLS))

    return rq, dt, val


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    creds   = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds)

    print(f"Building: https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit")

    SHEET_NAME = "Config"
    info = service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
    for s in info["sheets"]:
        if s["properties"]["title"] == SHEET_NAME:
            service.spreadsheets().batchUpdate(
                spreadsheetId=SHEET_ID,
                body={"requests": [{"deleteSheet": {"sheetId": s["properties"]["sheetId"]}}]}
            ).execute()

    result = service.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={"requests": [{"addSheet": {"properties": {"title": SHEET_NAME}}}]}
    ).execute()
    sid = result["replies"][0]["addSheet"]["properties"]["sheetId"]

    rq, dt, val = build_config(sid, SHEET_NAME)

    # Formatting + validation in one batchUpdate
    service.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID, body={"requests": rq + val}).execute()

    # Values (formulas use USER_ENTERED so VLOOKUP is interpreted)
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={"valueInputOption": "USER_ENTERED",
              "data": [{"range": to_a1(SHEET_NAME, r, c), "values": [[v]]} for r, c, v in dt]}
    ).execute()

    # Number format on amount column D
    service.spreadsheets().batchUpdate(spreadsheetId=SHEET_ID, body={"requests": [
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 120,
                      "startColumnIndex": 3, "endColumnIndex": 4},
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}},
            "fields": "userEnteredFormat.numberFormat"}}
    ]}).execute()

    print("\nDone:")
    print("  Account name columns → dropdown (sourced from Accounts table)")
    print("  Account number columns → VLOOKUP formula, grey/italic (read-only feel)")
    print("  Transfers: both From and To have independent dropdowns + VLOOKUPs")


if __name__ == "__main__":
    main()
