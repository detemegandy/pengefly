"""
PROTOTYPE — yearly overview sheet layout v2

Sections (rows) × 12 months (columns):
  STATUS       — per-month close state (Closed / Open / future)
  INCOME       — salary entered per month; June spikes with feriepenger
  TRANSFERS    — carry-forward amounts to savings buckets with account numbers;
                 status row per month (→ green when sent)
  FIXED        — carry-forward fixed expenses
  FLEX         — carry-forward budget vs actuals; orange cell = first month a budget changed;
                 grey = future carry-forward; no highlight = inherited unchanged
  BUFFER       — income − transfers − fixed − flex = what's left

Carry-forward: each month inherits the previous month's budget unless overridden.
Budget change rule: orange = value introduced here for the first time; grey = future carry-forward.

Run: uv run prototypes/yearly_overview.py
THROWAWAY — do not merge to main.
"""

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

CREDS_FILE = "/Users/andreas.saltveit/.config/gcp/google-sheets-sa.json"
SCOPES     = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
SHEET_ID   = "1gE5wc6lXaAOsiQ_dpjg7uGBnUrtGr0-oy-56UKyTdQc"

MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

# ── sample data ────────────────────────────────────────────────────────────────
SALARY = {
    "Andreas": [41000,41000,41000,41000,41000,70000,41000,41000,41000,41000,41000,41000],
    "Mona":    [50000,50000,50000,50000,50000,97000,50000,50000,50000,50000,50000,50000],
}

# (label, amount, account_number)
TRANSFERS = [
    ("General savings",    2000, "1228.66.84902"),
    ("Vacation savings",   2000, "1228.66.85038"),
    ("House savings",      1000, "1228.66.84945"),
    ("Pets fund",           500, "1228.66.85127"),
    ("SOS fund",            500, "1228.66.58178"),
    ("Liana",              1000, "1228.66.85089"),
    ("Baby",               1000, "1229.46.50049"),
    ("Maman (debt)",       2000, "—"),
    ("Andreas personal",   5483, "—"),
    ("Mona personal",      5500, "—"),
    ("Extra loan paydown", 1000, "1226.55.75135"),
]

FIXED = [
    ("Loan + insurance", 35105),
    ("Electricity",       4000),
    ("Internet",           500),
    ("Gym",                900),
    ("Barnehage",         3000),
]

# (name, monthly_budgets[12], jun_actual)
# monthly_budgets: list of 12 values — change value in a month = override; same = carry-forward
# orange highlight fires on the FIRST month where value differs from previous month
FLEX = [
    ("Groceries",
     [13000,13000,13000,13000,13000,13000,13000,14000,14000,14000,14000,14000],
     4622),   # Aug budget increase highlighted
    ("Entertainment",
     [1000,1000,1000,1000,1000,1000,1500,1500,1500,1000,1000,1000],
     3591),   # Jul increase, Oct drop back
    ("Shopping",
     [3000]*12, 1156),
    ("Other",
     [3000]*12,  237),
    ("Media",
     [288]*12,     0),
    ("Liana",
     [2000]*12, 2158),
    ("Pets",
     [1500]*12,  850),
]

# SENT[transfer_index][month_index] = amount in that cell (plan or actual), always filled.
# Carry-forward: each month inherits the previous month's value unless explicitly changed.
# Color: orange = value changed from previous month (first change OR change back both fire).
#         green  = closed month, value unchanged from previous.
#         amber  = open/current month, value unchanged from previous.
#         grey   = future month, value unchanged from previous.
# Demo: SOS and Liana short-paid in Mar (salary late), restored in Apr → both Mar and Apr are orange.
SENT = [
    [2000]*12,                                                        # General savings — unchanged
    [2000]*12,                                                        # Vacation savings — unchanged
    [1000]*12,                                                        # House savings — unchanged
    [ 500]*12,                                                        # Pets fund — unchanged
    [ 500, 500, 200, 500, 500, 500, 500, 500, 500, 500, 500, 500],  # SOS — short Mar, restored Apr
    [1000,1000, 600,1000,1000,1000,1000,1000,1000,1000,1000,1000],  # Liana — short Mar, restored Apr
    [1000]*12,                                                        # Baby — unchanged
    [2000]*12,                                                        # Maman (debt) — unchanged
    [5483]*12,                                                        # Andreas personal — unchanged
    [5500]*12,                                                        # Mona personal — unchanged
    [1000]*12,                                                        # Extra loan paydown — unchanged
]

# Month close state: "closed", "open", "future"
MONTH_STATE = ["closed","closed","closed","closed","closed","open",
               "future","future","future","future","future","future"]


# ── helpers ────────────────────────────────────────────────────────────────────

def rgb(r, g, b):
    return {"red": r/255, "green": g/255, "blue": b/255}

BLUE_DARK   = rgb(68,  114, 196)
BLUE_LIGHT  = rgb(180, 198, 231)
BLUE_PALE   = rgb(221, 235, 246)
AMBER       = rgb(255, 192,   0)
WHITE       = rgb(255, 255, 255)
GREEN_DARK  = rgb( 84, 130,  53)
GREEN       = rgb(198, 239, 206)
YELLOW      = rgb(255, 217, 102)
ORANGE      = rgb(255, 178, 102)   # budget change highlight
RED_LIGHT   = rgb(255, 199, 206)
GREY        = rgb(217, 217, 217)
GREY_TEXT   = rgb(120, 120, 120)
GREY_DARK   = rgb(150, 150, 150)

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

# Col layout: A=label | B=account | C=default | D-O=Jan-Dec | P=YTD
LABEL_COL   = 0
ACCT_COL    = 1
DEFAULT_COL = 2
JAN_COL     = 3   # months at cols 3-14
YTD_COL     = 15
NCOLS       = 16

def mcol(m):
    """0-indexed month → column index."""
    return JAN_COL + m

def to_a1(row, col):
    col_letter = ""
    c = col
    while True:
        col_letter = chr(ord('A') + c % 26) + col_letter
        c = c // 26 - 1
        if c < 0:
            break
    return f"'Yearly 2026'!{col_letter}{row+1}"


def build_yearly(sid):
    rq, dt = [], []

    # Column widths
    rq.append(cw(sid, LABEL_COL,   185))
    rq.append(cw(sid, ACCT_COL,    115))
    rq.append(cw(sid, DEFAULT_COL,  80))
    for c in range(JAN_COL, JAN_COL+12):
        rq.append(cw(sid, c, 68))
    rq.append(cw(sid, YTD_COL, 80))

    # ── Row 0: title ──────────────────────────────────────────────────────────
    rq.append(mrg(sid, 0, 0, 1, NCOLS))
    rq.append(rpt(sid, 0, 0, 1, NCOLS,
                  fmt(bg=BLUE_DARK, bold=True, fg=WHITE, size=13, halign="CENTER")))
    dt.append((0, 0, "2026 — Annual Overview"))

    # ── Row 1: column headers ─────────────────────────────────────────────────
    rq.append(rpt(sid, 1, 0, 1, NCOLS, fmt(bg=BLUE_LIGHT, bold=True, halign="CENTER")))
    dt.append((1, LABEL_COL, ""))
    dt.append((1, ACCT_COL, "Account"))
    dt.append((1, DEFAULT_COL, "Default"))
    for i, m in enumerate(MONTHS):
        state = MONTH_STATE[i]
        bg = BLUE_LIGHT if state == "closed" else (BLUE_PALE if state == "open" else GREY)
        fg = WHITE if state == "closed" else (None if state == "open" else GREY_TEXT)
        rq.append(rpt(sid, 1, mcol(i), 1, 1, fmt(bg=bg, bold=True, halign="CENTER", fg=fg)))
        dt.append((1, mcol(i), m))
    dt.append((1, YTD_COL, "YTD"))

    # ── Row 2: month close status ─────────────────────────────────────────────
    rq.append(rpt(sid, 2, 0, 1, NCOLS, fmt(bg=WHITE, bold=True, halign="CENTER")))
    dt.append((2, LABEL_COL, "Month status"))
    dt.append((2, ACCT_COL, ""))
    dt.append((2, DEFAULT_COL, ""))
    for i, state in enumerate(MONTH_STATE):
        if state == "closed":
            rq.append(rpt(sid, 2, mcol(i), 1, 1,
                          fmt(bg=GREEN, bold=True, fg=GREEN_DARK, halign="CENTER")))
            dt.append((2, mcol(i), "✓ Closed"))
        elif state == "open":
            rq.append(rpt(sid, 2, mcol(i), 1, 1,
                          fmt(bg=YELLOW, bold=True, halign="CENTER")))
            dt.append((2, mcol(i), "Open"))
        else:
            rq.append(rpt(sid, 2, mcol(i), 1, 1,
                          fmt(bg=GREY, fg=GREY_TEXT, halign="CENTER")))
            dt.append((2, mcol(i), "—"))

    rq.append(frz(sid, rows=3, cols=0))  # freeze title + header + status rows

    row = 3  # current writing row

    # ── INCOME ────────────────────────────────────────────────────────────────
    rq.append(mrg(sid, row, 0, row+1, NCOLS))
    rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=BLUE_DARK, bold=True, fg=WHITE)))
    dt.append((row, 0, "  INCOME")); row += 1

    for i, (name, monthly) in enumerate(SALARY.items()):
        bg = BLUE_PALE if i%2==0 else WHITE
        rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=bg)))
        dt.append((row, LABEL_COL, name))
        dt.append((row, DEFAULT_COL, monthly[0]))
        for m, v in enumerate(monthly):
            rq.append(rpt(sid, row, mcol(m), 1, 1,
                          fmt(bg=bg, halign="CENTER")))
            dt.append((row, mcol(m), v))
        row += 1

    # Total income row
    total_income = [SALARY["Andreas"][m] + SALARY["Mona"][m] for m in range(12)]
    rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=AMBER, bold=True)))
    dt.append((row, LABEL_COL, "Total income"))
    dt.append((row, DEFAULT_COL, total_income[0]))
    for m, v in enumerate(total_income):
        dt.append((row, mcol(m), v))
    dt.append((row, YTD_COL, sum(total_income[:6])))
    row += 2  # +spacer

    # ── TRANSFERS ─────────────────────────────────────────────────────────────
    rq.append(mrg(sid, row, 0, row+1, NCOLS))
    rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=BLUE_DARK, bold=True, fg=WHITE)))
    dt.append((row, 0, "  TRANSFERS  (from Joint Salary)")); row += 1

    for i, (name, amount, acct) in enumerate(TRANSFERS):
        bg = BLUE_PALE if i%2==0 else WHITE
        rq.append(rpt(sid, row, LABEL_COL, 1, 3, fmt(bg=bg)))
        dt.append((row, LABEL_COL, f"→ {name}"))
        dt.append((row, ACCT_COL,  acct))
        dt.append((row, DEFAULT_COL, amount))
        sent_row = SENT[i]
        prev = amount
        for m in range(12):
            val = sent_row[m]
            ms  = MONTH_STATE[m]
            if val != prev:
                cell_bg, cell_fg = ORANGE, None          # changed from previous — always orange
            elif ms == "future":
                cell_bg, cell_fg = GREY, GREY_TEXT       # carry-forward, not yet
            elif ms == "open":
                cell_bg, cell_fg = AMBER, None           # current month, plan unchanged
            else:
                cell_bg, cell_fg = GREEN, GREEN_DARK     # closed, sent as planned
            rq.append(rpt(sid, row, mcol(m), 1, 1,
                          fmt(bg=cell_bg, fg=cell_fg, halign="RIGHT")))
            dt.append((row, mcol(m), val))
            prev = val
        row += 1

    # Total transfers — sum actual SENT values per month
    total_tx_default = sum(a for _, a, _ in TRANSFERS)
    rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=AMBER, bold=True)))
    dt.append((row, LABEL_COL, "Total transfers"))
    dt.append((row, DEFAULT_COL, total_tx_default))
    for m in range(12):
        monthly_total = sum(SENT[i][m] for i in range(len(TRANSFERS)))
        dt.append((row, mcol(m), monthly_total))
    ytd_tx = sum(SENT[i][m] for i in range(len(TRANSFERS))
                 for m in range(6) if MONTH_STATE[m] in ("closed", "open"))
    dt.append((row, YTD_COL, ytd_tx))
    row += 2  # +spacer

    # ── FIXED EXPENSES ────────────────────────────────────────────────────────
    rq.append(mrg(sid, row, 0, row+1, NCOLS))
    rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=BLUE_DARK, bold=True, fg=WHITE)))
    dt.append((row, 0, "  FIXED EXPENSES")); row += 1

    total_fixed = sum(a for _, a in FIXED)
    for i, (name, amount) in enumerate(FIXED):
        bg = BLUE_PALE if i%2==0 else WHITE
        rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=bg)))
        dt.append((row, LABEL_COL, name))
        dt.append((row, DEFAULT_COL, amount))
        for m in range(12):
            rq.append(rpt(sid, row, mcol(m), 1, 1, fmt(bg=bg, halign="CENTER")))
            dt.append((row, mcol(m), amount))
        row += 1

    rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=AMBER, bold=True)))
    dt.append((row, LABEL_COL, "Total fixed"))
    dt.append((row, DEFAULT_COL, total_fixed))
    for m in range(12):
        dt.append((row, mcol(m), total_fixed))
    dt.append((row, YTD_COL, total_fixed * 6))
    row += 2  # +spacer

    # ── FLEX BUDGET ───────────────────────────────────────────────────────────
    rq.append(mrg(sid, row, 0, row+1, NCOLS))
    rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=BLUE_DARK, bold=True, fg=WHITE)))
    dt.append((row, 0, "  FLEX BUDGET  (orange = budget changed this month; grey = future plan)")); row += 1

    total_flex_default = sum(budgets[0] for _, budgets, _ in FLEX)
    total_flex_actual  = sum(actual    for _, _,       actual in FLEX)

    for i, (name, budgets, jun_actual) in enumerate(FLEX):
        bg_base = BLUE_PALE if i%2==0 else WHITE
        rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=bg_base)))
        dt.append((row, LABEL_COL, name))
        dt.append((row, DEFAULT_COL, budgets[0]))

        for m in range(12):
            col  = mcol(m)
            val  = budgets[m]
            prev = budgets[m-1] if m > 0 else budgets[0]
            state = MONTH_STATE[m]

            if state == "future":
                cell_bg = GREY
                cell_fg = GREY_TEXT
            elif m == 5 and jun_actual:   # Jun — show actual, color vs budget
                pct = jun_actual / val if val else 0
                cell_bg = GREEN if pct < 0.8 else (AMBER if pct <= 1.0 else RED_LIGHT)
                cell_fg = None
            elif val != prev:             # budget changed here — first occurrence
                cell_bg = ORANGE
                cell_fg = None
            else:
                cell_bg = bg_base
                cell_fg = None

            f = fmt(bg=cell_bg, halign="CENTER")
            if cell_fg:
                f["textFormat"] = {"foregroundColor": cell_fg}
            rq.append(rpt(sid, row, col, 1, 1, f))

            display = jun_actual if (m == 5 and jun_actual) else val
            dt.append((row, col, display))

        dt.append((row, YTD_COL, jun_actual if jun_actual else budgets[0]))
        row += 1

    rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=AMBER, bold=True)))
    dt.append((row, LABEL_COL, "Total flex"))
    dt.append((row, DEFAULT_COL, total_flex_default))
    for m in range(12):
        dt.append((row, mcol(m), sum(b[m] for _, b, _ in FLEX)))
    dt.append((row, YTD_COL, total_flex_actual))
    row += 2  # +spacer

    # ── BUFFER ────────────────────────────────────────────────────────────────
    rq.append(mrg(sid, row, 0, row+1, NCOLS))
    rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=BLUE_DARK, bold=True, fg=WHITE)))
    dt.append((row, 0, "  BUFFER")); row += 1

    total_out_default = total_tx_default + total_fixed + total_flex_default
    for m in range(12):
        income = SALARY["Andreas"][m] + SALARY["Mona"][m]
        out    = total_tx_default + total_fixed + sum(b[m] for _, b, _ in FLEX)
        buf    = income - out
        bg     = GREEN if buf >= 0 else RED_LIGHT
        rq.append(rpt(sid, row, mcol(m), 1, 1, fmt(bg=bg, bold=True, halign="CENTER")))
        dt.append((row, mcol(m), buf))

    buf_default = total_income[0] - total_out_default
    ytd_buf     = sum((SALARY["Andreas"][m] + SALARY["Mona"][m]) -
                      (total_tx_default + total_fixed + sum(b[m] for _, b, _ in FLEX))
                      for m in range(6))
    rq.append(rpt(sid, row, 0, 1, 3, fmt(bold=True)))
    dt.append((row, LABEL_COL, "Buffer (income − all out)"))
    dt.append((row, DEFAULT_COL, buf_default))
    dt.append((row, YTD_COL, ytd_buf))
    row += 1

    rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=YELLOW, italic=True)))
    dt.append((row, LABEL_COL, "  ↳ Extra / discretionary"))
    row += 1

    return rq, dt


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    creds   = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds)
    sid     = SHEET_ID

    print(f"Building: https://docs.google.com/spreadsheets/d/{sid}/edit")

    info = service.spreadsheets().get(spreadsheetId=sid).execute()
    for s in info["sheets"]:
        if s["properties"]["title"] == "Yearly 2026":
            service.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [
                {"deleteSheet": {"sheetId": s["properties"]["sheetId"]}}]}).execute()

    result = service.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [
        {"addSheet": {"properties": {"title": "Yearly 2026"}}}]}).execute()
    sheet_sid = result["replies"][0]["addSheet"]["properties"]["sheetId"]

    rq, dt = build_yearly(sheet_sid)

    service.spreadsheets().batchUpdate(
        spreadsheetId=sid, body={"requests": rq}).execute()

    data_payload = [{"range": to_a1(r, c), "values": [[v]]} for r, c, v in dt]
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=sid,
        body={"valueInputOption": "USER_ENTERED", "data": data_payload}).execute()

    # Number format: #,##0 on all value columns
    nok = {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}
    service.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [
        {"repeatCell": {
            "range": {"sheetId": sheet_sid, "startRowIndex": 3, "endRowIndex": 80,
                      "startColumnIndex": DEFAULT_COL, "endColumnIndex": YTD_COL+1},
            "cell": {"userEnteredFormat": nok},
            "fields": "userEnteredFormat.numberFormat"}}
    ]}).execute()

    print("\nDone — 'Yearly 2026' sheet:")
    print("  Row 2: month status  (✓ Closed / Open / —)")
    print("  Transfers: account numbers in column B")
    print("  Flex: orange = budget changed this month; grey = future carry-forward")
    print("  (Groceries +1k in Aug; Entertainment +500 Jul-Sep)")


if __name__ == "__main__":
    main()
