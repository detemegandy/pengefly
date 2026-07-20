"""
PROTOTYPE — yearly overview sheet layout v3

Changes from v2:
  - MONTH_STATE: Jun → closed, Jul → open (current date is July 2026)
  - Month status row is now an editable dropdown (✓ Closed / Open / —)
    The credit card tracker reads from this row via XLOOKUP to determine
    which months are closed.
  - Added CREDIT CARDS section (EIKA and Mona statement balances)
    Sheet rows: EIKA = row 54, Mona = row 55 (0-indexed 53/54)
    Jun column = col I (index 8 = JAN_COL+5)

Sections (rows) × 12 months (columns):
  STATUS       — per-month close state (Closed / Open / future), now editable dropdown
  INCOME       — salary entered per month; June spikes with feriepenger
  TRANSFERS    — carry-forward amounts to savings buckets with account numbers;
                 status row per month (→ green when sent)
  FIXED        — carry-forward fixed expenses
  FLEX         — carry-forward budget vs actuals; orange cell = first month a budget changed;
                 grey = future carry-forward; no highlight = inherited unchanged
  BUFFER       — income − transfers − fixed − flex = what's left
  CREDIT CARDS — statement balance per card per month (EIKA row 54, Mona row 55)

Run: uv run prototypes/yearly_overview.py
Run all tabs at once: uv run prototypes/run_all.py
THROWAWAY — do not merge to main.
"""

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

CREDS_FILE = "/Users/andreas.saltveit/.config/gcp/google-sheets-sa.json"
SCOPES     = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
SHEET_ID   = "1gE5wc6lXaAOsiQ_dpjg7uGBnUrtGr0-oy-56UKyTdQc"
TAB_NAME   = "Yearly 2026"

MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

# ── private config (real account numbers and salary) ─────────────────────────
# Copy prototypes/private.py.example → prototypes/private.py and fill in values.
# private.py is gitignored — it never gets committed.
import os as _os
try:
    import importlib.util as _iu
    _s = _iu.spec_from_file_location(
        "_priv", _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "private.py"))
    _priv = _iu.module_from_spec(_s); _s.loader.exec_module(_priv)
    _ACCT    = _priv.ACCOUNT_NUMBERS
    SALARY   = _priv.SALARY
    del _iu, _s, _priv
except (ImportError, FileNotFoundError, AttributeError):
    _ACCT  = {}
    SALARY = {
        "Andreas": [40000,40000,40000,40000,40000,65000,40000,40000,40000,40000,40000,40000],
        "Mona":    [45000,45000,45000,45000,45000,90000,45000,45000,45000,45000,45000,45000],
    }
del _os

def _acct(name):
    return _ACCT.get(name, "####.##.#####")

# ── sample data ────────────────────────────────────────────────────────────────
# (label, amount, account_number)
TRANSFERS = [
    ("General savings",    2000, _acct("Saving-General")),
    ("Vacation savings",   2000, _acct("Saving-Vacation")),
    ("House savings",      1000, _acct("Saving-House")),
    ("Pets fund",           500, _acct("Pets Fund")),
    ("SOS fund",            500, _acct("SOS")),
    ("Liana",              1000, _acct("Liana")),
    ("Baby",               1000, _acct("Baby")),
    ("Maman (debt)",       2000, "—"),
    ("Andreas personal",   5000, "—"),
    ("Mona personal",      5000, "—"),
    ("Extra loan paydown", 1000, _acct("House & Ins.")),
]

FIXED = [
    ("Loan + insurance", 35105),
    ("Electricity",       4000),
    ("Internet",           500),
    ("Gym",                900),
    ("Barnehage",         3000),
]

# (name, budgets[12], actuals[12])
FLEX = [
    ("Groceries",
     [13000,13000,13000,13000,13000,13000,13000,14000,14000,14000,14000,14000],
     [11200,13800,12400,13100,14200, 4622,  None, None, None, None, None, None]),
    ("Entertainment",
     [1000,1000,1000,1000,1000,1000,1500,1500,1500,1000,1000,1000],
     [  820, 1240,  680, 1100,  890, 3591,  None, None, None, None, None, None]),
    ("Shopping",
     [3000]*12,
     [ 2800, 1900, 4200, 2400, 3100, 1156,  None, None, None, None, None, None]),
    ("Other",
     [3000]*12,
     [  450,  220, 1800,  680,  340,  237,  None, None, None, None, None, None]),
    ("Media",
     [288]*12,
     [  288,  288,  288,  288,  288,    0,  None, None, None, None, None, None]),
    ("Liana",
     [2000]*12,
     [ 1800, 3200,  800, 2200, 1600, 2158,  None, None, None, None, None, None]),
    ("Pets",
     [1500]*12,
     [    0,  850,    0,    0, 1200,  850,  None, None, None, None, None, None]),
]

# SENT[transfer_index][month_index]
SENT = [
    [2000]*12,
    [2000]*12,
    [1000]*12,
    [ 500]*12,
    [ 500, 500, 200, 500, 500, 500, 500, 500, 500, 500, 500, 500],
    [1000,1000, 600,1000,1000,1000,1000,1000,1000,1000,1000,1000],
    [1000]*12,
    [2000]*12,
    [5483]*12,
    [5500]*12,
    [1000]*12,
]

# Month close state — Jun is now closed (current month = Jul 2026)
MONTH_STATE = ["closed","closed","closed","closed","closed","closed",
               "open","future","future","future","future","future"]

# Credit card statement balances per month (None = no statement yet)
CC_BALANCES = {
    "EIKA — statement balance":    [0,0,0,0,0,45000,None,None,None,None,None,None],
    "Mona — statement balance":    [0,0,0,0,0,28000,None,None,None,None,None,None],
}


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
ORANGE      = rgb(255, 178, 102)
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

# Col layout: A=label | B=account | C=default | D-O=Jan-Dec | P=YTD | (gap) | R=legend
LABEL_COL   = 0
ACCT_COL    = 1
DEFAULT_COL = 2
JAN_COL     = 3   # months at cols 3-14
YTD_COL     = 15
NCOLS       = 16
LEGEND_COL  = 17  # one blank gap col at 16

def mcol(m):
    """0-indexed month → column index."""
    return JAN_COL + m

def cletter(col):
    """Column index → A1 letter string (e.g. 3 → 'D')."""
    s = ""
    while True:
        s = chr(ord('A') + col % 26) + s
        col = col // 26 - 1
        if col < 0:
            break
    return s

def cell_ref(row_0idx, col_0idx):
    """Bare A1 ref (no sheet prefix) for intra-sheet formula strings."""
    return f"{cletter(col_0idx)}{row_0idx+1}"

def to_a1(row, col):
    return f"'{TAB_NAME}'!{cletter(col)}{row+1}"


# ── conditional format helpers ─────────────────────────────────────────────────
# These replace static Python-applied colors for cells whose value the user can edit.

def _changed_cfrule(sid, row):
    """
    ORANGE when a month's value differs from the cell immediately to its left.
    Applied to the full Jan-Dec range; relative refs shift correctly:
      - Jan (D): compares to DEFAULT_COL (C) = the planned default
      - Feb-Dec: compares to the previous month
    """
    sr       = row + 1
    jan_ltr  = cletter(JAN_COL)      # 'D'
    prev_ltr = cletter(DEFAULT_COL)  # 'C'
    return {"addConditionalFormatRule": {"rule": {
        "ranges": [{"sheetId": sid, "startRowIndex": row, "endRowIndex": row+1,
                    "startColumnIndex": JAN_COL, "endColumnIndex": JAN_COL+12}],
        "booleanRule": {
            "condition": {"type": "CUSTOM_FORMULA",
                          "values": [{"userEnteredValue": f"={jan_ltr}{sr}<>{prev_ltr}{sr}"}]},
            "format": {"backgroundColor": ORANGE},
        }}, "index": 0}}

def _state_cfrule(sid, row, status_str, bg, fg=None):
    """
    Highlight when the month's status-row cell (D3:O3) equals status_str.
    OFFSET($D$3, 0, COLUMN()-4) walks across the status row aligned to the month columns.
    """
    stat    = JAN_COL + 1   # 1-indexed column number of $D$3 (Jan status cell)
    formula = f'=OFFSET($D$3,0,COLUMN()-{stat})="{status_str}"'
    f       = {"backgroundColor": bg}
    if fg:
        f["textFormat"] = {"foregroundColor": fg}
    return {"addConditionalFormatRule": {"rule": {
        "ranges": [{"sheetId": sid, "startRowIndex": row, "endRowIndex": row+1,
                    "startColumnIndex": JAN_COL, "endColumnIndex": JAN_COL+12}],
        "booleanRule": {
            "condition": {"type": "CUSTOM_FORMULA",
                          "values": [{"userEnteredValue": formula}]},
            "format": f,
        }}, "index": 0}}

def tx_cfrules(sid, row):
    """
    CF rules for a TRANSFERS data row.
    Add in reverse priority order (each uses index=0, so last-added = highest priority):
      GREEN (lowest) → AMBER → GREY → ORANGE (highest)
    """
    return [
        _state_cfrule(sid, row, "✓ Closed", GREEN, GREEN_DARK),
        _state_cfrule(sid, row, "Open",     AMBER),
        _state_cfrule(sid, row, "—",        GREY,  GREY_TEXT),
        _changed_cfrule(sid, row),
    ]

def flex_bud_cfrules(sid, row):
    """
    CF rules for a FLEX budget row.
    GREY (lower) → ORANGE (higher priority).
    """
    return [
        _state_cfrule(sid, row, "—", GREY, GREY_TEXT),
        _changed_cfrule(sid, row),
    ]


def add_legend(sid, rq, dt, tx_row, fixed_row, flex_row):
    """Color legend anchored to section header rows."""
    def entry(r, text, bg, fg=None, bold=False):
        rq.append(rpt(sid, r, LEGEND_COL, 1, 1, fmt(bg=bg, bold=bold, fg=fg)))
        if text:
            dt.append((r, LEGEND_COL, f"  {text}"))

    entry(tx_row,   "— TRANSFERS —",                                   BLUE_LIGHT, bold=True)
    entry(tx_row+1, "Sent as planned  (closed, matches carry-forward)", GREEN,     GREEN_DARK)
    entry(tx_row+2, "Current month — plan not yet confirmed",           AMBER)
    entry(tx_row+3, "Changed from previous month  (change or revert)",  ORANGE)
    entry(tx_row+4, "Future month — carry-forward plan",                GREY,      GREY_TEXT)

    entry(fixed_row,   "— FIXED EXPENSES —",                           BLUE_LIGHT, bold=True)
    entry(fixed_row+1, "Plan amounts only — budget vs actual coming",   WHITE)

    entry(flex_row,   "— FLEX BUDGET  (Budget row = plan; ↳ actual row = spend) —", BLUE_LIGHT, bold=True)
    entry(flex_row+1, "Budget row: changed from previous month",         ORANGE)
    entry(flex_row+2, "Budget row: future month carry-forward plan",     GREY,      GREY_TEXT)
    entry(flex_row+3, "Actual < 80 % of budget  (well under)",          GREEN,     GREEN_DARK)
    entry(flex_row+4, "Actual 80–100 % of budget  (on track, close)",   AMBER)
    entry(flex_row+5, "Actual > 100 % of budget  (over!)",              RED_LIGHT)


def build_yearly(sid):
    rq, dt = [], []

    # Column widths
    rq.append(cw(sid, LABEL_COL,   185))
    rq.append(cw(sid, ACCT_COL,    115))
    rq.append(cw(sid, DEFAULT_COL,  80))
    for c in range(JAN_COL, JAN_COL+12):
        rq.append(cw(sid, c, 68))
    rq.append(cw(sid, YTD_COL, 80))
    rq.append(cw(sid, LEGEND_COL,  310))

    # ── Row 0: title ──────────────────────────────────────────────────────────
    rq.append(mrg(sid, 0, 0, 1, NCOLS))
    rq.append(rpt(sid, 0, 0, 1, NCOLS,
                  fmt(bg=BLUE_DARK, bold=True, fg=WHITE, size=13, halign="CENTER")))
    dt.append((0, 0, "2026 — Annual Overview"))

    # ── Row 1: column headers ─────────────────────────────────────────────────
    rq.append(rpt(sid, 1, 0, 1, NCOLS, fmt(bg=BLUE_LIGHT, bold=True, halign="CENTER")))
    dt.append((1, LABEL_COL, ""))
    dt.append((1, ACCT_COL, "Account"))
    dt.append((1, DEFAULT_COL, "Budget"))
    for i, m in enumerate(MONTHS):
        state = MONTH_STATE[i]
        bg = BLUE_LIGHT if state == "closed" else (BLUE_PALE if state == "open" else GREY)
        fg = WHITE if state == "closed" else (None if state == "open" else GREY_TEXT)
        rq.append(rpt(sid, 1, mcol(i), 1, 1, fmt(bg=bg, bold=True, halign="CENTER", fg=fg)))
        dt.append((1, mcol(i), m))
    dt.append((1, YTD_COL, "YTD"))

    # ── Row 2: month close status — editable dropdown ─────────────────────────
    # Credit card tracker reads: XLOOKUP("✓ Closed", 'Yearly 2026'!D3:O3, months, ..., -1)
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

    # Status cells are editable dropdowns — user marks months closed here
    rq.append({"setDataValidation": {
        "range": {"sheetId": sid, "startRowIndex": 2, "endRowIndex": 3,
                  "startColumnIndex": JAN_COL, "endColumnIndex": JAN_COL+12},
        "rule": {"condition": {
            "type": "ONE_OF_LIST",
            "values": [{"userEnteredValue": "✓ Closed"},
                       {"userEnteredValue": "Open"},
                       {"userEnteredValue": "—"}]},
        "showCustomUi": True, "strict": True}}})

    rq.append(frz(sid, rows=3, cols=0))

    row = 3  # current writing row

    # ── INCOME ────────────────────────────────────────────────────────────────
    rq.append(mrg(sid, row, 0, row+1, NCOLS))
    rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=BLUE_DARK, bold=True, fg=WHITE)))
    dt.append((row, 0, "  INCOME")); row += 1

    income_start_row = row
    for i, (name, monthly) in enumerate(SALARY.items()):
        bg = BLUE_PALE if i%2==0 else WHITE
        rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=bg)))
        dt.append((row, LABEL_COL, name))
        dt.append((row, DEFAULT_COL, monthly[0]))
        for m, v in enumerate(monthly):
            rq.append(rpt(sid, row, mcol(m), 1, 1,
                          fmt(bg=bg, halign="CENTER")))
            prev_v = monthly[m-1] if m > 0 else None
            if m == 0 or v != prev_v:
                dt.append((row, mcol(m), v))
            else:
                dt.append((row, mcol(m), f"={cell_ref(row, mcol(m-1))}"))
        row += 1
    income_end_row = row - 1

    total_income = [SALARY["Andreas"][m] + SALARY["Mona"][m] for m in range(12)]
    rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=AMBER, bold=True)))
    dt.append((row, LABEL_COL, "Total income"))
    dt.append((row, DEFAULT_COL, total_income[0]))
    for m in range(12):
        c = cletter(mcol(m))
        dt.append((row, mcol(m), f"=SUM({c}{income_start_row+1}:{c}{income_end_row+1})"))
    income_total_row = row
    dt.append((row, YTD_COL, f"=SUM(D{row+1}:{cletter(mcol(5))}{row+1})"))
    row += 2  # +spacer

    # ── TRANSFERS ─────────────────────────────────────────────────────────────
    tx_header_row = row
    rq.append(mrg(sid, row, 0, row+1, NCOLS))
    rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=BLUE_DARK, bold=True, fg=WHITE)))
    dt.append((row, 0, "  TRANSFERS  (from Joint Salary)")); row += 1

    tx_start_row = row
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
            rq.append(rpt(sid, row, mcol(m), 1, 1, fmt(bg=bg, halign="RIGHT")))
            if m == 0 or val != prev:
                dt.append((row, mcol(m), val))
            else:
                dt.append((row, mcol(m), f"={cell_ref(row, mcol(m-1))}"))
            prev = val
        rq.extend(tx_cfrules(sid, row))
        row += 1
    tx_end_row = row - 1

    total_tx_default = sum(a for _, a, _ in TRANSFERS)
    rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=AMBER, bold=True)))
    dt.append((row, LABEL_COL, "Total transfers"))
    dt.append((row, DEFAULT_COL, total_tx_default))
    for m in range(12):
        c = cletter(mcol(m))
        dt.append((row, mcol(m), f"=SUM({c}{tx_start_row+1}:{c}{tx_end_row+1})"))
    tx_total_row = row
    dt.append((row, YTD_COL, f"=SUM(D{row+1}:{cletter(mcol(5))}{row+1})"))
    row += 2  # +spacer

    # ── FIXED EXPENSES ────────────────────────────────────────────────────────
    fixed_header_row = row
    rq.append(mrg(sid, row, 0, row+1, NCOLS))
    rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=BLUE_DARK, bold=True, fg=WHITE)))
    dt.append((row, 0, "  FIXED EXPENSES")); row += 1

    total_fixed = sum(a for _, a in FIXED)
    fixed_start_row = row
    for i, (name, amount) in enumerate(FIXED):
        bg = BLUE_PALE if i%2==0 else WHITE
        rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=bg)))
        dt.append((row, LABEL_COL, name))
        dt.append((row, DEFAULT_COL, amount))
        for m in range(12):
            rq.append(rpt(sid, row, mcol(m), 1, 1, fmt(bg=bg, halign="CENTER")))
            dt.append((row, mcol(m), amount))
        row += 1
    fixed_end_row = row - 1

    rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=AMBER, bold=True)))
    dt.append((row, LABEL_COL, "Total fixed"))
    dt.append((row, DEFAULT_COL, total_fixed))
    for m in range(12):
        c = cletter(mcol(m))
        dt.append((row, mcol(m), f"=SUM({c}{fixed_start_row+1}:{c}{fixed_end_row+1})"))
    fixed_total_row = row
    dt.append((row, YTD_COL, f"=SUM(D{row+1}:{cletter(mcol(5))}{row+1})"))
    row += 2  # +spacer

    # ── FLEX BUDGET ───────────────────────────────────────────────────────────
    flex_header_row = row
    rq.append(mrg(sid, row, 0, row+1, NCOLS))
    rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=BLUE_DARK, bold=True, fg=WHITE)))
    dt.append((row, 0, "  FLEX BUDGET")); row += 1

    total_flex_budget = sum(budgets[0] for _, budgets, _ in FLEX)
    total_flex_actual = sum(
        sum(a for a in actuals if a is not None) for _, _, actuals in FLEX
    )
    flex_budget_rows = []   # 0-indexed row numbers of per-category budget rows
    flex_actual_rows = []   # 0-indexed row numbers of per-category actual rows

    for i, (name, budgets, actuals) in enumerate(FLEX):
        bg_base = BLUE_PALE if i%2==0 else WHITE

        # Budget row
        flex_budget_rows.append(row)
        rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=bg_base, bold=True)))
        dt.append((row, LABEL_COL, name))
        dt.append((row, DEFAULT_COL, budgets[0]))
        for m in range(12):
            budget = budgets[m]
            prev_b = budgets[m-1] if m > 0 else budgets[0]
            rq.append(rpt(sid, row, mcol(m), 1, 1, fmt(bg=bg_base, bold=True, halign="CENTER")))
            if m == 0 or budget != prev_b:
                dt.append((row, mcol(m), budget))
            else:
                dt.append((row, mcol(m), f"={cell_ref(row, mcol(m-1))}"))
        rq.extend(flex_bud_cfrules(sid, row))
        dt.append((row, YTD_COL, sum(budgets[:6])))
        row += 1

        # Actual row
        flex_actual_rows.append(row)
        rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=bg_base, italic=True)))
        dt.append((row, LABEL_COL, "  ↳ actual"))
        for m in range(12):
            actual = actuals[m]
            budget = budgets[m]
            if actual is None:
                cell_bg, cell_fg = GREY, GREY_TEXT
            else:
                pct = actual / budget if budget else 0
                if pct < 0.8:
                    cell_bg, cell_fg = GREEN, GREEN_DARK
                elif pct <= 1.0:
                    cell_bg, cell_fg = AMBER, None
                else:
                    cell_bg, cell_fg = RED_LIGHT, None
            f = fmt(bg=cell_bg, italic=True, halign="CENTER", fg=cell_fg)
            rq.append(rpt(sid, row, mcol(m), 1, 1, f))
            if actual is not None:
                dt.append((row, mcol(m), actual))
        ytd_actual = sum(a for a in actuals if a is not None)
        dt.append((row, YTD_COL, ytd_actual))
        row += 1

    rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=AMBER, bold=True)))
    dt.append((row, LABEL_COL, "Total flex budget"))
    dt.append((row, DEFAULT_COL, total_flex_budget))
    for m in range(12):
        c    = cletter(mcol(m))
        refs = ",".join(f"{c}{r+1}" for r in flex_budget_rows)
        dt.append((row, mcol(m), f"=SUM({refs})"))
    flex_budget_total_row = row
    dt.append((row, YTD_COL, f"=SUM(D{row+1}:{cletter(mcol(5))}{row+1})"))
    row += 1

    rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=AMBER, italic=True)))
    dt.append((row, LABEL_COL, "  ↳ total actual"))
    for m in range(12):
        c    = cletter(mcol(m))
        refs = ",".join(f"{c}{r+1}" for r in flex_actual_rows)
        dt.append((row, mcol(m), f"=SUM({refs})"))
    dt.append((row, YTD_COL, f"=SUM(D{row+1}:{cletter(mcol(5))}{row+1})"))
    row += 2  # +spacer

    # ── BUFFER ────────────────────────────────────────────────────────────────
    rq.append(mrg(sid, row, 0, row+1, NCOLS))
    rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=BLUE_DARK, bold=True, fg=WHITE)))
    dt.append((row, 0, "  BUFFER")); row += 1

    buf_sr = row + 1  # 1-indexed sheet row for the buffer cells
    rq.append(rpt(sid, row, JAN_COL, 1, 12, fmt(bg=GREEN, bold=True, halign="CENTER")))
    rq.append({"addConditionalFormatRule": {"rule": {
        "ranges": [{"sheetId": sid, "startRowIndex": row, "endRowIndex": row+1,
                    "startColumnIndex": JAN_COL, "endColumnIndex": JAN_COL+12}],
        "booleanRule": {
            "condition": {"type": "CUSTOM_FORMULA",
                          "values": [{"userEnteredValue": f"=D{buf_sr}<0"}]},
            "format": {"backgroundColor": RED_LIGHT},
        }}, "index": 0}})
    for m in range(12):
        c = cletter(mcol(m))
        dt.append((row, mcol(m),
                   f"={c}{income_total_row+1}"
                   f"-{c}{tx_total_row+1}"
                   f"-{c}{fixed_total_row+1}"
                   f"-{c}{flex_budget_total_row+1}"))

    buf_default = total_income[0] - (total_tx_default + total_fixed + total_flex_budget)
    rq.append(rpt(sid, row, 0, 1, 3, fmt(bold=True)))
    dt.append((row, LABEL_COL, "Buffer (income − all out)"))
    dt.append((row, DEFAULT_COL, buf_default))
    dt.append((row, YTD_COL, f"=SUM(D{buf_sr}:{cletter(mcol(5))}{buf_sr})"))
    row += 1

    rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=YELLOW, italic=True)))
    dt.append((row, LABEL_COL, "  ↳ Extra / discretionary"))
    row += 1

    # ── CREDIT CARDS ──────────────────────────────────────────────────────────
    # The credit card tracker (Credit Card Tracker tab) reads the status row
    # via XLOOKUP to show "Closed through". It reads statement balances from
    # the rows below for future cross-sheet linking.
    #
    # Sheet rows at current row counter (0-indexed → 1-indexed):
    #   CC header:  row   → sheet row row+1
    #   EIKA:       row+1 → sheet row row+2
    #   Mona:       row+2 → sheet row row+3
    row += 1   # spacer
    rq.append(mrg(sid, row, 0, row+1, NCOLS))
    rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=BLUE_DARK, bold=True, fg=WHITE)))
    dt.append((row, 0, "  CREDIT CARDS  (statement balance at month close)")); row += 1

    cc_names = list(CC_BALANCES.keys())
    for i, (name, balances) in enumerate(CC_BALANCES.items()):
        bg = BLUE_PALE if i%2==0 else WHITE
        rq.append(rpt(sid, row, 0, 1, NCOLS, fmt(bg=bg)))
        dt.append((row, LABEL_COL, name))
        for m, val in enumerate(balances):
            state = MONTH_STATE[m]
            if val is None:
                rq.append(rpt(sid, row, mcol(m), 1, 1, fmt(bg=GREY, fg=GREY_TEXT)))
            else:
                cell_bg = (GREEN if state == "closed" and val == 0
                           else AMBER if state == "closed"   # has a balance when closed
                           else YELLOW if state == "open"
                           else GREY)
                cell_fg = GREEN_DARK if state == "closed" and val == 0 else None
                rq.append(rpt(sid, row, mcol(m), 1, 1,
                              fmt(bg=cell_bg, fg=cell_fg, halign="RIGHT")))
                if val:
                    dt.append((row, mcol(m), val))
        row += 1

    print(f"  CC rows (0-indexed): {row-len(cc_names)}–{row-1}  (sheet rows {row-len(cc_names)+1}–{row})")
    print(f"  EIKA at sheet row {row-1}, Mona at sheet row {row}")

    add_legend(sid, rq, dt, tx_header_row, fixed_header_row, flex_header_row)

    return rq, dt


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    creds   = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds)
    sid     = SHEET_ID

    print(f"Building: https://docs.google.com/spreadsheets/d/{sid}/edit")

    # Only delete/recreate "Yearly 2026" — other tabs are preserved
    info = service.spreadsheets().get(spreadsheetId=sid).execute()
    for s in info["sheets"]:
        if s["properties"]["title"] == TAB_NAME:
            service.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [
                {"deleteSheet": {"sheetId": s["properties"]["sheetId"]}}]}).execute()

    result = service.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [
        {"addSheet": {"properties": {"title": TAB_NAME}}}]}).execute()
    sheet_sid = result["replies"][0]["addSheet"]["properties"]["sheetId"]

    rq, dt = build_yearly(sheet_sid)

    service.spreadsheets().batchUpdate(
        spreadsheetId=sid, body={"requests": rq}).execute()

    data_payload = [{"range": to_a1(r, c), "values": [[v]]} for r, c, v in dt]
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=sid,
        body={"valueInputOption": "USER_ENTERED", "data": data_payload}).execute()

    nok = {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}
    service.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [
        {"repeatCell": {
            "range": {"sheetId": sheet_sid, "startRowIndex": 3, "endRowIndex": 100,
                      "startColumnIndex": DEFAULT_COL, "endColumnIndex": YTD_COL+1},
            "cell": {"userEnteredFormat": nok},
            "fields": "userEnteredFormat.numberFormat"}}
    ]}).execute()

    print(f"\nDone — '{TAB_NAME}' (other tabs preserved)")
    print("  Row 3: month status dropdowns (✓ Closed / Open / —) — user editable")
    print("  Credit card tracker reads D3:O3 via XLOOKUP to get Closed Through")


if __name__ == "__main__":
    main()
