"""
PROTOTYPE — credit card tracker v4
Issue #5: What does the credit card tracker look like?

Bug fixes from v3:
  - Schedule IF formulas referenced B13/B14 instead of B14/B15 (off by one)
  - Google Sheets auto-converts "Jul 2026" to a date serial; replaced
    MATCH(sim_month, text_array) with VLOOKUP(TEXT(sim_month,"MMM YYYY"), schedule, 2, 0)
  - Conditional format simplified to check warning cell text

Pending (needs yearly overview tab):
  - "Closed through" and "Statement balance" will become XLOOKUP formulas pointing
    to 'Yearly 2026' once run_all.py creates all tabs together

Columns: A-E = EIKA (Andreas) | F = gap | G-K = DNB (Mona)

Sheet row refs (EIKA / Mona):
  Closed-through:    B4 / H4   Statement balance: B5 / H5
  Credit limit:      B6 / H6   Remaining credit:  B7 / H7  (=B6-B5)
  Annual rate:       B8 / H8   Min payment %:     B9 / H9
  Min floor (NOK):   B10 / H10  Payment due:      B11 / H11
  Sim month:         B14 / H14  Sim amount:       B15 / H15
  Schedule (opening bal): B20:B25 / H20:H25  (Jul–Dec 2026)

Run standalone:  uv run prototypes/credit_card_tracker.py
Run with all tabs: uv run prototypes/run_all.py
THROWAWAY — do not merge to main.
"""

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

CREDS_FILE = "/Users/andreas.saltveit/.config/gcp/google-sheets-sa.json"
SCOPES     = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
SHEET_ID   = "1gE5wc6lXaAOsiQ_dpjg7uGBnUrtGr0-oy-56UKyTdQc"
TAB_NAME   = "Credit Card Tracker"

NCOLS         = 11
LAST_CLOSED   = "Jun 2026"   # used as Python string in schedule header label
OPEN_MONTHS   = ["Jul 2026", "Aug 2026", "Sep 2026",
                 "Oct 2026", "Nov 2026", "Dec 2026"]

# Formula for the "Closed through" cells (B4 / H4).
# Reads the status row from Yearly 2026 and returns the last month marked "✓ Closed".
# IFERROR falls back to LAST_CLOSED so standalone runs still show a sensible value.
_ALL_MONTHS = ["Jan 2026","Feb 2026","Mar 2026","Apr 2026","May 2026","Jun 2026",
               "Jul 2026","Aug 2026","Sep 2026","Oct 2026","Nov 2026","Dec 2026"]
_months_arr = '{"' + '","'.join(_ALL_MONTHS) + '"}'
CLOSED_THROUGH_FX = (
    f'=IFERROR(XLOOKUP("✓ Closed",\'Yearly 2026\'!D3:O3,{_months_arr}'
    f',"{LAST_CLOSED}",0,-1),"{LAST_CLOSED}")'
)
DATA_ROW      = 20   # 1-indexed sheet row for Jul 2026
SCHED_END_ROW = DATA_ROW + len(OPEN_MONTHS) - 1   # 25

# Schedule column ranges for VLOOKUP (col 1 = month name, col 2 = opening balance)
SCHED_E = f"A{DATA_ROW}:E{SCHED_END_ROW}"   # EIKA
SCHED_M = f"G{DATA_ROW}:K{SCHED_END_ROW}"   # Mona


# ── helpers ────────────────────────────────────────────────────────────────────

def rgb(r, g, b):
    return {"red": r/255, "green": g/255, "blue": b/255}

BLUE_DARK   = rgb(68, 114, 196)
BLUE_LIGHT  = rgb(180, 198, 231)
BLUE_PALE   = rgb(221, 235, 246)
AMBER       = rgb(255, 192, 0)
WHITE       = rgb(255, 255, 255)
YELLOW_SOFT = rgb(255, 217, 102)
GREEN_SOFT  = rgb(198, 239, 206)
RED_SOFT    = rgb(255, 199, 206)
COMPUTED_BG = rgb(232, 244, 230)   # faint green = do not edit, computed

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

def rh(sid, row, px):
    return {"updateDimensionProperties": {
        "range": {"sheetId": sid, "dimension": "ROWS",
                  "startIndex": row, "endIndex": row+1},
        "properties": {"pixelSize": px}, "fields": "pixelSize"}}

def frz(sid, rows=0, cols=0):
    return {"updateSheetProperties": {
        "properties": {"sheetId": sid,
                       "gridProperties": {"frozenRowCount": rows, "frozenColumnCount": cols}},
        "fields": "gridProperties(frozenRowCount,frozenColumnCount)"}}

def nok_fmt(sid, r, c, nrows, ncols):
    return {"repeatCell": {
        "range": {"sheetId": sid, "startRowIndex": r, "endRowIndex": r+nrows,
                  "startColumnIndex": c, "endColumnIndex": c+ncols},
        "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}},
        "fields": "userEnteredFormat.numberFormat"}}

def pct_fmt(sid, r, c, nrows, ncols):
    return {"repeatCell": {
        "range": {"sheetId": sid, "startRowIndex": r, "endRowIndex": r+nrows,
                  "startColumnIndex": c, "endColumnIndex": c+ncols},
        "cell": {"userEnteredFormat": {"numberFormat": {"type": "PERCENT", "pattern": "0.0%"}}},
        "fields": "userEnteredFormat.numberFormat"}}

def cfmt(sid, r, c1, c2, formula, color, idx):
    return {"addConditionalFormatRule": {"rule": {
        "ranges": [{"sheetId": sid, "startRowIndex": r, "endRowIndex": r+1,
                    "startColumnIndex": c1, "endColumnIndex": c2}],
        "booleanRule": {
            "condition": {"type": "CUSTOM_FORMULA",
                          "values": [{"userEnteredValue": formula}]},
            "format": {"backgroundColor": color},
        }}, "index": idx}}

def ddv(sid, r, col, opts):
    """Strict dropdown: only listed values accepted."""
    return {"setDataValidation": {
        "range": {"sheetId": sid, "startRowIndex": r, "endRowIndex": r+1,
                  "startColumnIndex": col, "endColumnIndex": col+1},
        "rule": {"condition": {"type": "ONE_OF_LIST",
                               "values": [{"userEnteredValue": o} for o in opts]},
                 "showCustomUi": True, "strict": True}}}

def warn_formula(sched_rng, lim, sim_m, sim_a):
    """
    Warning: compares (opening balance at sim month + expense) against credit limit.
    Uses VLOOKUP(TEXT(date_serial,"MMM YYYY"), ...) because Google Sheets
    auto-converts "Jul 2026" dropdown selections to date serials.
    IFERROR guard keeps it clean if the date format is unexpected.
    """
    vlookup = f'IFERROR(VLOOKUP(TEXT({sim_m},"MMM YYYY"),{sched_rng},2,0),0)'
    return (
        f'=IF({vlookup}+{sim_a}>{lim},'
        f'"⚠  Over limit by "&TEXT({vlookup}+{sim_a}-{lim},"#,##0")&" NOK",'
        f'"✓  "&TEXT({lim}-{vlookup}-{sim_a},"#,##0")&" NOK remaining after expense")'
    )


# ── build Credit Card Tracker sheet ───────────────────────────────────────────

def build_tracker(sid):
    rq, dt = [], []

    for i, px in enumerate([200, 115, 90, 90, 115,  18,  200, 115, 90, 90, 115]):
        rq.append(cw(sid, i, px))

    # ── Row 1: title ──────────────────────────────────────────────────────────
    rq.append(mrg(sid, 0, 0, 1, NCOLS))
    rq.append(rpt(sid, 0, 0, 1, NCOLS,
                  fmt(bg=BLUE_DARK, bold=True, fg=WHITE, size=13, halign="CENTER")))
    dt.append({"range": "A1", "values": [["Credit Card Tracker"]]})

    rq.append(rh(sid, 1, 8))   # row 2 spacer

    # ── Row 3: card subheaders ────────────────────────────────────────────────
    rq.append(mrg(sid, 2, 0, 3, 5));  rq.append(mrg(sid, 2, 6, 3, NCOLS))
    rq.append(rpt(sid, 2, 0, 1, 5, fmt(bg=BLUE_LIGHT, bold=True, halign="CENTER")))
    rq.append(rpt(sid, 2, 6, 1, 5, fmt(bg=BLUE_LIGHT, bold=True, halign="CENTER")))
    dt.append({"range": "A3", "values": [["EIKA — Andreas"]]})
    dt.append({"range": "G3", "values": [["DNB — Mona"]]})

    # ── Rows 4-11: setup block ─────────────────────────────────────────────────
    # editable=False → faint green tint (computed cell, do not edit manually)
    setup = [
        ("Closed through",            CLOSED_THROUGH_FX, CLOSED_THROUGH_FX, False),
        ("Statement balance (NOK)",    45_000,         28_000,          True),
        ("Credit limit (NOK)",         50_000,         30_000,          True),
        ("Remaining credit (NOK)",     "=B6-B5",       "=H6-H5",        False),
        ("Annual interest rate",       0.22,           0.20,            True),
        ("Min payment %",              0.03,           0.03,            True),
        ("Min payment floor (NOK)",    250,            250,             True),
        ("Payment due",           "10th of month", "25th of month",    True),
    ]
    for i, (label, ev, mv, editable) in enumerate(setup):
        r     = i + 3
        shade = BLUE_PALE if i % 2 == 0 else WHITE
        vbg   = shade if editable else COMPUTED_BG
        rq.append(rpt(sid, r, 0, 1, 1, fmt(bg=shade, italic=True)))
        rq.append(rpt(sid, r, 1, 1, 1, fmt(bg=vbg,   bold=True, halign="RIGHT")))
        rq.append(rpt(sid, r, 6, 1, 1, fmt(bg=shade, italic=True)))
        rq.append(rpt(sid, r, 7, 1, 1, fmt(bg=vbg,   bold=True, halign="RIGHT")))
        dt.append({"range": f"A{r+1}:B{r+1}", "values": [[label, ev]]})
        dt.append({"range": f"G{r+1}:H{r+1}", "values": [[label, mv]]})

    # % format: rate (0-idx 7 = B8/H8) and min% (0-idx 8 = B9/H9)
    for ri in [7, 8]:
        rq.append(pct_fmt(sid, ri, 1, 1, 1));  rq.append(pct_fmt(sid, ri, 7, 1, 1))

    rq.append(rh(sid, 11, 8))   # row 12 spacer

    # ── Row 13: simulation header ─────────────────────────────────────────────
    rq.append(mrg(sid, 12, 0, 13, 5));  rq.append(mrg(sid, 12, 6, 13, NCOLS))
    rq.append(rpt(sid, 12, 0, 1, 5, fmt(bg=YELLOW_SOFT, bold=True, halign="CENTER")))
    rq.append(rpt(sid, 12, 6, 1, 5, fmt(bg=YELLOW_SOFT, bold=True, halign="CENTER")))
    dt.append({"range": "A13", "values": [["Simulate: add an expense"]]})
    dt.append({"range": "G13", "values": [["Simulate: add an expense"]]})

    # ── Rows 14-15: simulation inputs ─────────────────────────────────────────
    # B14/H14 = expense month (dropdown — open months only)
    # B15/H15 = expense amount (NOK)
    sim_inputs = [
        ("Add expense in month",  OPEN_MONTHS[0],  OPEN_MONTHS[0]),
        ("Expense amount (NOK)",  5_000,           3_000),
    ]
    for i, (label, ev, mv) in enumerate(sim_inputs):
        r     = i + 13
        shade = BLUE_PALE if i % 2 == 0 else WHITE
        rq.append(rpt(sid, r, 0, 1, 1, fmt(bg=shade, italic=True)))
        rq.append(rpt(sid, r, 1, 1, 1, fmt(bg=shade, bold=True, halign="RIGHT")))
        rq.append(rpt(sid, r, 6, 1, 1, fmt(bg=shade, italic=True)))
        rq.append(rpt(sid, r, 7, 1, 1, fmt(bg=shade, bold=True, halign="RIGHT")))
        dt.append({"range": f"A{r+1}:B{r+1}", "values": [[label, ev]]})
        dt.append({"range": f"G{r+1}:H{r+1}", "values": [[label, mv]]})

    # Strict dropdown on B14 / H14 — only open months accepted
    rq.append(ddv(sid, 13, 1, OPEN_MONTHS))   # B14
    rq.append(ddv(sid, 13, 7, OPEN_MONTHS))   # H14
    rq.append(nok_fmt(sid, 14, 1, 1, 1))       # B15 NOK format
    rq.append(nok_fmt(sid, 14, 7, 1, 1))       # H15 NOK format

    # ── Row 16: over-limit warning banner ─────────────────────────────────────
    # Uses VLOOKUP(TEXT(date_serial,"MMM YYYY"), schedule, 2, 0) to find the
    # opening balance at the sim month. Google Sheets stores dropdown dates as
    # serials (e.g. "Jul 2026" → 46204), so TEXT() converts back to the label
    # that VLOOKUP can match against column A of the schedule.
    eika_warn = warn_formula(SCHED_E, "$B$6", "$B$14", "$B$15")
    mona_warn = warn_formula(SCHED_M, "$H$6", "$H$14", "$H$15")

    rq.append(mrg(sid, 15, 0, 16, 5));  rq.append(mrg(sid, 15, 6, 16, NCOLS))
    rq.append(rpt(sid, 15, 0, 1, 5, fmt(bold=True, halign="CENTER")))
    rq.append(rpt(sid, 15, 6, 1, 5, fmt(bold=True, halign="CENTER")))
    dt.append({"range": "A16", "values": [[eika_warn]]})
    dt.append({"range": "G16", "values": [[mona_warn]]})

    # Conditional format: check the first char of the warning cell text
    rq.append(cfmt(sid, 15, 0, 5,  '=LEFT($A$16,1)="⚠"', RED_SOFT,   idx=0))
    rq.append(cfmt(sid, 15, 0, 5,  '=LEFT($A$16,1)="✓"', GREEN_SOFT, idx=1))
    rq.append(cfmt(sid, 15, 6, 11, '=LEFT($G$16,1)="⚠"', RED_SOFT,   idx=2))
    rq.append(cfmt(sid, 15, 6, 11, '=LEFT($G$16,1)="✓"', GREEN_SOFT, idx=3))

    rq.append(rh(sid, 16, 8))   # row 17 spacer

    # ── Row 18: schedule header ───────────────────────────────────────────────
    rq.append(mrg(sid, 17, 0, 18, 5));  rq.append(mrg(sid, 17, 6, 18, NCOLS))
    rq.append(rpt(sid, 17, 0, 1, 5, fmt(bg=BLUE_LIGHT, bold=True, halign="CENTER")))
    rq.append(rpt(sid, 17, 6, 1, 5, fmt(bg=BLUE_LIGHT, bold=True, halign="CENTER")))
    sched_label = f"Jul–Dec 2026 — minimum payments only (closed through {LAST_CLOSED})"
    dt.append({"range": "A18", "values": [[sched_label]]})
    dt.append({"range": "G18", "values": [[sched_label]]})

    # ── Row 19: column headers; freeze through here ───────────────────────────
    col_hdrs = ["Month", "Opening bal", "Interest", "Payment", "Closing bal"]
    rq.append(rpt(sid, 18, 0, 1, 5, fmt(bg=BLUE_DARK, bold=True, fg=WHITE, halign="CENTER")))
    rq.append(rpt(sid, 18, 6, 1, 5, fmt(bg=BLUE_DARK, bold=True, fg=WHITE, halign="CENTER")))
    dt.append({"range": "A19:E19", "values": [col_hdrs]})
    dt.append({"range": "G19:K19", "values": [col_hdrs]})
    rq.append(frz(sid, rows=19))

    # ── Rows 20-25: open-month schedule ───────────────────────────────────────
    # Opening balance:
    #   Jul (i=0): B5 (statement balance) + sim amount if MONTH(B14) = 7
    #   Aug (i=1): closing balance of Jul  + sim amount if MONTH(B14) = 8
    #   etc.
    # Uses MONTH($B$14) so the formula works whether B14 holds a text date or
    # date serial — MONTH() handles both.
    # Rate=B8, Min%=B9, Floor=B10. Sim month=B14, Sim amount=B15.
    for i, month in enumerate(OPEN_MONTHS):
        r  = i + 19
        sr = r + 1
        month_num = 7 + i   # Jul=7 … Dec=12
        shade = BLUE_PALE if i % 2 == 0 else WHITE
        rq.append(rpt(sid, r, 0, 1, 5, fmt(bg=shade)))
        rq.append(rpt(sid, r, 6, 1, 5, fmt(bg=shade)))

        if i == 0:
            eb = f'=B5+IF(MONTH($B$14)={month_num},$B$15,0)'
            mb = f'=H5+IF(MONTH($H$14)={month_num},$H$15,0)'
        else:
            prev = sr - 1
            eb = f'=E{prev}+IF(MONTH($B$14)={month_num},$B$15,0)'
            mb = f'=K{prev}+IF(MONTH($H$14)={month_num},$H$15,0)'

        ec = f'=IF(B{sr}>0,ROUND(B{sr}*$B$8/12,0),0)'
        ed = f'=IF(B{sr}=0,0,MIN(B{sr}+C{sr},MAX(ROUND((B{sr}+C{sr})*$B$9,0),$B$10)))'
        ee = f'=MAX(0,B{sr}+C{sr}-D{sr})'

        mc = f'=IF(H{sr}>0,ROUND(H{sr}*$H$8/12,0),0)'
        md = f'=IF(H{sr}=0,0,MIN(H{sr}+I{sr},MAX(ROUND((H{sr}+I{sr})*$H$9,0),$H$10)))'
        me = f'=MAX(0,H{sr}+I{sr}-J{sr})'

        dt.append({"range": f"A{sr}:E{sr}", "values": [[month, eb, ec, ed, ee]]})
        dt.append({"range": f"G{sr}:K{sr}", "values": [[month, mb, mc, md, me]]})

    # Green row when fully paid off
    for i in range(len(OPEN_MONTHS)):
        r  = i + 19;  sr = r + 1
        rq.append(cfmt(sid, r, 0, 5,  f"=E{sr}=0", GREEN_SOFT, idx=0))
        rq.append(cfmt(sid, r, 6, 11, f"=K{sr}=0", GREEN_SOFT, idx=0))

    # ── Total row ─────────────────────────────────────────────────────────────
    r_tot  = len(OPEN_MONTHS) + 19
    sr_tot = r_tot + 1
    rq.append(rpt(sid, r_tot, 0, 1, 5, fmt(bg=AMBER, bold=True)))
    rq.append(rpt(sid, r_tot, 6, 1, 5, fmt(bg=AMBER, bold=True)))
    dt.append({"range": f"A{sr_tot}:E{sr_tot}", "values": [[
        f"Total ({OPEN_MONTHS[0][:3]}–{OPEN_MONTHS[-1][:3]} 2026)",
        "", f"=SUM(C{DATA_ROW}:C{SCHED_END_ROW})",
            f"=SUM(D{DATA_ROW}:D{SCHED_END_ROW})", ""]]})
    dt.append({"range": f"G{sr_tot}:K{sr_tot}", "values": [[
        f"Total ({OPEN_MONTHS[0][:3]}–{OPEN_MONTHS[-1][:3]} 2026)",
        "", f"=SUM(I{DATA_ROW}:I{SCHED_END_ROW})",
            f"=SUM(J{DATA_ROW}:J{SCHED_END_ROW})", ""]]})

    return rq, dt


# ── sheet management ──────────────────────────────────────────────────────────

def recreate_tab(service, sid, title, index=None):
    """Delete the named tab if it exists, then create a fresh one. Returns new sheetId."""
    info = service.spreadsheets().get(spreadsheetId=sid).execute()
    to_del = [{"deleteSheet": {"sheetId": s["properties"]["sheetId"]}}
              for s in info["sheets"] if s["properties"]["title"] == title]
    if to_del:
        service.spreadsheets().batchUpdate(
            spreadsheetId=sid, body={"requests": to_del}).execute()
    props = {"title": title}
    if index is not None:
        props["index"] = index
    result = service.spreadsheets().batchUpdate(
        spreadsheetId=sid, body={"requests": [{"addSheet": {"properties": props}}]}
    ).execute()
    return result["replies"][0]["addSheet"]["properties"]["sheetId"]


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    creds   = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds)

    print(f"Building: https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit")

    keep = recreate_tab(service, SHEET_ID, TAB_NAME)

    rq, dt = build_tracker(keep)
    service.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID, body={"requests": rq}).execute()
    # Qualify every range with the tab name — bare ranges default to the first sheet,
    # which is "Yearly 2026" after run_all.py builds all tabs.
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={"valueInputOption": "USER_ENTERED",
              "data": [{"range": f"'{TAB_NAME}'!{d['range']}", "values": d["values"]} for d in dt]}
    ).execute()

    post = [
        nok_fmt(keep, DATA_ROW-1, 1, len(OPEN_MONTHS), 4),
        nok_fmt(keep, DATA_ROW-1, 7, len(OPEN_MONTHS), 4),
        nok_fmt(keep, DATA_ROW-1+len(OPEN_MONTHS), 2, 1, 2),
        nok_fmt(keep, DATA_ROW-1+len(OPEN_MONTHS), 8, 1, 2),
        nok_fmt(keep, 6, 1, 1, 1),   # remaining credit EIKA
        nok_fmt(keep, 6, 7, 1, 1),   # remaining credit Mona
    ]
    service.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID, body={"requests": post}).execute()

    print(f"Done — '{TAB_NAME}' (other tabs preserved)")


if __name__ == "__main__":
    main()
