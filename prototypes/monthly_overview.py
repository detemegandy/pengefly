"""
PROTOTYPE — Monthly Overview sheet.

Layout:
  Top: overall budget pie (all categories + surplus pool as slices)
  Top-right: financial health indicator (adherence score + verdict)
  Below: per-category columns (horizontal), each with:
    - Category name, budget, spent
    - SPARKLINE progress bar
    - Donut chart (fills rotationally; fully blue when over budget)
    - Live FILTER transaction list below the chart

Health score = (categories within budget / total) × 100
Verdict: On Track (≥86%), Caution (≥57%), Over Budget (<57%)

Run: uv run prototypes/monthly_overview.py
THROWAWAY — do not merge to main.
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import transaction_entry as te

CREDS_FILE = te.CREDS_FILE
SCOPES     = te.SCOPES
SHEET_ID   = te.SHEET_ID

TAB   = "Monthly Overview"
TRANS = te.TAB_NAME       # "Card Transactions Jun 2026 — Detailed"
T1    = te.DATA_ROW       # 14 (first transaction row, 1-indexed)
T2    = T1 + 499

CATS         = te.CATEGORIES  # 7 category names
BDGETS       = te.BUDGETS     # {cat: NOK amount}
N            = len(CATS)
TOTAL_BUDGET = sum(BDGETS.values())  # 23 788

# Per-category donut: Google Sheets API v4 does not expose per-slice color for
# pieChart specs — the "Remaining" arc will be Google's default 2nd-accent
# color. Fix in the UI: right-click the grey/orange arc → Color → white.

COLS_PER = 3  # columns per category block: Date | Merchant | Amount
CAT_COLS = [i * COLS_PER for i in range(N)]  # 0-indexed start col per cat

# Data zone: labels in col W (0-idx=22), values in col X (0-idx=23)
DZ_L, DZ_V = 22, 23


def col_letter(c0):
    s, n = "", c0 + 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# Shared formula builders
def spent_fx(cat):
    return (f"=SUMPRODUCT(('{TRANS}'!E{T1}:E{T2}=\"{cat}\")"
            f"*'{TRANS}'!D{T1}:D{T2})")


def sparkline_fx(cat):
    sfx = (f"SUMPRODUCT(('{TRANS}'!E{T1}:E{T2}=\"{cat}\")"
           f"*'{TRANS}'!D{T1}:D{T2})")
    b = BDGETS[cat]
    return (
        f'=SPARKLINE(MIN({sfx}/{b},1),'
        f'{{"charttype","bar";"max",1;'
        f'"color1",IF({sfx}>{b},"#E53935","#4472C4");'
        f'"color2","#E8F0FE"}})'
    )


def over_fx(cat, spent_cell):
    b = BDGETS[cat]
    return (
        f'=IF({spent_cell}>{b},'
        f'"Over: +NOK "&TEXT({spent_cell}-{b},"#,##0"),'
        f'"")'
    )


def filter_fx(cat):
    return (
        f'=IFERROR(FILTER(\'{TRANS}\'!B{T1}:D{T2},'
        f'\'{TRANS}\'!E{T1}:E{T2}="{cat}"),'
        f'{{"—","No transactions",0}})'
    )


# ── Row layout (1-indexed for formula strings, 0-indexed for API rowIndex) ─
PIE_HDR_1  = 1           # data zone: pie chart header
PIE_DATA_1 = 2           # data zone rows 2-8: category spents
PIE_SURP_1 = PIE_DATA_1 + N   # data zone row 9: surplus pool
HLTH_1     = PIE_SURP_1 + 2   # data zone row 11: health score
VERD_1     = HLTH_1 + 1       # data zone row 12: verdict
DONUT_1    = VERD_1 + 2       # data zone row 14: donut data start (2 per cat)

# Category section 0-indexed (displayed below the overall pie)
CAT_HDR   = 21   # row 22: category name
CAT_BUD   = 22   # row 23: Budget label
CAT_SPT   = 23   # row 24: Spent (formula)
CAT_BAR   = 24   # row 25: SPARKLINE progress bar
DNT_ANC   = 25   # row 26: donut chart anchor
DNT_H_PX  = 220
DNT_W_PX  = 200
OVER_ROW  = 37   # row 38: over-budget indicator (0-indexed)
TRANS_HDR = 38   # row 39: transaction list header
TRANS_DAT = 39   # row 40: FILTER formula


def main():
    creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    svc   = build("sheets", "v4", credentials=creds)

    info = svc.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
    tabs = {s["properties"]["title"]: s["properties"]["sheetId"]
            for s in info["sheets"]}

    # Delete and recreate the tab for idempotent reruns
    reqs = []
    if TAB in tabs:
        reqs.append({"deleteSheet": {"sheetId": tabs[TAB]}})
    reqs.append({"addSheet": {"properties": {"title": TAB}}})
    resp = svc.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID, body={"requests": reqs}).execute()
    sid = resp["replies"][-1]["addSheet"]["properties"]["sheetId"]
    print(f"Created '{TAB}' (sid={sid})")

    L = col_letter(DZ_L)   # "W"
    V = col_letter(DZ_V)   # "X"

    dt = []  # value ranges for batchUpdate

    # ── Data zone: overall pie chart (W1:X8) — spending categories only ──
    # Surplus is intentionally excluded: the pie shows only what was spent,
    # no "empty" surplus slice needed.
    pie_rows = [["Category", "Spent"]]
    for cat in CATS:
        pie_rows.append([cat, spent_fx(cat)])
    dt.append({"range": f"'{TAB}'!{L}{PIE_HDR_1}:{V}{PIE_HDR_1+N}",
               "values": pie_rows})

    # ── Data zone: health indicator (W10:X11, shifted since no surplus row) ─
    # References already-written spent values in X2:X8
    conditions = "+".join(
        f"IF({V}{PIE_DATA_1+i}<={BDGETS[CATS[i]]},1,0)" for i in range(N)
    )
    health_fx = f"=({conditions})/{N}*100"
    verdict_fx = (
        f'=IF({V}{HLTH_1}>=86,"On Track",'
        f'IF({V}{HLTH_1}>=57,"Caution","Over Budget"))'
    )
    dt.append({"range": f"'{TAB}'!{L}{HLTH_1}:{V}{VERD_1}",
               "values": [["Health Score", health_fx],
                          ["Verdict",      verdict_fx]]})

    # ── Data zone: per-category donut data (W14:X28, 2 rows per cat) ─────
    donut_rows = []
    for cat in CATS:
        sfx_bare = (f"SUMPRODUCT(('{TRANS}'!E{T1}:E{T2}=\"{cat}\")"
                    f"*'{TRANS}'!D{T1}:D{T2})")
        donut_rows.append(["Spent",     f"={sfx_bare}"])
        donut_rows.append(["Remaining", f"=MAX(0,{BDGETS[cat]}-{sfx_bare})"])
    dt.append({"range": f"'{TAB}'!{L}{DONUT_1}:{V}{DONUT_1+N*2-1}",
               "values": donut_rows})

    # ── Category section rows ─────────────────────────────────────────────
    hdr_row  = []
    bud_row  = []
    spt_row  = []  # store formula per cat for cross-referencing over_fx
    bar_row  = []
    over_row = []
    th_row   = []

    # Spent cell references in data zone (X2, X3, ... X8)
    spent_cells = [f"{V}{PIE_DATA_1+i}" for i in range(N)]

    for i, cat in enumerate(CATS):
        hdr_row.extend([cat,                            "", ""])
        bud_row.extend([f"Budget: NOK {BDGETS[cat]:,}", "", ""])
        spt_row.extend([f"={spent_cells[i]}",           "", ""])
        bar_row.extend([sparkline_fx(cat),              "", ""])
        over_row.extend([over_fx(cat, spent_cells[i]),  "", ""])
        th_row.extend(["Date", "Merchant", "Amount"])

    last_col = col_letter(N * COLS_PER - 1)  # "U"
    for row_vals, row_1idx in [
        (hdr_row,  CAT_HDR  + 1),
        (bud_row,  CAT_BUD  + 1),
        (spt_row,  CAT_SPT  + 1),
        (bar_row,  CAT_BAR  + 1),
        (over_row, OVER_ROW + 1),
        (th_row,   TRANS_HDR + 1),
    ]:
        dt.append({"range": f"'{TAB}'!A{row_1idx}:{last_col}{row_1idx}",
                   "values": [row_vals]})

    # FILTER formulas go per-category (each at its own start column)
    for i, cat in enumerate(CATS):
        cl = col_letter(CAT_COLS[i])
        dt.append({"range": f"'{TAB}'!{cl}{TRANS_DAT+1}",
                   "values": [[filter_fx(cat)]]})

    svc.spreadsheets().values().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={"valueInputOption": "USER_ENTERED", "data": dt},
    ).execute()
    print("  Data written")

    # ── Charts ───────────────────────────────────────────────────────────
    def pie_src(r0_start, r0_end_excl, col0):
        return {"sourceRange": {"sources": [{
            "sheetId":          sid,
            "startRowIndex":    r0_start,
            "endRowIndex":      r0_end_excl,
            "startColumnIndex": col0,
            "endColumnIndex":   col0 + 1,
        }]}}

    chart_reqs = []

    # Overall pie — centered over the 21-column category block.
    # 21 cols × 100px default = 2100px. Chart width 700px →
    # left edge at (2100-700)/2 = 700px = column H (0-idx=7), offsetX=0.
    # Shows spending categories only; no surplus slice.
    # Data rows: PIE_DATA_1-1 (0-idx) to PIE_DATA_1-1+N (exclusive) = rows 1..8
    chart_reqs.append({"addChart": {"chart": {
        "spec": {
            "title": "Spending Breakdown — Jun 2026",
            "backgroundColorStyle": {"rgbColor": {"red": 1, "green": 1, "blue": 1}},
            "pieChart": {
                "legendPosition": "LABELED_LEGEND",
                "threeDimensional": False,
                "domain": pie_src(PIE_DATA_1 - 1, PIE_DATA_1 - 1 + N, DZ_L),
                "series": pie_src(PIE_DATA_1 - 1, PIE_DATA_1 - 1 + N, DZ_V),
            },
        },
        "position": {"overlayPosition": {
            "anchorCell": {"sheetId": sid, "rowIndex": 0, "columnIndex": 7},
            "widthPixels": 700, "heightPixels": 300,
        }},
    }}})

    # Per-category donut charts.
    # NOTE: Google Sheets API v4 does not expose per-slice colors for pieChart.
    # The "Remaining" arc gets the default 2nd-accent color.
    # To make it white/invisible: right-click the arc in the Sheets UI → Color → white.
    for i, cat in enumerate(CATS):
        r0 = DONUT_1 - 1 + i * 2  # 0-indexed start row
        chart_reqs.append({"addChart": {"chart": {
            "spec": {
                "title": "",   # title shown in cells above, not on chart
                "backgroundColorStyle": {"rgbColor": {"red": 1, "green": 1, "blue": 1}},
                "pieChart": {
                    "legendPosition": "NO_LEGEND",
                    "pieHole": 0.5,
                    "threeDimensional": False,
                    "domain": pie_src(r0, r0 + 2, DZ_L),
                    "series": pie_src(r0, r0 + 2, DZ_V),
                },
            },
            "position": {"overlayPosition": {
                "anchorCell": {"sheetId": sid,
                               "rowIndex": DNT_ANC,
                               "columnIndex": CAT_COLS[i]},
                "widthPixels": DNT_W_PX, "heightPixels": DNT_H_PX,
            }},
        }}})

    svc.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID, body={"requests": chart_reqs}).execute()
    print(f"  Added overall pie + {N} donut charts")

    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit"
    print(f"\nDone. Open: {url}")


if __name__ == "__main__":
    main()
