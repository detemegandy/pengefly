"""
PROTOTYPE — Monthly Overview sheet.

Per-slice color trick: set spreadsheet theme ACCENT1=blue, ACCENT2=white.
Both the overall pie and each donut put their data as [Spent, Remaining] /
[Spending, Surplus] — the blue (ACCENT1) slice is the arc, the white
(ACCENT2) slice vanishes into the background, giving the "missing piece"
effect without any per-slice color API call.

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
TRANS = te.TAB_NAME
T1    = te.DATA_ROW
T2    = T1 + 499

CATS         = te.CATEGORIES
BDGETS       = te.BUDGETS
N            = len(CATS)
TOTAL_BUDGET = sum(BDGETS.values())   # 23 788

COLS_PER = 3
CAT_COLS = [i * COLS_PER for i in range(N)]   # 0-indexed start col per cat

# Data zone: col W (0-idx 22) = labels, col X (0-idx 23) = values
DZ_L, DZ_V = 22, 23

WHITE = {"red": 1.0, "green": 1.0, "blue": 1.0}


def col_letter(c0):
    s, n = "", c0 + 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def spent_bare(cat):
    """SUMPRODUCT expression without leading '='"""
    return (f"SUMPRODUCT(('{TRANS}'!E{T1}:E{T2}=\"{cat}\")"
            f"*'{TRANS}'!D{T1}:D{T2})")


def sparkline_fx(cat):
    sfx = spent_bare(cat)
    b = BDGETS[cat]
    return (
        f'=SPARKLINE(MIN({sfx}/{b},1),'
        f'{{"charttype","bar";"max",1;'
        f'"color1",IF({sfx}>{b},"#E53935","#4472C4");'
        f'"color2","#E8F0FE"}})'
    )


def over_fx(spent_cell, budget):
    return (
        f'=IF({spent_cell}>{budget},'
        f'"Over: +NOK "&TEXT({spent_cell}-{budget},"#,##0"),"")'
    )


def filter_fx(cat):
    return (
        f'=IFERROR(FILTER(\'{TRANS}\'!B{T1}:D{T2},'
        f'\'{TRANS}\'!E{T1}:E{T2}="{cat}"),'
        f'{{"—","No transactions",0}})'
    )


# ── Data zone row layout (1-indexed for formula strings) ─────────────────
# Overall pie: 2 slices in order [Spending, Surplus]
#   Slice 1 (Spending) → ACCENT1 = blue
#   Slice 2 (Surplus)  → ACCENT2 = white (invisible = "missing piece")
PIE_SPND_1 = 1   # "Spending" | =SUM(cat spents)
PIE_SURP_1 = 2   # "Surplus"  | =MAX(0, TOTAL_BUDGET - spending)
CAT_1      = 3   # rows 3–9: category labels + their SUMPRODUCT spents
HLTH_1     = 11  # health score
VERD_1     = 12  # verdict
DONUT_1    = 14  # donut data start (2 rows per cat: [Spent, Remaining])

# Category section 0-indexed rows (API rowIndex)
CAT_HDR   = 21   # row 22: category name
CAT_BUD   = 22   # row 23: Budget label
CAT_SPT   = 23   # row 24: Spent (formula)
CAT_BAR   = 24   # row 25: SPARKLINE progress bar
DNT_ANC   = 25   # row 26: donut chart anchor
DNT_H_PX  = 220
DNT_W_PX  = 200
OVER_ROW  = 37   # row 38: over-budget indicator
TRANS_HDR = 38   # row 39: transaction list header
TRANS_DAT = 39   # row 40: FILTER formula


def _set_theme(svc):
    """Set ACCENT1=blue (spent arcs) and ACCENT2=white (remaining arcs).
    ACCENT1 is already blue by default; only ACCENT2 changes from orange."""
    def c(r, g, b):
        return {"rgbColor": {"red": r/255, "green": g/255, "blue": b/255}}
    svc.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={"requests": [{
            "updateSpreadsheetProperties": {
                "properties": {
                    "spreadsheetTheme": {
                        "primaryFontFamily": "Roboto",
                        "themeColors": [
                            {"colorType": "TEXT",       "color": c(0,   0,   0)},
                            {"colorType": "BACKGROUND", "color": c(255, 255, 255)},
                            {"colorType": "ACCENT1",    "color": c(68,  114, 196)},  # #4472C4 blue
                            {"colorType": "ACCENT2",    "color": c(255, 255, 255)},  # white — remaining/surplus
                            {"colorType": "ACCENT3",    "color": c(255, 192, 0)},    # amber
                            {"colorType": "ACCENT4",    "color": c(70,  163, 71)},   # green
                            {"colorType": "ACCENT5",    "color": c(63,  147, 199)},  # teal
                            {"colorType": "ACCENT6",    "color": c(146, 99,  183)},  # purple
                            {"colorType": "LINK",       "color": c(17,  84,  181)},
                        ]
                    }
                },
                "fields": "spreadsheetTheme"
            }
        }]}
    ).execute()


def main():
    creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    svc   = build("sheets", "v4", credentials=creds)

    _set_theme(svc)
    print("  Theme: ACCENT1=blue, ACCENT2=white")

    info = svc.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
    tabs = {s["properties"]["title"]: s["properties"]["sheetId"]
            for s in info["sheets"]}

    reqs = []
    if TAB in tabs:
        reqs.append({"deleteSheet": {"sheetId": tabs[TAB]}})
    reqs.append({"addSheet": {"properties": {"title": TAB}}})
    resp = svc.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID, body={"requests": reqs}).execute()
    sid = resp["replies"][-1]["addSheet"]["properties"]["sheetId"]
    print(f"  Created '{TAB}' (sid={sid})")

    L = col_letter(DZ_L)   # "W"
    V = col_letter(DZ_V)   # "X"

    dt = []

    # ── Data zone: overall pie (W1:X2) ────────────────────────────────────
    total_spent_fx = f"=SUM({V}{CAT_1}:{V}{CAT_1+N-1})"    # =SUM(X3:X9)
    surplus_fx     = f"=MAX(0,{TOTAL_BUDGET}-{V}{PIE_SPND_1})"  # =MAX(0,23788-X1)
    dt.append({"range": f"'{TAB}'!{L}{PIE_SPND_1}:{V}{PIE_SURP_1}",
               "values": [["Spending", total_spent_fx],
                          ["Surplus",  surplus_fx]]})

    # ── Data zone: category spents (W3:X9) ────────────────────────────────
    dt.append({"range": f"'{TAB}'!{L}{CAT_1}:{V}{CAT_1+N-1}",
               "values": [[cat, f"={spent_bare(cat)}"] for cat in CATS]})

    # ── Data zone: health indicator (W11:X12) ─────────────────────────────
    adherence = "+".join(
        f"IF({V}{CAT_1+i}<={BDGETS[CATS[i]]},1,0)" for i in range(N)
    )
    dt.append({"range": f"'{TAB}'!{L}{HLTH_1}:{V}{VERD_1}",
               "values": [
                   ["Health Score", f"=({adherence})/{N}*100"],
                   ["Verdict",
                    f'=IF({V}{HLTH_1}>=86,"On Track",'
                    f'IF({V}{HLTH_1}>=57,"Caution","Over Budget"))'],
               ]})

    # ── Data zone: per-category donut data (W14:X28) ──────────────────────
    # [Spent, Remaining] → Spent gets ACCENT1=blue, Remaining gets ACCENT2=white
    donut_rows = []
    for i, cat in enumerate(CATS):
        ref = f"{V}{CAT_1+i}"   # already-computed SUMPRODUCT in X3:X9
        donut_rows.append(["Spent",     f"={ref}"])
        donut_rows.append(["Remaining", f"=MAX(0,{BDGETS[cat]}-{ref})"])
    dt.append({"range": f"'{TAB}'!{L}{DONUT_1}:{V}{DONUT_1+N*2-1}",
               "values": donut_rows})

    # ── Visible category section rows ─────────────────────────────────────
    hdr_row = []
    bud_row = []
    spt_row = []
    bar_row = []
    ovr_row = []
    thr_row = []

    for i, cat in enumerate(CATS):
        sc = f"{V}{CAT_1+i}"   # spent cell reference (X3, X4, …)
        hdr_row.extend([cat,                              "", ""])
        bud_row.extend([f"Budget: NOK {BDGETS[cat]:,}",  "", ""])
        spt_row.extend([f"={sc}",                         "", ""])
        bar_row.extend([sparkline_fx(cat),                "", ""])
        ovr_row.extend([over_fx(sc, BDGETS[cat]),         "", ""])
        thr_row.extend(["Date", "Merchant", "Amount"])

    last_col = col_letter(N * COLS_PER - 1)   # "U"
    for vals, row_1idx in [
        (hdr_row, CAT_HDR  + 1),
        (bud_row, CAT_BUD  + 1),
        (spt_row, CAT_SPT  + 1),
        (bar_row, CAT_BAR  + 1),
        (ovr_row, OVER_ROW + 1),
        (thr_row, TRANS_HDR + 1),
    ]:
        dt.append({"range": f"'{TAB}'!A{row_1idx}:{last_col}{row_1idx}",
                   "values": [vals]})

    # FILTER formulas — one per category at its starting column
    for i, cat in enumerate(CATS):
        dt.append({"range": f"'{TAB}'!{col_letter(CAT_COLS[i])}{TRANS_DAT+1}",
                   "values": [[filter_fx(cat)]]})

    svc.spreadsheets().values().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={"valueInputOption": "USER_ENTERED", "data": dt},
    ).execute()
    print("  Data written")

    # ── Charts ────────────────────────────────────────────────────────────
    def pie_src(r0_start, r0_end_excl, col0):
        return {"sourceRange": {"sources": [{
            "sheetId":          sid,
            "startRowIndex":    r0_start,
            "endRowIndex":      r0_end_excl,
            "startColumnIndex": col0,
            "endColumnIndex":   col0 + 1,
        }]}}

    chart_reqs = []

    # Overall pie — centered over 21-col block (col H=idx7, 700px wide, 300px tall)
    # 2 slices: Spending (ACCENT1=blue) + Surplus (ACCENT2=white = "missing piece")
    chart_reqs.append({"addChart": {"chart": {
        "spec": {
            "title": "",
            "backgroundColorStyle": {"rgbColor": WHITE},
            "pieChart": {
                "legendPosition": "NO_LEGEND",
                "threeDimensional": False,
                "domain": pie_src(0, 2, DZ_L),   # W1:W2
                "series": pie_src(0, 2, DZ_V),   # X1:X2
            },
        },
        "position": {"overlayPosition": {
            "anchorCell": {"sheetId": sid, "rowIndex": 0, "columnIndex": 7},
            "widthPixels": 700, "heightPixels": 300,
        }},
    }}})

    # Per-category donuts
    # 2 slices each: Spent (ACCENT1=blue) + Remaining (ACCENT2=white = invisible arc)
    for i in range(N):
        r0 = DONUT_1 - 1 + i * 2   # 0-indexed start of this cat's donut data
        chart_reqs.append({"addChart": {"chart": {
            "spec": {
                "title": "",
                "backgroundColorStyle": {"rgbColor": WHITE},
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
