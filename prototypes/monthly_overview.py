"""
PROTOTYPE — Monthly Overview sheet.

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
COL_W_PX = 100   # explicit column width so donut pixel positions are predictable

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
# Rows 1–7:  7 category labels + SUMPRODUCT spents (used for overall pie)
# Row  8:    "Surplus" | MAX(0, budget - SUM(X1:X7))
# Row  10:   Health score
# Row  11:   Verdict
# Rows 13–26: donut data [Spent, Remaining] per category (2 rows × 7 cats)
CAT_1   = 1
SURP_1  = 8
HLTH_1  = 10
VERD_1  = 11
DONUT_1 = 13

# Category section 0-indexed rows (API rowIndex)
CAT_HDR   = 21   # row 22: category name
CAT_BUD   = 22
CAT_SPT   = 23
CAT_BAR   = 24
DNT_ANC   = 25   # donut chart anchor
DNT_H_PX  = 220
DNT_W_PX  = 200
OVER_ROW  = 37
TRANS_HDR = 38
TRANS_DAT = 39


def _restore_theme(svc):
    """Restore spreadsheet theme to Google Sheets defaults.

    Reverts the earlier ACCENT2=white change that made some text invisible.

    Per-slice pie colors are NOT settable via the Sheets API v4 — the 'slices'
    field is rejected (confirmed HTTP 400). To make remaining/surplus arcs white:
    right-click each orange arc in the UI → Color → white. One-time manual step
    per sheet rebuild (7 donuts + 1 overall surplus = 8 clicks).
    """
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
                            {"colorType": "ACCENT1",    "color": c(68,  114, 196)},  # blue
                            {"colorType": "ACCENT2",    "color": c(237, 125, 49)},   # orange (restored)
                            {"colorType": "ACCENT3",    "color": c(165, 165, 165)},  # grey
                            {"colorType": "ACCENT4",    "color": c(255, 192, 0)},    # yellow
                            {"colorType": "ACCENT5",    "color": c(91,  155, 213)},  # light blue
                            {"colorType": "ACCENT6",    "color": c(112, 173, 71)},   # green
                            {"colorType": "LINK",       "color": c(17,  85,  204)},
                        ]
                    }
                },
                "fields": "spreadsheetTheme"
            }
        }]}
    ).execute()


def _set_column_widths(svc, sid):
    """Set cols A–U to COL_W_PX each so donut positions are deterministic."""
    svc.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={"requests": [{
            "updateDimensionProperties": {
                "range": {
                    "sheetId":    sid,
                    "dimension":  "COLUMNS",
                    "startIndex": 0,
                    "endIndex":   N * COLS_PER,   # 21 cols A-U
                },
                "properties": {"pixelSize": COL_W_PX},
                "fields": "pixelSize",
            }
        }]}
    ).execute()


def main():
    creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    svc   = build("sheets", "v4", credentials=creds)

    _restore_theme(svc)
    print("  Theme: restored ACCENT2 to default orange")

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

    _set_column_widths(svc, sid)
    print(f"  Set cols A-U to {COL_W_PX}px each")

    L = col_letter(DZ_L)   # "W"
    V = col_letter(DZ_V)   # "X"

    dt = []

    # ── Data zone: category spents (W1:X7) ────────────────────────────────
    dt.append({"range": f"'{TAB}'!{L}{CAT_1}:{V}{CAT_1+N-1}",
               "values": [[cat, f"={spent_bare(cat)}"] for cat in CATS]})

    # ── Data zone: surplus (W8:X8) ────────────────────────────────────────
    surplus_fx = f"=MAX(0,{TOTAL_BUDGET}-SUM({V}{CAT_1}:{V}{CAT_1+N-1}))"
    dt.append({"range": f"'{TAB}'!{L}{SURP_1}:{V}{SURP_1}",
               "values": [["Surplus", surplus_fx]]})

    # ── Data zone: health indicator (W10:X11) ─────────────────────────────
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

    # ── Data zone: per-category donut data (W13:X26) ──────────────────────
    donut_rows = []
    for i, cat in enumerate(CATS):
        ref = f"{V}{CAT_1+i}"   # X1, X2, … X7 already computed above
        donut_rows.append(["Spent",     f"={ref}"])
        donut_rows.append(["Remaining", f"=MAX(0,{BDGETS[cat]}-{ref})"])
    dt.append({"range": f"'{TAB}'!{L}{DONUT_1}:{V}{DONUT_1+N*2-1}",
               "values": donut_rows})

    # ── Visible category section rows ─────────────────────────────────────
    hdr_row, bud_row, spt_row, bar_row, ovr_row, thr_row = [], [], [], [], [], []
    for i, cat in enumerate(CATS):
        sc = f"{V}{CAT_1+i}"
        hdr_row.extend([cat,                             "", ""])
        bud_row.extend([f"Budget: NOK {BDGETS[cat]:,}", "", ""])
        spt_row.extend([f"={sc}",                        "", ""])
        bar_row.extend([sparkline_fx(cat),               "", ""])
        ovr_row.extend([over_fx(sc, BDGETS[cat]),        "", ""])
        thr_row.extend(["Date", "Merchant", "Amount"])

    last_col = col_letter(N * COLS_PER - 1)   # "U"
    for vals, row_1idx in [
        (hdr_row, CAT_HDR   + 1),
        (bud_row, CAT_BUD   + 1),
        (spt_row, CAT_SPT   + 1),
        (bar_row, CAT_BAR   + 1),
        (ovr_row, OVER_ROW  + 1),
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

    # Overall pie — 7 categories + surplus = 8 slices, labeled so names appear on arcs.
    # Domain/series: W1:W8 / X1:X8 (0-indexed rows 0–7).
    # Surplus arc will be ACCENT2=orange. To get the "missing piece" effect:
    # right-click that arc in the Sheets UI → Color → white (API won't do it).
    # Centered: 21 cols × 100px = 2100px; 700px pie anchored at col H (idx=7).
    chart_reqs.append({"addChart": {"chart": {
        "spec": {
            "title": "",
            "backgroundColorStyle": {"rgbColor": WHITE},
            "pieChart": {
                "legendPosition": "LABELED_LEGEND",
                "threeDimensional": False,
                "domain": pie_src(CAT_1 - 1, SURP_1, DZ_L),   # rows 0–7 (8 rows)
                "series": pie_src(CAT_1 - 1, SURP_1, DZ_V),
            },
        },
        "position": {"overlayPosition": {
            "anchorCell": {"sheetId": sid, "rowIndex": 0, "columnIndex": 7},
            "widthPixels": 700, "heightPixels": 300,
        }},
    }}})

    # Per-category donuts — centered within each 3-col × 100px = 300px block.
    # Blue arc = Spent (ACCENT1), orange arc = Remaining (ACCENT2).
    donut_offset_x = (COL_W_PX * COLS_PER - DNT_W_PX) // 2   # (300-200)//2 = 50px
    for i in range(N):
        r0 = DONUT_1 - 1 + i * 2   # 0-indexed start of this cat's 2-row donut block
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
                               "rowIndex":    DNT_ANC,
                               "columnIndex": CAT_COLS[i]},
                "offsetXPixels": donut_offset_x,
                "widthPixels":   DNT_W_PX,
                "heightPixels":  DNT_H_PX,
            }},
        }}})

    svc.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID, body={"requests": chart_reqs}).execute()
    print(f"  Added overall pie (8 slices, labeled) + {N} donut charts")

    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit"
    print(f"\nDone. Open: {url}")
    print()
    print("Remaining arcs are orange (ACCENT2 default — API cannot set per-slice colors).")
    print("For white arcs: right-click each orange arc → Color → white.")


if __name__ == "__main__":
    main()
