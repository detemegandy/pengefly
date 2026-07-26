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
TOTAL_BUDGET = sum(BDGETS.values())

COLS_PER = 3
CAT_COLS = [i * COLS_PER for i in range(N)]
COL_W_PX = 100   # all category cols fixed at 100px

# Data zone: col W (0-idx 22) = labels, col X (0-idx 23) = values
DZ_L, DZ_V = 22, 23

# ── Colours ───────────────────────────────────────────────────────────────
def rgb(r, g, b): return {"red": r/255, "green": g/255, "blue": b/255}

WHITE      = rgb(255, 255, 255)
BLUE_HDR   = rgb(68,  114, 196)   # #4472C4 — same blue as spent arcs
LIGHT_BLUE = rgb(235, 242, 252)   # #EBF2FC — very light blue card body
MID_BLUE   = rgb(189, 215, 238)   # #BDD7EE — transaction header
BORDER_CLR = rgb(155, 175, 215)   # medium blue-grey card border


def solid_border():
    return {"style": "SOLID", "colorStyle": {"rgbColor": BORDER_CLR}}


def col_letter(c0):
    s, n = "", c0 + 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def spent_bare(cat):
    return (f"SUMPRODUCT(('{TRANS}'!E{T1}:E{T2}=\"{cat}\")"
            f"*'{TRANS}'!D{T1}:D{T2})")


def sparkline_fx(cat):
    sfx = spent_bare(cat)
    b = BDGETS[cat]
    return (
        f'=SPARKLINE(MIN({sfx}/{b},1),'
        f'{{"charttype","bar";"max",1;'
        f'"color1",IF({sfx}>{b},"#E53935","#4472C4");'
        f'"color2","#EBF2FC"}})'   # remaining fill matches card bg
    )


def over_fx(spent_cell, budget):
    return (
        f'=IF({spent_cell}>{budget},'
        f'"⚠ Over: +NOK "&TEXT({spent_cell}-{budget},"#,##0"),"")'
    )


def filter_fx(cat):
    return (
        f'=IFERROR(FILTER(\'{TRANS}\'!B{T1}:D{T2},'
        f'\'{TRANS}\'!E{T1}:E{T2}="{cat}"),'
        f'{{"—","No transactions",0}})'
    )


# ── Data zone row layout (1-indexed) ──────────────────────────────────────
CAT_1   = 1    # rows 1–7: category labels + SUMPRODUCT spents
SURP_1  = 8    # surplus
HLTH_1  = 10   # health score
VERD_1  = 11   # verdict
DONUT_1 = 13   # donut data rows 13–26: [Spent, Remaining] × 7 cats

# Category section 0-indexed rows (API rowIndex)
CAT_HDR   = 21   # row 22: category name (blue card header)
CAT_BUD   = 22   # row 23: budget label
CAT_SPT   = 23   # row 24: spent amount
CAT_BAR   = 24   # row 25: SPARKLINE bar
DNT_ANC   = 25   # row 26: donut anchor
DNT_H_PX  = 220
DNT_W_PX  = 200
OVER_ROW  = 37   # row 38: over-budget indicator
TRANS_HDR = 38   # row 39: transaction list header
TRANS_DAT = 39   # row 40: FILTER formula

# Overview area rows (left of the top pie chart, cols A–G)
OV_TITLE_R  = 0    # row 1: title
OV_MONTH_R  = 2    # row 3: month label
OV_VERD_R   = 4    # row 5: verdict value
OV_SPENT_R  = 7    # row 8: total spent
OV_BUDGET_R = 9    # row 10: budget line


def _restore_theme(svc):
    """Restore spreadsheet theme to Google Sheets defaults (ACCENT2=orange).
    Reverts the earlier ACCENT2=white change that was causing text to disappear.

    Per-slice pie colors are NOT settable via the Sheets API v4 ('slices' field
    returns HTTP 400). To make remaining/surplus arcs white: right-click each
    orange arc → Color → white (7 donuts + 1 overall surplus = 8 clicks total).
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
                            {"colorType": "ACCENT1",    "color": c(68,  114, 196)},
                            {"colorType": "ACCENT2",    "color": c(237, 125, 49)},
                            {"colorType": "ACCENT3",    "color": c(165, 165, 165)},
                            {"colorType": "ACCENT4",    "color": c(255, 192, 0)},
                            {"colorType": "ACCENT5",    "color": c(91,  155, 213)},
                            {"colorType": "ACCENT6",    "color": c(112, 173, 71)},
                            {"colorType": "LINK",       "color": c(17,  85,  204)},
                        ]
                    }
                },
                "fields": "spreadsheetTheme"
            }
        }]}
    ).execute()


def _set_column_widths(svc, sid):
    svc.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={"requests": [
            # Category cols A–U: 100px each
            {"updateDimensionProperties": {
                "range": {"sheetId": sid, "dimension": "COLUMNS",
                          "startIndex": 0, "endIndex": N * COLS_PER},
                "properties": {"pixelSize": COL_W_PX},
                "fields": "pixelSize",
            }},
            # Data zone cols W–X: 1px (hidden, still readable by charts)
            {"updateDimensionProperties": {
                "range": {"sheetId": sid, "dimension": "COLUMNS",
                          "startIndex": DZ_L, "endIndex": DZ_V + 1},
                "properties": {"pixelSize": 1},
                "fields": "pixelSize",
            }},
            # Category header rows: taller
            {"updateDimensionProperties": {
                "range": {"sheetId": sid, "dimension": "ROWS",
                          "startIndex": CAT_HDR, "endIndex": CAT_HDR + 1},
                "properties": {"pixelSize": 30},
                "fields": "pixelSize",
            }},
            # Overview title rows: taller
            {"updateDimensionProperties": {
                "range": {"sheetId": sid, "dimension": "ROWS",
                          "startIndex": OV_TITLE_R, "endIndex": OV_TITLE_R + 2},
                "properties": {"pixelSize": 28},
                "fields": "pixelSize",
            }},
        ]}
    ).execute()


def _apply_styling(svc, sid):
    reqs = []

    # ── Overview area: left of the top pie (cols A–G, rows 1–15) ──────────
    # Title bar rows 1–2
    reqs.append({"repeatCell": {
        "range": {"sheetId": sid,
                  "startRowIndex": OV_TITLE_R, "endRowIndex": OV_TITLE_R + 2,
                  "startColumnIndex": 0, "endColumnIndex": 7},
        "cell": {"userEnteredFormat": {
            "backgroundColor": BLUE_HDR,
            "textFormat": {"bold": True, "foregroundColor": WHITE, "fontSize": 13},
            "verticalAlignment": "MIDDLE",
            "padding": {"left": 10},
        }},
        "fields": "userEnteredFormat(backgroundColor,textFormat,verticalAlignment,padding)",
    }})

    # Info/health area rows 3–15
    reqs.append({"repeatCell": {
        "range": {"sheetId": sid,
                  "startRowIndex": OV_TITLE_R + 2, "endRowIndex": 15,
                  "startColumnIndex": 0, "endColumnIndex": 7},
        "cell": {"userEnteredFormat": {"backgroundColor": LIGHT_BLUE}},
        "fields": "userEnteredFormat.backgroundColor",
    }})

    # Label cells (rows OV_MONTH_R, OV_VERD_R, OV_SPENT_R, OV_BUDGET_R): italic label style
    for label_row in (OV_MONTH_R, OV_VERD_R, OV_SPENT_R, OV_BUDGET_R):
        reqs.append({"repeatCell": {
            "range": {"sheetId": sid,
                      "startRowIndex": label_row, "endRowIndex": label_row + 1,
                      "startColumnIndex": 0, "endColumnIndex": 7},
            "cell": {"userEnteredFormat": {
                "textFormat": {"italic": True,
                               "foregroundColor": rgb(100, 120, 160)},
                "padding": {"left": 10},
            }},
            "fields": "userEnteredFormat(textFormat,padding)",
        }})

    # Value cells one row below each label: larger, bold
    for val_row in (OV_MONTH_R + 1, OV_VERD_R + 1, OV_SPENT_R + 1, OV_BUDGET_R + 1):
        reqs.append({"repeatCell": {
            "range": {"sheetId": sid,
                      "startRowIndex": val_row, "endRowIndex": val_row + 1,
                      "startColumnIndex": 0, "endColumnIndex": 7},
            "cell": {"userEnteredFormat": {
                "textFormat": {"bold": True, "fontSize": 13},
                "padding": {"left": 10},
            }},
            "fields": "userEnteredFormat(textFormat,padding)",
        }})

    # ── Category cards ──────────────────────────────────────────────────────
    for i in range(N):
        cs, ce = CAT_COLS[i], CAT_COLS[i] + COLS_PER

        # Header row: blue bg, white bold text
        reqs.append({"repeatCell": {
            "range": {"sheetId": sid,
                      "startRowIndex": CAT_HDR, "endRowIndex": CAT_HDR + 1,
                      "startColumnIndex": cs, "endColumnIndex": ce},
            "cell": {"userEnteredFormat": {
                "backgroundColor": BLUE_HDR,
                "textFormat": {"bold": True, "foregroundColor": WHITE, "fontSize": 11},
                "verticalAlignment": "MIDDLE",
                "padding": {"left": 8},
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat,verticalAlignment,padding)",
        }})

        # Card body: budget through over-budget row — light blue bg
        reqs.append({"repeatCell": {
            "range": {"sheetId": sid,
                      "startRowIndex": CAT_BUD, "endRowIndex": OVER_ROW + 1,
                      "startColumnIndex": cs, "endColumnIndex": ce},
            "cell": {"userEnteredFormat": {"backgroundColor": LIGHT_BLUE}},
            "fields": "userEnteredFormat.backgroundColor",
        }})

        # Budget label: grey italic
        reqs.append({"repeatCell": {
            "range": {"sheetId": sid,
                      "startRowIndex": CAT_BUD, "endRowIndex": CAT_BUD + 1,
                      "startColumnIndex": cs, "endColumnIndex": ce},
            "cell": {"userEnteredFormat": {
                "textFormat": {"italic": True, "foregroundColor": rgb(100, 120, 160),
                               "fontSize": 9},
                "padding": {"left": 6, "top": 2},
            }},
            "fields": "userEnteredFormat(textFormat,padding)",
        }})

        # Spent amount: large bold number
        reqs.append({"repeatCell": {
            "range": {"sheetId": sid,
                      "startRowIndex": CAT_SPT, "endRowIndex": CAT_SPT + 1,
                      "startColumnIndex": cs, "endColumnIndex": ce},
            "cell": {"userEnteredFormat": {
                "numberFormat": {"type": "NUMBER", "pattern": "#,##0"},
                "textFormat": {"bold": True, "fontSize": 12},
                "padding": {"left": 6},
            }},
            "fields": "userEnteredFormat(numberFormat,textFormat,padding)",
        }})

        # Transaction header: mid-blue, bold
        reqs.append({"repeatCell": {
            "range": {"sheetId": sid,
                      "startRowIndex": TRANS_HDR, "endRowIndex": TRANS_HDR + 1,
                      "startColumnIndex": cs, "endColumnIndex": ce},
            "cell": {"userEnteredFormat": {
                "backgroundColor": MID_BLUE,
                "textFormat": {"bold": True, "fontSize": 9},
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat)",
        }})

        # Card outer border (header through transaction header)
        b = solid_border()
        reqs.append({"updateBorders": {
            "range": {"sheetId": sid,
                      "startRowIndex": CAT_HDR, "endRowIndex": TRANS_HDR + 1,
                      "startColumnIndex": cs, "endColumnIndex": ce},
            "top": b, "bottom": b, "left": b, "right": b,
        }})

        # Transaction data rows: background + left/right borders to stay visually
        # distinct from adjacent category blocks.
        # Alternating very-light-blue vs near-white so neighbours are distinct.
        tx_bg = LIGHT_BLUE if i % 2 == 0 else rgb(248, 250, 255)
        TRANS_END = TRANS_DAT + 80   # cover up to 80 transaction rows
        reqs.append({"repeatCell": {
            "range": {"sheetId": sid,
                      "startRowIndex": TRANS_DAT, "endRowIndex": TRANS_END,
                      "startColumnIndex": cs, "endColumnIndex": ce},
            "cell": {"userEnteredFormat": {
                "backgroundColor": tx_bg,
                "textFormat": {"fontSize": 9},
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat)",
        }})
        b_med = {"style": "SOLID_MEDIUM",
                 "colorStyle": {"rgbColor": BORDER_CLR}}
        reqs.append({"updateBorders": {
            "range": {"sheetId": sid,
                      "startRowIndex": TRANS_DAT, "endRowIndex": TRANS_END,
                      "startColumnIndex": cs, "endColumnIndex": ce},
            "left": b_med, "right": b_med,
        }})

    svc.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID, body={"requests": reqs}
    ).execute()


def main():
    creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    svc   = build("sheets", "v4", credentials=creds)

    _restore_theme(svc)
    print("  Theme: ACCENT2 restored to orange")

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
    print("  Column widths set")

    L = col_letter(DZ_L)   # "W"
    V = col_letter(DZ_V)   # "X"

    dt = []

    # ── Overview context cells (left of pie, cols A–G) ────────────────────
    dt += [
        {"range": f"'{TAB}'!A1",  "values": [["Monthly Overview"]]},
        {"range": f"'{TAB}'!A3",  "values": [["Financial Health"]]},
        {"range": f"'{TAB}'!A5",  "values": [[f"={V}{VERD_1}"]]},          # verdict
        {"range": f"'{TAB}'!A8",  "values": [["Total Spent"]]},
        {"range": f"'{TAB}'!A9",  "values": [[f"=SUM({V}{CAT_1}:{V}{CAT_1+N-1})"]]},
        {"range": f"'{TAB}'!A10", "values": [["Budget"]]},
        {"range": f"'{TAB}'!A11", "values": [[f"=TEXT({TOTAL_BUDGET},\"#,##0\")&\" NOK\""]]},
        {"range": f"'{TAB}'!A12", "values": [["Health Score"]]},
        {"range": f"'{TAB}'!A13", "values": [[f"={V}{HLTH_1}&\"%\""]]},
    ]

    # ── Data zone: category spents (W1:X7) ────────────────────────────────
    dt.append({"range": f"'{TAB}'!{L}{CAT_1}:{V}{CAT_1+N-1}",
               "values": [[cat, f"={spent_bare(cat)}"] for cat in CATS]})

    # ── Data zone: surplus (W8:X8) ────────────────────────────────────────
    dt.append({"range": f"'{TAB}'!{L}{SURP_1}:{V}{SURP_1}",
               "values": [["Surplus",
                            f"=MAX(0,{TOTAL_BUDGET}-SUM({V}{CAT_1}:{V}{CAT_1+N-1}))"]]})

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
        ref = f"{V}{CAT_1+i}"
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

    for i, cat in enumerate(CATS):
        dt.append({"range": f"'{TAB}'!{col_letter(CAT_COLS[i])}{TRANS_DAT+1}",
                   "values": [[filter_fx(cat)]]})

    svc.spreadsheets().values().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={"valueInputOption": "USER_ENTERED", "data": dt},
    ).execute()
    print("  Data written")

    _apply_styling(svc, sid)
    print("  Styling applied")

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

    # Overall pie: 7 categories + surplus = 8 slices with labels.
    # Anchored at col H (idx=7), 700px wide, 300px tall — centered over
    # the 7-category area (21 cols × 100px = 2100px; pie at 700–1400px = center).
    # Surplus arc will be orange (ACCENT2); right-click → Color → white for
    # the "missing piece" effect — per-slice color not settable via API.
    chart_reqs.append({"addChart": {"chart": {
        "spec": {
            "title": "",
            "backgroundColorStyle": {"rgbColor": WHITE},
            "pieChart": {
                "legendPosition": "LABELED_LEGEND",
                "threeDimensional": False,
                "domain": pie_src(CAT_1 - 1, SURP_1, DZ_L),   # W1:W8
                "series": pie_src(CAT_1 - 1, SURP_1, DZ_V),   # X1:X8
            },
        },
        "position": {"overlayPosition": {
            "anchorCell": {"sheetId": sid, "rowIndex": 0, "columnIndex": 7},
            "widthPixels": 700, "heightPixels": 300,
        }},
    }}})

    # Per-category donuts — centered within 3-col × 100px = 300px block.
    # Blue arc = Spent (ACCENT1), orange arc = Remaining (ACCENT2).
    donut_offset_x = (COL_W_PX * COLS_PER - DNT_W_PX) // 2   # 50px
    for i in range(N):
        r0 = DONUT_1 - 1 + i * 2   # 0-indexed data row
        chart_reqs.append({"addChart": {"chart": {
            "spec": {
                "title": "",
                "backgroundColorStyle": {"rgbColor": LIGHT_BLUE},  # matches card bg
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
    print(f"  Added 1 overview pie + {N} donut charts")

    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit"
    print(f"\nDone. Open: {url}")
    print()
    print("Remaining arcs are orange (API limitation — no per-slice color support).")
    print("For white arcs: right-click each orange arc → Color → white.")


if __name__ == "__main__":
    main()
