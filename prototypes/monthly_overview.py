"""
PROTOTYPE — Monthly Overview sheet.

Savings categories (Liana, Pets) are configured in private.py:
  - Ceiling = savings account balance (not monthly budget)
  - Only shown in the overview when they have transactions this month
  - Label shows "Balance" not "Budget"

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

CATS         = te.CATEGORIES           # all 7, fixed order
BDGETS       = te.BUDGETS
TOTAL_BUDGET = sum(BDGETS.values())

# ── Savings category configuration ───────────────────────────────────────
# private.py stores balances; if the file doesn't exist fall back gracefully
try:
    import private as _priv
    SAVINGS_BALANCES: dict[str, int] = _priv.SAVINGS_BALANCES
except (ImportError, AttributeError):
    SAVINGS_BALANCES = {}

SAVINGS_CATS  = set(SAVINGS_BALANCES.keys())
SPENDING_CATS = [c for c in CATS if c not in SAVINGS_CATS]
N_SPENDING    = len(SPENDING_CATS)

COLS_PER = 3
COL_W_PX = 100
PIE_W_PX = 700

# Data zone: col W (0-idx 22) = labels, col X (0-idx 23) = values (always fixed)
DZ_L, DZ_V = 22, 23

# Data zone row layout — 1-indexed.
# Rows 1..7    : all 7 cats in CATS order (spending first, savings last)
# Row  8       : surplus (TOTAL_BUDGET - SUM(X1:X7))
# Row  10      : health score (spending cats only)
# Row  11      : verdict
# Rows 13..26  : donut data — [Spent, Remaining] × 7 cats (2 rows each)
CAT_1   = 1
SURP_1  = 8
HLTH_1  = 10
VERD_1  = 11
DONUT_1 = 13

# Category section 0-indexed rows (API rowIndex)
CAT_HDR   = 21
CAT_BUD   = 22
CAT_SPT   = 23
CAT_BAR   = 24
DNT_ANC   = 25
DNT_H_PX  = 220
DNT_W_PX  = 200
OVER_ROW  = 37
TRANS_HDR = 38
TRANS_DAT = 39
TRANS_END = TRANS_DAT + 80

# Overview panel rows (left of top pie, cols A-G)
OV_TITLE_R = 0

# ── Colours ───────────────────────────────────────────────────────────────
def rgb(r, g, b): return {"red": r/255, "green": g/255, "blue": b/255}

WHITE      = rgb(255, 255, 255)
BLUE_HDR   = rgb(68,  114, 196)
LIGHT_BLUE = rgb(235, 242, 252)
MID_BLUE   = rgb(189, 215, 238)
BORDER_CLR = rgb(155, 175, 215)

def solid_border():
    return {"style": "SOLID", "colorStyle": {"rgbColor": BORDER_CLR}}


def col_letter(c0):
    s, n = "", c0 + 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# ── Per-category helpers ──────────────────────────────────────────────────
def cat_dz_row(cat):
    """1-indexed data zone row for this category's spent value."""
    return CAT_1 + CATS.index(cat)

def cat_spent_ref(cat, V):
    return f"{V}{cat_dz_row(cat)}"

def get_ceiling(cat):
    """Budget (spending) or savings balance (savings) as sparkline denominator."""
    if cat in SAVINGS_CATS:
        b = SAVINGS_BALANCES.get(cat, 0)
        return b if b > 0 else BDGETS.get(cat, 1)
    return BDGETS[cat]

def get_label(cat):
    if cat in SAVINGS_CATS:
        b = SAVINGS_BALANCES.get(cat, 0)
        return f"Balance: NOK {b:,}" if b > 0 else "Balance: (set in private.py)"
    return f"Budget: NOK {BDGETS[cat]:,}"

def spent_bare(cat):
    return (f"SUMPRODUCT(('{TRANS}'!E{T1}:E{T2}=\"{cat}\")"
            f"*'{TRANS}'!D{T1}:D{T2})")

def sparkline_fx(cat, V):
    sfx    = f"SUMPRODUCT(('{TRANS}'!E{T1}:E{T2}=\"{cat}\")*'{TRANS}'!D{T1}:D{T2})"
    ceil   = get_ceiling(cat)
    return (
        f'=SPARKLINE(MIN({sfx}/{ceil},1),'
        f'{{"charttype","bar";"max",1;'
        f'"color1",IF({sfx}>{ceil},"#E53935","#4472C4");'
        f'"color2","#EBF2FC"}})'
    )

def over_fx(cat, spent_cell):
    ceil = get_ceiling(cat)
    return (
        f'=IF({spent_cell}>{ceil},'
        f'"⚠ Over: +NOK "&TEXT({spent_cell}-{ceil},"#,##0"),"")'
    )

def filter_fx(cat):
    return (
        f'=IFERROR(FILTER(\'{TRANS}\'!B{T1}:D{T2},'
        f'\'{TRANS}\'!E{T1}:E{T2}="{cat}"),'
        f'{{"—","No transactions",0}})'
    )


def _restore_theme(svc):
    """Restore ACCENT2 to default orange — reverts the earlier white-arc attempt."""
    def c(r, g, b): return {"rgbColor": {"red": r/255, "green": g/255, "blue": b/255}}
    svc.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={"requests": [{"updateSpreadsheetProperties": {
            "properties": {"spreadsheetTheme": {
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
                ],
            }},
            "fields": "spreadsheetTheme",
        }}]}
    ).execute()


def _set_column_widths(svc, sid, N_active):
    svc.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={"requests": [
            {"updateDimensionProperties": {
                "range": {"sheetId": sid, "dimension": "COLUMNS",
                          "startIndex": 0, "endIndex": N_active * COLS_PER},
                "properties": {"pixelSize": COL_W_PX}, "fields": "pixelSize",
            }},
            {"updateDimensionProperties": {
                "range": {"sheetId": sid, "dimension": "COLUMNS",
                          "startIndex": DZ_L, "endIndex": DZ_V + 1},
                "properties": {"pixelSize": 1}, "fields": "pixelSize",
            }},
            {"updateDimensionProperties": {
                "range": {"sheetId": sid, "dimension": "ROWS",
                          "startIndex": CAT_HDR, "endIndex": CAT_HDR + 1},
                "properties": {"pixelSize": 30}, "fields": "pixelSize",
            }},
            {"updateDimensionProperties": {
                "range": {"sheetId": sid, "dimension": "ROWS",
                          "startIndex": OV_TITLE_R, "endIndex": OV_TITLE_R + 2},
                "properties": {"pixelSize": 28}, "fields": "pixelSize",
            }},
        ]}
    ).execute()


def _apply_styling(svc, sid, active_cats):
    N_active = len(active_cats)
    reqs = []

    # Overview panel (left of pie): title + health info
    reqs.append({"repeatCell": {
        "range": {"sheetId": sid,
                  "startRowIndex": OV_TITLE_R, "endRowIndex": OV_TITLE_R + 2,
                  "startColumnIndex": 0, "endColumnIndex": 7},
        "cell": {"userEnteredFormat": {
            "backgroundColor": BLUE_HDR,
            "textFormat": {"bold": True, "foregroundColor": WHITE, "fontSize": 13},
            "verticalAlignment": "MIDDLE", "padding": {"left": 10},
        }},
        "fields": "userEnteredFormat(backgroundColor,textFormat,verticalAlignment,padding)",
    }})
    reqs.append({"repeatCell": {
        "range": {"sheetId": sid,
                  "startRowIndex": OV_TITLE_R + 2, "endRowIndex": 15,
                  "startColumnIndex": 0, "endColumnIndex": 7},
        "cell": {"userEnteredFormat": {"backgroundColor": LIGHT_BLUE}},
        "fields": "userEnteredFormat.backgroundColor",
    }})
    for label_row in (2, 4, 7, 9):
        reqs.append({"repeatCell": {
            "range": {"sheetId": sid,
                      "startRowIndex": label_row, "endRowIndex": label_row + 1,
                      "startColumnIndex": 0, "endColumnIndex": 7},
            "cell": {"userEnteredFormat": {
                "textFormat": {"italic": True, "foregroundColor": rgb(100, 120, 160)},
                "padding": {"left": 10},
            }},
            "fields": "userEnteredFormat(textFormat,padding)",
        }})
    for val_row in (3, 5, 8, 10):
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

    # Per-category cards
    for idx, cat in enumerate(active_cats):
        cs, ce = idx * COLS_PER, idx * COLS_PER + COLS_PER
        tx_bg = LIGHT_BLUE if idx % 2 == 0 else rgb(248, 250, 255)

        reqs.append({"repeatCell": {
            "range": {"sheetId": sid,
                      "startRowIndex": CAT_HDR, "endRowIndex": CAT_HDR + 1,
                      "startColumnIndex": cs, "endColumnIndex": ce},
            "cell": {"userEnteredFormat": {
                "backgroundColor": BLUE_HDR,
                "textFormat": {"bold": True, "foregroundColor": WHITE, "fontSize": 11},
                "verticalAlignment": "MIDDLE", "padding": {"left": 8},
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat,verticalAlignment,padding)",
        }})
        reqs.append({"repeatCell": {
            "range": {"sheetId": sid,
                      "startRowIndex": CAT_BUD, "endRowIndex": OVER_ROW + 1,
                      "startColumnIndex": cs, "endColumnIndex": ce},
            "cell": {"userEnteredFormat": {"backgroundColor": LIGHT_BLUE}},
            "fields": "userEnteredFormat.backgroundColor",
        }})
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
        b = solid_border()
        reqs.append({"updateBorders": {
            "range": {"sheetId": sid,
                      "startRowIndex": CAT_HDR, "endRowIndex": TRANS_HDR + 1,
                      "startColumnIndex": cs, "endColumnIndex": ce},
            "top": b, "bottom": b, "left": b, "right": b,
        }})
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
        b_med = {"style": "SOLID_MEDIUM", "colorStyle": {"rgbColor": BORDER_CLR}}
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

    # ── Determine active categories ────────────────────────────────────────
    # Savings cats only appear if they have transactions this month.
    tx_col = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range=f"'{TRANS}'!E{T1}:E{T2}",
    ).execute().get("values", [])
    savings_with_tx = {r[0] for r in tx_col if r and r[0] in SAVINGS_CATS}

    active_savings = [c for c in CATS if c in SAVINGS_CATS and c in savings_with_tx]
    active_cats    = SPENDING_CATS + active_savings
    N_active       = len(active_cats)
    cat_cols_map   = {cat: i * COLS_PER for i, cat in enumerate(active_cats)}

    print(f"  Active categories ({N_active}): {', '.join(active_cats)}")
    if SAVINGS_CATS - savings_with_tx:
        hidden = SAVINGS_CATS - savings_with_tx
        print(f"  Hidden (no transactions): {', '.join(hidden)}")

    # ── Create / recreate the tab ──────────────────────────────────────────
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

    _set_column_widths(svc, sid, N_active)

    L = col_letter(DZ_L)   # "W"
    V = col_letter(DZ_V)   # "X"
    dt = []

    # ── Overview panel cells (left of pie, cols A-G) ───────────────────────
    dt += [
        {"range": f"'{TAB}'!A1",  "values": [["Monthly Overview"]]},
        {"range": f"'{TAB}'!A3",  "values": [["Financial Health"]]},
        {"range": f"'{TAB}'!A4",  "values": [[f"={V}{VERD_1}"]]},
        {"range": f"'{TAB}'!A8",  "values": [["Total Spent"]]},
        {"range": f"'{TAB}'!A9",  "values": [[f"=SUM({V}{CAT_1}:{V}{CAT_1+len(CATS)-1})"]]},
        {"range": f"'{TAB}'!A10", "values": [["Budget"]]},
        {"range": f"'{TAB}'!A11", "values": [[f"=TEXT({TOTAL_BUDGET},\"#,##0\")&\" NOK\""]]},
        {"range": f"'{TAB}'!A12", "values": [["Health Score"]]},
        {"range": f"'{TAB}'!A13", "values": [[f"={V}{HLTH_1}&\"%\""]]},
    ]

    # ── Data zone: all 7 category spents (W1:X7, fixed regardless of active) ──
    dt.append({"range": f"'{TAB}'!{L}{CAT_1}:{V}{CAT_1+len(CATS)-1}",
               "values": [[cat, f"={spent_bare(cat)}"] for cat in CATS]})

    # ── Data zone: surplus (W8:X8) ────────────────────────────────────────
    surplus_fx = (f"=MAX(0,{TOTAL_BUDGET}"
                  f"-SUM({V}{CAT_1}:{V}{CAT_1+len(CATS)-1}))")
    dt.append({"range": f"'{TAB}'!{L}{SURP_1}:{V}{SURP_1}",
               "values": [["Surplus", surplus_fx]]})

    # ── Data zone: health indicator (W10:X11) — spending cats only ────────
    adherence = "+".join(
        f"IF({V}{cat_dz_row(c)}<={BDGETS[c]},1,0)" for c in SPENDING_CATS
    )
    dt.append({"range": f"'{TAB}'!{L}{HLTH_1}:{V}{VERD_1}",
               "values": [
                   ["Health Score", f"=({adherence})/{N_SPENDING}*100"],
                   ["Verdict",
                    f'=IF({V}{HLTH_1}>=86,"On Track",'
                    f'IF({V}{HLTH_1}>=57,"Caution","Over Budget"))'],
               ]})

    # ── Data zone: per-cat donut data (W13:X26) — all 7 cats ─────────────
    donut_rows = []
    for cat in CATS:
        ref   = f"{V}{cat_dz_row(cat)}"
        ceil  = get_ceiling(cat)
        donut_rows.append(["Spent",     f"={ref}"])
        donut_rows.append(["Remaining", f"=MAX(0,{ceil}-{ref})"])
    dt.append({"range": f"'{TAB}'!{L}{DONUT_1}:{V}{DONUT_1+len(CATS)*2-1}",
               "values": donut_rows})

    # ── Visible category section rows (only active_cats) ──────────────────
    hdr_row, bud_row, spt_row, bar_row, ovr_row, thr_row = [], [], [], [], [], []
    for cat in active_cats:
        sc = cat_spent_ref(cat, V)
        hdr_row.extend([cat,              "", ""])
        bud_row.extend([get_label(cat),   "", ""])
        spt_row.extend([f"={sc}",         "", ""])
        bar_row.extend([sparkline_fx(cat, V), "", ""])
        ovr_row.extend([over_fx(cat, sc), "", ""])
        thr_row.extend(["Date", "Merchant", "Amount"])

    last_col_vis = col_letter(N_active * COLS_PER - 1)
    for vals, row_1idx in [
        (hdr_row, CAT_HDR   + 1),
        (bud_row, CAT_BUD   + 1),
        (spt_row, CAT_SPT   + 1),
        (bar_row, CAT_BAR   + 1),
        (ovr_row, OVER_ROW  + 1),
        (thr_row, TRANS_HDR + 1),
    ]:
        dt.append({"range": f"'{TAB}'!A{row_1idx}:{last_col_vis}{row_1idx}",
                   "values": [vals]})

    for cat in active_cats:
        cs = cat_cols_map[cat]
        dt.append({"range": f"'{TAB}'!{col_letter(cs)}{TRANS_DAT+1}",
                   "values": [[filter_fx(cat)]]})

    svc.spreadsheets().values().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={"valueInputOption": "USER_ENTERED", "data": dt},
    ).execute()
    print("  Data written")

    _apply_styling(svc, sid, active_cats)
    print("  Styling applied")

    # ── Charts ────────────────────────────────────────────────────────────
    def src_range(r0, r1_excl, col0):
        return {"sheetId": sid,
                "startRowIndex": r0, "endRowIndex": r1_excl,
                "startColumnIndex": col0, "endColumnIndex": col0 + 1}

    def pie_src_multi(row_ranges, col0):
        """Non-contiguous sourceRange for pie charts."""
        return {"sourceRange": {"sources": [
            src_range(r0, r1, col0) for r0, r1 in row_ranges
        ]}}

    # Overall pie: spending cats (rows 1–N_SPENDING) + surplus (row SURP_1)
    # Saved cats are intentionally excluded — they use a different ceiling and
    # should not dilute the monthly spending picture.
    spending_rows = (CAT_1 - 1, CAT_1 - 1 + N_SPENDING)   # 0-indexed: 0..4
    surplus_rows  = (SURP_1 - 1, SURP_1)                   # 0-indexed: 7..7

    total_px = N_active * COLS_PER * COL_W_PX
    pie_left = total_px // 2 - PIE_W_PX // 2
    pie_col  = pie_left // COL_W_PX
    pie_offx = pie_left - pie_col * COL_W_PX

    chart_reqs = [{"addChart": {"chart": {
        "spec": {
            "title": "",
            "backgroundColorStyle": {"rgbColor": WHITE},
            "pieChart": {
                "legendPosition": "LABELED_LEGEND",
                "threeDimensional": False,
                "domain": pie_src_multi([spending_rows, surplus_rows], DZ_L),
                "series": pie_src_multi([spending_rows, surplus_rows], DZ_V),
            },
        },
        "position": {"overlayPosition": {
            "anchorCell": {"sheetId": sid, "rowIndex": 0, "columnIndex": pie_col},
            "offsetXPixels": pie_offx,
            "widthPixels": PIE_W_PX, "heightPixels": 300,
        }},
    }}}]

    # Per-category donuts — only for active categories
    donut_offset_x = (COL_W_PX * COLS_PER - DNT_W_PX) // 2   # centre in 3-col block
    for cat in active_cats:
        cat_idx = CATS.index(cat)
        r0 = DONUT_1 - 1 + cat_idx * 2   # 0-indexed data zone row for this cat
        cs = cat_cols_map[cat]
        chart_reqs.append({"addChart": {"chart": {
            "spec": {
                "title": "",
                "backgroundColorStyle": {"rgbColor": LIGHT_BLUE},
                "pieChart": {
                    "legendPosition": "NO_LEGEND",
                    "pieHole": 0.5,
                    "threeDimensional": False,
                    "domain": {"sourceRange": {"sources": [src_range(r0, r0+2, DZ_L)]}},
                    "series": {"sourceRange": {"sources": [src_range(r0, r0+2, DZ_V)]}},
                },
            },
            "position": {"overlayPosition": {
                "anchorCell": {"sheetId": sid, "rowIndex": DNT_ANC, "columnIndex": cs},
                "offsetXPixels": donut_offset_x,
                "widthPixels":   DNT_W_PX,
                "heightPixels":  DNT_H_PX,
            }},
        }}})

    svc.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID, body={"requests": chart_reqs}).execute()
    print(f"  Added overview pie + {N_active} donut charts")

    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit"
    print(f"\nDone. Open: {url}")
    print()
    print("Savings categories: update SAVINGS_BALANCES in private.py with current")
    print("account balances, then re-run this script.")


if __name__ == "__main__":
    main()
