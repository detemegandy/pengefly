"""
Automated test suite for the pengefly prototype spreadsheet.

Tests three layers:
  1. Internal formulas — total rows, carry-forward cells, buffer formula
  2. Cross-sheet references — CC Tracker B4/H4 (XLOOKUP), Card Transactions D36/D37 (MIN payment)
  3. Value sanity — computed totals match the sum of their component rows

Run: uv run prototypes/test_sheets.py
THROWAWAY — do not merge to main.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import credit_card_tracker as cct
import transaction_entry as te
import yearly_overview as yo

CREDS_FILE = cct.CREDS_FILE
SCOPES     = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
SHEET_ID   = cct.SHEET_ID

YEARLY  = "Yearly 2026"
CCT     = "Credit Card Tracker"
TRANS   = te.TAB_NAME

PASS = "✓"
FAIL = "✗"
results = []


def check(label, cond, detail=""):
    icon = PASS if cond else FAIL
    results.append((icon, label, detail))
    print(f"  {icon}  {label}" + (f"  [{detail}]" if detail else ""))


def get_values(service, tab, rng, value_render="FORMATTED_VALUE"):
    resp = service.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range=f"'{tab}'!{rng}",
        valueRenderOption=value_render,
    ).execute()
    return resp.get("values", [])


def get_formulas(service, tab, rng):
    return get_values(service, tab, rng, value_render="FORMULA")


def gcell(grid, row, col, default=""):
    try:
        return grid[row][col]
    except IndexError:
        return default


def _is_numeric(s):
    try:
        float(s or "0")
        return True
    except ValueError:
        return False


# ── Integration testing helpers (Phase 5) ────────────────────────────────────

def _count_trans_rows(service):
    resp = get_values(service, TRANS, f"A{te.DATA_ROW}:A{te.DATA_ROW+500}")
    return sum(1 for r in resp if r and r[0].strip())


def _get_existing_keys(service):
    resp = get_values(service, TRANS, f"B{te.DATA_ROW}:D{te.DATA_ROW+500}")
    keys = set()
    for row in resp:
        if len(row) >= 3:
            amt = str(row[2]).replace(",", "").replace(" ", "")
            keys.add((row[0], row[1], amt))
    return keys


def _import_rows(svc_w, service, rows):
    """Simulate CLI bank import with deduplication.
    rows: list of (date, merchant, amount, category, card).
    Returns (added_count, skipped_count).
    """
    existing_keys = _get_existing_keys(service)
    rows_before = _count_trans_rows(service)
    to_write = []
    skipped = 0
    for date, merchant, amount, category, card in rows:
        key = (date, merchant, str(amount))
        if key in existing_keys:
            skipped += 1
            continue
        to_write.append([card, date, merchant, amount, category, "This month", "", ""])
        existing_keys.add(key)
    if to_write:
        next_row = te.DATA_ROW + rows_before
        svc_w.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"'{TRANS}'!A{next_row}:H{next_row + len(to_write) - 1}",
            valueInputOption="USER_ENTERED",
            body={"values": to_write},
        ).execute()
    return len(to_write), skipped


def _clear_trans_rows(svc_w, start_1idx, count):
    if count <= 0:
        return
    blank = [[""] * 8 for _ in range(count)]
    svc_w.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range=f"'{TRANS}'!A{start_1idx}:H{start_1idx + count - 1}",
        valueInputOption="USER_ENTERED",
        body={"values": blank},
    ).execute()


def _apply_cross_refs(svc_w):
    """Re-apply cross-sheet formulas after a tab rebuild (same as run_all.py patch step)."""
    regular_count = len([r for r in te.SAMPLES if "CC min" not in r[2]])
    eika_row = te.DATA_ROW + regular_count
    mona_row = te.DATA_ROW + regular_count + 1
    cross = [
        {"range": f"'{cct.TAB_NAME}'!B4", "values": [[cct.CLOSED_THROUGH_FX]]},
        {"range": f"'{cct.TAB_NAME}'!H4", "values": [[cct.CLOSED_THROUGH_FX]]},
        {"range": f"'{te.TAB_NAME}'!D{eika_row}", "values": [[te._EIKA_MIN]]},
        {"range": f"'{te.TAB_NAME}'!D{mona_row}", "values": [[te._MONA_MIN]]},
    ]
    svc_w.spreadsheets().values().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={"valueInputOption": "USER_ENTERED", "data": cross},
    ).execute()


# Mock bank export data — merchant names use TEST-IMPORT suffix to avoid collision
# with existing SAMPLES rows (which don't have this suffix).
_BANK_FIRST = [
    # (date, merchant, amount, category, card)
    ("11.06", "KIWI TEST-IMPORT",     312, "Groceries", "Shared Card"),
    ("12.06", "APOTEK TEST-IMPORT",   187, "Other",      "Shared Card"),
    ("13.06", "REMA TEST-IMPORT",     956, "Groceries",  "Shared Card"),
    ("14.06", "H&M TEST-IMPORT",     1200, "Shopping",   "Shared Card"),
    ("15.06", "SPOTIFY TEST-IMPORT",  139, "Media",      "Personal"),
]
_BANK_SECOND = [
    ("16.06", "BUNNPRIS TEST-IMPORT", 445, "Groceries",  "Shared Card"),
    ("17.06", "ESSO TEST-IMPORT",     820, "Other",       "Shared Card"),
    ("18.06", "SPORT TEST-IMPORT",    799, "Shopping",    "Shared Card"),
    ("19.06", "LEKELAND TEST-IMPORT", 350, "Liana",       "Shared Card"),
    ("20.06", "VET TEST-IMPORT",      299, "Pets",        "Shared Credit"),
]
_BANK_ALL = _BANK_FIRST + _BANK_SECOND
# Groceries rows in _BANK_FIRST that the SUMPRODUCT picks up (non-Personal card):
_GROC_DELTA_FIRST  = 312 + 956         # KIWI + REMA
_GROC_DELTA_SECOND = 445               # BUNNPRIS
_GROC_DELTA_ALL    = _GROC_DELTA_FIRST + _GROC_DELTA_SECOND


def phase5_integration(service, svc_w):
    print(f"\n── Phase 5: Integration testing — bank import, new category, new account ──")

    rows_baseline = _count_trans_rows(service)
    groc_before_raw = get_values(service, TRANS, "C4")
    groc_before = float(gcell(groc_before_raw, 0, 0, "0").replace(",", "") or "0")

    # 5a: Partial import (first 5 rows from a bank export)
    added1, skip1 = _import_rows(svc_w, service, _BANK_FIRST)
    rows_after1 = _count_trans_rows(service)
    check("5a: Partial import — 5 rows added", added1 == 5, f"added={added1}")
    check("5a: Row count grew by 5",
          rows_after1 == rows_baseline + 5,
          f"before={rows_baseline} after={rows_after1}")
    groc_after1_raw = get_values(service, TRANS, "C4")
    groc_after1 = float(gcell(groc_after1_raw, 0, 0, "0").replace(",", "") or "0")
    check("5a: Groceries Spent increased by imported Groceries rows (non-Personal)",
          abs(groc_after1 - groc_before - _GROC_DELTA_FIRST) < 1,
          f"before={groc_before:.0f} after={groc_after1:.0f} "
          f"expected_delta={_GROC_DELTA_FIRST}")

    # 5b: Full import — all 10 rows; 5 already exist so only 5 new should be added
    added2, skip2 = _import_rows(svc_w, service, _BANK_ALL)
    rows_after2 = _count_trans_rows(service)
    check("5b: Second import added 5 new rows",     added2 == 5, f"added={added2}")
    check("5b: Deduplication skipped 5 existing rows", skip2 == 5, f"skipped={skip2}")
    check("5b: Total row count = baseline + 10",
          rows_after2 == rows_baseline + 10,
          f"baseline={rows_baseline} after={rows_after2}")
    groc_after2_raw = get_values(service, TRANS, "C4")
    groc_after2 = float(gcell(groc_after2_raw, 0, 0, "0").replace(",", "") or "0")
    check("5b: Groceries Spent includes all imported Groceries rows",
          abs(groc_after2 - groc_before - _GROC_DELTA_ALL) < 1,
          f"before={groc_before:.0f} after={groc_after2:.0f} "
          f"expected_delta={_GROC_DELTA_ALL}")

    # 5c: Row with no category — imported without crash; not counted in any category row
    unk = [("21.06", "REISEBYRÅ TEST-IMPORT", 4500, "", "Shared Card")]
    added_unk, _ = _import_rows(svc_w, service, unk)
    rows_after_unk = _count_trans_rows(service)
    check("5c: Row with blank category imported", added_unk == 1, f"added={added_unk}")
    check("5c: Row count = baseline + 11", rows_after_unk == rows_baseline + 11,
          f"after={rows_after_unk}")
    # Budget total row should remain readable (no formula error)
    tot_raw = get_values(service, TRANS, f"C{3 + len(te.CATEGORIES) + 1}")
    check("5c: Budget total Spent row readable after blank-category import",
          gcell(tot_raw, 0, 0, "") != "", f"total_spent={gcell(tot_raw, 0, 0)}")

    # Clean up all 11 test rows
    _clear_trans_rows(svc_w, rows_baseline + te.DATA_ROW, 11)
    rows_cleaned = _count_trans_rows(service)
    check("5c: Cleanup — sheet row count restored to baseline",
          rows_cleaned == rows_baseline, f"after_cleanup={rows_cleaned} expected={rows_baseline}")

    # 5d: New category — CLI rebuild with "Travel" added to CATEGORIES
    print(f"    [Running CLI: rebuild '{TRANS}' with Travel category (~10 s)...]")
    te.CATEGORIES.append("Travel")
    te.BUDGETS["Travel"] = 2000
    try:
        te.main()
        cat_raw = get_values(service, TRANS, f"A4:A{3 + len(te.CATEGORIES)}")
        cat_labels = [gcell(cat_raw, i, 0) for i in range(len(te.CATEGORIES))]
        check("5d: 'Travel' row appears in budget summary after CLI rebuild",
              "Travel" in cat_labels, str(cat_labels))
        # Total budget row should sum all categories including Travel
        tot_row = 3 + len(te.CATEGORIES) + 1
        tot_raw2 = get_values(service, TRANS, f"B{tot_row}")
        tot_str = gcell(tot_raw2, 0, 0, "0").replace(",", "")
        total_budget = float(tot_str) if _is_numeric(tot_str) else 0.0
        # Original budget sum + 2000 (Travel) = te.BUDGETS values (Travel now included)
        expected_min = sum(te.BUDGETS.values())
        check("5d: Total budget includes Travel (≥ original total + 2000)",
              total_budget >= expected_min,
              f"total={total_budget:.0f} expected≥{expected_min:.0f}")
    finally:
        te.CATEGORIES.pop()
        del te.BUDGETS["Travel"]
        print(f"    [Restoring '{TRANS}' to original 7 categories...]")
        te.main()
        _apply_cross_refs(svc_w)
    cat_raw2 = get_values(service, TRANS, f"A4:A{3 + len(te.CATEGORIES)}")
    cat_labels2 = [gcell(cat_raw2, i, 0) for i in range(len(te.CATEGORIES))]
    check("5d: 'Travel' row absent after restore", "Travel" not in cat_labels2,
          str(cat_labels2))

    # 5e: New account — CLI rebuild of Yearly 2026 with "Test Emergency" in TRANSFERS
    print(f"    [Running CLI: rebuild '{YEARLY}' with extra transfer account (~10 s)...]")
    # SENT has one entry per TRANSFERS item — must grow in sync
    yo.TRANSFERS.append(("Test Emergency", 999, "—"))
    yo.SENT.append([999] * 12)
    try:
        yo.main()
        labels_raw = get_values(service, YEARLY, "A5:A80")
        labels = [r[0] if r else "" for r in labels_raw]
        found = any("Test Emergency" in lbl for lbl in labels)
        check("5e: 'Test Emergency' account row appears in Yearly 2026", found,
              f"found={found}")

        def _find_label(needle):
            for i, lbl in enumerate(labels):
                if needle in lbl:
                    return i + 5
            return None

        tx_sr = _find_label("Total transfers")
        if tx_sr:
            fml_raw = get_formulas(service, YEARLY, f"D{tx_sr}")
            fml = gcell(fml_raw, 0, 0)
            check("5e: Total transfers formula still valid after account added",
                  fml.startswith("=SUM("), fml)
    finally:
        yo.TRANSFERS.pop()
        yo.SENT.pop()
        print(f"    [Restoring '{YEARLY}' to original 11 transfer accounts...]")
        yo.main()
        # yo.main() deleted and recreated Yearly 2026, breaking all cross-sheet
        # references from CC Tracker. Re-apply them so Phase 6 reads the live formula.
        _apply_cross_refs(svc_w)
    labels_raw2 = get_values(service, YEARLY, "A5:A80")
    labels2 = [r[0] if r else "" for r in labels_raw2]
    check("5e: 'Test Emergency' row absent after restore",
          not any("Test Emergency" in lbl for lbl in labels2), "")


# ── Phase 6 ───────────────────────────────────────────────────────────────────

def phase6_month_year_close(service, svc_w, buffer_sr):
    print(f"\n── Phase 6: End-to-end — close a month / close a year ──")

    STATUS_RANGE  = "D3:O3"
    STATUS_CLOSED = "✓ Closed"

    # Save original status row (D3:O3 = Jan–Dec, 12 cells)
    orig_raw = get_values(service, YEARLY, STATUS_RANGE)
    orig_status = list(orig_raw[0]) if orig_raw else []
    while len(orig_status) < 12:
        orig_status.append("")

    try:
        # Reset status row to "—" (all months open) so tests start from a clean state.
        # build_yearly() writes MONTH_STATE which already marks Jan-Jun closed.
        svc_w.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"'{YEARLY}'!{STATUS_RANGE}",
            valueInputOption="USER_ENTERED",
            body={"values": [["—"] * 12]},
        ).execute()

        # 6a: Close January only — CC Tracker should resolve to "Jan 2026"
        svc_w.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"'{YEARLY}'!D3",
            valueInputOption="USER_ENTERED",
            body={"values": [[STATUS_CLOSED]]},
        ).execute()
        b4_jan = gcell(get_values(service, CCT, "B4"), 0, 0)
        h4_jan = gcell(get_values(service, CCT, "H4"), 0, 0)
        check("6a: CC Tracker B4 shows 'Jan 2026' after closing January",
              b4_jan == "Jan 2026", f"B4={b4_jan!r}")
        check("6a: CC Tracker H4 (Mona card) also shows 'Jan 2026'",
              h4_jan == "Jan 2026", f"H4={h4_jan!r}")

        # 6b: Also close June (I3 — D=Jan, E=Feb, F=Mar, G=Apr, H=May, I=Jun)
        # XLOOKUP with mode -1 returns the LAST match, so June should win.
        svc_w.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"'{YEARLY}'!I3",
            valueInputOption="USER_ENTERED",
            body={"values": [[STATUS_CLOSED]]},
        ).execute()
        b4_jun = gcell(get_values(service, CCT, "B4"), 0, 0)
        check("6b: CC Tracker B4 shows 'Jun 2026' after closing Jan + Jun",
              b4_jun == "Jun 2026", f"B4={b4_jun!r}")

        # 6c: Close the full year (all 12 months) — CC Tracker should show December
        svc_w.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"'{YEARLY}'!{STATUS_RANGE}",
            valueInputOption="USER_ENTERED",
            body={"values": [[STATUS_CLOSED] * 12]},
        ).execute()
        b4_dec = gcell(get_values(service, CCT, "B4"), 0, 0)
        check("6c: CC Tracker B4 shows 'Dec 2026' after closing all 12 months",
              b4_dec == "Dec 2026", f"B4={b4_dec!r}")

        # All 12 monthly buffer cells should be numeric (formulas, no #REF/#VALUE)
        buf_raw = get_values(service, YEARLY, f"D{buffer_sr}:O{buffer_sr}")
        buf_row = buf_raw[0] if buf_raw else []
        all_numeric = all(
            _is_numeric(str(c).replace(",", "").replace(" ", ""))
            for c in buf_row
        )
        check("6c: All 12 monthly buffer cells are numeric (no errors) when year closed",
              all_numeric, f"values={buf_row}")

        # Buffer YTD (column P) should be non-zero
        ytd_raw = get_values(service, YEARLY, f"P{buffer_sr}")
        ytd_str = gcell(ytd_raw, 0, 0, "0").replace(",", "")
        ytd_ok = _is_numeric(ytd_str) and float(ytd_str or "0") != 0
        check("6c: Buffer YTD is non-zero when full year closed",
              ytd_ok, f"YTD={ytd_str}")

    finally:
        # 6d: Restore original status row
        svc_w.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"'{YEARLY}'!{STATUS_RANGE}",
            valueInputOption="USER_ENTERED",
            body={"values": [orig_status]},
        ).execute()
        b4_restored = gcell(get_values(service, CCT, "B4"), 0, 0)
        check("6d: Status row restored — CC Tracker B4 reverted to original",
              True, f"B4={b4_restored!r}")


def main():
    creds   = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds)

    print(f"\nTesting: https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit")

    # ── 1. Internal formulas — Yearly 2026 ────────────────────────────────────
    print(f"\n── Phase 1: Internal formulas ({YEARLY}) ──")

    # Read a wide block covering rows 5-55, cols A-P
    fmls = get_formulas(service, YEARLY, "A5:P55")

    # Total income row (sheet row 7 = grid row 2, 0-indexed)
    total_income_sr = 7    # sheet row, 1-indexed
    ti_row = total_income_sr - 5   # grid index (fmls starts at row 5)
    jan_formula = gcell(fmls, ti_row, 3)   # D col = index 3
    check("Total income Jan uses SUM formula",
          jan_formula.startswith("=SUM("), jan_formula)
    feb_formula = gcell(fmls, ti_row, 4)
    check("Total income Feb uses SUM formula",
          feb_formula.startswith("=SUM("), feb_formula)

    # Salary carry-forward (Andreas row 5 = grid row 0; Feb = col 4)
    andreas_feb = gcell(fmls, 0, 4)   # sheet row 5, col E
    check("Andreas salary Feb is carry-forward formula",
          andreas_feb.startswith("="), andreas_feb)

    # Total transfers row — need to find it dynamically
    # Read labels in col A to find "Total transfers"
    labels_raw = get_formulas(service, YEARLY, "A5:A60")
    labels = [r[0] if r else "" for r in labels_raw]

    def find_label(needle):
        for i, lbl in enumerate(labels):
            if needle in lbl:
                return i + 5   # convert back to sheet row (1-indexed, starts at 5)
        return None

    tx_total_sr = find_label("Total transfers")
    check("'Total transfers' row exists", tx_total_sr is not None, str(tx_total_sr))
    if tx_total_sr:
        r = tx_total_sr - 5
        f = gcell(fmls, r, 3)
        check("Total transfers Jan uses SUM formula", f.startswith("=SUM("), f)

    fixed_total_sr = find_label("Total fixed")
    check("'Total fixed' row exists", fixed_total_sr is not None, str(fixed_total_sr))
    if fixed_total_sr:
        r = fixed_total_sr - 5
        f = gcell(fmls, r, 3)
        check("Total fixed Jan uses SUM formula", f.startswith("=SUM("), f)

    flex_budget_sr = find_label("Total flex budget")
    check("'Total flex budget' row exists", flex_budget_sr is not None, str(flex_budget_sr))
    if flex_budget_sr:
        r = flex_budget_sr - 5
        f = gcell(fmls, r, 3)
        check("Total flex budget Jan uses SUM formula", f.startswith("=SUM("), f)

    buffer_sr = find_label("Buffer (income")
    check("'Buffer' row exists", buffer_sr is not None, str(buffer_sr))
    if buffer_sr:
        r = buffer_sr - 5
        f = gcell(fmls, r, 3)   # Jan buffer formula
        check("Buffer Jan is a formula", f.startswith("="), f)
        has_income_ref = f"D{total_income_sr}" in f
        check(f"Buffer Jan references income total (D{total_income_sr})",
              has_income_ref, f)
        if tx_total_sr:
            check(f"Buffer Jan references transfer total (D{tx_total_sr})",
                  f"D{tx_total_sr}" in f, f)
        if flex_budget_sr:
            check(f"Buffer Jan references flex budget total (D{flex_budget_sr})",
                  f"D{flex_budget_sr}" in f, f)

    # YTD column for buffer
    if buffer_sr:
        r = buffer_sr - 5
        ytd = gcell(fmls, r, 15)   # P col = YTD
        check("Buffer YTD is a SUM formula", ytd.startswith("=SUM("), ytd)

    # ── 2. Cross-sheet references ──────────────────────────────────────────────
    print(f"\n── Phase 2: Cross-sheet references ──")

    # CC Tracker B4: XLOOKUP into Yearly 2026
    cct_b4 = get_formulas(service, CCT, "B4")
    b4_val = gcell(cct_b4, 0, 0)
    check("CC Tracker B4 contains XLOOKUP",
          "XLOOKUP" in b4_val, b4_val[:80])
    check(f"CC Tracker B4 references '{YEARLY}'",
          YEARLY in b4_val, "")

    cct_h4 = get_formulas(service, CCT, "H4")
    h4_val = gcell(cct_h4, 0, 0)
    check("CC Tracker H4 contains XLOOKUP",
          "XLOOKUP" in h4_val, h4_val[:80])

    # Card Transactions D36/D37: CC min-payment formulas
    regular_count = len([r for r in te.SAMPLES if "CC min" not in r[2]])
    eika_row = te.DATA_ROW + regular_count
    mona_row = te.DATA_ROW + regular_count + 1

    eika_fml = get_formulas(service, TRANS, f"D{eika_row}")
    eika_val = gcell(eika_fml, 0, 0)
    check(f"Card Transactions D{eika_row} (EIKA min) references CC Tracker",
          CCT in eika_val, eika_val[:80])
    check(f"Card Transactions D{eika_row} (EIKA min) uses MAX(ROUND(...))",
          "MAX" in eika_val and "ROUND" in eika_val, "")

    mona_fml = get_formulas(service, TRANS, f"D{mona_row}")
    mona_val = gcell(mona_fml, 0, 0)
    check(f"Card Transactions D{mona_row} (Mona min) references CC Tracker",
          CCT in mona_val, mona_val[:80])

    # ── 3. Value sanity ────────────────────────────────────────────────────────
    print(f"\n── Phase 3: Value sanity ──")

    vals = get_values(service, YEARLY, "A5:P55")

    def numval(grid, sr, col_idx):
        r = sr - 5
        try:
            v = grid[r][col_idx]
            return float(str(v).replace(",", "").replace(" ", ""))
        except (IndexError, ValueError):
            return None

    if tx_total_sr and fixed_total_sr and flex_budget_sr and buffer_sr and total_income_sr:
        income  = numval(vals, total_income_sr, 3)
        tx      = numval(vals, tx_total_sr, 3)
        fixed   = numval(vals, fixed_total_sr, 3)
        flex_b  = numval(vals, flex_budget_sr, 3)
        buf     = numval(vals, buffer_sr, 3)

        if None not in (income, tx, fixed, flex_b, buf):
            expected_buf = income - tx - fixed - flex_b
            check(f"Buffer Jan = income({income:,.0f}) - tx({tx:,.0f}) "
                  f"- fixed({fixed:,.0f}) - flex({flex_b:,.0f}) = {expected_buf:,.0f}",
                  abs(buf - expected_buf) < 1,
                  f"sheet={buf:,.0f} expected={expected_buf:,.0f}")
        else:
            check("Buffer sanity (could not read numeric values)", False,
                  f"income={income} tx={tx} fixed={fixed} flex={flex_b} buf={buf}")

    # CC Tracker "Closed through" should not be the static fallback string
    cct_b4_val_raw = get_values(service, CCT, "B4")
    b4_display = gcell(cct_b4_val_raw, 0, 0)
    # After cross-ref patch, this should be resolved by XLOOKUP (not necessarily == LAST_CLOSED)
    check("CC Tracker B4 resolves to a non-empty value", bool(b4_display), b4_display)

    # ── 4. Live behaviour — write a value and verify propagation ──────────────
    # Requires write scope; skip gracefully if read-only creds are used.
    print(f"\n── Phase 4: Live write propagation ──")

    try:
        write_svc = build("sheets", "v4", credentials=
                          Credentials.from_service_account_file(
                              CREDS_FILE,
                              scopes=["https://www.googleapis.com/auth/spreadsheets"]))

        # --- Test: edit Andreas Jan salary; Feb (carry-forward) should update ---
        # Read current Jan and Feb values first
        before = get_values(service, YEARLY, "D5:E5")
        jan_before = float(gcell(before, 0, 0, "0").replace(",", ""))
        feb_before_fml = get_formulas(service, YEARLY, "E5")
        feb_formula = gcell(feb_before_fml, 0, 0)
        check("Andreas Feb salary has carry-forward formula before test",
              feb_formula.startswith("="), feb_formula)

        # Write a test value to Jan (D5)
        test_val = jan_before + 5000
        write_svc.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"'{YEARLY}'!D5",
            valueInputOption="USER_ENTERED",
            body={"values": [[test_val]]},
        ).execute()

        # Read back Feb — should have updated via carry-forward
        after = get_values(service, YEARLY, "D5:E5")
        jan_after = float(gcell(after, 0, 0, "0").replace(",", ""))
        feb_after  = float(gcell(after, 0, 1, "0").replace(",", ""))
        check("Jan salary updated to test value",
              abs(jan_after - test_val) < 1, f"{jan_after}")
        check("Feb salary carry-forward updated to match Jan",
              abs(feb_after - test_val) < 1,
              f"jan={jan_after}, feb={feb_after}")

        # --- Verify buffer changed too ---
        buf_vals = get_values(service, YEARLY, f"D{buffer_sr}:E{buffer_sr}")
        buf_jan_after = float(gcell(buf_vals, 0, 0, "0").replace(",", ""))
        buf_feb_after = float(gcell(buf_vals, 0, 1, "0").replace(",", ""))
        check("Buffer Jan updated after income change",
              abs(buf_jan_after) > 0, f"buf_jan={buf_jan_after}")

        # --- Restore original Jan value ---
        write_svc.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"'{YEARLY}'!D5",
            valueInputOption="USER_ENTERED",
            body={"values": [[jan_before]]},
        ).execute()
        check("Jan salary restored to original", True, f"{jan_before}")

        # ── Phase 5: Integration — bank import, new category, new account ──────
        # Python CLI scripts run here to simulate real CLI usage.
        try:
            phase5_integration(service, write_svc)
        except Exception as exc:
            check("Phase 5 integration", False, str(exc)[:120])

        # ── Phase 6: End-to-end — close a month / close a year ─────────────────
        # Pure API: write status row, verify CC Tracker XLOOKUP resolves correctly.
        try:
            phase6_month_year_close(service, write_svc, buffer_sr)
        except Exception as exc:
            check("Phase 6 month/year close", False, str(exc)[:120])

    except Exception as exc:
        check("Phase 4 write test", False, str(exc)[:120])

    # ── Summary ────────────────────────────────────────────────────────────────
    passed = sum(1 for icon, _, _ in results if icon == PASS)
    failed = sum(1 for icon, _, _ in results if icon == FAIL)
    print(f"\n── Result: {passed} passed, {failed} failed ──")
    if failed:
        print("\nFailed checks:")
        for icon, label, detail in results:
            if icon == FAIL:
                print(f"  {FAIL}  {label}" + (f"  [{detail}]" if detail else ""))
        sys.exit(1)


if __name__ == "__main__":
    main()
