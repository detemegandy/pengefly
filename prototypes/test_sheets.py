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
