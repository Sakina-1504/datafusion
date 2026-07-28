"""
test_exporter_fix.py
=====================

Demo/test script for the bug:

    TypeError: cannot unpack non-iterable bool object

which happened in ui/dashboard.py at:

    success, reason = exporter.open_file_with_highlighted_columns(...)

ROOT CAUSE
----------
backend/exporter.py's open_file_with_highlighted_columns() only ever
returned a plain True/False. dashboard.py, however, expects a 2-item
tuple (success, reason) so it can show *why* something failed in the
"Couldn't Open File" popup. Unpacking a single bool into two variables
raises the TypeError you see in the log file.

THE FIX
-------
open_file_with_highlighted_columns() now returns (success, reason) on
every code path:
    - (True, None)                          -> everything worked
    - (False, "<human readable message>")   -> something failed, and why

HOW TO USE THIS FILE DURING YOUR PRESENTATION
----------------------------------------------
1. Drop this file into the root of DataFusionPlatform_v6 (next to app.py),
   so it can be run as: `python test_exporter_fix.py`
2. Copy the fixed exporter.py over backend/exporter.py (back up the old
   one first if you want to show the "before" crash).
3. Run the script. It exercises 4 scenarios end-to-end using a real
   temporary .xlsx file, and prints PASS/FAIL for each — no GUI needed.

To show the ORIGINAL bug live: temporarily rename backend/exporter.py's
fixed function to return only `success` instead of `(success, reason)`
(or restore your old backup exporter file) and re-run this script —
scenario 1 will crash with the exact TypeError from the log, proving
the bug is reproducible outside the full GUI. Then swap the fixed
version back in and re-run to show it passing.
"""

import os
import sys
import shutil
import tempfile

# Make sure "backend" package is importable when this file sits at the
# project root (same folder as app.py).
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from openpyxl import Workbook
from backend import exporter


def make_test_workbook(path, sheet_name="Sheet1", columns=None, missing_columns=False):
    """Creates a tiny .xlsx with a header row + a couple of data rows,
    so we have a realistic file for open_file_with_highlighted_columns
    to highlight and 'open'."""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    if missing_columns:
        headers = ["Date", "Narration", "Amount"]  # no Debit/Credit cols
    else:
        headers = columns or ["Date", "Narration", "Debit", "Credit"]

    ws.append(headers)
    ws.append(["2026-07-01", "Opening balance", 1000, 0])
    ws.append(["2026-07-02", "Purchase", 0, 500])
    wb.save(path)


def run_scenario(title, func):
    print(f"\n--- {title} ---")
    try:
        result = func()
        print(f"RESULT: {result}")
        print("STATUS: PASS (no crash, tuple returned as expected)")
        return True
    except TypeError as e:
        print(f"STATUS: FAIL -- TypeError: {e}")
        print("(This is the exact bug from the log file.)")
        return False
    except Exception as e:
        print(f"STATUS: FAIL -- Unexpected error: {e}")
        return False


def main():
    tmp_dir = tempfile.mkdtemp(prefix="datafusion_test_")
    print(f"Working in temp folder: {tmp_dir}")

    # Avoid actually launching Excel/OS file-open during the automated
    # test -- we only care that the function's return value can be
    # safely unpacked as (success, reason), which is what dashboard.py
    # does at line 1368.
    exporter.open_file = lambda path: True

    results = []

    # Scenario 1: happy path -- file exists, sheet exists, columns exist
    good_file = os.path.join(tmp_dir, "good_sample.xlsx")
    make_test_workbook(good_file)

    def scenario_1():
        success, reason = exporter.open_file_with_highlighted_columns(
            good_file, "Sheet1", ["Debit", "Credit"]
        )
        assert success is True and reason is None
        return (success, reason)

    results.append(run_scenario(
        "Scenario 1: valid file + valid columns (happy path)", scenario_1))

    # Scenario 2: file does not exist
    def scenario_2():
        success, reason = exporter.open_file_with_highlighted_columns(
            os.path.join(tmp_dir, "does_not_exist.xlsx"), "Sheet1", ["Debit", "Credit"]
        )
        assert success is False and isinstance(reason, str)
        return (success, reason)

    results.append(run_scenario(
        "Scenario 2: missing file (this is the exact case from the log)", scenario_2))

    # Scenario 3: sheet name does not exist in the workbook
    def scenario_3():
        success, reason = exporter.open_file_with_highlighted_columns(
            good_file, "NoSuchSheet", ["Debit", "Credit"]
        )
        assert success is False and isinstance(reason, str)
        return (success, reason)

    results.append(run_scenario(
        "Scenario 3: sheet name not found", scenario_3))

    # Scenario 4: columns to highlight aren't present in the header row
    missing_cols_file = os.path.join(tmp_dir, "missing_cols.xlsx")
    make_test_workbook(missing_cols_file, missing_columns=True)

    def scenario_4():
        success, reason = exporter.open_file_with_highlighted_columns(
            missing_cols_file, "Sheet1", ["Debit", "Credit"]
        )
        assert success is False and isinstance(reason, str)
        return (success, reason)

    results.append(run_scenario(
        "Scenario 4: Debit/Credit columns missing from header", scenario_4))

    # Scenario 5: output copy is locked (simulates the file being open in
    # Excel already) -- should fall back to a fresh filename instead of
    # failing outright.
    def scenario_5():
        from backend.settings import _user_data_dir
        import time

        out_dir = os.path.join(_user_data_dir(), "highlighted_references")
        os.makedirs(out_dir, exist_ok=True)
        base_name = os.path.splitext(os.path.basename(good_file))[0]
        locked_path = os.path.join(out_dir, f"{base_name}__Sheet1__highlighted.xlsx")

        # Pre-create the target file, then monkeypatch Workbook.save so the
        # *first* save attempt raises PermissionError (as Windows would if
        # the file were open in Excel), forcing the fallback path.
        from openpyxl import Workbook as WB
        WB().save(locked_path)

        import backend.exporter as exp
        original_save = exp.load_workbook  # placeholder, not used
        from openpyxl.workbook.workbook import Workbook as OpenpyxlWorkbook
        real_save = OpenpyxlWorkbook.save
        call_count = {"n": 0}

        def flaky_save(self, path):
            call_count["n"] += 1
            if call_count["n"] == 1:
                raise PermissionError(13, "Permission denied", path)
            return real_save(self, path)

        OpenpyxlWorkbook.save = flaky_save
        try:
            success, reason = exporter.open_file_with_highlighted_columns(
                good_file, "Sheet1", ["Debit", "Credit"]
            )
        finally:
            OpenpyxlWorkbook.save = real_save

        assert success is True and reason is None
        return (success, reason)

    results.append(run_scenario(
        "Scenario 5: output file locked (already open in Excel) -> fallback filename", scenario_5))

    shutil.rmtree(tmp_dir, ignore_errors=True)

    print("\n================ SUMMARY ================")
    passed = sum(results)
    print(f"{passed}/{len(results)} scenarios passed")
    if all(results):
        print("All good -- the (success, reason) tuple fix works end-to-end.")
    else:
        print("Some scenarios failed -- see FAIL lines above for details.")


if __name__ == "__main__":
    main()