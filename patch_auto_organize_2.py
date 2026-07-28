"""
patch_auto_organize_2.py

Run this once from inside your DataFusionPlatform_v6 folder:

    python patch_auto_organize_2.py

FIXES THE REAL BUG behind the "Account" / "Account (2)" mess (verified
against your actual GSR Foods II LLC.xlsx file): when a source file's
per-company grand-total row has the word "TOTAL" sitting alone in its
own mostly-empty column (very common in QuickBooks Trial Balance
exports), Auto-Organize was mistaking that column for a real text
column and letting it steal the "Account" name -- which caused the
REAL account-label column right next to it to get silently dropped as
a "duplicate". That's why account labels were showing up blank for
some companies.

This patch relocates that lone "TOTAL" marker into the real
neighboring column instead, so:
  - The real account-label column keeps the "Account" name (no more
    stolen name, no more dropped data).
  - The TOTAL row still shows "TOTAL" in the right place, same as
    before.
  - Files that only have one genuine Account column, or genuinely have
    both "Account" and "Account (2)" as real headers, are unaffected.

Tested directly against your uploaded GSR_Foods_II_LLC.xlsx plus three
other scenarios (real Account+Account(2) headers, a plain clean file,
and a fully-empty unnamed column) to confirm nothing else changes.

Makes a backup first as backend/auto_organize_backup2.py.
"""

import shutil
import sys

PATH = "backend/auto_organize.py"
BACKUP = "backend/auto_organize_backup2.py"

try:
    with open(PATH, "r", encoding="utf-8") as f:
        content = f.read()
except FileNotFoundError:
    print(f"ERROR: could not find {PATH}")
    print("Make sure you're running this from inside the DataFusionPlatform_v6 folder.")
    sys.exit(1)

old_block = '''    if not rows:
        return None
    header_idx = detect_header_row(rows)
    if header_idx is None:
        header_idx = 0  # nothing messy found -- treat row 0 as header

    num_cols = max(len(r) for r in rows)
    header_row = rows[header_idx] + [""] * (num_cols - len(rows[header_idx]))
    metadata_cols = _metadata_columns(rows, header_idx, num_cols)
    title = _build_title(rows, header_idx, metadata_cols)'''

new_block = '''    if not rows:
        return None
    rows = [list(r) for r in rows]  # ensure every row is mutable
    header_idx = detect_header_row(rows)
    if header_idx is None:
        header_idx = 0  # nothing messy found -- treat row 0 as header

    num_cols = max(len(r) for r in rows)
    header_row = rows[header_idx] + [""] * (num_cols - len(rows[header_idx]))

    # A column whose ONLY value in the whole sheet is the literal word
    # "TOTAL" is a grand-total marker carried over from the source
    # file (usually meant to line up with the real label column next
    # to it). Relocate it into that neighboring column so it doesn't
    # steal the "Account" name away from the real label column, and
    # doesn't turn into its own near-empty flagged column -- it just
    # becomes part of the real column's data, exactly where a person
    # reading the original file would expect the TOTAL label to sit.
    data_rows = rows[header_idx + 1:]
    for c in range(num_cols):
        vals = [(ri, str(r[c]).strip()) for ri, r in enumerate(data_rows) if c < len(r) and str(r[c]).strip() != ""]
        if len(vals) == 1 and vals[0][1].upper() == "TOTAL":
            ri, _ = vals[0]
            target_row = data_rows[ri]
            for nc in (c + 1, c - 1):
                if 0 <= nc < num_cols:
                    while len(target_row) <= max(nc, c):
                        target_row.append("")
                    if str(target_row[nc]).strip() == "":
                        target_row[nc] = "TOTAL"
                        target_row[c] = ""
                        break

    metadata_cols = _metadata_columns(rows, header_idx, num_cols)
    title = _build_title(rows, header_idx, metadata_cols)'''

if old_block in content:
    content = content.replace(old_block, new_block)
elif "grand-total marker carried over" in content:
    print("Nothing to change -- this fix is already applied.")
    sys.exit(0)
else:
    print("ERROR: Could not find the expected code block to change.")
    print("No changes were made. Please share this message with support.")
    sys.exit(1)

shutil.copy(PATH, BACKUP)
with open(PATH, "w", encoding="utf-8") as f:
    f.write(content)

print("SUCCESS. Real account labels will no longer be dropped because of a lone TOTAL marker.")
print(f"Backup of the previous version saved as {BACKUP}")
print("Now run: python app.py  and generate a fresh report.")
