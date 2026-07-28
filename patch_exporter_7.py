"""
patch_exporter_7.py

Run this once from inside your DataFusionPlatform_v6 folder:

    python patch_exporter_7.py

FIXES A BUG introduced by patch_exporter_6.py: that patch dropped the
"Account" column whenever "Account (2)" was also present anywhere in
the merged data -- but for some source files (e.g. GSR Foods III LLC)
"Account" is the ONLY column with real account labels, so dropping it
deleted real data for those rows.

This patch replaces that with a per-row MERGE instead: for each row,
it keeps whichever of "Account" / "Account (2)" actually has a value
(preferring "Account (2)" when both have one), combines them into a
single "Account" column, and drops the now-redundant "Account (2)".
No data is lost either way. Makes a backup first as
backend/exporter_backup7.py.
"""

import shutil
import sys

PATH = "backend/exporter.py"
BACKUP = "backend/exporter_backup7.py"

try:
    with open(PATH, "r", encoding="utf-8") as f:
        content = f.read()
except FileNotFoundError:
    print(f"ERROR: could not find {PATH}")
    print("Make sure you're running this from inside the DataFusionPlatform_v6 folder.")
    sys.exit(1)

# The buggy version from patch_exporter_6.py
old_block = '''    if "Account" in columns and "Account (2)" in columns:
        columns = [c for c in columns if c != "Account"]
        rows = [{k: v for k, v in row.items() if k != "Account"} for row in rows]'''

new_block = '''    if "Account" in columns and "Account (2)" in columns:
        # Merge per row instead of dropping one column outright -- some
        # source files only put their real account label in "Account"
        # (no "Account (2)" at all in that file), others only put it in
        # "Account (2)". Keep whichever one actually has a value.
        for row in rows:
            merged_val = row.get("Account (2)") or row.get("Account") or ""
            row["Account"] = merged_val
            row.pop("Account (2)", None)
        columns = [c for c in columns if c != "Account (2)"]'''

if old_block in content:
    content = content.replace(old_block, new_block)
elif "merged_val = row.get" in content:
    print("Nothing to change -- this fix is already applied.")
    sys.exit(0)
else:
    print("ERROR: Could not find the expected code block to change.")
    print("This usually means patch_exporter_6.py hasn't been run on this file,")
    print("or this file is a different version than expected.")
    print("No changes were made. Please share this message with support.")
    sys.exit(1)

shutil.copy(PATH, BACKUP)
with open(PATH, "w", encoding="utf-8") as f:
    f.write(content)

print("SUCCESS. 'Account' and 'Account (2)' are now merged per row -- no data lost either way.")
print(f"Backup of the previous version saved as {BACKUP}")
print("Now run: python app.py  and generate a fresh report.")
