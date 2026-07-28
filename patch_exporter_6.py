"""
patch_exporter_6.py

Run this once from inside your DataFusionPlatform_v6 folder:

    python patch_exporter_6.py

Removes the "Account" column from the Consolidated Data sheet (and
everything derived from it) whenever the file also has an "Account
(2)" column -- since in that case "Account" was the mostly-blank,
redundant one causing the TOTAL row label to visually land in the
wrong column. Nothing else changes. If a file only has "Account" and
no "Account (2)", it's left alone. Makes a backup first as
backend/exporter_backup6.py.
"""

import shutil
import sys

PATH = "backend/exporter.py"
BACKUP = "backend/exporter_backup6.py"

try:
    with open(PATH, "r", encoding="utf-8") as f:
        content = f.read()
except FileNotFoundError:
    print(f"ERROR: could not find {PATH}")
    print("Make sure you're running this from inside the DataFusionPlatform_v6 folder.")
    sys.exit(1)

old_block = '''    # Company Name always leads the sheet since every row is tagged with it.
    if "Company Name" in columns:
        columns = ["Company Name"] + [c for c in columns if c != "Company Name"]
    has_source = bool(rows) and "__source_file" in rows[0]
    display_columns = columns + (["Source File"] if has_source else [])'''

new_block = '''    # Company Name always leads the sheet since every row is tagged with it.
    if "Company Name" in columns:
        columns = ["Company Name"] + [c for c in columns if c != "Company Name"]
    # "Account" is redundant/mostly blank whenever a more specific
    # "Account (2)" column is also present -- drop it so the TOTAL row
    # label doesn't visually land in the wrong column.
    if "Account" in columns and "Account (2)" in columns:
        columns = [c for c in columns if c != "Account"]
        rows = [{k: v for k, v in row.items() if k != "Account"} for row in rows]
    has_source = bool(rows) and "__source_file" in rows[0]
    display_columns = columns + (["Source File"] if has_source else [])'''

if old_block not in content:
    if '"Account" in columns and "Account (2)" in columns' in content:
        print("Nothing to change -- this fix is already applied.")
        sys.exit(0)
    print("ERROR: Could not find the expected code block to change.")
    print("No changes were made. Please share this message with support.")
    sys.exit(1)

content = content.replace(old_block, new_block)

shutil.copy(PATH, BACKUP)
with open(PATH, "w", encoding="utf-8") as f:
    f.write(content)

print("SUCCESS. The redundant 'Account' column will now be dropped when 'Account (2)' is also present.")
print(f"Backup of the previous version saved as {BACKUP}")
print("Now run: python app.py  and generate a fresh report.")
