"""
patch_exporter.py

Run this once from inside your DataFusionPlatform_v6 folder:

    python patch_exporter.py

It edits backend/exporter.py in place to remove the GST Summary,
Monthly Summary, and Party Summary sheets from the exported report
(and from the Index / Report Guide sheet). Makes a backup copy first
as backend/exporter_backup.py.
"""

import re
import shutil
import sys

PATH = "backend/exporter.py"
BACKUP = "backend/exporter_backup.py"

try:
    with open(PATH, "r", encoding="utf-8") as f:
        content = f.read()
except FileNotFoundError:
    print(f"ERROR: could not find {PATH}")
    print("Make sure you're running this from inside the DataFusionPlatform_v6 folder.")
    sys.exit(1)

original_content = content
changes_made = []

# 1) Remove the 3 lines from the docstring listing (items 6,7,8) and
#    renumber "9. Issues Found" down to "6. Issues Found".
pattern1 = re.compile(
    r"\n\s*6\.\s*GST Summary.*?\n\s*7\.\s*Monthly Summary.*?\n\s*8\.\s*Party Summary.*?\n(\s*)9\.\s*Issues Found",
    re.DOTALL,
)
new_content, n = pattern1.subn(r"\n\g<1>6. Issues Found", content)
if n:
    content = new_content
    changes_made.append("docstring list")

# 2) Remove the 3 guide rows shown on the Index sheet.
pattern2 = re.compile(
    r'\s*\("GST Summary",.*?\),\s*\n\s*\("Monthly Summary",.*?\),\s*\n\s*\("Party Summary",.*?\),\n',
    re.DOTALL,
)
new_content, n = pattern2.subn("\n", content)
if n:
    content = new_content
    changes_made.append("guide rows")

# 3) Remove the actual sheet-generation code blocks for all three sheets.
pattern3 = re.compile(
    r'\s*# ---------- 6\. GST Summary sheet ----------.*?'
    r'# ---------- 9\. Issues Found sheet ----------',
    re.DOTALL,
)
new_content, n = pattern3.subn("\n    # ---------- 6. Issues Found sheet ----------", content)
if n:
    content = new_content
    changes_made.append("sheet-generation code")

if not changes_made:
    if "GST Summary" not in content:
        print("Nothing to change -- this file already has GST/Monthly/Party Summary removed.")
        sys.exit(0)
    else:
        print("ERROR: Could not find the expected code patterns to remove.")
        print("This file may already be a different version than expected.")
        print("No changes were made. Please share this message with support.")
        sys.exit(1)

shutil.copy(PATH, BACKUP)
with open(PATH, "w", encoding="utf-8") as f:
    f.write(content)

print("SUCCESS. Changes made to:", ", ".join(changes_made))
print(f"Backup of the original saved as {BACKUP}")
print("Now run: python app.py  and generate a fresh report.")