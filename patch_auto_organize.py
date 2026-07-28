"""
patch_auto_organize.py

Run this once from inside your DataFusionPlatform_v6 folder:

    python patch_auto_organize.py

When you use Auto-Organize on a messy file, any column that has no
header text AND no data in it at all will now be dropped silently
instead of being kept and shown as a red-flagged "No heading in
source file" column that needs manual cleanup on the Preview screen.
Columns with no header but that DO have data are still kept and named
as before. Makes a backup first as backend/auto_organize_backup.py.
"""

import shutil
import sys

PATH = "backend/auto_organize.py"
BACKUP = "backend/auto_organize_backup.py"

try:
    with open(PATH, "r", encoding="utf-8") as f:
        content = f.read()
except FileNotFoundError:
    print(f"ERROR: could not find {PATH}")
    print("Make sure you're running this from inside the DataFusionPlatform_v6 folder.")
    sys.exit(1)

old_block = '''    # Name each kept column: use its own header text if present,
    # otherwise guess "Account" for a text-heavy label column, else a
    # generic placeholder.
    final_columns = []
    for c in kept_cols:
        text = str(header_row[c]).strip()
        if text:
            final_columns.append(text)
        else:
            data_vals = [str(r[c]).strip() for r in rows[header_idx + 1:] if c < len(r) and str(r[c]).strip() != ""]
            looks_textual = any(not v.replace(",", "").replace(".", "").replace("-", "").isdigit() for v in data_vals)
            final_columns.append("Account" if looks_textual and data_vals else NO_HEADER_TEXT)
    final_columns = _clean_headers(final_columns)'''

new_block = '''    # Name each kept column: use its own header text if present,
    # otherwise guess "Account" for a text-heavy label column. A column
    # with no header text AND no data in it at all is dropped silently
    # here -- there's nothing useful to show or ask the user to name,
    # so it shouldn't turn into a red-flagged "No heading in source
    # file" column on the Preview screen.
    final_columns = []
    real_kept_cols = []
    for c in kept_cols:
        text = str(header_row[c]).strip()
        if text:
            final_columns.append(text)
            real_kept_cols.append(c)
            continue
        data_vals = [str(r[c]).strip() for r in rows[header_idx + 1:] if c < len(r) and str(r[c]).strip() != ""]
        if not data_vals:
            continue  # fully empty, unnamed column -- drop silently
        looks_textual = any(not v.replace(",", "").replace(".", "").replace("-", "").isdigit() for v in data_vals)
        final_columns.append("Account" if looks_textual else NO_HEADER_TEXT)
        real_kept_cols.append(c)
    kept_cols = real_kept_cols
    final_columns = _clean_headers(final_columns)'''

if old_block not in content:
    if "real_kept_cols" in content:
        print("Nothing to change -- this fix is already applied.")
        sys.exit(0)
    print("ERROR: Could not find the expected code block to change.")
    print("No changes were made. Please share this message with support.")
    sys.exit(1)

content = content.replace(old_block, new_block)

shutil.copy(PATH, BACKUP)
with open(PATH, "w", encoding="utf-8") as f:
    f.write(content)

print("SUCCESS. Fully-empty unnamed columns are now dropped silently by Auto-Organize.")
print(f"Backup of the previous version saved as {BACKUP}")
print("Now run: python app.py  and try Auto-Organize on that file again.")
