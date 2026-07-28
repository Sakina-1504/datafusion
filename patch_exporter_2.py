"""
patch_exporter_2.py

Run this once from inside your DataFusionPlatform_v6 folder:

    python patch_exporter_2.py

Makes the Index / Report Guide sheet only list "Issues Found" as a row
when the report actually has data-quality issues in it (i.e. only when
the Issues Found tab actually exists). Makes a backup first as
backend/exporter_backup2.py.
"""

import shutil
import sys

PATH = "backend/exporter.py"
BACKUP = "backend/exporter_backup2.py"

try:
    with open(PATH, "r", encoding="utf-8") as f:
        content = f.read()
except FileNotFoundError:
    print(f"ERROR: could not find {PATH}")
    print("Make sure you're running this from inside the DataFusionPlatform_v6 folder.")
    sys.exit(1)

old_guide_block = '''    guide_rows = [
        ("Summary", "Headline totals, files used, credit/debit checkpoint, CGST/SGST/IGST and data-quality snapshot."),
        ("Source Files", "Every source file (sub file) that went into this report, with the Company Name assigned to it and how many rows it contributed. Click a 'Source File' cell on Consolidated Data to jump here."),
        ("Consolidated Data", "Every row from every imported file, merged, with a Company Name column, a clickable Source File reference, and a Grand Total row."),
        ("Company Summary", "Taxable value, tax and row count per company -- useful when you're handling more than one client at once."),
        ("Issues Found", "Every data-quality issue detected (missing fields, invalid GSTIN, duplicates, tax mismatches) with its exact file/sheet/row reference."),
    ]'''

new_guide_block = '''    guide_rows = [
        ("Summary", "Headline totals, files used, credit/debit checkpoint, CGST/SGST/IGST and data-quality snapshot."),
        ("Source Files", "Every source file (sub file) that went into this report, with the Company Name assigned to it and how many rows it contributed. Click a 'Source File' cell on Consolidated Data to jump here."),
        ("Consolidated Data", "Every row from every imported file, merged, with a Company Name column, a clickable Source File reference, and a Grand Total row."),
        ("Company Summary", "Taxable value, tax and row count per company -- useful when you're handling more than one client at once."),
    ]
    # "Issues Found" is only listed in the guide if the report actually
    # ends up with an Issues Found sheet (i.e. there's at least one
    # real data-quality issue to show).
    _has_issues = bool(validation_issues) and any(
        len(v) for v in validation_issues.values()
    ) if isinstance(validation_issues, dict) else bool(validation_issues)
    if _has_issues:
        guide_rows.append(
            ("Issues Found", "Every data-quality issue detected (missing fields, invalid GSTIN, duplicates, tax mismatches) with its exact file/sheet/row reference.")
        )'''

if old_guide_block not in content:
    if "_has_issues" in content:
        print("Nothing to change -- this fix is already applied.")
        sys.exit(0)
    print("ERROR: Could not find the expected code block to change.")
    print("No changes were made. Please share this message with support.")
    sys.exit(1)

content = content.replace(old_guide_block, new_guide_block)

shutil.copy(PATH, BACKUP)
with open(PATH, "w", encoding="utf-8") as f:
    f.write(content)

print("SUCCESS. The Index guide will now skip 'Issues Found' when there are no issues.")
print(f"Backup of the previous version saved as {BACKUP}")
print("Now run: python app.py  and generate a fresh report.")
