"""
patch_exporter_3.py

Run this once from inside your DataFusionPlatform_v6 folder:

    python patch_exporter_3.py

Company Summary will now skip the "Taxable Total" and "Tax Total"
columns whenever every company's totals come out as 0.00 (i.e. the
imported data has no GST/tax columns to begin with, like a plain
trial balance) -- it'll just show Company Name and Row Count instead.
If any company has a real non-zero total, all 4 columns still show
as before. Makes a backup first as backend/exporter_backup3.py.
"""

import shutil
import sys

PATH = "backend/exporter.py"
BACKUP = "backend/exporter_backup3.py"

try:
    with open(PATH, "r", encoding="utf-8") as f:
        content = f.read()
except FileNotFoundError:
    print(f"ERROR: could not find {PATH}")
    print("Make sure you're running this from inside the DataFusionPlatform_v6 folder.")
    sys.exit(1)

old_block = '''        ws_co = wb.create_sheet("Company Summary")
        co_rows = []
        for name, vals in company_summary["companies"].items():
            co_rows.append({
                "Company Name": name,
                "Taxable Total": vals["taxable_total"],
                "Tax Total": vals["tax_total"],
                "Row Count": vals["row_count"],
            })
        _write_table(ws_co, ["Company Name", "Taxable Total", "Tax Total", "Row Count"], co_rows, add_grand_total=False)'''

new_block = '''        ws_co = wb.create_sheet("Company Summary")
        co_companies = company_summary["companies"]
        # Skip the Taxable/Tax Total columns entirely if every company's
        # total is zero -- that means the imported data had no GST/tax
        # columns to begin with (e.g. a plain trial balance), so showing
        # two all-zero columns would just be noise.
        has_tax_data = any(
            vals.get("taxable_total") or vals.get("tax_total")
            for vals in co_companies.values()
        )
        if has_tax_data:
            co_columns = ["Company Name", "Taxable Total", "Tax Total", "Row Count"]
        else:
            co_columns = ["Company Name", "Row Count"]
        co_rows = []
        for name, vals in co_companies.items():
            row = {"Company Name": name, "Row Count": vals["row_count"]}
            if has_tax_data:
                row["Taxable Total"] = vals["taxable_total"]
                row["Tax Total"] = vals["tax_total"]
            co_rows.append(row)
        _write_table(ws_co, co_columns, co_rows, add_grand_total=False)'''

if old_block not in content:
    if "has_tax_data" in content:
        print("Nothing to change -- this fix is already applied.")
        sys.exit(0)
    print("ERROR: Could not find the expected code block to change.")
    print("No changes were made. Please share this message with support.")
    sys.exit(1)

content = content.replace(old_block, new_block)

shutil.copy(PATH, BACKUP)
with open(PATH, "w", encoding="utf-8") as f:
    f.write(content)

print("SUCCESS. Company Summary will now hide Taxable/Tax Total columns when they're all zero.")
print(f"Backup of the previous version saved as {BACKUP}")
print("Now run: python app.py  and generate a fresh report.")
