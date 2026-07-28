"""
patch_exporter_4.py

Run this once from inside your DataFusionPlatform_v6 folder:

    python patch_exporter_4.py

Adds "Total Debit" and "Total Credit" columns to the Company Summary
sheet, computed per company from the Debit/Credit columns in your
consolidated data -- only shown if your data actually has Debit and
Credit columns (e.g. trial balance style data). Makes a backup first
as backend/exporter_backup4.py.
"""

import shutil
import sys

PATH = "backend/exporter.py"
BACKUP = "backend/exporter_backup4.py"

try:
    with open(PATH, "r", encoding="utf-8") as f:
        content = f.read()
except FileNotFoundError:
    print(f"ERROR: could not find {PATH}")
    print("Make sure you're running this from inside the DataFusionPlatform_v6 folder.")
    sys.exit(1)

old_block = '''        ws_co = wb.create_sheet("Company Summary")
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
        # If the data has Debit/Credit columns (e.g. a trial balance),
        # add per-company Total Debit / Total Credit columns too.
        co_debit_credit_mapping = detect_columns(columns)
        co_debit_col = co_debit_credit_mapping.get("debit")
        co_credit_col = co_debit_credit_mapping.get("credit")
        has_debit_credit = bool(co_debit_col and co_credit_col)
        co_debit_totals = {}
        co_credit_totals = {}
        if has_debit_credit:
            for row in rows:
                name = row.get("Company Name", "")
                co_debit_totals[name] = co_debit_totals.get(name, 0.0) + _to_number(row.get(co_debit_col))
                co_credit_totals[name] = co_credit_totals.get(name, 0.0) + _to_number(row.get(co_credit_col))

        co_columns = ["Company Name"]
        if has_tax_data:
            co_columns += ["Taxable Total", "Tax Total"]
        if has_debit_credit:
            co_columns += ["Total Debit", "Total Credit"]
        co_columns.append("Row Count")

        co_rows = []
        for name, vals in co_companies.items():
            row = {"Company Name": name, "Row Count": vals["row_count"]}
            if has_tax_data:
                row["Taxable Total"] = vals["taxable_total"]
                row["Tax Total"] = vals["tax_total"]
            if has_debit_credit:
                row["Total Debit"] = round(co_debit_totals.get(name, 0.0), 2)
                row["Total Credit"] = round(co_credit_totals.get(name, 0.0), 2)
            co_rows.append(row)
        _write_table(ws_co, co_columns, co_rows, add_grand_total=False)'''

if old_block not in content:
    if "co_debit_credit_mapping" in content:
        print("Nothing to change -- this fix is already applied.")
        sys.exit(0)
    print("ERROR: Could not find the expected code block to change.")
    print("This usually means patch_exporter_3.py hasn't been run yet on")
    print("this file, or this file is a different version than expected.")
    print("No changes were made. Please share this message with support.")
    sys.exit(1)

content = content.replace(old_block, new_block)

shutil.copy(PATH, BACKUP)
with open(PATH, "w", encoding="utf-8") as f:
    f.write(content)

print("SUCCESS. Company Summary will now show Total Debit / Total Credit per company.")
print(f"Backup of the previous version saved as {BACKUP}")
print("Now run: python app.py  and generate a fresh report.")
