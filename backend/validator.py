"""
validator.py

Runs data-quality checks on an uploaded sheet before it's trusted
for consolidation or reconciliation. Covers:

  1. GSTIN format/checksum validation
  2. Required-field completeness (blank invoice no, blank amount, etc.)
  3. Duplicate invoice detection (exact duplicate rows AND same invoice
     number with different data, which is more dangerous)
  4. Missing invoice number detection (gaps in a numeric sequence)

Every check is skipped gracefully (with a note) if the relevant
column couldn't be detected in the uploaded file, rather than
crashing.
"""

import re
from backend.column_mapper import detect_columns
from backend.filters import _to_number

# GSTIN structure: 2 digit state code + 10 char PAN + 1 entity code
# + 1 default 'Z' + 1 checksum character. We validate the shape.
GSTIN_PATTERN = re.compile(r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$")


def is_valid_gstin(gstin):
    """Returns True if the GSTIN matches the standard 15-character format."""
    if not isinstance(gstin, str):
        return False
    gstin = gstin.strip().upper()
    return bool(GSTIN_PATTERN.match(gstin))


def review_sheet(data, columns):
    """
    Lightweight pre-consolidation check used by the "Review Data" screen.
    Flags columns that came in with no header (excel_engine.py already
    renames these to "No heading in source file" at import time, so we
    just look for that marker), columns that are entirely
    empty, and rows that are entirely empty -- the sort of thing you
    want to catch by eye before consolidating, rather than after.

    Never raises: every value is safely stringified before comparison,
    so unusual cell types (numbers, dates, None) can't crash this.
    """

    columns = list(columns or [])
    data = data or []

    missing_header_columns = [c for c in columns if str(c).startswith("No heading in source file")]

    empty_columns = []
    for col in columns:
        if col in missing_header_columns:
            continue
        if data and all(str(row.get(col, "")).strip() == "" for row in data):
            empty_columns.append(col)

    empty_rows = []
    for i, row in enumerate(data):
        if columns and all(str(row.get(col, "")).strip() == "" for col in columns):
            empty_rows.append(i)

    return {
        "missing_header_columns": missing_header_columns,
        "empty_columns": empty_columns,
        "empty_rows": empty_rows,
        "row_count": len(data),
        "column_count": len(columns),
    }


def validate_sheet(data, columns):
    """
    data: list of row dicts (as produced by excel_engine.py)
    columns: list of column names for this sheet

    Returns a report dict with detected column mapping, skipped
    checks, per-issue-type lists of problems, and a summary string.
    """

    mapping = detect_columns(columns)
    skipped = []
    issues = {
        "missing_required_fields": [],
        "invalid_gstin": [],
        "exact_duplicate_rows": [],
        "duplicate_invoice_numbers": [],
        "missing_invoice_numbers": [],
    }

    # ---- 1. Required field completeness ----
    # This check only makes sense for an invoice-level GST register --
    # an "Invoice No" column is what actually marks a sheet as one. A
    # trial balance, balance sheet, or fixed-asset register can easily
    # have its own column literally named "Amount" (e.g. a fixed asset
    # purchase amount) -- that's not a taxable/GST value, it just
    # happens to share the word "amount". Flagging every blank cell in
    # that column as a missing "Taxable Amount" would be a false
    # alarm on a sheet that was never a GST document to begin with, so
    # skip this whole section unless the sheet has an invoice number
    # column (the one reliable signal it's actually a GST-style sheet).
    is_gst_style_sheet = "invoice_no" in mapping
    required_fields = ["invoice_no", "taxable_amount"]
    for field in required_fields:
        if not is_gst_style_sheet:
            skipped.append(f"required-field check for '{field}' (not a GST-style invoice sheet)")
            continue
        if field not in mapping:
            skipped.append(f"required-field check for '{field}' (column not found)")
            continue
        col = mapping[field]
        for i, row in enumerate(data):
            value = row.get(col, "")
            if value == "" or value is None:
                issues["missing_required_fields"].append({"row": i, "field": field})

    # ---- 2. GSTIN validation ----
    if "gstin" not in mapping:
        skipped.append("GSTIN validation (column not found)")
    else:
        col = mapping["gstin"]
        for i, row in enumerate(data):
            value = row.get(col, "")
            if value == "" or value is None:
                continue
            if not is_valid_gstin(str(value)):
                issues["invalid_gstin"].append({"row": i, "value": value})

    # ---- 3. Exact duplicate rows ----
    seen = {}
    for i, row in enumerate(data):
        key = tuple(sorted(row.items()))
        seen.setdefault(key, []).append(i)
    for key, rows in seen.items():
        if len(rows) > 1:
            issues["exact_duplicate_rows"].append(rows)

    # ---- 4. Duplicate invoice numbers ----
    if "invoice_no" not in mapping:
        skipped.append("duplicate/missing invoice number checks (column not found)")
    else:
        col = mapping["invoice_no"]
        by_invoice = {}
        for i, row in enumerate(data):
            inv = str(row.get(col, "")).strip()
            if inv == "":
                continue
            by_invoice.setdefault(inv, []).append(i)

        for inv, rows in by_invoice.items():
            if len(rows) > 1:
                issues["duplicate_invoice_numbers"].append({"invoice_no": inv, "rows": rows})

        # ---- 5. Missing invoice numbers (gap detection) ----
        numeric_invoices = []
        prefix = None
        width = None
        for inv in by_invoice.keys():
            m = re.match(r"^([A-Za-z\-\/]*)(\d+)$", inv)
            if m:
                this_prefix, num_str = m.group(1), m.group(2)
                if prefix is None:
                    prefix = this_prefix
                    width = len(num_str)
                if this_prefix == prefix:
                    numeric_invoices.append(int(num_str))

        if len(numeric_invoices) >= 2:
            numeric_invoices.sort()
            for n in range(numeric_invoices[0], numeric_invoices[-1] + 1):
                if n not in numeric_invoices:
                    missing_id = f"{prefix}{str(n).zfill(width)}"
                    issues["missing_invoice_numbers"].append(missing_id)
        else:
            skipped.append("missing-invoice-number gap detection (numbers not in a consistent pattern)")

    total_issues = sum(len(v) for v in issues.values())
    summary = f"{total_issues} issue(s) found across {len(data)} row(s)."
    if skipped:
        summary += f" {len(skipped)} check(s) skipped due to missing columns."

    return {
        "mapping": mapping,
        "skipped_checks": skipped,
        "issues": issues,
        "summary": summary,
    }


def humanize_issue(issue_type, item):
    """
    Turns one raw issue entry (as produced above -- a dict, a list of row
    numbers, or a plain string, depending on issue_type) into a plain
    English sentence for the exported report, instead of a raw Python
    dict/list printed as text (e.g. "{'row': 4, 'field': 'taxable_amount'}").

    Row numbers are shown 1-based ("Row 1" = the first data row under the
    header), which is how a non-technical reader expects rows to be
    counted, rather than the 0-based index used internally.
    """
    def _field_label(field):
        return str(field).replace("_", " ").strip().title()

    try:
        if issue_type == "missing_required_fields" and isinstance(item, dict):
            return f"Row {item['row'] + 1}: '{_field_label(item.get('field'))}' is missing or blank."

        if issue_type == "invalid_gstin" and isinstance(item, dict):
            return f"Row {item['row'] + 1}: GSTIN '{item.get('value')}' is not a valid GSTIN format."

        if issue_type == "exact_duplicate_rows" and isinstance(item, (list, tuple)):
            row_nums = [r + 1 for r in item]
            if len(row_nums) == 2:
                return f"Rows {row_nums[0]} and {row_nums[1]} are exact duplicates of each other."
            row_list = ", ".join(str(r) for r in row_nums[:-1]) + f" and {row_nums[-1]}"
            return f"Rows {row_list} are all exact duplicates of each other."

        if issue_type == "duplicate_invoice_numbers" and isinstance(item, dict):
            row_nums = [r + 1 for r in item.get("rows", [])]
            row_list = ", ".join(str(r) for r in row_nums)
            return f"Invoice number '{item.get('invoice_no')}' is repeated on rows {row_list}."

        if issue_type == "missing_invoice_numbers":
            return f"Invoice number '{item}' is missing from the sequence (expected but not found)."
    except Exception:
        pass

    # Fallback for any unrecognised shape -- still better than a raw repr.
    return str(item)


def check_debit_credit_balance(file_results):
    """
    Pre-consolidation checkpoint: for every sheet across the given
    file results, detect Debit / Credit columns (via column_mapper)
    and total them up. Used to warn the user -- before they consolidate
    -- if debits and credits don't tie out, and exactly which
    file/sheet is responsible.

    file_results: list of result dicts, each shaped like:
        {"file_name": ..., "file_path": ..., "sheets": {sheet_name: {"columns": [...], "data": [...]}}}

    Returns:
        {
            "checked": bool,       # True if at least one sheet had BOTH a debit
                                    # and credit column detected
            "balanced": bool,      # True if totals match (within a cent)
            "total_debit": float,
            "total_credit": float,
            "difference": float,
            "mismatches": [
                {
                    "file_name": str, "file_path": str, "sheet_name": str,
                    "debit_column": str, "credit_column": str,
                    "debit_total": float, "credit_total": float,
                    "difference": float, "row_count": int,
                },
                ...
            ],
        }

    Sheets where a debit or credit column simply couldn't be found are
    silently skipped (not every sheet is a ledger), rather than being
    reported as a mismatch.
    """
    total_debit = 0.0
    total_credit = 0.0
    mismatches = []
    checked_any = False

    for result in file_results or []:
        fname = result.get("file_name", "")
        fpath = result.get("file_path", "")
        for sheet_name, sheet in (result.get("sheets") or {}).items():
            columns = sheet.get("columns", [])
            data = sheet.get("data", [])
            mapping = detect_columns(columns)

            debit_col = mapping.get("debit")
            credit_col = mapping.get("credit")
            if not debit_col or not credit_col:
                continue

            checked_any = True
            sheet_debit = sum(_to_number(row.get(debit_col), 0) or 0 for row in data)
            sheet_credit = sum(_to_number(row.get(credit_col), 0) or 0 for row in data)

            total_debit += sheet_debit
            total_credit += sheet_credit

            diff = round(sheet_debit - sheet_credit, 2)
            if abs(diff) > 0.01:
                mismatches.append({
                    "file_name": fname,
                    "file_path": fpath,
                    "sheet_name": sheet_name,
                    "debit_column": debit_col,
                    "credit_column": credit_col,
                    "debit_total": sheet_debit,
                    "credit_total": sheet_credit,
                    "difference": diff,
                    "row_count": len(data),
                })

    overall_diff = round(total_debit - total_credit, 2)

    return {
        "checked": checked_any,
        "balanced": checked_any and abs(overall_diff) <= 0.01,
        "total_debit": total_debit,
        "total_credit": total_credit,
        "difference": overall_diff,
        "mismatches": mismatches,
    }