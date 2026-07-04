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

# GSTIN structure: 2 digit state code + 10 char PAN + 1 entity code
# + 1 default 'Z' + 1 checksum character. We validate the shape.
GSTIN_PATTERN = re.compile(r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$")


def is_valid_gstin(gstin):
    """Returns True if the GSTIN matches the standard 15-character format."""
    if not isinstance(gstin, str):
        return False
    gstin = gstin.strip().upper()
    return bool(GSTIN_PATTERN.match(gstin))


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
    required_fields = ["invoice_no", "taxable_amount"]
    for field in required_fields:
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
