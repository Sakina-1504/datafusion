"""
consolidator.py

Merges multiple uploaded Excel files/sheets into one dataset, and
produces the summary numbers a CA would otherwise build by hand:

  1. Merge multiple files into a single combined table
  2. GST rate-wise summary (group by tax rate e.g. 5%/12%/18%/28%)
  3. CGST/SGST/IGST totals
  4. Tax-calculation mismatch check (declared tax vs taxable value x rate)
  5. Party-wise (vendor/customer) totals

As with validator.py, every step gracefully skips (with a note)
when a required column can't be detected, instead of crashing.
"""

import os
from datetime import datetime
from backend.column_mapper import detect_columns

# Tolerance for rounding differences when comparing declared tax
# against a recomputed taxable_value * rate. Real-world files round
# to the nearest rupee, so small differences aren't real errors.
ROUNDING_TOLERANCE = 1.0

COMPANY_COLUMN = "Company Name"


def _to_number(value, default=0.0):
    try:
        if value == "" or value is None:
            return default
        return float(value)
    except (ValueError, TypeError):
        return default


def merge_files(file_results, company_names=None):
    """
    file_results: list of structured results from ExcelEngine.load_excel(),
    each shaped like {"file_name": ..., "sheets": {sheet_name: {"data":..., "columns":...}}}

    company_names: optional dict of {file_name: "Company Name"} entered
    by the user when importing each file (see ui/dashboard.py). Every
    row coming from that file gets a real, always-filled "Company Name"
    column -- never left blank -- so the consolidated report and every
    export clearly show which company each row belongs to. If no name
    was given for a file, its own file_result["company_name"] is used,
    and if that's blank too we fall back to the file name itself so the
    column is still never empty.

    Returns a single flat list of rows (each row tagged with its
    source file/sheet so you can always trace it back), plus the
    union of all columns seen, with "Company Name" always placed first.
    """

    company_names = company_names or {}
    merged_rows = []
    all_columns = [COMPANY_COLUMN]

    for file_result in file_results:
        file_name = file_result.get("file_name", "unknown")

        company = (
            company_names.get(file_name)
            or file_result.get("company_name")
            or os.path.splitext(file_name)[0]
        ).strip() or os.path.splitext(file_name)[0]

        for sheet_name, sheet in file_result.get("sheets", {}).items():
            for col in sheet["columns"]:
                if col not in all_columns:
                    all_columns.append(col)
            for row in sheet["data"]:
                tagged_row = {COMPANY_COLUMN: company}
                tagged_row.update(row)
                tagged_row["__source_file"] = file_name
                tagged_row["__source_sheet"] = sheet_name
                merged_rows.append(tagged_row)

    return {"rows": merged_rows, "columns": all_columns}


def company_wise_summary(data, columns):
    """Totals taxable value and tax per company -- useful the moment
    you're consolidating data for more than one client/entity at once."""

    mapping = detect_columns(columns)
    taxable_col = mapping.get("taxable_amount")
    gst_col = mapping.get("gst_amount")

    companies = {}
    for row in data:
        name = str(row.get(COMPANY_COLUMN, "")).strip() or "(Unknown Company)"
        c = companies.setdefault(name, {"taxable_total": 0.0, "tax_total": 0.0, "row_count": 0})
        if taxable_col:
            c["taxable_total"] += _to_number(row.get(taxable_col))
        if gst_col:
            c["tax_total"] += _to_number(row.get(gst_col))
        c["row_count"] += 1

    for c in companies.values():
        c["taxable_total"] = round(c["taxable_total"], 2)
        c["tax_total"] = round(c["tax_total"], 2)

    return {"companies": companies, "mapping": mapping}


def month_wise_summary(data, columns):
    """
    Groups rows by month (from the detected date column) and totals
    taxable value + tax per month -- the number a CA needs for
    month-wise GST filing trend or a management report.
    Skips gracefully (with a note) if no usable date column is found.
    """

    mapping = detect_columns(columns)
    if "date" not in mapping:
        return {"error": "Cannot compute month-wise summary: no date column found."}

    date_col = mapping["date"]
    taxable_col = mapping.get("taxable_amount")
    gst_col = mapping.get("gst_amount")

    months = {}
    unparsed = 0
    for row in data:
        raw = row.get(date_col, "")
        month_key = _parse_month(raw)
        if month_key is None:
            unparsed += 1
            continue
        m = months.setdefault(month_key, {"taxable_total": 0.0, "tax_total": 0.0, "row_count": 0})
        if taxable_col:
            m["taxable_total"] += _to_number(row.get(taxable_col))
        if gst_col:
            m["tax_total"] += _to_number(row.get(gst_col))
        m["row_count"] += 1

    for m in months.values():
        m["taxable_total"] = round(m["taxable_total"], 2)
        m["tax_total"] = round(m["tax_total"], 2)

    return {"months": months, "unparsed_dates": unparsed, "mapping": mapping}


def _parse_month(raw_value):
    """Best-effort conversion of a date cell (datetime, string, excel
    serial, etc.) into a sortable 'YYYY-Mon' label. Returns None if it
    can't confidently parse it, rather than guessing wrong."""
    if raw_value in ("", None):
        return None
    if isinstance(raw_value, datetime):
        return raw_value.strftime("%Y-%b")
    if hasattr(raw_value, "strftime"):
        try:
            return raw_value.strftime("%Y-%b")
        except Exception:
            pass
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%b-%Y", "%d %b %Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(str(raw_value).strip(), fmt).strftime("%Y-%b")
        except (ValueError, TypeError):
            continue
    return None


def gst_rate_summary(data, columns):
    """
    Groups rows by GST rate (computed from gst_amount / taxable_amount
    if there's no explicit rate column) and totals taxable value + tax
    per rate bucket. This is the number that goes straight into GSTR-1
    rate-wise filing.

    Also flags rows where the declared tax doesn't match
    taxable_value x rate within ROUNDING_TOLERANCE — a common manual
    entry error.
    """

    mapping = detect_columns(columns)
    skipped = []

    if "taxable_amount" not in mapping:
        return {"error": "Cannot compute GST summary: no taxable amount column found."}
    taxable_col = mapping["taxable_amount"]

    has_rate_col = "gst_rate" in mapping
    has_gst_amount_col = "gst_amount" in mapping
    has_split = "cgst" in mapping or "sgst" in mapping or "igst" in mapping

    if not has_rate_col and not has_gst_amount_col:
        skipped.append("GST rate summary (no rate or GST-amount column found)")

    buckets = {}
    mismatches = []
    cgst_total = sgst_total = igst_total = 0.0

    for i, row in enumerate(data):
        taxable = _to_number(row.get(taxable_col))

        if has_rate_col:
            rate = _to_number(row.get(mapping["gst_rate"]))
            gst_amt = round(taxable * rate / 100, 2)
        elif has_gst_amount_col:
            gst_amt = _to_number(row.get(mapping["gst_amount"]))
            rate = round((gst_amt / taxable) * 100, 2) if taxable else 0.0
        else:
            continue

        # bucket by rounded rate (5, 12, 18, 28 etc.)
        rate_bucket = round(rate)
        b = buckets.setdefault(rate_bucket, {"taxable_total": 0.0, "tax_total": 0.0, "row_count": 0})
        b["taxable_total"] += taxable
        b["tax_total"] += gst_amt
        b["row_count"] += 1

        if has_split:
            cgst_total += _to_number(row.get(mapping.get("cgst"), 0))
            sgst_total += _to_number(row.get(mapping.get("sgst"), 0))
            igst_total += _to_number(row.get(mapping.get("igst"), 0))

        # mismatch check only meaningful when we have an explicit rate col
        # AND an explicit gst_amount col to cross-check against each other
        if has_rate_col and has_gst_amount_col:
            declared = _to_number(row.get(mapping["gst_amount"]))
            expected = round(taxable * rate / 100, 2)
            if abs(declared - expected) > ROUNDING_TOLERANCE:
                mismatches.append({
                    "row": i,
                    "declared_tax": declared,
                    "expected_tax": expected,
                    "difference": round(declared - expected, 2),
                })

    if not (has_rate_col and has_gst_amount_col):
        skipped.append("tax-mismatch cross-check (needs both a rate column and a GST-amount column)")
    if not has_split:
        skipped.append("CGST/SGST/IGST totals (no split columns found)")

    return {
        "rate_wise": buckets,
        "cgst_total": round(cgst_total, 2),
        "sgst_total": round(sgst_total, 2),
        "igst_total": round(igst_total, 2),
        "tax_mismatches": mismatches,
        "skipped_checks": skipped,
        "mapping": mapping,
    }


def party_wise_summary(data, columns):
    """Totals taxable value and tax per customer/vendor (party_name column)."""

    mapping = detect_columns(columns)
    if "party_name" not in mapping:
        return {"error": "Cannot compute party-wise summary: no party/customer name column found."}

    party_col = mapping["party_name"]
    taxable_col = mapping.get("taxable_amount")
    gst_col = mapping.get("gst_amount")

    parties = {}
    for row in data:
        name = str(row.get(party_col, "")).strip() or "(Unknown)"
        p = parties.setdefault(name, {"taxable_total": 0.0, "tax_total": 0.0, "invoice_count": 0})
        if taxable_col:
            p["taxable_total"] += _to_number(row.get(taxable_col))
        if gst_col:
            p["tax_total"] += _to_number(row.get(gst_col))
        p["invoice_count"] += 1

    for p in parties.values():
        p["taxable_total"] = round(p["taxable_total"], 2)
        p["tax_total"] = round(p["tax_total"], 2)

    return {"parties": parties, "mapping": mapping}
