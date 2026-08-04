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
                tagged_row["__report_title"] = sheet.get("title", "")
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


def credit_debit_checkpoint(rows, columns, file_paths=None):
    """
    Checks whether Total Debit equals Total Credit -- a common tie-out
    check before a report is considered final -- both overall AND
    broken down by the exact file/sheet each row came from, so a
    mismatch can be traced straight back to its source instead of just
    showing one grand "doesn't match" number.

    For each mismatched file/sheet, also tries to point at the single
    most likely culprit row -- either a row whose lone Debit/Credit
    value exactly (or closely) matches the shortfall, which usually
    means a duplicated/misplaced entry, or -- if nothing matches --
    a plain-English note that the fix is probably a missing entry
    rather than a bad one. This is what powers the "Open File" button
    jumping straight to a specific row instead of just the sheet.

    file_paths: optional {file_name: file_path} so the UI can offer an
    "Open File" button next to each mismatched file/sheet.

    Returns None if the data doesn't have both a debit and a credit
    column (nothing to check, so the checkpoint is simply skipped
    rather than showing a false mismatch). Otherwise returns:
        {
            "debit_col": ..., "credit_col": ...,
            "total_debit": ..., "total_credit": ..., "difference": ...,
            "matched": bool,
            "mismatches": [
                {"file_name":..., "sheet_name":..., "file_path":...,
                 "debit_col":..., "credit_col":...,
                 "total_debit":..., "total_credit":..., "difference":...,
                 "row_count":...,
                 "suspect_row": int or None,      # 1-based Excel row number
                 "suspect_column": str or None,    # debit_col or credit_col
                 "suspect_value": float or None,
                 "suspect_match": "exact"|"closest"|None,
                 "guidance": str},                 # plain-English next step
                ...
            ],
        }
    """
    mapping = detect_columns(columns)
    debit_col = mapping.get("debit")
    credit_col = mapping.get("credit")
    if not debit_col or not credit_col:
        return None

    file_paths = file_paths or {}
    groups = {}   # (file_name, sheet_name) -> running totals + per-row detail
    order = []
    total_debit = total_credit = 0.0

    for row in rows:
        fname = row.get("__source_file", "Unknown")
        sheet = row.get("__source_sheet", "Unknown")
        key = (fname, sheet)
        if key not in groups:
            groups[key] = {"total_debit": 0.0, "total_credit": 0.0, "row_count": 0, "detail": []}
            order.append(key)
        d = _to_number(row.get(debit_col))
        c = _to_number(row.get(credit_col))
        g = groups[key]
        g["row_count"] += 1
        # 1-based Excel row number, assuming row 1 is the header (same
        # assumption exporter.open_file_with_highlighted_columns makes).
        g["detail"].append({"excel_row": g["row_count"] + 1, "debit": d, "credit": c})
        g["total_debit"] += d
        g["total_credit"] += c
        total_debit += d
        total_credit += c

    def _find_suspect(detail, diff, debit_col, credit_col):
        """diff = total_debit - total_credit for this sheet. Looks for a
        single row whose value in the "heavier" column matches the
        shortfall -- the classic sign of a duplicated or misplaced
        entry. Falls back to the closest candidate if nothing matches
        exactly, and gives up (returns all-None) if nothing is even
        close, since guessing wrong is worse than not guessing."""
        target = abs(diff)
        if target < 0.01:
            return None, None, None, None
        # If Debit is the heavier side, a lone/duplicated Debit entry is
        # the likely culprit; otherwise look at Credit.
        heavy_col_name = debit_col if diff > 0 else credit_col
        heavy_key = "debit" if diff > 0 else "credit"

        best = None
        best_gap = None
        exact = None
        for d in detail:
            val = d[heavy_key]
            if val <= 0:
                continue
            gap = abs(val - target)
            if gap < 0.01:
                exact = d
                break
            if best_gap is None or gap < best_gap:
                best, best_gap = d, gap

        if exact:
            return exact["excel_row"], heavy_col_name, exact[heavy_key], "exact"
        # Only surface a "closest" guess if it's reasonably close --
        # otherwise it's just noise that sends the user to the wrong row.
        if best is not None and best_gap <= target * 0.5:
            return best["excel_row"], heavy_col_name, best[heavy_key], "closest"
        return None, None, None, None

    mismatches = []
    for fname, sheet in order:
        g = groups[(fname, sheet)]
        diff = round(g["total_debit"] - g["total_credit"], 2)
        if abs(diff) >= 0.01:
            s_row, s_col, s_val, s_match = _find_suspect(g["detail"], diff, debit_col, credit_col)

            heavier = debit_col if diff > 0 else credit_col
            lighter = credit_col if diff > 0 else debit_col
            if s_match == "exact":
                guidance = (
                    f"Row {s_row}'s '{heavier}' value of {s_val:,.2f} exactly matches the "
                    f"shortfall -- it's likely a duplicated or misplaced entry. Check whether "
                    f"it should be removed, corrected, or moved to '{lighter}'."
                )
            elif s_match == "closest":
                guidance = (
                    f"No exact match, but Row {s_row}'s '{heavier}' value of {s_val:,.2f} is the "
                    f"closest single entry to the shortfall of {abs(diff):,.2f} -- worth checking "
                    f"first, though it may not be the actual cause."
                )
            else:
                guidance = (
                    f"No single row accounts for the {abs(diff):,.2f} shortfall -- it's likely a "
                    f"transaction missing from '{lighter}' rather than a bad entry already in the "
                    f"sheet. Add the missing entry to '{lighter}' rather than editing an existing row."
                )

            mismatches.append({
                "file_name": fname,
                "sheet_name": sheet,
                "file_path": file_paths.get(fname),
                "debit_col": debit_col,
                "credit_col": credit_col,
                "total_debit": round(g["total_debit"], 2),
                "total_credit": round(g["total_credit"], 2),
                "difference": diff,
                "row_count": g["row_count"],
                "suspect_row": s_row,
                "suspect_column": s_col,
                "suspect_value": s_val,
                "suspect_match": s_match,
                "guidance": guidance,
            })

    difference = round(total_debit - total_credit, 2)
    # IMPORTANT: "matched" is NOT just "does the grand total tie out".
    # Two files/sheets with opposite mismatches (one debit-heavy, one
    # credit-heavy) can cancel out and make the grand total look fine
    # even though every individual file is genuinely wrong -- so this
    # only counts as matched if the grand total ties out AND every
    # individual file/sheet in `mismatches` is also clean.
    return {
        "debit_col": debit_col,
        "credit_col": credit_col,
        "total_debit": round(total_debit, 2),
        "total_credit": round(total_credit, 2),
        "difference": difference,
        "matched": abs(difference) < 0.01 and not mismatches,
        "mismatches": mismatches,
    }


def journal_wise_summary(data, columns):
    """
    Totals Debit and Credit per Journal Number + Journal Name -- the
    ledger-style breakdown most CA users actually want, instead of a
    flat row list. Grouped on the PAIR of (journal number, journal
    name) so two different journals that happen to share a name still
    show up as separate lines.

    Returns None if neither a Journal Number nor a Journal Name column
    could be detected (nothing to group by). Otherwise:
        {
            "journal_no_col": ..., "journal_name_col": ...,
            "debit_col": ..., "credit_col": ...,
            "journals": {
                (journal_no, journal_name): {
                    "debit_total": ..., "credit_total": ...,
                    "difference": ..., "row_count": ...
                }, ...
            }
        }
    """
    mapping = detect_columns(columns)
    journal_no_col = mapping.get("journal_no")
    journal_name_col = mapping.get("journal_name")
    debit_col = mapping.get("debit")
    credit_col = mapping.get("credit")

    if not journal_no_col and not journal_name_col:
        return None

    journals = {}
    order = []
    for row in data:
        jno = str(row.get(journal_no_col, "")).strip() if journal_no_col else ""
        jname = str(row.get(journal_name_col, "")).strip() if journal_name_col else ""
        key = (jno or "(No Journal No.)", jname or "(No Journal Name)")
        if key not in journals:
            journals[key] = {"debit_total": 0.0, "credit_total": 0.0, "row_count": 0}
            order.append(key)
        if debit_col:
            journals[key]["debit_total"] += _to_number(row.get(debit_col))
        if credit_col:
            journals[key]["credit_total"] += _to_number(row.get(credit_col))
        journals[key]["row_count"] += 1

    result_journals = {}
    for key in order:
        vals = journals[key]
        result_journals[key] = {
            "debit_total": round(vals["debit_total"], 2),
            "credit_total": round(vals["credit_total"], 2),
            "difference": round(vals["debit_total"] - vals["credit_total"], 2),
            "row_count": vals["row_count"],
        }

    return {
        "journal_no_col": journal_no_col, "journal_name_col": journal_name_col,
        "debit_col": debit_col, "credit_col": credit_col,
        "journals": result_journals,
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