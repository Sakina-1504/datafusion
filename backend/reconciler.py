"""
reconciler.py

Compares two datasets (e.g. your purchase register vs the GSTR-2A/2B
downloaded from the government portal) and reports what matches and
what doesn't. This is usually one of the most time-consuming manual
tasks for a CA during GST filing/audit.

The user chooses which field(s) count as a "match" per run
(invoice_no, gstin, taxable_amount, date - any combination), since
GSTIN isn't always available (unregistered/composition suppliers)
and a rigid single rule breaks on real data.
"""

from backend.column_mapper import detect_columns

AVAILABLE_MATCH_FIELDS = ["invoice_no", "gstin", "taxable_amount", "date"]


def _to_number(value, default=0.0):
    try:
        if value == "" or value is None:
            return default
        return float(value)
    except (ValueError, TypeError):
        return default


def _build_key(row, mapping, match_fields):
    """Builds a tuple key from the chosen match fields for one row."""
    key_parts = []
    for field in match_fields:
        col = mapping.get(field)
        if col is None:
            key_parts.append(None)
            continue
        value = row.get(col, "")
        if field == "taxable_amount":
            # round to nearest rupee so trivial rounding doesn't break a match
            key_parts.append(round(_to_number(value)))
        else:
            key_parts.append(str(value).strip().upper())
    return tuple(key_parts)


def reconcile(data_a, columns_a, data_b, columns_b, match_fields=None):
    """
    data_a / columns_a: rows + columns of file A (e.g. your register)
    data_b / columns_b: rows + columns of file B (e.g. GSTR-2A/2B)
    match_fields: list from AVAILABLE_MATCH_FIELDS chosen by the user.
                  Defaults to ["invoice_no"] if not given.

    Returns:
        {
            "matched": [ {a_row, b_row}, ... ],            # present in both, amounts agree
            "matched_but_differs": [ {a_row, b_row, diff}, ... ],  # same key, different amount
            "only_in_a": [ row_index, ... ],
            "only_in_b": [ row_index, ... ],
            "match_fields_used": [...],
            "skipped_fields": [...fields requested but not found in one/both files...],
        }
    """

    if not match_fields:
        match_fields = ["invoice_no"]

    mapping_a = detect_columns(columns_a)
    mapping_b = detect_columns(columns_b)

    usable_fields = []
    skipped_fields = []
    for field in match_fields:
        if field in mapping_a and field in mapping_b:
            usable_fields.append(field)
        else:
            skipped_fields.append(field)

    if not usable_fields:
        return {"error": "None of the chosen match fields were found in both files."}

    index_b = {}
    for j, row in enumerate(data_b):
        key = _build_key(row, mapping_b, usable_fields)
        index_b.setdefault(key, []).append(j)

    matched = []
    matched_but_differs = []
    only_in_a = []
    used_b_indices = set()

    amount_col_a = mapping_a.get("taxable_amount")
    amount_col_b = mapping_b.get("taxable_amount")

    for i, row_a in enumerate(data_a):
        key = _build_key(row_a, mapping_a, usable_fields)
        candidates = index_b.get(key, [])
        # take the first unused candidate
        b_idx = next((j for j in candidates if j not in used_b_indices), None)

        if b_idx is None:
            only_in_a.append(i)
            continue

        used_b_indices.add(b_idx)
        row_b = data_b[b_idx]

        if amount_col_a and amount_col_b and "taxable_amount" not in usable_fields:
            amt_a = _to_number(row_a.get(amount_col_a))
            amt_b = _to_number(row_b.get(amount_col_b))
            if abs(amt_a - amt_b) > 1.0:
                matched_but_differs.append({
                    "a_row": i, "b_row": b_idx,
                    "amount_a": amt_a, "amount_b": amt_b,
                    "difference": round(amt_a - amt_b, 2),
                })
                continue

        matched.append({"a_row": i, "b_row": b_idx})

    only_in_b = [j for j in range(len(data_b)) if j not in used_b_indices]

    return {
        "matched": matched,
        "matched_but_differs": matched_but_differs,
        "only_in_a": only_in_a,
        "only_in_b": only_in_b,
        "match_fields_used": usable_fields,
        "skipped_fields": skipped_fields,
        "summary": (
            f"{len(matched)} matched, {len(matched_but_differs)} matched-but-differ, "
            f"{len(only_in_a)} only in file A, {len(only_in_b)} only in file B."
        ),
    }
