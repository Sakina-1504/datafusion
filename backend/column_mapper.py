"""
column_mapper.py

Different Excel files call the same thing by different names —
"Invoice No" vs "Bill No." vs "Inv#", "Amount" vs "Taxable Value", etc.

This module looks at the column headers of an uploaded sheet and
guesses which real-world field each column represents, so the rest
of the app (validator, consolidator, reconciler) doesn't need to
know the exact column names used in any particular file.
"""

import re

# Each field maps to a list of keyword patterns we'll look for inside
# a (normalized) column name. Order matters: more specific patterns
# should come first within a field's list where there's overlap.
FIELD_PATTERNS = {
    "company_name": ["companyname", "legalname", "businessname", "firmname",
                      "entityname", "organisationname", "organizationname", "companyid"],
    "invoice_no": ["invoiceno", "invoicenum", "billno", "billnum", "invno", "invnumber"],
    "gstin": ["gstin", "gstno", "gstnumber"],
    "party_name": ["customername", "partyname", "vendorname", "suppliername",
                    "clientname", "customer", "party", "vendor", "supplier"],
    "date": ["invoicedate", "billdate", "date"],
    "taxable_amount": ["taxablevalue", "taxableamount", "amount", "value", "netamount"],
    "cgst": ["cgst"],
    "sgst": ["sgst"],
    "igst": ["igst"],
    "gst_rate": ["gstrate", "taxrate", "rate"],
    "gst_amount": ["gst", "tax", "totaltax", "taxamount"],
    "debit": ["debitamount", "debitamt", "debit", "dramount", "dr"],
    "credit": ["creditamount", "creditamt", "credit", "cramount", "cr"],
    "entry_type": ["drcr", "type", "entrytype", "transactiontype"],
}

# Order in which we try to resolve fields. Fields listed first "claim"
# a column before more generic fields (like gst_amount) get a chance to.
FIELD_PRIORITY = [
    "company_name", "invoice_no", "gstin", "date", "cgst", "sgst", "igst", "debit", "credit",
    "entry_type", "gst_rate", "taxable_amount", "gst_amount", "party_name",
]


def _normalize(col_name):
    """Lowercase and strip everything except letters/numbers, so
    'Invoice No.', 'invoice_no', 'Invoice  No' all normalize the same."""
    return re.sub(r"[^a-z0-9]", "", str(col_name).lower())


def detect_columns(columns):
    """
    Given a list of raw column names from an uploaded sheet, returns a
    dict mapping field -> actual column name, for every field it could
    confidently identify.

    Example return value:
        {
            "invoice_no": "Invoice No",
            "party_name": "Customer Name",
            "taxable_amount": "Amount",
            "gst_amount": "GST",
            "date": "Date",
        }

    Fields that couldn't be found are simply absent from the dict —
    callers must check with `.get()` and handle missing fields
    gracefully (skip that check, tell the user, etc.) rather than
    assuming every field is always present.
    """

    normalized = {col: _normalize(col) for col in columns}
    mapping = {}
    used_columns = set()

    for field in FIELD_PRIORITY:
        patterns = FIELD_PATTERNS[field]
        for col, norm in normalized.items():
            if col in used_columns:
                continue
            if any(pattern in norm for pattern in patterns):
                mapping[field] = col
                used_columns.add(col)
                break

    return mapping


def describe_mapping(mapping, all_columns):
    """
    Returns a human-readable summary of what was detected and what
    was not, so the UI can show the user exactly what the app assumed
    before running checks on their data.
    """

    lines = []
    for field in FIELD_PRIORITY:
        if field in mapping:
            lines.append(f"  ✔ {field} -> '{mapping[field]}'")
        else:
            lines.append(f"  ✘ {field} -> not found")

    unused = [c for c in all_columns if c not in mapping.values()]
    if unused:
        lines.append(f"  (unused columns: {', '.join(unused)})")

    return "\n".join(lines)
