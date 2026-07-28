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
    "journal_no": ["journalno", "journalnum", "journalnumber", "jvno", "jvnum",
                    "jvnumber", "voucherno", "vouchernum", "vouchernumber", "journalref"],
    "journal_name": ["journalname", "jvname", "vouchername", "journaltype", "journal"],
}

FIELD_PRIORITY = [
    "company_name", "invoice_no", "gstin", "date", "cgst", "sgst", "igst",
    "journal_no", "journal_name", "debit", "credit",
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
    """
    normalized = {col: _normalize(col) for col in columns}
    mapping = {}
    used_columns = set()

    EXACT_ONLY_LENGTH = 2

    for field in FIELD_PRIORITY:
        patterns = FIELD_PATTERNS[field]
        for col, norm in normalized.items():
            if col in used_columns:
                continue
            matched = False
            for pattern in patterns:
                if len(pattern) <= EXACT_ONLY_LENGTH:
                    if norm == pattern:
                        matched = True
                        break
                elif pattern in norm:
                    matched = True
                    break
            if matched:
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


# ---------------------------------------------------------------- #
# Content-based header guessing -- for when a column has no heading
# at all in the source file. Looks at the actual cell VALUES instead
# of the header text, so "Review & Clean Data" can suggest a sensible
# heading for the user to accept or edit.
# ---------------------------------------------------------------- #

_ACCOUNT_NAME_KEYWORDS = [
    "cash", "bank", "inventory", "stock", "receivable", "payable",
    "debtor", "creditor", "capital", "equity", "reserve", "provision",
    "loan", "revenue", "sales", "purchase", "expense", "salary", "wages",
    "rent", "depreciation", "gst", "tds", "asset", "liability",
    "insurance", "commission", "interest", "discount", "freight",
    "carriage", "furniture", "equipment", "vehicle", "investment",
    "drawings", "suspense", "goodwill", "prepaid", "outstanding",
    "office", "fixed asset", "current asset",
]

_GSTIN_RE = re.compile(r"^\d{2}[A-Z]{5}\d{4}[A-Z][A-Z0-9]Z[A-Z0-9]$", re.IGNORECASE)
_DATE_RE = re.compile(r"^\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}$")
_INVOICE_LIKE_RE = re.compile(r"^[A-Za-z]{0,6}[-/#]?\d{2,}[A-Za-z]{0,3}$")
_ALPHA_WORDS_RE = re.compile(r"^[A-Za-z][A-Za-z .&'()\-]*$")

# Whole-word Debit/Credit hints -- checked FIRST, before assuming a
# mostly-numeric column is just a generic "Amount".
_DEBIT_WORDS_RE = re.compile(r"\bdr\b|\bdebit\b", re.IGNORECASE)
_CREDIT_WORDS_RE = re.compile(r"\bcr\b|\bcredit\b", re.IGNORECASE)


def _clean_sample(values, limit=40):
    """Strips blanks and caps the sample size, keeping guessing fast
    even on very long columns."""
    out = []
    for v in values:
        if v is None:
            continue
        s = str(v).strip()
        if s == "":
            continue
        out.append(s)
        if len(out) >= limit:
            break
    return out


def _looks_numeric(s):
    t = s.replace(",", "").replace("\u20b9", "").replace("$", "").strip()
    if t.startswith("(") and t.endswith(")"):
        t = "-" + t[1:-1]
    if t.endswith("%"):
        t = t[:-1]
    try:
        float(t)
        return True
    except ValueError:
        return False


def guess_header_from_values(values):
    """
    Best-effort guess at what a headerless column represents, judged
    purely from its own data. Returns a guessed header string, or
    None if the sample is too small or too mixed to guess confidently.
    """
    sample = _clean_sample(values)
    if len(sample) < 3:
        return None
    total = len(sample)

    debit_hits = sum(1 for s in sample if _DEBIT_WORDS_RE.search(s))
    credit_hits = sum(1 for s in sample if _CREDIT_WORDS_RE.search(s))
    if debit_hits and debit_hits >= credit_hits:
        return "Debit"
    if credit_hits:
        return "Credit"

    if sum(1 for s in sample if _looks_numeric(s)) / total >= 0.8:
        return "Amount"

    if sum(1 for s in sample if _GSTIN_RE.match(s)) / total >= 0.6:
        return "GSTIN"

    date_hits = 0
    for s in sample:
        if _DATE_RE.match(s):
            date_hits += 1
    if date_hits / total >= 0.6:
        return "Date"

    lowered = [s.lower() for s in sample]
    if sum(1 for s in lowered if any(k in s for k in _ACCOUNT_NAME_KEYWORDS)) / total >= 0.4:
        return "Account Name"

    if sum(1 for s in sample if _INVOICE_LIKE_RE.match(s)) / total >= 0.7:
        return "Invoice No"

    if sum(1 for s in sample if _ALPHA_WORDS_RE.match(s)) / total >= 0.8:
        avg_len = sum(len(s) for s in sample) / total
        return "Party Name" if avg_len <= 30 else "Description"

    return None