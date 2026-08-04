"""
auto_organize.py

Handles messy exported reports (QuickBooks Trial Balance / Balance
Sheet style files) that have title rows, a company-name row, a date
row, and a split-looking header mixed in above the real data -- the
"Choose Import Format" feature (Keep As-Is vs Auto-Organize).

This module only READS raw rows (list of lists, no header assumption)
and returns a cleaned table. It never touches the normal load_excel()
import path used everywhere else in the app.
"""

import re
from .excel_engine import _clean_headers, NO_HEADER_TEXT

# Cells matching any of these (case-insensitive) count as "this looks
# like a real column heading" -- used to find the true header row.
HEADER_KEYWORDS = [
    "debit", "credit", "amount", "balance", "date", "account",
    "particulars", "description", "gstin", "narration", "voucher",
    "quantity", "qty", "rate", "total", "code",
]

# Headings whose column should always be KEPT even if every value in
# it happens to be blank or identical -- these are real financial
# columns, not repeated metadata.
NEVER_DROP_KEYWORDS = [
    "debit", "credit", "amount", "balance", "date", "qty", "quantity",
    "rate", "total",
]

# Plain label words that shouldn't be used as part of a report title.
TITLE_SKIP_WORDS = {"company name", "name", "report", "account name"}

_DATE_RE = re.compile(r"^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$")


def _is_blank(v):
    return str(v).strip() == ""


def _header_score(row):
    """How many cells in this row look like real column headings."""
    score = 0
    for cell in row:
        text = str(cell).strip().lower()
        if not text:
            continue
        if any(kw in text for kw in HEADER_KEYWORDS):
            score += 1
    return score


def detect_header_row(rows, max_scan=20):
    """
    Returns the index of the row that looks like the real header row,
    or None if nothing in the first `max_scan` rows looks like one.
    Requires at least 2 header-keyword matches so a simple
    "Company Name | Acme Corp" title row doesn't get mistaken for it.
    """
    best_idx, best_score = None, 0
    for idx, row in enumerate(rows[:max_scan]):
        score = _header_score(row)
        if score >= 2 and score > best_score:
            best_idx, best_score = idx, score
    return best_idx


def _looks_like_number(s):
    s = s.replace(",", "").replace("(", "-").replace(")", "").replace("$", "").strip()
    if s == "":
        return False
    try:
        float(s)
        return True
    except ValueError:
        return False


def _detect_single_value_header(rows, max_scan=20):
    """Fallback for reports like a Balance Sheet or Profit & Loss where
    the only column heading is a single period label sitting by itself
    above a column of numbers (e.g. "Mar 31, 25" or "Jan - Mar 25"),
    with no Debit/Credit/Amount-style heading anywhere for
    `detect_header_row` above to key off. Returns the header row
    index, or None if nothing like that is found."""
    for idx, row in enumerate(rows[:max_scan]):
        if _header_score(row) > 0:
            continue  # would already have matched the keyword path above
        non_blank = [(c, str(v).strip()) for c, v in enumerate(row) if str(v).strip() != ""]
        if len(non_blank) != 1:
            continue
        col, _text = non_blank[0]
        if col == 0:
            continue  # a single cell in column 0 is a title/company-name row, not a heading
        below = [r[col] for r in rows[idx + 1: idx + 30] if col < len(r)]
        numeric_below = [v for v in below if _looks_like_number(str(v))]
        if len(numeric_below) >= 3:
            return idx
    return None


def _merge_indent_columns(rows, header_idx, num_cols):
    """Balance Sheet / P&L style reports spread the account label
    across several columns depending on how deeply it's nested (e.g.
    "ASSETS" in column A, "Current Assets" in column B, "Cash -
    Checking" in column D, one level per row, mutually exclusive).
    Detected only via `_detect_single_value_header` above, this
    collapses those columns into one real "Account" column by taking
    whichever one is actually populated on each row -- exactly the
    label a person reading the original sheet would see for that
    line -- instead of leaving several near-empty flagged columns."""
    header_row = rows[header_idx] + [""] * (num_cols - len(rows[header_idx]))
    value_cols = [c for c in range(num_cols) if str(header_row[c]).strip() != ""]
    if not value_cols:
        return
    first_value_col = min(value_cols)
    label_cols = [c for c in range(first_value_col) if str(header_row[c]).strip() == ""]
    if len(label_cols) < 2:
        return  # only one label column already -- nothing to merge

    data_rows = rows[header_idx + 1:]
    # These should be mutually-exclusive indentation levels (at most
    # one populated per row) if this really is an indent hierarchy.
    # A rare row where a leaf account repeats its parent's exact label
    # (e.g. "Opening Balance Equity" as both the section and the
    # account) is harmless -- the rightmost cell is still the right
    # value to keep. Only bail out if violations are common enough
    # that this doesn't look like an indent hierarchy at all.
    populated_rows = 0
    violations = 0
    for r in data_rows:
        populated = [c for c in label_cols if c < len(r) and str(r[c]).strip() != ""]
        if populated:
            populated_rows += 1
        if len(populated) > 1:
            violations += 1
    if populated_rows and violations / populated_rows > 0.05:
        return

    target = label_cols[0]
    for r in data_rows:
        while len(r) <= max(label_cols + [target]):
            r.append("")
        value = ""
        for c in reversed(label_cols):
            if str(r[c]).strip() != "":
                value = r[c]
                break
        for c in label_cols:
            r[c] = ""
        r[target] = value


def needs_organizing(raw_result):
    """True if ANY sheet in this raw (header=None) file has its real
    header row somewhere below row 0 -- i.e. the file has title rows
    above the data and plain load_excel() would misread it."""
    if not raw_result or "error" in raw_result:
        return False
    for rows in raw_result.get("sheets", {}).values():
        if not rows:
            continue
        header_idx = detect_header_row(rows)
        if header_idx is not None and header_idx > 0:
            return True
    return False


def _metadata_columns(rows, header_idx, num_cols):
    """Column indexes whose data-area values are all identical (a
    repeated label like a company name), and whose header text isn't
    a real financial/date field we should always keep."""
    header_row = rows[header_idx] if header_idx < len(rows) else []
    data_rows = rows[header_idx + 1:]
    metadata = set()
    for c in range(num_cols):
        header_text = str(header_row[c]).strip().lower() if c < len(header_row) else ""
        if any(kw in header_text for kw in NEVER_DROP_KEYWORDS):
            continue
        values = [str(r[c]).strip() for r in data_rows if c < len(r) and str(r[c]).strip() != ""]
        if len(values) >= 2 and len(set(values)) == 1:
            metadata.add(c)
    return metadata


def _build_title(rows, header_idx, metadata_cols):
    parts = []
    for row in rows[:header_idx]:
        for c, cell in enumerate(row):
            if c in metadata_cols:
                continue
            text = str(cell).strip()
            if not text or text.lower() in TITLE_SKIP_WORDS:
                continue
            parts.append(text)
            break  # one phrase per title row is enough
    # de-duplicate while keeping order
    seen = set()
    ordered = []
    for p in parts:
        if p not in seen:
            seen.add(p)
            ordered.append(p)
    return " \u2013 ".join(ordered)


def organize_sheet(rows):
    """
    rows: raw grid (list of list of str) for one sheet, as returned by
    ExcelEngine.load_excel_raw().

    Returns {"title": str, "row_count": int, "columns": [...], "data": [...]}
    in the same shape as a normal load_excel() sheet entry, or None if
    the sheet is empty / has no detectable header.
    """
    if not rows:
        return None
    rows = [list(r) for r in rows]  # ensure every row is mutable
    header_idx = detect_header_row(rows)
    used_fallback_header = False
    if header_idx is None:
        header_idx = _detect_single_value_header(rows)
        if header_idx is not None:
            used_fallback_header = True
        else:
            header_idx = 0  # nothing messy found -- treat row 0 as header

    num_cols = max(len(r) for r in rows)
    header_row = rows[header_idx] + [""] * (num_cols - len(rows[header_idx]))

    if used_fallback_header:
        _merge_indent_columns(rows, header_idx, num_cols)


    # A column whose ONLY value in the whole sheet is the literal word
    # "TOTAL" is a grand-total marker carried over from the source
    # file (usually meant to line up with the real label column next
    # to it). Relocate it into that neighboring column so it doesn't
    # steal the "Account" name away from the real label column, and
    # doesn't turn into its own near-empty flagged column -- it just
    # becomes part of the real column's data, exactly where a person
    # reading the original file would expect the TOTAL label to sit.
    data_rows = rows[header_idx + 1:]
    for c in range(num_cols):
        vals = [(ri, str(r[c]).strip()) for ri, r in enumerate(data_rows) if c < len(r) and str(r[c]).strip() != ""]
        if len(vals) == 1 and vals[0][1].upper() == "TOTAL":
            ri, _ = vals[0]
            target_row = data_rows[ri]
            for nc in (c + 1, c - 1):
                if 0 <= nc < num_cols:
                    while len(target_row) <= max(nc, c):
                        target_row.append("")
                    if str(target_row[nc]).strip() == "":
                        target_row[nc] = "TOTAL"
                        target_row[c] = ""
                        break

    metadata_cols = _metadata_columns(rows, header_idx, num_cols)
    title = _build_title(rows, header_idx, metadata_cols)

    kept_cols = [c for c in range(num_cols) if c not in metadata_cols]

    # Name each kept column: use its own header text if present,
    # otherwise guess "Account" for a text-heavy label column. A column
    # with no header text AND no data in it at all is dropped silently
    # here -- there's nothing useful to show or ask the user to name,
    # so it shouldn't turn into a red-flagged "No heading in source
    # file" column on the Preview screen.
    final_columns = []
    real_kept_cols = []
    for c in kept_cols:
        text = str(header_row[c]).strip()
        if text:
            final_columns.append(text)
            real_kept_cols.append(c)
            continue
        data_vals = [str(r[c]).strip() for r in rows[header_idx + 1:] if c < len(r) and str(r[c]).strip() != ""]
        if not data_vals:
            continue  # fully empty, unnamed column -- drop silently
        looks_textual = any(not v.replace(",", "").replace(".", "").replace("-", "").isdigit() for v in data_vals)
        guessed_name = "Account" if looks_textual else NO_HEADER_TEXT
        # Don't create a confusing duplicate column (e.g. "Account (2)")
        # when a genuine column with this same name already exists in
        # this sheet -- drop this extra column silently instead.
        if guessed_name in final_columns:
            continue
        final_columns.append(guessed_name)
        real_kept_cols.append(c)
    kept_cols = real_kept_cols
    final_columns = _clean_headers(final_columns)

    value_col_positions = [
        i for i, name in enumerate(final_columns)
        if any(kw in name.lower() for kw in NEVER_DROP_KEYWORDS)
    ]

    data = []
    for r in rows[header_idx + 1:]:
        r = r + [""] * (num_cols - len(r))
        if all(_is_blank(r[c]) for c in kept_cols):
            continue  # fully empty row
        if value_col_positions:
            row_vals = [r[kept_cols[i]] for i in value_col_positions if i < len(kept_cols)]
            if all(_is_blank(v) for v in row_vals):
                continue  # section-label row with no actual amounts -- drop
        data.append({final_columns[i]: r[kept_cols[i]] for i in range(len(kept_cols))})

    return {"title": title, "row_count": len(data), "columns": final_columns, "data": data}


def organize_workbook_raw(raw_result):
    """Applies organize_sheet() to every sheet in a raw (header=None)
    file result, returning a structured result in the same shape
    load_excel() normally produces (file_name, file_path, sheets)."""
    sheets = {}
    for sheet_name, rows in raw_result.get("sheets", {}).items():
        organized = organize_sheet(rows)
        if organized is None:
            continue
        sheets[sheet_name] = {
            "row_count": organized["row_count"],
            "columns": organized["columns"],
            "data": organized["data"],
            "title": organized["title"],
        }
    return {
        "file_name": raw_result["file_name"],
        "file_path": raw_result["file_path"],
        "company_name": "",
        "sheets": sheets,
    }


def _preview_lines(rows, max_rows=8, max_col_width=18):
    lines = []
    for row in rows[:max_rows]:
        cells = []
        for v in row:
            text = str(v).strip()
            if not text:
                text = "(blank)"
            cells.append(text[:max_col_width])
        lines.append(" | ".join(cells))
    return "\n".join(lines)


def build_previews(raw_result, sheet_name=None):
    """Returns (raw_preview_text, organized_preview_text) for the first
    (or given) sheet of a raw file -- used to fill the Choose Import
    Format dialog."""
    sheets = raw_result.get("sheets", {})
    if not sheets:
        return "", ""
    name = sheet_name if sheet_name in sheets else next(iter(sheets))
    rows = sheets[name]

    raw_preview = _preview_lines(rows, max_rows=8)

    organized = organize_sheet(rows)
    if organized:
        header_line = " | ".join(organized["columns"])
        body_lines = []
        for row in organized["data"][:6]:
            body_lines.append(" | ".join(str(row.get(c, "")).strip() or "(blank)" for c in organized["columns"]))
        org_lines = []
        if organized["title"]:
            org_lines.append(f"Report: {organized['title']}")
            org_lines.append("")
        org_lines.append(header_line)
        org_lines.extend(body_lines)
        organized_preview = "\n".join(org_lines)
    else:
        organized_preview = raw_preview

    return raw_preview, organized_preview