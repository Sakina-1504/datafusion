"""
cleanup.py

Looks at one imported sheet (columns + rows, exactly as produced by
ExcelEngine.load_excel) and answers two practical questions a user
cares about before they consolidate:

  1. "Does this sheet have fully-empty rows or columns I should strip
      out first?"
  2. "Does this sheet have a name/party column with some blank
      entries I should fill in (or explicitly mark) before merging?"

Nothing here mutates data in place -- every function returns new
columns/rows so the caller (the Review screen) can preview a change
before committing to it.
"""

import re

from backend.column_mapper import detect_columns

NO_NAME_PLACEHOLDER = "(No Name Provided)"

# Internal/bookkeeping columns are never counted as "empty" data columns.
_INTERNAL_PREFIX = "__"


def _display_columns(columns):
    return [c for c in columns if not str(c).startswith(_INTERNAL_PREFIX)]


def _is_blank(value):
    return str(value).strip() == ""


def detect_name_column(columns):
    """Best-effort guess at which column holds a person/party's name for
    this sheet (Customer Name, Party Name, Vendor Name, Student's Name,
    plain 'Name', etc.), so we can flag blanks in it. Deliberately
    skips 'Company Name' -- that's the consolidation-level column, not
    a per-row name field."""
    display_cols = _display_columns(columns)

    mapping = detect_columns(display_cols)
    if "party_name" in mapping:
        return mapping["party_name"]

    for col in display_cols:
        low = re.sub(r"[^a-z0-9]", "", str(col).lower())
        if low == "companyname":
            continue
        if low.endswith("name") or low == "name":
            return col

    return None


def analyze_sheet(columns, rows):
    """
    Returns:
        {
            "empty_rows": [row indices that are fully blank across every
                           display column],
            "empty_cols": [column names that are fully blank across every
                           row],
            "name_column": detected name column, or None,
            "missing_name_rows": [row indices (excluding empty rows) where
                                  the name column is blank],
        }
    """
    display_cols = _display_columns(columns)

    empty_rows = []
    for i, row in enumerate(rows):
        if display_cols and all(_is_blank(row.get(c, "")) for c in display_cols):
            empty_rows.append(i)

    empty_cols = []
    for c in display_cols:
        if rows and all(_is_blank(row.get(c, "")) for row in rows):
            empty_cols.append(c)

    name_column = detect_name_column(columns)
    missing_name_rows = []
    if name_column:
        empty_set = set(empty_rows)
        for i, row in enumerate(rows):
            if i in empty_set:
                continue
            if _is_blank(row.get(name_column, "")):
                missing_name_rows.append(i)

    return {
        "empty_rows": empty_rows,
        "empty_cols": empty_cols,
        "name_column": name_column,
        "missing_name_rows": missing_name_rows,
    }


def has_issues(analysis):
    return bool(analysis["empty_rows"] or analysis["empty_cols"] or analysis["missing_name_rows"])


def clean_empty(columns, rows):
    """Strips fully-empty rows and fully-empty columns. Returns
    (new_columns, new_rows) -- internal/bookkeeping columns (prefixed
    with '__') always survive untouched."""
    analysis = analyze_sheet(columns, rows)
    empty_cols = set(analysis["empty_cols"])
    empty_rows = set(analysis["empty_rows"])

    new_columns = [c for c in columns if c not in empty_cols]
    new_rows = [
        {c: row.get(c, "") for c in new_columns}
        for i, row in enumerate(rows)
        if i not in empty_rows
    ]
    return new_columns, new_rows


def apply_name_fixes(columns, rows, name_column, name_by_row, use_placeholder=True):
    """
    name_by_row: {row_index: "typed name"} for rows the user actually
    filled in. Any other row that's still blank in name_column is left
    alone unless use_placeholder is True, in which case it gets
    NO_NAME_PLACEHOLDER so it's never silently blank in the merged data.
    """
    if not name_column:
        return rows

    new_rows = []
    for i, row in enumerate(rows):
        r = dict(row)
        if _is_blank(r.get(name_column, "")):
            if i in name_by_row and str(name_by_row[i]).strip():
                r[name_column] = str(name_by_row[i]).strip()
            elif use_placeholder:
                r[name_column] = NO_NAME_PLACEHOLDER
        new_rows.append(r)
    return new_rows