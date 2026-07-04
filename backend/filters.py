"""
filters.py

Two ways of narrowing down consolidated data before analysing or
exporting it:

  1. quick_search()  -> one search box, matches ANY column, no setup.
  2. apply_filters() -> a list of precise conditions the user builds
                         (column + operator + value), all combined
                         with AND.

Both work directly on the list-of-dict rows produced by
consolidator.merge_files(), so they can be reused by the Filters
screen, Analytics screen, and Export screen alike.
"""

OPERATORS = [
    "equals", "not equals", "contains", "does not contain",
    "greater than", "less than", "between", "is empty", "is not empty",
]


def _to_number(value, default=None):
    try:
        if value == "" or value is None:
            return default
        return float(value)
    except (ValueError, TypeError):
        return default


def quick_search(rows, columns, term):
    """Case-insensitive substring match across every column."""
    if not term:
        return rows
    term = str(term).strip().lower()
    if not term:
        return rows

    matched = []
    for row in rows:
        for col in columns:
            if term in str(row.get(col, "")).lower():
                matched.append(row)
                break
    return matched


def _row_passes(row, condition):
    col = condition["column"]
    op = condition["operator"]
    val = condition.get("value", "")
    cell = row.get(col, "")
    cell_str = str(cell).strip().lower()
    val_str = str(val).strip().lower()

    if op == "equals":
        return cell_str == val_str
    if op == "not equals":
        return cell_str != val_str
    if op == "contains":
        return val_str in cell_str
    if op == "does not contain":
        return val_str not in cell_str
    if op == "is empty":
        return cell_str == ""
    if op == "is not empty":
        return cell_str != ""
    if op == "greater than":
        num_cell, num_val = _to_number(cell), _to_number(val)
        return num_cell is not None and num_val is not None and num_cell > num_val
    if op == "less than":
        num_cell, num_val = _to_number(cell), _to_number(val)
        return num_cell is not None and num_val is not None and num_cell < num_val
    if op == "between":
        # value expected as "low,high"
        try:
            low_str, high_str = str(val).split(",", 1)
            low, high = _to_number(low_str), _to_number(high_str)
            num_cell = _to_number(cell)
            return num_cell is not None and low is not None and high is not None and low <= num_cell <= high
        except ValueError:
            return False
    return True


def apply_filters(rows, conditions):
    """
    conditions: list of {"column": str, "operator": str, "value": str}
    All conditions are combined with AND. Conditions with a missing
    column/operator are ignored rather than raising.
    """
    if not conditions:
        return rows

    valid_conditions = [c for c in conditions if c.get("column") and c.get("operator")]
    if not valid_conditions:
        return rows

    result = []
    for row in rows:
        if all(_row_passes(row, c) for c in valid_conditions):
            result.append(row)
    return result


def numeric_columns(rows, columns, sample_size=50):
    """Guesses which columns are numeric by sampling the first few rows.
    Used to populate 'Y-axis / Values' dropdowns with sensible choices."""
    numeric_cols = []
    sample = rows[:sample_size] if rows else []
    for col in columns:
        if col.startswith("__"):
            continue
        seen_value = False
        all_numeric = True
        for row in sample:
            v = row.get(col, "")
            if v == "" or v is None:
                continue
            seen_value = True
            if _to_number(v) is None:
                all_numeric = False
                break
        if seen_value and all_numeric:
            numeric_cols.append(col)
    return numeric_cols


def grand_totals(rows, columns):
    """Sums every numeric column across all rows — the 'Grand Total'
    row that goes at the bottom of exported sheets."""
    totals = {}
    for col in numeric_columns(rows, columns, sample_size=len(rows) or 1):
        total = 0.0
        for row in rows:
            v = _to_number(row.get(col))
            if v is not None:
                total += v
        totals[col] = round(total, 2)
    return totals
