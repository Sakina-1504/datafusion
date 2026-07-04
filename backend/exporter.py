"""
exporter.py

Turns the app's in-memory data (consolidated rows, GST summary, party
summary, validation issues, filtered views) into a single, properly
formatted .xlsx workbook -- the kind a CA would actually want to open:
bold coloured headers, frozen header row, autofilter, sensible column
widths, number formatting on amount columns, and a bold Grand Total
row at the bottom of every data sheet.

Also provides open_file()/open_containing_folder() helpers so the UI
can take the user straight into Excel after export, cross-platform.
"""

import os
import platform
import subprocess
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from backend.filters import numeric_columns, grand_totals

HEADER_FILL = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
TOTAL_FILL = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
TOTAL_FONT = Font(bold=True, size=11)
TITLE_FONT = Font(bold=True, size=16, color="1E3A8A")
THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"),
)


def _autofit_columns(ws, columns, rows, min_width=10, max_width=45):
    for idx, col in enumerate(columns, start=1):
        longest = len(str(col))
        for row in rows[:500]:
            val_len = len(str(row.get(col, "")))
            if val_len > longest:
                longest = val_len
        ws.column_dimensions[get_column_letter(idx)].width = max(min_width, min(longest + 3, max_width))


def _write_table(ws, columns, rows, start_row=1, add_grand_total=True):
    for c_idx, col in enumerate(columns, start=1):
        cell = ws.cell(row=start_row, column=c_idx, value=str(col))
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER

    num_cols = numeric_columns(rows, columns, sample_size=len(rows) or 1)

    r = start_row
    for row in rows:
        r += 1
        for c_idx, col in enumerate(columns, start=1):
            value = row.get(col, "")
            cell = ws.cell(row=r, column=c_idx, value=value)
            cell.border = THIN_BORDER
            if col in num_cols and value not in ("", None):
                try:
                    cell.value = float(value)
                    cell.number_format = "#,##0.00"
                except (ValueError, TypeError):
                    pass

    if add_grand_total and rows:
        r += 1
        totals = grand_totals(rows, columns)
        label_written = False
        for c_idx, col in enumerate(columns, start=1):
            cell = ws.cell(row=r, column=c_idx)
            cell.fill = TOTAL_FILL
            cell.font = TOTAL_FONT
            cell.border = THIN_BORDER
            if col in totals:
                cell.value = totals[col]
                cell.number_format = "#,##0.00"
            elif not label_written:
                cell.value = "GRAND TOTAL"
                label_written = True

    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
    if rows:
        ws.auto_filter.ref = f"A{start_row}:{get_column_letter(len(columns))}{start_row + len(rows)}"
    _autofit_columns(ws, columns, rows)
    return r + 2


def export_full_report(output_path, consolidated, gst_summary=None, party_summary=None,
                        validation_issues=None, source_files=None, company_summary=None,
                        month_summary=None, consolidated_at=None):
    """
    Builds the complete multi-sheet workbook:
      1. Index               - what's in this workbook and why (so nobody
                                has to guess what each tab is for)
      2. Summary              - headline numbers + which files went in
      3. Consolidated Data    - every merged row + Grand Total row
      4. Company Summary      - totals per company (from the always-filled
                                 "Company Name" column)
      5. GST Summary          - rate-wise, CGST/SGST/IGST totals
      6. Monthly Summary      - totals per month, for filing/MIS trend
      7. Party Summary        - customer/vendor wise totals
      8. Issues Found         - every validation problem, with reference
    Any section is skipped gracefully if its data wasn't supplied.
    """

    wb = Workbook()
    rows = consolidated.get("rows", [])
    columns = [c for c in consolidated.get("columns", []) if not c.startswith("__")]
    # Guard against any stray blank header slipping through to export --
    # every column must show something a user can understand.
    columns = [c if str(c).strip() else "Reference Column (no header found)" for c in columns]
    # Company Name always leads the sheet since every row is tagged with it.
    if "Company Name" in columns:
        columns = ["Company Name"] + [c for c in columns if c != "Company Name"]
    has_source = bool(rows) and "__source_file" in rows[0]
    display_columns = columns + (["Source File"] if has_source else [])

    consolidated_at_text = (consolidated_at or datetime.now()).strftime("%d-%b-%Y %I:%M %p")

    # ---------- 1. Index / how-to-use sheet ----------
    ws_idx = wb.active
    ws_idx.title = "Index"
    ws_idx["A1"] = "DataFusion Platform -- Report Guide"
    ws_idx["A1"].font = TITLE_FONT
    ws_idx["A2"] = f"Consolidated on: {consolidated_at_text}   |   Exported on: {datetime.now().strftime('%d-%b-%Y %I:%M %p')}"
    ws_idx["A2"].font = Font(italic=True, size=10, color="666666")

    guide_rows = [
        ("Summary", "Headline totals, files used, CGST/SGST/IGST and data-quality snapshot."),
        ("Consolidated Data", "Every row from every imported file, merged, with a Company Name column and Grand Total row."),
        ("Company Summary", "Taxable value, tax and row count per company -- useful when you're handling more than one client at once."),
        ("GST Summary", "Rate-wise (5%/12%/18%/28%) taxable and tax totals, ready for GSTR-1 rate-wise filing."),
        ("Monthly Summary", "Taxable and tax totals grouped by month -- for filing trend or a management report."),
        ("Party Summary", "Taxable value, tax and invoice count per customer/vendor."),
        ("Issues Found", "Every data-quality issue detected (missing fields, invalid GSTIN, duplicates, tax mismatches) with its exact file/sheet/row reference."),
    ]
    r = 4
    ws_idx.cell(row=r, column=1, value="Sheet").font = HEADER_FONT
    ws_idx.cell(row=r, column=1).fill = HEADER_FILL
    ws_idx.cell(row=r, column=2, value="What it's for").font = HEADER_FONT
    ws_idx.cell(row=r, column=2).fill = HEADER_FILL
    for name, desc in guide_rows:
        r += 1
        ws_idx.cell(row=r, column=1, value=name).font = Font(bold=True)
        ws_idx.cell(row=r, column=2, value=desc)
    ws_idx.column_dimensions["A"].width = 22
    ws_idx.column_dimensions["B"].width = 95

    # ---------- 2. Summary sheet ----------
    ws = wb.create_sheet("Summary")
    ws["A1"] = "DataFusion Platform -- Consolidation Summary"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Consolidated on: {consolidated_at_text}"
    ws["A2"].font = Font(italic=True, size=10, color="666666")
    ws["A3"] = f"Report generated/exported on: {datetime.now().strftime('%d-%b-%Y %I:%M %p')}"
    ws["A3"].font = Font(italic=True, size=10, color="666666")

    r = 5
    ws.cell(row=r, column=1, value="Files included in this export:").font = Font(bold=True)
    r += 1
    for f in (source_files or []):
        ws.cell(row=r, column=1, value=f"  - {f}")
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="Grand Totals").font = Font(bold=True, size=13, color="1E3A8A")
    r += 1
    totals = grand_totals(rows, columns)
    ws.cell(row=r, column=1, value="Total Rows").font = Font(bold=True)
    ws.cell(row=r, column=2, value=len(rows))
    r += 1
    for col, total in totals.items():
        ws.cell(row=r, column=1, value=f"Total {col}").font = Font(bold=True)
        cell = ws.cell(row=r, column=2, value=total)
        cell.number_format = "#,##0.00"
        r += 1

    if gst_summary and "error" not in gst_summary:
        r += 1
        ws.cell(row=r, column=1, value="CGST Total").font = Font(bold=True)
        ws.cell(row=r, column=2, value=gst_summary.get("cgst_total", 0)).number_format = "#,##0.00"
        r += 1
        ws.cell(row=r, column=1, value="SGST Total").font = Font(bold=True)
        ws.cell(row=r, column=2, value=gst_summary.get("sgst_total", 0)).number_format = "#,##0.00"
        r += 1
        ws.cell(row=r, column=1, value="IGST Total").font = Font(bold=True)
        ws.cell(row=r, column=2, value=gst_summary.get("igst_total", 0)).number_format = "#,##0.00"
        r += 1
        mismatches = gst_summary.get("tax_mismatches", [])
        colr = "B91C1C" if mismatches else "16A34A"
        ws.cell(row=r, column=1, value="Tax Mismatches Found").font = Font(bold=True, color=colr)
        ws.cell(row=r, column=2, value=len(mismatches))

    if validation_issues is not None:
        r += 2
        if isinstance(validation_issues, dict):
            total_issues = sum(len(v) for v in validation_issues.values())
        else:
            total_issues = len(validation_issues)
        colr = "B91C1C" if total_issues else "16A34A"
        ws.cell(row=r, column=1, value="Data Quality Issues Found").font = Font(bold=True, color=colr)
        ws.cell(row=r, column=2, value=total_issues)

    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 20

    # ---------- 3. Consolidated Data sheet ----------
    ws2 = wb.create_sheet("Consolidated Data")
    export_rows = []
    for row in rows:
        flat = {c: row.get(c, "") for c in columns}
        if "Source File" in display_columns:
            flat["Source File"] = row.get("__source_file", "")
        export_rows.append(flat)
    _write_table(ws2, display_columns, export_rows)

    # ---------- 3. Company Summary sheet ----------
    if company_summary and "error" not in company_summary and company_summary.get("companies"):
        ws_co = wb.create_sheet("Company Summary")
        co_rows = []
        for name, vals in company_summary["companies"].items():
            co_rows.append({
                "Company Name": name,
                "Taxable Total": vals["taxable_total"],
                "Tax Total": vals["tax_total"],
                "Row Count": vals["row_count"],
            })
        _write_table(ws_co, ["Company Name", "Taxable Total", "Tax Total", "Row Count"], co_rows, add_grand_total=False)

    # ---------- 4. GST Summary sheet ----------
    if gst_summary and "error" not in gst_summary:
        ws3 = wb.create_sheet("GST Summary")
        gst_rows = []
        for rate, vals in sorted(gst_summary.get("rate_wise", {}).items()):
            gst_rows.append({
                "GST Rate (%)": rate,
                "Taxable Total": vals["taxable_total"],
                "Tax Total": vals["tax_total"],
                "Row Count": vals["row_count"],
            })
        _write_table(ws3, ["GST Rate (%)", "Taxable Total", "Tax Total", "Row Count"], gst_rows, add_grand_total=False)

    # ---------- 5. Monthly Summary sheet ----------
    if month_summary and "error" not in month_summary and month_summary.get("months"):
        ws_mo = wb.create_sheet("Monthly Summary")
        month_rows = []
        for month, vals in sorted(month_summary["months"].items()):
            month_rows.append({
                "Month": month,
                "Taxable Total": vals["taxable_total"],
                "Tax Total": vals["tax_total"],
                "Row Count": vals["row_count"],
            })
        _write_table(ws_mo, ["Month", "Taxable Total", "Tax Total", "Row Count"], month_rows, add_grand_total=False)

    # ---------- 6. Party Summary sheet ----------
    if party_summary and "error" not in party_summary:
        ws4 = wb.create_sheet("Party Summary")
        party_rows = []
        for name, vals in party_summary.get("parties", {}).items():
            party_rows.append({
                "Party Name": name,
                "Taxable Total": vals["taxable_total"],
                "Tax Total": vals["tax_total"],
                "Invoice Count": vals["invoice_count"],
            })
        _write_table(ws4, ["Party Name", "Taxable Total", "Tax Total", "Invoice Count"], party_rows, add_grand_total=False)

    # ---------- 7. Issues Found sheet ----------
    if validation_issues:
        ws5 = wb.create_sheet("Issues Found")
        issue_rows = []
        for issue_type, items in validation_issues.items():
            for item in items:
                issue_rows.append({
                    "Issue Type": issue_type.replace("_", " ").title(),
                    "Details": str(item),
                })
        if issue_rows:
            _write_table(ws5, ["Issue Type", "Details"], issue_rows, add_grand_total=False)
        else:
            ws5.cell(row=1, column=1, value="No issues found").font = Font(bold=True, color="16A34A")

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    return output_path


def export_filtered_view(output_path, columns, rows, title="Filtered Data"):
    """Simpler single-sheet export used by the Filters screen and the
    Analytics screen (exporting whatever the user is currently looking at)."""
    wb = Workbook()
    ws = wb.active
    ws.title = (title[:31] if title else "Filtered Data")
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Generated on: {datetime.now().strftime('%d-%b-%Y %I:%M %p')}  |  {len(rows)} row(s)"
    ws["A2"].font = Font(italic=True, size=10, color="666666")

    display_columns = [c if str(c).strip() else "Reference Column (no header found)" for c in columns if not str(c).startswith("__")]
    _write_table(ws, display_columns, rows, start_row=4)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    return output_path


def export_pivot_view(output_path, pivot_columns, pivot_rows, title="Pivot Table"):
    """Exports a pivot-shaped table (list of flat dict rows) with the
    same styling as everything else."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Pivot Table"
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Generated on: {datetime.now().strftime('%d-%b-%Y %I:%M %p')}"
    ws["A2"].font = Font(italic=True, size=10, color="666666")
    _write_table(ws, pivot_columns, pivot_rows, start_row=4, add_grand_total=False)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    return output_path


# ---------------- Cross-platform "open file / folder" helpers ---------------- #

def open_file(path):
    """Opens the given file with whatever application the OS has
    associated with .xlsx (i.e. Excel, if installed) -- this is what
    takes the user directly into Excel after export."""
    try:
        system = platform.system()
        if system == "Windows":
            os.startfile(path)  # noqa: only exists on Windows
        elif system == "Darwin":
            subprocess.run(["open", path], check=False)
        else:
            subprocess.run(["xdg-open", path], check=False)
        return True
    except Exception:
        return False


def open_containing_folder(path):
    """Opens the folder containing the file (selecting the file itself
    where the OS supports it)."""
    folder = os.path.dirname(os.path.abspath(path))
    try:
        system = platform.system()
        if system == "Windows":
            subprocess.run(["explorer", "/select,", os.path.abspath(path)], check=False)
        elif system == "Darwin":
            subprocess.run(["open", "-R", os.path.abspath(path)], check=False)
        else:
            subprocess.run(["xdg-open", folder], check=False)
        return True
    except Exception:
        return False
