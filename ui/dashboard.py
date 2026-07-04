import os
import platform
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog
import customtkinter as ctk
import pandas as pd

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from ui.file_dialog import select_excel_files
from ui.widgets import StatCard, ZoomableTable, ZoomableTextbox
from ui import dialogs

from backend.excel_engine import ExcelEngine
from backend.database import DataStore
from backend.validator import validate_sheet
from backend.consolidator import (
    merge_files, gst_rate_summary, party_wise_summary,
    company_wise_summary, month_wise_summary,
)
from backend.filters import quick_search, apply_filters, numeric_columns, grand_totals, OPERATORS, _to_number
from backend import exporter
from backend import settings as settings_backend

ACTIVE_COLOR = "#F59E0B"
DEFAULT_BTN_COLOR = "#2563EB"


class Dashboard:

    def __init__(self, root):
        self.root = root

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.root.title("DataFusion Platform")
        self.root.geometry("1440x840")
        self.root.minsize(1150, 700)

        # ---------- core state ----------
        self.settings = settings_backend.load_settings()
        self.engine = ExcelEngine()
        self.store = DataStore()

        self.loaded_results = []          # one entry per imported file
        self.file_check_vars = {}         # file_name -> tk.BooleanVar
        self.company_names = {}           # file_name -> "Company Name" entered by the user

        self.consolidated = None          # {"rows": [...], "columns": [...]}
        self.consolidated_at = None        # datetime of the last consolidation run
        self.gst_summary = None
        self.party_summary = None
        self.company_summary = None
        self.month_summary = None
        self.source_files_used = []
        self.all_issues = {}              # aggregated validation issues, labeled

        self.filtered_rows = None         # None => no filter active, show all consolidated rows
        self.filter_condition_rows = []   # widgets for the advanced filter builder
        self.last_export_path = None

        self.pivot_result_columns = []
        self.pivot_result_rows = []

        self.menu_buttons = {}
        self.current_screen = "home"

        self.create_header()
        self.create_body()

    # ================= HEADER ================= #

    def create_header(self):
        header = ctk.CTkFrame(self.root, height=70, corner_radius=0, fg_color="#1E3A8A")
        header.pack(fill="x")

        ctk.CTkLabel(header, text="DataFusion Platform", font=("Segoe UI", 26, "bold"),
                     text_color="white").pack(side="left", padx=25, pady=15)

        ctk.CTkLabel(header, text="Enterprise Data Consolidation & Analysis for Finance/CA teams",
                     font=("Segoe UI", 13), text_color="#CBD5E1").pack(side="right", padx=25)

    # ================= BODY / SIDEBAR ================= #

    def create_body(self):
        body = ctk.CTkFrame(self.root, fg_color="#F5F7FA", corner_radius=0)
        body.pack(fill="both", expand=True)

        menu = ctk.CTkFrame(body, width=225, fg_color="#23395B", corner_radius=0)
        menu.pack(side="left", fill="y")
        menu.pack_propagate(False)

        ctk.CTkLabel(menu, text="MENU", font=("Segoe UI", 20, "bold"), text_color="white").pack(pady=(28, 16))

        nav_items = [
            ("home", "\U0001F3E0 Dashboard", self.show_home),
            ("import", "\U0001F4C2 Import Files", self.import_files),
            ("validate", "\u2714 Validate Data", self.validate_data),
            ("consolidate", "\U0001F4CA Consolidate", self.consolidate_data),
            ("filters", "\U0001F50E Filters & Search", self.show_filters),
            ("analytics", "\U0001F4C8 Analytics", self.show_analytics),
            ("pivot", "\U0001F9EE Pivot Table", self.show_pivot),
            ("export", "\U0001F4E4 Export", self.export_data),
            ("settings", "\u2699 Settings", self.show_settings),
        ]

        for key, label, cmd in nav_items:
            btn = ctk.CTkButton(menu, text=label, width=185, height=40, fg_color=DEFAULT_BTN_COLOR,
                                 command=cmd)
            btn.pack(pady=6)
            self.menu_buttons[key] = btn

        self.workspace = ctk.CTkScrollableFrame(body, fg_color="#F5F7FA")
        self.workspace.pack(side="left", fill="both", expand=True, padx=20, pady=20)

        self.show_home()

    def _set_active(self, key):
        self.current_screen = key
        for k, btn in self.menu_buttons.items():
            btn.configure(fg_color=ACTIVE_COLOR if k == key else DEFAULT_BTN_COLOR)

    def _clear_workspace(self):
        for widget in self.workspace.winfo_children():
            widget.destroy()

    def _screen_heading(self, title, subtitle=None):
        ctk.CTkLabel(self.workspace, text=title, font=("Segoe UI", 24, "bold"),
                     text_color="#1E3A8A").pack(anchor="w", pady=(0, 2))
        if subtitle:
            ctk.CTkLabel(self.workspace, text=subtitle, font=("Segoe UI", 13),
                         text_color="#64748B").pack(anchor="w", pady=(0, 14))
        else:
            ctk.CTkFrame(self.workspace, height=8, fg_color="transparent").pack()

    # ================= HOME / DASHBOARD ================= #

    def show_home(self):
        self._set_active("home")
        self._clear_workspace()
        self._screen_heading("Welcome to DataFusion Platform",
                              "Enterprise offline data consolidation, validation, filtering, analytics and reporting.")

        # ---------- stat cards ----------
        cards_row = ctk.CTkFrame(self.workspace, fg_color="transparent")
        cards_row.pack(fill="x", pady=(0, 20))

        total_issue_count = sum(len(v) for v in self.all_issues.values())
        consolidated_rows = len(self.consolidated["rows"]) if self.consolidated else 0
        filtered_rows_count = len(self.filtered_rows) if self.filtered_rows is not None else consolidated_rows

        cards = [
            ("Files Imported", len(self.loaded_results), "#1E3A8A"),
            ("Consolidated Rows", consolidated_rows, "#0F766E"),
            ("Issues Found", total_issue_count, "#DC2626" if total_issue_count else "#16A34A"),
            ("Rows After Filter", filtered_rows_count, "#7C3AED"),
        ]
        for title, value, color in cards:
            card = StatCard(cards_row, title, value, color=color, width=250, height=90)
            card.pack(side="left", padx=(0, 14))

        if not self.loaded_results:
            ctk.CTkLabel(
                self.workspace,
                text="No files imported yet. Click '\U0001F4C2 Import Files' in the menu to select one or more Excel files\n"
                     "(hold Ctrl or Shift while browsing to pick several at once).",
                justify="left", anchor="w", font=("Segoe UI", 14)
            ).pack(fill="x", pady=20)
            return

        # ---------- imported files tick list ---------- #
        header_row = ctk.CTkFrame(self.workspace, fg_color="transparent")
        header_row.pack(fill="x", pady=(6, 6))

        ctk.CTkLabel(header_row, text=f"Imported Files ({len(self.loaded_results)}) \u2014 tick which ones to use",
                     font=("Segoe UI", 16, "bold"), text_color="#1E3A8A").pack(side="left")

        ctk.CTkButton(header_row, text="Select All", width=100, command=self._select_all_files).pack(side="right", padx=(6, 0))
        ctk.CTkButton(header_row, text="Deselect All", width=110, fg_color="#B91C1C", hover_color="#7f1d1d",
                      command=self._deselect_all_files).pack(side="right", padx=6)

        file_list_frame = ctk.CTkScrollableFrame(self.workspace, fg_color="white", height=230)
        file_list_frame.pack(fill="x", pady=(0, 10))

        for result in self.loaded_results:
            fname = result["file_name"]
            if fname not in self.file_check_vars:
                self.file_check_vars[fname] = tk.BooleanVar(value=True)

            total_rows = sum(sheet["row_count"] for sheet in result["sheets"].values())
            sheet_count = len(result["sheets"])

            row = ctk.CTkFrame(file_list_frame, fg_color="transparent")
            row.pack(fill="x", padx=5, pady=4)

            company = self.company_names.get(fname) or os.path.splitext(fname)[0]
            ctk.CTkCheckBox(row, text=f"{fname}   ({sheet_count} sheet(s), {total_rows} row(s))",
                             variable=self.file_check_vars[fname], font=("Segoe UI", 13)).pack(side="left", padx=5)

            ctk.CTkLabel(row, text=f"\U0001F3E2 {company}", font=("Segoe UI", 12, "italic"),
                         text_color="#1E3A8A").pack(side="left", padx=(14, 0))

            ctk.CTkButton(row, text="\u2715 Remove", width=80, height=26, fg_color="#94A3B8",
                          hover_color="#64748B",
                          command=lambda f=fname: self._remove_file(f)).pack(side="right", padx=5)
            ctk.CTkButton(row, text="\u270E Edit Company", width=120, height=26, fg_color="#2563EB",
                          hover_color="#1d4ed8",
                          command=lambda f=fname: self._edit_company_name(f)).pack(side="right", padx=5)

        selected_count = sum(1 for v in self.file_check_vars.values() if v.get())
        ctk.CTkLabel(
            self.workspace,
            text=f"{selected_count} of {len(self.loaded_results)} file(s) ticked \u2014 "
                 f"Validate / Consolidate will only use ticked files.",
            font=("Segoe UI", 12, "italic"), text_color="#16A34A"
        ).pack(anchor="w", pady=(0, 10))

    def _select_all_files(self):
        for var in self.file_check_vars.values():
            var.set(True)
        self.show_home()

    def _deselect_all_files(self):
        for var in self.file_check_vars.values():
            var.set(False)
        self.show_home()

    def _remove_file(self, fname):
        self.loaded_results = [r for r in self.loaded_results if r["file_name"] != fname]
        self.file_check_vars.pop(fname, None)
        self.company_names.pop(fname, None)
        self.show_home()

    def _edit_company_name(self, fname):
        current = self.company_names.get(fname) or os.path.splitext(fname)[0]
        new_name = dialogs.ask_text(self.root, "Edit Company Name",
                                     f"Company name for '{fname}':", default=current)
        if new_name:
            self.company_names[fname] = new_name.strip() or current
            self.show_home()

    def _get_selected_results(self):
        return [r for r in self.loaded_results
                if self.file_check_vars.get(r["file_name"], tk.BooleanVar(value=True)).get()]

    # ================= IMPORT FILES ================= #

    def import_files(self):
        self._set_active("import")
        new_paths = select_excel_files()

        if not new_paths:
            self.show_home()
            return

        errors = []
        added = 0
        newly_added_files = []
        existing_names = {r["file_name"] for r in self.loaded_results}

        for path in new_paths:
            result = self.engine.load_excel(path)
            if "error" in result:
                errors.append(f"{os.path.basename(path)}: {result['error']}")
                continue
            # re-importing a file with the same name replaces the old copy
            if result["file_name"] in existing_names:
                self.loaded_results = [r for r in self.loaded_results if r["file_name"] != result["file_name"]]
            self.loaded_results.append(result)
            self.file_check_vars[result["file_name"]] = tk.BooleanVar(value=True)
            self.store.add_upload(result["file_name"], result)
            existing_names.add(result["file_name"])
            newly_added_files.append(result["file_name"])
            added += 1

        # ONE popup listing every file just imported -- not one popup per
        # file. Each row is pre-filled with a guessed company name; edit
        # any of them here, or click Skip to keep the guessed names (you
        # can still rename any of them later with "Edit Company").
        if newly_added_files:
            names = dialogs.ask_company_names(self.root, newly_added_files)
            if names:
                self.company_names.update(names)
            else:
                for f in newly_added_files:
                    self.company_names.setdefault(f, os.path.splitext(f)[0])

        self.show_home()

        if errors:
            dialogs.show_error(self.root, "Import Errors", "\n".join(errors))
        if added:
            dialogs.show_info(
                self.root, "Import Complete",
                f"Loaded {added} file(s) successfully.\n"
                f"Total files now available: {len(self.loaded_results)}"
            )

    # ================= VALIDATE ================= #

    def validate_data(self):
        self._set_active("validate")
        if not self.loaded_results:
            dialogs.show_warning(self.root, "No Data", "Please import files first.")
            self.show_home()
            return

        selected = self._get_selected_results()
        if not selected:
            dialogs.show_warning(self.root, "No Files Ticked", "Tick at least one imported file on the Dashboard screen first.")
            self.show_home()
            return

        self._clear_workspace()
        self._screen_heading("Validation Results", "Every issue below tells you exactly which file, sheet and row it is in.")

        self.all_issues = {}
        output_lines = [f"Validating {len(selected)} of {len(self.loaded_results)} imported file(s) (ticked only)", ""]

        for result in selected:
            fname = result["file_name"]
            output_lines.append(f"FILE: {fname}")
            for sheet_name, sheet in result["sheets"].items():
                report = validate_sheet(sheet["data"], sheet["columns"])
                output_lines.append(f"  Sheet: {sheet_name}")
                output_lines.append(f"  {report['summary']}")
                for check in report["skipped_checks"]:
                    output_lines.append(f"    (skipped) {check}")
                for issue_type, items in report["issues"].items():
                    if items:
                        output_lines.append(f"    {issue_type}: {len(items)} found")
                        for item in items[:10]:
                            output_lines.append(f"       - Row {item if isinstance(item, int) else item}")
                        if len(items) > 10:
                            output_lines.append(f"       ... and {len(items) - 10} more")
                        labeled = self.all_issues.setdefault(issue_type, [])
                        for item in items:
                            labeled.append(f"File: {fname} | Sheet: {sheet_name} | {item}")
                output_lines.append("")
            output_lines.append("")

        total_issues = sum(len(v) for v in self.all_issues.values())
        if total_issues:
            banner = ctk.CTkLabel(self.workspace, text=f"\u26A0 {total_issues} issue(s) found across the ticked files \u2014 see details below and in the exported 'Issues Found' sheet.",
                                   font=("Segoe UI", 13, "bold"), text_color="#B91C1C")
        else:
            banner = ctk.CTkLabel(self.workspace, text="\u2714 No issues found in the ticked files.",
                                   font=("Segoe UI", 13, "bold"), text_color="#16A34A")
        banner.pack(anchor="w", pady=(0, 8))

        box = ZoomableTextbox(self.workspace, height=430)
        box.pack(fill="both", expand=True)
        box.set_text("\n".join(output_lines))

    # ================= CONSOLIDATE ================= #

    def consolidate_data(self):
        self._set_active("consolidate")
        if not self.loaded_results:
            dialogs.show_warning(self.root, "No Data", "Please import files first.")
            self.show_home()
            return

        selected = self._get_selected_results()
        if not selected:
            dialogs.show_warning(self.root, "No Files Ticked", "Tick at least one imported file on the Dashboard screen first.")
            self.show_home()
            return

        merged = merge_files(selected, company_names=self.company_names)
        self.store.set_consolidated(merged)
        self.consolidated = merged
        self.consolidated_at = datetime.now()
        self.source_files_used = [r["file_name"] for r in selected]
        self.filtered_rows = None  # reset any previous filter when re-consolidating

        self.gst_summary = gst_rate_summary(merged["rows"], merged["columns"])
        self.party_summary = party_wise_summary(merged["rows"], merged["columns"])
        self.company_summary = company_wise_summary(merged["rows"], merged["columns"])
        self.month_summary = month_wise_summary(merged["rows"], merged["columns"])

        self._clear_workspace()
        self._screen_heading("Consolidation Results",
                              f"Merged from {len(selected)} of {len(self.loaded_results)} ticked file(s) "
                              f"on {self.consolidated_at.strftime('%d-%b-%Y %I:%M %p')}.")

        # ---- headline numbers ---- #
        totals = grand_totals(merged["rows"], merged["columns"])
        cards_row = ctk.CTkFrame(self.workspace, fg_color="transparent")
        cards_row.pack(fill="x", pady=(0, 14))
        StatCard(cards_row, "Total Rows", len(merged["rows"]), color="#0F766E", width=220, height=85).pack(side="left", padx=(0, 12))
        for col, total in list(totals.items())[:3]:
            StatCard(cards_row, f"Total {col}", f"{total:,.2f}", color="#1E3A8A", width=220, height=85).pack(side="left", padx=(0, 12))

        if self.gst_summary and "error" not in self.gst_summary and self.gst_summary.get("tax_mismatches"):
            ctk.CTkLabel(self.workspace, text=f"\u26A0 {len(self.gst_summary['tax_mismatches'])} tax mismatch(es) detected \u2014 check the Issues report.",
                         font=("Segoe UI", 13, "bold"), text_color="#B91C1C").pack(anchor="w", pady=(0, 8))

        ctk.CTkLabel(self.workspace, text="Merged Data (Grand Total row at the bottom)",
                     font=("Segoe UI", 15, "bold"), text_color="#1E3A8A").pack(anchor="w", pady=(6, 4))

        table = ZoomableTable(self.workspace, height=16)
        table.pack(fill="both", expand=True, pady=(0, 12))
        self._load_table_with_total(table, merged["columns"], merged["rows"])

        # ---- GST / Party summaries as text ---- #
        summary_lines = ["GST RATE-WISE SUMMARY"]
        if self.gst_summary and "error" not in self.gst_summary:
            for rate, vals in sorted(self.gst_summary["rate_wise"].items()):
                summary_lines.append(f"  {rate}% -> Taxable: {vals['taxable_total']:,.2f}, Tax: {vals['tax_total']:,.2f}, Rows: {vals['row_count']}")
            summary_lines.append(f"  CGST total: {self.gst_summary['cgst_total']:,.2f}  |  SGST total: {self.gst_summary['sgst_total']:,.2f}  |  IGST total: {self.gst_summary['igst_total']:,.2f}")
        else:
            summary_lines.append(f"  {self.gst_summary.get('error', 'n/a')}")

        summary_lines.append("")
        summary_lines.append("COMPANY-WISE SUMMARY")
        if self.company_summary and self.company_summary.get("companies"):
            for name, vals in self.company_summary["companies"].items():
                summary_lines.append(f"  {name}: Taxable {vals['taxable_total']:,.2f}, Tax {vals['tax_total']:,.2f}, Rows {vals['row_count']}")
        else:
            summary_lines.append("  (no company data)")

        summary_lines.append("")
        summary_lines.append("PARTY-WISE SUMMARY")
        if self.party_summary and "error" not in self.party_summary:
            for name, vals in self.party_summary["parties"].items():
                summary_lines.append(f"  {name}: Taxable {vals['taxable_total']:,.2f}, Tax {vals['tax_total']:,.2f}, Invoices {vals['invoice_count']}")
        else:
            summary_lines.append(f"  {self.party_summary.get('error', 'n/a')}")

        box = ZoomableTextbox(self.workspace, height=180)
        box.pack(fill="both", pady=(0, 10))
        box.set_text("\n".join(summary_lines))

        ctk.CTkButton(self.workspace, text="\U0001F50E Go to Filters & Search \u2192", height=38,
                      command=self.show_filters).pack(anchor="w")

    def _load_table_with_total(self, table_widget, columns, rows):
        """Populates a ZoomableTable and appends a synthetic GRAND TOTAL row for display."""
        display_cols = [c for c in columns if not str(c).startswith("__")]
        totals = grand_totals(rows, display_cols)
        total_row = {c: "" for c in display_cols}
        label_set = False
        for c in display_cols:
            if c in totals:
                total_row[c] = totals[c]
            elif not label_set:
                total_row[c] = "GRAND TOTAL"
                label_set = True
        display_rows = [{c: r.get(c, "") for c in display_cols} for r in rows] + ([total_row] if rows else [])
        table_widget.set_data(display_cols, display_rows)

    # ================= FILTERS & SEARCH ================= #

    def show_filters(self):
        self._set_active("filters")
        if not self.consolidated:
            dialogs.show_warning(self.root, "No Data", "Please consolidate data first (Import Files \u2192 Consolidate).")
            self.show_home()
            return

        self._clear_workspace()
        self._screen_heading("Filters & Search", "Combine a quick keyword search with precise column conditions. Results feed straight into Analytics.")

        columns = [c for c in self.consolidated["columns"] if not str(c).startswith("__")]

        # ---- quick search ---- #
        search_row = ctk.CTkFrame(self.workspace, fg_color="white", corner_radius=8)
        search_row.pack(fill="x", pady=(0, 10), ipady=8)
        ctk.CTkLabel(search_row, text="Quick Search:", font=("Segoe UI", 13, "bold")).pack(side="left", padx=(12, 8))
        self.search_entry = ctk.CTkEntry(search_row, placeholder_text="Type anything... searches every column", width=400)
        self.search_entry.pack(side="left", padx=(0, 8))
        ctk.CTkButton(search_row, text="Search", width=90, command=self._apply_filters_action).pack(side="left", padx=4)
        ctk.CTkButton(search_row, text="Clear All", width=90, fg_color="#B91C1C", hover_color="#7f1d1d",
                      command=self._clear_filters_action).pack(side="left", padx=4)

        # ---- advanced conditions ---- #
        adv_header = ctk.CTkFrame(self.workspace, fg_color="transparent")
        adv_header.pack(fill="x", pady=(6, 4))
        ctk.CTkLabel(adv_header, text="Advanced Conditions (all combined with AND)",
                     font=("Segoe UI", 14, "bold"), text_color="#1E3A8A").pack(side="left")
        ctk.CTkButton(adv_header, text="+ Add Condition", width=140, command=lambda: self._add_condition_row(columns)).pack(side="right")

        self.conditions_frame = ctk.CTkFrame(self.workspace, fg_color="white", corner_radius=8)
        self.conditions_frame.pack(fill="x", pady=(0, 10))
        self.filter_condition_rows = []
        self._add_condition_row(columns)

        ctk.CTkButton(self.workspace, text="\u2705 Apply Filters", height=38, fg_color="#16A34A", hover_color="#15803d",
                      command=self._apply_filters_action).pack(anchor="w", pady=(0, 10))

        # ---- results table ---- #
        self.filters_result_label = ctk.CTkLabel(self.workspace, text="", font=("Segoe UI", 13, "italic"), text_color="#64748B")
        self.filters_result_label.pack(anchor="w")

        self.filters_table = ZoomableTable(self.workspace, height=16)
        self.filters_table.pack(fill="both", expand=True, pady=(4, 10))

        ctk.CTkButton(self.workspace, text="\U0001F4E4 Export This Filtered View", height=38,
                      command=self._export_filtered_view).pack(anchor="w")

        self._refresh_filters_table(columns)

    def _add_condition_row(self, columns):
        row_frame = ctk.CTkFrame(self.conditions_frame, fg_color="transparent")
        row_frame.pack(fill="x", padx=10, pady=6)

        col_var = tk.StringVar(value=columns[0] if columns else "")
        op_var = tk.StringVar(value=OPERATORS[0])
        val_entry = ctk.CTkEntry(row_frame, placeholder_text="value (use low,high for 'between')", width=260)

        ctk.CTkOptionMenu(row_frame, values=columns, variable=col_var, width=180).pack(side="left", padx=4)
        ctk.CTkOptionMenu(row_frame, values=OPERATORS, variable=op_var, width=150).pack(side="left", padx=4)
        val_entry.pack(side="left", padx=4)

        entry = {"column_var": col_var, "operator_var": op_var, "value_entry": val_entry, "frame": row_frame}

        def remove_this():
            row_frame.destroy()
            self.filter_condition_rows.remove(entry)

        ctk.CTkButton(row_frame, text="\u2715", width=32, fg_color="#94A3B8", hover_color="#64748B",
                      command=remove_this).pack(side="left", padx=4)

        self.filter_condition_rows.append(entry)

    def _apply_filters_action(self):
        columns = [c for c in self.consolidated["columns"] if not str(c).startswith("__")]
        rows = self.consolidated["rows"]

        term = self.search_entry.get().strip() if hasattr(self, "search_entry") else ""
        if term:
            rows = quick_search(rows, columns, term)

        conditions = []
        for entry in self.filter_condition_rows:
            conditions.append({
                "column": entry["column_var"].get(),
                "operator": entry["operator_var"].get(),
                "value": entry["value_entry"].get(),
            })
        # ignore rows left at their default blank value so an untouched
        # condition row doesn't silently filter everything out
        conditions = [c for c in conditions if c["value"] or c["operator"] in ("is empty", "is not empty")]

        rows = apply_filters(rows, conditions)
        self.filtered_rows = rows
        self._refresh_filters_table(columns)

    def _clear_filters_action(self):
        if hasattr(self, "search_entry"):
            self.search_entry.delete(0, "end")
        for entry in list(self.filter_condition_rows):
            entry["frame"].destroy()
        self.filter_condition_rows = []
        columns = [c for c in self.consolidated["columns"] if not str(c).startswith("__")]
        self._add_condition_row(columns)
        self.filtered_rows = None
        self._refresh_filters_table(columns)

    def _refresh_filters_table(self, columns):
        rows = self.filtered_rows if self.filtered_rows is not None else self.consolidated["rows"]
        self.filters_result_label.configure(
            text=f"Showing {len(rows):,} of {len(self.consolidated['rows']):,} total row(s)."
        )
        self._load_table_with_total(self.filters_table, columns, rows)

    def _export_filtered_view(self):
        if not self.consolidated:
            dialogs.show_warning(self.root, "No Data", "Nothing to export yet.")
            return
        rows = self.filtered_rows if self.filtered_rows is not None else self.consolidated["rows"]
        columns = [c for c in self.consolidated["columns"] if not str(c).startswith("__")]
        self._run_export(lambda path: exporter.export_filtered_view(path, columns, rows, title="Filtered Data"),
                          default_name="Filtered_Export")

    # ================= ANALYTICS ================= #

    def show_analytics(self):
        self._set_active("analytics")
        if not self.consolidated:
            dialogs.show_warning(self.root, "No Data", "Please consolidate data first (Import Files \u2192 Consolidate).")
            self.show_home()
            return

        self._clear_workspace()
        self._screen_heading("Analytics", "Build your own chart: choose what to group by, what to measure, and how. "
                                            "Uses your active filters automatically.")

        columns = [c for c in self.consolidated["columns"] if not str(c).startswith("__")]
        num_cols = numeric_columns(self.consolidated["rows"], columns) or columns
        source_rows = self.filtered_rows if self.filtered_rows is not None else self.consolidated["rows"]

        ctk.CTkLabel(self.workspace,
                     text=f"Data source: {'Filtered view (' + str(len(source_rows)) + ' rows)' if self.filtered_rows is not None else 'Full consolidated data (' + str(len(source_rows)) + ' rows)'}",
                     font=("Segoe UI", 12, "italic"), text_color="#64748B").pack(anchor="w", pady=(0, 8))

        builder = ctk.CTkFrame(self.workspace, fg_color="white", corner_radius=8)
        builder.pack(fill="x", pady=(0, 12), ipady=10)

        self.chart_type_var = tk.StringVar(value=self.settings.get("default_chart_type", "Bar"))
        self.chart_group_var = tk.StringVar(value=columns[0] if columns else "")
        self.chart_value_var = tk.StringVar(value=num_cols[0] if num_cols else "")
        self.chart_agg_var = tk.StringVar(value=self.settings.get("default_aggregation", "Sum"))

        f1 = ctk.CTkFrame(builder, fg_color="transparent"); f1.pack(side="left", padx=14)
        ctk.CTkLabel(f1, text="Chart Type", font=("Segoe UI", 12)).pack(anchor="w")
        ctk.CTkOptionMenu(f1, values=["Bar", "Pie", "Line"], variable=self.chart_type_var, width=140).pack()

        f2 = ctk.CTkFrame(builder, fg_color="transparent"); f2.pack(side="left", padx=14)
        ctk.CTkLabel(f2, text="Group By (X-axis)", font=("Segoe UI", 12)).pack(anchor="w")
        ctk.CTkOptionMenu(f2, values=columns, variable=self.chart_group_var, width=170).pack()

        f3 = ctk.CTkFrame(builder, fg_color="transparent"); f3.pack(side="left", padx=14)
        ctk.CTkLabel(f3, text="Value Field (Y-axis)", font=("Segoe UI", 12)).pack(anchor="w")
        ctk.CTkOptionMenu(f3, values=num_cols, variable=self.chart_value_var, width=170).pack()

        f4 = ctk.CTkFrame(builder, fg_color="transparent"); f4.pack(side="left", padx=14)
        ctk.CTkLabel(f4, text="Aggregation", font=("Segoe UI", 12)).pack(anchor="w")
        ctk.CTkOptionMenu(f4, values=["Sum", "Count", "Average"], variable=self.chart_agg_var, width=130).pack()

        ctk.CTkButton(builder, text="\U0001F4CA Generate Chart", height=36,
                      command=lambda: self._generate_chart(source_rows)).pack(side="left", padx=20)

        self.chart_container = ctk.CTkFrame(self.workspace, fg_color="white", corner_radius=8, height=420)
        self.chart_container.pack(fill="both", expand=True, pady=(0, 10))

        self.analytics_export_btn = ctk.CTkButton(self.workspace, text="\U0001F4E4 Export This Analysis", height=38,
                                                    state="disabled", command=self._export_analysis)
        self.analytics_export_btn.pack(anchor="w")

        self._last_chart_data = None
        self._generate_chart(source_rows)

    def _generate_chart(self, rows):
        for w in self.chart_container.winfo_children():
            w.destroy()

        group_col = self.chart_group_var.get()
        value_col = self.chart_value_var.get()
        agg = self.chart_agg_var.get()
        chart_type = self.chart_type_var.get()

        buckets = {}
        for row in rows:
            key = str(row.get(group_col, "") or "(blank)")
            b = buckets.setdefault(key, {"sum": 0.0, "count": 0})
            v = _to_number(row.get(value_col))
            if v is not None:
                b["sum"] += v
            b["count"] += 1

        agg_data = {}
        for key, b in buckets.items():
            if agg == "Sum":
                agg_data[key] = round(b["sum"], 2)
            elif agg == "Count":
                agg_data[key] = b["count"]
            else:
                agg_data[key] = round(b["sum"] / b["count"], 2) if b["count"] else 0

        # keep the chart readable: top 15 categories by value
        sorted_items = sorted(agg_data.items(), key=lambda x: x[1], reverse=True)[:15]
        labels = [k for k, _ in sorted_items]
        values = [v for _, v in sorted_items]

        if not labels:
            ctk.CTkLabel(self.chart_container, text="No data available for this combination.",
                         font=("Segoe UI", 14)).pack(pady=40)
            self.analytics_export_btn.configure(state="disabled")
            return

        fig = Figure(figsize=(9, 4.6), dpi=100)
        ax = fig.add_subplot(111)

        if chart_type == "Bar":
            ax.bar(labels, values, color="#1E3A8A")
            ax.set_ylabel(f"{agg} of {value_col}")
            ax.tick_params(axis="x", rotation=45)
        elif chart_type == "Pie":
            ax.pie(values, labels=labels, autopct="%1.1f%%")
        else:  # Line
            ax.plot(labels, values, marker="o", color="#1E3A8A")
            ax.set_ylabel(f"{agg} of {value_col}")
            ax.tick_params(axis="x", rotation=45)

        ax.set_title(f"{agg} of {value_col} by {group_col}")
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.chart_container)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=8, pady=(8, 0))

        # NavigationToolbar2Tk gives built-in pan / zoom-to-rectangle / save-as-image
        toolbar_frame = ctk.CTkFrame(self.chart_container, fg_color="transparent")
        toolbar_frame.pack(fill="x")
        toolbar = NavigationToolbar2Tk(canvas, toolbar_frame)
        toolbar.update()

        self._last_chart_data = {"group_col": group_col, "value_col": value_col, "agg": agg, "rows": [
            {group_col: k, f"{agg} of {value_col}": v} for k, v in agg_data.items()
        ]}
        self.analytics_export_btn.configure(state="normal")

    def _export_analysis(self):
        if not self._last_chart_data:
            return
        rows = self._last_chart_data["rows"]
        columns = list(rows[0].keys()) if rows else []
        self._run_export(lambda path: exporter.export_filtered_view(path, columns, rows, title="Analysis Result"),
                          default_name="Analytics_Export")

    # ================= PIVOT TABLE ================= #

    def show_pivot(self):
        self._set_active("pivot")
        if not self.consolidated:
            dialogs.show_warning(self.root, "No Data", "Please consolidate data first (Import Files \u2192 Consolidate).")
            self.show_home()
            return

        self._clear_workspace()
        self._screen_heading("Pivot Table", "Summarize your data the way you would in Excel: pick rows, columns and a value field.")

        columns = [c for c in self.consolidated["columns"] if not str(c).startswith("__")]
        num_cols = numeric_columns(self.consolidated["rows"], columns) or columns
        source_rows = self.filtered_rows if self.filtered_rows is not None else self.consolidated["rows"]

        builder = ctk.CTkFrame(self.workspace, fg_color="white", corner_radius=8)
        builder.pack(fill="x", pady=(0, 12), ipady=10)

        self.pivot_rows_var = tk.StringVar(value=columns[0] if columns else "")
        self.pivot_cols_var = tk.StringVar(value="(None)")
        self.pivot_values_var = tk.StringVar(value=num_cols[0] if num_cols else "")
        self.pivot_agg_var = tk.StringVar(value=self.settings.get("default_aggregation", "Sum"))

        f1 = ctk.CTkFrame(builder, fg_color="transparent"); f1.pack(side="left", padx=14)
        ctk.CTkLabel(f1, text="Rows", font=("Segoe UI", 12)).pack(anchor="w")
        ctk.CTkOptionMenu(f1, values=columns, variable=self.pivot_rows_var, width=170).pack()

        f2 = ctk.CTkFrame(builder, fg_color="transparent"); f2.pack(side="left", padx=14)
        ctk.CTkLabel(f2, text="Columns (optional)", font=("Segoe UI", 12)).pack(anchor="w")
        ctk.CTkOptionMenu(f2, values=["(None)"] + columns, variable=self.pivot_cols_var, width=170).pack()

        f3 = ctk.CTkFrame(builder, fg_color="transparent"); f3.pack(side="left", padx=14)
        ctk.CTkLabel(f3, text="Values", font=("Segoe UI", 12)).pack(anchor="w")
        ctk.CTkOptionMenu(f3, values=num_cols, variable=self.pivot_values_var, width=150).pack()

        f4 = ctk.CTkFrame(builder, fg_color="transparent"); f4.pack(side="left", padx=14)
        ctk.CTkLabel(f4, text="Aggregation", font=("Segoe UI", 12)).pack(anchor="w")
        ctk.CTkOptionMenu(f4, values=["Sum", "Count", "Average", "Min", "Max"], variable=self.pivot_agg_var, width=130).pack()

        ctk.CTkButton(builder, text="\U0001F9EE Generate Pivot", height=36,
                      command=lambda: self._generate_pivot(source_rows)).pack(side="left", padx=20)

        self.pivot_table = ZoomableTable(self.workspace, height=18)
        self.pivot_table.pack(fill="both", expand=True, pady=(0, 10))

        ctk.CTkButton(self.workspace, text="\U0001F4E4 Export This Pivot Table", height=38,
                      command=self._export_pivot).pack(anchor="w")

        self._generate_pivot(source_rows)

    def _generate_pivot(self, rows):
        if not rows:
            self.pivot_table.clear()
            return

        row_field = self.pivot_rows_var.get()
        col_field = self.pivot_cols_var.get()
        value_field = self.pivot_values_var.get()
        agg_choice = self.pivot_agg_var.get()
        agg_map = {"Sum": "sum", "Count": "count", "Average": "mean", "Min": "min", "Max": "max"}

        try:
            df = pd.DataFrame(rows)
            df[value_field] = pd.to_numeric(df[value_field], errors="coerce")

            if col_field and col_field != "(None)":
                pivot = pd.pivot_table(df, index=row_field, columns=col_field, values=value_field,
                                        aggfunc=agg_map[agg_choice], fill_value=0)
                pivot.columns = [str(c) for c in pivot.columns]
            else:
                pivot = pd.pivot_table(df, index=row_field, values=value_field,
                                        aggfunc=agg_map[agg_choice], fill_value=0)

            pivot = pivot.reset_index()
            pivot_columns = [str(c) for c in pivot.columns]
            pivot_rows = pivot.round(2).to_dict(orient="records")

            self.pivot_result_columns = pivot_columns
            self.pivot_result_rows = pivot_rows
            self.pivot_table.set_data(pivot_columns, pivot_rows)
        except Exception as e:
            dialogs.show_error(self.root, "Pivot Error", f"Could not build pivot table: {e}")

    def _export_pivot(self):
        if not self.pivot_result_rows:
            dialogs.show_warning(self.root, "No Data", "Generate a pivot table first.")
            return
        self._run_export(lambda path: exporter.export_pivot_view(path, self.pivot_result_columns, self.pivot_result_rows),
                          default_name="Pivot_Export")

    # ================= SETTINGS ================= #

    def show_settings(self):
        self._set_active("settings")
        self._clear_workspace()
        self._screen_heading("Settings", "These defaults are remembered every time you open DataFusion Platform.")

        panel = ctk.CTkFrame(self.workspace, fg_color="white", corner_radius=10)
        panel.pack(fill="x", pady=(0, 10), ipady=14)

        # default export folder
        row1 = ctk.CTkFrame(panel, fg_color="transparent")
        row1.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(row1, text="Default export folder:", font=("Segoe UI", 13, "bold"), width=220, anchor="w").pack(side="left")
        self.settings_folder_var = tk.StringVar(value=self.settings.get("default_export_folder", ""))
        ctk.CTkEntry(row1, textvariable=self.settings_folder_var, width=420).pack(side="left", padx=8)
        ctk.CTkButton(row1, text="Browse...", width=90, command=self._browse_default_folder).pack(side="left")

        # default zoom
        row2 = ctk.CTkFrame(panel, fg_color="transparent")
        row2.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(row2, text="Default table zoom:", font=("Segoe UI", 13, "bold"), width=220, anchor="w").pack(side="left")
        self.settings_zoom_var = tk.StringVar(value=f"{self.settings.get('default_zoom', 100)}%")
        ctk.CTkOptionMenu(row2, values=[f"{z}%" for z in range(60, 201, 10)], variable=self.settings_zoom_var, width=120).pack(side="left")

        # default chart type
        row3 = ctk.CTkFrame(panel, fg_color="transparent")
        row3.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(row3, text="Default chart type:", font=("Segoe UI", 13, "bold"), width=220, anchor="w").pack(side="left")
        self.settings_chart_var = tk.StringVar(value=self.settings.get("default_chart_type", "Bar"))
        ctk.CTkOptionMenu(row3, values=["Bar", "Pie", "Line"], variable=self.settings_chart_var, width=120).pack(side="left")

        # default aggregation
        row4 = ctk.CTkFrame(panel, fg_color="transparent")
        row4.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(row4, text="Default aggregation:", font=("Segoe UI", 13, "bold"), width=220, anchor="w").pack(side="left")
        self.settings_agg_var = tk.StringVar(value=self.settings.get("default_aggregation", "Sum"))
        ctk.CTkOptionMenu(row4, values=["Sum", "Count", "Average"], variable=self.settings_agg_var, width=120).pack(side="left")

        # auto-open toggle
        row5 = ctk.CTkFrame(panel, fg_color="transparent")
        row5.pack(fill="x", padx=20, pady=10)
        self.settings_autoopen_var = tk.BooleanVar(value=self.settings.get("auto_open_after_export", True))
        ctk.CTkCheckBox(row5, text="Always ask to open the file in Excel right after exporting",
                         variable=self.settings_autoopen_var, font=("Segoe UI", 13)).pack(side="left")

        ctk.CTkButton(self.workspace, text="\U0001F4BE Save Settings", height=40, fg_color="#16A34A", hover_color="#15803d",
                      command=self._save_settings_action).pack(anchor="w", pady=10)

    def _browse_default_folder(self):
        folder = filedialog.askdirectory(title="Choose default export folder")
        if folder:
            self.settings_folder_var.set(folder)

    def _save_settings_action(self):
        self.settings["default_export_folder"] = self.settings_folder_var.get()
        self.settings["default_zoom"] = int(self.settings_zoom_var.get().replace("%", ""))
        self.settings["default_chart_type"] = self.settings_chart_var.get()
        self.settings["default_aggregation"] = self.settings_agg_var.get()
        self.settings["auto_open_after_export"] = self.settings_autoopen_var.get()
        settings_backend.save_settings(self.settings)
        dialogs.show_info(self.root, "Settings Saved", "Your defaults have been saved and will be used from now on.")

    # ================= EXPORT ================= #

    def export_data(self):
        self._set_active("export")
        if not self.consolidated:
            dialogs.show_warning(self.root, "No Data", "Please consolidate data first (Import Files \u2192 Consolidate).")
            self.show_home()
            return

        rows_to_export = self.consolidated["rows"]
        title_note = "full consolidated report"

        if self.filtered_rows is not None and len(self.filtered_rows) != len(self.consolidated["rows"]):
            choice = dialogs.ask_yes_no_cancel(
                self.root, "Export Options",
                "Filters are currently active on the Filters & Search screen.\n\n"
                "Yes  = Export the FULL consolidated report (all sheets, all rows)\n"
                "No   = Export only the FILTERED rows currently shown\n"
                "Cancel = don't export"
            )
            if choice is None:
                return
            if not choice:
                columns = [c for c in self.consolidated["columns"] if not str(c).startswith("__")]
                self._run_export(lambda path: exporter.export_filtered_view(path, columns, self.filtered_rows, title="Filtered Export"),
                                  default_name="Filtered_Export")
                return

        self._run_export(
            lambda path: exporter.export_full_report(
                path, self.consolidated, gst_summary=self.gst_summary, party_summary=self.party_summary,
                validation_issues=self.all_issues, source_files=self.source_files_used,
                company_summary=self.company_summary, month_summary=self.month_summary,
                consolidated_at=self.consolidated_at
            ),
            default_name="DataFusion_Export"
        )

    def _run_export(self, export_fn, default_name="Export"):
        """Shared save-dialog + save + 'open now?' flow used by every export button in the app."""
        default_folder = self.settings.get("default_export_folder") or os.path.expanduser("~")
        os.makedirs(default_folder, exist_ok=True)
        default_filename = f"{default_name}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

        path = filedialog.asksaveasfilename(
            title="Save Excel Report",
            initialdir=default_folder,
            initialfile=default_filename,
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not path:
            return

        try:
            export_fn(path)
        except Exception as e:
            dialogs.show_error(self.root, "Export Failed", f"Could not save the file:\n{e}")
            return

        self.last_export_path = path
        self._show_export_success_screen(path)

        if self.settings.get("auto_open_after_export", True):
            if dialogs.ask_yes_no(self.root, "Export Complete", "Your Excel file has been saved successfully.\n\nOpen it now in Excel?"):
                if not exporter.open_file(path):
                    dialogs.show_warning(self.root, "Couldn't Open", "The file was saved, but it couldn't be opened automatically.\n"
                                                              f"You can find it at:\n{path}")

    def _show_export_success_screen(self, path):
        self._clear_workspace()
        self._screen_heading("Export Complete", "Your report has been saved and formatted, ready to use.")

        ctk.CTkLabel(self.workspace, text=f"\u2714 Saved to:\n{path}", font=("Segoe UI", 14),
                     justify="left", wraplength=900).pack(anchor="w", pady=(0, 20))

        btn_row = ctk.CTkFrame(self.workspace, fg_color="transparent")
        btn_row.pack(anchor="w")

        ctk.CTkButton(btn_row, text="\U0001F4C2 Open File in Excel", height=42, width=220, fg_color="#16A34A",
                      hover_color="#15803d", command=lambda: exporter.open_file(path)).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_row, text="\U0001F4C1 Open Containing Folder", height=42, width=220,
                      command=lambda: exporter.open_containing_folder(path)).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_row, text="\u2B05 Back to Dashboard", height=42, width=180,
                      fg_color="#94A3B8", hover_color="#64748B", command=self.show_home).pack(side="left")
