import os
import platform
import re
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog
import customtkinter as ctk
import pandas as pd

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from ui.file_dialog import select_excel_files, select_excel_folder, select_parent_folder, list_candidate_folders
from ui.widgets import StatCard, ZoomableTable, ZoomableTextbox
from ui import dialogs

from backend.excel_engine import ExcelEngine
from backend.database import DataStore
from backend.validator import validate_sheet, humanize_issue
from backend.consolidator import (
    merge_files, gst_rate_summary, party_wise_summary,
    company_wise_summary, month_wise_summary, credit_debit_checkpoint,
)
from backend.filters import quick_search, apply_filters, numeric_columns, grand_totals, OPERATORS, _to_number, format_indian_number
from backend import exporter
from backend import settings as settings_backend
from backend import auto_organize

ACTIVE_COLOR = "#F59E0B"
DEFAULT_BTN_COLOR = "#2563EB"

# Matches the auto-generated placeholder the backend gives a column when the
# source Excel sheet had a blank header (e.g. "Column D (no header in source
# file)") -- used to visually flag those entries on the Review Data page.
NO_HEADER_PATTERN = re.compile(r"no head(er|ing) in source file", re.IGNORECASE)


DEFAULT_REPORT_SECTIONS = {
    "checkpoint": True, "company": True, "source_files": True,
    "gst": False, "monthly": False, "party": False, "custom": [],
}


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
        self.credit_debit_summary = None
        self.report_sections = None
        self.source_files_used = []
        self._pending_consolidation_selection = None  # last {file_name:[sheet,...]} attempt, for Checkpoint Refresh
        self.all_issues = {}              # aggregated validation issues, labeled

        self.filtered_rows = None         # None => no filter active, show all consolidated rows
        self.filter_condition_rows = []   # widgets for the advanced filter builder
        self.last_export_path = None

        self.pivot_result_columns = []
        self.pivot_result_rows = []

        self.menu_buttons = {}
        self.current_screen = "home"

        # ---------- Back / Forward / Refresh screen history ----------
        # Tracks the simple, no-argument "screen views" (Dashboard,
        # Review Data, the Consolidate preview, Filters, Analytics,
        # Pivot, Settings). Deliberately does NOT track steps that pop
        # a system file dialog (Import Files/Folder) or the final
        # Consolidation Results screen (that one is the product of a
        # multi-step wizard with popups, not a plain re-callable view --
        # replaying it would re-run a merge and re-open dialogs).
        self._nav_stack = []
        self._nav_index = -1
        self._nav_replaying = False

        self.create_header()
        self.create_body()

    # ================= HEADER ================= #

    def create_header(self):
        header = ctk.CTkFrame(self.root, height=70, corner_radius=0, fg_color="#1E3A8A")
        header.pack(fill="x")

        nav_box = ctk.CTkFrame(header, fg_color="transparent")
        nav_box.pack(side="left", padx=(20, 4), pady=15)

        self.nav_back_btn = ctk.CTkButton(
            nav_box, text="\u25C0", width=44, height=44, corner_radius=22,
            fg_color="#3B5BA5", hover_color="#2C4884", font=("Segoe UI", 16),
            command=self._nav_go_back
        )
        self.nav_back_btn.pack(side="left", padx=(0, 8))

        self.nav_fwd_btn = ctk.CTkButton(
            nav_box, text="\u25B6", width=44, height=44, corner_radius=22,
            fg_color="#3B5BA5", hover_color="#2C4884", font=("Segoe UI", 16),
            command=self._nav_go_forward
        )
        self.nav_fwd_btn.pack(side="left", padx=(0, 8))

        self.nav_refresh_btn = ctk.CTkButton(
            nav_box, text="\u21BB", width=44, height=44, corner_radius=22,
            fg_color="#3B5BA5", hover_color="#2C4884", font=("Segoe UI", 16),
            command=self._nav_refresh
        )
        self.nav_refresh_btn.pack(side="left")

        ctk.CTkLabel(header, text="DataFusion Platform", font=("Segoe UI", 26, "bold"),
                     text_color="white").pack(side="left", padx=(10, 25), pady=15)

        ctk.CTkLabel(header, text="Enterprise Data Consolidation & Analysis for Finance/CA teams",
                     font=("Segoe UI", 13), text_color="#CBD5E1").pack(side="right", padx=25)

        self._update_nav_buttons()

    # ---------- Back / Forward / Refresh ---------- #

    def _record_nav(self, screen_fn):
        """
        Called at the top of each simple, replayable screen method.
        Pushes it onto the history stack so Back/Forward can step
        through it -- unless we're currently REPLAYING history (Back/
        Forward triggered this same call), in which case we just leave
        the stack alone.
        """
        if self._nav_replaying:
            return
        # Cut any "forward" history once the user navigates somewhere new.
        self._nav_stack = self._nav_stack[:self._nav_index + 1]
        self._nav_stack.append(screen_fn)
        self._nav_index = len(self._nav_stack) - 1
        self._update_nav_buttons()

    def _update_nav_buttons(self):
        if not hasattr(self, "nav_back_btn"):
            return
        self.nav_back_btn.configure(state="normal" if self._nav_index > 0 else "disabled")
        self.nav_fwd_btn.configure(state="normal" if self._nav_index < len(self._nav_stack) - 1 else "disabled")
        self.nav_refresh_btn.configure(state="normal" if self._nav_stack else "disabled")

    def _nav_go_back(self):
        if self._nav_index <= 0:
            return
        self._nav_index -= 1
        self._nav_replaying = True
        try:
            self._nav_stack[self._nav_index]()
        finally:
            self._nav_replaying = False
        self._update_nav_buttons()

    def _nav_go_forward(self):
        if self._nav_index >= len(self._nav_stack) - 1:
            return
        self._nav_index += 1
        self._nav_replaying = True
        try:
            self._nav_stack[self._nav_index]()
        finally:
            self._nav_replaying = False
        self._update_nav_buttons()

    def _nav_refresh(self):
        """Redraws the current screen with whatever data is current now --
        does not change your position in the Back/Forward history."""
        if not self._nav_stack:
            return
        self._nav_replaying = True
        try:
            self._nav_stack[self._nav_index]()
        finally:
            self._nav_replaying = False

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
            ("import_folder", "\U0001F4C1 Import Folder", self.import_folder),
            ("review", "\U0001F4DD Review Data", self.show_review_data),
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

        # A permanent footer bar OUTSIDE the scrollable workspace. Screens
        # that want an always-reachable call-to-action row (like Review
        # Data's "Continue to Consolidate") build it here instead of
        # inside self.workspace -- that's what makes it 100% immune to
        # the scrollable frame's scroll-region ever hiding it, instead of
        # relying on a manual scrollregion refresh that isn't reliable.
        self.footer_bar = ctk.CTkFrame(body, fg_color="white", corner_radius=0, border_width=0)
        self.footer_bar.pack(side="bottom", fill="x")

        self.workspace = ctk.CTkScrollableFrame(body, fg_color="#F5F7FA")
        self.workspace.pack(side="left", fill="both", expand=True, padx=20, pady=20)

        self.show_home()

    def _clear_footer(self):
        for widget in self.footer_bar.winfo_children():
            widget.destroy()

    def _set_active(self, key):
        self.current_screen = key
        for k, btn in self.menu_buttons.items():
            btn.configure(fg_color=ACTIVE_COLOR if k == key else DEFAULT_BTN_COLOR)

    def _clear_workspace(self):
        for widget in self.workspace.winfo_children():
            widget.destroy()
        self._clear_footer()

    def _scroll_workspace_to_top(self):
        # Ensures the page always opens showing its top content instead of
        # staying at whatever scroll position was left over from before.
        try:
            self.workspace._parent_canvas.yview_moveto(0.0)
        except Exception:
            pass

    def _refresh_workspace_scrollregion(self):
        # CTkScrollableFrame doesn't always recompute its scrollable area
        # correctly right after inserting a large, variable-height widget
        # (like the merged-data table or a summary textbox), which can
        # leave the bottom of the page unreachable. Force a recompute.
        try:
            self.workspace.update_idletasks()
            canvas = self.workspace._parent_canvas
            canvas.configure(scrollregion=canvas.bbox("all"))
        except Exception:
            pass

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
        self._record_nav(self.show_home)
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
        # Old validation results were computed against the file list that
        # just changed -- drop them so a later export can't accidentally
        # include issues that belonged to a file that's no longer loaded.
        self.all_issues = {}
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

    def _get_consolidated_scope_results(self):
        """
        Same as _get_selected_results(), but narrowed down further to
        exactly the individual SHEETS the user picked in "Choose Files &
        Sheets" the last time Consolidate ran (self._pending_consolidation_
        selection) -- not every sheet in every ticked file.

        Example: a file has sheets "FA Q1 2025", "TB Q1 2025", "BS Q1
        2025" and the user only ticked "TB Q1 2025" to consolidate.
        Without this, a report generated after Export would still run
        validation against (and could show issues from) the FA and BS
        sheets too, even though neither was actually included in the
        consolidated data. This makes sure only what was actually
        selected/consolidated is ever considered.

        Falls back to the plain ticked-file selection if Consolidate
        hasn't been run yet in this session (nothing to narrow down to).
        """
        selection = self._pending_consolidation_selection
        if not selection:
            return self._get_selected_results()
        scoped = []
        for r in self.loaded_results:
            fname = r["file_name"]
            if fname not in selection:
                continue
            wanted_sheets = set(selection[fname])
            filtered = dict(r)
            filtered["sheets"] = {s: data for s, data in r["sheets"].items() if s in wanted_sheets}
            if filtered["sheets"]:
                scoped.append(filtered)
        return scoped

    # ================= IMPORT FILES ================= #

    def import_files(self):
        self._set_active("import")
        self._run_import_session(start_mode="files")

    def import_folder(self):
        """
        Lets the user pick whole folders instead of ctrl-clicking
        individual workbooks. Every Excel file found directly inside a
        chosen folder is imported and tagged with that folder's name,
        so "Choose Files to Consolidate" can group them under a folder
        heading (folder -> workbook -> sheets).
        """
        self._set_active("import_folder")
        self._run_import_session(start_mode="folder")

    def _run_import_session(self, start_mode):
        """
        Builds up ONE import batch out of any mix of individual files
        and whole folders -- e.g. two folders plus a couple of loose
        files -- all picked in the same sitting, before anything is
        actually loaded. That way there's a single "Company Names"
        prompt and a single "Import Complete" summary at the end,
        instead of one per pick.

        start_mode: "files" or "folder" -- which picker opens first,
        based on which menu button (Import Files / Import Folder) was
        clicked. After that, the user can keep adding either kind.
        """
        path_folder_pairs = []  # [(path, folder_name_or_None), ...]
        mode = start_mode
        cancelled = False

        while mode:
            if mode == "cancel":
                cancelled = True
                break
            if mode == "files":
                picked = select_excel_files(self.root)
                if picked:
                    path_folder_pairs.extend((p, None) for p in picked)
            else:  # "folder"
                path_folder_pairs.extend(self._pick_multiple_folders())

            mode = dialogs.ask_add_more(self.root, added_so_far=len(path_folder_pairs))

        if cancelled or not path_folder_pairs:
            self.show_home()
            return

        self._import_paths(path_folder_pairs)

    def _pick_multiple_folders(self):
        """
        Windows' native folder browser only ever returns ONE folder per
        click -- Shift/Ctrl there just highlights, the dialog still
        reports a single path back, because that's how the OS dialog
        itself works. This is the actual workaround: pick a parent
        folder once, then tick as many of its subfolders as you want
        in a single checklist screen.

        Returns a flat list of (file_path, folder_name) pairs covering
        every ticked folder, or [] if the user cancelled at any step.
        """
        base_folder = select_parent_folder(self.root)
        if not base_folder:
            return []

        candidates = list_candidate_folders(base_folder)
        if not candidates:
            dialogs.show_warning(
                self.root, "No Excel Files Found",
                "That folder (and its subfolders) has no Excel files in it directly. "
                "Try picking a different parent folder."
            )
            return []

        chosen_paths = dialogs.ask_pick_folders(self.root, base_folder, candidates)
        if not chosen_paths:
            return []

        pairs = []
        for folder_path in chosen_paths:
            folder_name = os.path.basename(folder_path.rstrip("/\\")) or folder_path
            for name in sorted(os.listdir(folder_path)):
                if name.startswith("~$"):
                    continue
                if name.lower().endswith((".xlsx", ".xls")):
                    pairs.append((os.path.join(folder_path, name), folder_name))
        return pairs

    def _import_paths(self, new_paths, source_folder=None):
        """
        new_paths: either a flat list of file paths (all tagged with the
        single `source_folder`, or no folder at all for plain Import
        Files), or a list of (path, folder_name) tuples when multiple
        folders were queued up together by Import Folder.
        """
        if new_paths and isinstance(new_paths[0], tuple):
            path_folder_pairs = new_paths
        else:
            path_folder_pairs = [(p, source_folder) for p in new_paths]

        errors = []
        added = 0
        newly_added_files = []
        folders_used = []
        existing_names = {r["file_name"] for r in self.loaded_results}

        # ---- Pass 1: load every file, with a real (non-stuck) progress bar,
        # and detect which ones are messy exported reports. ----
        progress = dialogs.show_import_progress(self.root, total=len(path_folder_pairs))
        results_by_path = {}
        raw_by_path = {}
        messy_paths = []
        try:
            for i, (path, folder_name) in enumerate(path_folder_pairs, start=1):
                progress.update(i, len(path_folder_pairs), os.path.basename(path))
                result = self.engine.load_excel(path)
                if "error" in result:
                    errors.append(f"{os.path.basename(path)}: {result['error']}")
                    continue
                results_by_path[path] = result
                raw = self.engine.load_excel_raw(path)
                if "error" not in raw and auto_organize.needs_organizing(raw):
                    raw_by_path[path] = raw
                    messy_paths.append(path)
        finally:
            progress.close()

        # ---- If any file looks messy, ask ONCE for the whole batch. ----
        import_mode = "raw"
        if messy_paths:
            raw_preview, organized_preview = auto_organize.build_previews(raw_by_path[messy_paths[0]])
            import_mode = dialogs.ask_import_format(self.root, raw_preview, organized_preview) or "raw"

        ambiguous_by_file = {}  # file_name -> {sheet_name: [{"placeholder","samples"}, ...]}

        for path, folder_name in path_folder_pairs:
            if path not in results_by_path:
                continue
            result = results_by_path[path]
            if import_mode == "auto" and path in raw_by_path:
                organized = auto_organize.organize_workbook_raw(raw_by_path[path])
                if organized["sheets"]:
                    result = organized
                    if organized.get("ambiguous_columns"):
                        ambiguous_by_file[organized["file_name"]] = organized["ambiguous_columns"]
            # re-importing a file with the same name replaces the old copy
            if result["file_name"] in existing_names:
                self.loaded_results = [r for r in self.loaded_results if r["file_name"] != result["file_name"]]
            if folder_name:
                result["source_folder"] = folder_name
                if folder_name not in folders_used:
                    folders_used.append(folder_name)
            self.loaded_results.append(result)
            self.file_check_vars[result["file_name"]] = tk.BooleanVar(value=True)
            self.store.add_upload(result["file_name"], result)
            existing_names.add(result["file_name"])
            newly_added_files.append(result["file_name"])
            added += 1

        # ---- If Auto-Organize left any headerless-but-has-data columns,
        # ask the user to name them ONCE for the whole batch -- with real
        # sample values shown -- instead of silently guessing "Account",
        # "Account (2)", etc. Blank/skipped answers keep the placeholder
        # name, still safely editable later from the red-flagged box on
        # the Review Data / Preview screens. ----
        if ambiguous_by_file:
            renames = dialogs.ask_column_names(self.root, ambiguous_by_file)
            if renames:
                for r in self.loaded_results:
                    sheet_renames = renames.get(r["file_name"])
                    if not sheet_renames:
                        continue
                    for sheet_name, col_renames in sheet_renames.items():
                        sheet = r["sheets"].get(sheet_name)
                        if not sheet or not col_renames:
                            continue
                        sheet["columns"] = [col_renames.get(c, c) for c in sheet["columns"]]
                        sheet["data"] = [
                            {col_renames.get(k, k): v for k, v in row.items()}
                            for row in sheet["data"]
                        ]

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

        if added:
            self.show_consolidate_preview()
        else:
            self.show_home()

        if errors:
            dialogs.show_error(self.root, "Import Errors", "\n".join(errors))
        if added:
            folder_note = f" from folder(s): {', '.join(folders_used)}" if folders_used else ""
            dialogs.show_info(
                self.root, "Import Complete",
                f"Loaded {added} file(s){folder_note} successfully.\n"
                f"Total files now available: {len(self.loaded_results)}\n\n"
                f"Review the column headings and remove any empty rows below, "

                f"then continue to Validate or Consolidate."
            )

    # ================= REVIEW DATA ================= #

    def show_review_data(self, preserve_scroll=False):
        scroll_pos = None
        if preserve_scroll:
            try:
                scroll_pos = self.workspace._parent_canvas.yview()[0]
            except Exception:
                scroll_pos = None

        self._set_active("review")
        self._record_nav(self.show_review_data)
        self._clear_workspace()
        self._screen_heading(
            "Review & Clean Data",
            "Rename column headings and remove fully empty rows before you Validate or Consolidate."
        )

        if not self.loaded_results:
            ctk.CTkLabel(
                self.workspace,
                text="No files imported yet. Click '\U0001F4C2 Import Files' in the menu first.",
                justify="left", anchor="w", font=("Segoe UI", 14)
            ).pack(fill="x", pady=20)
            return

        self._review_header_entries = {}  # (file_name, sheet_name) -> [entry widgets]

        for result in self.loaded_results:
            fname = result["file_name"]
            file_box = ctk.CTkFrame(self.workspace, fg_color="white", corner_radius=8)
            file_box.pack(fill="x", pady=(0, 14), padx=2)

            ctk.CTkLabel(file_box, text=f"\U0001F4C4 {fname}", font=("Segoe UI", 15, "bold"),
                         text_color="#1E3A8A").pack(anchor="w", padx=14, pady=(12, 4))

            for sheet_name, sheet in result["sheets"].items():
                sheet_frame = ctk.CTkFrame(file_box, fg_color="#F8FAFC", corner_radius=6)
                sheet_frame.pack(fill="x", padx=14, pady=6)

                top_row = ctk.CTkFrame(sheet_frame, fg_color="transparent")
                top_row.pack(fill="x", padx=10, pady=(10, 4))
                ctk.CTkLabel(top_row, text=f"Sheet: {sheet_name}   ({sheet['row_count']} row(s))",
                             font=("Segoe UI", 13, "bold"), text_color="#334155").pack(side="left")

                empty_row_count = sum(
                    1 for row in sheet["data"]
                    if sheet["columns"] and all(str(row.get(c, "")).strip() == "" for c in sheet["columns"])
                )
                if empty_row_count:
                    ctk.CTkButton(top_row, text=f"\U0001F9F9 Clean Empty Rows ({empty_row_count})", width=190, height=28,
                                  fg_color="#DC2626", hover_color="#991b1b", font=("Segoe UI", 12),
                                  command=lambda f=fname, s=sheet_name: self._remove_empty_rows(f, s)
                                  ).pack(side="right")
                else:
                    ctk.CTkLabel(top_row, text="\u2714 No empty rows", font=("Segoe UI", 11, "italic"),
                                 text_color="#16A34A").pack(side="right")

                ctk.CTkLabel(sheet_frame, text="Column headings (edit text, then click 'Save Headings'):",
                             font=("Segoe UI", 12, "italic"), text_color="#64748B").pack(anchor="w", padx=10)

                headers_area = ctk.CTkFrame(sheet_frame, fg_color="transparent")
                headers_area.pack(fill="x", padx=10, pady=(4, 4))

                entries = []
                missing_count = 0
                for col in sheet["columns"]:
                    is_missing = bool(NO_HEADER_PATTERN.search(str(col)))
                    if is_missing:
                        missing_count += 1

                    col_row = ctk.CTkFrame(headers_area, fg_color="transparent")
                    col_row.pack(fill="x", pady=2)

                    if is_missing:
                        ctk.CTkLabel(
                            col_row,
                            text="\u26A0 No heading found in the source file \u2014 type one below:",
                            font=("Segoe UI", 11, "bold"), text_color="#B91C1C"
                        ).pack(anchor="w")

                    entry_row = ctk.CTkFrame(col_row, fg_color="transparent")
                    entry_row.pack(fill="x")

                    entry = ctk.CTkEntry(
                        entry_row, height=30, font=("Segoe UI", 12),
                        fg_color="#FEF2F2" if is_missing else None,
                        border_color="#DC2626" if is_missing else None,
                        border_width=2 if is_missing else None,
                        text_color="#991B1B" if is_missing else None,
                    )
                    entry.pack(side="left", fill="x", expand=True)
                    entry.insert(0, col)
                    entries.append(entry)

                    if is_missing:
                        ctk.CTkButton(
                            entry_row, text="\U0001F5D1", width=32, height=30,
                            fg_color="#DC2626", hover_color="#991b1b", font=("Segoe UI", 13),
                            command=lambda f=fname, s=sheet_name, c=col: self._delete_column(f, s, c)
                        ).pack(side="left", padx=(6, 0))

                self._review_header_entries[(fname, sheet_name)] = entries

                if missing_count:
                    ctk.CTkLabel(
                        sheet_frame,
                        text=f"\u26A0 {missing_count} column(s) above had no heading in the source file "
                             f"\u2014 highlighted in red. Please type a heading for each before saving.",
                        font=("Segoe UI", 11, "italic"), text_color="#B91C1C"
                    ).pack(anchor="w", padx=10, pady=(0, 4))

                ctk.CTkButton(sheet_frame, text="\U0001F4BE Save Headings", width=150, height=30,
                              fg_color="#2563EB", hover_color="#1d4ed8", font=("Segoe UI", 12),
                              command=lambda f=fname, s=sheet_name: self._save_headers(f, s)
                              ).pack(anchor="e", padx=10, pady=(2, 10))

        # Built in the permanent footer bar (outside the scrollable
        # workspace) so it's ALWAYS visible regardless of how much
        # content is above it or how the scrollable frame's internal
        # scroll region behaves -- no more relying on a scroll-region
        # refresh that could silently get overwritten.
        action_row = ctk.CTkFrame(self.footer_bar, fg_color="white")
        action_row.pack(fill="x", padx=22, pady=12)
        ctk.CTkLabel(
            action_row,
            text="Happy with the headings and rows? Continue to Validate or Consolidate.",
            font=("Segoe UI", 12), text_color="#64748B"
        ).pack(side="left")
        ctk.CTkButton(action_row, text="\U0001F4CA Continue to Consolidate \u2192", height=36,
                      fg_color="#0F766E", hover_color="#0b5c54",
                      command=self.consolidate_data).pack(side="right", padx=(6, 0))
        ctk.CTkButton(action_row, text="\u2714 Continue to Validate Data \u2192", height=36,
                      command=self.validate_data).pack(side="right", padx=(6, 0))

        if preserve_scroll and scroll_pos is not None:
            self.root.after(50, lambda: self.workspace._parent_canvas.yview_moveto(scroll_pos))
        else:
            self.root.after(50, self._scroll_workspace_to_top)

    def _remove_empty_rows(self, fname, sheet_name, return_to=None):
        """Deletes rows where every column for that row is blank.
        return_to: which screen to redraw afterward (defaults to Review Data)."""
        result = next((r for r in self.loaded_results if r["file_name"] == fname), None)
        if not result:
            return
        sheet = result["sheets"].get(sheet_name)
        if not sheet:
            return

        columns = sheet["columns"]
        before = len(sheet["data"])
        sheet["data"] = [
            row for row in sheet["data"]
            if any(str(row.get(c, "")).strip() != "" for c in columns)
        ]
        sheet["row_count"] = len(sheet["data"])
        removed = before - sheet["row_count"]

        dialogs.show_info(
            self.root, "Empty Rows Removed",
            f"Removed {removed} fully empty row(s) from '{sheet_name}' in {fname}."
        )
        target_screen = return_to or self.show_review_data
        target_screen(preserve_scroll=True)

    def _delete_column(self, fname, sheet_name, col_name, return_to=None):
        """Removes one column entirely (heading + every row's value for it)
        -- used for the trash button on 'No heading in source file' columns
        in Review Data, so an empty/unwanted column doesn't need a typed
        heading just to move past it."""
        result = next((r for r in self.loaded_results if r["file_name"] == fname), None)
        if not result:
            return
        sheet = result["sheets"].get(sheet_name)
        if not sheet or col_name not in sheet["columns"]:
            return

        if not dialogs.ask_yes_no(
            self.root, "Delete Column",
            f"Delete the column '{col_name}' from '{sheet_name}' in {fname}? "
            f"This removes it from every row and can't be undone."
        ):
            return

        sheet["columns"] = [c for c in sheet["columns"] if c != col_name]
        for row in sheet["data"]:
            row.pop(col_name, None)

        # Deferred + scroll-preserving: rebuilding immediately inside the
        # trash button's own click handler could crash on some systems
        # (a pending hover/redraw callback firing on a button that's
        # already been destroyed), and jumping back to the top of a long
        # Review Data page is disorienting -- keep the user right where
        # they were.
        target_screen = return_to or self.show_review_data
        self.root.after(10, lambda: target_screen(preserve_scroll=True))

    def _save_headers(self, fname, sheet_name, return_to=None):
        """Applies the edited column heading text boxes back onto the sheet's data.
        return_to: which screen to redraw afterward (defaults to Review Data)."""
        result = next((r for r in self.loaded_results if r["file_name"] == fname), None)
        if not result:
            return
        sheet = result["sheets"].get(sheet_name)
        if not sheet:
            return

        entries = self._review_header_entries.get((fname, sheet_name), [])
        old_columns = sheet["columns"]

        new_columns = []
        seen = {}
        for entry, old in zip(entries, old_columns):
            text = entry.get().strip() or old
            if text in seen:
                seen[text] += 1
                text = f"{text} ({seen[text]})"
            else:
                seen[text] = 1
            new_columns.append(text)

        rename_map = dict(zip(old_columns, new_columns))
        sheet["data"] = [{rename_map.get(k, k): v for k, v in row.items()} for row in sheet["data"]]
        sheet["columns"] = new_columns

        # No confirmation popup here on purpose -- it's an unnecessary
        # interruption for a routine save. Just refresh in place,
        # keeping the user's current scroll position either way.
        target_screen = return_to or self.show_review_data
        self.root.after(10, lambda: target_screen(preserve_scroll=True))

    # ================= VALIDATE ================= #

    def _compute_all_issues(self, selected, output_lines=None):
        """
        Rebuilds the validation-issues dict FROM SCRATCH for exactly the
        file list passed in (`selected`). This is the single source of
        truth for "Issues Found" -- both the Validate Data screen and
        Export call this instead of trusting whatever self.all_issues
        happened to hold before, so removed/deselected files can never
        leak into a later export.

        If `output_lines` (a list) is passed, the human-readable preview
        lines used by the Validate Data screen are appended to it.
        """
        all_issues = {}
        for result in selected:
            fname = result["file_name"]
            if output_lines is not None:
                output_lines.append(f"FILE: {fname}")
            for sheet_name, sheet in result["sheets"].items():
                report = validate_sheet(sheet["data"], sheet["columns"])
                if output_lines is not None:
                    output_lines.append(f"  Sheet: {sheet_name}")
                    output_lines.append(f"  {report['summary']}")
                    for check in report["skipped_checks"]:
                        output_lines.append(f"    (skipped) {check}")
                for issue_type, items in report["issues"].items():
                    if items:
                        if output_lines is not None:
                            output_lines.append(f"    {issue_type}: {len(items)} found")
                            for item in items[:10]:
                                output_lines.append(f"       - {humanize_issue(issue_type, item)}")
                            if len(items) > 10:
                                output_lines.append(f"       ... and {len(items) - 10} more")
                        labeled = all_issues.setdefault(issue_type, [])
                        for item in items:
                            labeled.append(f"File: {fname} | Sheet: {sheet_name} | {humanize_issue(issue_type, item)}")
                if output_lines is not None:
                    output_lines.append("")
            if output_lines is not None:
                output_lines.append("")
        return all_issues

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

        self._record_nav(self.validate_data)
        self._clear_workspace()
        self._screen_heading("Validation Results", "Every issue below tells you exactly which file, sheet and row it is in.")

        output_lines = [f"Validating {len(selected)} of {len(self.loaded_results)} imported file(s) (ticked only)", ""]
        self.all_issues = self._compute_all_issues(selected, output_lines=output_lines)

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

        self.root.after(50, self._scroll_workspace_to_top)

    # ================= CONSOLIDATE ================= #

    def _apply_custom_commands(self, sections):
        """
        The "+" box on the report-sections popup lets someone type
        anything. Most of it can only be carried through as a plain
        reminder note (there's no safe way to invent a brand-new
        calculation from free text) -- but a couple of common requests
        ARE things this app can actually turn on, so those are
        recognized and genuinely applied here instead of just being
        noted down. Anything not recognized is left as a plain note,
        same as before.

        Mutates and returns `sections` in place. Sets
        sections["_executed_custom"] = [the items that were actually
        applied] so the results screen can show a clear confirmation
        that they took effect, not just got written down.
        """
        executed = []
        remaining = []
        for item in sections.get("custom", []):
            low = item.strip().lower()
            if "full summary" in low or "complete summary" in low or "everything" in low or "all summary" in low:
                # "Full Summary" -> force every section on, overriding
                # whatever was individually unticked above.
                for key in ("checkpoint", "company", "source_files"):
                    sections[key] = True
                executed.append(item)
            elif "grand total" in low:
                # "Grand Total" -> guarantees a clearly labelled Grand
                # Total recap is shown on-screen and in the export,
                # beyond the Grand Total row that's already always on
                # every data sheet.
                sections["emphasize_grand_total"] = True
                executed.append(item)
            else:
                remaining.append(item)

        sections["custom"] = remaining
        sections["_executed_custom"] = executed
        return sections

    def consolidate_data(self):
        self._set_active("consolidate")
        if not self.loaded_results:
            dialogs.show_warning(self.root, "No Data", "Please import files first.")
            self.show_home()
            return
        self.show_consolidate_preview()

    def show_consolidate_preview(self, preserve_scroll=False):
        """
        Shown right when the user clicks Consolidate, BEFORE the file/
        sheet picker. Gives them an actual, Excel-style look at every
        imported sheet's real data (not just headers/row-counts like
        Review Data) so they know exactly what they're about to merge.

        Also where flagged/blank column headings get fixed or deleted --
        this is now the main stop after Import (Review Data still
        exists in the menu if you want it, it's just no longer forced
        on you right after importing).
        """
        scroll_pos = None
        if preserve_scroll:
            try:
                scroll_pos = self.workspace._parent_canvas.yview()[0]
            except Exception:
                scroll_pos = None

        self._set_active("consolidate")
        self._record_nav(self.show_consolidate_preview)
        self._clear_workspace()
        self._screen_heading(
            "Preview Before Consolidation",
            "This is exactly what's in each imported sheet. Fix any red-flagged column headings below its "
            "grid, then continue to pick which files/sheets to merge."
        )

        self._review_header_entries = {}  # (file_name, sheet_name) -> [entry widgets], reused by _save_headers

        for result in self.loaded_results:
            fname = result["file_name"]
            file_box = ctk.CTkFrame(self.workspace, fg_color="white", corner_radius=8)
            file_box.pack(fill="x", pady=(0, 14), padx=2)

            ctk.CTkLabel(file_box, text=f"\U0001F4C4 {fname}", font=("Segoe UI", 15, "bold"),
                         text_color="#1E3A8A").pack(anchor="w", padx=14, pady=(12, 4))

            for sheet_name, sheet in result["sheets"].items():
                sheet_frame = ctk.CTkFrame(file_box, fg_color="#F8FAFC", corner_radius=6)
                sheet_frame.pack(fill="x", padx=14, pady=6)

                top_row = ctk.CTkFrame(sheet_frame, fg_color="transparent")
                top_row.pack(fill="x", padx=10, pady=(10, 4))
                ctk.CTkLabel(top_row, text=f"Sheet: {sheet_name}   ({sheet['row_count']} row(s))",
                             font=("Segoe UI", 13, "bold"), text_color="#334155").pack(side="left")

                if sheet.get("title"):
                    ctk.CTkLabel(sheet_frame, text=sheet["title"],
                                 font=("Segoe UI", 12, "bold"), text_color="#1E3A8A").pack(anchor="w", padx=10, pady=(0, 4))

                empty_row_count = sum(
                    1 for row in sheet["data"]
                    if sheet["columns"] and all(str(row.get(c, "")).strip() == "" for c in sheet["columns"])
                )
                if empty_row_count:
                    ctk.CTkButton(top_row, text=f"\U0001F9F9 Clean Empty Rows ({empty_row_count})", width=190, height=28,
                                  fg_color="#DC2626", hover_color="#991b1b", font=("Segoe UI", 12),
                                  command=lambda f=fname, s=sheet_name: self._remove_empty_rows(
                                      f, s, return_to=self.show_consolidate_preview)
                                  ).pack(side="right")
                else:
                    ctk.CTkLabel(top_row, text="\u2714 No empty rows", font=("Segoe UI", 11, "italic"),
                                 text_color="#16A34A").pack(side="right")

                # ---- the actual Excel-style data grid, read-only ---- #
                grid = ZoomableTable(sheet_frame, height=8)
                grid.pack(fill="x", padx=10, pady=(0, 6))
                grid.set_data(sheet["columns"], sheet["data"], max_rows=300)

                # ---- editable ONLY for flagged/problem headings ---- #
                missing_cols = [c for c in sheet["columns"] if NO_HEADER_PATTERN.search(str(c))]
                if missing_cols:
                    ctk.CTkLabel(sheet_frame, text="\u26A0 Fix the flagged column heading(s) below, then Save:",
                                 font=("Segoe UI", 11, "italic"), text_color="#B91C1C").pack(anchor="w", padx=10)

                headers_area = ctk.CTkFrame(sheet_frame, fg_color="transparent")
                headers_area.pack(fill="x", padx=10, pady=(4, 4))

                entries = []
                for col in sheet["columns"]:
                    is_missing = bool(NO_HEADER_PATTERN.search(str(col)))

                    entry_row = ctk.CTkFrame(headers_area, fg_color="transparent")
                    entry_row.pack(fill="x", pady=2)

                    entry = ctk.CTkEntry(
                        entry_row, height=30, font=("Segoe UI", 12),
                        fg_color="#FEF2F2" if is_missing else None,
                        border_color="#DC2626" if is_missing else None,
                        border_width=2 if is_missing else None,
                        text_color="#991B1B" if is_missing else None,
                    )
                    entry.pack(side="left", fill="x", expand=True)
                    entry.insert(0, col)  # must insert BEFORE disabling -- disabled entries reject .insert()
                    if not is_missing:
                        entry.configure(state="disabled")  # read-only: not a flagged/problem cell
                    entries.append(entry)

                    if is_missing:
                        ctk.CTkButton(
                            entry_row, text="\U0001F5D1", width=32, height=30,
                            fg_color="#DC2626", hover_color="#991b1b", font=("Segoe UI", 13),
                            command=lambda f=fname, s=sheet_name, c=col: self._delete_column(
                                f, s, c, return_to=self.show_consolidate_preview)
                        ).pack(side="left", padx=(6, 0))
                self._review_header_entries[(fname, sheet_name)] = entries

                if missing_cols:
                    ctk.CTkButton(sheet_frame, text="\U0001F4BE Save Headings", width=150, height=30,
                                  fg_color="#2563EB", hover_color="#1d4ed8", font=("Segoe UI", 12),
                                  command=lambda f=fname, s=sheet_name: self._save_headers(
                                      f, s, return_to=self.show_consolidate_preview)
                                  ).pack(anchor="e", padx=10, pady=(2, 10))

        action_row = ctk.CTkFrame(self.footer_bar, fg_color="white")
        action_row.pack(fill="x", padx=22, pady=12)
        ctk.CTkLabel(
            action_row,
            text="Data looks right? Continue to pick which files/sheets to consolidate.",
            font=("Segoe UI", 12), text_color="#64748B"
        ).pack(side="left")
        ctk.CTkButton(action_row, text="\u25C0 Back to Dashboard", height=36,
                      fg_color="#94A3B8", hover_color="#64748B",
                      command=self.show_home).pack(side="right", padx=(6, 0))
        ctk.CTkButton(action_row, text="\U0001F4CA Continue \u2192 Choose Files & Sheets", height=36,
                      fg_color="#0F766E", hover_color="#0b5c54",
                      command=self._proceed_to_consolidation_choice).pack(side="right", padx=(6, 0))

        if preserve_scroll and scroll_pos is not None:
            self.root.after(50, lambda: self.workspace._parent_canvas.yview_moveto(scroll_pos))
        else:
            self.root.after(50, self._scroll_workspace_to_top)

    def _proceed_to_consolidation_choice(self):
        """The original Consolidate flow: pick file/sheet, then merge.
        Now reached only AFTER the data preview screen above."""
        self._set_active("consolidate")
        all_file_names = [r["file_name"] for r in self.loaded_results]
        preselected = [r["file_name"] for r in self._get_selected_results()] or all_file_names
        chosen = dialogs.ask_consolidation_choice(self.root, self.loaded_results, preselected=preselected)

        if chosen is None:
            # User clicked Cancel/Back -- go back to the preview, not all the way to Dashboard.
            self.show_consolidate_preview()
            return
        if not chosen:
            dialogs.show_warning(self.root, "No Files Chosen", "Select at least one file/sheet to consolidate.")
            self.show_consolidate_preview()
            return

        # Keep the Dashboard's tick list in sync with what was just chosen.
        for fname, var in self.file_check_vars.items():
            var.set(fname in chosen)

        sections = dialogs.ask_report_sections(self.root, preselected=self.report_sections)
        self.report_sections = sections if sections is not None else (self.report_sections or dict(DEFAULT_REPORT_SECTIONS))
        self._apply_custom_commands(self.report_sections)

        # chosen is {file_name: [sheet_name, ...]} -- build a filtered
        # copy of loaded_results that only includes the ticked sheets,
        # without touching self.loaded_results itself.
        selected = []
        for r in self.loaded_results:
            if r["file_name"] not in chosen:
                continue
            wanted_sheets = set(chosen[r["file_name"]])
            filtered = dict(r)
            filtered["sheets"] = {s: data for s, data in r["sheets"].items() if s in wanted_sheets}
            selected.append(filtered)

        merged = merge_files(selected, company_names=self.company_names)
        self._pending_consolidation_selection = {r["file_name"]: list(r["sheets"].keys()) for r in selected}

        # ---- Debit / Credit checkpoint -- hard gate, checked BEFORE
        # committing this consolidation, so a mismatch can never sneak
        # through to the results screen or export. ---- #
        # Skips silently (returns None) when the data has no debit/credit
        # columns at all -- only shown when there's something to check.
        file_paths = {r["file_name"]: r.get("file_path") for r in self.loaded_results}
        credit_debit_summary = credit_debit_checkpoint(merged["rows"], merged["columns"], file_paths=file_paths)
        if (credit_debit_summary and not credit_debit_summary["matched"]
                and self.report_sections.get("checkpoint", True)):
            self.credit_debit_summary = credit_debit_summary
            dialogs.show_credit_debit_checkpoint(
                self.root, credit_debit_summary,
                on_open_file=self._open_checkpoint_reference,
                on_refresh=self._refresh_credit_debit_checkpoint,
            )
            self.show_review_data()
            return

        self._finalize_consolidation(merged, selected, credit_debit_summary)

    def _finalize_consolidation(self, merged, selected, credit_debit_summary):
        """Commits a merge that has already cleared the Debit/Credit
        checkpoint (or never needed one) and renders the Consolidation
        Results screen. Pulled out of _proceed_to_consolidation_choice
        so _refresh_credit_debit_checkpoint can call this directly too
        -- once 'Refresh' shows the totals now tie out, the
        consolidation finishes right there instead of making the user
        close the popup and repeat the whole file/sheet-picker and
        report-sections flow from scratch."""
        self.store.set_consolidated(merged)
        self.consolidated = merged
        self.consolidated_at = datetime.now()
        self.source_files_used = [r["file_name"] for r in selected]
        self.filtered_rows = None  # reset any previous filter when re-consolidating

        self.gst_summary = gst_rate_summary(merged["rows"], merged["columns"])
        self.party_summary = party_wise_summary(merged["rows"], merged["columns"])
        self.company_summary = company_wise_summary(merged["rows"], merged["columns"])
        self.month_summary = month_wise_summary(merged["rows"], merged["columns"])
        self.credit_debit_summary = credit_debit_summary

        self._clear_workspace()
        self._screen_heading("Consolidation Results",
                              f"Merged from {len(selected)} of {len(self.loaded_results)} ticked file(s) "
                              f"on {self.consolidated_at.strftime('%d-%b-%Y %I:%M %p')}.")

        # ---- headline numbers ---- #
        totals = grand_totals(merged["rows"], merged["columns"])
        totals = {c: v for c, v in totals.items() if not NO_HEADER_PATTERN.search(str(c))}
        cards_row = ctk.CTkFrame(self.workspace, fg_color="transparent")
        cards_row.pack(fill="x", pady=(0, 14))
        StatCard(cards_row, "Total Rows", len(merged["rows"]), color="#0F766E", width=220, height=85).pack(side="left", padx=(0, 12))
        if self.credit_debit_summary and self.credit_debit_summary.get("total_debit") is not None:
            StatCard(cards_row, "Total Debit", f"{self.credit_debit_summary['total_debit']:,.2f}",
                      color="#1E3A8A", width=220, height=85).pack(side="left", padx=(0, 12))
            StatCard(cards_row, "Total Credit", f"{self.credit_debit_summary['total_credit']:,.2f}",
                      color="#1E3A8A", width=220, height=85).pack(side="left", padx=(0, 12))
        else:
            for col, total in list(totals.items())[:2]:
                StatCard(cards_row, f"Total {col}", f"{total:,.2f}", color="#1E3A8A", width=220, height=85).pack(side="left", padx=(0, 12))

        if self.gst_summary and "error" not in self.gst_summary and self.gst_summary.get("tax_mismatches"):
            ctk.CTkLabel(self.workspace, text=f"\u26A0 {len(self.gst_summary['tax_mismatches'])} tax mismatch(es) detected \u2014 check the Issues report.",
                         font=("Segoe UI", 13, "bold"), text_color="#B91C1C").pack(anchor="w", pady=(0, 8))

        if self.credit_debit_summary and self.report_sections.get("checkpoint", True):
            if self.credit_debit_summary["matched"]:
                ctk.CTkLabel(self.workspace, text="\u2714 Debit / Credit Checkpoint: totals tie out.",
                             font=("Segoe UI", 13, "bold"), text_color="#16A34A").pack(anchor="w", pady=(0, 8))
            else:
                cd_row = ctk.CTkFrame(self.workspace, fg_color="transparent")
                cd_row.pack(fill="x", pady=(0, 8))
                ctk.CTkLabel(cd_row, text=f"\u26A0 Debit / Credit Checkpoint: totals don't match "
                                          f"(difference {abs(self.credit_debit_summary['difference']):,.2f}).",
                             font=("Segoe UI", 13, "bold"), text_color="#B91C1C").pack(side="left")
                ctk.CTkButton(cd_row, text="View Details", width=110, height=26, fg_color="#F59E0B",
                              hover_color="#b45309", font=("Segoe UI", 11),
                              command=self._show_credit_debit_checkpoint).pack(side="left", padx=(10, 0))

        executed_custom = self.report_sections.get("_executed_custom") or []
        if executed_custom:
            applied_text = "; ".join(f"'{item}'" for item in executed_custom)
            ctk.CTkLabel(self.workspace, text=f"\u2714 Applied your custom request(s): {applied_text}",
                         font=("Segoe UI", 12, "bold"), text_color="#16A34A",
                         anchor="w", justify="left", wraplength=900).pack(anchor="w", pady=(0, 8))

        if self.report_sections.get("emphasize_grand_total"):
            gt_card = ctk.CTkFrame(self.workspace, fg_color="#EFF6FF", corner_radius=8, border_width=1, border_color="#93C5FD")
            gt_card.pack(fill="x", pady=(0, 10))
            ctk.CTkLabel(gt_card, text="\U0001F522 Grand Total Recap (as requested)", font=("Segoe UI", 13, "bold"),
                         text_color="#1E3A8A", anchor="w").pack(anchor="w", padx=12, pady=(8, 2))
            gt_text = f"Total Rows: {len(merged['rows']):,}   |   " + "   |   ".join(
                f"Total {col}: {total:,.2f}" for col, total in totals.items()
            )
            ctk.CTkLabel(gt_card, text=gt_text, font=("Segoe UI", 12), text_color="#334155",
                         anchor="w", justify="left", wraplength=880).pack(anchor="w", padx=12, pady=(0, 10))

        ctk.CTkLabel(self.workspace, text="Merged Data (Grand Total row at the bottom)",
                     font=("Segoe UI", 15, "bold"), text_color="#1E3A8A").pack(anchor="w", pady=(6, 4))

        table = ZoomableTable(self.workspace, height=16)
        table.pack(fill="both", expand=True, pady=(0, 12))
        self._load_table_with_total(table, merged["columns"], merged["rows"])

        # ---- GST / Party summaries as text (only the sections the user asked for) ---- #
        summary_lines = []
        if self.report_sections.get("gst", False):
            summary_lines.append("GST RATE-WISE SUMMARY")
            if self.gst_summary and "error" not in self.gst_summary:
                for rate, vals in sorted(self.gst_summary["rate_wise"].items()):
                    summary_lines.append(f"  {rate}% -> Taxable: {vals['taxable_total']:,.2f}, Tax: {vals['tax_total']:,.2f}, Rows: {vals['row_count']}")
                summary_lines.append(f"  CGST total: {self.gst_summary['cgst_total']:,.2f}  |  SGST total: {self.gst_summary['sgst_total']:,.2f}  |  IGST total: {self.gst_summary['igst_total']:,.2f}")
            else:
                summary_lines.append(f"  {self.gst_summary.get('error', 'n/a')}")
            summary_lines.append("")

        if self.report_sections.get("company", True):
            summary_lines.append("COMPANY-WISE SUMMARY")
            if self.company_summary and self.company_summary.get("companies"):
                for name, vals in self.company_summary["companies"].items():
                    summary_lines.append(f"  {name}: Taxable {vals['taxable_total']:,.2f}, Tax {vals['tax_total']:,.2f}, Rows {vals['row_count']}")
            else:
                summary_lines.append("  (no company data)")
            summary_lines.append("")

        if self.report_sections.get("party", False):
            summary_lines.append("PARTY-WISE SUMMARY")
            if self.party_summary and "error" not in self.party_summary:
                for name, vals in self.party_summary["parties"].items():
                    summary_lines.append(f"  {name}: Taxable {vals['taxable_total']:,.2f}, Tax {vals['tax_total']:,.2f}, Invoices {vals['invoice_count']}")
            else:
                summary_lines.append(f"  {self.party_summary.get('error', 'n/a')}")
            summary_lines.append("")

        custom_notes = self.report_sections.get("custom") or []
        if custom_notes:
            summary_lines.append("CUSTOM ITEMS REQUESTED (not auto-built -- follow up manually)")
            for item in custom_notes:
                summary_lines.append(f"  - {item}")

        if not summary_lines:
            summary_lines.append("(No optional sections selected -- see the Merged Data table above.)")

        box = ZoomableTextbox(self.workspace, height=180)
        box.pack(fill="both", pady=(0, 10))
        box.set_text("\n".join(summary_lines))

        results_actions = ctk.CTkFrame(self.footer_bar, fg_color="white")
        results_actions.pack(fill="x", padx=22, pady=12)
        ctk.CTkButton(results_actions, text="\U0001F50E Go to Filters & Search \u2192", height=38,
                      command=self.show_filters).pack(side="left")
        ctk.CTkButton(results_actions, text="\U0001F4E4 Export Consolidated Data \u2192", height=38,
                      fg_color="#0F766E", hover_color="#0b5c54",
                      command=self.export_data).pack(side="left", padx=(10, 0))

        self.root.after(50, self._scroll_workspace_to_top)

    def _load_table_with_total(self, table_widget, columns, rows):
        """Populates a ZoomableTable and appends a synthetic GRAND TOTAL row for display.
        Numeric columns are shown in Indian numbering style (e.g. 12,34,567,
        no decimals) -- the same format now used in the exported Excel
        file -- so what's shown on screen matches what you'll see when
        you open the report."""
        display_cols = [c for c in columns if not str(c).startswith("__")]
        totals = grand_totals(rows, display_cols)
        num_cols = set(numeric_columns(rows, display_cols))

        def _fmt(col, value):
            if col in num_cols and value not in ("", None):
                return format_indian_number(value)
            return value

        total_row = {c: "" for c in display_cols}
        label_set = False
        for c in display_cols:
            if c in totals:
                total_row[c] = format_indian_number(totals[c])
            elif not label_set:
                total_row[c] = "GRAND TOTAL"
                label_set = True
        display_rows = [{c: _fmt(c, r.get(c, "")) for c in display_cols} for r in rows] + ([total_row] if rows else [])
        table_widget.set_data(display_cols, display_rows)

    def _open_checkpoint_reference(self, mismatch):
        """Opens a Debit/Credit Checkpoint reference's source file with
        its Debit and Credit columns highlighted in yellow, so the
        mismatch is immediately visible in Excel."""
        success, reason = exporter.open_file_with_highlighted_columns(
            mismatch.get("file_path"), mismatch.get("sheet_name"),
            [mismatch.get("debit_col"), mismatch.get("credit_col")]
        )
        if not success:
            dialogs.show_error(self.root, "Couldn't Open File", reason or "An unknown error occurred.")

    def _refresh_credit_debit_checkpoint(self):
        """Re-imports every file involved in the last consolidation
        attempt straight from disk, recomputes the Debit/Credit
        checkpoint, and re-opens the popup with the up-to-date numbers
        -- the 'Refresh' button on that popup.

        Clicking 'Open File' opens a highlighted COPY of the source
        file (so the original on disk is never touched automatically).
        That means the fix the user just made in Excel usually lives in
        that copy, not in the original file -- so for each sheet
        involved, this checks for a saved highlighted copy first and
        reads the corrected numbers from there if one exists; only
        falls back to the original file if no such copy is present.
        The file's recorded file_path is kept as the original either
        way, so 'Open File', exports, and everything else downstream
        keep pointing at the real source file, not the copy."""
        selection = self._pending_consolidation_selection
        if not selection:
            return

        for r in self.loaded_results:
            fname = r.get("file_name")
            if fname in selection and r.get("file_path"):
                source_path = r["file_path"]

                # Prefer a saved highlighted copy (the file "Open File"
                # opened and the user may have edited + saved) over the
                # untouched original, for any of this file's sheets that
                # are part of the pending consolidation.
                for sheet_name in selection.get(fname, []):
                    copy_path = exporter.highlighted_copy_path(source_path, sheet_name)
                    if copy_path and os.path.exists(copy_path):
                        source_path = copy_path
                        break

                fresh = self.engine.load_excel(source_path)
                if "error" not in fresh:
                    fresh["file_path"] = r["file_path"]  # keep the ORIGINAL path recorded
                    fresh["source_folder"] = r.get("source_folder")
                    idx = next((i for i, x in enumerate(self.loaded_results) if x["file_name"] == fname), None)
                    if idx is not None:
                        self.loaded_results[idx] = fresh

        selected = []
        for r in self.loaded_results:
            wanted_sheets = set(selection.get(r["file_name"], []))
            if not wanted_sheets:
                continue
            filtered = dict(r)
            filtered["sheets"] = {s: data for s, data in r["sheets"].items() if s in wanted_sheets}
            selected.append(filtered)

        merged = merge_files(selected, company_names=self.company_names)
        file_paths = {r["file_name"]: r.get("file_path") for r in self.loaded_results}
        credit_debit_summary = credit_debit_checkpoint(merged["rows"], merged["columns"], file_paths=file_paths)
        self.credit_debit_summary = credit_debit_summary

        if credit_debit_summary and not credit_debit_summary["matched"]:
            dialogs.show_credit_debit_checkpoint(
                self.root, credit_debit_summary,
                on_open_file=self._open_checkpoint_reference,
                on_refresh=self._refresh_credit_debit_checkpoint,
            )
        else:
            dialogs.show_info(
                self.root, "Checkpoint Passed",
                "Debit and Credit totals now match. Finishing the consolidation now..."
            )
            self._finalize_consolidation(merged, selected, credit_debit_summary)

    def _show_credit_debit_checkpoint(self):
        """Re-opens the Debit/Credit Checkpoint popup on demand (e.g. via
        the 'View Details' link on the Consolidation Results screen)."""
        if self.credit_debit_summary:
            dialogs.show_credit_debit_checkpoint(
                self.root, self.credit_debit_summary,
                on_open_file=self._open_checkpoint_reference,
                on_refresh=self._refresh_credit_debit_checkpoint,
            )

    # ================= FILTERS & SEARCH ================= #

    def show_filters(self):
        self._set_active("filters")
        if not self.consolidated:
            dialogs.show_warning(self.root, "No Data", "Please consolidate data first (Import Files \u2192 Consolidate).")
            self.show_home()
            return

        self._record_nav(self.show_filters)
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
        results_header = ctk.CTkFrame(self.workspace, fg_color="transparent")
        results_header.pack(fill="x")
        self.filters_result_label = ctk.CTkLabel(results_header, text="", font=("Segoe UI", 13, "italic"), text_color="#64748B")
        self.filters_result_label.pack(side="left", anchor="w")

        row_actions = ctk.CTkFrame(results_header, fg_color="transparent")
        row_actions.pack(side="right")
        ctk.CTkButton(row_actions, text="\u2795 Add Row", width=120, height=30,
                      fg_color="#16A34A", hover_color="#15803d",
                      command=self._add_row_action).pack(side="left", padx=(0, 6))
        ctk.CTkButton(row_actions, text="\U0001F5D1 Remove Selected Row(s)", width=190, height=30,
                      fg_color="#B91C1C", hover_color="#7f1d1d",
                      command=self._remove_selected_rows_action).pack(side="left")

        self.filters_table = ZoomableTable(self.workspace, height=16)
        self.filters_table.pack(fill="both", expand=True, pady=(4, 10))

        ctk.CTkButton(self.workspace, text="\U0001F4E4 Export This Filtered View", height=38,
                      command=self._export_filtered_view).pack(anchor="w")

        self._refresh_filters_table(columns)

    def _add_row_action(self):
        columns = [c for c in self.consolidated["columns"] if not str(c).startswith("__")]
        new_row = dialogs.ask_new_row(self.root, columns)
        if new_row is None:
            return
        # Goes straight into the real dataset (self.consolidated), so it
        # carries through to Analytics and every export -- not just a
        # cosmetic addition to this screen's table.
        self.consolidated["rows"].append(new_row)
        if self.filtered_rows is not None:
            self.filtered_rows.append(new_row)
        self._refresh_filters_table(columns)

    def _remove_selected_rows_action(self):
        selected = self.filters_table.get_selected_rows()
        if not selected:
            dialogs.show_warning(self.root, "No Rows Selected",
                                  "Click a row (Ctrl/Shift+click for more) in the table, then try again.")
            return
        confirmed = dialogs.ask_yes_no(
            self.root, "Remove Row(s)?",
            f"Remove {len(selected)} row(s) from the dataset? This affects Analytics and Export too, "
            "not just this filtered view."
        )
        if not confirmed:
            return
        # Match by identity, not equality -- two genuinely different
        # rows can have identical values (e.g. two $0.00 lines), and we
        # only ever want to remove the exact rows the user selected.
        selected_ids = {id(r) for r in selected}
        self.consolidated["rows"] = [r for r in self.consolidated["rows"] if id(r) not in selected_ids]
        if self.filtered_rows is not None:
            self.filtered_rows = [r for r in self.filtered_rows if id(r) not in selected_ids]
        columns = [c for c in self.consolidated["columns"] if not str(c).startswith("__")]
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

        self._record_nav(self.show_analytics)
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

        self._record_nav(self.show_pivot)
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
        self._record_nav(self.show_settings)
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

        # Always rebuild "Issues Found" right now, from only the files that
        # are CURRENTLY loaded and ticked -- never reuse self.all_issues as
        # it may have been computed earlier for a different set of files
        # (e.g. before an old file was removed/deselected and a new one
        # imported), which would otherwise leak stale rows into the export.
        self.all_issues = self._compute_all_issues(self._get_consolidated_scope_results())

        file_paths = {r["file_name"]: r.get("file_path") for r in self.loaded_results}
        self._run_export(
            lambda path: exporter.export_full_report(
                path, self.consolidated, gst_summary=self.gst_summary, party_summary=self.party_summary,
                validation_issues=self.all_issues, source_files=self.source_files_used,
                company_summary=self.company_summary, month_summary=self.month_summary,
                consolidated_at=self.consolidated_at, sections=self.report_sections, file_paths=file_paths
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

        # The success screen above already has an "Open File in Excel"
        # button, so asking "Open it now?" in a second stacked popup was
        # just a redundant, confusing extra step. "Auto-open" now means
        # exactly that -- it opens automatically, no second prompt.
        if self.settings.get("auto_open_after_export", True):
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