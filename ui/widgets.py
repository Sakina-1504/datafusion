"""
widgets.py

Reusable UI building blocks shared across dashboard screens:

  ZoomableTable   - a data grid (ttk.Treeview) with +/- zoom buttons
                    and Ctrl+MouseWheel zoom, so nothing is ever too
                    small to read on a big consolidated sheet.
  ZoomableTextbox - same zoom behaviour for the plain-text report
                    panels (Validate results, etc).
  StatCard        - small coloured summary card used on the Dashboard
                    home screen (Files Imported, Rows, etc).
"""

import re
import tkinter as tk
from tkinter import ttk
import customtkinter as ctk

from backend.filters import numeric_columns

# Column names containing any of these are treated as "this column
# holds a real number" as a quick fallback when there isn't enough
# data to sample -- numeric_columns() (data-driven) is tried first.
_VALUE_COL_KEYWORDS = ("debit", "credit", "amount", "balance", "qty", "quantity", "rate", "total")
_TOTAL_TEXT_RE = re.compile(r"\btotal(s)?\b", re.IGNORECASE)


def _is_special_row(row, columns, value_cols):
    """True for a row that's structure, not a data line: either it
    already says TOTAL/Sub Total somewhere, or every value column is
    blank while some other cell has real text (a bare section header
    like "ASSETS" or "Current Assets")."""
    values = [row.get(c, "") for c in columns]
    if any(_TOTAL_TEXT_RE.search(str(v)) for v in values):
        return True
    if not value_cols:
        return False
    all_values_blank = all(str(row.get(c, "")).strip() == "" for c in value_cols)
    has_other_text = any(str(row.get(c, "")).strip() != "" for c in columns if c not in value_cols)
    return all_values_blank and has_other_text


class StatCard(ctk.CTkFrame):
    def __init__(self, parent, title, value, color="#1E3A8A", **kwargs):
        super().__init__(parent, corner_radius=10, fg_color="white", border_width=1,
                          border_color="#E2E8F0", **kwargs)
        self.value_label = ctk.CTkLabel(self, text=str(value), font=("Segoe UI", 26, "bold"), text_color=color)
        self.value_label.pack(pady=(14, 0), padx=16)
        ctk.CTkLabel(self, text=title, font=("Segoe UI", 12), text_color="#64748B").pack(pady=(0, 12), padx=16)

    def set_value(self, value, color=None):
        self.value_label.configure(text=str(value))
        if color:
            self.value_label.configure(text_color=color)


class EditableDataGrid(ctk.CTkFrame):
    """
    A Treeview-based grid for the Review Data page: shows the ACTUAL
    rows of a sheet (not just column headers), and lets the user
    double-click any cell to edit its value directly -- Excel-style,
    right there before consolidating. Edits write straight back into
    the underlying row dict passed to set_data(), so there's no
    separate "save" step for cell edits (unlike the column-heading
    editor, which needs an explicit Save).

    mismatch=True tints every row in this grid light red -- used when
    this sheet is one of the Debit/Credit Checkpoint's mismatched
    references, so the user can see exactly what to fix and edit it
    right here instead of hunting through the popup.
    """

    def __init__(self, parent, height=8, **kwargs):
        super().__init__(parent, fg_color="white", corner_radius=8, **kwargs)
        self._columns = []
        self._rows = []
        self._edit_entry = None

        tree_frame = tk.Frame(self, bg="white")
        tree_frame.pack(fill="both", expand=True, padx=4, pady=4)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        style_name = f"Edit{id(self)}.Treeview"
        style = ttk.Style()
        style.configure(style_name, rowheight=24, font=("Segoe UI", 10))
        style.configure(f"{style_name}.Heading", font=("Segoe UI", 10, "bold"))

        self.tree = ttk.Treeview(tree_frame, style=style_name, yscrollcommand=vsb.set,
                                  xscrollcommand=hsb.set, height=height)
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)

        self.tree.tag_configure("mismatch", background="#FEE2E2")
        self.tree.tag_configure("empty_row", background="#FEF9C3")

        self.tree.bind("<Double-1>", self._on_double_click)
        # scrolling/resizing can leave a stray edit box floating in
        # the wrong place -- close it rather than let that happen.
        self.tree.bind("<MouseWheel>", lambda e: self._destroy_editor())
        vsb.bind("<B1-Motion>", lambda e: self._destroy_editor())

    def set_data(self, columns, rows, mismatch=False, max_rows=1000):
        """rows: list[dict] -- the SAME objects the caller holds, so
        edits made here are immediately visible to the caller too."""
        self._destroy_editor()
        self.tree.delete(*self.tree.get_children())
        self._columns = list(columns)
        self._rows = rows
        self.tree["columns"] = self._columns
        self.tree["show"] = "headings"
        for col in self._columns:
            sample_lengths = [len(str(col))] + [len(str(r.get(col, ""))) for r in rows[:200]]
            width = min(max(70, max(sample_lengths) * 8), 240)
            self.tree.heading(col, text=str(col))
            self.tree.column(col, width=width, anchor="w")

        shown = rows[:max_rows]
        for row in shown:
            values = [row.get(c, "") for c in self._columns]
            is_empty = all(str(v).strip() == "" for v in values)
            if mismatch:
                tags = ("mismatch",)
            elif is_empty:
                tags = ("empty_row",)
            else:
                tags = ()
            self.tree.insert("", "end", values=values, tags=tags)

    def _destroy_editor(self):
        if self._edit_entry is not None:
            try:
                self._edit_entry.destroy()
            except Exception:
                pass
            self._edit_entry = None

    def _on_double_click(self, event):
        self._destroy_editor()
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return
        try:
            col_index = int(col_id.replace("#", "")) - 1
        except ValueError:
            return
        if col_index < 0 or col_index >= len(self._columns):
            return
        bbox = self.tree.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, width, height = bbox
        col_name = self._columns[col_index]
        current_value = self.tree.set(row_id, col_name)

        entry = tk.Entry(self.tree, font=("Segoe UI", 10))
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, current_value)
        entry.select_range(0, "end")
        entry.focus()
        self._edit_entry = entry

        def _commit(event=None):
            new_value = entry.get()
            self.tree.set(row_id, col_name, new_value)
            idx = self.tree.index(row_id)
            if 0 <= idx < len(self._rows):
                self._rows[idx][col_name] = new_value
            self._destroy_editor()

        def _cancel(event=None):
            self._destroy_editor()

        entry.bind("<Return>", _commit)
        entry.bind("<FocusOut>", _commit)
        entry.bind("<Escape>", _cancel)


class ZoomableTable(ctk.CTkFrame):
    """A Treeview-based data grid with zoom controls and a status bar.
    Use set_data(columns, rows) to (re)populate it."""

    def __init__(self, parent, height=18, **kwargs):
        super().__init__(parent, fg_color="white", corner_radius=8, **kwargs)

        self._zoom = 100
        self._base_font_size = 10
        self._style_name = f"Zoom{id(self)}.Treeview"

        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.pack(fill="x", padx=6, pady=(6, 0))

        self.status_label = ctk.CTkLabel(toolbar, text="No data loaded", font=("Segoe UI", 11), text_color="#64748B")
        self.status_label.pack(side="left", padx=(4, 0))

        zoom_box = ctk.CTkFrame(toolbar, fg_color="transparent")
        zoom_box.pack(side="right")

        ctk.CTkButton(zoom_box, text="-", width=28, height=26, command=self.zoom_out).pack(side="left", padx=2)
        self.zoom_label = ctk.CTkLabel(zoom_box, text="100%", width=45, font=("Segoe UI", 11))
        self.zoom_label.pack(side="left", padx=2)
        ctk.CTkButton(zoom_box, text="+", width=28, height=26, command=self.zoom_in).pack(side="left", padx=2)
        ctk.CTkButton(zoom_box, text="Reset", width=55, height=26, command=self.zoom_reset).pack(side="left", padx=(6, 2))

        tree_frame = tk.Frame(self, bg="white")
        tree_frame.pack(fill="both", expand=True, padx=6, pady=6)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")

        # 'clam' is the only built-in ttk theme that reliably honours
        # custom Treeview colors on Windows -- the default theme there
        # ignores background/foreground options on headings, so without
        # this the header stays the OS's plain gray regardless of what
        # we configure below.
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass
        self.style.configure(self._style_name, rowheight=24, font=("Segoe UI", self._base_font_size),
                              background="white", fieldbackground="white", borderwidth=0)
        self.style.configure(f"{self._style_name}.Heading", font=("Segoe UI", self._base_font_size, "bold"),
                              background="#1E3A8A", foreground="white", relief="flat")
        self.style.map(f"{self._style_name}.Heading", background=[("active", "#1E3A8A")])

        self.tree = ttk.Treeview(tree_frame, style=self._style_name, selectmode="extended",
                                  yscrollcommand=vsb.set, xscrollcommand=hsb.set, height=height)
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)

        # Ctrl + mouse wheel to zoom (Windows/Linux use <MouseWheel>/<Button-4/5>)
        self.tree.bind("<Control-MouseWheel>", self._on_ctrl_wheel)
        self.tree.bind("<Control-Button-4>", lambda e: self.zoom_in())
        self.tree.bind("<Control-Button-5>", lambda e: self.zoom_out())

        # Plain (non-Ctrl) mouse wheel: scroll the table itself while
        # there are more rows to show, then hand scrolling off to the
        # page the table lives inside once the table hits its own top
        # or bottom -- otherwise a table full of rows swallows every
        # wheel event and the rest of the page becomes unreachable.
        self.tree.bind("<MouseWheel>", self._on_wheel)
        self.tree.bind("<Button-4>", self._on_wheel)
        self.tree.bind("<Button-5>", self._on_wheel)

        self._columns = []
        self._rows = []
        self._col_widths = {}

    def _on_wheel(self, event):
        delta = getattr(event, "delta", 0)
        if delta:
            direction = -1 if delta > 0 else 1
        elif getattr(event, "num", None) == 4:
            direction = -1
        elif getattr(event, "num", None) == 5:
            direction = 1
        else:
            return None

        first, last = self.tree.yview()
        at_top = first <= 0.0001
        at_bottom = last >= 0.9999

        if (direction < 0 and at_top) or (direction > 0 and at_bottom):
            outer_canvas = self._find_outer_canvas()
            if outer_canvas is not None:
                outer_canvas.yview_scroll(direction, "units")
            return "break"

        self.tree.yview_scroll(direction, "units")
        return "break"

    def _find_outer_canvas(self):
        """Walks up the widget tree to find the CTkScrollableFrame this
        table has been placed inside (it exposes a `_parent_canvas`)."""
        widget = self.master
        while widget is not None:
            if hasattr(widget, "_parent_canvas"):
                return widget._parent_canvas
            widget = widget.master
        return None

    def _on_ctrl_wheel(self, event):
        if event.delta > 0:
            self.zoom_in()
        else:
            self.zoom_out()

    def zoom_in(self):
        self._apply_zoom(min(self._zoom + 10, 200))

    def zoom_out(self):
        self._apply_zoom(max(self._zoom - 10, 60))

    def zoom_reset(self):
        self._apply_zoom(100)

    def _apply_zoom(self, level):
        self._zoom = level
        self.zoom_label.configure(text=f"{level}%")
        size = max(7, int(self._base_font_size * level / 100))
        row_h = max(18, int(24 * level / 100))
        self.style.configure(self._style_name, rowheight=row_h, font=("Segoe UI", size))
        self.style.configure(f"{self._style_name}.Heading", font=("Segoe UI", size, "bold"))
        # widen columns proportionally so zoomed text isn't clipped
        for col in self._columns:
            base_w = self._col_widths.get(col, 120)
            self.tree.column(col, width=int(base_w * level / 100))

    def set_data(self, columns, rows, max_rows=2000):
        """columns: list[str], rows: list[dict]. Caps at max_rows for
        UI responsiveness (a settings-controlled limit)."""
        self.tree.delete(*self.tree.get_children())
        self._columns = list(columns)
        self._rows = rows
        self.tree["columns"] = self._columns
        self.tree["show"] = "headings"
        # Plain two-tone banding (white / very light blue) instead of a
        # flat, undifferentiated grid -- makes wide rows easier to
        # follow without adding extra colors.
        self.tree.tag_configure("evenrow", background="white")
        self.tree.tag_configure("oddrow", background="#EAF1FB")
        self.tree.tag_configure("evenrow_total", background="white",
                                 foreground="#1E3A8A", font=("Segoe UI", self._base_font_size, "bold"))
        self.tree.tag_configure("oddrow_total", background="#EAF1FB",
                                 foreground="#1E3A8A", font=("Segoe UI", self._base_font_size, "bold"))

        self._col_widths = {}
        for col in self._columns:
            sample_lengths = [len(str(col))] + [len(str(r.get(col, ""))) for r in rows[:200]]
            width = min(max(60, max(sample_lengths) * 8), 260)
            self._col_widths[col] = width
            self.tree.heading(col, text=str(col))
            self.tree.column(col, width=width, anchor="w")

        shown = rows[:max_rows]
        # Value columns are detected from the actual data (numeric_columns),
        # so an amount column named after a date (e.g. a Balance Sheet's
        # "Mar 31, 25") is still recognized correctly -- falls back to a
        # simple name-keyword guess only if that finds nothing at all.
        value_cols = numeric_columns(rows, self._columns, sample_size=min(len(rows), 300))
        if not value_cols:
            value_cols = [c for c in self._columns if any(
                kw in str(c).lower() for kw in ("debit", "credit", "amount", "balance", "qty", "quantity", "rate", "total")
            )]
        for i, row in enumerate(shown):
            values = [row.get(c, "") for c in self._columns]
            # A "TOTAL"/"Sub Total" line, or a bare section-header row
            # (e.g. "ASSETS", "Current Assets" -- label with no figure
            # of its own) gets blue, bold text so the hierarchy reads
            # as structure instead of blending into the data rows.
            is_special = _is_special_row(row, self._columns, value_cols)
            even = (i % 2 == 0)
            band_tag = ("evenrow_total" if even else "oddrow_total") if is_special else ("evenrow" if even else "oddrow")
            self.tree.insert("", "end", values=values, tags=(band_tag,))

        truncated_note = f" (showing first {max_rows:,})" if len(rows) > max_rows else ""
        self.status_label.configure(text=f"{len(rows):,} row(s){truncated_note}  |  {len(self._columns)} column(s)")
        self._apply_zoom(self._zoom)

    def get_selected_rows(self):
        """Returns the actual row dicts (same objects passed into
        set_data) that are currently selected in the table -- lets a
        caller delete exactly those rows from the real dataset, not
        just hide them from view."""
        selected = []
        for iid in self.tree.selection():
            idx = self.tree.index(iid)
            if idx < len(self._rows):
                selected.append(self._rows[idx])
        return selected

    def clear(self):
        self.tree.delete(*self.tree.get_children())
        self.status_label.configure(text="No data loaded")

    def get_selected_rows(self):
        """Returns the actual row dict(s) (from the list passed to
        set_data) corresponding to whatever's currently selected in
        the table -- used by screens that let the user delete rows."""
        selected = []
        for iid in self.tree.selection():
            idx = self.tree.index(iid)
            if 0 <= idx < len(self._rows):
                selected.append(self._rows[idx])
        return selected


class ZoomableTextbox(ctk.CTkFrame):
    """CTkTextbox with +/- zoom controls, for text-based report panels."""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)
        self._font_size = 13

        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.pack(fill="x")
        zoom_box = ctk.CTkFrame(toolbar, fg_color="transparent")
        zoom_box.pack(side="right")
        ctk.CTkButton(zoom_box, text="-", width=28, height=26, command=self.zoom_out).pack(side="left", padx=2)
        self.zoom_label = ctk.CTkLabel(zoom_box, text="13pt", width=40, font=("Segoe UI", 11))
        self.zoom_label.pack(side="left", padx=2)
        ctk.CTkButton(zoom_box, text="+", width=28, height=26, command=self.zoom_in).pack(side="left", padx=2)

        self.box = ctk.CTkTextbox(self, font=("Consolas", self._font_size), wrap="word")
        self.box.pack(fill="both", expand=True, pady=(4, 0))
        self.box.bind("<Control-MouseWheel>", lambda e: self.zoom_in() if e.delta > 0 else self.zoom_out())
        # Plain (non-Ctrl) wheel: scroll the textbox itself, then hand off
        # to the page it lives inside once the textbox hits its own top
        # or bottom -- same reasoning as ZoomableTable._on_wheel above.
        self.box.bind("<MouseWheel>", self._on_wheel)
        self.box.bind("<Button-4>", self._on_wheel)
        self.box.bind("<Button-5>", self._on_wheel)

    def _on_wheel(self, event):
        delta = getattr(event, "delta", 0)
        if delta:
            direction = -1 if delta > 0 else 1
        elif getattr(event, "num", None) == 4:
            direction = -1
        elif getattr(event, "num", None) == 5:
            direction = 1
        else:
            return None

        first, last = self.box.yview()
        at_top = first <= 0.0001
        at_bottom = last >= 0.9999

        if (direction < 0 and at_top) or (direction > 0 and at_bottom):
            outer_canvas = self._find_outer_canvas()
            if outer_canvas is not None:
                outer_canvas.yview_scroll(direction, "units")
            return "break"

        self.box.yview_scroll(direction, "units")
        return "break"

    def _find_outer_canvas(self):
        widget = self.master
        while widget is not None:
            if hasattr(widget, "_parent_canvas"):
                return widget._parent_canvas
            widget = widget.master
        return None

    def zoom_in(self):
        self._font_size = min(self._font_size + 1, 28)
        self._refresh()

    def zoom_out(self):
        self._font_size = max(self._font_size - 1, 8)
        self._refresh()

    def _refresh(self):
        self.box.configure(font=("Consolas", self._font_size))
        self.zoom_label.configure(text=f"{self._font_size}pt")

    def set_text(self, text):
        self.box.configure(state="normal")
        self.box.delete("1.0", "end")
        self.box.insert("1.0", text)
        self.box.configure(state="disabled")