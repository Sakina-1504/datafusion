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

import tkinter as tk
from tkinter import ttk
import customtkinter as ctk


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

        self.style = ttk.Style()
        try:
            self.style.theme_use(self.style.theme_use())
        except Exception:
            pass
        self.style.configure(self._style_name, rowheight=24, font=("Segoe UI", self._base_font_size))
        self.style.configure(f"{self._style_name}.Heading", font=("Segoe UI", self._base_font_size, "bold"))

        self.tree = ttk.Treeview(tree_frame, style=self._style_name,
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

        self._columns = []
        self._rows = []
        self._col_widths = {}

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

        self._col_widths = {}
        for col in self._columns:
            sample_lengths = [len(str(col))] + [len(str(r.get(col, ""))) for r in rows[:200]]
            width = min(max(60, max(sample_lengths) * 8), 260)
            self._col_widths[col] = width
            self.tree.heading(col, text=str(col))
            self.tree.column(col, width=width, anchor="w")

        shown = rows[:max_rows]
        for row in shown:
            values = [row.get(c, "") for c in self._columns]
            self.tree.insert("", "end", values=values)

        truncated_note = f" (showing first {max_rows:,})" if len(rows) > max_rows else ""
        self.status_label.configure(text=f"{len(rows):,} row(s){truncated_note}  |  {len(self._columns)} column(s)")
        self._apply_zoom(self._zoom)

    def clear(self):
        self.tree.delete(*self.tree.get_children())
        self.status_label.configure(text="No data loaded")


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
