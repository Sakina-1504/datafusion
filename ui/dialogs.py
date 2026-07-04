"""
dialogs.py

Good-looking, on-brand replacements for tkinter.messagebox.
Every popup in the app (import results, warnings, errors, export
confirmations, the "Company Name" prompt, etc.) should use these
instead of the plain grey OS message boxes.

Drop-in style usage:

    from ui import dialogs

    dialogs.show_info(self.root, "Import Complete", "Loaded 3 file(s).")
    dialogs.show_error(self.root, "Import Errors", "Demo1.xlsx: bad file")
    dialogs.show_warning(self.root, "No Data", "Please import files first.")
    if dialogs.ask_yes_no(self.root, "Export Complete", "Open it now?"):
        ...
    choice = dialogs.ask_yes_no_cancel(self.root, "Export Options", "...")
    name = dialogs.ask_text(self.root, "Company Name", "Which company...", default="Acme Pvt Ltd")
"""

import os
import tkinter as tk
import customtkinter as ctk

BRAND = "#1E3A8A"
INFO_COLOR = "#16A34A"
ERROR_COLOR = "#DC2626"
WARNING_COLOR = "#F59E0B"
QUESTION_COLOR = "#2563EB"
MUTED_TEXT = "#64748B"


class _BaseDialog(ctk.CTkToplevel):
    def __init__(self, parent, title, message, accent_color, icon_text,
                 width=440, extra_builder=None):
        super().__init__(parent)
        self.result = None
        self.title(title)
        self.geometry(f"{width}x{240 if not extra_builder else 300}")
        self.resizable(False, False)
        self.configure(fg_color="#F5F7FA")
        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

        # coloured accent bar across the top
        ctk.CTkFrame(self, height=6, corner_radius=0, fg_color=accent_color).pack(fill="x", side="top")

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=26, pady=(18, 14))

        top_row = ctk.CTkFrame(body, fg_color="transparent")
        top_row.pack(fill="x")

        badge = ctk.CTkLabel(top_row, text=icon_text, font=("Segoe UI", 22, "bold"),
                              text_color="white", fg_color=accent_color,
                              width=44, height=44, corner_radius=22)
        badge.pack(side="left", anchor="n")

        text_col = ctk.CTkFrame(top_row, fg_color="transparent")
        text_col.pack(side="left", fill="both", expand=True, padx=(14, 0))

        ctk.CTkLabel(text_col, text=title, font=("Segoe UI", 16, "bold"),
                     text_color=BRAND, anchor="w", justify="left").pack(fill="x")
        ctk.CTkLabel(text_col, text=message, font=("Segoe UI", 13),
                     text_color="#334155", anchor="w", justify="left",
                     wraplength=width - 120).pack(fill="x", pady=(6, 0))

        self.extra_area = ctk.CTkFrame(body, fg_color="transparent")
        self.extra_area.pack(fill="x", pady=(12, 0))
        if extra_builder:
            extra_builder(self.extra_area)

        self.button_row = ctk.CTkFrame(self, fg_color="transparent")
        self.button_row.pack(fill="x", padx=26, pady=(0, 18), side="bottom")

        self.after(50, self._center_on_parent)
        self.bind("<Return>", lambda e: self._on_default_enter())
        self.bind("<Escape>", lambda e: self._on_cancel())

    def _center_on_parent(self):
        self.update_idletasks()
        try:
            px, py = self.master.winfo_rootx(), self.master.winfo_rooty()
            pw, ph = self.master.winfo_width(), self.master.winfo_height()
            w, h = self.winfo_width(), self.winfo_height()
            self.geometry(f"+{px + (pw - w) // 2}+{py + (ph - h) // 2}")
        except Exception:
            pass
        self.focus_force()

    def _on_default_enter(self):
        self._on_cancel()

    def _on_cancel(self):
        self.result = None
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

    def _add_button(self, text, color, hover, value, default=False, side="right"):
        def _click():
            self.result = value
            try:
                self.grab_release()
            except Exception:
                pass
            self.destroy()
        btn = ctk.CTkButton(self.button_row, text=text, width=120, height=38,
                             fg_color=color, hover_color=hover, command=_click,
                             font=("Segoe UI", 13, "bold" if default else "normal"))
        btn.pack(side=side, padx=(8, 0) if side == "right" else (0, 8))
        if default:
            self._on_default_enter = _click
        return btn


def _run(parent, dialog):
    parent.wait_window(dialog)
    # Windows in particular can leave the parent window looking "dead"/
    # unresponsive after a modal Toplevel closes if focus isn't handed
    # back explicitly -- so always reclaim it here.
    try:
        parent.focus_force()
        parent.lift()
    except Exception:
        pass
    return dialog.result


def show_info(parent, title, message):
    d = _BaseDialog(parent, title, message, INFO_COLOR, "\u2714")
    d._add_button("OK", INFO_COLOR, "#15803d", True, default=True)
    return _run(parent, d)


def show_error(parent, title, message):
    d = _BaseDialog(parent, title, message, ERROR_COLOR, "\u2715")
    d._add_button("OK", ERROR_COLOR, "#991b1b", True, default=True)
    return _run(parent, d)


def show_warning(parent, title, message):
    d = _BaseDialog(parent, title, message, WARNING_COLOR, "\u26A0")
    d._add_button("OK", WARNING_COLOR, "#b45309", True, default=True)
    return _run(parent, d)


def ask_yes_no(parent, title, message):
    """Returns True / False."""
    d = _BaseDialog(parent, title, message, QUESTION_COLOR, "?")
    d._add_button("Yes", INFO_COLOR, "#15803d", True, default=True)
    d._add_button("No", "#94A3B8", "#64748B", False)
    return _run(parent, d)


def ask_yes_no_cancel(parent, title, message):
    """Returns True / False / None (Cancel)."""
    d = _BaseDialog(parent, title, message, QUESTION_COLOR, "?", width=480)
    d._add_button("Cancel", "#94A3B8", "#64748B", None)
    d._add_button("No", ERROR_COLOR, "#991b1b", False)
    d._add_button("Yes", INFO_COLOR, "#15803d", True, default=True)
    return _run(parent, d)


def ask_company_names(parent, file_names, defaults=None, title="Company Names",
                       message="One entry per file. Edit any name below, then click Save.\n"
                               "This fills the 'Company Name' column for every row from that file."):
    """
    Shows ONE popup listing every file with an editable company-name box
    next to it (pre-filled with a sensible guess), instead of asking one
    file at a time. Returns a dict {file_name: company_name}, or None if
    the user clicked Skip/cancelled (in which case the caller should fall
    back to the defaults).
    """
    defaults = defaults or {}
    entry_holder = {}

    height = min(520, 160 + 46 * len(file_names))

    d = ctk.CTkToplevel(parent)
    d.result = None
    d.title(title)
    d.geometry(f"560x{height}")
    d.resizable(False, False)
    d.configure(fg_color="#F5F7FA")
    d.transient(parent)
    d.grab_set()

    ctk.CTkFrame(d, height=6, corner_radius=0, fg_color=QUESTION_COLOR).pack(fill="x", side="top")

    body = ctk.CTkFrame(d, fg_color="transparent")
    body.pack(fill="both", expand=True, padx=24, pady=(16, 10))

    ctk.CTkLabel(body, text=title, font=("Segoe UI", 16, "bold"), text_color=BRAND,
                 anchor="w").pack(fill="x")
    ctk.CTkLabel(body, text=message, font=("Segoe UI", 12), text_color="#334155",
                 anchor="w", justify="left", wraplength=500).pack(fill="x", pady=(4, 12))

    rows_frame = ctk.CTkScrollableFrame(body, fg_color="white", corner_radius=8)
    rows_frame.pack(fill="both", expand=True)

    for fname in file_names:
        row = ctk.CTkFrame(rows_frame, fg_color="transparent")
        row.pack(fill="x", padx=6, pady=6)
        ctk.CTkLabel(row, text=fname, font=("Segoe UI", 12, "bold"), text_color="#1E3A8A",
                     width=190, anchor="w", wraplength=180, justify="left").pack(side="left")
        entry = ctk.CTkEntry(row, height=34, font=("Segoe UI", 12))
        entry.pack(side="left", fill="x", expand=True, padx=(8, 0))
        entry.insert(0, defaults.get(fname, os.path.splitext(fname)[0]))
        entry_holder[fname] = entry

    button_row = ctk.CTkFrame(d, fg_color="transparent")
    button_row.pack(fill="x", padx=24, pady=(0, 16), side="bottom")

    def _submit():
        d.result = {f: (e.get().strip() or os.path.splitext(f)[0]) for f, e in entry_holder.items()}
        try:
            d.grab_release()
        except Exception:
            pass
        d.destroy()

    def _skip():
        d.result = None
        try:
            d.grab_release()
        except Exception:
            pass
        d.destroy()

    d.protocol("WM_DELETE_WINDOW", _skip)
    ctk.CTkButton(button_row, text="Skip (use file names)", width=170, height=38,
                  fg_color="#94A3B8", hover_color="#64748B", command=_skip).pack(side="right", padx=(8, 0))
    ctk.CTkButton(button_row, text="Save Company Names", width=180, height=38,
                  fg_color=INFO_COLOR, hover_color="#15803d", font=("Segoe UI", 13, "bold"),
                  command=_submit).pack(side="right", padx=(8, 0))

    d.after(50, lambda: (d.update_idletasks(), d.focus_force()))
    return _run(parent, d)
    """Returns the entered string, or None if cancelled."""
    entry_holder = {}

    def _builder(area):
        entry = ctk.CTkEntry(area, placeholder_text=placeholder, height=38,
                              font=("Segoe UI", 13))
        entry.pack(fill="x")
        if default:
            entry.insert(0, default)
        entry.focus_set()
        entry.select_range(0, "end")
        entry_holder["entry"] = entry

    d = _BaseDialog(parent, title, message, QUESTION_COLOR, "\u270E",
                     extra_builder=_builder)

    def _submit():
        d.result = entry_holder["entry"].get().strip()
        try:
            d.grab_release()
        except Exception:
            pass
        d.destroy()

    def _cancel():
        d.result = None
        try:
            d.grab_release()
        except Exception:
            pass
        d.destroy()

    entry_holder["entry"].bind("<Return>", lambda e: _submit())
    cancel_btn = ctk.CTkButton(d.button_row, text="Skip", width=100, height=38,
                                fg_color="#94A3B8", hover_color="#64748B", command=_cancel)
    cancel_btn.pack(side="right", padx=(8, 0))
    ok_btn = ctk.CTkButton(d.button_row, text="Save", width=120, height=38,
                            fg_color=INFO_COLOR, hover_color="#15803d",
                            font=("Segoe UI", 13, "bold"), command=_submit)
    ok_btn.pack(side="right", padx=(8, 0))
    d._on_default_enter = _submit

    return _run(parent, d)
