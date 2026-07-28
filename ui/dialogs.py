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

from backend.filters import format_indian_number

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

        self.button_row = ctk.CTkFrame(self, fg_color="transparent")
        self.button_row.pack(fill="x", padx=26, pady=(0, 18), side="bottom")

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

    def _add_button(self, text, color, hover, value, default=False, side="right", width=None):
        def _click():
            self.result = value
            try:
                self.grab_release()
            except Exception:
                pass
            self.destroy()
        # No fixed width by default -- CTkButton sizes itself to fit the
        # label. A fixed 120px width clipped longer labels (e.g. "Done
        # -- Import Now" rendering as "Imp") once more than 2-3 buttons
        # shared one row. width=None lets CTk auto-size; pass an actual
        # number only when a specific width is genuinely wanted.
        kwargs = {"width": width} if width is not None else {}
        btn = ctk.CTkButton(self.button_row, text=text, height=38,
                             fg_color=color, hover_color=hover, command=_click,
                             font=("Segoe UI", 13, "bold" if default else "normal"), **kwargs)
        btn.pack(side=side, padx=(8, 0) if side == "right" else (0, 8))
        if default:
            self._on_default_enter = _click
        return btn

    def _add_back_button(self):
        """
        Adds a '<- Back' button on the LEFT side of the button row.
        It does exactly one thing: close this popup immediately and
        go back to whatever was open before it (result = None) --
        same as pressing Escape. No recompute, no save, no side
        effects, so it's always instant.
        """
        return self._add_button("\u25C0 Back", "#94A3B8", "#64748B", None, side="left")


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


def _fit_on_screen(d, width, desired_height):
    """
    Clamps a custom Toplevel's height to comfortably fit the actual
    screen (leaving room for the taskbar/title bar) and centers it on
    screen. Without this, a dialog sized purely from its content could
    end up taller than the visible screen on a smaller laptop display,
    which pushes its Cancel/Consolidate buttons off the bottom edge
    even though they're correctly packed inside the window -- the
    window itself is just bigger than the screen. Returns the height
    actually used, so callers can still use it (e.g. for scroll-area
    sizing) if needed.
    """
    d.update_idletasks()
    screen_h = d.winfo_screenheight()
    screen_w = d.winfo_screenwidth()
    max_h = max(340, screen_h - 140)  # leave room for taskbar + title bar
    height = min(desired_height, max_h)
    x = max(0, (screen_w - width) // 2)
    y = max(0, (screen_h - height) // 2 - 20)
    d.geometry(f"{width}x{height}+{x}+{y}")
    return height


def show_import_progress(parent, total):
    """
    Small, non-blocking "Importing Files..." dialog shown while
    Dashboard._import_paths loads a batch of Excel files one at a time
    (no background thread involved, so the progress bar only moves
    because update() below pumps the Tk event loop itself).

    Returns an object with:
        progress.update(current, total, filename="")  -- call once per
            file, right before/while loading it.
        progress.close()  -- call in a finally block once the batch is
            done, so the window always closes even if loading raises.
    """
    d = ctk.CTkToplevel(parent)
    d.result = None
    d.title("Importing Files\u2026")
    d.configure(fg_color="#F5F7FA")
    d.transient(parent)
    d.resizable(False, False)
    # No close button while a batch is mid-flight -- there's nothing
    # sensible to cancel back to partway through loading.
    d.protocol("WM_DELETE_WINDOW", lambda: None)
    _fit_on_screen(d, 420, 190)

    ctk.CTkFrame(d, height=6, corner_radius=0, fg_color=QUESTION_COLOR).pack(fill="x", side="top")

    body = ctk.CTkFrame(d, fg_color="transparent")
    body.pack(fill="both", expand=True, padx=24, pady=20)

    ctk.CTkLabel(body, text="\U0001F4C2 Importing Files\u2026", font=("Segoe UI", 15, "bold"),
                 text_color=BRAND, anchor="w").pack(fill="x")

    status_label = ctk.CTkLabel(body, text="Preparing\u2026", font=("Segoe UI", 12),
                                 text_color="#334155", anchor="w", wraplength=370, justify="left")
    status_label.pack(fill="x", pady=(10, 10))

    bar = ctk.CTkProgressBar(body, height=14, corner_radius=7, progress_color=INFO_COLOR)
    bar.pack(fill="x")
    bar.set(0)

    count_label = ctk.CTkLabel(body, text=f"0 / {total}", font=("Segoe UI", 11),
                                text_color="#64748B", anchor="e")
    count_label.pack(fill="x", pady=(6, 0))

    d.update_idletasks()
    d.lift()
    d.focus_force()

    class _ImportProgress:
        def update(self, current, total_count, filename=""):
            try:
                frac = (current / total_count) if total_count else 0
                bar.set(min(max(frac, 0), 1))
                status_label.configure(
                    text=f"Loading \u201c{filename}\u201d\u2026" if filename else "Loading\u2026"
                )
                count_label.configure(text=f"{current} / {total_count}")
                # No background thread is doing this loading, so the
                # window needs an explicit pump here or it would just
                # sit frozen-looking until the whole batch finishes.
                d.update_idletasks()
                d.update()
            except Exception:
                pass  # window may already be gone; never let progress UI break an import

        def close(self):
            try:
                d.grab_release()
            except Exception:
                pass
            try:
                d.destroy()
            except Exception:
                pass

    return _ImportProgress()


def show_info(parent, title, message):
    d = _BaseDialog(parent, title, message, INFO_COLOR, "\u2714")
    d._add_back_button()
    d._add_button("OK", INFO_COLOR, "#15803d", True, default=True)
    return _run(parent, d)


def show_error(parent, title, message):
    d = _BaseDialog(parent, title, message, ERROR_COLOR, "\u2715")
    d._add_back_button()
    d._add_button("OK", ERROR_COLOR, "#991b1b", True, default=True)
    return _run(parent, d)


def show_warning(parent, title, message):
    d = _BaseDialog(parent, title, message, WARNING_COLOR, "\u26A0")
    d._add_back_button()
    d._add_button("OK", WARNING_COLOR, "#b45309", True, default=True)
    return _run(parent, d)


def ask_yes_no(parent, title, message):
    """Returns True / False."""
    d = _BaseDialog(parent, title, message, QUESTION_COLOR, "?")
    d._add_back_button()
    d._add_button("Yes", INFO_COLOR, "#15803d", True, default=True)
    d._add_button("No", "#94A3B8", "#64748B", False)
    return _run(parent, d)


def ask_yes_no_cancel(parent, title, message):
    """Returns True / False / None (Cancel)."""
    d = _BaseDialog(parent, title, message, QUESTION_COLOR, "?", width=480)
    d._add_back_button()
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

    button_row = ctk.CTkFrame(d, fg_color="transparent")
    button_row.pack(fill="x", padx=24, pady=(0, 16), side="bottom")
    ctk.CTkButton(button_row, text="Skip (use file names)", width=170, height=38,
                  fg_color="#94A3B8", hover_color="#64748B", command=_skip).pack(side="right", padx=(8, 0))
    ctk.CTkButton(button_row, text="Save Company Names", width=180, height=38,
                  fg_color=INFO_COLOR, hover_color="#15803d", font=("Segoe UI", 13, "bold"),
                  command=_submit).pack(side="right", padx=(8, 0))
    ctk.CTkButton(button_row, text="\u25C0 Back", width=100, height=38,
                  fg_color="#94A3B8", hover_color="#64748B", command=_skip).pack(side="left")

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

    d.after(50, lambda: (d.update_idletasks(), d.focus_force()))
    return _run(parent, d)


def ask_column_names(parent, ambiguous_by_file):
    """
    Shown once per import batch, right after Auto-Organize, for every
    column that had NO header text in the source file but does have
    real data in it (a blank column with no data is dropped
    automatically and never reaches here).

    Instead of silently guessing "Account", "Account (2)", ... for
    each one, this shows a few real sample values from that exact
    column and lets the user type the real name -- e.g. seeing
    "Hancock Whitney # 8577, Origin Operating *4291" makes it obvious
    that column is a bank/account name, not an amount.

    ambiguous_by_file: {file_name: {sheet_name: [{"placeholder", "samples"}, ...]}}

    Returns {file_name: {sheet_name: {placeholder: new_name}}} built
    only from boxes the user actually typed something into (blank
    boxes are left out, so that column keeps its placeholder name --
    still safely editable later from the red-flagged box on the
    Review Data / Preview screens). Returns None if cancelled/skipped
    entirely, in which case the caller should leave every placeholder
    as-is.
    """
    total_cols = sum(len(cols) for sheets in ambiguous_by_file.values() for cols in sheets.values())
    height = min(720, max(360, 220 + 92 * min(total_cols, 7)))

    d = ctk.CTkToplevel(parent)
    d.result = None
    d.title("Name These Columns")
    d.configure(fg_color="#F5F7FA")
    d.transient(parent)
    d.grab_set()
    _fit_on_screen(d, 620, height)

    ctk.CTkFrame(d, height=6, corner_radius=0, fg_color=QUESTION_COLOR).pack(fill="x", side="top")

    entry_holder = {}  # (file_name, sheet_name, placeholder) -> entry widget

    def _submit():
        result = {}
        for (fname, sheet_name, placeholder), entry in entry_holder.items():
            typed = entry.get().strip()
            if not typed:
                continue
            result.setdefault(fname, {}).setdefault(sheet_name, {})[placeholder] = typed
        d.result = result
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
    d.bind("<Escape>", lambda e: _skip())

    button_row = ctk.CTkFrame(d, fg_color="transparent")
    button_row.pack(fill="x", padx=24, pady=(10, 18), side="bottom")
    ctk.CTkButton(button_row, text="Skip (keep placeholders)", width=190, height=38,
                  fg_color="#94A3B8", hover_color="#64748B", command=_skip).pack(side="right", padx=(8, 0))
    ctk.CTkButton(button_row, text="\U0001F4BE Save Names", width=160, height=38,
                  fg_color=INFO_COLOR, hover_color="#15803d", font=("Segoe UI", 13, "bold"),
                  command=_submit).pack(side="right", padx=(8, 0))

    body = ctk.CTkFrame(d, fg_color="transparent")
    body.pack(fill="both", expand=True, padx=24, pady=(16, 0))

    ctk.CTkLabel(body, text="\u2753 Name These Columns", font=("Segoe UI", 16, "bold"),
                 text_color=BRAND, anchor="w").pack(fill="x")
    ctk.CTkLabel(body, text="Auto-Organize found column(s) with no heading in the source file, but "
                            "with real data in them. A few sample values from each are shown below "
                            "-- type the real column name, or leave blank to fix it later.",
                 font=("Segoe UI", 12), text_color="#334155", anchor="w", justify="left",
                 wraplength=560).pack(fill="x", pady=(4, 12))

    list_frame = ctk.CTkScrollableFrame(body, fg_color="#EEF2F7", corner_radius=8)
    list_frame.pack(fill="both", expand=True)

    for fname, sheets in ambiguous_by_file.items():
        ctk.CTkLabel(list_frame, text=f"\U0001F4C4 {fname}", font=("Segoe UI", 13, "bold"),
                     text_color=BRAND, anchor="w").pack(fill="x", padx=4, pady=(10, 2))
        for sheet_name, cols in sheets.items():
            if len(sheets) > 1:
                ctk.CTkLabel(list_frame, text=f"Sheet: {sheet_name}", font=("Segoe UI", 11, "italic"),
                             text_color="#64748B", anchor="w").pack(fill="x", padx=14, pady=(0, 2))
            for col in cols:
                card = ctk.CTkFrame(list_frame, fg_color="white", corner_radius=8,
                                     border_width=1, border_color="#E2E8F0")
                card.pack(fill="x", padx=14, pady=5)

                samples_text = ", ".join(col["samples"]) if col["samples"] else "(no sample values)"
                ctk.CTkLabel(card, text=f"Sample values: {samples_text}", font=("Segoe UI", 11, "italic"),
                             text_color="#64748B", anchor="w", justify="left", wraplength=540
                             ).pack(fill="x", padx=12, pady=(10, 4))

                entry = ctk.CTkEntry(card, height=34, font=("Segoe UI", 12),
                                      placeholder_text="Type the real column name\u2026")
                entry.pack(fill="x", padx=12, pady=(0, 10))
                entry_holder[(fname, sheet_name, col["placeholder"])] = entry

    d.after(50, lambda: (d.update_idletasks(), d.focus_force()))
    return _run(parent, d)


def ask_new_row(parent, columns):
    """
    Shown from Filters & Search's "+ Add Row" button. One entry box per
    column in the current dataset. Returns a {column: value} dict built
    from every box (blanks kept as ""), or None if cancelled.
    """
    height = min(720, max(320, 200 + 46 * min(len(columns), 10)))

    d = ctk.CTkToplevel(parent)
    d.result = None
    d.title("Add Row")
    d.configure(fg_color="#F5F7FA")
    d.transient(parent)
    d.grab_set()
    _fit_on_screen(d, 520, height)

    ctk.CTkFrame(d, height=6, corner_radius=0, fg_color=QUESTION_COLOR).pack(fill="x", side="top")

    entry_holder = {}  # column -> entry widget

    def _submit():
        d.result = {col: entry.get() for col, entry in entry_holder.items()}
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

    d.protocol("WM_DELETE_WINDOW", _cancel)
    d.bind("<Escape>", lambda e: _cancel())

    button_row = ctk.CTkFrame(d, fg_color="transparent")
    button_row.pack(fill="x", padx=24, pady=(10, 18), side="bottom")
    ctk.CTkButton(button_row, text="Cancel", width=110, height=38,
                  fg_color="#94A3B8", hover_color="#64748B", command=_cancel).pack(side="right", padx=(8, 0))
    ctk.CTkButton(button_row, text="\u2795 Add Row", width=140, height=38,
                  fg_color=INFO_COLOR, hover_color="#15803d", font=("Segoe UI", 13, "bold"),
                  command=_submit).pack(side="right", padx=(8, 0))

    body = ctk.CTkFrame(d, fg_color="transparent")
    body.pack(fill="both", expand=True, padx=24, pady=(16, 0))

    ctk.CTkLabel(body, text="\u2795 Add Row", font=("Segoe UI", 16, "bold"),
                 text_color=BRAND, anchor="w").pack(fill="x")
    ctk.CTkLabel(body, text="Fill in as many fields as you need -- the rest are left blank. "
                            "The new row is added to the full dataset, so it flows into "
                            "Analytics and Export just like every imported row.",
                 font=("Segoe UI", 12), text_color="#334155", anchor="w", justify="left",
                 wraplength=460).pack(fill="x", pady=(4, 12))

    list_frame = ctk.CTkScrollableFrame(body, fg_color="#EEF2F7", corner_radius=8)
    list_frame.pack(fill="both", expand=True)

    for col in columns:
        row = ctk.CTkFrame(list_frame, fg_color="transparent")
        row.pack(fill="x", padx=4, pady=5)
        ctk.CTkLabel(row, text=str(col), font=("Segoe UI", 12, "bold"),
                     text_color="#334155", width=160, anchor="w").pack(side="left", padx=(6, 8))
        entry = ctk.CTkEntry(row, height=32, font=("Segoe UI", 12))
        entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
        entry_holder[col] = entry

    if entry_holder:
        first_entry = next(iter(entry_holder.values()))
        d.after(50, lambda: (d.update_idletasks(), first_entry.focus_set()))

    return _run(parent, d)


def ask_pick_folders(parent, base_folder, candidates):
    """
    Lets the user tick as many folders as they want in ONE screen --
    the real fix for "select multiple folders at once", since Windows'
    native folder browser can only ever return a single folder per
    pick (Shift/Ctrl there highlights multiple, but the dialog still
    reports just one -- that's a Windows limitation, not something an
    app can change). This screen is the workaround: pick a parent
    folder once, then tick every subfolder you actually want here.

    base_folder: the parent folder the user picked (shown in the title
    for context).
    candidates: [(label, full_path), ...] from
        file_dialog.list_candidate_folders(base_folder)

    Returns a list of the ticked full_paths, or None if the user
    clicked Cancel / closed the window.
    """
    desired_height = min(640, max(300, 260 + 42 * max(1, len(candidates))))

    d = ctk.CTkToplevel(parent)
    d.result = None
    d.title("Select Folders to Import")
    d.minsize(460, 300)
    d.configure(fg_color="#F5F7FA")
    d.transient(parent)
    d.grab_set()
    _fit_on_screen(d, 600, desired_height)

    ctk.CTkFrame(d, height=6, corner_radius=0, fg_color=QUESTION_COLOR).pack(fill="x", side="top")

    button_row = ctk.CTkFrame(d, fg_color="transparent")
    button_row.pack(fill="x", padx=24, pady=(10, 18), side="bottom")

    vars_by_path = {}  # full_path -> BooleanVar

    def _submit():
        d.result = [p for p, v in vars_by_path.items() if v.get()]
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

    d.protocol("WM_DELETE_WINDOW", _cancel)
    d.bind("<Escape>", lambda e: _cancel())
    d.bind("<Return>", lambda e: _submit())

    ctk.CTkButton(button_row, text="Cancel", width=110, height=38, fg_color="#94A3B8",
                  hover_color="#64748B", command=_cancel).pack(side="right", padx=(8, 0))
    ctk.CTkButton(button_row, text="\U0001F4C2 Import Selected", width=170, height=38,
                  fg_color=INFO_COLOR, hover_color="#15803d", font=("Segoe UI", 13, "bold"),
                  command=_submit).pack(side="right", padx=(8, 0))
    ctk.CTkButton(button_row, text="\u25C0 Back", width=100, height=38,
                  fg_color="#94A3B8", hover_color="#64748B", command=_cancel).pack(side="left")

    body = ctk.CTkFrame(d, fg_color="transparent")
    body.pack(fill="both", expand=True, padx=24, pady=(16, 0))

    ctk.CTkLabel(body, text="Select Folders to Import", font=("Segoe UI", 16, "bold"),
                 text_color=BRAND, anchor="w").pack(fill="x")
    ctk.CTkLabel(body, text=f"Tick every folder you want imported from '{os.path.basename(base_folder.rstrip(chr(92)+'/'))}'. "
                            f"Each one keeps its own name when grouping files for consolidation.",
                 font=("Segoe UI", 12), text_color="#334155", anchor="w",
                 justify="left", wraplength=520).pack(fill="x", pady=(4, 10))

    if not candidates:
        ctk.CTkLabel(body, text="No subfolders with Excel files were found in that location.",
                     font=("Segoe UI", 12, "italic"), text_color="#64748B", anchor="w").pack(fill="x", pady=10)
        d.after(50, lambda: (d.update_idletasks(), d.focus_force()))
        return _run(parent, d)

    select_all_var = tk.BooleanVar(value=True)

    def _toggle_all():
        val = select_all_var.get()
        for v in vars_by_path.values():
            v.set(val)

    ctk.CTkCheckBox(body, text="Select All / None", font=("Segoe UI", 12, "bold"),
                     text_color=BRAND, variable=select_all_var, command=_toggle_all).pack(anchor="w", pady=(0, 6))

    list_frame = ctk.CTkScrollableFrame(body, fg_color="white", corner_radius=8)
    list_frame.pack(fill="both", expand=True)

    for label, path in candidates:
        var = tk.BooleanVar(value=True)
        vars_by_path[path] = var
        ctk.CTkCheckBox(list_frame, text=label, font=("Segoe UI", 13),
                         variable=var).pack(anchor="w", padx=10, pady=4)

    d.after(50, lambda: (d.update_idletasks(), d.focus_force()))
    return _run(parent, d)


def ask_report_sections(parent, preselected=None):
    """
    Shown right after picking files/sheets to consolidate: lets the
    user choose exactly which sections they want in the Consolidation
    Results screen and the exported Excel report, instead of always
    building everything. The "+" box at the bottom lets them note down
    anything not already listed -- there's no way to compute a brand
    new numeric section without knowing what to calculate, so a custom
    note is carried through and shown clearly in the Summary sheet/
    screen instead, flagged as something to follow up on manually.

    preselected: a previous selection dict to restore (e.g. re-opening
    this after changing files), or None for the sensible
    everything-on default.

    Returns a dict:
        {"checkpoint": bool, "company": bool, "source_files": bool, "custom": [str, ...]}
    or None if the user clicked Skip/closed the window (callers should
    fall back to "everything on" in that case -- this step is a
    convenience, not a gate).
    """
    options = [
        ("checkpoint", "Debit / Credit Checkpoint", "Flags any file/sheet where totals don't tie out."),
        ("company", "Company-wise Summary", "Taxable value, tax and row count per company."),
        ("source_files", "Source Files List", "Every sub file that went into this report, with row counts."),
    ]

    preselected = preselected or {}
    custom_items = list(preselected.get("custom", []))

    d = ctk.CTkToplevel(parent)
    d.result = None
    d.title("What Should This Consolidation Show?")
    d.minsize(460, 420)
    d.configure(fg_color="#F5F7FA")
    d.transient(parent)
    d.grab_set()
    _fit_on_screen(d, 580, 640)

    ctk.CTkFrame(d, height=6, corner_radius=0, fg_color=QUESTION_COLOR).pack(fill="x", side="top")

    button_row = ctk.CTkFrame(d, fg_color="transparent")
    button_row.pack(fill="x", padx=24, pady=(10, 18), side="bottom")

    section_vars = {}

    def _submit():
        result = {key: var.get() for key, var in section_vars.items()}
        result["custom"] = list(custom_items)
        d.result = result
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
    d.bind("<Escape>", lambda e: _skip())

    ctk.CTkButton(button_row, text="Skip (Show Everything)", width=180, height=38, fg_color="#94A3B8",
                  hover_color="#64748B", command=_skip).pack(side="right", padx=(8, 0))
    ctk.CTkButton(button_row, text="\u2714 Consolidate", width=150, height=38,
                  fg_color=INFO_COLOR, hover_color="#15803d", font=("Segoe UI", 13, "bold"),
                  command=_submit).pack(side="right", padx=(8, 0))
    ctk.CTkButton(button_row, text="\u25C0 Back", width=100, height=38,
                  fg_color="#94A3B8", hover_color="#64748B", command=_skip).pack(side="left")

    body = ctk.CTkFrame(d, fg_color="transparent")
    body.pack(fill="both", expand=True, padx=24, pady=(16, 0))

    ctk.CTkLabel(body, text="What Should This Consolidation Show?", font=("Segoe UI", 16, "bold"),
                 text_color=BRAND, anchor="w").pack(fill="x")
    ctk.CTkLabel(body, text="Tick what you want in the results screen and the exported report. "
                            "Use the + box below for anything not listed.",
                 font=("Segoe UI", 12), text_color="#334155", anchor="w",
                 justify="left", wraplength=520).pack(fill="x", pady=(4, 10))

    list_frame = ctk.CTkScrollableFrame(body, fg_color="#EEF2F7", corner_radius=8)
    list_frame.pack(fill="both", expand=True)

    for key, label, desc in options:
        var = tk.BooleanVar(value=preselected.get(key, True))
        section_vars[key] = var
        row = ctk.CTkFrame(list_frame, fg_color="white", corner_radius=6)
        row.pack(fill="x", padx=4, pady=4)
        ctk.CTkCheckBox(row, text=label, font=("Segoe UI", 13, "bold"), text_color=BRAND,
                         variable=var).pack(anchor="w", padx=10, pady=(8, 0))
        ctk.CTkLabel(row, text=desc, font=("Segoe UI", 11), text_color="#64748B",
                     anchor="w", justify="left", wraplength=460).pack(anchor="w", padx=(34, 10), pady=(0, 8))

    ctk.CTkFrame(list_frame, height=1, fg_color="#CBD5E1").pack(fill="x", pady=10, padx=4)

    ctk.CTkLabel(list_frame, text="Anything else you want noted (added as a clear reminder in the report, "
                                   "since it isn't something that can be auto-built):",
                 font=("Segoe UI", 12, "italic"), text_color="#334155",
                 anchor="w", justify="left", wraplength=470).pack(fill="x", padx=6, pady=(0, 6))

    custom_list_frame = ctk.CTkFrame(list_frame, fg_color="transparent")
    custom_list_frame.pack(fill="x", padx=4)

    def _refresh_custom_list():
        for w in custom_list_frame.winfo_children():
            w.destroy()
        for i, item in enumerate(custom_items):
            row = ctk.CTkFrame(custom_list_frame, fg_color="#EFF6FF", corner_radius=6)
            row.pack(fill="x", pady=3)
            ctk.CTkLabel(row, text=item, font=("Segoe UI", 12), text_color="#1E3A8A",
                         anchor="w", justify="left", wraplength=380).pack(side="left", padx=10, pady=6, fill="x", expand=True)
            ctk.CTkButton(row, text="\u2715", width=26, height=26, fg_color="#DC2626",
                          hover_color="#991b1b", command=lambda idx=i: _remove_custom(idx)).pack(side="right", padx=6)

    def _remove_custom(idx):
        del custom_items[idx]
        _refresh_custom_list()

    entry_row = ctk.CTkFrame(list_frame, fg_color="transparent")
    entry_row.pack(fill="x", padx=4, pady=(4, 10))
    custom_entry = ctk.CTkEntry(entry_row, placeholder_text="e.g. Vendor-wise GST breakup")
    custom_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

    def _add_custom():
        text = custom_entry.get().strip()
        if text:
            custom_items.append(text)
            custom_entry.delete(0, "end")
            _refresh_custom_list()

    ctk.CTkButton(entry_row, text="+ Add", width=80, height=32, fg_color=QUESTION_COLOR,
                  hover_color="#1d4ed8", command=_add_custom).pack(side="right")
    custom_entry.bind("<Return>", lambda e: _add_custom())

    _refresh_custom_list()
    d.after(50, lambda: (d.update_idletasks(), d.focus_force()))
    return _run(parent, d)


def ask_add_more(parent, added_so_far=0):
    """
    Shown after each file/folder pick during an import session, so the
    user can keep adding more of either kind before anything is loaded
    -- e.g. two folders plus a couple of loose files, all in one go.

    Returns "files" (open the file picker again), "folder" (open the
    folder picker again), "cancel" (discard everything queued so far
    and import nothing), or None (Done -- import everything queued up
    so far).
    """
    message = (
        f"{added_so_far} file(s) queued so far.\n\n"
        f"Add more individual files, add another folder, finish and import everything now, "
        f"or cancel and discard everything queued so far?"
    )
    d = _BaseDialog(parent, "Add More Files or Folders?", message, QUESTION_COLOR, "+", width=660)
    d._add_button("Cancel", "#94A3B8", "#64748B", "cancel", side="left", width=100)
    d._add_button("Add Files", "#2563EB", "#1d4ed8", "files", width=130)
    d._add_button("Add Folder", "#0F766E", "#0b5c54", "folder", width=140)
    d._add_button("Done \u2014 Import Now", INFO_COLOR, "#15803d", None, default=True, width=210)
    return _run(parent, d)


def ask_import_format(parent, raw_preview, organized_preview):
    """
    Shown once per import batch when any file looks like a messy
    exported report (title/company-name/date rows sitting above the
    real header row -- see backend.auto_organize.needs_organizing).
    Lets the user compare both interpretations before committing the
    whole batch to one mode.

    raw_preview / organized_preview: plain preview text as returned by
    backend.auto_organize.build_previews() (a few pipe-delimited rows).

    Returns "raw" (keep every file exactly as exported) or "auto"
    (auto-clean every messy file in the batch), or None if the user
    cancelled -- dashboard.py falls back to "raw" via `... or "raw"`.
    """
    d = ctk.CTkToplevel(parent)
    d.result = None
    d.title("Choose Import Format")
    d.configure(fg_color="#F5F7FA")
    d.transient(parent)
    d.grab_set()
    _fit_on_screen(d, 640, 560)

    ctk.CTkFrame(d, height=6, corner_radius=0, fg_color=QUESTION_COLOR).pack(fill="x", side="top")

    def _choose(mode):
        d.result = mode
        try:
            d.grab_release()
        except Exception:
            pass
        d.destroy()

    d.protocol("WM_DELETE_WINDOW", lambda: _choose(None))
    d.bind("<Escape>", lambda e: _choose(None))

    button_row = ctk.CTkFrame(d, fg_color="transparent")
    button_row.pack(fill="x", padx=24, pady=(10, 18), side="bottom")
    ctk.CTkButton(button_row, text="Cancel", width=100, height=38, fg_color="#94A3B8",
                  hover_color="#64748B", command=lambda: _choose(None)).pack(side="right", padx=(8, 0))
    ctk.CTkButton(button_row, text="\u2728 Auto-Organize", width=170, height=38,
                  fg_color=INFO_COLOR, hover_color="#15803d", font=("Segoe UI", 13, "bold"),
                  command=lambda: _choose("auto")).pack(side="right", padx=(8, 0))
    ctk.CTkButton(button_row, text="\U0001F4C4 Keep As-Is", width=150, height=38,
                  fg_color="#2563EB", hover_color="#1d4ed8", font=("Segoe UI", 13, "bold"),
                  command=lambda: _choose("raw")).pack(side="right", padx=(8, 0))

    body = ctk.CTkFrame(d, fg_color="transparent")
    body.pack(fill="both", expand=True, padx=24, pady=(16, 0))

    ctk.CTkLabel(body, text="Choose Import Format", font=("Segoe UI", 16, "bold"),
                 text_color=BRAND, anchor="w").pack(fill="x")
    ctk.CTkLabel(body, text="One or more files look like exported reports with title/company rows "
                            "sitting above the real data. Preview both options below, then pick one "
                            "for this whole batch.", font=("Segoe UI", 12), text_color="#334155",
                 anchor="w", justify="left", wraplength=580).pack(fill="x", pady=(4, 14))

    ctk.CTkLabel(body, text="\U0001F4C4 Keep As-Is  \u2014  raw rows, unmodified", font=("Segoe UI", 13, "bold"),
                 text_color="#2563EB", anchor="w").pack(fill="x")
    raw_box = ctk.CTkTextbox(body, height=110, font=("Consolas", 11), fg_color="white",
                              corner_radius=6, wrap="none")
    raw_box.pack(fill="x", pady=(4, 14))
    raw_box.insert("1.0", raw_preview or "(no preview available)")
    raw_box.configure(state="disabled")

    ctk.CTkLabel(body, text="\u2728 Auto-Organize  \u2014  title/metadata rows removed, real headers detected",
                 font=("Segoe UI", 13, "bold"), text_color="#15803d", anchor="w").pack(fill="x")
    org_box = ctk.CTkTextbox(body, height=110, font=("Consolas", 11), fg_color="white",
                              corner_radius=6, wrap="none")
    org_box.pack(fill="x", pady=(4, 4))
    org_box.insert("1.0", organized_preview or "(no preview available)")
    org_box.configure(state="disabled")

    d.after(50, lambda: (d.update_idletasks(), d.focus_force()))
    return _run(parent, d)


def ask_consolidation_choice(parent, loaded_results, preselected=None):
    """
    The "Choose Files to Consolidate" picker (called from Dashboard's
    Consolidate button). Lets the user tick individual sheets, not just
    whole files.

    loaded_results: the full list of imported file dicts (Dashboard's
    self.loaded_results), each shaped like
        {"file_name": ..., "sheets": {sheet_name: {...}, ...}, "source_folder": <optional>}

    Files that share the same "source_folder" (set when the user used
    "Import Folder" instead of picking files one at a time) are grouped
    under a folder heading, so the picker reads:

        Folder Name
            Workbook.xlsx
                Sheet 1
                Sheet 2

    Files imported individually have no "source_folder" and are listed
    the same way but without a folder heading.

    preselected: list/set of file_names that should start ticked
    (defaults to every file + every sheet ticked).

    Returns {file_name: [sheet_name, ...]} for every ticked sheet, or
    None if the user clicked Cancel/closed the window.
    """
    preselected = set(preselected) if preselected is not None else {r["file_name"] for r in loaded_results}

    # group files by source_folder; files with no folder share the
    # `None` bucket and are rendered flat (no extra heading), same as
    # the picker looked before folder-import existed.
    folders = {}
    for r in loaded_results:
        folders.setdefault(r.get("source_folder"), []).append(r)

    # Size the window to how much content there actually is, so with
    # only a few files the Cancel/Consolidate buttons sit right below
    # the list instead of being stranded at the bottom of an
    # unnecessarily tall, mostly-empty window.
    total_rows = sum(1 + len(r["sheets"]) for r in loaded_results) + sum(1 for f in folders if f)
    height = min(720, max(380, 260 + 40 * total_rows))

    d = ctk.CTkToplevel(parent)
    d.result = None
    d.title("Choose Files to Consolidate")
    d.minsize(480, 340)
    d.configure(fg_color="#F5F7FA")
    d.transient(parent)
    d.grab_set()
    _fit_on_screen(d, 580, height)

    ctk.CTkFrame(d, height=6, corner_radius=0, fg_color=QUESTION_COLOR).pack(fill="x", side="top")

    # ---- Buttons are packed to the bottom BEFORE the scrollable list
    # is built, and outside/below it in the widget tree. That guarantees
    # they always stay visible no matter how long the file/sheet list
    # gets -- this is what was missing before (the OK/Consolidate button
    # was never reachable because nothing reserved space for it). ----
    button_row = ctk.CTkFrame(d, fg_color="transparent")
    button_row.pack(fill="x", padx=24, pady=(10, 18), side="bottom")

    file_vars = {}    # file_name -> BooleanVar
    sheet_vars = {}   # (file_name, sheet_name) -> BooleanVar

    def _submit():
        chosen = {}
        for r in loaded_results:
            fname = r["file_name"]
            # A ticked sheet counts on its own -- don't require the
            # file-level box to ALSO be ticked. Before, picking just a
            # sheet while the file box happened to be unticked silently
            # dropped that whole file, producing an empty selection and
            # a confusing "No Files Chosen" warning even though a sheet
            # was clearly picked.
            sheets = [s for s in r["sheets"] if sheet_vars[(fname, s)].get()]
            if sheets:
                chosen[fname] = sheets
        d.result = chosen
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

    d.protocol("WM_DELETE_WINDOW", _cancel)
    d.bind("<Escape>", lambda e: _cancel())

    ctk.CTkButton(button_row, text="Cancel", width=110, height=38, fg_color="#94A3B8",
                  hover_color="#64748B", command=_cancel).pack(side="right", padx=(8, 0))
    ok_btn = ctk.CTkButton(button_row, text="\U0001F4CA Consolidate", width=170, height=38,
                            fg_color=INFO_COLOR, hover_color="#15803d",
                            font=("Segoe UI", 13, "bold"), command=_submit)
    ok_btn.pack(side="right", padx=(8, 0))
    ctk.CTkButton(button_row, text="\u25C0 Back", width=100, height=38,
                  fg_color="#94A3B8", hover_color="#64748B", command=_cancel).pack(side="left")
    d._on_default_enter = _submit
    d.bind("<Return>", lambda e: _submit())

    body = ctk.CTkFrame(d, fg_color="transparent")
    body.pack(fill="both", expand=True, padx=24, pady=(16, 0))

    ctk.CTkLabel(body, text="Choose Files to Consolidate", font=("Segoe UI", 16, "bold"),
                 text_color=BRAND, anchor="w").pack(fill="x")
    ctk.CTkLabel(body, text="Tick the files and sheets you want merged into this consolidation, "
                            "then click Consolidate.", font=("Segoe UI", 12), text_color="#334155",
                 anchor="w", justify="left", wraplength=520).pack(fill="x", pady=(4, 10))

    # ---- Search box: filters the list below by file name or sheet
    # name as the user types. This is purely a view filter -- it never
    # touches file_vars/sheet_vars, so ticks made before/after searching
    # are preserved no matter what's currently visible. ----
    search_row = ctk.CTkFrame(body, fg_color="white", corner_radius=8,
                               border_width=1, border_color="#CBD5E1")
    search_row.pack(fill="x", pady=(0, 10))

    ctk.CTkLabel(search_row, text="\U0001F50D", font=("Segoe UI", 14),
                 text_color="#94A3B8", width=20).pack(side="left", padx=(10, 0))

    search_var = tk.StringVar()
    search_entry = ctk.CTkEntry(
        search_row, textvariable=search_var,
        placeholder_text="Search by file or sheet name\u2026",
        font=("Segoe UI", 13), height=36, border_width=0,
        fg_color="transparent",
    )
    search_entry.pack(side="left", fill="x", expand=True, padx=(6, 4), pady=2)

    clear_btn = ctk.CTkButton(
        search_row, text="\u2715", width=26, height=26, corner_radius=13,
        font=("Segoe UI", 11, "bold"), fg_color="#E2E8F0", hover_color="#CBD5E1",
        text_color="#475569", command=lambda: search_var.set(""),
    )
    # only shown once there's something to clear (packed/hidden in _apply_filter)

    list_frame = ctk.CTkScrollableFrame(body, fg_color="#EEF2F7", corner_radius=8)
    list_frame.pack(fill="both", expand=True)

    no_results_label = ctk.CTkLabel(
        list_frame, text="No files or sheets match your search.",
        font=("Segoe UI", 12, "italic"), text_color="#64748B",
    )

    selection_note = ctk.CTkLabel(body, text="", font=("Segoe UI", 11, "italic"), text_color="#64748B", anchor="w")
    selection_note.pack(fill="x", pady=(6, 0))

    def _update_selection_note():
        n_files = sum(1 for v in file_vars.values() if v.get())
        n_sheets = sum(1 for v in sheet_vars.values() if v.get())
        selection_note.configure(text=f"{n_files} file(s), {n_sheets} sheet(s) selected.")

    def _toggle_file(fname, sheet_names):
        val = file_vars[fname].get()
        for s in sheet_names:
            sheet_vars[(fname, s)].set(val)
        _update_selection_note()

    def _toggle_sheet(fname, sheet_names):
        # Reflect the sheet ticks back onto the file box: ticked if any
        # sheet is on, unticked only once every sheet is off -- purely
        # visual, selection itself no longer depends on this box.
        file_vars[fname].set(any(sheet_vars[(fname, s)].get() for s in sheet_names))
        _update_selection_note()

    folder_headers = {}   # folder_name -> header label widget (truthy folder names only)
    card_entries = []     # ordered list of {fname, folder, card, card_pad, sheet_rows: [(name, checkbox)]}

    for folder_name, results in folders.items():
        if folder_name:
            header = ctk.CTkLabel(list_frame, text=f"\U0001F4C1 {folder_name}", font=("Segoe UI", 13, "bold"),
                                   text_color="#0F766E", anchor="w")
            header.pack(fill="x", padx=4, pady=(10, 4))
            folder_headers[folder_name] = header

        for r in results:
            fname = r["file_name"]
            sheet_names = list(r["sheets"].keys())
            file_vars[fname] = tk.BooleanVar(value=fname in preselected)

            card_pad = 16 if folder_name else 2
            card = ctk.CTkFrame(list_frame, fg_color="white", corner_radius=8,
                                 border_width=1, border_color="#E2E8F0")
            card.pack(fill="x", padx=card_pad, pady=5)

            file_row = ctk.CTkFrame(card, fg_color="transparent")
            file_row.pack(fill="x", padx=14, pady=(10, 4))
            ctk.CTkCheckBox(file_row, text=f"\U0001F4C4 {fname}", font=("Segoe UI", 13, "bold"),
                             text_color=BRAND, variable=file_vars[fname],
                             command=lambda f=fname, s=sheet_names: _toggle_file(f, s)).pack(anchor="w")

            sheets_box = ctk.CTkFrame(card, fg_color="#F8FAFC", corner_radius=6)
            sheets_box.pack(fill="x", padx=14, pady=(0, 10))
            sheet_rows = []
            for sheet_name in sheet_names:
                sheet_vars[(fname, sheet_name)] = tk.BooleanVar(value=fname in preselected)
                row_count = r["sheets"][sheet_name].get("row_count", 0)
                cb = ctk.CTkCheckBox(sheets_box, text=f"{sheet_name}   ({row_count} row(s))",
                                      font=("Segoe UI", 12), text_color="#334155",
                                      variable=sheet_vars[(fname, sheet_name)],
                                      command=lambda f=fname, s=sheet_names: _toggle_sheet(f, s))
                cb.pack(anchor="w", padx=12, pady=4)
                sheet_rows.append((sheet_name, cb))

            card_entries.append({
                "fname": fname, "folder": folder_name, "card": card,
                "card_pad": card_pad, "sheet_rows": sheet_rows,
            })

    def _apply_filter(*_args):
        query = search_var.get().strip().lower()

        if query:
            clear_btn.pack(side="right", padx=(0, 8))
        else:
            clear_btn.pack_forget()

        any_visible = False
        for entry in card_entries:
            fname = entry["fname"]
            file_match = (not query) or (query in fname.lower())

            visible_sheets = 0
            for sheet_name, cb in entry["sheet_rows"]:
                show_sheet = file_match or (query in sheet_name.lower())
                if show_sheet:
                    cb.pack(anchor="w", padx=12, pady=4)
                    visible_sheets += 1
                else:
                    cb.pack_forget()

            if visible_sheets > 0:
                entry["card"].pack(fill="x", padx=entry["card_pad"], pady=5)
                any_visible = True
            else:
                entry["card"].pack_forget()

        for folder_name, header in folder_headers.items():
            folder_has_visible = any(
                e["folder"] == folder_name and e["card"].winfo_ismapped()
                for e in card_entries
            )
            if folder_has_visible:
                header.pack(fill="x", padx=4, pady=(10, 4))
            else:
                header.pack_forget()

        if query and not any_visible:
            no_results_label.pack(pady=24)
        else:
            no_results_label.pack_forget()

    search_var.trace_add("write", _apply_filter)
    search_entry.bind("<Escape>", lambda e: search_var.set(""))

    _update_selection_note()
    d.after(50, lambda: (d.update_idletasks(), d.focus_force()))
    return _run(parent, d)


def show_credit_debit_checkpoint(parent, summary, on_open_file=None, on_refresh=None):
    """
    "Debit / Credit Checkpoint" popup -- shown when Total Debit and
    Total Credit don't tie out across the consolidated data. Lists
    every file/sheet that's individually out of balance, with an
    "Open File" button next to each one so the user can jump straight
    to the source and fix it, and a "Refresh" button that re-imports
    those file(s) from disk and recomputes the checkpoint -- so once
    the file's been corrected and saved, this popup can be updated
    without closing it and re-running the whole consolidation by hand.

    summary: the dict returned by backend.consolidator.credit_debit_checkpoint(...)
    on_open_file: optional callback(mismatch_dict) -- called when an
    "Open File" button is clicked, with that reference's full dict
    (file_name, sheet_name, file_path, debit_col, credit_col, ...) so
    the caller can open the file with those columns highlighted (e.g.
    backend.exporter.open_file_with_highlighted_columns).
    on_refresh: optional callback() -- called when "Refresh" is
    clicked. This popup closes itself first; the caller is expected to
    re-check the totals and either re-open a fresh copy of this same
    popup (still mismatched) or show a "Checkpoint Passed" message.

    This is a hard stop, not just a heads-up: there's no "proceed
    anyway" button here on purpose. The caller (Dashboard) is expected
    to block the consolidation from completing whenever this shows,
    and only let the user through once the totals actually tie out.
    Closing this window (Close button, X, or Escape) just dismisses
    the popup so the user can go fix the file(s) listed.
    """
    mismatches = summary.get("mismatches", [])
    height = min(720, 420 + 70 * max(1, len(mismatches)))

    d = ctk.CTkToplevel(parent)
    d.result = True
    d.title("Debit / Credit Checkpoint")
    d.minsize(520, 380)
    d.configure(fg_color="#F5F7FA")
    d.transient(parent)
    d.grab_set()
    _fit_on_screen(d, 620, height)

    ctk.CTkFrame(d, height=6, corner_radius=0, fg_color=WARNING_COLOR).pack(fill="x", side="top")

    def _close():
        d.result = True
        try:
            d.grab_release()
        except Exception:
            pass
        d.destroy()

    d.protocol("WM_DELETE_WINDOW", _close)
    d.bind("<Escape>", lambda e: _close())
    d.bind("<Return>", lambda e: _close())

    def _refresh():
        # Close this popup first, then hand control back to the
        # caller (Dashboard._refresh_credit_debit_checkpoint), which
        # re-imports the relevant file(s) from disk, recomputes the
        # checkpoint, and either re-opens this popup with fresh
        # numbers or shows "Checkpoint Passed".
        _close()
        if on_refresh:
            on_refresh()

    # ---- pinned footer buttons ---- #
    button_row = ctk.CTkFrame(d, fg_color="transparent")
    button_row.pack(fill="x", padx=24, pady=(10, 18), side="bottom")
    ctk.CTkButton(button_row, text="Close \u2014 I'll Fix the File(s)", height=40,
                  fg_color="#94A3B8", hover_color="#64748B", font=("Segoe UI", 13, "bold"),
                  command=_close).pack(side="right", fill="x", expand=True)
    if on_refresh:
        ctk.CTkButton(button_row, text="\U0001F504 Refresh", width=130, height=40,
                      fg_color=QUESTION_COLOR, hover_color="#1d4ed8", font=("Segoe UI", 13, "bold"),
                      command=_refresh).pack(side="left", padx=(0, 8))
    ctk.CTkButton(button_row, text="\u25C0 Back", width=100, height=40,
                  fg_color="#94A3B8", hover_color="#64748B", command=_close).pack(side="left", padx=(0, 8))

    body = ctk.CTkFrame(d, fg_color="transparent")
    body.pack(fill="both", expand=True, padx=24, pady=(16, 0))

    # ---- header: warning badge + "Debit and Credit totals don't match" ---- #
    top_row = ctk.CTkFrame(body, fg_color="transparent")
    top_row.pack(fill="x")
    ctk.CTkLabel(top_row, text="\u26A0", font=("Segoe UI", 22, "bold"), text_color="white",
                 fg_color=WARNING_COLOR, width=44, height=44, corner_radius=22).pack(side="left", anchor="n")

    text_col = ctk.CTkFrame(top_row, fg_color="transparent")
    text_col.pack(side="left", fill="both", expand=True, padx=(14, 0))
    ctk.CTkLabel(text_col, text="Debit and Credit totals don't match", font=("Segoe UI", 16, "bold"),
                 text_color=BRAND, anchor="w").pack(fill="x")

    totals_row = ctk.CTkFrame(text_col, fg_color="transparent")
    totals_row.pack(fill="x", pady=(6, 0))
    ctk.CTkLabel(totals_row, text=f"Total Debit: {format_indian_number(summary['total_debit'], decimals=2)}",
                 font=("Segoe UI", 12, "bold"), text_color=ERROR_COLOR).pack(side="left")
    ctk.CTkLabel(totals_row, text=f"   |   Total Credit: {format_indian_number(summary['total_credit'], decimals=2)}",
                 font=("Segoe UI", 12, "bold"), text_color=ERROR_COLOR).pack(side="left")
    ctk.CTkLabel(totals_row, text=f"   |   Difference: {format_indian_number(abs(summary['difference']), decimals=2)}",
                 font=("Segoe UI", 12, "bold"), text_color="#334155").pack(side="left")

    # ---- "How to fix this" step-by-step guidance box ---- #
    steps_box = ctk.CTkFrame(body, fg_color="#EFF6FF", corner_radius=6,
                              border_width=1, border_color="#BFDBFE")
    steps_box.pack(fill="x", pady=(12, 4))
    steps_inner = ctk.CTkFrame(steps_box, fg_color="transparent")
    steps_inner.pack(fill="x", padx=14, pady=10)
    ctk.CTkLabel(steps_inner, text="How to fix this", font=("Segoe UI", 12, "bold"),
                 text_color=BRAND, anchor="w").pack(fill="x")
    for step_text in [
        "1.  Click \u201COpen File\u201D on the entry below \u2014 it jumps straight to the sheet and columns that don't tie out.",
        "2.  Fix the mismatched Debit/Credit values in Excel.",
        "3.  Press Ctrl+S to save the file (keep it as .xlsx) and close it.",
        "4.  Come back to this popup and click \u201CRefresh\u201D to re-check the totals.",
    ]:
        ctk.CTkLabel(steps_inner, text=step_text, font=("Segoe UI", 11), text_color="#1E3A8A",
                     anchor="w", justify="left", wraplength=540).pack(fill="x", pady=(4, 0))

    ctk.CTkLabel(body, text="Full references for every file/sheet that's out of balance are listed below.",
                 font=("Segoe UI", 12), text_color="#64748B", anchor="w").pack(fill="x", pady=(10, 8))

    # ---- scrollable list of mismatched file/sheet references ---- #
    list_frame = ctk.CTkScrollableFrame(body, fg_color="transparent")
    list_frame.pack(fill="both", expand=True)

    def _open(mismatch):
        if mismatch.get("file_path") and on_open_file:
            on_open_file(mismatch)

    if not mismatches:
        ctk.CTkLabel(list_frame, text="No individual file/sheet is out of balance on its own -- "
                                       "the mismatch only shows up once everything is combined.",
                     font=("Segoe UI", 12, "italic"), text_color="#64748B",
                     anchor="w", justify="left", wraplength=520).pack(fill="x", pady=10)

    for m in mismatches:
        card = ctk.CTkFrame(list_frame, fg_color="#FEF2F2", corner_radius=6,
                             border_width=1, border_color="#FCA5A5")
        card.pack(fill="x", pady=6)

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=12, pady=10, side="left", expand=True)

        ctk.CTkLabel(inner, text=f"File: {m['file_name']}   |   Sheet: {m['sheet_name']}",
                     font=("Segoe UI", 12, "bold"), text_color="#334155",
                     anchor="w", justify="left").pack(fill="x")
        ctk.CTkLabel(inner,
                     text=(f"Debit column '{m['debit_col']}': {format_indian_number(m['total_debit'], decimals=2)}   |   "
                           f"Credit column '{m['credit_col']}': {format_indian_number(m['total_credit'], decimals=2)}   |   "
                           f"Difference: {format_indian_number(m['difference'], decimals=2)}   |   Rows: {m['row_count']}"),
                     font=("Segoe UI", 11), text_color="#B91C1C",
                     anchor="w", justify="left", wraplength=380).pack(fill="x", pady=(2, 0))

        guidance = m.get("guidance")
        if guidance:
            ctk.CTkLabel(inner, text=guidance, font=("Segoe UI", 11, "italic"),
                         text_color="#7C2D12", anchor="w", justify="left",
                         wraplength=380).pack(fill="x", pady=(2, 0))

        ctk.CTkButton(card, text="\U0001F4C2 Open File", width=110, height=32,
                      fg_color="#2563EB", hover_color="#1d4ed8", font=("Segoe UI", 11),
                      state="normal" if m.get("file_path") else "disabled",
                      command=lambda mm=m: _open(mm)).pack(side="right", padx=12)

    d.after(50, lambda: (d.update_idletasks(), d.focus_force()))
    return _run(parent, d)



def ask_text(parent, title, message, placeholder="", default=""):
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
    d._add_back_button()
    cancel_btn = ctk.CTkButton(d.button_row, text="Skip", width=100, height=38,
                                fg_color="#94A3B8", hover_color="#64748B", command=_cancel)
    cancel_btn.pack(side="right", padx=(8, 0))
    ok_btn = ctk.CTkButton(d.button_row, text="Save", width=120, height=38,
                            fg_color=INFO_COLOR, hover_color="#15803d",
                            font=("Segoe UI", 13, "bold"), command=_submit)
    ok_btn.pack(side="right", padx=(8, 0))
    d._on_default_enter = _submit

    return _run(parent, d)