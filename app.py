"""
app.py

The entry point for DataFusion Platform. Launches the main
dashboard directly (no login/sign-in screen).

Run with: python app.py

Also installs a top-level crash handler. Without this, any
unexpected error would either print a raw Python traceback to a
console window most users will never see, or -- in a PyInstaller
"--noconsole" build -- just silently disappear with zero explanation.
Instead, any crash now: (1) writes a full traceback to a log file in
the same per-user app-data folder used for settings, and (2) shows a
plain, friendly error window pointing at that log file.
"""

import os
import sys
import traceback
from datetime import datetime

from backend.settings import _user_data_dir


def _crash_log_path():
    folder = _user_data_dir()
    try:
        os.makedirs(folder, exist_ok=True)
    except OSError:
        folder = os.path.expanduser("~")
    return os.path.join(folder, "datafusion_crash_log.txt")


def _write_crash_log(exc_type, exc_value, exc_tb):
    path = _crash_log_path()
    try:
        with open(path, "a", encoding="utf-8") as f:
            f.write(f"\n{'=' * 60}\n{datetime.now().strftime('%d-%b-%Y %I:%M %p')}\n")
            traceback.print_exception(exc_type, exc_value, exc_tb, file=f)
    except OSError:
        pass
    return path


def _show_crash_dialog(log_path):
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "DataFusion Platform - Something Went Wrong",
            "The app hit an unexpected error and needs to close.\n\n"
            f"Details were saved to:\n{log_path}\n\n"
            "Please share that file if you need support with this."
        )
        root.destroy()
    except Exception:
        pass


def _install_callback_exception_handler(root):
    def handler(exc_type, exc_value, exc_tb):
        log_path = _write_crash_log(exc_type, exc_value, exc_tb)
        try:
            from tkinter import messagebox
            messagebox.showwarning(
                "DataFusion Platform - Action Failed",
                "That action couldn't be completed due to an unexpected error.\n\n"
                f"Details were saved to:\n{log_path}\n\n"
                "You can keep using the app; if this keeps happening, please "
                "share that file for support."
            )
        except Exception:
            pass
    root.report_callback_exception = handler


def launch_dashboard():
    import customtkinter as ctk
    from ui.dashboard import Dashboard

    dashboard_root = ctk.CTk()
    _install_callback_exception_handler(dashboard_root)
    Dashboard(dashboard_root)
    dashboard_root.mainloop()


def main():
    try:
        launch_dashboard()
    except Exception:
        exc_type, exc_value, exc_tb = sys.exc_info()
        log_path = _write_crash_log(exc_type, exc_value, exc_tb)
        _show_crash_dialog(log_path)
        sys.exit(1)


if __name__ == "__main__":
    main()