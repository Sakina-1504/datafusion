"""
app.py

The entry point for DataFusion Platform. Launches the main
dashboard directly (no login/sign-in screen).

Run with: python app.py
"""

import customtkinter as ctk

from ui.dashboard import Dashboard


def launch_dashboard():
    dashboard_root = ctk.CTk()
    Dashboard(dashboard_root)
    dashboard_root.mainloop()


def main():
    launch_dashboard()


if __name__ == "__main__":
    main()
