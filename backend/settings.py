"""
settings.py

Stores user preferences (default export folder, default zoom level,
default chart type, etc.) in a small JSON file at data/settings.json
so the app remembers them the next time it's opened. This is what
the Settings screen in the dashboard reads/writes.
"""

import json
import os
import platform


def _user_data_dir():
    """Returns a writable, per-user folder to store this app's data in."""
    system = platform.system()
    if system == "Windows":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
        path = os.path.join(base, "DataFusionPlatform")
    elif system == "Darwin":
        path = os.path.join(os.path.expanduser("~"), "Library", "Application Support", "DataFusionPlatform")
    else:
        base = os.environ.get("XDG_DATA_HOME") or os.path.join(os.path.expanduser("~"), ".local", "share")
        path = os.path.join(base, "DataFusionPlatform")
    return path


SETTINGS_PATH = os.path.join(_user_data_dir(), "settings.json")

DEFAULT_SETTINGS = {
    "default_export_folder": os.path.join(os.path.expanduser("~"), "Desktop"),
    "default_zoom": 100,            # percentage, applies to tables on open
    "default_chart_type": "Bar",    # Bar / Pie / Line
    "default_aggregation": "Sum",   # Sum / Count / Average
    "auto_open_after_export": True,  # ask to open file right after export
    "rows_per_page": 500,            # safety cap so huge files don't freeze the UI
}


def load_settings():
    """Returns saved settings merged over the defaults."""
    settings = dict(DEFAULT_SETTINGS)
    try:
        if os.path.exists(SETTINGS_PATH):
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                saved = json.load(f)
            settings.update(saved)
    except (json.JSONDecodeError, OSError):
        pass
    return settings


def save_settings(settings):
    """Returns True on success, False if the file couldn't be written
    (permissions, read-only install location, disk full, etc.) --
    never raises."""
    try:
        os.makedirs(os.path.dirname(SETTINGS_PATH), exist_ok=True)
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
        return True
    except OSError:
        return False

def get_setting(key, default=None):
    return load_settings().get(key, default)


def set_setting(key, value):
    settings = load_settings()
    settings[key] = value
    save_settings(settings)
    return settings
