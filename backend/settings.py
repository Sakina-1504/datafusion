"""
settings.py

Stores user preferences (default export folder, default zoom level,
default chart type, etc.) in a small JSON file at data/settings.json
so the app remembers them the next time it's opened. This is what
the Settings screen in the dashboard reads/writes.
"""

import json
import os

SETTINGS_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "settings.json"
)

DEFAULT_SETTINGS = {
    "default_export_folder": os.path.join(os.path.expanduser("~"), "Desktop"),
    "default_zoom": 100,            # percentage, applies to tables on open
    "default_chart_type": "Bar",    # Bar / Pie / Line
    "default_aggregation": "Sum",   # Sum / Count / Average
    "auto_open_after_export": True,  # ask to open file right after export
    "rows_per_page": 500,            # safety cap so huge files don't freeze the UI
}


def load_settings():
    """Returns saved settings merged over the defaults (so new keys
    added later always have a sensible fallback)."""
    settings = dict(DEFAULT_SETTINGS)
    if os.path.exists(SETTINGS_PATH):
        try:
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                saved = json.load(f)
            settings.update(saved)
        except (json.JSONDecodeError, OSError):
            pass
    return settings


def save_settings(settings):
    os.makedirs(os.path.dirname(SETTINGS_PATH), exist_ok=True)
    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2)


def get_setting(key, default=None):
    return load_settings().get(key, default)


def set_setting(key, value):
    settings = load_settings()
    settings[key] = value
    save_settings(settings)
    return settings
