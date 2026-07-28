from tkinter import filedialog
import os


def select_excel_file(parent=None):
    """
    Open file explorer and allow the user to choose ONE Excel file.
    """
    if parent is not None:
        parent.lift()
        parent.update_idletasks()

    file = filedialog.askopenfilename(
        parent=parent,
        title="Select Excel File",
        filetypes=[
            ("Excel Files", "*.xlsx *.xls"),
            ("Excel Workbook", "*.xlsx"),
            ("Excel 97-2003", "*.xls")
        ]
    )

    if parent is not None:
        parent.lift()
        parent.focus_force()

    return file


def select_excel_files(parent=None):
    """
    Open file explorer and allow the user to choose MULTIPLE Excel files
    at once (e.g. several months of a sales register to consolidate together).
    This is what dashboard.py's "Import Files" button actually calls.

    `parent` should be the app's main window. Without it, on some Windows
    setups the native file-picker can open *behind* the app window (with
    no taskbar flash) and look like nothing happened when the button is
    clicked -- tying the dialog to `parent` and forcing focus keeps it on
    top and returns focus to the app afterwards.
    """
    if parent is not None:
        parent.lift()
        parent.update_idletasks()

    files = filedialog.askopenfilenames(
        parent=parent,
        title="Select Excel File(s)",
        filetypes=[
            ("Excel Files", "*.xlsx *.xls"),
            ("Excel Workbook", "*.xlsx"),
            ("Excel 97-2003", "*.xls")
        ]
    )

    if parent is not None:
        parent.lift()
        parent.focus_force()

    return list(files)


def select_excel_folder(parent=None):
    """
    Open the folder browser and return (folder_name, [file_paths]) for
    every Excel file found directly inside the chosen folder.

    Lets a user pick a whole folder of workbooks in one go instead of
    picking files one at a time. Dashboard tags every file loaded this
    way with that folder's name, so the "Choose Files to Consolidate"
    picker can show it as:
        folder name
            workbook
                sheets

    Returns (None, []) if the user cancelled, or if the folder has no
    Excel files in it.
    """
    if parent is not None:
        parent.lift()
        parent.update_idletasks()

    folder = filedialog.askdirectory(parent=parent, title="Select Folder of Excel Files")

    if parent is not None:
        parent.lift()
        parent.focus_force()

    if not folder:
        return None, []

    files = []
    for name in sorted(os.listdir(folder)):
        # skip Excel's temporary "lock" files (start with ~$) and anything
        # that isn't actually a workbook
        if name.startswith("~$"):
            continue
        if name.lower().endswith((".xlsx", ".xls")):
            files.append(os.path.join(folder, name))

    if not files:
        return None, []

    return os.path.basename(folder.rstrip("/\\")), files


def _excel_files_in(folder):
    """Excel files (skipping Excel's ~$ lock files) directly inside `folder`."""
    try:
        names = sorted(os.listdir(folder))
    except OSError:
        return []
    return [
        os.path.join(folder, name) for name in names
        if not name.startswith("~$") and name.lower().endswith((".xlsx", ".xls"))
    ]


def select_parent_folder(parent=None):
    """
    Open the folder browser ONCE to pick a parent location (e.g. a
    "Clients" or "Batches" folder that contains one subfolder per
    company/batch). Returns the chosen path, or None if cancelled.
    This is the unavoidable single native pick -- Windows' folder
    browser only ever returns one folder per click, no matter how it's
    selected (that's an OS limitation). What happens next
    (list_candidate_folders + dialogs.ask_pick_folders) is what
    actually lets the user multi-select several folders at once.
    """
    if parent is not None:
        parent.lift()
        parent.update_idletasks()

    folder = filedialog.askdirectory(parent=parent, title="Select Parent Folder")

    if parent is not None:
        parent.lift()
        parent.focus_force()

    return folder or None


def list_candidate_folders(base_folder):
    """
    Returns [(label, full_path), ...] for every folder worth offering
    as an import candidate under base_folder: the base folder itself
    (if it directly holds Excel files) plus every immediate subfolder
    that directly holds Excel files. Used to build the multi-select
    checklist in dialogs.ask_pick_folders so a user can tick e.g. 10
    client subfolders in one screen instead of picking them one at a
    time.
    """
    candidates = []
    if _excel_files_in(base_folder):
        candidates.append((f"(this folder) {os.path.basename(base_folder.rstrip('/\\'))}", base_folder))

    try:
        names = sorted(os.listdir(base_folder))
    except OSError:
        names = []

    for name in names:
        full = os.path.join(base_folder, name)
        if os.path.isdir(full) and _excel_files_in(full):
            candidates.append((name, full))

    return candidates