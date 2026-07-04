from tkinter import filedialog


def select_excel_file():
    """
    Open file explorer and allow the user to choose ONE Excel file.
    """

    file = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[
            ("Excel Files", "*.xlsx *.xls"),
            ("Excel Workbook", "*.xlsx"),
            ("Excel 97-2003", "*.xls")
        ]
    )

    return file


def select_excel_files():
    """
    Open file explorer and allow the user to choose MULTIPLE Excel files
    at once (e.g. several months of a sales register to consolidate together).
    This is what dashboard.py's "Import Files" button actually calls.
    """

    files = filedialog.askopenfilenames(
        title="Select Excel File(s)",
        filetypes=[
            ("Excel Files", "*.xlsx *.xls"),
            ("Excel Workbook", "*.xlsx"),
            ("Excel 97-2003", "*.xls")
        ]
    )

    return list(files)