import re
import pandas as pd
import os
from datetime import datetime

# Text used whenever a column had no real heading in the source file --
# kept as one constant so every screen (Review & Clean Data) and every
# exported workbook shows the exact same wording.
NO_HEADER_TEXT = "No heading in source file"


def _clean_headers(raw_columns):
    """
    Replaces blank or Excel-auto-generated headers (pandas gives blank
    columns names like 'Unnamed: 3') with a clear, human-readable label
    -- "No heading in source file" -- instead of leaving an empty
    header, or a confusing 'Unnamed' label, in the app or exports.
    """
    cleaned = []
    seen = {}
    for idx, col in enumerate(raw_columns):
        text = str(col).strip()
        if text == "" or re.match(r"^Unnamed:\s*\d+$", text):
            text = NO_HEADER_TEXT
        # keep duplicate headers unique so no data silently overwrites another
        if text in seen:
            seen[text] += 1
            text = f"{text} ({seen[text]})"
        else:
            seen[text] = 1
        cleaned.append(text)
    return cleaned



class ExcelEngine:
    def __init__(self):
        self.loaded_files = {}

    def load_excel(self, file_path):
        """
        Loads an Excel file and returns cleaned, structured data.

        Cleaning happens here, immediately after pandas reads the file,
        so every screen (Review Data, Preview Before Consolidation,
        exports) always sees the same already-cleaned data:
          1. Rows where every cell is blank/NaN are dropped.
          2. Columns where every cell is blank/NaN are dropped.
          3. Column names are stripped of stray whitespace.
          4. Blank / Excel-auto-generated ("Unnamed: 3") headers are
             turned into a readable placeholder, then that placeholder
             column is removed outright -- no real heading means it's
             not real data the user asked to see, so it's dropped
             automatically instead of being left for manual deletion.
        This applies to every sheet of every file, current or future,
        not just one file's layout.
        """

        if not os.path.exists(file_path):
            return {"error": "File not found"}

        try:
            excel_data = pd.read_excel(file_path, sheet_name=None)

            structured = {
                "file_name": os.path.basename(file_path),
                "file_path": os.path.abspath(file_path),  # full path, so the UI can re-open this exact file later
                "uploaded_at": str(datetime.now()),
                "company_name": "",   # filled in by the UI right after import
                "sheets": {}
            }

            for sheet_name, df in excel_data.items():

                # 1) Fully blank rows -- gone.
                df = df.dropna(how="all")

                # 2) Fully blank columns -- gone.
                df = df.dropna(axis=1, how="all")

                # 3) Strip stray whitespace from whatever header text remains.
                df.columns = [str(col).strip() for col in df.columns]

                # 4) Blank / "Unnamed: N" headers -> a readable, consistent
                #    placeholder (also de-duplicates repeated headers).
                df.columns = _clean_headers(list(df.columns))

                # 5) Any column that still has no real heading is dropped
                #    outright -- whether or not it happens to hold data.
                df = df.loc[:, ~df.columns.str.contains("no heading", case=False)]

                # NaN -> empty string for anything left in the grid.
                df = df.fillna("")

                columns = list(df.columns)
                data = df.to_dict(orient="records")

                structured["sheets"][sheet_name] = {
                    "row_count": len(data),
                    "columns": columns,
                    "data": data
                }

            self.loaded_files[file_path] = structured

            return structured

        except Exception as e:
            return {"error": str(e)}

    def load_excel_raw(self, file_path):
        """
        Reads a workbook with NO header assumption at all (header=None),
        so every physical row -- title rows, split headers, section
        labels, everything -- comes back exactly as it sits in the
        sheet. Used only for detecting messy reports (Auto-Organize);
        the normal load_excel() above is untouched and still drives
        every existing screen exactly as before.
        """
        if not os.path.exists(file_path):
            return {"error": "File not found"}
        try:
            excel_data = pd.read_excel(file_path, sheet_name=None, header=None)
            sheets = {}
            for sheet_name, df in excel_data.items():
                df = df.fillna("")
                rows = [[str(v).strip() for v in row] for row in df.values.tolist()]
                sheets[sheet_name] = rows
            return {
                "file_name": os.path.basename(file_path),
                "file_path": os.path.abspath(file_path),
                "sheets": sheets
            }
        except Exception as e:
            return {"error": str(e)}

    def get_file(self, file_path):
        return self.loaded_files.get(file_path, None)

    def clear(self):
        self.loaded_files = {}