import re
import pandas as pd
import os
from datetime import datetime
from openpyxl.utils import get_column_letter


def _clean_headers(raw_columns):
    """
    Replaces blank or Excel-auto-generated headers (pandas gives blank
    columns names like 'Unnamed: 3') with a clear reference to which
    Excel column it originally was -- e.g. "Column D (no header in
    source file)" -- instead of leaving an empty header, or a confusing
    'Unnamed' label, in the app or exports.
    """
    cleaned = []
    seen = {}
    for idx, col in enumerate(raw_columns):
        text = str(col).strip()
        if text == "" or re.match(r"^Unnamed:\s*\d+$", text):
            letter = get_column_letter(idx + 1)
            text = f"Column {letter} (no header in source file)"
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
        Loads Excel file and returns structured data
        """

        if not os.path.exists(file_path):
            return {"error": "File not found"}

        try:
            excel_data = pd.read_excel(file_path, sheet_name=None)

            structured = {
                "file_name": os.path.basename(file_path),
                "uploaded_at": str(datetime.now()),
                "company_name": "",   # filled in by the UI right after import
                "sheets": {}
            }

            for sheet_name, df in excel_data.items():

                # clean columns -- blank/auto-generated headers become a
                # readable reference (e.g. "Unnamed Column (D)") instead
                # of staying empty and confusing the user later on.
                df.columns = _clean_headers(list(df.columns))

                # convert NaN → empty string
                df = df.fillna("")

                structured["sheets"][sheet_name] = {
                    "row_count": len(df),
                    "columns": list(df.columns),
                    "data": df.to_dict(orient="records")
                }

            self.loaded_files[file_path] = structured

            return structured

        except Exception as e:
            return {"error": str(e)}

    def get_file(self, file_path):
        return self.loaded_files.get(file_path, None)

    def clear(self):
        self.loaded_files = {}