from backend.excel_engine import ExcelEngine
from backend.database import DataStore

engine = ExcelEngine()
store = DataStore()

print("=== DATAFUSION TEST ===")

file_path = input("Enter Excel file path (Demo1.xlsx etc): ")

result = engine.load_excel(file_path)

if "error" in result:
    print("ERROR:", result["error"])
else:
    file_name = result["file_name"]

    store.add_upload(file_name, result)

    print("\nUPLOAD SUCCESS ✔")
    print("File:", file_name)

    print("\nSHEETS FOUND:", list(result["sheets"].keys()))

    print("\nFIRST SHEET SAMPLE:")
    first_sheet = list(result["sheets"].keys())[0]
    print(result["sheets"][first_sheet]["data"][:3])