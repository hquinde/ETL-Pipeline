import sys
from pathlib import Path
import xlwings as xw

# Add the 'code' directory to the Python path
sys.path.insert(0, str(Path(__file__).parent / "code"))

from excel_extract import Extract
from excel_transform import Transform
from excel_load import Load

@xw.func
def main():
    try:
        # Get the calling Excel workbook
        wb = xw.Book.caller()
        # Get the name of the active sheet
        sheet_name = wb.sheets.active.name

        # Inform the user about the processing
        xw.apps.active.alert(f"Processing sheet: {sheet_name}", "ETL Pipeline", True)

        # Initialize the extractor with the workbook and active sheet
        extractor = Extract(wb, sheet_name)
        raw_data = extractor.extract_data()

        if raw_data is not None:
            transformer = Transform(raw_data)
            # Initialize the loader with the transformer and the workbook
            loader = Load(transformer, wb)
            loader.export_all()
            xw.apps.active.alert("Processing complete!", "ETL Pipeline", True)
        else:
            xw.apps.active.alert("Failed to extract data. Check Excel file format.", "ETL Pipeline Error", True)

    except Exception as e:
        xw.apps.active.alert(f"An unexpected error occurred: {e}", "ETL Pipeline Error", True)