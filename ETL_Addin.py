import xlwings as xw
import sys
import traceback
import os
from datetime import datetime
from excel_extract import Extract
from excel_transform import Transform
from excel_load import Load

def log_error(message):
    """Write error to log file for debugging"""
    try:
        log_path = os.path.join(os.path.dirname(__file__), "etl_error.log")
        with open(log_path, "a") as f:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"\n{'='*60}\n")
            f.write(f"[{timestamp}]\n")
            f.write(f"{message}\n")
    except:
        pass

def main():
    try:
        log_error("ETL Pipeline started")

        # Connect to the active Excel instance
        # When running as standalone executable, we connect to the existing Excel app
        app = xw.apps.active
        log_error(f"Excel app connection: {app}")

        if app is None:
            error_msg = "Error: No active Excel instance found."
            log_error(error_msg)
            print(error_msg)
            input("Press Enter to close...")
            sys.exit(1)

        # Get the active workbook
        wb = app.books.active
        log_error(f"Active workbook: {wb}")

        if wb is None:
            error_msg = "No workbook is open. Please open a workbook first."
            log_error(error_msg)
            app.alert(error_msg, "ETL Pipeline Error")
            sys.exit(1)

        # Get the name of the active sheet
        sheet_name = wb.sheets.active.name
        log_error(f"Active sheet: {sheet_name}")

        # Inform the user about the processing
        app.alert(f"Processing sheet: {sheet_name}", "ETL Pipeline")

        # Initialize the extractor with the workbook and active sheet
        extractor = Extract(wb, sheet_name)
        raw_data = extractor.extract_data()

        if raw_data is not None:
            log_error("Data extracted successfully")

            # Transform the data
            transformer = Transform(raw_data)

            # Load the transformed data into output sheets
            loader = Load(transformer, wb)
            loader.export_all()
            app.alert("Processing complete!", "ETL Pipeline")
            log_error("Processing completed successfully")
        else:
            error_msg = "Failed to extract data. Check Excel file format."
            log_error(error_msg)
            app.alert(error_msg, "ETL Pipeline Error")

    except Exception as e:
        error_details = f"Exception: {str(e)}\n{traceback.format_exc()}"
        log_error(error_details)
        try:
            xw.apps.active.alert(f"An unexpected error occurred: {e}", "ETL Pipeline Error")
        except:
            print(f"Error: {e}")
            print(traceback.format_exc())
            input("Press Enter to close...")
        sys.exit(1)

# Entry point for standalone executable
if __name__ == "__main__":
    main()