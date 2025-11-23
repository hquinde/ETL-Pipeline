import pandas as pd
import xlwings as xw

class Extract:
    def __init__(self, workbook: xw.Book, sheet_name: str):
        self.workbook = workbook
        self.sheet_name = sheet_name
        self.header_row_index = 1
        self.cols = ("Sample ID", "Sample Type", "Mean (per analysis type)", "PPM", "Adjusted ABS")


    def extract_data(self):
        wanted_columns = set()
        for column in self.cols:
            wanted_columns.add(column.strip())
        
        # Read the entire sheet into a DataFrame, then filter columns
        try:
            # Read all data from the active sheet to allow for column filtering post-read
            # xlwings reads directly into a DataFrame
            # Get the used range and convert to DataFrame with proper headers
            ws = self.workbook.sheets[self.sheet_name]
            df = ws.used_range.options(pd.DataFrame, header=1, index=False).value

            # Filter columns after reading
            # Convert column names to strings and strip whitespace
            actual_columns = [col for col in df.columns if isinstance(col, str) and col.strip() in wanted_columns]
            df = df[actual_columns]
            
            if len(df) > 0:
                xw.apps.active.alert(f"Successfully loaded {len(df)} rows from sheet '{self.sheet_name}'", "ETL Pipeline")
                return df
            else:
                xw.apps.active.alert(f"No data found in sheet '{self.sheet_name}'.", "ETL Pipeline Error")
                return None

        except Exception as e:
            xw.apps.active.alert(f"Error extracting data from sheet '{self.sheet_name}': {e}", "ETL Pipeline Error")
            return None