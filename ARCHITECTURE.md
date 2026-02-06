# ETL Pipeline Architecture

## Overview

This document provides detailed technical reference for the ETL Pipeline's architecture, data flow, calculations, and design decisions.

## ETL Pipeline Flow

The pipeline follows a classic three-stage ETL architecture with class-based separation of concerns:

### 1. Extract (`excel_extract.py`)
- Reads data from the active Excel sheet using xlwings
- Filters specific columns: "Sample ID", "Sample Type", "Mean (per analysis type)", "PPM", "Adjusted ABS"
- Returns a pandas DataFrame

### 2. Transform (`excel_transform.py`)
- Filters samples by type ("Samples" only)
- Groups data by Sample ID
- Calculates analytical metrics: mean PPM, RPD (Relative Percent Difference), percent recovery (%R)
- Converts PPM to µmol/L using molecular weight (12.01057 for carbon)

### 3. Load (`excel_load.py`)
- Orchestrates final transformation and formatting
- Creates three output sheets: "QC", "Samples", "Reported Results"
- Applies conditional formatting (red text for out-of-bounds values)
- Writes directly to the Excel workbook using xlwings

### Key Classes

- **Extract**: Takes `xw.Book` and sheet name, extracts data into DataFrame
- **Transform**: Takes DataFrame, provides transformation methods
- **Load**: Takes Transform object and `xw.Book`, orchestrates output and formatting

### Data Flow Pattern

```
Excel Workbook (Active Sheet)
  → Extract.extract_data() → DataFrame
  → Transform (various methods) → Transformed Data
  → Load.export_all() → Three Output Sheets in Same Workbook
```

## Domain-Specific Logic

### Sample Type Categories

- **Regular Samples**: Excluded from pattern `^(MDL|ICV|ICB|CCV\d+|CCB\d+|Rinse)$`
- **QC Samples**: MDL, ICV, CCV (with optional numbers)
- **Blank Samples**: ICB, CCB (with optional numbers)

### Quality Control Thresholds

**QC %R bounds** (`excel_load.py:18`):
- Normal QC (ICV/CCV): 90-110%
- MDL: 45-145%

**RPD bounds** (`excel_load.py:22`):
- Maximum: 10%

**QC Targets** (`excel_load.py:69`):
- MDL: 0.2 ppm
- ICV: 18.0 ppm
- CCV: 10.0 ppm

## Calculation Formulas

### RPD (Relative Percent Difference)

**Location:** `excel_transform.py:34`

```python
rpd = abs(v1 - v2) / mean_ppm * 100.0
```

Uses the last two values from the PPM column for each sample.

### Percent Recovery (%R)

**Location:** `excel_transform.py:42`

```python
percent_r = mean_value / target * 100.0
```

Where `target` is the known concentration for QC samples (MDL: 0.2 ppm, ICV: 18.0 ppm, CCV: 10.0 ppm).

### PPM to µmol/L Conversion

**Location:** `excel_transform.py:56`

```python
umol_per_L = ppm_value * 1000.0 / molecular_weight
```

**Molecular weight for carbon:** 12.01057 g/mol

## Critical Design Decisions

### 1. xlwings Integration
All user feedback uses `xw.apps.active.alert()` for native Excel pop-ups instead of console output.

### 2. In-place Updates
The Load class writes output sheets directly back to the calling workbook, not to separate files. If sheets exist, they are cleared and reused.

### 3. Conditional Formatting
Applied programmatically via `r_cell.font.color` (xlwings font property) for out-of-bounds values (red text), enabling real-time highlighting without saving/reopening. Uses RGB tuple format `(255, 0, 0)` for red text.

### 4. Data Grouping
Maintains insertion order when grouping samples to preserve the original sequence from the input data.

### 5. String Matching
Sample IDs are stripped of whitespace and matched using regex patterns (case-insensitive for QC samples).

### 6. No File I/O
The `input_files/` and `output_files/` directories are legacy - the codebase no longer uses them. All operations happen on live workbooks.

### 7. Index Column Removal
DataFrames are written to Excel using `ws.range('A1').options(index=False).value = df` to exclude unnecessary row index numbers from output sheets.

### 8. Error Logging
All errors are logged to `etl_error.log` in the same directory as the executable for troubleshooting in production environments.

## Refactoring History

This codebase was **refactored from a file-based CLI tool** to the current Excel Add-in architecture:

### Old Model (File-based CLI)
```
input_files/ → run.py → output_files/
```

### New Model (Hybrid Excel Add-in)
```
Live Excel workbook → ETL_Addin.py (via xlwings) → Same workbook (new sheets)
```

### Key Changes Made

1. `run.py` replaced with `ETL_Addin.py` using `@xw.func` and `xw.Book.caller()`
2. `Extract` class now takes `xw.Book` instead of file paths
3. `Load` class writes to live workbook, not files
4. All `print()` statements replaced with `xw.apps.active.alert()`
5. Conditional formatting uses `.api` property for live application
6. **Code modules moved to root:** `code/` directory removed, all `.py` files moved to root for PyInstaller compatibility

## Hybrid Deployment Model

This project uses a **two-part architecture** designed for offline lab computers:

### 1. Excel Add-in (`.xlam` file)
- Provides the UI: custom ribbon tab and "Process Data" button
- Contains VBA macro that launches the Python executable
- Lightweight - only handles user interface

### 2. Standalone Python Executable (`ETL_Processor.exe`)
- Built with PyInstaller from `ETL_Addin.py`
- Self-contained: includes Python interpreter, all libraries (pandas, xlwings), and codebase
- Runs on any Windows machine without Python installation

**Why this architecture?** Lab computers are shared, offline, and lack Python. This approach bundles everything into a single executable while maintaining a native Excel experience.

## PyInstaller Packaging

**IMPORTANT:** All Python modules (`excel_extract.py`, `excel_transform.py`, `excel_load.py`) must be in the same directory as `ETL_Addin.py`. PyInstaller cannot follow dynamic path modifications (like `sys.path.insert`), so keeping all modules in the root directory ensures they are automatically included.

**Build command (Windows):**
```bash
.venv\Scripts\pyinstaller.exe --onefile --name ETL_Processor --clean ETL_Addin.py
```

Creates: `dist\ETL_Processor.exe` (Windows 64-bit executable, ~29MB)

## Module Dependencies

```python
# ETL_Addin.py
import xlwings as xw
from excel_extract import Extract
from excel_transform import Transform
from excel_load import Load

# excel_extract.py
import pandas as pd
import xlwings as xw

# excel_transform.py
import pandas as pd
import re

# excel_load.py
import pandas as pd
import xlwings as xw
import re
```

## File Locations Reference

| Component | Development Path | Deployment Path |
|-----------|------------------|-----------------|
| Python Executable | `dist/ETL_Processor.exe` | `%APPDATA%\ETL_Pipeline\ETL_Processor.exe` |
| Excel Add-in | `ETL_Addin.xlsm` → `.xlam` | User's Excel AddIns folder |
| VBA Macro | `ProcessData.vba` | Embedded in `.xlam` |
| Error Log | N/A | Same directory as executable |

## End-User Workflow

1. Open raw data Excel file
2. Click "Lab Tools" ribbon tab
3. Click "Process Data" button
4. VBA macro launches `ETL_Processor.exe`
5. Executable connects to Excel via xlwings COM automation
6. Three new sheets appear in the same workbook
7. Out-of-bounds values highlighted in red
8. Pop-up confirms completion
