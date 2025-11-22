# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an ETL (Extract, Transform, Load) pipeline for processing Excel laboratory data, specifically designed for chemical analysis results. The pipeline processes sample data, quality control metrics, and generates formatted reports with conditional formatting.

**Critical Context:** This is a **hybrid Excel Add-in application** designed for offline lab computers. It provides a "one-button" solution directly within Excel, requiring no Python installation on end-user machines.

## Current Status

**Completed on macOS:**
- ✅ Python code refactored from CLI to xlwings Excel integration
- ✅ All modules moved to root directory for PyInstaller compatibility
- ✅ macOS executable built successfully (29MB, all dependencies included)
- ✅ Code tested with real lab data and confirmed working
- ✅ Project structure cleaned and documented

**Pending on Windows:**
- ⏳ Build Windows executable (`ETL_Processor.exe`)
- ⏳ Create VBA macro in Excel to call the executable
- ⏳ Test full workflow with macro button in Excel
- ⏳ Package as `.xlam` add-in file
- ⏳ Deploy to lab computers

## Environment Setup

**Python Version:** Python 3.13 (uses `.venv` virtual environment)

**Install dependencies:**
```bash
pip install -r requirements.txt
```

**Required libraries:**
- pandas (data manipulation)
- openpyxl (Excel file operations)
- xlwings (Excel integration and COM automation)
- pyinstaller (for building standalone executables)

## End-User Experience

Lab members interact with this tool entirely within Excel:

1. Open any raw data Excel file
2. Click the custom "Lab Tools" ribbon tab
3. Click the "Process Data" button
4. Within moments, three new sheets appear: "QC", "Samples", "Reported Results"
5. Out-of-bounds values are automatically highlighted in red
6. A pop-up confirms completion

**No Python installation required** - the tool runs as a self-contained executable called by the Excel Add-in.

## Development Setup

**Python Version:** Python 3.13 (uses `.venv` virtual environment)

**Install dependencies:**
```bash
pip install -r requirements.txt
```

**Running during development:**
- Open `ETL_Addin.xlsm` in Excel
- Run the VBA macro that calls the `main()` function from `ETL_Addin.py`
- The ETL process runs on the active sheet

## Architecture

### Hybrid Deployment Model

This project uses a **two-part architecture** designed for offline lab computers:

1. **Excel Add-in (`.xlam` file)**:
   - Provides the UI: custom ribbon tab and "Process Data" button
   - Contains VBA macro that launches the Python executable
   - Lightweight - only handles user interface

2. **Standalone Python Executable (`ETL_Processor.exe`)**:
   - Built with PyInstaller from `ETL_Addin.py`
   - Self-contained: includes Python interpreter, all libraries (pandas, xlwings), and codebase
   - Runs on any Windows machine without Python installation

**Why this architecture?** Lab computers are shared, offline, and lack Python. This approach bundles everything into a single executable while maintaining a native Excel experience.

### ETL Pipeline Flow

The pipeline follows a classic three-stage ETL architecture with class-based separation of concerns:

1. **Extract** (`excel_extract.py`):
   - Reads data from the active Excel sheet using xlwings
   - Filters specific columns: "Sample ID", "Sample Type", "Mean (per analysis type)", "PPM", "Adjusted ABS"
   - Returns a pandas DataFrame

2. **Transform** (`excel_transform.py`):
   - Filters samples by type ("Samples" only)
   - Groups data by Sample ID
   - Calculates analytical metrics: mean PPM, RPD (Relative Percent Difference), percent recovery (%R)
   - Converts PPM to µmol/L using molecular weight (12.01057 for carbon)

3. **Load** (`excel_load.py`):
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

**QC %R bounds (excel_load.py:18):**
- Normal QC: 90-110%
- MDL: 45-145%

**RPD bounds (excel_load.py:22):**
- Maximum: 10%

**QC Targets (excel_load.py:69):**
- MDL: 0.2 ppm
- ICV: 18.0 ppm
- CCV: 10.0 ppm

### Calculation Formulas

**RPD (excel_transform.py:34):**
```python
rpd = abs(v1 - v2) / mean_ppm * 100.0
```
Uses last two values from PPM column.

**Percent Recovery (excel_transform.py:42):**
```python
percent_r = mean_value / target * 100.0
```

**PPM to µmol/L (excel_transform.py:56):**
```python
umol_per_L = ppm_value * 1000.0 / molecular_weight
```
Molecular weight for carbon: 12.01057

## Coding Style

- **Class-based organization**: Each ETL stage is a separate class
- **Naming convention**: `snake_case` for functions and variables
- **Type hints**: Used for function parameters
- **Error handling**: xlwings alerts for user-facing errors
- **Modular design**: Clear separation between extraction, transformation, and loading

## Build and Deployment

### Building the Executable

**macOS (for testing - completed):**
```bash
source .venv/bin/activate
pyinstaller --onefile --name ETL_Processor ETL_Addin.py
```
Creates: `dist/ETL_Processor` (macOS executable)

**Note:** xlwings has limited VBA macro support on macOS. The macOS build is for verifying the PyInstaller packaging works correctly, but full Excel integration testing should be done on Windows.

**Windows (for lab deployment - PENDING):**
```bash
# On Windows machine with Python 3.13 and dependencies installed
pyinstaller --onefile --name ETL_Processor ETL_Addin.py
```
Creates: `dist/ETL_Processor.exe` (Windows 64-bit executable)

**IMPORTANT:** All Python modules (`excel_extract.py`, `excel_transform.py`, `excel_load.py`) must be in the same directory as `ETL_Addin.py`. PyInstaller cannot follow dynamic path modifications (like `sys.path.insert`), so keeping all modules in the root directory ensures they are automatically included.

### Creating the Excel Add-in (Windows)

**Step 1: Create the VBA Macro**
1. Open `ETL_Addin.xlsm` in Excel
2. Enable the Developer tab (File → Options → Customize Ribbon → check Developer)
3. Click Developer → Insert → Button (Form Control)
4. Draw a button on the sheet
5. When prompted, create a new macro named `RunETLPipeline`
6. In the VBA editor, add this code:
```vba
Sub RunETLPipeline()
    ' Path to the ETL_Processor.exe (update as needed)
    Dim exePath As String
    exePath = "C:\Path\To\ETL_Processor.exe"

    ' Run the executable
    Shell exePath, vbNormalFocus
End Sub
```

**Step 2: Save as Excel Add-in**
1. File → Save As
2. Choose "Excel Add-in (*.xlam)" format
3. Save to a permanent location

**Alternative (simpler for testing):**
Keep using `ETL_Addin.xlsm` during development, only create `.xlam` when ready to deploy.

### Deployment to Lab Computers

Distribute two files:
1. `ETL_Processor.exe` - Copy to a permanent location on the lab computer
2. `ETL_Addin.xlam` - Install via Excel's Add-in Manager

**Note:** The `.gitignore` excludes build artifacts (`dist/`, `build/`, `*.spec`, `.xlwings/`) from version control.

## Important Implementation Details

### Refactoring History

This codebase was **refactored from a file-based CLI tool** to the current Excel Add-in architecture:

- **Old model**: `input_files/` → Python script → `output_files/`
- **New model**: Live Excel workbook → xlwings → Same workbook (new sheets)

**Key changes made:**
- `run.py` replaced with `ETL_Addin.py` using `@xw.func` and `xw.Book.caller()`
- `Extract` class now takes `xw.Book` instead of file paths
- `Load` class writes to live workbook, not files
- All `print()` statements replaced with `xw.apps.active.alert()`
- Conditional formatting uses `.api` property for live application
- **Code modules moved to root:** `code/` directory removed, all `.py` files moved to root for PyInstaller compatibility

### Critical Design Decisions

1. **xlwings Integration**: All user feedback uses `xw.apps.active.alert()` for native Excel pop-ups.

2. **In-place Updates**: The Load class writes output sheets directly back to the calling workbook, not to separate files. If sheets exist, they are cleared and reused.

3. **Conditional Formatting**: Applied programmatically via `ws.api.Font.Color` for out-of-bounds values (red text), enabling real-time highlighting without saving/reopening.

4. **Data Grouping**: Maintains insertion order when grouping samples to preserve the original sequence from the input data.

5. **String Matching**: Sample IDs are stripped of whitespace and matched using regex patterns (case-insensitive for QC samples).

6. **No File I/O**: The `input_files/` and `output_files/` directories are legacy - the codebase no longer uses them.
