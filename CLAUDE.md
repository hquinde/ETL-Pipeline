# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an ETL (Extract, Transform, Load) pipeline for processing Excel laboratory data, specifically designed for chemical analysis results. The pipeline processes sample data, quality control metrics, and generates formatted reports with conditional formatting.

**Critical Context:** This is a **hybrid Excel Add-in application** designed for offline lab computers. It provides a "one-button" solution directly within Excel, requiring no Python installation on end-user machines.

## Current Status

**Development Complete - Ready for Lab Deployment:**
- ✅ Python code refactored from CLI to xlwings Excel integration
- ✅ All modules moved to root directory for PyInstaller compatibility
- ✅ Windows executable built successfully (`dist/ETL_Processor.exe`, ~29MB)
- ✅ Code tested with real lab data and confirmed working on Windows
- ✅ Bounds checking working correctly (MDL: 45-145%, ICV/CCV: 90-110%, RPD: ≤10%)
- ✅ Index column removed from output sheets
- ✅ VBA macro created (`ProcessData.vba`)
- ✅ Deployment documentation complete (`DEPLOYMENT_GUIDE.md`)
- ✅ All functionality verified and working

**Pending Lab Deployment:**
- ⏳ Add VBA macro to `ETL_Addin.xlsm`
- ⏳ Save as `ETL_Addin.xlam` add-in file
- ⏳ Add "Process Data" button to custom ribbon tab
- ⏳ Copy `ETL_Processor.exe` to `%APPDATA%\ETL_Pipeline\` on each lab computer
- ⏳ Install `.xlam` add-in on each lab computer

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

**Windows (for lab deployment - COMPLETED):**
```bash
.venv\Scripts\pyinstaller.exe --onefile --name ETL_Processor --clean ETL_Addin.py
```
Creates: `dist\ETL_Processor.exe` (Windows 64-bit executable, ~29MB)

**IMPORTANT:** All Python modules (`excel_extract.py`, `excel_transform.py`, `excel_load.py`) must be in the same directory as `ETL_Addin.py`. PyInstaller cannot follow dynamic path modifications (like `sys.path.insert`), so keeping all modules in the root directory ensures they are automatically included.

**Build Status:** ✅ Successfully built and tested on Windows with real lab data.

### Creating the Excel Add-in (Windows)

**Complete deployment instructions are available in `DEPLOYMENT_GUIDE.md`**

**Quick Summary:**

1. **Add VBA Macro** (code available in `ProcessData.vba`):
   - Open `ETL_Addin.xlsm` → Press `Alt+F11`
   - Insert → Module
   - Paste the ProcessData macro
   - Macro expects executable at: `%APPDATA%\ETL_Pipeline\ETL_Processor.exe`

2. **Save as Excel Add-in**:
   - File → Save As → Choose "Excel Add-in (*.xlam)"
   - Excel will suggest the AddIns folder - accept this location
   - Name it `ETL_Addin.xlam`

3. **Add Custom Ribbon Button** (no internet required):
   - File → Options → Customize Ribbon
   - Create new tab "Lab Tools"
   - Add "ProcessData" macro to the tab
   - This creates a permanent button available in all Excel files

**For detailed step-by-step instructions, see `DEPLOYMENT_GUIDE.md`**

### Deployment to Lab Computers

**Installation Location (no admin rights required):**
```
%APPDATA%\ETL_Pipeline\ETL_Processor.exe
```

**Files to Deploy:**
1. `dist\ETL_Processor.exe` → Copy to `%APPDATA%\ETL_Pipeline\` on each lab computer
2. `ETL_Addin.xlam` → Install via Excel's Add-in Manager (File → Options → Add-ins)

**Key Constraints Addressed:**
- ✅ No internet required (offline lab computers)
- ✅ No Python installation needed (standalone executable)
- ✅ No admin rights required (user-level %APPDATA% installation)
- ✅ Works on shared computers (user-specific installation)
- ✅ Native Excel interface (custom ribbon button)

**See `DEPLOYMENT_GUIDE.md` for complete deployment instructions.**

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

3. **Conditional Formatting**: Applied programmatically via `r_cell.font.color` (xlwings font property) for out-of-bounds values (red text), enabling real-time highlighting without saving/reopening. Uses RGB tuple format `(255, 0, 0)` for red text.

4. **Data Grouping**: Maintains insertion order when grouping samples to preserve the original sequence from the input data.

5. **String Matching**: Sample IDs are stripped of whitespace and matched using regex patterns (case-insensitive for QC samples).

6. **No File I/O**: The `input_files/` and `output_files/` directories are legacy - the codebase no longer uses them.

7. **Index Column Removal**: DataFrames are written to Excel using `ws.range('A1').options(index=False).value = df` to exclude unnecessary row index numbers from output sheets.

8. **Error Logging**: All errors are logged to `etl_error.log` in the same directory as the executable for troubleshooting in production environments.
