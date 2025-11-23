# ETL Pipeline Deployment Guide

## Overview
This guide walks you through deploying the ETL Pipeline as an Excel Add-in on lab computers that are offline and don't have Python installed.

## What You'll Deploy

1. **ETL_Processor.exe** - Standalone Python executable (no Python installation needed)
2. **ETL_Addin.xlam** - Excel Add-in with custom ribbon and VBA macro

## Pre-Deployment: Build the Executable

If you haven't already built the Windows executable:

```bash
.venv\Scripts\pyinstaller.exe --onefile --name ETL_Processor --clean ETL_Addin.py
```

This creates `dist\ETL_Processor.exe` (~29MB, includes all dependencies).

## Step 1: Create the VBA Macro in Excel

1. **Open ETL_Addin.xlsm** in Excel
2. Press `Alt+F11` to open the VBA Editor
3. In the VBA Editor, insert a new module: `Insert → Module`
4. Paste the following VBA code:

```vba
' ETL Pipeline VBA Macro
' This macro calls the standalone Python executable

Sub ProcessData()
    ' Path to the ETL_Processor.exe
    ' IMPORTANT: Update this path to match where you install the executable
    Dim exePath As String
    exePath = Environ("APPDATA") & "\ETL_Pipeline\ETL_Processor.exe"

    ' Check if the executable exists
    If Dir(exePath) = "" Then
        MsgBox "ETL Processor not found at: " & vbCrLf & exePath & vbCrLf & vbCrLf & _
               "Please verify the installation.", vbCritical, "ETL Pipeline Error"
        Exit Sub
    End If

    ' Ensure a workbook is open
    If Application.Workbooks.Count = 0 Then
        MsgBox "Please open a workbook first.", vbExclamation, "ETL Pipeline"
        Exit Sub
    End If

    ' Run the executable
    ' The executable will connect to the active Excel instance
    Dim result As Double
    On Error Resume Next
    result = Shell(exePath, vbNormalFocus)

    If Err.Number <> 0 Then
        MsgBox "Failed to launch ETL Processor: " & Err.Description, vbCritical, "ETL Pipeline Error"
    End If
    On Error GoTo 0
End Sub
```

5. **Save the file** (keep as .xlsm for now)

## Step 2: Create a Custom Ribbon (Optional but Recommended)

Creating a custom ribbon tab gives users a professional "Lab Tools" button.

### Option A: Simple Approach - Quick Access Toolbar

1. In Excel, right-click the Quick Access Toolbar
2. Select "Customize Quick Access Toolbar"
3. Choose "Macros" from the dropdown
4. Select "ProcessData" and click "Add"
5. Click "Modify" to change the icon and name

### Option B: Custom Ribbon Tab (More Professional)

This requires using the Custom UI Editor for Microsoft Office:

1. **Download Office RibbonX Editor** (free tool):
   - https://github.com/fernandreu/office-ribbonx-editor/releases

2. **Open ETL_Addin.xlsm** in the RibbonX Editor

3. **Insert a Custom UI Part** (Office 2010+ Custom UI)

4. **Paste this XML**:

```xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="LabToolsTab" label="Lab Tools">
        <group id="ETLGroup" label="Data Processing">
          <button id="ProcessDataBtn"
                  label="Process Data"
                  size="large"
                  onAction="ProcessData"
                  imageMso="DatabaseTableAnalyze" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

5. **Save and close** the RibbonX Editor

6. **Open the file in Excel** - you should see a "Lab Tools" tab with a "Process Data" button

## Step 3: Save as Excel Add-in (.xlam)

1. In Excel, open `ETL_Addin.xlsm`
2. Go to `File → Save As`
3. Choose `Excel Add-in (*.xlam)` from the file type dropdown
4. Name it `ETL_Addin.xlam`
5. Save to a temporary location (you'll move it during deployment)

## Step 4: Package for Deployment

Create a deployment folder with these files:

```
ETL_Pipeline_Deploy/
├── ETL_Processor.exe          (from dist\ folder)
├── ETL_Addin.xlam             (the add-in you just created)
└── INSTALL_INSTRUCTIONS.txt   (installation steps for lab users)
```

## Step 5: Deploy to Lab Computers

### Recommended Installation Path

Install to a user-specific location that doesn't require admin rights:

**Executable location:**
```
%APPDATA%\ETL_Pipeline\ETL_Processor.exe
```

This translates to:
```
C:\Users\[Username]\AppData\Roaming\ETL_Pipeline\ETL_Processor.exe
```

### Installation Steps for Each Lab Computer

1. **Create the ETL_Pipeline folder:**
   - Press `Win+R`
   - Type `%APPDATA%` and press Enter
   - Create a new folder named `ETL_Pipeline`
   - Copy `ETL_Processor.exe` into this folder

2. **Install the Excel Add-in:**
   - Open Excel
   - Go to `File → Options → Add-ins`
   - At the bottom, select "Excel Add-ins" and click "Go..."
   - Click "Browse"
   - Navigate to where you saved `ETL_Addin.xlam`
   - Select it and click "OK"
   - **Check the box** next to "ETL_Addin" to enable it

3. **Verify Installation:**
   - You should see the "Lab Tools" tab in the Excel ribbon (if using custom ribbon)
   - OR the macro should appear in the Quick Access Toolbar
   - Click "Process Data" to test

## Step 6: Create User Instructions

Save this as `USER_GUIDE.md` in your deployment folder:

```markdown
# ETL Pipeline - User Guide

## How to Use

1. Open your raw data Excel file (the one exported from the instrument)
2. Make sure the raw data sheet is active
3. Click the "Lab Tools" tab in the Excel ribbon
4. Click the "Process Data" button
5. Wait a few seconds while the data is processed
6. Three new sheets will appear:
   - **QC** - Quality control samples with bounds checking
   - **Samples** - Regular samples with RPD calculations
   - **Reported Results** - Final results ready for reporting

## Color Coding

- **Red text** = Value is out of acceptable bounds
  - MDL: Outside 45-145% recovery
  - ICV/CCV: Outside 90-110% recovery
  - RPD: Greater than 10%

## Troubleshooting

**Error: "ETL Processor not found"**
- Contact your lab manager - the software needs to be reinstalled

**Error: "No active Excel instance found"**
- Make sure you have an Excel file open before clicking "Process Data"

**Error: "Failed to extract data"**
- Verify your raw data file is in the correct format
- Make sure you're on the correct sheet before processing
```

## Alternative: Portable Installation

If lab computers have restrictions on %APPDATA%, you can use a shared network drive:

1. Copy `ETL_Processor.exe` to a network location like:
   ```
   \\LabServer\Shared\ETL_Pipeline\ETL_Processor.exe
   ```

2. Update the VBA macro to use the network path:
   ```vba
   exePath = "\\LabServer\Shared\ETL_Pipeline\ETL_Processor.exe"
   ```

## Testing Before Deployment

1. Test the executable standalone:
   - Open a raw data Excel file
   - Run `dist\ETL_Processor.exe` directly
   - Verify it processes the data correctly

2. Test the add-in:
   - Install the .xlam file on your development machine
   - Open a raw data file
   - Click the "Process Data" button
   - Verify all three output sheets are created correctly

## Maintenance & Updates

To update the tool:

1. Rebuild the executable with new code changes
2. Replace `ETL_Processor.exe` on each lab computer
3. The .xlam file only needs updating if you change the VBA macro or ribbon

## Constraints & Considerations

✅ **No Python installation required** - Everything is bundled in the .exe
✅ **Offline operation** - No internet connection needed
✅ **Shared computers** - User-level installation, no admin rights needed
✅ **One-button operation** - Simple workflow for lab staff
✅ **Native Excel experience** - Familiar ribbon interface

⚠️ **File size** - The .exe is ~29MB due to bundled dependencies
⚠️ **Windows only** - This deployment is for Windows lab computers
⚠️ **Excel version** - Requires Excel 2010 or later for custom ribbon support

## Security Notes

- The executable runs with the same permissions as Excel
- No network access is required
- Data never leaves the local machine
- Log files are written to the same directory as the executable for troubleshooting
