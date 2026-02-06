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

## Step 5: USB Transfer Process

For lab computers without network access, prepare a USB drive with the deployment files:

### Files to Copy to USB Drive

Transfer these files from your development machine:

1. **`dist\ETL_Processor.exe`** (~29MB) - The standalone executable
2. **`ETL_Addin.xlam`** - The Excel Add-in file you created in Step 3
3. **`DEPLOYMENT_GUIDE.md`** (this file) - Step-by-step deployment instructions
4. **`ProcessData.vba`** (optional) - VBA code for reference

### USB Folder Structure

```
USB:\ETL_Pipeline_Deploy\
├── ETL_Processor.exe
├── ETL_Addin.xlam
├── DEPLOYMENT_GUIDE.md
└── ProcessData.vba (optional)
```

## Step 6: Deploy to Lab Computers

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

**Step 6.1: Install the Python Executable**

1. Navigate to `%APPDATA%` folder:
   - Press `Win+R`
   - Type `%APPDATA%` and press Enter
   - Creates path like: `C:\Users\[Username]\AppData\Roaming\`

2. Create the ETL Pipeline folder:
   - Right-click → New → Folder
   - Name it: `ETL_Pipeline`

3. Copy the executable:
   - Copy `ETL_Processor.exe` from USB into `%APPDATA%\ETL_Pipeline\`
   - Final path should be: `%APPDATA%\ETL_Pipeline\ETL_Processor.exe`

**Step 6.2: Install the Excel Add-in**

1. **Enable the Add-in:**
   - Open Excel
   - Go to `File → Options → Add-ins`
   - At the bottom, select "Excel Add-ins" and click "Go..."
   - Click "Browse"
   - Navigate to USB drive and select `ETL_Addin.xlam`
   - Click OK
   - **Check the box** next to "ETL_Addin" to enable it

2. **Verify Installation:**
   - You should see the "Lab Tools" tab in the Excel ribbon (if using custom ribbon)
   - OR the macro should appear in the Quick Access Toolbar

**Step 6.3: Test the Installation**

1. Open a raw data Excel file (or use a test file from USB)
2. Look for the "Lab Tools" ribbon tab
3. Click the "Process Data" button
4. Verify that:
   - Three new sheets appear: "QC", "Samples", "Reported Results"
   - Out-of-bounds values are highlighted in red
   - A completion message pop-up appears

### Deployment Checklist

Use this checklist to track deployment across multiple lab computers:

```
Computer Name: ____________  Date: __________
[ ] Step 1: ETL_Processor.exe copied to %APPDATA%\ETL_Pipeline\
[ ] Step 2: ETL_Addin.xlam installed via Add-ins Manager
[ ] Step 3: "Lab Tools" ribbon tab visible (or macro in Quick Access Toolbar)
[ ] Step 4: Tested with sample data file
[ ] Step 5: Verified all three output sheets generate correctly
[ ] Step 6: Confirmed conditional formatting (red highlights) works
[ ] Step 7: Verified completion pop-up message appears

Notes: _________________________________________________
```

## Troubleshooting

### Common Installation Issues

**Error: "Cannot run the macro..."**
- Ensure `ETL_Processor.exe` is at: `%APPDATA%\ETL_Pipeline\ETL_Processor.exe`
- Verify the path in the VBA macro matches this location
- Check that the file copied correctly from USB (should be ~29MB)

**Error: "ETL Processor not found"**
- Open File Explorer and navigate to `%APPDATA%`
- Verify the `ETL_Pipeline` folder exists
- Verify `ETL_Processor.exe` is inside this folder
- If missing, recopy from USB drive

**Error: "Python script failed"**
- Check `etl_error.log` in the same folder as `ETL_Processor.exe` (`%APPDATA%\ETL_Pipeline\`)
- Verify the input data has the required columns: "Sample ID", "Sample Type", "Mean (per analysis type)", "PPM", "Adjusted ABS"
- Ensure you're on the correct sheet before clicking "Process Data"

**Ribbon button doesn't appear:**
- Verify the add-in is enabled: File → Options → Add-ins
- Look for "ETL_Addin" in the list and ensure it's checked
- If using custom ribbon, restart Excel
- If still missing, check that the VBA macro was properly added (Step 1)

**Error: "No active Excel instance found"**
- Make sure you have an Excel workbook open before clicking "Process Data"
- The workbook must contain raw data in the active sheet

**Add-in disappears after Excel restart:**
- The `.xlam` file may have been moved or deleted
- Reinstall the add-in following Step 6.2
- Consider copying the `.xlam` file to a permanent location before installing

### Common Usage Issues

**No output sheets appear:**
- Check that the active sheet contains the expected column names
- Look for error messages in pop-ups
- Check `etl_error.log` for detailed error information

**Incorrect data in output sheets:**
- Verify the input data format matches the expected structure
- Check that Sample IDs are properly formatted
- Ensure QC samples are named correctly (MDL, ICV, CCV, etc.)

**Process runs but takes a long time:**
- Large datasets may take 30-60 seconds to process
- Do not close Excel or click other buttons while processing

## Step 7: Create User Instructions

Save this as `USER_GUIDE.txt` or print it for lab users:

```
# ETL Pipeline - User Guide

## How to Use

1. Open your raw data Excel file (the one exported from the instrument)
2. Make sure the raw data sheet is active
3. Click the "Lab Tools" tab in the Excel ribbon
4. Click the "Process Data" button
5. Wait a few seconds while the data is processed
6. Three new sheets will appear:
   - QC - Quality control samples with bounds checking
   - Samples - Regular samples with RPD calculations
   - Reported Results - Final results ready for reporting

## Color Coding

Red text = Value is out of acceptable bounds
  - MDL: Outside 45-145% recovery
  - ICV/CCV: Outside 90-110% recovery
  - RPD: Greater than 10%

## If You See an Error

Contact your lab manager if you encounter:
- "ETL Processor not found"
- "No active Excel instance found"
- "Failed to extract data"

Make sure you have the data sheet active before clicking "Process Data"
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
