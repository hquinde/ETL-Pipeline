' ETL Pipeline VBA Macro
' This macro calls the standalone Python executable
'
' Installation: Copy this code into a VBA module in ETL_Addin.xlsm

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
