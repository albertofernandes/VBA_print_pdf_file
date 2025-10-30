# VBA Print to PDF — from a database sheet to a formatted design sheet
**What this does**
Takes rows from a database sheet, writes selected values into a design_page sheet, and exports one PDF per row using a safe filename pattern.

**Why you might use it**
You have a table of records and a formatted page you want to print per record.
You want deterministic file names (e.g., ID@@Name.pdf).

**Requirements**
1) Excel for Windows (VBA). Tested on Excel 2019.
2) A workbook with two sheets (download from this project folder a working example):
  2.1. database — your data table lives here. First row = headers.
  2.2. design_page — the single-page layout you want to export. Set its print area.
  If you rename the sheets, update the constants in the code.

**Quick start**
Open the VBA editor (Alt + F11), insert a Module, and paste the code below.
Adjust the constants: sheet names, output folder, column numbers.
Map the cells on design_page that you want to populate.
Run PrintDatabaseRowsToPDF.

**The code (drop-in)**

Option Explicit

' ====== USER SETTINGS ======
' Sheet names
Private Const DATABASE_SHEET As String = "database"
Private Const DESIGN_SHEET   As String = "design_page"

' Output folder (ensure you have write permission)
Private Const OUTPUT_FOLDER  As String = "C:\Users\user\Documents\exports"  ' no trailing slash needed

' Column mappings in the database sheet (numbers, 1-based)
Private Const KEY_COL        As Long = 1    ' column used to detect last used row (e.g., an ID or Name)
' Add here your variables and paste them with the design_sheet layout down in the loop
Private Const COL_EXAMPLE1   As Long = 1    ' example: column 1
Private Const COL_EXAMPLE2   As Long = 2   ' example: column 2
Private Const COL_EXAMPLE3   As Long = 3   ' example: column 3

' ====== MACRO ENTRYPOINT ======
Public Sub PrintDatabaseRowsToPDF()
    Dim wsData As Worksheet, wsDesign As Worksheet
    Dim lastRow As Long, i As Long
    Dim outPath As String, fileName As String
    Dim val1 As Variant, val2 As Variant, val3 As Variant
    Dim prevCalc As XlCalculation

    On Error GoTo CleanFail

    Set wsData = ThisWorkbook.Worksheets(DATABASE_SHEET)
    Set wsDesign = ThisWorkbook.Worksheets(DESIGN_SHEET)

    ' Performance toggles
    Application.ScreenUpdating = False
    prevCalc = Application.Calculation
    Application.Calculation = xlCalculationManual

    outPath = EnsureTrailingSlash(OUTPUT_FOLDER)
    CreateFolderIfMissing outPath

    ' Last data row determined by KEY_COL (must have no gaps in that column)
    lastRow = wsData.Cells(wsData.Rows.Count, KEY_COL).End(xlUp).Row
    If lastRow < 2 Then Err.Raise vbObjectError + 100, , "No data rows found in '" & DATABASE_SHEET & "'"

    For i = 2 To lastRow   ' assume row 1 = headers
        ' Read values from the database sheet
        val1 = wsData.Cells(i, COL_EXAMPLE1).Value
        val2 = wsData.Cells(i, COL_EXAMPLE2).Value
        val3 = wsData.Cells(i, COL_EXAMPLE3).Value

        ' === Write values to the design_page (map your cells here) ===
        ' Example mapping: put a composed sentence into B3
        wsDesign.Range("B3").Value = "This is a test " & val1 & " Ending Test"
        ' Add your own mappings like:
        wsDesign.Range("B4").Value = val2
        wsDesign.Range("B5").Value = val3

        ' Build safe filename and export a single-page PDF
        fileName = outPath & SanitizeFileName(CStr(val2) & "@@" & CStr(val3)) & ".pdf"

        wsDesign.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            fileName:=fileName, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            From:=1, To:=1, OpenAfterPublish:=False
    Next i

    MsgBox "Done: " & (lastRow - 1) & " PDFs created in " & outPath, vbInformation

CleanExit:
    ' Restore Excel state
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    MsgBox "Error: " & Err.Number & " — " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' ====== HELPERS ======
Private Function EnsureTrailingSlash(ByVal p As String) As String
    If Len(p) = 0 Then EnsureTrailingSlash = "" Else EnsureTrailingSlash = p & IIf(Right$(p, 1) = "\", "", "\")
End Function

Private Function SanitizeFileName(ByVal s As String) As String
    Dim badChars As Variant, c As Variant
    badChars = Array("<", ">", ":", Chr$(34), "/", "\", "|", "?", "*")
    For Each c In badChars
        s = Replace$(s, c, "_")
    Next
    ' trim and collapse whitespace
    s = Trim$(s)
    s = Replace$(s, vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop
    SanitizeFileName = s
End Function

Private Sub CreateFolderIfMissing(ByVal p As String)
    If Len(p) = 0 Then Exit Sub
    If Dir$(p, vbDirectory) = vbNullString Then MkDir p
End Sub
```

**Contributing**
Keep the macro single-purpose.
Prefer constants for configuration.
Document any additional mapping patterns you add in this README.

**License**
MIT. Do whatever you want; no warranty.

Credits
Original idea by @albertofernandes. Cleanup using AI.
