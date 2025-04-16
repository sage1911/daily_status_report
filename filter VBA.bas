' === PASTE THIS ENTIRE BLOCK INTO Module1 in the VBA Editor ===
' === Corrected Sunday Calc, Friday Name, No Unblinded Sheet ===

Sub CreateFilteredSheetsWithHyperlinks()
    Dim wsSource As Worksheet
    Dim wsToday As Worksheet, wsFriday As Worksheet ' Renamed variable
    Dim dueDateCol As Long
    Dim today As Date, nextSundayDate As Date ' Variable still holds Sunday date

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    ' --- Configuration ---
    Const SOURCE_SHEET_NAME As String = "RawData"
    Const TODAY_SHEET_NAME As String = "Today"
    Const FRIDAY_SHEET_NAME As String = "Friday" ' Sheet name is Friday

    ' --- Set Source Sheet ---
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets(SOURCE_SHEET_NAME)
    On Error GoTo ErrorHandler
    If wsSource Is Nothing Then
        Debug.Print "ERROR: Source sheet '" & SOURCE_SHEET_NAME & "' not found!"
        GoTo CleanUp
    End If

    ' --- Find Columns ---
    dueDateCol = FindColumnIndex(wsSource, "Task Due Date")
    If dueDateCol = 0 Then
         Debug.Print "ERROR: Required column 'Task Due Date' not found on '" & SOURCE_SHEET_NAME & "' sheet!"
         GoTo CleanUp
    End If

    ' --- Calculate Dates ---
    today = Date
    ' *** Use correct Sunday calculation ***
    nextSundayDate = GetNextSunday(today)
    Debug.Print "VBA: Filtering 'Friday' sheet using date: " & Format(nextSundayDate, "yyyy-mm-dd") ' Add info

    ' --- Delete Existing Sheets ---
    DeleteSheetIfExists TODAY_SHEET_NAME
    DeleteSheetIfExists FRIDAY_SHEET_NAME    ' Delete the target name "Friday"
    DeleteSheetIfExists "Next Sunday"      ' Delete the old name just in case

    ' --- Create New Sheets ---
    Set wsToday = ThisWorkbook.Sheets.Add(After:=wsSource)
    wsToday.Name = TODAY_SHEET_NAME

    Set wsFriday = ThisWorkbook.Sheets.Add(After:=wsToday) ' Use wsFriday variable
    wsFriday.Name = FRIDAY_SHEET_NAME                  ' Name the sheet "Friday"

    ' --- Filter and Copy ---
    Debug.Print "VBA Filtering for Today..."
    FilterAndCopy wsSource, wsToday, dueDateCol, "<=" & CDbl(today)

    Debug.Print "VBA Filtering for Friday (using Sunday date)..."
    ' *** Filter using the calculated Sunday date ***
    FilterAndCopy wsSource, wsFriday, dueDateCol, "<=" & CDbl(nextSundayDate)

    wsSource.AutoFilterMode = False
    Debug.Print "VBA: Filtered sheets ('Today', 'Friday') created successfully."

CleanUp:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Debug.Print "!!! VBA ERROR in CreateFilteredSheetsWithHyperlinks: " & Err.Description
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not wsSource Is Nothing Then wsSource.AutoFilterMode = False
    On Error GoTo 0
    Resume CleanUp
End Sub


' --- Helper Functions ---

Function GetNextSunday(startDate As Date) As Date
    ' *** RESTORED Correct logic to calculate upcoming Sunday (incl. today if Sunday) ***
    ' Weekday(startDate, vbSunday) returns 1 for Sunday, 2 for Mon... 7 for Sat
    GetNextSunday = startDate + (8 - Weekday(startDate, vbSunday))
End Function

Sub FilterAndCopy(src As Worksheet, dst As Worksheet, col As Long, criteria As String)
    Dim srcRange As Range, visibleRange As Range
    Dim lastRow As Long, lastCol As Long

    On Error GoTo FilterErrorHandler
    src.AutoFilterMode = False

    lastRow = src.Cells(src.Rows.Count, "A").End(xlUp).Row
    If lastRow <= 1 And IsEmpty(src.Range("A1").Value) Then
        Debug.Print "Source sheet '" & src.Name & "' appears empty or header only. Skipping copy for '" & dst.Name & "'."
        Exit Sub
    End If
    lastCol = src.Cells(1, src.Columns.Count).End(xlToLeft).Column
    Set srcRange = src.Range(src.Cells(1, 1), src.Cells(lastRow, lastCol))

    src.Rows(1).Copy
    dst.Rows(1).PasteSpecial xlPasteAll
    dst.Rows(1).PasteSpecial xlPasteColumnWidths

    srcRange.AutoFilter Field:=col, Criteria1:=criteria

    On Error Resume Next
    Set visibleRange = Nothing
    Set visibleRange = srcRange.Offset(1, 0).Resize(srcRange.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
    On Error GoTo FilterErrorHandler

    If Not visibleRange Is Nothing Then
        visibleRange.Copy
        dst.Range("A2").PasteSpecial xlPasteAll
        Dim areaCount As Long, rowCount As Long
        rowCount = 0 ' Initialize count
        For areaCount = 1 To visibleRange.Areas.Count
           rowCount = rowCount + visibleRange.Areas(areaCount).Rows.Count
        Next areaCount
        Debug.Print "Copied approx " & rowCount & " data rows to '" & dst.Name & "'"
    Else
        Debug.Print "No data rows found for '" & dst.Name & "' with criteria: " & criteria
    End If

    src.AutoFilterMode = False
    dst.Columns.AutoFit
    Application.CutCopyMode = False
    Exit Sub

FilterErrorHandler:
     Debug.Print "!!! VBA ERROR during filter/copy for sheet '" & dst.Name & "': " & Err.Description
     On Error Resume Next
     src.AutoFilterMode = False
     Application.CutCopyMode = False
     On Error GoTo 0
End Sub

Sub DeleteSheetIfExists(sheetName As String)
   Dim ws As Worksheet
   On Error Resume Next
   Set ws = ThisWorkbook.Sheets(sheetName)
   On Error GoTo 0
   If Not ws Is Nothing Then
       Application.DisplayAlerts = False
       ws.Delete
       Application.DisplayAlerts = True
       Debug.Print "Deleted existing sheet: " & sheetName
   End If
End Sub

Function FindColumnIndex(ws As Worksheet, headerName As String) As Long
    Dim headerRange As Range
    Dim foundCell As Range

    FindColumnIndex = 0
    If ws Is Nothing Then Exit Function
    ' Check if UsedRange exists before accessing rows count
    On Error Resume Next
    If ws.UsedRange Is Nothing Then Exit Function ' Handle completely empty sheet
    If ws.UsedRange.Rows.Count = 0 Then Exit Function
    On Error GoTo 0 ' Resume normal error handling


    On Error Resume Next
    Set headerRange = ws.Rows(1)
    If Err.Number <> 0 Then
        Debug.Print "Error accessing header row 1 on sheet '" & ws.Name & "'"
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Set foundCell = headerRange.Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False, SearchFormat:=False)

    If Not foundCell Is Nothing Then
        FindColumnIndex = foundCell.Column
        Debug.Print "Found column '" & headerName & "' at index: " & FindColumnIndex
    Else
         Debug.Print "Column '" & headerName & "' NOT FOUND."
    End If
End Function