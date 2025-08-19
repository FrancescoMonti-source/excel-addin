Attribute VB_Name = "range_utils"
' === RangeUtils Module ===
' Safe and efficient range handling for robust macro development

' ?? Get the last non-empty row in a specific column
Function LastNonEmptyRow(col As Variant, Optional ws As Worksheet) As Long
    If ws Is Nothing Then Set ws = ActiveSheet
    On Error Resume Next
    LastNonEmptyRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' ?? Get the value of the last non-empty cell in a column (skips blanks)
Function LastNonEmptyValue(col As Variant, Optional ws As Worksheet) As Variant
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim r As Long
    For r = ws.Rows.Count To 1 Step -1
        If Trim(ws.Cells(r, col).Value) <> "" Then
            LastNonEmptyValue = ws.Cells(r, col).Value
            Exit Function
        End If
    Next r
    LastNonEmptyValue = ""
End Function



' ?? Trim unused rows/columns and reset used range
Sub TrimAndResetUsedRange(Optional ws As Worksheet)
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim lastRow As Long, lastCol As Long

    On Error Resume Next
    lastRow = ws.Cells.Find("*", , , , xlByRows, xlPrevious).Row
    lastCol = ws.Cells.Find("*", , , , xlByColumns, xlPrevious).Column

    If lastRow < ws.Rows.Count Then
        ws.Rows(lastRow + 1 & ":" & ws.Rows.Count).Delete
    End If
    If lastCol < ws.Columns.Count Then
        ws.Columns(lastCol + 1 & ":" & ws.Columns.Count).Delete
    End If

    ws.UsedRange ' Force Excel to reset internal used range tracking
    MsgBox "Trimmed unused rows and columns. Save & reopen to finalize.", vbInformation
End Sub

