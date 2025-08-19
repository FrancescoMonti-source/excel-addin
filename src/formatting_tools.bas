Attribute VB_Name = "formatting_tools"
Function IsValidRGB(rgbArray() As String) As Boolean
    IsValidRGB = False
    If UBound(rgbArray) = 2 Then
        For i = 0 To 2
            If Not IsNumeric(rgbArray(i)) Or CInt(rgbArray(i)) < 0 Or CInt(rgbArray(i)) > 255 Then Exit Function
        Next i
        IsValidRGB = True
    End If
End Function

Sub FormatMatchingText()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim rng As Range, cell As Range
    Dim startPos As Integer
    Dim searchText As String, fontStyle As String, rgbString As String
    Dim rgbArray() As String
    Dim myColor As Long
    Dim formattedCells As Long

    ' Get user inputs
    searchText = InputBox("Enter the text you want to format:")
    If searchText = "" Then GoTo Cleanup

    fontStyle = InputBox("Enter the font style: Bold, Italic, Regular, Underline")
    rgbString = InputBox("Enter the RGB values separated by commas (R,G,B), or leave empty:")

    ' Convert RGB string to individual color components if not empty
    If rgbString <> "" Then
        rgbArray = Split(rgbString, ",")
        If IsValidRGB(rgbArray) Then
            myColor = RGB(CInt(rgbArray(0)), CInt(rgbArray(1)), CInt(rgbArray(2)))
        Else
            MsgBox "Invalid RGB input"
            GoTo Cleanup
        End If
    End If

    Set rng = Selection
    For Each cell In rng
        If Len(cell.Value) > 0 Then
            startPos = 1
            Do
                startPos = InStr(startPos, cell.Value, searchText)
                If startPos > 0 Then
                    formattedCells = formattedCells + 1
                    With cell.Characters(startPos, Len(searchText)).Font
                        .Bold = (UCase(fontStyle) = "BOLD")
                        .Italic = (UCase(fontStyle) = "ITALIC")
                        .Underline = (UCase(fontStyle) = "UNDERLINE")
                        If UCase(fontStyle) = "REGULAR" Then
                            .Bold = False
                            .Italic = False
                            .Underline = xlUnderlineStyleNone
                        End If
                        If rgbString <> "" Then .Color = myColor
                    End With
                    startPos = startPos + Len(searchText)
                Else
                    Exit Do
                End If
            Loop While startPos > 0
        End If
    Next cell

    MsgBox formattedCells & " instances of text were formatted.", vbInformation

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub


Sub sub_remove_accent()
    Dim rng As Range
    Dim cell As Range
    Dim str As String
    Dim i As Integer

    ' Define the selected range
    Set rng = Selection
    
    ' Loop through each cell in the selected range
    For Each cell In rng
        If cell.HasFormula = False Then ' Skip cells with formulas
            str = cell.Value
            For i = Len(str) To 1 Step -1
                If Asc(Mid(str, i, 1)) > 127 Then
                    Mid(str, i, 1) = REMOVE_ACCENT(Mid(str, i, 1))
                End If
            Next i
            cell.Value = str
        End If
    Next cell
End Sub

Function REMOVE_ACCENT(ByVal c As String) As String
    Dim s As String
    s = c
    
    ' Acute
    s = Replace(s, "á", "a")
    s = Replace(s, "é", "e")
    s = Replace(s, "í", "i")
    s = Replace(s, "ó", "o")
    s = Replace(s, "ú", "u")
    s = Replace(s, "Á", "A")
    s = Replace(s, "É", "E")
    s = Replace(s, "Í", "I")
    s = Replace(s, "Ó", "O")
    s = Replace(s, "Ú", "U")
    s = Replace(s, "ý", "y")
    s = Replace(s, "Ý", "Y")
    
    ' Grave
    s = Replace(s, "à", "a")
    s = Replace(s, "è", "e")
    s = Replace(s, "ì", "i")
    s = Replace(s, "ò", "o")
    s = Replace(s, "ù", "u")
    s = Replace(s, "À", "A")
    s = Replace(s, "È", "E")
    s = Replace(s, "Ì", "I")
    s = Replace(s, "Ò", "O")
    s = Replace(s, "Ù", "U")
    
    ' Circumflex
    s = Replace(s, "â", "a")
    s = Replace(s, "ê", "e")
    s = Replace(s, "î", "i")
    s = Replace(s, "ô", "o")
    s = Replace(s, "û", "u")
    s = Replace(s, "Â", "A")
    s = Replace(s, "Ê", "E")
    s = Replace(s, "Î", "I")
    s = Replace(s, "Ô", "O")
    s = Replace(s, "Û", "U")
    
    ' Tilde
    s = Replace(s, "ã", "a")
    s = Replace(s, "õ", "o")
    s = Replace(s, "Ã", "A")
    s = Replace(s, "Õ", "O")
    s = Replace(s, "ñ", "n")
    s = Replace(s, "Ñ", "N")
    
    ' Umlaut
    s = Replace(s, "ä", "a")
    s = Replace(s, "ë", "e")
    s = Replace(s, "ï", "i")
    s = Replace(s, "ö", "o")
    s = Replace(s, "ü", "u")
    s = Replace(s, "Ä", "A")
    s = Replace(s, "Ë", "E")
    s = Replace(s, "Ï", "I")
    s = Replace(s, "Ö", "O")
    s = Replace(s, "Ü", "U")
    s = Replace(s, "ÿ", "y")
    
    ' Cedil
    s = Replace(s, "ç", "c")
    s = Replace(s, "Ç", "C")
    
    REMOVE_ACCENT = s
End Function


Sub ToLowerCase()
    Dim cell As Range
    
    ' Loop through each cell in the selected range
    For Each cell In Selection
        ' Skip cells with formulas
        If cell.HasFormula = False Then
            cell.Value = LCase(cell.Value)
        End If
    Next cell
End Sub


Sub ToUpperCase()
    Dim cell As Range
    
    ' Loop through each cell in the selected range
    For Each cell In Selection
        ' Skip cells with formulas
        If cell.HasFormula = False Then
            cell.Value = UCase(cell.Value)
        End If
    Next cell
End Sub

' ?? Fix formulas in selected range — only the actual used portion
Sub FixValuesInPlace()
    Dim col As Range
    Dim targetRange As Range
    Dim lastRow As Long
    Dim ws As Worksheet
    Set ws = ActiveSheet

    For Each col In Selection.Columns
        lastRow = LastNonEmptyRow(col.Column, ws)
        If lastRow >= col.Row Then
            If targetRange Is Nothing Then
                Set targetRange = ws.Range(ws.Cells(col.Row, col.Column), ws.Cells(lastRow, col.Column))
            Else
                Set targetRange = Union(targetRange, ws.Range(ws.Cells(col.Row, col.Column), ws.Cells(lastRow, col.Column)))
            End If
        End If
    Next col

    If Not targetRange Is Nothing Then
        targetRange.Value = targetRange.Value
    End If
End Sub

