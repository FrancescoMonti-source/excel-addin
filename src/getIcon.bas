Attribute VB_Name = "getIcon"
Sub ReplaceLinksWithImages()
    Dim cell As Range
    Dim tmpFile As String
    Dim http As Object
    Dim imgData() As Byte
    Dim fileNum As Integer
    Dim ws As Worksheet
    Dim pic As Picture

    Set ws = ActiveSheet
    Set http = CreateObject("MSXML2.XMLHTTP")

    For Each cell In Selection
        If cell.Value Like "http*" Then
            On Error Resume Next
            ' Get the image
            http.Open "GET", cell.Value, False
            http.Send
            If http.Status = 200 Then
                imgData = http.responseBody
                tmpFile = Environ("TEMP") & "\temp_img" & Format(Now, "yyyymmddhhmmss") & ".jpg"
                
                ' Save image to temp file
                fileNum = FreeFile
                Open tmpFile For Binary As #fileNum
                    Put #fileNum, , imgData
                Close #fileNum

                ' Insert the image (now from local file)
                Set pic = ws.Pictures.Insert(tmpFile)
                Kill tmpFile

                With pic
                    .Left = cell.Left
                    .Top = cell.Top
                    .Width = cell.Width
                    .Height = cell.Height
                    .Placement = xlMoveAndSize
                End With

                ' Optional: Delete the cell content
                ' cell.ClearContents
            End If
            On Error GoTo 0
        End If
    Next cell
End Sub

