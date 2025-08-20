Option Explicit

' Wrapper callbacks for Ribbon buttons
' Each Sub matches an onAction in customUI.xml

Public Sub ToUpperCase_UI(control As IRibbonControl)
    ToUpperCase
End Sub

Public Sub ToLowerCase_UI(control As IRibbonControl)
    ToLowerCase
End Sub

Public Sub RemoveDiacritics_UI(control As IRibbonControl)
    ' Works whether you implemented REMOVE_ACCENT as a Function(text) -> text
    ' or a Sub named sub_remove_accent that operates on the selection.
    If TypeName(Selection) <> "Range" Then Exit Sub

    On Error GoTo FallbackSub  ' try function-style first

    Dim c As Range, v As Variant
    For Each c In Selection.Cells
        If Not IsError(c.Value) And Len(c.Value) > 0 Then
            ' Call the function version dynamically: REMOVE_ACCENT(text) -> text
            v = Application.Run("REMOVE_ACCENT", CStr(c.Value))
            c.Value = v
        End If
    Next
    Exit Sub

FallbackSub:
    ' If the function signature isn't available, try the sub that edits selection in-place
    Err.Clear
    On Error Resume Next
    Application.Run "sub_remove_accent"
End Sub


Public Sub FormatMatchingText_UI(control As IRibbonControl)
    FormatMatchingText
End Sub

Public Sub ColorID_UI(control As IRibbonControl)
    ColorID
End Sub

Public Sub GrabImageFromUrl_UI(control As IRibbonControl)
    ReplaceLinksWithImages
End Sub

Public Sub FixValuesInPlace_UI(control As IRibbonControl)
    FixValuesInPlace
End Sub

Public Sub TrimAndResetUsedRange_UI(control As IRibbonControl)
    TrimAndResetUsedRange
End Sub

Public Sub Ribbon_OnLoad(r As IRibbonUI)
    ' optional: MsgBox "Ribbon loaded"
End Sub

Public Sub Ping_UI(control As IRibbonControl)
    MsgBox "Button works", vbInformation
End Sub