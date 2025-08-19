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
    REMOVE_ACCENT
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
    MsgBox "Ribbon loaded", vbInformation
End Sub

Public Sub Ping_UI(control As IRibbonControl)
    MsgBox "Button works", vbInformation
End Sub
