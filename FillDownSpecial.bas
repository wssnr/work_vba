Sub FillDownSpecial()
'
' FillDownSpecial Macro
'
' Keyboard Shortcut: Ctrl+Shift+F
'
    While 1 = 1
    
    Range(Selection, Selection.End(xlDown).Offset(-1, 0)).Select
    
    Selection.FillDown
    
    Selection.End(xlDown).Select

    Wend
End Sub
