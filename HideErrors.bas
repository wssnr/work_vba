Sub HideErrors()
'
' HideErrors Macro
' Macro recorded 7/18/2008 by Jay Woessner
'

'
    Selection.NumberFormat = _
        "[Black]_($* #,##0.000_);[Black]_($* (#,##0.000);[Black]_($* ""-""??_);_(@_)"
    With Selection.Font
          .ColorIndex = 2
    End With
