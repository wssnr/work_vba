End Sub

Public Sub DeleteEmptyRows(ByRef DeleteRange As Range)
' Deletes all empty rows in DeleteRange
' Example: DeleteEmptyRows Selection
' Example: DeleteEmptyRows Range("A1:D100")
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    
        Dim rCount As Long, r As Long
        
        If DeleteRange Is Nothing Then Exit Sub
        If DeleteRange.Areas.Count > 1 Then Exit Sub
        
        With DeleteRange
            rCount = .Rows.Count
            For r = rCount To 1 Step -1
                If Application.CountA(.Rows(r)) = 0 Then
                    .Rows(r).EntireRow.Delete
                End If
            Next r
        
        End With
        
        .Calculation = xlAutomatic
        .ScreenUpdating = True
    
    End With
        
    
End Sub
