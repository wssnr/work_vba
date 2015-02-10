Public Function FileBrowse(ByRef objCommonDialog As Object) As String


On Error GoTo cmdBrowse_Click_Err
        
    ChDrive ("C")
    ChDir ("C:\")
      
    objCommonDialog.Filter = "Excel WorkBooks (*.xls)|*.xls"
    objCommonDialog.FilterIndex = 1
    objCommonDialog.Action = 1
    
    If objCommonDialog.Filename <> "" Then
        FileBrowse = objCommonDialog.Filename
    End If
cmdBrowse_Click_Exit:
    Exit Function

cmdBrowse_Click_Err:
    MsgBox Err.Description, , "cmdBrowse_Click"
    Resume cmdBrowse_Click_Exit

End Function
