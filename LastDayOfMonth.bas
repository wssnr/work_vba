Function LastDayOfMonth(dtDay As Date) As Date

'    Dim day As Integer
 '   Dim month As Integer
  '  Dim year As Integer
    
    LastDayOfMonth = DateAdd("d", -1, DateValue(Month(DateAdd("m", 1, dtDay)) & "/1/" & Year(dtDay)))
    
    




End Function
