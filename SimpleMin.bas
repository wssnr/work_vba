Function jMin(a As Double, b As Double) As Double

    jMin = -(a < b) * a + -(b < a) * b + -(a = b) * a
    
    

End Function
