Sub LineOut(ArrayOut() As Variant, ByRef rngOut As Range)
' Usage: Print out 1 line stored in ArrayOut left to right starting with rngOut


rngOut.Resize(1, (UBound(ArrayOut) - LBound(ArrayOut)) + 1) = ArrayOut

End Sub
