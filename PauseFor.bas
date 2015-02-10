Sub PauseFor(seconds As Integer)

newHour = Hour(Now())
newMinute = Minute(Now())
newSecond = Second(Now()) + seconds
waitTime = TimeSerial(newHour, newMinute, newSecond)
Application.Wait waitTime
 

End Sub
