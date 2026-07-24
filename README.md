# E-Screen

#If VBA7 Then
    Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
    Public TimerID As LongPtr
#Else
    Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
    Public TimerID As Long
#End If

' Starts or resets the 2-minute timer
Sub StartInactivityTimer()
    StopInactivityTimer ' Clear any existing timer first
    ' 120000 milliseconds = 120 seconds (2 minutes)
    TimerID = SetTimer(0, 0, 120000, AddressOf TimerTriggered)
End Sub

' Stops the timer
Sub StopInactivityTimer()
    If TimerID <> 0 Then
        KillTimer 0, TimerID
        TimerID = 0
    End If
End Sub

' What happens when 2 minutes is up
Sub TimerTriggered()
    StopInactivityTimer
    On Error Resume Next
    ' Jump back to slide 1
    SlideShowWindows(1).View.GotoSlide 1
End Sub

' Automatically runs every time a slide changes
Sub OnSlideShowPageChange(ByVal Wn As SlideShowWindow)
    ' If we are on Slide 1, stop the timer (so it doesn't keep looping)
    If Wn.View.CurrentShowPosition = 1 Then
        StopInactivityTimer
    Else
        ' If we are on any other slide, start/reset the 2-minute timer
        StartInactivityTimer
    End If
End Sub
