Attribute VB_Name = "Module1"
Declare Function GetCurrentTime Lib "Kernel32" Alias "GetTickCount" () As Long
Public Sub MilSec(WAIT As Long)
Dim START, c_tim, k, k0
     If WAIT < 1 Then
        k0 = WAIT * ForNexts_Milsec
        For k = 1 To k0: Next
        Exit Sub
     End If
     START = GetCurrentTime()
     Do
        c_tim = GetCurrentTime()
        DoEvents
     Loop Until c_tim >= START + WAIT
End Sub
Function miliseconds(WAIT As Long)
    Dim START, c_tim, k, k0
     START = GetCurrentTime()
     Do
        c_tim = GetCurrentTime()
        DoEvents
     Loop Until c_tim >= START + WAIT
End Function
