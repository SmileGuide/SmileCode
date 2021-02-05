Attribute VB_Name = "Module"


Private Declare Function APIBeep Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, _
ByVal dwDuration As Long) As Long
Sub r()
Dim BB(24) As Integer, I As Integer, J As Integer
    BB(1) = 784: BB(2) = 659: BB(3) = 523: BB(4) = 784
    BB(5) = 659: BB(6) = 523: BB(7) = 880: BB(8) = 698
    BB(9) = 587: BB(10) = 880: BB(11) = 698: BB(12) = 587
    BB(13) = 1568: BB(14) = 1318: BB(15) = 1046
    BB(16) = 1568: BB(17) = 1318: BB(18) = 1046
    BB(19) = 1760: BB(20) = 1396: BB(21) = 1174
    BB(22) = 1760: BB(23) = 1396: BB(24) = 1174
    For I = 1 To 2
    For J = 1 To 24
    APIBeep BB(J), 600
    Next J
    Next I
    times = 1
End Sub


