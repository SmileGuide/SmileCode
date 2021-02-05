Attribute VB_Name = "PLUS"

Sub foodp()
Open "F:\BBSeed Files\foodstate.bdf" For Input As #5
Input #5, e
Close #5
If e = 0 Then
nofood.Show
ElseIf e = 1 Then
main.Hide
food.Show
ElseIf e = 3 Then
sorryfood.Show
Else
MsgBox "你提交的养料请求正在等待小抱抱审核，请耐心等待"
End If
End Sub
Sub 刷新()
Open "F:\BBSeed Files\nr.bdf" For Input As #1
Input #1, a
Close #1
Open "F:\BBSeed Files\fd.bdf" For Input As #2
Input #2, b
Close #2
Open "F:\BBSeed Files\wr.bdf" For Input As #3
Input #3, c
Close #3
Open "F:\BBSeed Files\date.bdf" For Input As #4
Input #4, d
Close #4
Open "F:\BBSeed Files\hopestate.bdf" For Input As #5
Input #5, h
Close #5
If h > 0 Then
main.Label8.ForeColor = QBColor(8)
Else
main.Label8.ForeColor = vbBlack
End If
If d <> Date Then
f = 0
Open "F:\BBSeed Files\wr.bdf" For Output As #4
Print #4, f
Close #4
End If
Text1.Text = a
a1 = a
a2 = a
If a > 10 And a < 100 Then
a1 = 10
a11 = "100%+"
a12 = Int(a) & "%"
ElseIf a > 100 Then
a1 = 10
a11 = "100%+"
a2 = 100
a12 = "100%+"
Else
a11 = Int(a * 10) & "%"
a12 = Int(a) & "%"
End If
ProgressBar1.Value = a1 / 10
main.Label3.Caption = a11
ProgressBar2.Value = a2 / 100
main.Label5.Caption = a12
If c = 0 Then
main.Label7.Caption = "今天未浇水"
Else
main.Label7.Caption = "今天已浇水"
End If
End Sub
Sub news()
Open "F:\BBSeed Files\hopestate.bdf" For Input As #1
Input #1, a
Close #1
Open "F:\BBSeed Files\foodstate.bdf" For Input As #3
Input #3, c
Close #3
If a <> 0 Or c <> 0 Then
End If
End Sub
