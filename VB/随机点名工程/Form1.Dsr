VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form1 
   Caption         =   "随机点名"
   ClientHeight    =   2475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4860
   Icon            =   "Form1.dsx":0000
   OleObjectBlob   =   "Form1.dsx":08CA
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Randomize
Dim nam(7) As String
nam(1) = "向美霖"
nam(2) = "王艳霞"
nam(3) = "李俊漾"
nam(4) = "刘子敬"
nam(5) = "迟舒月"
nam(6) = "王  尊"
nam(7) = "赵静涛"
a = Int(Rnd * 6) + 1
If lastnam <> nam(a) Or lastlastnam <> nam(a) Then
lastlastnam = lastnam
lastnam = nam(a)
Label1.Caption = nam(a)
Else
If a = 7 Then
lastlastnam = lastnam
lastnam = nam(a - 1)
Label1.Caption = nam(a - 1)
Else
lastlastnam = lastnam
lastnam = nam(a + 1)
Label1.Caption = nam(a + 1)
End If
End If
cv = cv + 1
End Sub

Private Sub CommandButton2_Click()
Label1.Caption = ""
End Sub


Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
plusb.Show
End Sub
