VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "记事本"
   ClientHeight    =   6465
   ClientLeft      =   -60
   ClientTop       =   255
   ClientWidth     =   9180
   Icon            =   "nn.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "nn.dsx":C84A
   StartUpPosition =   2  '屏幕中心
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Open "F:\WPE Files\times.bdf" For Input As #88
Input #88, ac
Close #88
ac = ac + 1
Open "F:\WPE Files\times.bdf" For Output As #45
Print #45, ac
Close #45
a = Date
b = TextBox1.Text
c = TextBox2.Text
If a = "" Or b = "" Then
TextBox1.Text = ""
TextBox1.BackColor = RGB(250, 121, 107)
TextBox2.Text = ""
TextBox2.BackColor = RGB(250, 121, 107)
GoTo 1
MsgBox ("你还没有把内容填写完整")
Else
Open "F:\note1.txt" For Output As #1
Print #1, a
Close #1
Open "F:\note2.txt" For Output As #2
Print #2, b
Close #2
Open "F:\note3.txt" For Output As #3
Print #3, c
Close #3
TextBox1.BackColor = vbWhite
TextBox2.BackColor = vbWhite
MsgBox ("已保存")
UserForm2.Hide
1 End If
End Sub

Private Sub CommandButton10_Click()
MsgBox ("记事本功能：在老师布置交流作业后即可填写记事本内容，在交流即将开始时只需按下主界面的“使用记事本内容”按钮")
End Sub

Private Sub CommandButton8_Click()
Open "F:\WPE Files\number.txt" For Input As #1
Input #1, a
Close #1
TextBox1.Text = a
TextBox2.SetFocus
End Sub

Private Sub CommandButton9_Click()
If MsgBox("是否清空记事本？", vbOKCancel) = vbOK Then
a = ""
b = ""
c = ""
d = ""
Open "F:\note1.txt" For Output As #1
Print #1, b
Close #1
Open "F:\note2.txt" For Output As #2
Print #2, b
Close #2
Open "F:\note3.txt" For Output As #3
Print #3, c
Close #3
Open "F:\note4.txt" For Output As #4
Print #4, d
Close #4
MsgBox "已清空"
Else
GoTo 300
End If
300 End Sub

