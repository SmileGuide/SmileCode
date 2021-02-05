VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "第一组、第六组组内交流快捷程序"
   ClientHeight    =   13080
   ClientLeft      =   -60
   ClientTop       =   255
   ClientWidth     =   17115
   Icon            =   "note.dsx":0000
   MinButton       =   0   'False
   OleObjectBlob   =   "note.dsx":C84A
   StartUpPosition =   2  '屏幕中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()
b = 1
If CheckBox1.Value = 0 Then GoTo 1
Open "F:\WPE Files\aa.txt" For Output As #1
Print #1, b
Close #1
GoTo 10
1 c = 0
Open "F:\WPE Files\aa.txt" For Output As #1
Print #1, c
Close #1
10 End Sub

Private Sub CheckBox2_Click()
If CheckBox2.Value = False Then
If MsgBox("是否取消小蛋蛋模式？", vbOKCancel) = vbOK Then
Unload plusb
c = 0
Open "F:\BBSeed Files\nr.bdf" For Output As #2
Print #2, c
Close #1
End If
Else
plusb.Show
End If
1 End Sub

Private Sub CommandButton1_Click()
Open "F:\WPE Files\times.bdf" For Input As #88
Input #88, ac
Close #88
ac = ac + 1
Open "F:\WPE Files\times.bdf" For Output As #45
Print #45, ac
Close #45
If Val(TextBox1.Text) < 0 Then
TextBox1.BackColor = RGB(250, 121, 107)
MsgBox ("请输入正确的序号")
TextBox1.Text = ""
ElseIf TextBox2.Text = "" Or TextBox3.Text = "" Then
TextBox2.BackColor = RGB(250, 121, 107)
TextBox3.BackColor = RGB(250, 121, 107)
TextBox2.Text = ""
TextBox3.Text = ""
MsgBox ("请输入正确的内容")
GoTo 33
Else
GoTo 22
End If
22 If Val(TextBox1.Text) = 0 Then
If MsgBox("是否修复功能？", vbOKCancel) = vbOK Then
TextBox1.BackColor = vbWhite
TextBox2.BackColor = vbWhite
TextBox3.BackColor = vbWhite
TextBox1.Text = ""
TextBox2.Text = ""
TextBox3.Text = ""
Open "F:\WPE Files\number.txt" For Output As #2
a = 1
Print #2, a
Close #2
b = 0
Open "F:\WPE Files\aa.txt" For Output As #1
Close #1
Open "F:\note1.txt" For Output As #1
Close #1
Open "F:\note2.txt" For Output As #2
Close #2
Open "F:\note3.txt" For Output As #3
Close #3
MsgBox "修复成功！"
TextBox4.Text = Now & Chr(10) & "数据与文件已恢复"
TextBox5.Text = "您已可以开始正常使用"
TextBox6.Text = "请做好准备"
GoTo 33
Else
End If
End If
TextBox1.BackColor = vbWhite
TextBox2.BackColor = vbWhite
TextBox3.BackColor = vbWhite
Open "F:\WPE Files\number.txt" For Output As #2
a = Val(TextBox1.Text) + 1
Print #2, a
Close #2
b = 0
Open "F:\WPE Files\aa.txt" For Output As #1
Print #1, b
Close #1
If a <> 1 Then
TextBox4.Text = Now & Chr(10) & "第" & TextBox1.Text & "次交流“" & TextBox2.Text & "”即将开始，交流人员：" & TextBox3.Text & "   请大家做好准备"
TextBox5.Text = "好的，交流开始"
TextBox6.Text = "本次交流到此结束，谢谢大家的配合！"
End If
33 End Sub

Private Sub CommandButton10_Click()
If MsgBox("是否清空内部数据?清除后在交流序号里输入“0”后随便填写内容，再次按下“完成”按钮部分因数据删除导致的错误功能即可恢复", vbOKCancel) = vbOK Then
If Dir("F:\WPE Files\note1.txt") <> "" Then
Kill ("F:\WPE Files\note1.txt")
End If
If Dir("F:\WPE Files\note2.txt") <> "" Then
Kill ("F:\WPE Files\note2.txt")
End If
If Dir("F:\WPE Files\note3.txt") <> "" Then
Kill ("F:\WPE Files\note3.txt")
End If
If Dir("F:\WPE Files\number.txt") <> "" Then
Kill ("F:\WPE Files\number.txt")
End If
If Dir("F:\WPE Files\aa.txt") <> "" Then
Kill ("F:\WPE Files\aa.txt")
End If
Else
GoTo 1
End If
1 End Sub

Private Sub CommandButton12_Click()
MsgBox "在交流序号里输入“0”后随便填写内容，再次按下“完成”按钮部分错误文件即可恢复"
End Sub

Private Sub CommandButton13_Click()
MsgBox "请将文件移植至C:\Program Files (x86)\WPE，在桌面创建图标，确保F盘可正常写入、读出"
End Sub

Private Sub CommandButton2_Click()
Open "F:\WPE Files\aa.txt" For Input As #1
Input #1, a
Close #1
Clipboard.Clear
Clipboard.SetText TextBox4.Text
If a = 0 Then GoTo 1
MsgBox "复制成功！"
1 End Sub

Private Sub CommandButton4_Click()
Open "F:\WPE Files\aa.txt" For Input As #1
Input #1, a
Close #1
Clipboard.Clear
Clipboard.SetText TextBox6.Text
MsgBox "复制成功！"
1 End Sub

Private Sub CommandButton5_Click()
Open "F:\WPE Files\aa.txt" For Input As #1
Input #1, a
Close #1
Clipboard.Clear
Clipboard.SetText TextBox5.Text
If a = 0 Then GoTo 1
MsgBox "复制成功！"
1 End Sub

Private Sub CommandButton6_Click()
Open "F:\WPE Files\aa.txt" For Input As #1
Input #1, a
Close #1
Clipboard.Clear
Clipboard.SetText TextBox7.Text
If a = 0 Then GoTo 1
MsgBox "复制成功！"
1 End Sub

Private Sub CommandButton7_Click()
UserForm2.Show
End Sub

Private Sub CommandButton8_Click()
Open "F:\WPE Files\times.bdf" For Input As #88
Input #88, ac
Close #88
ac = ac + 1
Open "F:\WPE Files\times.bdf" For Output As #55
Print #55, ac
Close #55

Open "F:\WPE Files\number.txt" For Input As #1
Input #1, a
Close #1
TextBox1.Text = a
TextBox2.SetFocus
End Sub

Private Sub CommandButton9_Click()
Open "F:\WPE Files\times.bdf" For Input As #88
Input #88, ac
Close #88
ac = ac + 1
Open "F:\WPE Files\times.bdf" For Output As #45
Print #45, ac
Close #45
Open "F:\WPE Files\note1.txt" For Input As #1
Input #1, a
Close #1
Open "F:\WPE Files\note2.txt" For Input As #2
Input #2, b
Close #2
Open "F:\WPE Files\note3.txt" For Input As #3
Input #3, c
Close #3
If a = "" Then
MsgBox ("你的记事本还没有内容呢")
ElseIf b = "" Or c = "" Then
If MsgBox("记事本内容不完整，是否清空？", vbOKCancel) = vbOK Then GoTo 1 Else GoTo 300
Else
TextBox1.Text = b
TextBox2.Text = c
TextBox3.Text = d
End If
GoTo 300
1 a = ""
b = ""
c = ""
d = ""
Open "F:\WPE Files\note1.txt" For Output As #1
Write #1, a
Close #1
Open "F:\WPE Files\note2.txt" For Output As #1
Write #1, b
Close #1
Open "F:\WPE Files\note3.txt" For Output As #1
Write #1, c
Close #1
Open "F:\WPE Files\note4.txt" For Output As #1
Write #1, d
Close #1
MsgBox "已清空"
300 End Sub

Private Sub OptionButton10_Click()
Open "F:\WPE Files\times.bdf" For Input As #88
Input #88, ac
Close #88
ac = ac + 1
Open "F:\WPE Files\times.bdf" For Output As #45
Print #45, ac
Close #45
TextBox2.Text = "语文小作文"
End Sub

Private Sub OptionButton11_Click()
Open "F:\WPE Files\times.bdf" For Input As #88
Input #88, ac
Close #88
ac = ac + 1
Open "F:\WPE Files\times.bdf" For Output As #45
Print #45, ac
Close #45
TextBox2.Text = "英语笔记纠错"
TextBox3.SetFocus
End Sub

Private Sub OptionButton12_Click()
Open "F:\WPE Files\times.bdf" For Input As #88
Input #88, ac
Close #88
ac = ac + 1
Open "F:\WPE Files\times.bdf" For Output As #45
Print #45, ac
Close #45
TextBox2.Text = "语文大作文"
TextBox3.SetFocus
End Sub

Private Sub OptionButton13_Click()
Open "F:\WPE Files\times.bdf" For Input As #88
Input #88, ac
Close #88
ac = ac + 1
Open "F:\WPE Files\times.bdf" For Output As #45
Print #45, ac
Close #45
TextBox2.Text = "英语笔记讨论"
TextBox3.SetFocus
End Sub

Private Sub OptionButton14_Click()
Open "F:\WPE Files\times.bdf" For Input As #88
Input #88, ac
Close #88
ac = ac + 1
Open "F:\WPE Files\times.bdf" For Output As #45
Print #45, ac
Close #45
TextBox2.Text = "英语题讨论"
TextBox3.SetFocus
End Sub

Private Sub OptionButton15_Click()
Open "F:\WPE Files\times.bdf" For Input As #88
Input #88, ac
Close #88
ac = ac + 1
Open "F:\WPE Files\times.bdf" For Output As #45
Print #45, ac
Close #45
TextBox2.Text = "语文题讨论"
TextBox3.SetFocus
End Sub

Private Sub OptionButton20_Click()
TextBox7.Text = "写得不错嘛"
Open "F:\WPE Files\times.bdf" For Input As #88
Input #88, ac
Close #88
ac = ac + 1
Open "F:\WPE Files\times.bdf" For Output As #45
Print #45, ac
Close #45
End Sub

Private Sub OptionButton19_Click()
TextBox7.Text = "生动形象地写出了"
Open "F:\WPE Files\times.bdf" For Input As #88
Input #88, ac
Close #88
ac = ac + 1
Open "F:\WPE Files\times.bdf" For Output As #45
Print #45, ac
Close #45
End Sub

Private Sub OptionButton17_Click()
TextBox7.Text = "继续加油！"
Open "F:\WPE Files\times.bdf" For Input As #88
Input #88, ac
Close #88
ac = ac + 1
Open "F:\WPE Files\times.bdf" For Output As #45
Print #45, ac
Close #45
End Sub

Private Sub OptionButton18_Click()
TextBox7.Text = "表达了"
Open "F:\WPE Files\times.bdf" For Input As #88
Input #88, ac
Close #88
ac = ac + 1
Open "F:\WPE Files\times.bdf" For Output As #45
Print #45, ac
Close #45
End Sub

Private Sub OptionButton16_Click()
TextBox7.Text = "我没有看到你的呢"
Open "F:\WPE Files\times.bdf" For Input As #88
Input #88, ac
Close #88
ac = ac + 1
Open "F:\WPE Files\times.bdf" For Output As #45
Print #45, ac
Close #45
End Sub

Private Sub OptionButton7_Click()
Open "F:\WPE Files\times.bdf" For Input As #88
Input #88, ac
Close #88
ac = ac + 1
Open "F:\WPE Files\times.bdf" For Output As #45
Print #45, ac
Close #45
TextBox7.Text = "大家也快快来发言吧"
End Sub

Private Sub OptionButton8_Click()
TextBox3.Text = "金美杉、刘子敬、迟舒月、王 尊、赵静涛"
TextBox7.SetFocus
End Sub

Private Sub OptionButton9_Click()
TextBox3.Text = "金圣皓、向美霖、王艳霞、李俊漾"
TextBox7.SetFocus
End Sub

Private Sub UserForm_Initialize()
If Dir("F:\WPE Files\note1.text") <> "" Then
Open "F:\WPE Files\note1.txt" For Input As #1
Input #1, a
Close #1
Else
GoTo 10
End If
If a <> Date Then
a = ""
b = ""
c = ""
d = ""
Open "F:\WPE Files\note1.txt" For Output As #1
Write #1, a
Close #1
Open "F:\WPE Files\note2.txt" For Output As #1
Write #1, b
Close #1
Open "F:\WPE Files\note3.txt" For Output As #1
Write #1, c
Close #1
Open "F:\WPE Files\note4.txt" For Output As #1
Write #1, d
Close #1
Else
GoTo 10
End If
10 End Sub
