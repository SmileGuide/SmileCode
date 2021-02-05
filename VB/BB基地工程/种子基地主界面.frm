VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form main 
   BackColor       =   &H00C0FFFF&
   Caption         =   "抱抱种子基地"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9090
   FillColor       =   &H00004000&
   Icon            =   "种子基地主界面.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   9090
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "领取"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "领取"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C000&
      Caption         =   "给种子浇水>"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      MaskColor       =   &H00C0C000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FF80&
      Caption         =   "去赚抱抱种子>"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H000080FF&
      Caption         =   "抱抱养料>"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   960
      TabIndex        =   5
      Top             =   960
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   1
      Max             =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   615
      Left            =   960
      TabIndex        =   7
      Top             =   2760
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   1
      Max             =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C4F2E9&
      Caption         =   "消息>"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "修改即将被审核的消息"
      ForeColor       =   &H8000000A&
      Height          =   495
      Left            =   7560
      TabIndex        =   16
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   7920
      X2              =   8160
      Y1              =   3240
      Y2              =   3720
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "当前"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "个抱抱种子长大"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   3600
      TabIndex        =   13
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   3720
      TabIndex        =   12
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "叠被"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   6120
      TabIndex        =   11
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   3840
      TabIndex        =   10
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "愿望"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   6120
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H0097FBCB&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "今天未浇水"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   615
      Left            =   3120
      TabIndex        =   8
      Top             =   4440
      Width           =   1215
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If MsgBox("仅限领取一次，要继续吗？", vbOKCancel) = vbOK Then
If Text1.Text >= 10 Then
makebad.Show
GoTo 22
Else
MsgBox "你还没有让足够的抱抱种子长大"
GoTo 33
End If
End If
22 cc = Val(Text1.Text) - 10
Open "F:\BBSeed Files\nr.bdf" For Output As #1
Print #1, cc
Close #1
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
Label8.ForeColor = QBColor(8)
Else
Label8.ForeColor = vbBlack
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
Label3.Caption = a11
ProgressBar2.Value = a2 / 100
Label5.Caption = a12
If c = 0 Then
Label7.Caption = "今天未浇水"
Else
Label7.Caption = "今天已浇水"
End If
33 End Sub

Private Sub Command2_Click()
If MsgBox("仅限领取一次，要继续吗？", vbOKCancel) = vbOK Then
If Val(Text1.Text) < 100 Then
MsgBox "你还没有让足够的抱抱种子长大"
GoTo 3
End If
Open "F:\BBSeed Files\hopestate.bdf" For Input As #1
Input #1, a
Close #1
If a = 3 Then
hope.Show
GoTo 3
ElseIf a = 1 Then
hopecon.Show
GoTo 3
ElseIf a = 2 Then
sorry.Show
GoTo 3
Else
MsgBox "您提交的愿望“" & a & "”已经提交，正在等待小抱审核"
GoTo 3
End If
cc = Val(Text1.Text) - 100
Open "F:\BBSeed Files\nr.bdf" For Output As #1
Print #1, cc
Close #1
Open "F:\BBSeed Files\hope.bdf" For Input As #1
Input #1, a
Close #1
End If
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
Label8.ForeColor = QBColor(8)
Else
Label8.ForeColor = vbBlack
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
Label3.Caption = a11
ProgressBar2.Value = a2 / 100
Label5.Caption = a12
If c = 0 Then
Label7.Caption = "今天未浇水"
Else
Label7.Caption = "今天已浇水"
End If
3 End Sub

Private Sub Command3_Click()
If Label7.Caption = "今天未浇水" Then
water.Show
Unload main
Else
MsgBox "你进天已经浇水了呢"
End If
End Sub

Private Sub Command4_Click()
main.Hide
turn.Show
End Sub

Private Sub Command5_Click()
Call foodp
End Sub

Private Sub Command6_Click()
main.Hide
letter.Show
End Sub

Private Sub Form_Load()
Load Me
Call news
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
Label8.ForeColor = QBColor(8)
Else
Label8.ForeColor = vbBlack
End If
If d <> Date Then
f = 0
Open "F:\BBSeed Files\wr.bdf" For Output As #4
Print #4, f
Close #4
Else
f = 1
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
ElseIf a = 100 Then
a1 = 10
a11 = "100%"
a2 = 100
a12 = "100%"
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
Label3.Caption = a11
ProgressBar2.Value = a2 / 100
Label5.Caption = a12
If c = 0 Then
Label7.Caption = "今天未浇水"
Else
Label7.Caption = "今天已浇水"
End If
End Sub




Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call news
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub Label8_Click()
Open "F:\BBSeed Files\hopestate.bdf" For Input As #1
Input #1, a
Close #1
If a = 0 Then
ct.Show
Else
End If
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Open "F:\BBSeed Files\hopestate.bdf" For Input As #1
Input #1, a
Close #1
If a = 0 Then
Label8.BorderStyle = 0
Else
End If
End Sub

Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Open "F:\BBSeed Files\hopestate.bdf" For Input As #1
Input #1, a
Close #1
If a = 0 Then
Label8.BorderStyle = 1
Else
End If
End Sub
Sub waterp()
If Label7.Caption = "今天已浇水" Then
MsgBox ("你今天已经浇过水了呢")
Else
water.Show
main.Hide
End If
End Sub
