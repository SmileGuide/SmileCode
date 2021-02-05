VERSION 5.00
Begin VB.Form main 
   BackColor       =   &H00C0C000&
   Caption         =   "养料工厂"
   ClientHeight    =   2730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5700
   FillColor       =   &H00FFFF00&
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   5700
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "更多..."
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "重置"
      Height          =   735
      Left            =   4680
      TabIndex        =   5
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FAB4FA&
      Caption         =   "完成"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "提示"
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "kg"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "要生产多少克养料？"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "一千克养料可以让一个种子长大"
End Sub

Private Sub Command2_Click()
n = Val(Text1.Text)
If n = 0 Then
MsgBox "请输入正确的数字！"
Text1.Text = ""
GoTo 99
End If
b = 1
j = Now
Open "F:\BBSeed Files\fd.bdf" For Input As #1
Input #1, a
Close #1
Open "F:\BBSeed Files\fd.bdf" For Output As #2
Print #2, n
Close #2
Open "F:\BBSeed Files\foodstate.bdf" For Output As #3
Print #3, b
Close #3
Open "F:\BBSeed Files\foodtime.bdf" For Output As #4
Print #4, j
Close #4
MsgBox "生产完成！"
Text1.Text = ""
99 End Sub

Private Sub Command3_Click()
a = 0
b = 1
Open "F:\BBSeed Files\fd.bdf" For Output As #1
Print #1, a
Close #1
Open "F:\BBSeed Files\foodstate.bdf" For Output As #2
Print #2, b
Close #2
MsgBox "重置成功，现在小蛋蛋有0个抱抱种子"
End Sub



Private Sub Command5_Click()
more.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
