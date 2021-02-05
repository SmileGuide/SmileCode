VERSION 5.00
Begin VB.Form hope 
   Caption         =   "愿望提交"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3540
   Icon            =   "hope.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3510
   ScaleWidth      =   3540
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "提交"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "愿望解释（不限字数）："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "愿望名称（1-10字）："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "hope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Len(Text1.Text) > 10 Or Text2.Text = "" Then
MsgBox " 输入的愿望名称不合规 "
Else
a = Text1.Text
b = Text2.Text
c = 0
Open "F:\BBSeed Files\hopec.bdf" For Output As #1
Print #1, a
Close #1
Open "F:\BBSeed Files\hopereason.bdf" For Output As #2
Print #2, b
Close #2
Open "F:\BBSeed Files\hopestate.bdf" For Output As #3
Print #3, c
Close #3
MsgBox "提交成功！"
hope.Hide
main.Label8.ForeColor = vbBlack
End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
main.Show
End Sub
