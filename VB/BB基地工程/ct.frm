VERSION 5.00
Begin VB.Form ct 
   Caption         =   "愿望修改"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3450
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   3450
   StartUpPosition =   3  '窗口缺省
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
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "完成"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   2415
   End
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
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1680
      Width           =   3015
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
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   2655
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
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   2895
   End
End
Attribute VB_Name = "ct"
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
Open "F:\BBSeed Files\hopec.bdf" For Output As #1
Print #1, a
Close #1
Open "F:\BBSeed Files\hopereason.bdf" For Output As #2
Print #2, b
Close #2
MsgBox "修改成功！"
ct.Hide
End If
End Sub


Private Sub Form_Load()
Open "F:\BBSeed Files\hopec.bdf" For Input As #1
Input #1, a
Close #1
Open "F:\BBSeed Files\hopereason.bdf" For Input As #2
Input #2, b
Text1.Text = a
Text2.Text = b
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.Show
End Sub
