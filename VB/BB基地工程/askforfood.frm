VERSION 5.00
Begin VB.Form askforfood 
   BackColor       =   &H00FFFF80&
   Caption         =   "讨养料"
   ClientHeight    =   4140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "askforfood.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FAB4FA&
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
      Height          =   735
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "随机备注"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1680
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF80&
      Caption         =   "备注"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      Caption         =   "kg"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   "讨养料数量"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "askforfood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Randomize
a = Int(6 * Rnd() + 1)
If a = 1 Then
Text2.Text = "小抱抱，我的抱抱种子有些饿了，给我些养料吧(*^▽^*)Thanks"
ElseIf a = 2 Then
Text2.Text = "小抱抱，喜欢你呦￣ω￣  =~(＾◇^)/你真好"
ElseIf a = 3 Then
Text2.Text = "快给我种子吧(￣.￣)  ヾ(*′ー`*)゛谢谢 "
ElseIf a = 4 Then
Text2.Text = "我想要种子了(￣３￣)a"
ElseIf a = 5 Then
Text2.Text = "给我抱抱种子,我请你吃好吃的(灬°ω°灬) "
Else
Text2.Text = "啊！小抱抱的养料是我种子的生命之源！多么善良，亲爱的小抱！你赋予了我的抱抱种子以生命!好感动!ヾ(o′▽`o)ノ°° .°谢谢° .°"
End If
End Sub

Private Sub Command2_Click()
a = Val(Text1.Text)
b = Text2.Text
c = 2
Open "F:\BBSeed Files\askn.bdf" For Output As #1
Print #1, a
Close #1
Open "F:\BBSeed Files\askc.bdf" For Output As #2
Print #2, b
Close #2
Open "F:\BBSeed Files\foodstate.bdf" For Output As #3
Print #3, c
Close #3
MsgBox "提交成功！"
 Unload askforfood
 Unload nofood
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.Show
End Sub
