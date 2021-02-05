VERSION 5.00
Begin VB.Form 种子操作台 
   Caption         =   "种子操作台"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6315
   Icon            =   "ctrl.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   6315
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   1080
      TabIndex        =   15
      Top             =   4560
      Width           =   2895
      Begin VB.OptionButton Option10 
         Caption         =   "已浇水"
         Height          =   495
         Left            =   1320
         TabIndex        =   17
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton Option9 
         Caption         =   "未浇水"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   360
      TabIndex        =   9
      Top             =   2400
      Width           =   5415
      Begin VB.OptionButton Option5 
         Caption         =   "未有养料申请"
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option6 
         Caption         =   "养料申请审核中"
         Height          =   735
         Left            =   1320
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option7 
         Caption         =   "养料给予"
         Height          =   615
         Left            =   2760
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option8 
         Caption         =   "养料驳回"
         Height          =   615
         Left            =   3960
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   5775
      Begin VB.OptionButton Option1 
         Caption         =   "未有愿望"
         Height          =   735
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "愿望审核中"
         Height          =   735
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "愿望实现"
         Height          =   615
         Left            =   3000
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "愿望驳回"
         Height          =   615
         Left            =   4320
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "未/已浇水"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "种子数量"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "养料状态"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "愿望状态"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "种子操作台"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Open "F:\BBSeed Files\nr.bdf" For Input As #1
Input #1, a
Close #1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Open "F:\BBSeed Files\nr.bdf" For Input As #1
Input #1, a
Close #1
End Sub

Private Sub Option1_Click()
a = 3
Open "F:\BBSeed Files\hopestate.bdf" For Output As #1
Print #1, a
Close #1
End Sub

Private Sub Option10_Click()
a = 1
Open "F:\BBSeed Files\wr.bdf" For Output As #1
Print #1, a
Close #1
End Sub

Private Sub Option2_Click()
a = 0
Open "F:\BBSeed Files\hopestate.bdf" For Output As #1
Print #1, a
Close #1
End Sub

Private Sub Option3_Click()
a = 1
Open "F:\BBSeed Files\hopestate.bdf" For Output As #1
Print #1, a
Close #1
End Sub

Private Sub Option4_Click()
a = 2
Open "F:\BBSeed Files\hopestate.bdf" For Output As #1
Print #1, a
Close #1
End Sub

Private Sub Option5_Click()
a = 0
Open "F:\BBSeed Files\foodstate.bdf" For Output As #1
Print #1, a
Close #1
End Sub

Private Sub Option6_Click()
a = 2
Open "F:\BBSeed Files\hopestate.bdf" For Output As #1
Print #1, a
Close #1
End Sub

Private Sub Option7_Click()
a = 1
Open "F:\BBSeed Files\hopestate.bdf" For Output As #1
Print #1, a
Close #1
End Sub

Private Sub Option8_Click()
a = 3
Open "F:\BBSeed Files\hopestate.bdf" For Output As #1
Print #1, a
Close #1
End Sub

Private Sub Option9_Click()
a = 0
Open "F:\BBSeed Files\wr.bdf" For Output As #1
Print #1, a
Close #1
End Sub

Private Sub Text1_Change()
a = Text1.Text
Open "F:\BBSeed Files\nr.bdf" For Output As #1
Print #1, a
Close #1
End Sub

