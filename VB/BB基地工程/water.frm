VERSION 5.00
Begin VB.Form water 
   BackColor       =   &H00FFFF00&
   Caption         =   "浇水"
   ClientHeight    =   3810
   ClientLeft      =   5835
   ClientTop       =   4170
   ClientWidth     =   4680
   Icon            =   "water.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   4680
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   3120
      Picture         =   "water.frx":324A
      ScaleHeight     =   1215
      ScaleWidth      =   1335
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   2640
      Picture         =   "water.frx":3F53
      ScaleHeight     =   1215
      ScaleWidth      =   1335
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   2160
      Picture         =   "water.frx":4C5C
      ScaleHeight     =   1215
      ScaleWidth      =   1335
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   1680
      Picture         =   "water.frx":5965
      ScaleHeight     =   1215
      ScaleWidth      =   1335
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   120
      Picture         =   "water.frx":666E
      ScaleHeight     =   2295
      ScaleWidth      =   2775
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00F8D047&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2400
         Shape           =   3  'Circle
         Top             =   1560
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00F8D047&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2400
         Shape           =   3  'Circle
         Top             =   1320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00F8D047&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2400
         Shape           =   3  'Circle
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "双击我浇水↖"
         Height          =   495
         Left            =   1920
         TabIndex        =   5
         Top             =   1800
         Width           =   855
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      Caption         =   "种子长大444"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      Caption         =   "已浇水！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E335DF&
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "water"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
main.Show
End Sub

Private Sub Label3_Click()
Label3.Caption = "种子  `   "
End Sub

Private Sub Picture1_DblClick()
Label3.Caption = "种子长大3个"
Label2.Visible = True
Label3.Visible = True
a = 1
b = Date
Open "F:\BBSeed Files\wr.bdf" For Output As #2
Print #2, a
Close #2
Open "F:\BBSeed Files\date.bdf" For Output As #3
Print #3, b
Close #3
Open "F:\BBSeed Files\nr.bdf" For Input As #4
Input #4, n
Close #4
n = n + 3
Open "F:\BBSeed Files\nr.bdf" For Output As #5
Print #5, n
Close #5
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.Visible = False
Shape2.Visible = False
Shape3.Visible = False
Shape1.Visible = True
Shape2.Visible = True
Shape3.Visible = True
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.Visible = True
Shape2.Visible = True
Shape3.Visible = True
Shape1.Visible = False
Shape2.Visible = False
Shape3.Visible = False
End Sub
