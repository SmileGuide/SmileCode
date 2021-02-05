VERSION 5.00
Begin VB.Form food 
   BackColor       =   &H000080FF&
   Caption         =   "抱抱养料"
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6285
   Icon            =   "food.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   6285
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture5 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   4080
      Picture         =   "food.frx":0CCA
      ScaleHeight     =   1215
      ScaleWidth      =   1335
      TabIndex        =   4
      Top             =   3840
      Width           =   1335
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   3600
      Picture         =   "food.frx":19D3
      ScaleHeight     =   1215
      ScaleWidth      =   1335
      TabIndex        =   3
      Top             =   3840
      Width           =   1335
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   3120
      Picture         =   "food.frx":26DC
      ScaleHeight     =   1215
      ScaleWidth      =   1335
      TabIndex        =   2
      Top             =   3840
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   2640
      Picture         =   "food.frx":33E5
      ScaleHeight     =   1215
      ScaleWidth      =   1335
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      Height          =   3855
      Left            =   0
      Picture         =   "food.frx":40EE
      ScaleHeight     =   3795
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         Caption         =   "双击我施肥↑"
         Height          =   495
         Left            =   1200
         TabIndex        =   5
         Top             =   3120
         Width           =   855
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0080FFFF&
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2760
         Shape           =   3  'Circle
         Top             =   3240
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0080FFFF&
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2760
         Shape           =   3  'Circle
         Top             =   2880
         Width           =   495
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0080FFFF&
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2760
         Shape           =   3  'Circle
         Top             =   2520
         Width           =   495
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   3600
      TabIndex        =   7
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Caption         =   "已施肥！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   3600
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   3135
   End
End
Attribute VB_Name = "food"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Open "F:\BBSeed Files\fd.bdf" For Input As #5
Input #5, e
Close #5
Label3.Caption = "小抱给了你" & e & "kg养料，可以让" & e & "个抱抱种子长大！"
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.Show
End Sub
Private Sub Picture1_DblClick()
If Label2.Visible = False Then
Label2.Visible = True
a = 0
Open "F:\BBSeed Files\foodstate.bdf" For Output As #3
Print #3, a
Close #3
Open "F:\BBSeed Files\fd.bdf" For Input As #5
Input #5, e
Close #5
Open "F:\BBSeed Files\nr.bdf" For Input As #3
Input #3, c
Close #3
f = e + c
Open "F:\BBSeed Files\nr.bdf" For Output As #3
Print #3, f
Close #3
main.Text1.Text = f
End If
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


