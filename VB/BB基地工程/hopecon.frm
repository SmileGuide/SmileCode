VERSION 5.00
Begin VB.Form hopecon 
   Caption         =   "愿望"
   ClientHeight    =   2610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4170
   Icon            =   "hopecon.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2610
   ScaleWidth      =   4170
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "打印奖状"
      Height          =   615
      Left            =   0
      MaskColor       =   &H00F8D047&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "打印承诺书"
      Height          =   615
      Left            =   0
      MaskColor       =   &H00F8D047&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   0
      Picture         =   "hopecon.frx":324A
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "愿望内容"
      BeginProperty Font 
         Name            =   "华文彩云"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Kunstler Script"
         Size            =   42
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   855
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "hopecon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
hopeconpromprint.PrintForm
End Sub

Private Sub Command2_Click()
hopeconpromprint.PrintForm
End Sub

Private Sub Form_Load()
Open "F:\BBSeed Files\hopec.bdf" For Input As #1
Input #1, a
Close #1
Label1.Caption = a
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.Show
End Sub
