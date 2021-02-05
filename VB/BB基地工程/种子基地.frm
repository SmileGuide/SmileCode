VERSION 5.00
Begin VB.Form makebad 
   Caption         =   "Form2"
   ClientHeight    =   2610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3375
   Icon            =   "种子基地.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2610
   ScaleWidth      =   3375
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "打印奖状"
      Height          =   615
      Left            =   120
      MaskColor       =   &H00F8D047&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "打印承诺书"
      Height          =   615
      Left            =   120
      MaskColor       =   &H00F8D047&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFF00&
      Height          =   1095
      Left            =   120
      Picture         =   "种子基地.frx":324A
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   0
      Top             =   120
      Width           =   1095
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
      Left            =   1440
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "叠被"
      BeginProperty Font 
         Name            =   "华文彩云"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "makebad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
makebadpromprint.Show
makebadpromprint.PrintForm
End Sub

Private Sub Command2_Click()
makebadcmprint.Show
makebadcmprint.PrintForm
End Sub
