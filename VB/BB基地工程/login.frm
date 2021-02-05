VERSION 5.00
Begin VB.Form login 
   BackColor       =   &H00808080&
   Caption         =   "抱抱养料工厂"
   ClientHeight    =   1725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4230
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   4230
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "插进去了"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   360
      Picture         =   "login.frx":0CCA
      ScaleHeight     =   1575
      ScaleWidth      =   1095
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "请插入钥匙"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
login.Hide
check0.Show
End Sub

