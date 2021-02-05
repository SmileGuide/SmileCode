VERSION 5.00
Begin VB.Form check1 
   Caption         =   "身份验证"
   ClientHeight    =   1350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4440
   Icon            =   "check1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1350
   ScaleWidth      =   4440
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "A关机 B画板 C剪切 D写字板"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "clipboard对象与以下哪个内容有关？"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "check1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()
If LCase(Text1.Text) = "c" Then
check1.Hide
check2.Show
Else
MsgBox "错了"
Unload Me
End If
End Sub
