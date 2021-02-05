VERSION 5.00
Begin VB.Form check2 
   Caption         =   "身份验证"
   ClientHeight    =   1440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4455
   Icon            =   "check2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1440
   ScaleWidth      =   4455
   StartUpPosition =   2  '屏幕中心
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
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "请写出一个给变量a赋值VBA窗体第一文本框的内容的语句"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "check2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()
If Len(Text1.Text) = 15 Then
If LCase(Text1.Text) = "a=textbox1.text" Then
check2.Hide
main.Show
Else
Unload check1
Unload check2
Unload check0
Unload login
MsgBox "错了"
End If
End If
End Sub
