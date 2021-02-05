VERSION 5.00
Begin VB.Form check0 
   Caption         =   "请把下面的语言描述转换成盘符和路径"
   ClientHeight    =   1215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5160
   Icon            =   "findlogin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1215
   ScaleWidth      =   5160
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
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "D盘的22文件夹里的ss文件夹里的VB工程文件（10进制）QQ"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "check0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()
If Len(Text1.Text) = 15 Then
If LCase(Text1.Text) = "d:\22\ss\qq.vbp" Then
check0.Hide
check1.Show
Else
Unload check1
Unload check2
Unload check0
Unload login
MsgBox "错了"
End If
End If
End Sub
