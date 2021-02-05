VERSION 5.00
Begin VB.Form Abme 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关于-剪切板存图助手"
   ClientHeight    =   1085
   ClientLeft      =   15
   ClientTop       =   350
   ClientWidth     =   3625
   ControlBox      =   0   'False
   Icon            =   "Abme.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1085
   ScaleWidth      =   3625
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label LblEr 
      BackColor       =   &H00C0FFFF&
      Caption         =   "作者："
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   245
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   545
   End
   Begin VB.Label LblName 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Smile Guide"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   185
      Left            =   720
      TabIndex        =   2
      Top             =   0
      Width           =   1145
   End
   Begin VB.Label LblFit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "         适用场合：截图后粘贴保存文件嫌太麻烦时"
      BeginProperty Font 
         Name            =   "华文行楷"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   485
      Left            =   0
      TabIndex        =   1
      Top             =   540
      Width           =   3605
   End
   Begin VB.Label LblCap 
      BackColor       =   &H00C0FFFF&
      Caption         =   "实时从剪贴板中获取图片并存储到文件。"
      BeginProperty Font 
         Name            =   "华文行楷"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   185
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   3605
   End
End
Attribute VB_Name = "Abme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_LostFocus()
Unload Abme
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If LblName.ForeColor <> vbBlack Then LblName.ForeColor = vbBlack
Screen.MousePointer = 1
End Sub

Private Sub LblCap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = 1
End Sub

Private Sub LblEr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = 1
End Sub

Private Sub LblName_Click()
Shell "explorer https://blog.csdn.net/edwfvqhewjyh"
End Sub

Private Sub LblName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblName.ForeColor = vbRed
Screen.MousePointer = 1
End Sub
