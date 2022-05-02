VERSION 5.00
Begin VB.Form Abme 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关于-微笑课程表"
   ClientHeight    =   1866
   ClientLeft      =   9450
   ClientTop       =   -3450
   ClientWidth     =   3612
   ControlBox      =   0   'False
   Icon            =   "Abme.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1866
   ScaleWidth      =   3612
   StartUpPosition =   1  '所有者中心
   Begin VB.Label LblEm 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "邮箱："
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   244
      Left            =   0
      TabIndex        =   7
      Top             =   540
      Width           =   544
   End
   Begin VB.Label LblEMail 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SmileGuide@163.com"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   186
      Left            =   780
      TabIndex        =   6
      Top             =   540
      Width           =   1626
   End
   Begin VB.Label LblCap 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "介绍：为用户提供一个专业实用、个性化的效率应用"
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
      Height          =   545
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   3605
   End
   Begin VB.Label LblAge 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "14岁"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   184
      Left            =   780
      TabIndex        =   5
      Top             =   300
      Width           =   1144
   End
   Begin VB.Label LblAg 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "年龄："
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   244
      Left            =   0
      TabIndex        =   4
      Top             =   300
      Width           =   544
   End
   Begin VB.Label LblEr 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "开发者："
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   244
      Left            =   0
      TabIndex        =   3
      Top             =   60
      Width           =   784
   End
   Begin VB.Label LblName 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
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
      Height          =   184
      Left            =   780
      TabIndex        =   2
      Top             =   60
      Width           =   1144
   End
   Begin VB.Label LblFit 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "适用年龄：3+"
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
      Top             =   1380
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
