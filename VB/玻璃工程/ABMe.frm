VERSION 5.00
Begin VB.Form ABMe 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于-玻璃"
   ClientHeight    =   1050
   ClientLeft      =   21
   ClientTop       =   336
   ClientWidth     =   4480
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   4480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      ForeColor       =   &H00FF0000&
      Height          =   763
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   4480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "→戳 访问作者个人网站 戳←"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   259
      Left            =   0
      TabIndex        =   1
      Top             =   756
      Width           =   4480
   End
End
Attribute VB_Name = "ABMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
''''''''''''''''''''
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE& = &H1
Private Const SWP_NOMOVE& = &H2

'''''''''''''
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2
Const LWA_COLORKEY = &H1
''''''''''''''

Private Sub Form_Load()
SetWindowPos ABMe.hwnd, HWND_TOPMOST&, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
Text1.Text = "玻璃 " & Trim$(Str$(App.Major)) & "." & Format$(App.Minor, "##00") & "." & Format$(App.Revision, "0000") & " ，功能强大，由S.G.G.制作。"
End Sub

Private Sub Label1_Click()
Shell "explorer https://blog.csdn.net/edwfvqhewjyh"
End Sub


Private Sub Text1_DblClick()
Shell "explorer https://blog.csdn.net/edwfvqhewjyh"
End Sub
