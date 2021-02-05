VERSION 5.00
Begin VB.Form mini 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   585
   ClientLeft      =   9540
   ClientTop       =   2955
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      Height          =   555
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   7290
      TabIndex        =   0
      Top             =   0
      Width           =   7350
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   1950
         Top             =   15
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "玻璃筐已最小化到右下角系统托盘,双击图标召回"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   525
         Left            =   480
         TabIndex        =   1
         Top             =   60
         Width           =   6870
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   0
         Picture         =   "msgmini.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
End
Attribute VB_Name = "mini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
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


Private Sub Form_Load()
mini.Move (Screen.Width - mini.Width) / 2, (Screen.Height - mini.Height) / 2
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
'''''''''''''''
Dim rtn As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, vbWhite, 150, LWA_ALPHA
End Sub


Private Sub Timer1_Timer()
Timer1.Enabled = False
Unload mini
End Sub
