VERSION 5.00
Begin VB.Form FrmMsgAsk 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1236
   ClientLeft      =   9846
   ClientTop       =   1308
   ClientWidth     =   5028
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1236
   ScaleWidth      =   5028
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "华文仿宋"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   426
      Left            =   2820
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   780
      Width           =   1026
   End
   Begin VB.CommandButton CmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "华文仿宋"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   426
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   780
      Width           =   1026
   End
   Begin VB.Timer Tmr 
      Interval        =   2000
      Left            =   3300
      Top             =   60
   End
   Begin VB.Label LblText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "abc课程表？"
      BeginProperty Font 
         Name            =   "华文仿宋"
         Size            =   26.1
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2778
   End
End
Attribute VB_Name = "FrmMsgAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
Private Const SWP_NOSIZE& = &H1
Private Const SWP_NOMOVE& = &H2

'''''''''''''
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2
Const LWA_COLORKEY = &H1


Private Sub CmdCancel_Click()
MsgYN = False
Unload Me
End Sub

Private Sub CmdOK_Click()
MsgYN = True
Unload Me
End Sub

Private Sub Form_Load()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
'''''''''''''''
Dim rtn As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn

End Sub




Private Sub Form_Paint()
FrmMsg.Width = LblText.Width
EnMiddle Me
End Sub

Private Sub Tmr_Timer()
Unload Me
End Sub
