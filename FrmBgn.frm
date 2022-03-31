VERSION 5.00
Begin VB.Form FrmBgn 
   BackColor       =   &H00FDEEBF&
   ClientHeight    =   936
   ClientLeft      =   24
   ClientTop       =   24
   ClientWidth     =   6972
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   936
   ScaleWidth      =   6972
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Tmr 
      Interval        =   100
      Left            =   180
      Top             =   480
   End
   Begin VB.Label LblText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "启动中..."
      BeginProperty Font 
         Name            =   "华文仿宋"
         Size            =   26.1
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   588
      Left            =   2340
      TabIndex        =   0
      Top             =   120
      Width           =   1944
   End
End
Attribute VB_Name = "FrmBgn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Dim Doc As Integer
Private Sub Form_Paint()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
EnMiddle Me

        Dim rtn As Long
        rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
        rtn = rtn Or Me.BackColor
        SetWindowLong hwnd, GWL_EXSTYLE, rtn
        SetLayeredWindowAttributes hwnd, 0, 210, LWA_ALPHA
End Sub

Private Sub Tmr_Timer()
FrmWel.Show
If FrmWelLoad = True Then Unload Me: Set FrmBgn = Nothing
End Sub
