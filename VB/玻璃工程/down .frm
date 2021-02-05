VERSION 5.00
Begin VB.Form down 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   495
   ClientLeft      =   2115
   ClientTop       =   1710
   ClientWidth     =   495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "down .frx":0000
   ScaleHeight     =   495
   ScaleWidth      =   495
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   75
      Top             =   -30
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "再次点击结束程序"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   660
      TabIndex        =   0
      Top             =   105
      Width           =   1980
   End
End
Attribute VB_Name = "down"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim candown As Integer
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

Private Sub Form_Click()
down.Width = 2775
candown = candown + 1
down.Left = down.Left - 2775 + 585
Timer1.Enabled = True
Mom.Timerre.Enabled = False
If candown = 2 Then End
End Sub

Private Sub Form_Load()
If mor Then
down.Left = Mom.Left - 585
down.Top = Mom.Top
Dim rtn As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, vbWhite, 150, LWA_ALPHA
End If
End Sub

Private Sub Form_Paint()
If Not mor Then
down.Hide
down.Left = Mom.Left - 585
down.Top = Mom.Top
Dim rtn As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, vbWhite, 150, LWA_ALPHA
End If
End Sub

Private Sub Timer1_Timer()
down.Width = 585
down.Left = down.Left - 2775 + 585
Timer1.Enabled = False
Mom.Timerre.Enabled = True
candown = False
End Sub
