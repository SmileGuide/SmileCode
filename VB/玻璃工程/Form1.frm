VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Tim 
      Interval        =   1000
      Left            =   150
      Top             =   195
   End
   Begin VB.Label Label1 
      Caption         =   "已锁定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   1275
      TabIndex        =   0
      Top             =   1305
      Visible         =   0   'False
      Width           =   1755
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim odr As Long
Dim cmo As Boolean







Private Sub Form_KeyPress(KeyAscii As Integer)
If Ascii = 27 Or KeyAscii = 8 Then Unload Me
End Sub

Private Sub Form_Load()
Form1.Top = Mom.Top
Form1.Left = Mom.Width
cmo = True
odr = 1
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
'''''''''''''''
Dim rtn As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, vbWhite, 150, LWA_ALPHA
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If cmo Then
Dim ReturnVal As Long
X = ReleaseCapture()
ReturnVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End If
End Sub



Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And Int(odr / 2) = odr / 2 Then
cmo = True
odr = odr + 1
Label1.Caption = "已解锁"
ElseIf Button = 2 And Int(odr / 2) <> odr / 2 Then
Label1.Caption = "已锁定"
cmo = False
odr = odr + 1
End If
Tim.Enabled = True
Label1.Visible = True

''
'If down.Left <> mom.Left - 585 Then down.Left = mom.Left - 585
'If down.Top <> mom.Top Then down.Top = mom.Top
'If mom.Shapek.Width <> Int(HS.Value / 10) Then mom.Shapek.Width = Int(HS.Value / 10)
'If mom.Shapek.Height <> Int(VS.Value / 10) Then mom.Shapek.Height = Int(VS.Value / 10)
'On Error GoTo 66
'If slow.Value = 1 And (mom.Shapek.BackColor <> ClSet(od + 1)) And od < 4 Then
'mom.Shapek.BackColor = ClSet(od + 1)
'End If
'66 If Err.Number = 9 And slow.Value = 1 Then
'mom.Shapek.BackColor = ClSet(0)
'End If
End Sub



Private Sub Tim_Timer()
Label1.Visible = False
Tim.Enabled = False
End Sub
