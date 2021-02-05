VERSION 5.00
Begin VB.Form mini 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "白屏"
   ClientHeight    =   3135
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu st 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu se 
         Caption         =   "设置背景颜色"
      End
   End
End
Attribute VB_Name = "mini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bc As Single
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

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo er
st: Open "settings.ini" For Input As #1
Input #1, bc
Close #1
mini.BackColor = bc
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
'''''''''''''''
Dim rtn As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, &HC0FFFF, 150, LWA_ALPHA
Exit Sub
er: Open "settings.ini" For Output As #2
Print #2, vbWhite
Close #2
GoTo st
End Sub



Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu st
End Sub

Private Sub se_Click()
On Error GoTo ed
mini.BackColor = "&h" & InputBox("请输入十六进制颜色码(6位）", "设置")
Open "settings.ini" For Output As #2
Print #2, mini.BackColor
Close #2
ed: End Sub



