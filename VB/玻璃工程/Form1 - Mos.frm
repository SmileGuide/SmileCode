VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FormSP 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3136
   ClientLeft      =   42
   ClientTop       =   42
   ClientWidth     =   4683
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1 - Mos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3136
   ScaleWidth      =   4683
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton More 
      Caption         =   "M"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      ToolTipText     =   "进入堆叠模式"
      Top             =   510
      Width           =   255
   End
   Begin VB.CommandButton Guide 
      Caption         =   "G"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "导航"
      Top             =   510
      Width           =   255
   End
   Begin VB.CommandButton About 
      Caption         =   "A"
      Height          =   255
      Left            =   495
      TabIndex        =   4
      ToolTipText     =   "关于"
      Top             =   255
      Width           =   255
   End
   Begin VB.CommandButton Unl 
      Caption         =   "E"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "关闭"
      Top             =   255
      Width           =   255
   End
   Begin MSComDlg.CommonDialog Cdi 
      Left            =   780
      Top             =   1020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Color           =   16744576
   End
   Begin VB.CommandButton Cl 
      Caption         =   "C"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "设置玻璃颜色"
      Top             =   0
      Width           =   255
   End
   Begin VB.VScrollBar VS 
      Height          =   2835
      LargeChange     =   500
      Left            =   0
      Min             =   3105
      SmallChange     =   100
      TabIndex        =   1
      Top             =   255
      Value           =   3105
      Width           =   255
   End
   Begin VB.HScrollBar HS 
      Height          =   255
      LargeChange     =   500
      Left            =   240
      Min             =   4635
      SmallChange     =   100
      TabIndex        =   0
      Top             =   0
      Value           =   4635
      Width           =   4395
   End
End
Attribute VB_Name = "FormSP"
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

Private Sub About_Click()
ABMe.Show
End Sub

Private Sub About_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Cl_Click()
On Error GoTo 33
Cdi.ShowColor
FormSP.BackColor = Cdi.Color
Open "settings" For Output As #1
Print #1, Form1.BackColor
Close #1
33 End Sub

Private Sub Cl_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If Ascii = 27 Or KeyAscii = 8 Then Unload Me
End Sub

Private Sub Form_Load()
If App.PrevInstance Then
    Unload Me
End If
If Command$ <> "" Then
    Open Mid(Command$, 2, Len(Command$) - 2) For Input As #1
    On Error GoTo 33
    Input #1, mbcr
    Input #1, mrcr
    Input #1, ClSet(0)
    Input #1, ClSet(1)
    Input #1, ClSet(2)
    Input #1, ClSet(3)
    Input #1, ClSet(4)
    Close #1
    Form1.BackColor = mbcr
    FormSP.BackColor = mrcr
    MsgBox "打开 “" & Mid(Command$, 2, Len(Command$) - 2) & "” 玻璃配置文件成功！", , "玻璃"
End If
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
'''''''''''''''
Dim rtn As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, vbWhite, 150, LWA_ALPHA
HS.Min = 4635
HS.Max = Screen.Width
VS.Min = 3105
VS.Max = Screen.Height
Exit Sub
33 MsgBox "无法打开该文件。", vbCritical, "玻璃"
End
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ReturnVal As Long
X = ReleaseCapture()
ReturnVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
If Button = 1 Then
Cl.Visible = False
HS.Visible = False
VS.Visible = False
Unl.Visible = False
Guide.Visible = False
About.Visible = False
More.Visible = False
ElseIf Button = 2 Then
Cl.Visible = True
HS.Visible = True
VS.Visible = True
Unl.Visible = True
Guide.Visible = True
About.Visible = True
More.Visible = True
End If

End Sub



Private Sub Form_Unload(Cancel As Integer)
If Not mor Then End
End Sub

Private Sub Guide_Click()
SetWindowPos Form1.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
MsgBox "点按鼠标右键显示设置按钮（C为设置玻璃颜色，E为关闭玻璃，A为关于玻璃，G为玻璃导航，M为开启堆叠模式，），点按鼠标左键隐藏设置按钮， 敲击Esc键或退格键关闭玻璃。", , "导航"
SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub

Private Sub Guide_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub HS_Change()
FormSP.Width = HS.Value
End Sub

Private Sub HS_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub More_Click()
mor = True
Mom.Show
Unload FormSP
End Sub

Private Sub unl_Click()
Unload Me
End Sub

Private Sub Unl_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub VS_Change()
FormSP.Height = VS.Value
End Sub

Private Sub VS_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub
