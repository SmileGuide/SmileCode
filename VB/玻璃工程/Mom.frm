VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{60CC5D62-2D08-11D0-BDBE-00AA00575603}#1.0#0"; "systray.ocx"
Begin VB.Form Mom 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "玻璃筐"
   ClientHeight    =   2443
   ClientLeft      =   8127
   ClientTop       =   2086
   ClientWidth     =   4683
   Icon            =   "Mom.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Mom.frx":237BE
   ScaleHeight     =   2443
   ScaleWidth      =   4683
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "关于..."
      Height          =   385
      Left            =   3969
      TabIndex        =   12
      Top             =   0
      Width           =   763
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2175
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "调出白板"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton Back 
      BackColor       =   &H00C0FFFF&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "返回"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton Savef 
      BackColor       =   &H00C0FFFF&
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1845
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "存储为配色文件"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton Openf 
      BackColor       =   &H00C0FFFF&
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1530
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "打开配色文件"
      Top             =   0
      Width           =   315
   End
   Begin SysTrayCtl.cSysTray cs 
      Left            =   2220
      Top             =   1110
      _ExtentX        =   966
      _ExtentY        =   966
      InTray          =   0   'False
      TrayIcon        =   "Mom.frx":23B00
      TrayTip         =   "VB 5 - SysTray Control."
   End
   Begin VB.Timer Timerre 
      Interval        =   1
      Left            =   3255
      Top             =   1725
   End
   Begin VB.CommandButton Cl 
      BackColor       =   &H00C0FFFF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   900
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "设置非序列下玻璃颜色"
      Top             =   0
      Width           =   315
   End
   Begin VB.HScrollBar HS 
      Height          =   255
      LargeChange     =   500
      Left            =   255
      Min             =   4635
      SmallChange     =   100
      TabIndex        =   5
      Top             =   375
      Value           =   4635
      Width           =   4395
   End
   Begin VB.VScrollBar VS 
      Height          =   1785
      LargeChange     =   500
      Left            =   30
      Min             =   3105
      SmallChange     =   100
      TabIndex        =   4
      Top             =   630
      Value           =   3105
      Width           =   255
   End
   Begin VB.CommandButton Ode 
      BackColor       =   &H00C0FFFF&
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   615
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "设置序列"
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox slow 
      BackColor       =   &H00C0FFFF&
      Caption         =   "使用颜色序列"
      Height          =   330
      Left            =   2580
      TabIndex        =   2
      Top             =   0
      Width           =   1545
   End
   Begin VB.CommandButton clear 
      BackColor       =   &H00C0FFFF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   15
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "清除所有玻璃"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton hlp 
      BackColor       =   &H00C0FFFF&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   315
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "帮助"
      Top             =   0
      Width           =   315
   End
   Begin MSComDlg.CommonDialog Cdi 
      Left            =   3690
      Top             =   1755
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Color           =   16744576
      DefaultExt      =   "。gss"
      DialogTitle     =   "玻璃配色文件"
      FileName        =   "配色"
      Filter          =   "玻璃配色文件|*.gss"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "预览1:10"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   345
      TabIndex        =   7
      Top             =   690
      Width           =   930
   End
   Begin VB.Shape Shapek 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   660
      Left            =   285
      Top             =   615
      Width           =   975
   End
   Begin VB.Menu a 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu sba 
         Caption         =   "显示玻璃筐"
      End
      Begin VB.Menu sto 
         Caption         =   "结束"
      End
   End
End
Attribute VB_Name = "Mom"
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


Private Sub Back_Click()
rt = True
Unload Mom
FormSP.Show
mor = False
End Sub

Private Sub Cl_Click()
On Error GoTo 33
Cdi.ShowColor
Open "settings" For Output As #1
Print #1, Form1.BackColor
Close #1
If slow.Value <> 1 Then Shapek.BackColor = Cdi.Color: Form1.BackColor = Cdi.Color
slow.Value = 0
33 End Sub

Private Sub clear_Click()
Do
aq = aq + 1
On Error GoTo 33
Unload fm(aq)
Loop
33 End Sub

Private Sub Command1_Click()
WBoard.Show
End Sub

Private Sub Command2_Click()
ABMe.Show
End Sub

Private Sub cs_MouseDblClick(Button As Integer, Id As Long)
cs.InTray = False
Mom.Visible = True
WBoard.Visible = True
down.Visible = True
End Sub



Private Sub cs_MouseUp(Button As Integer, Id As Long)
If Button = 2 Then PopupMenu a
End Sub

Private Sub Form_Initialize()
HS.Min = 4635
HS.Max = Screen.Width
VS.Min = 3105
VS.Max = Screen.Height
Cdi.Color = Form1.BackColor
Shapek.BackColor = Form1.BackColor
down.Show
od = -1
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
nm = nm + 1
od = od + 1
If od > 4 Then od = 0
On Error GoTo 46
fm(nm).Show
fm(nm).Width = HS.Value
fm(nm).Height = VS.Value
If slow.Value = 1 Then fm(nm).BackColor = ClSet(od) Else fm(nm).BackColor = Cdi.Color
fm(nm).Top = Mom.Top
fm(nm).Left = Mom.Width
Dim ReturnVal As Long
X = ReleaseCapture()
ReturnVal = SendMessage(fm(nm).hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
Exit Sub
46 MsgBox "玻璃太多了，计算机会承受不住的。"
End If
End Sub



Private Sub Form_Paint()
mor = True
HS.Min = 4635
HS.Max = Screen.Width
VS.Min = 3105
VS.Max = Screen.Height
Cdi.Color = Form1.BackColor
Shapek.BackColor = Form1.BackColor
down.Show
od = -1
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Unload(Cancel As Integer)
If mor And Not rt Then
mini.Show
Cancel = 1
Mom.Visible = False
WBoard.Visible = False
down.Visible = False
cs.InTray = True
End If
End Sub

Private Sub hlp_Click()
MsgBox "温馨提示：" & vbCrLf & "    -从玻璃筐向外进行拖放操作，即召唤出一张新玻璃。" & vbCrLf & "    -该操作有利于生产多块玻璃(多块玻璃不能随时点按右键召唤设置按钮，需在召唤前设置。在召唤出的玻璃上点按右键，锁定窗体，点按左键解除。" & vbCrLf & "    -可开启颜色序列，设置（O按钮）后，召唤出玻璃的颜色会按序列改变。" & vbCrLf & "    - X为清除所有玻璃，？为帮助，C为设置召唤窗体颜色（非序列下），B为返回单一模式。" & vbCrLf & "    -右键单击召唤出的玻璃锁定玻璃（无法移动），再次按下解锁。", , "帮助"
End Sub



Private Sub Ode_Click()
CLOrder.Show
End Sub


Private Sub Openf_Click()
On Error GoTo 44
Cdi.ShowOpen
Open Cdi.FileName For Input As #1
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
    MsgBox "打开 “" & Cdi.FileName & "” 玻璃配置文件成功！，选中“使用颜色序列”来使用文件中关于颜色序列的内容。", , "玻璃"
Exit Sub
33 MsgBox "类型不符，打开 “" & Cdi.FileName & "” 文件失败，请重新生成配色文件。", vbCritical, "玻璃"
44 End Sub

Private Sub Savef_Click()
On Error GoTo 44
Cdi.ShowSave
Open Cdi.FileName For Output As #1
Print #1, Form1.BackColor
Print #1, FormSP.BackColor
Print #1, ClSet(0)
Print #1, ClSet(1)
Print #1, ClSet(2)
Print #1, ClSet(3)
Print #1, ClSet(4)
Close #1
Close #1
MsgBox "玻璃配置文件已生成！请用本程"
44 End Sub

Private Sub sba_Click()
cs.InTray = False
Mom.Visible = True
down.Visible = True
End Sub

Private Sub slow_Click()
If slow.Value = 1 Then
Shapek.BackStyle = 0
Else
Shapek.BackStyle = 1
End If
End Sub

Private Sub sto_Click()
cs.InTray = False
Do
i = i + 1
On Error GoTo 99
Unload fm(i)
Loop
99 End
End Sub

Private Sub Timerre_Timer()
If down.Left <> Mom.Left - 585 Then down.Left = Mom.Left - 585
If down.Top <> Mom.Top Then down.Top = Mom.Top
If Shapek.Width <> Int(HS.Value / 10) Then Shapek.Width = Int(HS.Value / 10)
If Shapek.Height <> Int(VS.Value / 10) Then Shapek.Height = Int(VS.Value / 10)
End Sub

