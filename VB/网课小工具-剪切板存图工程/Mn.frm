VERSION 5.00
Begin VB.Form Mn 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3880
   ClientLeft      =   5
   ClientTop       =   5
   ClientWidth     =   2515
   ControlBox      =   0   'False
   Icon            =   "Mn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3880
   ScaleWidth      =   2515
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdNew 
      BackColor       =   &H00C0E0FF&
      Caption         =   "新建文件夹"
      Height          =   245
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "最小化"
      Top             =   1740
      Width           =   2525
   End
   Begin VB.PictureBox PicEx 
      Height          =   125
      Left            =   1740
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   125
   End
   Begin VB.Timer TMR 
      Interval        =   1
      Left            =   1200
      Top             =   240
   End
   Begin VB.CommandButton CmdMini 
      BackColor       =   &H00C0E0FF&
      Caption         =   "最小化"
      Height          =   245
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "最小化"
      Top             =   3660
      Width           =   2525
   End
   Begin VB.ComboBox CboFormat 
      Height          =   220
      ItemData        =   "Mn.frx":424A
      Left            =   0
      List            =   "Mn.frx":425A
      TabIndex        =   8
      Text            =   "JPG"
      Top             =   3240
      Width           =   2525
   End
   Begin VB.FileListBox FleFL 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "目录内文件"
      Top             =   1980
      Width           =   2525
   End
   Begin VB.CommandButton CmdAbout 
      BackColor       =   &H00FFFFC0&
      Caption         =   "…"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   245
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "关于"
      Top             =   0
      Width           =   245
   End
   Begin VB.DirListBox DirFL 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1150
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "目录"
      Top             =   600
      Width           =   2525
   End
   Begin VB.DriveListBox DrvFL 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   190
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "驱动器"
      Top             =   420
      Width           =   2525
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00FFFFC0&
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   245
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "退出"
      Top             =   0
      Width           =   245
   End
   Begin VB.CommandButton CmdMove 
      BackColor       =   &H00C0E0FF&
      Caption         =   "剪贴板存图助手"
      Height          =   245
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "移动"
      Top             =   0
      Width           =   2045
   End
   Begin VB.Line LinUD 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      DrawMode        =   14  'Copy Pen
      X1              =   0
      X2              =   2520
      Y1              =   3540
      Y2              =   3540
   End
   Begin VB.Label LblFormat 
      BackStyle       =   0  'Transparent
      Caption         =   "保存格式"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   185
      Left            =   0
      TabIndex        =   7
      Top             =   3060
      Width           =   2525
   End
   Begin VB.Label LblFL 
      BackStyle       =   0  'Transparent
      Caption         =   "保存目录"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   185
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   2525
   End
End
Attribute VB_Name = "Mn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ExpFL As String
Dim Nw As Boolean
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2



Private Sub CmdAbout_Click()
Abme.Show
End Sub



Private Sub CmdAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = 1
End Sub

Private Sub CmdExit_Click()
End
End Sub

Private Sub CmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = 1
End Sub

Private Sub CmdMini_Click()
Mn.WindowState = 1
End Sub

Private Sub CmdMove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ReturnVal As Long
X = ReleaseCapture()
ReturnVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub



Private Sub CmdMove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = 15
End Sub

Private Sub CmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = 1
End Sub



Private Sub CmdNew_Click()
ExpFL = DirFL.Path _
& "\" _
& Format(Date, "yyyymmdd") _
& "-" _
& Format(Now, "hhmmss")
On Error GoTo Frequen
MkDir ExpFL
On Error GoTo Frequen
DirFL.Path = ExpFL
Exit Sub
Frequen: MsgBox "请勿操作过于频繁，间隔小于1s。"
End Sub

Private Sub DirFL_Change()
FleFL.Path = DirFL.Path
SaveSetting "PctFrmClp", "Setting", "Dir", DirFL.Path
End Sub

Private Sub DrvFL_Change()
If Not Nw Then
DirFL.Path = DrvFL.Drive
SaveSetting "PctFrmClp", "Setting", "Drive", DrvFL.Drive
End If
End Sub

Private Sub Form_Load()
Nw = True
DrvFL.Drive = _
GetSetting("PctFrmClp", "Setting", "Drive", "C:\")
DirFL.Path = _
GetSetting("PctFrmClp", "Setting", "Dir", DrvFL.Drive)
Nw = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = 1
End Sub

Private Sub Label1_Click()

End Sub

Private Sub LblFL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = 1
End Sub

Private Sub TMR_Timer()
On Error GoTo cannot
PicEx.Picture = Clipboard.GetData
If PicEx.Picture <> 0 Then
    On Error GoTo no
    SavePicture _
    PicEx.Picture, _
    DirFL.Path _
    & "\" _
    & Format(Date, "yyyymmdd") _
    & "-" _
    & Format(Now, "hhmmss") _
    & "." & Format(CboFormat.Text, "<")
    Clipboard.Clear
    DirFL.Refresh
    FleFL.Refresh
End If
Exit Sub
no: If Err.Number = 53 Then MsgBox "未找到目录，请重新选择"
cannot: End Sub
