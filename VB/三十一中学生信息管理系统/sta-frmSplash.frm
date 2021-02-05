VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   4050
      Left            =   135
      TabIndex        =   0
      Top             =   45
      Width           =   7080
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   330
         Left            =   5595
         TabIndex        =   5
         Top             =   1770
         Width           =   255
      End
      Begin VB.Timer Timer2 
         Left            =   705
         Top             =   75
      End
      Begin VB.Timer Timer1 
         Interval        =   800
         Left            =   285
         Top             =   75
      End
      Begin VB.Image unl 
         Height          =   480
         Left            =   6480
         Picture         =   "sta-frmSplash.frx":0000
         ToolTipText     =   "关闭"
         Top             =   195
         Width           =   480
      End
      Begin VB.Shape Shst 
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000C0&
         Height          =   345
         Left            =   285
         Shape           =   4  'Rounded Rectangle
         Top             =   3180
         Visible         =   0   'False
         Width           =   6630
      End
      Begin VB.Line staterr 
         BorderColor     =   &H0000C000&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         X1              =   300
         X2              =   6870
         Y1              =   3555
         Y2              =   3555
      End
      Begin VB.Label state 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "请将程序文件夹内的日志编辑器打开来为本产品提供数据源"
         BeginProperty Font 
            Name            =   "华文中宋"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E335DF&
         Height          =   375
         Left            =   315
         TabIndex        =   4
         Top             =   3180
         Width           =   6510
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   90
         Picture         =   "sta-frmSplash.frx":26EBA
         Stretch         =   -1  'True
         Top             =   795
         Width           =   2085
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H00C0FFFF&
         Caption         =   "S.G.G.公司"
         BeginProperty Font 
            Name            =   "华文中宋"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   5940
         TabIndex        =   1
         Top             =   3690
         Width           =   1065
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "使用统计"
         BeginProperty Font 
            Name            =   "华文中宋"
            Size            =   32.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   2355
         TabIndex        =   3
         Top             =   1185
         Width           =   2580
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "三十一中学生信息管理系统"
         BeginProperty Font 
            Name            =   "华文中宋"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   2220
         TabIndex        =   2
         Top             =   705
         Width           =   4320
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2


Private Sub Command1_Click()
Open "Log-S.ini" For Output As #1
Print #1, 2
'程序状态：
'0 关闭
'1 统计端已打开，正在监听
'2 编辑器已打开，正在监听
'3 编辑器端已接受请求，统计端可开始运行
'4 编辑器发送数据
'5 统计端发送数据
Close #1
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ReturnVal As Long
X = ReleaseCapture()
ReturnVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub



Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ReturnVal As Long
X = ReleaseCapture()
ReturnVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub





Private Sub imgLogo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ReturnVal As Long
X = ReleaseCapture()
ReturnVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub



Private Sub lblCompany_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ReturnVal As Long
X = ReleaseCapture()
ReturnVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub


Private Sub lblCompanyProduct_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ReturnVal As Long
X = ReleaseCapture()
ReturnVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub



Private Sub lblProductName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ReturnVal As Long
X = ReleaseCapture()
ReturnVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub state_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ReturnVal As Long
X = ReleaseCapture()
ReturnVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub Timer1_Timer()
staterr.Visible = Not staterr.Visible
Shst.Visible = Not Shst.Visible
End Sub

Private Sub Form_Load()
lblProductName.Caption = "使用统计 " & App.Major & "." & App.Minor & "." & App.Revision
Open "Log-S.ini" For Input As #1
Input #1, exstate
'程序状态：
'0 关闭
'1 统计端已打开，正在监听
'2 编辑器已打开，正在监听
'3 编辑器端已接受请求，统计端可开始运行
'4 编辑器发送数据
'5 统计端发送数据
Close #1
 If exstate <> 2 Then
exstate = 1
Open "Log-S.ini" For Output As #2
Print #2, exstate
Close #2
Else
exstate = 3
Open "Log-S.ini" For Output As #3
Print #3, exstate
Close #3
Unload frmSplash
Form1.Show
End If
End Sub

Private Sub unl_Click()
Unload Me
End Sub

Private Sub unl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
unl.BorderStyle = 1
End Sub

Private Sub unl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
unl.BorderStyle = 0
End Sub
