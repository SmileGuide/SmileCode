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
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.CommandButton ttLog 
         BackColor       =   &H00C0FFFF&
         Caption         =   "日志编辑器"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1845
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2175
         Width           =   1455
      End
      Begin VB.CommandButton ttMain 
         BackColor       =   &H00C0FFFF&
         Caption         =   "主程序"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2190
         Width           =   1455
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "转到："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   330
         Left            =   180
         TabIndex        =   4
         Top             =   1785
         Width           =   930
      End
      Begin VB.Image unl 
         Height          =   480
         Left            =   6480
         Picture         =   "beg-frmSplash - .frx":0000
         ToolTipText     =   "关闭"
         Top             =   195
         Width           =   480
      End
      Begin VB.Image imgLogo 
         Height          =   990
         Left            =   750
         Picture         =   "beg-frmSplash - .frx":26EBA
         Stretch         =   -1  'True
         Top             =   735
         Width           =   915
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
         Caption         =   "使用向导"
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
         Top             =   1065
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



Private Sub Label1_Click()

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


Private Sub Form_Load()
lblProductName.Caption = "使用向导 " & App.Major & "." & App.Minor & "." & App.Revision
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
