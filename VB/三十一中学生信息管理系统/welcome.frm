VERSION 5.00
Object = "{60CC5D62-2D08-11D0-BDBE-00AA00575603}#1.0#0"; "systray.ocx"
Begin VB.Form welcome 
   BackColor       =   &H00C0FFFF&
   Caption         =   "欢迎！"
   ClientHeight    =   3960
   ClientLeft      =   9765
   ClientTop       =   5040
   ClientWidth     =   7035
   Icon            =   "welcome.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3960
   ScaleWidth      =   7035
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5895
      Top             =   3345
   End
   Begin SysTrayCtl.cSysTray Ic 
      Left            =   6345
      Top             =   3315
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   -1  'True
      TrayIcon        =   "welcome.frx":1B692
      TrayTip         =   "三十一中学生信息管理系统"
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   15
      ScaleHeight     =   3615
      ScaleWidth      =   6735
      TabIndex        =   0
      Top             =   15
      Width           =   6735
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "版权所有，违者必究！"
         Height          =   255
         Left            =   2400
         TabIndex        =   22
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   3495
         Left            =   1440
         Picture         =   "welcome.frx":36D34
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "查1.3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "查1.4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   0
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "31中学生信息管理系统"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   1635
         TabIndex        =   19
         Top             =   165
         Width           =   3975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "查1.5"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   0
         TabIndex        =   18
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "查1.7"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   0
         TabIndex        =   17
         Top             =   1800
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "查1.8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   0
         TabIndex        =   16
         Top             =   2400
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "查1.9"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   0
         TabIndex        =   15
         Top             =   3000
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "查1.10"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   1680
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "查1.11"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   1680
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "查2.1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   1680
         TabIndex        =   12
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "查2.2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   1680
         TabIndex        =   11
         Top             =   1800
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "查2.3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   1680
         TabIndex        =   10
         Top             =   2400
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "查2.4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   1680
         TabIndex        =   9
         Top             =   3000
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "查2.5"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   3360
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "查2.6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   3360
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "查3.3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   3360
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "查3.4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   3360
         TabIndex        =   5
         Top             =   1800
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "查3.8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   3360
         TabIndex        =   4
         Top             =   2400
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "查3.9"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   3360
         TabIndex        =   3
         Top             =   3000
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "查3.10"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   5040
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "搜索全校"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2175
         Left            =   5595
         TabIndex        =   1
         Top             =   1080
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Menu tray 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu we 
         Caption         =   "显示欢迎界面"
      End
      Begin VB.Menu log 
         Caption         =   "日志"
      End
      Begin VB.Menu chusername 
         Caption         =   "更改用户信息"
      End
      Begin VB.Menu cc 
         Caption         =   "收起菜单"
      End
      Begin VB.Menu exi 
         Caption         =   "退出程序"
      End
   End
End
Attribute VB_Name = "welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
Private Const SWP_NOSIZE& = &H1
Private Const SWP_NOMOVE& = &H2

Private Sub cc_Click()
tray.Visible = False
End Sub

Private Sub exi_Click()
Ic.InTray = False
c14.Visible = False
welcome.Visible = False
Open "log.log" For Append As #1
Print #1, Now & vbCrLf & "End"
Close #1
End
End Sub





Private Sub Form_Load()
SetWindowPos Me.hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
welcome.Caption = "欢迎来到三十一中学生信息管理系统 " & Trim$(Str$(App.Major)) & "." & Format$(App.Minor, "##00") & "." & Format$(App.Revision, "0000") & " ！"
End Sub

Private Sub Form_Resize()
On Error GoTo 33
Picture1.Top = welcome.Height / 2 - Picture1.Height / 2 - 300
Picture1.Left = welcome.Width / 2 - Picture1.Width / 2
If welcome.Width < Picture1.Width Then welcome.Width = Picture1.Width
If welcome.Height < Picture1.Height Then welcome.Height = Picture1.Height
33 End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 2
welcome.Visible = False
mini.Show
Open "log.log" For Append As #1
Print #1, Now, vbCrLf & "Minimized"
Close #1
End Sub

Private Sub Ic_MouseUp(Button As Integer, Id As Long)
If Button = 2 Then PopupMenu tray
If Button = 1 Then c14.Hide: welcome.Show
End Sub

Private Sub Image1_Click()
Label3.Visible = True
Label2.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True
Label12.Visible = True
Label13.Visible = True
Label14.Visible = True
Label15.Visible = True
Label16.Visible = True
Label17.Visible = True
Label18.Visible = True
Label19.Visible = True
Label20.Visible = True
Label21.Visible = True
Label22.Visible = True
Image1.Visible = False
Label4.Visible = False
Label1.Visible = False
End Sub

Private Sub Label10_Click()
se = "reco111"
c14.Show
welcome.Hide: tipstart = True
End Sub

Private Sub Label11_Click()
se = "reco201"
c14.Show
welcome.Hide: tipstart = True
End Sub

Private Sub Label12_Click()
se = "reco202"
c14.Show
welcome.Hide: tipstart = True
End Sub

Private Sub Label13_Click()
se = "reco203"
c14.Show
welcome.Hide: tipstart = True
End Sub

Private Sub Label14_Click()
se = "reco204"
c14.Show
welcome.Hide: tipstart = True
End Sub

Private Sub Label15_Click()
se = "reco205"
c14.Show
welcome.Hide: tipstart = True
End Sub

Private Sub Label16_Click()
se = "reco206"
c14.Show
welcome.Hide: tipstart = True
End Sub

Private Sub Label17_Click()
se = "reco303"
c14.Show
welcome.Hide: tipstart = True
End Sub

Private Sub Label18_Click()
se = "reco304"
c14.Show
welcome.Hide: tipstart = True
End Sub

Private Sub Label19_Click()
se = "reco308"
c14.Show
welcome.Hide: tipstart = True
End Sub

Private Sub Label2_Click()
se = "reco103"
c14.Show
welcome.Hide: tipstart = True
End Sub

Private Sub Label20_Click()
se = "reco309"
c14.Show
welcome.Hide: tipstart = True
End Sub

Private Sub Label21_Click()
se = "reco310"
c14.Show
welcome.Hide: tipstart = True
End Sub



Private Sub Label22_Click()
se = "recoall"
c14.Show
welcome.Hide: tipstart = True
End Sub

Private Sub Label3_Click()
se = "reco104"
c14.Show
welcome.Hide: tipstart = True
End Sub

Private Sub Label5_Click()
se = "reco105"
c14.Show
welcome.Hide: tipstart = True
End Sub

Private Sub Label6_Click()
se = "reco109"
c14.Show
welcome.Hide: tipstart = True
End Sub

Private Sub Label7_Click()
se = "reco107"
c14.Show
welcome.Hide: tipstart = True
End Sub

Private Sub Label8_Click()
se = "reco108"
c14.Show
welcome.Hide: tipstart = True
End Sub

Private Sub Label9_Click()
se = "reco110"
c14.Show
welcome.Hide: tipstart = True
End Sub




Private Sub log_Click()
On Error GoTo 45
Open App.Path & "\log.log" For Input As #1
Close #1
Shell (App.Path & "日志编辑器.exe " & App.Path & "\log.log")
GoTo 88
45 MsgBox "日志不存在，未记录或名称不正确。请确保日志位于程序文件夹，已记录内容且名称为“log.log”"
88 End Sub

Private Sub we_Click()
welcome.Visible = True
c14.Hide
End Sub
