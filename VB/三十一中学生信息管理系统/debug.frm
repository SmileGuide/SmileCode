VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "控件错误解决程序"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   Icon            =   "debug.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "解决配置文件丢失错误"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   105
      TabIndex        =   4
      Top             =   1545
      Width           =   4515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "复制Active X控件解决口令"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   90
      TabIndex        =   2
      Top             =   960
      Width           =   4515
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "稍后请再次运行发生错误的产品，若错误还未解决，我们深表歉意，请将程序文件夹内的log.log日志文件上传给我们，我们将分析并解决问题。"
      Height          =   1485
      Left            =   45
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   4530
   End
   Begin VB.Label Label2 
      Caption         =   "    如果错误关于Active X控件，请点击下方复制口令按钮来复制注册控件cmd口令，在弹出的cmd窗口中按Ctrl+V来粘贴口令，点按回车。"
      Height          =   690
      Left            =   90
      TabIndex        =   1
      Top             =   375
      Width           =   4170
   End
   Begin VB.Label Label1 
      Caption         =   "您在使用本产品时，可能遇到运行错误。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   0
      Top             =   60
      Width           =   4560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
Private Const SWP_NOSIZE& = &H1
Private Const SWP_NOMOVE& = &H2

Private Sub Command1_Click()
Clipboard.Clear
Clipboard.SetText "regsvr32 /s systray.ocx"
a = Shell("cmd", vbNormalFocus)
Label1.Visible = False
Label2.Visible = False
Command1.Visible = False
Label3.Visible = True
Label1.Caption = "粘贴口令到命令行后请再次运行发生错误的产品，若错误还未解决，我们深表歉意，请联系我们，我们将分析并解决问题。"
End Sub


Private Sub Command2_Click()
If MsgBox("此操作会导致程序恢复默认设置，是否继续进行？", vbExclamation + vbYesNo, "错误解决向导") = vbYes Then
Open App.Path & "\bkcolor-in-log.ini" For Output As #1
Print #1, vbBlack
Close #1
Open App.Path & "\frcolor-in-log.ini" For Output As #2
Print #2, vbBlack
Close #2
ft = "宋体"
fs = 12
Open App.Path & "\font-in-log.ini" For Output As #3
Print #3, ft
Print #3, fs
Close #3
Label1.Caption = "已解决！"
End If
End Sub

Private Sub Form_Load()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub
