VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form sets 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                              设置 - 夜间模式背景颜色  默认文字颜色 默认字体"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10545
   Icon            =   "settings.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton newini 
      BackColor       =   &H00FFC0FF&
      Caption         =   "恢复默认"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   8295
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   195
      Width           =   2010
   End
   Begin MSComDlg.CommonDialog cbk 
      Left            =   6825
      Top             =   1500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "设置 - 夜间模式背景颜色  默认文字颜色 默认字体"
      FontName        =   "宋体"
   End
   Begin VB.CommandButton fos 
      BackColor       =   &H00C0FFC0&
      Caption         =   "默认字体设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   5295
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   195
      Width           =   2655
   End
   Begin VB.CommandButton woc 
      BackColor       =   &H00C0FFC0&
      Caption         =   "默认文字颜色设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   2730
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   180
      Width           =   2580
   End
   Begin VB.CommandButton bkco 
      BackColor       =   &H00C0FFC0&
      Caption         =   "夜间模式背景颜色设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   180
      Width           =   2580
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   8130
      X2              =   8130
      Y1              =   345
      Y2              =   885
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "1234567890+-*/[]{}|@#￥$%……*（）~"
      Height          =   1920
      Index           =   2
      Left            =   435
      TabIndex        =   2
      Top             =   4230
      Width           =   9840
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "ABCDEFGHIJKLMNOPQRESUVWXYZ,.;:'""?!\"
      Height          =   1260
      Index           =   1
      Left            =   435
      TabIndex        =   1
      Top             =   2880
      Width           =   9990
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "感谢您使用日志编辑器，。；;‘“？！、——【】·"
      Height          =   1305
      Index           =   0
      Left            =   390
      TabIndex        =   0
      Top             =   1485
      Width           =   9885
   End
   Begin VB.Shape modb 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   5055
      Left            =   30
      Shape           =   4  'Rounded Rectangle
      Top             =   1290
      Width           =   10470
   End
   Begin VB.Label Labelu 
      BackStyle       =   0  'Transparent
      Caption         =   "综合预览"
      Height          =   195
      Left            =   810
      TabIndex        =   6
      Top             =   1080
      Width           =   1230
   End
End
Attribute VB_Name = "sets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
Private Const SWP_NOSIZE& = &H1
Private Const SWP_NOMOVE& = &H2

Private Sub bkco_Click()
On Error GoTo Ex
cbk.ShowColor
modb.BackColor = cbk.Color
bkc = cbk.Color
Open App.Path & "\bkcolor-in-log.ini" For Output As #1
Print #1, bkc
Close #1
If Form1.night.Value = 1 Then Form1.Cont.BackColor = bkc
Ex: End Sub

Private Sub Form_Load()
SetWindowPos sets.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub fos_Click()
ftbag = Label(0).Font
fsbag = Label(0).Font.Size
ft = InputBox("请输入要改变的字体，字体不合法视为默认宋体。", "日志编辑器")
On Error GoTo 999
fs = InputBox("请输入要改变的字号。", "日志编辑器")
For i = 0 To 2
On Error GoTo er
Label(i).Font = ft
On Error GoTo er
Label(i).Font.Size = fs
Next i
Open App.Path & "\Font-in-log.ini" For Output As #1
Print #1, ft
Print #1, fs
Close #1
Exit Sub
er: For j = 0 To 2
Label(j).Font = ft
Label(j).Font.Size = fs
Next j
Open App.Path & "\Font-in-log.ini" For Output As #1
Print #1, ftbag
Print #1, fsbag
Close #1
999 End Sub

Private Sub newini_Click()
If MsgBox("是否将所有配置恢复默认？", vbQuestion + vbYesNo, "日志编辑器") = vbYes Then
    bkb = vbWhite
    Open App.Path & "\BkColor-in-log.ini" For Output As #1
        Print #1, bkb
    Close #1
    If Form1.night.Value = 1 Then Form1.Cont.BackColor = bkb
    frb = vbBlack
    Open App.Path & "\FrColor-in-log.ini" For Output As #2
        Print #2, frb
    Close #2
    ft = "宋体"
    fs = 12
    Open App.Path & "\Font-in-log.ini" For Output As #3
        Print #3, ft
        Print #3, fs
    Close #3
    MsgBox "初始化成功!", , "日志编辑器"
End If
End Sub

Private Sub woc_Click()
On Error GoTo Ex
cbk.ShowColor
frc = cbk.Color
For i = 0 To 2
Label(i).ForeColor = frc
Next i
Open App.Path & "\frcolor-in-log.ini" For Output As #1
Print #1, frc
Close #1
Ex: End Sub
