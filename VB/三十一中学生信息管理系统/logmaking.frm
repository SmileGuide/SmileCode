VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "日志编辑器"
   ClientHeight    =   7365
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   13095
   Icon            =   "logmaking.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   13095
   StartUpPosition =   3  '窗口缺省
   Begin ComctlLib.Toolbar tb 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   953
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   8
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "copy"
            Object.ToolTipText     =   "复制"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cut"
            Object.ToolTipText     =   "剪切"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "lock"
            Object.ToolTipText     =   "锁定切换"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "turn"
            Object.ToolTipText     =   "转到（行）"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "open"
            Object.ToolTipText     =   "打开日志文件"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "font"
            Object.ToolTipText     =   "查看字体"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "color"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "save"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComDlg.CommonDialog cbk 
         Left            =   8535
         Top             =   45
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   5460
         TabIndex        =   5
         Top             =   45
         Width           =   2130
         Begin VB.CommandButton tr 
            Height          =   390
            Left            =   765
            Picture         =   "logmaking.frx":7D42
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   30
            Width           =   405
         End
         Begin VB.Image settingS 
            Height          =   390
            Left            =   60
            Picture         =   "logmaking.frx":80CC
            Stretch         =   -1  'True
            Top             =   15
            Width           =   405
         End
      End
      Begin VB.CheckBox night 
         Caption         =   "夜间模式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   4260
         TabIndex        =   3
         Top             =   120
         Width           =   1245
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   9495
         Top             =   45
      End
      Begin MSComDlg.CommonDialog C 
         Left            =   8970
         Top             =   30
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "打开日志文件"
         FileName        =   "未标题"
         Filter          =   "日志文件(*.log)|*.log|所有文件(*.*)|*.*"
         FontName        =   "宋体"
         FontSize        =   12
      End
   End
   Begin ComctlLib.StatusBar st 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   7005
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   635
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   5609
            Text            =   "无文件"
            TextSave        =   "无文件"
            Object.Tag             =   ""
            Object.ToolTipText     =   "文件名"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   5609
            Text            =   "文本未锁定"
            TextSave        =   "文本未锁定"
            Object.Tag             =   ""
            Object.ToolTipText     =   "文本锁定"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   5609
            Text            =   "第1行"
            TextSave        =   "第1行"
            Object.Tag             =   ""
            Object.ToolTipText     =   "光标位置"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   5609
            Text            =   "无文件状态"
            TextSave        =   "无文件状态"
            Object.Tag             =   ""
            Object.ToolTipText     =   "文件状态"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "logmaking.frx":8996
   End
   Begin VB.TextBox Cont 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3315
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   510
      Width           =   6690
   End
   Begin VB.Label bk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "读取文件线程洪荒之力爆发中..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1245
      Left            =   2175
      TabIndex        =   4
      Top             =   4410
      Width           =   9255
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3870
      Top             =   165
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   23
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "logmaking.frx":8CB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "logmaking.frx":932A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "logmaking.frx":9644
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "logmaking.frx":995E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "logmaking.frx":9FD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "logmaking.frx":A2F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "logmaking.frx":A60C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "logmaking.frx":A926
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu wjdk 
      Caption         =   "wjdk"
      Visible         =   0   'False
      Begin VB.Menu filenew 
         Caption         =   "新建"
      End
      Begin VB.Menu fileopen 
         Caption         =   "打开"
      End
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

'''''''''''''
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2
Const LWA_COLORKEY = &H1



Private Sub Cont_Change()
saved = False
End Sub

Private Sub Form_Load()
Load mdd
saved = True
tred = False
Form1.Caption = "日志编辑器 " & Trim$(Str$(App.Major)) & "." & Format$(App.Minor, "##00") & "." & Format$(App.Revision, "0000")
On Error GoTo 55
Open App.Path & "\bkcolor-in-log.ini" For Input As #1
Input #1, bkc
Close #1
On Error GoTo 55
Open App.Path & "\frcolor-in-log.ini" For Input As #2
Input #2, frc
Close #2
Cont.ForeColor = frc
On Error GoTo 55
Open App.Path & "\font-in-log.ini" For Input As #3
On Error GoTo 55
Line Input #3, ft
On Error GoTo 55
Line Input #3, fsb
fs = fsb
Close #3
Cont.Font = ft
Cont.Font.Size = fs
If Command <> "" Then
On Error GoTo 59
    Open Mid(Command, 2, Len(Command) - 1) For Input As #1
    Cont.Visible = False
    st.Panels(4).Text = "正在解析文件，请勿操作..."
        Cont.Text = ""
        Do Until EOF(1)
        nu = nu + 1
        Line Input #1, Con(nu)
        If nu <> 1 Then Cont.Text = Cont.Text & vbCrLf & Con(nu) Else Cont.Text = Cont.Text & Con(nu)
        DoEvents
        Loop
         Close #1
        st.Panels(1).Text = C.FileName
        Cont.Visible = True
        st.Panels(4).Text = "文件解析成功，文件正常"
End If
Exit Sub
55 MsgBox "配置文件未找到或已损坏，请到程序目录内寻找“错误解决程序.exe”运行，点击“解决配置文件丢失错误”按钮来解决错误", vbCritical, "日志编辑器"
Close
End
59 Close
MsgBox "命令行参数非文件名或文件不存在，无法打开。", vbCritical
End Sub

Private Sub Form_Resize()
bk.Top = Cont.Top + 1000
bk.Left = Cont.Left
Cont.Width = Form1.Width - 100
bk.Width = Form1.Width - 100
On Error Resume Next
Cont.Height = Form1.Height - 1310
On Error Resume Next
bk.Height = Form1.Height - 1310
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not saved Then
If MsgBox("是否保存更改？", vbYesNo, "日志编辑器") = vbYes Then
If C.FileName <> "未标题" Then Open C.FileName For Output As #1 Else GoTo 55
Print #1, Cont.Text
Close #1
End If
Unload Me
Exit Sub
55 On Error GoTo 66
C.ShowSave
Open C.FileName For Output As #1
Print #1, Cont.Text
Close #1
End If
66 Close
End Sub

Private Sub night_Click()
If night.Value = 1 Then
    Cont.BackColor = bkc
    If Cont.ForeColor = &H80000008 Then
        Cont.ForeColor = &H80000005
    End If
Else
    Cont.BackColor = vbWhite
    If Cont.ForeColor = &H80000005 Then
        Cont.ForeColor = &H80000008
    End If
End If
End Sub




Private Sub settings_Click()
sets.Show
End Sub

Private Sub tb_ButtonClick(ByVal Button As ComctlLib.Button)
If Button.Key = "copy" And Not Cont.Locked Then
    Clipboard.Clear
    Clipboard.SetText Cont.SelText
ElseIf Button.Key = "cut" And Not Cont.Locked Then
    Clipboard.Clear
    Clipboard.SetText Cont.SelText
    Cont.SelText = ""
ElseIf Button.Key = "lock" Then
    Cont.Locked = Not Cont.Locked
    If Cont.Locked = True Then st.Panels(2) = "文件已锁定，再按锁头解锁" Else st.Panels(2) = "文件未锁定"
ElseIf Button.Key = "open" Then
        C.FileName = App.Path & "\log.log"
        On Error GoTo 34
        C.ShowOpen
        Cont.Visible = False
        Form1.SetFocus
        st.Panels(4).Text = "正在解析文件，请勿操作..."
        Open C.FileName For Input As #1
        Cont.Text = ""
        Do Until EOF(1)
        nu = nu + 1
        Line Input #1, Con(nu)
        If nu <> 1 Then Cont.Text = Cont.Text & vbCrLf & Con(nu) Else Cont.Text = Cont.Text & Con(nu)
        DoEvents
        Loop
         Close #1
        st.Panels(1).Text = C.FileName
        Cont.Visible = True
        st.Panels(4).Text = "文件解析成功，文件正常"
ElseIf Button.Key = "font" Then
   C.ShowFont
   Debug.Print C.FontSize
   With Cont.Font
   .Bold = C.FontBold
   .Italic = C.FontItalic
   .Name = C.FontName
   .Size = C.FontSize
   .Strikethrough = C.FontStrikethru
   .Underline = C.FontUnderline
    End With
ElseIf Button.Key = "turn" Then
    cc = InputBox("请输入要转到的行号", "转到")
    For i = 1 To Len(Cont.Text)
        A = Mid(Cont.Text, i, 1)
        If A = Chr(10) Then lin = lin + 1
        If lin = Val(cc) Then Cont.SelStart = i: Exit Sub
    Next i
ElseIf Button.Key = "color" Then
    C.ShowColor
    Cont.ForeColor = C.Color
Else
    If C.FileName <> "未标题" Then Open C.FileName For Output As #1 Else GoTo 55
        Print #1, Cont.Text
        Close #1
        saved = True
        Exit Sub
55         On Error GoTo 34
            C.ShowSave
            Open C.FileName For Output As #1
            Print #1, Cont.Text
            Close #1
            saved = True
Exit Sub
                On Error GoTo 34
                Cont.Visible = False
             Form1.SetFocus
             st.Panels(4).Text = "正在解析文件，请勿操作..."
              Open C.FileName For Input As #1
                Cont.Text = ""
             Do Until EOF(1)
                nu = nu + 1
                Line Input #1, Con(nu)
             If nu <> 1 Then Cont.Text = Cont.Text & vbCrLf & Con(nu) Else Cont.Text = Cont.Text & Con(nu)
            DoEvents
             Loop
            Close #1
            st.Panels(1).Text = C.FileName
            Cont.Visible = True
            st.Panels(4).Text = "文件解析成功，文件正常"
   End If
34 End Sub




Private Sub tb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu wjdk
End Sub

Private Sub Timer1_Timer()
lie = 0
For i = 1 To Cont.SelStart - 1
If Mid(Cont.Text, i, 1) = Chr(10) Then lie = lie + 1
Next i
If Cont.SelLength = 0 Then st.Panels(3).Text = "第" & lie + 1 & "行 位置:第" & Cont.SelStart + 1 & "字符" Else st.Panels(3).Text = "选中：" & Cont.SelLength & " 位置：" & Cont.SelStart + 1 & "-" & Cont.SelStart + 1 + Cont.SelLength & "字符"
33 End Sub

Private Sub tr_Click()
tred = Not tred
If tred Then
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, &HC0FFFF, 150, LWA_ALPHA
Else
mdd.Text1.Text = Cont.Text
xx = Form1.Left
yy = Form1.Top
w = Form1.Width
h = Form1.Height
saved = True
Unload Form1
Form1.Show
Form1.Left = xx
Form1.Top = yy
Form1.Width = w
Form1.Height = h
Cont.Text = mdd.Text1
If Cont.Text <> "" Then saved = False
End If
End Sub
