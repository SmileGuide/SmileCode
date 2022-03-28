VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmCED 
   BackColor       =   &H00000000&
   Caption         =   "高级编辑"
   ClientHeight    =   6732
   ClientLeft      =   24
   ClientTop       =   360
   ClientWidth     =   11160
   Icon            =   "FrmCED.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6732
   ScaleWidth      =   11160
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CmdC 
      Left            =   2160
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Color           =   16777215
   End
   Begin RichTextLib.RichTextBox TxtCvr 
      Height          =   1326
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1086
      _ExtentX        =   1916
      _ExtentY        =   2339
      _Version        =   393217
      TextRTF         =   $"FrmCED.frx":424A
   End
   Begin RichTextLib.RichTextBox TxtCode 
      Height          =   6726
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11166
      _ExtentX        =   19696
      _ExtentY        =   11864
      _Version        =   393217
      BackColor       =   789516
      ScrollBars      =   3
      TextRTF         =   $"FrmCED.frx":42E7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmCED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LastPsn
Dim CRLFED
Dim CauText As String

Private Sub Command1_Click()
ReflashCmdFormat
End Sub

Private Sub Form_Load()
CodeThemeRead
On Error GoTo 99
Me.BackColor = SknColor
TxtCode.BackColor = SknColor
Exit Sub

99 CodeThemeReset
End Sub

Private Sub Form_Paint()
TxtCode.Width = Me.Width
TxtCode.Height = Me.Height
End Sub
Private Sub TxtCode_Change()
Saved = False
If UnDis Then UnDis = False: Exit Sub
If Covered Then
    Select Case CvrText
    Dim lstWord
    Case "clear"
    lstWord = Right(TxtCode.Text, 3)
        If lstWord = "y" & vbCrLf Then
            TxtCode.SelStart = 0
            TxtCode.SelLength = Len(TxtCode.Text)
            TxtCode.SelText = ""
        Else
            TxtCode.Text = TxtCvr.Text
        End If
    Case "quit"
        lstWord = Right(TxtCode.Text, 3)
        If lstWord = "y" & vbCrLf Then
            Saved = True
            Unload Me
        Else
            TxtCode.Text = TxtCvr.Text
        End If
    End Select
Covered = False
End If
On Error Resume Next
Dim p
p = TxtCode.SelStart
TxtCode.SelStart = LastPsn
TxtCode.SelLength = Len(TxtCode.Text) - 1
LastPsn = TxtCode.SelStart + TxtCode.SelLength
TxtCode.SelColor = TxtColor
TxtCode.SelLength = 0
TxtCode.SelStart = p

On Error GoTo 77

If Mid(TxtCode.Text, TxtCode.SelStart, 2) = Chr(10) Or Mid(TxtCode.Text, TxtCode.SelStart, 1) = Chr(10) Then CRLFED = True: GoTo 45
CRLFED = False
33 TxtCode.SelStart = TxtCode.SelStart - 1
TxtCode.SelLength = 1
TxtCode.SelColor = TxtColor



If Right(LCase(TxtCode.Text), 5) = "clear" Then
TxtCode.SelStart = TxtCode.SelStart - 4
TxtCode.SelLength = 5
TxtCode.SelColor = CmdColor
TxtCode.SelStart = TxtCode.SelStart + 4

ElseIf Right(LCase(TxtCode.Text), 4) = "quit" Then
TxtCode.SelStart = TxtCode.SelStart - 3
TxtCode.SelLength = 4
TxtCode.SelColor = CmdColor
TxtCode.SelStart = TxtCode.SelStart + 3

ElseIf Right(LCase(TxtCode.Text), 4) = "help" Then
TxtCode.SelStart = TxtCode.SelStart - 3
TxtCode.SelLength = 4
TxtCode.SelColor = CmdColor
TxtCode.SelStart = TxtCode.SelStart + 3


ElseIf Right(LCase(TxtCode.Text), 6) = "format" Then
TxtCode.SelStart = TxtCode.SelStart - 5
TxtCode.SelLength = 6
TxtCode.SelColor = CmdColor
TxtCode.SelStart = TxtCode.SelStart + 5

ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart, 2) = "**" Then
        Dim pp
        pp = TxtCode.SelStart
        FrmCED.TxtCode.SelStart = 0
        FrmCED.TxtCode.SelLength = Len(FrmCED.TxtCode.Text)
        FrmCED.TxtCode.SelColor = TxtColor
        
        Dim i
        For i = 0 To Len(FrmCED.TxtCode.Text) - 1
        FrmCED.TxtCode.SelStart = i
        FrmCED.TxtCode.SelLength = 1
        
        Select Case FrmCED.TxtCode.SelText
        Case Chr(34), Chr(44), Chr(58)
        FrmCED.TxtCode.SelColor = SpcColor
        
        Case Chr(46), Chr(59), Chr(123), Chr(125)
        FrmCED.TxtCode.SelColor = CmdColor
        Case 0 To 9
        FrmCED.TxtCode.SelColor = NumColor
        End Select

        FrmCED.TxtCode.SelLength = 2
        If FrmCED.TxtCode.SelText = "**" Then FrmCED.TxtCode.SelText = ""

        FrmCED.TxtCode.SelLength = 3
        Select Case FrmCED.TxtCode.SelText
        Case "day"
        FrmCED.TxtCode.SelColor = CmdColor
        End Select
        
        FrmCED.TxtCode.SelLength = 5
        Select Case FrmCED.TxtCode.SelText
        Case "clear"
        FrmCED.TxtCode.SelColor = CmdColor
        End Select
        
        
        Next
        FrmCED.TxtCode.SelLength = 0
        TxtCode.SelStart = pp
    
77 ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart + 1, 1) = "*" Then
TxtCode.SelLength = 1
TxtCode.SelColor = CmdColor
TxtCode.SelStart = TxtCode.SelStart + 1

ElseIf Right(LCase(TxtCode.Text), 9) = "skin-dark" Then
TxtCode.SelStart = TxtCode.SelStart - 8
TxtCode.SelLength = 9
TxtCode.SelColor = CmdColor
TxtCode.SelStart = TxtCode.SelStart + 8

ElseIf Right(LCase(TxtCode.Text), 11) = "skin-bright" Then
TxtCode.SelStart = TxtCode.SelStart - 10
TxtCode.SelLength = 11
TxtCode.SelColor = CmdColor
TxtCode.SelStart = TxtCode.SelStart + 10

ElseIf Right(LCase(TxtCode.Text), 11) = "skin-custom" Then
TxtCode.SelStart = TxtCode.SelStart - 10
TxtCode.SelLength = 11
TxtCode.SelColor = CmdColor
TxtCode.SelStart = TxtCode.SelStart + 10

End If

'贴尾归位
TxtCode.SelStart = TxtCode.SelStart + 1
99 TxtCode.SelLength = 0
If Not Covered And CRLFED Then ReflashCmdFormat
Exit Sub





45
If Mid(TxtCode.Text, TxtCode.SelStart - 1, 2) = Chr(46) & vbCrLf Then
        ''''''''''''保存
    ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 5, 6) = "quit" & vbCrLf Then
        CvrText = "quit"
        UnDis = False
        TxtCvr.Text = TxtCode.Text
        TxtCvr.Text = Left(TxtCvr.Text, Len(TxtCvr.Text) - 6)

        UnDis = True
        
        CauText = "警告：你还没有保存课程表。" & Chr(10) & "确定要执行此命令，请输入y后回车；撤回此命令，请连续按下回车。>>"
        TxtCode.Text = TxtCode.Text & vbCrLf & CauText
        ReflashCmdFormat
        TxtCode.SelStart = Len(TxtCode.Text) - Len(CauText)
        TxtCode.SelLength = 14
        TxtCode.SelColor = vbRed
        TxtCode.SelLength = 0
        TxtCode.SelStart = Len(TxtCode.Text) - 34
        TxtCode.SelLength = 34
        TxtCode.SelColor = TxtColor
        TxtCode.SelStart = Len(TxtCode.Text)
        TxtCode.SelLength = 0
        Covered = True
        
    ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 6, 7) = "clear" & vbCrLf Then
        CvrText = "clear"
        UnDis = False
        TxtCvr.Text = TxtCode.Text
        TxtCvr.Text = Left(TxtCvr.Text, Len(TxtCvr.Text) - 7)

        UnDis = True
        
        CauText = "警告：此操作是不可逆的。" & Chr(10) & "确定要执行此命令，请输入y后回车；撤回此命令，请连续按下回车。>>"
        TxtCode.Text = TxtCode.Text & vbCrLf & CauText
        ReflashCmdFormat
        TxtCode.SelStart = Len(TxtCode.Text) - Len(CauText)
        TxtCode.SelLength = 13
        TxtCode.SelColor = vbRed
        TxtCode.SelLength = 0
        TxtCode.SelStart = Len(TxtCode.Text) - 34
        TxtCode.SelLength = 34
        TxtCode.SelColor = TxtColor
        TxtCode.SelStart = Len(TxtCode.Text)
        TxtCode.SelLength = 0
        Covered = True

        ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 7, 8) = "format" & vbCrLf Then
        






       
    ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 10, 11) = "skin-dark" & vbCrLf Then
        TxtCode.SelStart = TxtCode.SelStart - 11
        TxtCode.SelLength = 11
        TxtCode.SelText = ""
        TxtCode.SelStart = TxtCode.SelStart + 11
        SaveSetting "SmileTimetable", "Code", "BgColor", &HC0C0C
        SaveSetting "SmileTimetable", "Code", "SpecialColor", &HFF00FF
    SaveSetting "SmileTimetable", "Code", "CommandColor", &HFF00&
        SaveSetting "SmileTimetable", "Code", "NumColor", &HFFFF&
        SknColor = GetSetting("SmileTimetable", "Code", "BgColor")
        NumColor = GetSetting("SmileTimetable", "Code", "NumColor")
        SaveSetting "SmileTimetable", "Code", "TxtColor", vbWhite
        Me.BackColor = SknColor
        TxtCode.BackColor = SknColor
        TxtColor = vbWhite
        SpcColor = &HFF00FF
        CmdColor = &HFF00&
        ReflashCmdFormat
ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 12, 13) = "skin-bright" & vbCrLf Then
        TxtCode.SelStart = TxtCode.SelStart - 13
        TxtCode.SelLength = 13
        TxtCode.SelText = ""
        TxtCode.SelStart = TxtCode.SelStart + 13
        SaveSetting "SmileTimetable", "Code", "BgColor", &HFFFFFF
        SaveSetting "SmileTimetable", "Code", "NumColor", &H2469F6
        SaveSetting "SmileTimetable", "Code", "TxtColor", vbBlack
        SaveSetting "SmileTimetable", "Code", "SpecialColor", &HFF00FF
        SaveSetting "SmileTimetable", "Code", "CommandColor", &HFF00&
        NumColor = GetSetting("SmileTimetable", "Code", "NumColor")
        SknColor = GetSetting("SmileTimetable", "Code", "BgColor")
        Me.BackColor = SknColor
        TxtCode.BackColor = SknColor
        TxtColor = vbBlack
                SpcColor = &HFF00FF
        CmdColor = &HFF00&
        ReflashCmdFormat
ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 12, 13) = "skin-custom" & vbCrLf Then
        TxtCode.SelStart = TxtCode.SelStart - 13
        TxtCode.SelLength = 13
        TxtCode.SelText = ""
        TxtCode.SelStart = TxtCode.SelStart + 13

        SknColor = GetSetting("SmileTimetable", "Code", "BgColorCustom")
        SaveSetting "SmileTimetable", "Code", "BgColor", SknColor
        NumColor = GetSetting("SmileTimetable", "Code", "NumColorCustom")
        SaveSetting "SmileTimetable", "Code", "NumColor", NumColor
        TxtColor = GetSetting("SmileTimetable", "Code", "TxtColorCustom")
        SaveSetting "SmileTimetable", "Code", "TxtColor", TxtColor
        Me.BackColor = SknColor
        TxtCode.BackColor = SknColor
        
        
ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 20, 21) = "skin-custom-setting" & vbCrLf Then
        TxtCode.SelStart = TxtCode.SelStart - 21
        TxtCode.SelLength = 21
        TxtCode.SelText = ""
        TxtCode.SelStart = TxtCode.SelStart + 21
        
        
    End If
    If Not Covered Then ReflashCmdFormat
    Exit Sub

999 End Sub

'''''''''''''''''''''''
'补充命令的反馈
'help
'format格式套half双书签
'cmd字体
