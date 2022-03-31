VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmCED 
   BackColor       =   &H0000FFFF&
   Caption         =   "高级编辑"
   ClientHeight    =   7734
   ClientLeft      =   24
   ClientTop       =   360
   ClientWidth     =   12252
   Icon            =   "FrmCED.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7734
   ScaleWidth      =   12252
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox LstL 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2184
      ItemData        =   "FrmCED.frx":424A
      Left            =   5400
      List            =   "FrmCED.frx":424C
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2820
      Width           =   1684
   End
   Begin VB.ListBox LstTm 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2184
      ItemData        =   "FrmCED.frx":424E
      Left            =   7560
      List            =   "FrmCED.frx":4250
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2820
      Width           =   2226
   End
   Begin VB.Timer Tmr 
      Interval        =   1
      Left            =   4920
      Top             =   1320
   End
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
      Enabled         =   -1  'True
      TextRTF         =   $"FrmCED.frx":4252
   End
   Begin RichTextLib.RichTextBox TxtCode 
      Height          =   6426
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11166
      _ExtentX        =   19696
      _ExtentY        =   11335
      _Version        =   393217
      BackColor       =   12648447
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"FrmCED.frx":42EF
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





Private Sub Form_Load()
CodeThemeRead
StName = "testing"
On Error GoTo 99
Me.BackColor = SknColor
TxtCode.BackColor = SknColor
TxtCode.Font.Name = TxtFont
FrmCED.Caption = "课程表高级编辑窗口：" & StName & "-" & "[未保存]"
Exit Sub

99 CodeThemeReset
End Sub

Private Sub Form_Paint()
TxtCode.Width = Me.Width
TxtCode.Height = Me.Height - 500

End Sub

Private Sub Tmr_Timer()
        ITRef = GetSetting("SmileTimetable", "Code", "ReflashFormatWhenCrlf", True)

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
            ReflashCmdFormat
            TxtCode.SelStart = Len(TxtCode.Text)
        End If
    Case "quit"
        lstWord = Right(TxtCode.Text, 3)
        If lstWord = "y" & vbCrLf Then
            Saved = True
            Unload Me
        Else
            TxtCode.Text = TxtCvr.Text
                        ReflashCmdFormat
            TxtCode.SelStart = Len(TxtCode.Text)
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

ElseIf Right(LCase(TxtCode.Text), 1) = "." Then
TxtCode.SelStart = TxtCode.SelStart
TxtCode.SelLength = 1
TxtCode.SelColor = CmdColor
TxtCode.SelStart = TxtCode.SelStart

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



ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart, 3) = "*-" Then
TxtCode.SelStart = TxtCode.SelStart - 2
TxtCode.SelLength = 4
                        TxtCode.SelText = ""
TxtCode.SelStart = TxtCode.SelStart + 2

                        TxtCode.SelStart = TxtCode.SelStart - 4
                        TxtCode.SelLength = 4

                        TxtCode.SelStart = TxtCode.SelStart + 4
            ITRef = False
            SaveSetting "SmileTimetable", "Code", "ReflashFormatWhenCrlf", ITRef

ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart, 3) = "*+" Then
    TxtCode.SelStart = TxtCode.SelStart - 2
TxtCode.SelLength = 4
                        TxtCode.SelText = ""
TxtCode.SelStart = TxtCode.SelStart + 2

                        TxtCode.SelStart = TxtCode.SelStart - 4
                        TxtCode.SelLength = 4

                        TxtCode.SelStart = TxtCode.SelStart + 4
            ITRef = True
            SaveSetting "SmileTimetable", "Code", "ReflashFormatWhenCrlf", ITRef



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
        FrmCED.TxtCode.SelColor = Numcolor
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
If Not Covered And CRLFED Then
If ITRef Then ReflashCmdFormat
End If
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
            TxtCode.SelStart = TxtCode.SelStart - 8
        TxtCode.SelLength = 8
        TxtCode.SelText = ""
        TxtCode.SelStart = TxtCode.SelStart + 8
        SampleTxt
        TxtCode.SelStart = Len(TxtCode.Text)






       
    ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 10, 11) = "skin-dark" & vbCrLf Then
        TxtCode.SelStart = TxtCode.SelStart - 11
        TxtCode.SelLength = 11
        TxtCode.SelText = ""
        TxtCode.SelStart = TxtCode.SelStart + 11
        SetTheme &HC0C0C, &HFFFF&, vbWhite, &HFF00FF, &HFF00&, "宋体"
ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 12, 13) = "skin-bright" & vbCrLf Then
        TxtCode.SelStart = TxtCode.SelStart - 13
        TxtCode.SelLength = 13
        TxtCode.SelText = ""
        TxtCode.SelStart = TxtCode.SelStart + 13
        SetTheme &HFFFFFF, &H2469F6, vbBlack, &HFF00FF, &HFF00&, "宋体"
        
ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 12, 13) = "skin-forest" & vbCrLf Then
        TxtCode.SelStart = TxtCode.SelStart - 13
        TxtCode.SelLength = 13
        TxtCode.SelText = ""
        TxtCode.SelStart = TxtCode.SelStart + 13
        SetTheme RGB(107, 255, 192), RGB(255, 251, 17), RGB(0, 85, 242), RGB(253, 114, 79), RGB(231, 128, 93), "幼圆"

ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 12, 13) = "skin-forest" & vbCrLf Then
        TxtCode.SelStart = TxtCode.SelStart - 13
        TxtCode.SelLength = 13
        TxtCode.SelText = ""
        TxtCode.SelStart = TxtCode.SelStart + 13
        SetTheme &HFFFFFF, &H2469F6, vbBlack, &HFF00FF, &HFF00&, "宋体"

ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 12, 13) = "skin-custom" & vbCrLf Then
        TxtCode.SelStart = TxtCode.SelStart - 13
        TxtCode.SelLength = 13
        TxtCode.SelText = ""
        TxtCode.SelStart = TxtCode.SelStart + 13
 

        SknColor = GetSetting("SmileTimetable", "Code", "BgColorCustom")
        SaveSetting "SmileTimetable", "Code", "BgColor", SknColor
        Numcolor = GetSetting("SmileTimetable", "Code", "NumColorCustom")
        SaveSetting "SmileTimetable", "Code", "NumColor", Numcolor
        TxtColor = GetSetting("SmileTimetable", "Code", "TxtColorCustom")
        SaveSetting "SmileTimetable", "Code", "TxtColor", TxtColor
        
        
        
ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 20, 21) = "skin-custom-setting" & vbCrLf Then
        TxtCode.SelStart = TxtCode.SelStart - 21
        TxtCode.SelLength = 21
        TxtCode.SelText = ""
        TxtCode.SelStart = TxtCode.SelStart + 21
        
ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 2, 3) = "." & vbCrLf Then
        ''''''''''''''''''''''''保存

        AllTxt = TxtCode.Text
        AllTxt = Replace(AllTxt, vbCrLf, "")
        AllTxt = Replace(AllTxt, Chr(10), "")

        Dim StStart, StStop, rec
        For rec = 1 To Len(AllTxt)
        If Mid(AllTxt, rec, 1) = "{" Then StStart = rec
        If Mid(AllTxt, rec, 1) = "}" Then StStop = rec
        Next
        AllTxt = Mid(AllTxt, StStart + 1, StStop - StStart - 1) '去{}
        AllTxt = Replace(AllTxt, " ", "")
        AllTxt = LCase(AllTxt)
        Dim DeM
        For DeM = 0 To 9
        AllTxt = Replace(AllTxt, "day" & DeM, "\")
        GruDay = Split(AllTxt, "\")
        Next
        Dim ADCF
        For ADCF = 0 To UBound(GruDay)
        GruDay(ADCF) = Replace(GruDay(ADCF), ";", vbCrLf)
        If Right(GruDay(ADCF), 2) = vbCrLf Then GruDay(ADCF) = Left(GruDay(ADCF), Len(GruDay(ADCF)) - 2)
        Debug.Print GruDay(ADCF)
        Next
        
        
        Dim j
        For j = 1 To 7

        
        Open App.Path & "\SmTab\" & StName & ".smtab" & j For Output As #1
        Print #1, GruDay(j)
        Close #1
        Next j
        
        
        Dim recyy
        For recyy = 1 To 7
        
        LstTm.Clear
        LstL.Clear
        Open App.Path & "\SmTab\" & StName & ".smtab" & recyy For Input As #1
        Dim CTM, CL
        For p = 1 To 100
        If EOF(1) = True Then Exit For
        Input #1, CTM, CL
        LstTm.AddItem CTM
        LstL.AddItem CL
        Next p
        Close #1
        
        

        
        
        Open App.Path & "\SmTab\" & StName & ".smtab" & recyy For Output As #1
        
        For j = 0 To LstTm.ListCount - 1
        rrrrrrrrrrWrite #1, Format(Left(LstTm.List(j), 2), "00") & ":" & Format(Right(LstTm.List(j), 2)), LstL.List(j)
        Next j
        Close #1
        
        
        
        
        LstTm.Clear
        LstL.Clear
        Open App.Path & "\SmTab\" & StName & ".smtab" & recyy For Input As #1
        For p = 1 To 100
        If EOF(1) = True Then Exit For
        Input #1, CTM, CL
        LstTm.AddItem CTM
        LstL.AddItem CL
        Next p
        Close #1
        NumberTbl LstTm, LstL
        
        
        
        
        Open App.Path & "\SmTab\" & StName & ".smtab" & recyy For Output As #1
        
        For j = 0 To LstTm.ListCount - 1
        Write #1, LstTm.List(j), LstL.List(j)
        Next j
        Close #1
        
        Next recyy
        '时间必须00:00

        
        
End If
    If Not Covered Then
    If ITRef Then
    ReflashCmdFormat
    End If
    End If
    Exit Sub

999 End Sub

'''''''''''''''''''''''
'补充命令的反馈
'help
'format格式套half双书签
'cmd字体
