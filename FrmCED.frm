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
   Begin VB.PictureBox Pc 
      Height          =   1326
      Left            =   2820
      ScaleHeight     =   1302
      ScaleWidth      =   1002
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   1026
   End
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      ScrollBars      =   3
      Appearance      =   0
      OLEDropMode     =   0
      TextRTF         =   $"FrmCED.frx":42EF
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "华文中宋"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu Rt 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu Scp 
         Caption         =   "复制"
         Shortcut        =   ^C
      End
      Begin VB.Menu Spt 
         Caption         =   "粘贴"
         Shortcut        =   ^V
      End
      Begin VB.Menu UsS 
         Caption         =   "使用选中的命令"
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu Sp 
         Caption         =   "-"
      End
      Begin VB.Menu SRTF 
         Caption         =   "<save-as-rtf> 保存对话记录为富文本文档"
      End
      Begin VB.Menu STxt 
         Caption         =   "<save-as-text> 保存对话记录为文本文档"
      End
      Begin VB.Menu SScr 
         Caption         =   "<screenshot> 窗口截屏"
      End
   End
End
Attribute VB_Name = "FrmCED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
TxtCode.Font.Size = TxtSize
FrmCED.Caption = "课程表高级编辑窗口：" & StName & "-" & "[未保存]"
FstOpen = True
ITRef = GetSetting("SmileTimetable", "Code", "ReflashFormatWhenCrlf")

Exit Sub

99 CodeThemeReset
End Sub

Private Sub Form_Paint()
TxtCode.Width = Me.Width
TxtCode.Height = Me.Height - 500
If FstOpen Then
TxtCode.Text = "//Coded-Editing-Form @SmileTable,S.G.G. [Vision：" & App.Revision & "]"
TxtCode.SelStart = Len(TxtCode.Text)
TxtCode.Text = TxtCode.Text & vbCrLf
        If ITRef Then ReflashCmdFormat TxtCode
FstOpen = False
TxtCode.SelStart = Len(TxtCode.Text)
End If
End Sub

Private Sub OWord_Updated(Code As Integer)

End Sub

Private Sub Rt_Click()
Dim CLct

CLct = Clipboard.GetText
If CLct = "" Then Spt.Enabled = False Else Spt.Enabled = True
If TxtCode.SelLength > 0 Then
Scp.Enabled = True
UsS.Enabled = True
Else
Scp.Enabled = False
UsS.Enabled = False
End If
End Sub

Private Sub Scp_Click()
Clipboard.Clear
Clipboard.SetText TxtCode.SelText

End Sub

Private Sub Spt_Click()

TxtCode.SelText = Clipboard.GetText
End Sub

Private Sub SRTF_Click()
SARTF
End Sub

Private Sub SScr_Click()
Scr
End Sub

Private Sub STxt_Click()
SATxt
End Sub

Private Sub Tmr_Timer()
        ITRef = GetSetting("SmileTimetable", "Code", "ReflashFormatWhenCrlf", True)
        If TxtCode.Width <> Me.Width Then TxtCode.Width = Me.Width
        If TxtCode.Height <> Me.Height - 500 Then TxtCode.Height = Me.Height - 500
End Sub

Private Sub TxtCode_Change()


On Error Resume Next
Debug.Print Right(LCase(TxtCode.Text), 2)
Dim fstL As String
fstL = Right(TxtCode.Text, 1)
Debug.Print Left(TxtCode.Text, Len(TxtCode.Text) - 1)
If Len(TxtCode.Text) > 50 Then
If (Left(TxtCode.Text, Len(TxtCode.Text) - 3) = "//Coded-Editing-Form @SmileTable,S.G.G. [Vision：" & App.Revision & "]" And Not FstOpen) Or (Left(TxtCode.Text, Len(TxtCode.Text) - 4) = "//Coded-Editing-Form @SmileTable,S.G.G. [Vision：" & App.Revision & "]" And Not FstOpen) Then TxtCode.Text = "": TxtCode.Text = fstL: TxtCode.SelStart = 0: TxtCode.SelLength = 1: TxtCode.SelColor = TxtColor: TxtCode.SelLength = 0
End If
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
            TxtCode.SelText = "//Coded-Editing-Form @SmileTable,S.G.G. [Vision：" & App.Revision & "]"
            FstOpen = False
        Else
            TxtCode.Text = TxtCvr.Text
                   If ITRef Then ReflashCmdFormat TxtCode
            TxtCode.SelStart = Len(TxtCode.Text)
        End If
    Covered = False
                CvrText = ""
    Case "quit"
        lstWord = Right(TxtCode.Text, 3)
        If lstWord = "y" & vbCrLf Then
            Saved = True
            Unload Me
        Else
            TxtCode.Text = TxtCvr.Text
                            If ITRef Then ReflashCmdFormat TxtCode
            TxtCode.SelStart = Len(TxtCode.Text)
        End If
        Covered = False
                    CvrText = ""
    Case "formaterror"
            lstWord = Right(TxtCode.Text, 2)
            If lstWord = vbCrLf Then
            TxtCode.SelStart = Len(TxtCode.Text) - Len(CauText) - 2
            TxtCode.SelLength = Len(CauText)
            TxtCode.SelText = ""
            End If
            Covered = False
                        CvrText = ""
    Case "font"
                FontC = True
                If Right(TxtCode.Text, 2) = vbCrLf Then
                lstWord = Right(TxtCode.Text, 10)
                If lstWord Like "*>>*" & vbCrLf Then
                TxtCode.Text = TxtCvr.Text

                MxFTS = Split(lstWord, ">>")
                MxFTS(1) = Replace(MxFTS(1), vbCrLf, "")
                TxtSize = MxFTS(1)
                SaveSetting "SmileTimetable", "Code", "FontSize", TxtSize
                TxtCode.SelStart = 0
                TxtCode.SelLength = Len(TxtCode.Text)
                FrmCED.TxtCode.Font.Size = TxtSize
                       If ITRef Then ReflashCmdFormat TxtCode
                        TxtCode.SelStart = Len(TxtCode.Text)
                End If
                Covered = False
                                            CvrText = ""
                End If

    Case "rtf", "screenshot"
                TxtCode.Text = TxtCvr.Text
                   If ITRef Then ReflashCmdFormat TxtCode
            TxtCode.SelStart = Len(TxtCode.Text)
                Covered = False
                            CvrText = ""
    Case "help"
                lstWord = Right(TxtCode.Text, 1)
                If lstWord = "y" Then
            TxtCode.Text = TxtCvr.Text
                   If ITRef Then ReflashCmdFormat TxtCode
            TxtCode.SelStart = Len(TxtCode.Text)

                End If
                                Covered = False
                                            CvrText = ""
    End Select

            
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
If UsSed Then GoTo 45
If Mid(TxtCode.Text, TxtCode.SelStart, 2) = Chr(10) Or Mid(TxtCode.Text, TxtCode.SelStart, 1) = Chr(10) Then CRLFED = True: GoTo 45
CRLFED = False
33 TxtCode.SelStart = TxtCode.SelStart - 1
TxtCode.SelLength = 1
TxtCode.SelColor = TxtColor





If Right(LCase(TxtCode.Text), 1) = "." Then
TxtCode.SelStart = TxtCode.SelStart
TxtCode.SelLength = 1
TxtCode.SelColor = CmdColor
TxtCode.SelStart = TxtCode.SelStart

ElseIf Right(LCase(TxtCode.Text), 2) = "*-" Then

TxtCode.Text = Left(TxtCode.Text, Len(TxtCode.Text) - 2)
TxtCode.SelStart = Len(TxtCode.Text)

            ITRef = False
            SaveSetting "SmileTimetable", "Code", "ReflashFormatWhenCrlf", ITRef
TxtCode.SelStart = 0
TxtCode.SelLength = Len(TxtCode.Text)
TxtCode.SelColor = TxtColor
TxtCode.SelStart = Len(TxtCode.Text)
ElseIf Right(LCase(TxtCode.Text), 2) = "*+" Then
TxtCode.Text = Left(TxtCode.Text, Len(TxtCode.Text) - 2)
TxtCode.SelStart = Len(TxtCode.Text)

            ITRef = True
            SaveSetting "SmileTimetable", "Code", "ReflashFormatWhenCrlf", ITRef
         ReflashCmdFormat TxtCode


ElseIf Right(LCase(TxtCode.Text), 2) = "**" Then
TxtCode.Text = Left(TxtCode.Text, Len(TxtCode.Text) - 2)

      ReflashCmdFormat TxtCode
      TxtCode.SelStart = Len(TxtCode.Text)
'        Dim pp
'        pp = TxtCode.SelStart
'        FrmCED.TxtCode.SelStart = 0
'        FrmCED.TxtCode.SelLength = Len(FrmCED.TxtCode.Text)
'        FrmCED.TxtCode.SelColor = TxtColor
'
'        Dim i
'        For i = 0 To Len(FrmCED.TxtCode.Text) - 1
'        FrmCED.TxtCode.SelStart = i
'        FrmCED.TxtCode.SelLength = 1
'
'        Select Case FrmCED.TxtCode.SelText
'        Case Chr(34), Chr(44), Chr(58)
'        FrmCED.TxtCode.SelColor = SpcColor
'
'        Case Chr(46), Chr(59), Chr(123), Chr(125)
'        FrmCED.TxtCode.SelColor = CmdColor
'        Case 0 To 9
'        FrmCED.TxtCode.SelColor = Numcolor
'        End Select
'
'        FrmCED.TxtCode.SelLength = 2
'        If FrmCED.TxtCode.SelText = "**" Then FrmCED.TxtCode.SelText = ""
'
'        FrmCED.TxtCode.SelLength = 3
'        Select Case FrmCED.TxtCode.SelText
'        Case "day"
'        FrmCED.TxtCode.SelColor = CmdColor
'        End Select
'
'        FrmCED.TxtCode.SelLength = 5
'        Select Case FrmCED.TxtCode.SelText
'        Case "clear"
'        FrmCED.TxtCode.SelColor = CmdColor
'        End Select
'
'
'        Next
'        FrmCED.TxtCode.SelLength = 0
'        TxtCode.SelStart = pp

ElseIf Right(LCase(TxtCode.Text), 5) = "clear" Then
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

ElseIf Right(LCase(TxtCode.Text), 4) = "font" Then
TxtCode.SelStart = TxtCode.SelStart - 3
TxtCode.SelLength = 4
TxtCode.SelColor = CmdColor
TxtCode.SelStart = TxtCode.SelStart + 3

ElseIf Right(LCase(TxtCode.Text), 6) = "format" Then
TxtCode.SelStart = TxtCode.SelStart - 5
TxtCode.SelLength = 6
TxtCode.SelColor = CmdColor
TxtCode.SelStart = TxtCode.SelStart + 5

ElseIf Right(LCase(TxtCode.Text), 7) = "example" Or Right(LCase(TxtCode.Text), 7) = "preview" Then
TxtCode.SelStart = TxtCode.SelStart - 6
TxtCode.SelLength = 7
TxtCode.SelColor = CmdColor
TxtCode.SelStart = TxtCode.SelStart + 6

ElseIf Right(LCase(TxtCode.Text), 9) = "skin-dark" Then
TxtCode.SelStart = TxtCode.SelStart - 8
TxtCode.SelLength = 9
TxtCode.SelColor = CmdColor
TxtCode.SelStart = TxtCode.SelStart + 8

ElseIf Right(LCase(TxtCode.Text), 10) = "screenshot" Then
TxtCode.SelStart = TxtCode.SelStart - 9
TxtCode.SelLength = 10
TxtCode.SelColor = CmdColor
TxtCode.SelStart = TxtCode.SelStart + 9


ElseIf Right(LCase(TxtCode.Text), 11) = "save-as-rtf" Then
TxtCode.SelStart = TxtCode.SelStart - 10
TxtCode.SelLength = 11
TxtCode.SelColor = CmdColor
TxtCode.SelStart = TxtCode.SelStart + 10




77 ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart + 1, 1) = "*" Then
TxtCode.SelLength = 1
TxtCode.SelColor = CmdColor
TxtCode.SelStart = TxtCode.SelStart + 1



ElseIf Right(LCase(TxtCode.Text), 11) = "skin-bright" Then
TxtCode.SelStart = TxtCode.SelStart - 10
TxtCode.SelLength = 11
TxtCode.SelColor = CmdColor
TxtCode.SelStart = TxtCode.SelStart + 10

ElseIf Right(LCase(TxtCode.Text), 11) = "skin-custom" Or Right(LCase(TxtCode.Text), 11) = "skin-forest" Then
TxtCode.SelStart = TxtCode.SelStart - 10
TxtCode.SelLength = 11
TxtCode.SelColor = CmdColor
TxtCode.SelStart = TxtCode.SelStart + 10

End If

'贴尾归位
TxtCode.SelStart = TxtCode.SelStart + 1
99 TxtCode.SelLength = 0
If Not Covered And CRLFED Then
If ITRef Then ReflashCmdFormat TxtCode
End If
Exit Sub





45 If UsSed Then UsSed = False: TxtCode.SelStart = Len(TxtCode.Text)
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
        ReflashCmdFormat TxtCode
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
        
        
ElseIf Right(LCase(TxtCode.Text), 2) = "*-" Then

TxtCode.Text = Left(TxtCode.Text, Len(TxtCode.Text) - 2)
TxtCode.SelStart = Len(TxtCode.Text)

            ITRef = False
            SaveSetting "SmileTimetable", "Code", "ReflashFormatWhenCrlf", ITRef
TxtCode.SelStart = 0
TxtCode.SelLength = Len(TxtCode.Text)
TxtCode.SelColor = TxtColor
TxtCode.SelStart = Len(TxtCode.Text)
ElseIf Right(LCase(TxtCode.Text), 2) = "*+" Then
TxtCode.Text = Left(TxtCode.Text, Len(TxtCode.Text) - 2)
TxtCode.SelStart = Len(TxtCode.Text)

            ITRef = True
            SaveSetting "SmileTimetable", "Code", "ReflashFormatWhenCrlf", ITRef
         ReflashCmdFormat TxtCode


ElseIf Right(LCase(TxtCode.Text), 2) = "**" Then
TxtCode.Text = Left(TxtCode.Text, Len(TxtCode.Text) - 2)

      ReflashCmdFormat TxtCode
      TxtCode.SelStart = Len(TxtCode.Text)
        
        
        
        
ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 5, 6) = "font" & vbCrLf Then
        CvrText = "font"
        UnDis = False
        TxtCvr.Text = TxtCode.Text
        TxtCvr.Text = Left(TxtCvr.Text, Len(TxtCvr.Text) - 6)

        UnDis = True
        
        CauText = "请输入要设置的字号。>>"
        TxtCode.Text = TxtCode.Text & vbCrLf & CauText
        ReflashCmdFormat TxtCode
        Covered = True
        TxtCode.SelStart = Len(TxtCode.Text)
ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 5, 6) = "help" & vbCrLf Then
                TxtCvr.Text = TxtCode.Text
        TxtCvr.Text = Left(TxtCvr.Text, Len(TxtCvr.Text) - 6)
        Open App.Path & "\help.txt" For Input As #8
        Dim HelpCon, HelpAdd
        
        Do Until EOF(8)
        Line Input #8, HelpAdd
        
        HelpCon = HelpCon & HelpAdd & vbCrLf
        Loop
        Close #8
        HelpCon = HelpCon & vbCrLf & "输入y来隐藏帮助。>>"
        TxtCode.Text = TxtCode.Text & vbCrLf & HelpCon


        
        If ITRef Then ReflashCmdFormat TxtCode Else EnWhite
        
        CvrText = "help"
        Covered = True

        TxtCode.SelStart = Len(TxtCode.Text)
        
        
        
    ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 6, 7) = "clear" & vbCrLf Then
        CvrText = "clear"
        UnDis = False
        TxtCvr.Text = TxtCode.Text
        TxtCvr.Text = Left(TxtCvr.Text, Len(TxtCvr.Text) - 7)

        UnDis = True
        
        CauText = "警告：此操作是不可逆的。" & Chr(10) & "确定要执行此命令，请输入y后回车；撤回此命令，请连续按下回车。>>"
        TxtCode.Text = TxtCode.Text & vbCrLf & CauText
                If ITRef Then ReflashCmdFormat TxtCode
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
        
ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 8, 9) = "example" & vbCrLf Then
            TxtCode.SelStart = TxtCode.SelStart - 9
        TxtCode.SelLength = 8
        TxtCode.SelText = ""
        TxtCode.SelStart = TxtCode.SelStart + 9
        EgTxt
        TxtCode.SelStart = Len(TxtCode.Text)
        
ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 8, 9) = "preview" & vbCrLf Then
            TxtCode.SelStart = TxtCode.SelStart - 9
        TxtCode.SelLength = 8
        TxtCode.SelText = ""
        Open App.Path & "\" & StName & "预览.txt" For Output As #16
        AllTxt = TxtCode.Text
        For rec = 1 To Len(AllTxt)
        If Mid(AllTxt, rec, 1) = "{" Then StStart = rec
        If Mid(AllTxt, rec, 1) = "}" Then StStop = rec
        Next
                AllTxt = Mid(AllTxt, StStart, StStop - StStart + 1)
        Spa = Split(AllTxt, vbCrLf)
        

        Dim iii
        For iii = 0 To 1000
        On Error Resume Next
        Print #16, Spa(iii)
        Next iii
    
        Close #16
        
                TxtCode.Text = TxtCode.Text & vbCrLf & "保存成功。" & vbCrLf & "按任意键来继续。>>"
        Shell "explorer /n,/select," & App.Path & "\" & StName & "预览.txt", vbNormalFocus
        TxtCode.SelStart = Len(TxtCode.Text)
               If ITRef Then ReflashCmdFormat TxtCode Else EnWhite
        TxtCode.SelStart = Len(TxtCode.Text) - 17
        TxtCode.SelLength = 5
        TxtCode.SelColor = CmdColor
                TxtCode.SelStart = Len(TxtCode.Text)
        CvrText = "text"
        Covered = True
    ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 10, 11) = "skin-dark" & vbCrLf Then
        TxtCode.SelStart = TxtCode.SelStart - 11
        TxtCode.SelLength = 11
        TxtCode.SelText = ""
        TxtCode.SelStart = TxtCode.SelStart + 11
        SetTheme &HC0C0C, &HFFFF&, vbWhite, &HFF00FF, &HFF00&, "宋体"

ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 11, 12) = "screenshot" & vbCrLf Then
Scr




       

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
        
        ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 12, 13) = "save-as-rtf" & vbCrLf Then

        SARTF
ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 13, 14) = "save-as-text" & vbCrLf Then

        SATxt
ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 20, 21) = "skin-custom-setting" & vbCrLf Then
        TxtCode.SelStart = TxtCode.SelStart - 21
        TxtCode.SelLength = 21
        TxtCode.SelText = ""
        TxtCode.SelStart = TxtCode.SelStart + 21
        
ElseIf Mid(LCase(TxtCode.Text), TxtCode.SelStart - 2, 3) = "." & vbCrLf Then
        ''''''''''''''''''''''''保存
       ' On Error GoTo SynFail
        FrmCED.Caption = "课程表高级编辑窗口：" & StName & "-" & "[已保存]"
        AllTxt = TxtCode.Text
        AllTxt = Replace(AllTxt, vbCrLf, "")
        AllTxt = Replace(AllTxt, Chr(10), "")

        
        For rec = 1 To Len(AllTxt)
        If Mid(AllTxt, rec, 1) = "{" Then StStart = rec
        If Mid(AllTxt, rec, 1) = "}" Then StStop = rec
        Next
        AllTxt = Mid(AllTxt, StStart + 1, StStop - StStart - 1) '去{}
        AllTxt = Replace(AllTxt, " ", "")
        Dim DeM
        For DeM = 0 To 9
        AllTxt = Replace(AllTxt, "day" & DeM, "\")

        GruDay = Split(AllTxt, "\")
        Next
                AllTxt = Replace(AlText, "none", "")
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
        Dim SplTm As Variant
        For j = 0 To LstTm.ListCount - 1
        '下句几乎写完了
        SplTm = Split(LstTm.List(j), ":")
        Write #1, Format(SplTm(0), "00") & ":" & Format(SplTm(1), "00"), LstL.List(j)
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


        
        
End If
    If Not Covered Then
    If ITRef Then
            If ITRef Then ReflashCmdFormat TxtCode
    End If
    End If
    Exit Sub
SynFail:        CauText = "错误：对该代码的语法审查不合格，无法保存为课程表。" & vbCrLf & "      若要将该代码以文本文档形式存储，请使用preview命令。" & vbCrLf & "      使用help命令获取更多帮助。" & Chr(10) & "请按回车键表示确定。>>"
        TxtCode.Text = TxtCode.Text & vbCrLf & CauText
                If ITRef Then ReflashCmdFormat TxtCode
        TxtCode.SelStart = Len(TxtCode.Text) - Len(CauText)
        TxtCode.SelLength = Len(CauText)
        TxtCode.SelColor = TxtColor
        TxtCode.SelLength = 0
        TxtCode.SelStart = Len(TxtCode.Text) - Len(CauText)
        TxtCode.SelLength = Len(CauText) - 74
        TxtCode.SelColor = vbRed
        TxtCode.SelStart = Len(TxtCode.Text)
        TxtCode.SelLength = 0
        Covered = True
        CvrText = "formaterror"
999 End Sub

'''''''''''''''''''''''
'补充命令的反馈
'help
'cmd字号、禁止拖放
'打开文件  文件名



''''''
'standard Description
'{
'Day1
'"14:20","geography";
'"2:30","maths";
'"12:50","physics";
'Day2
'"12:50","physics";
'Day 3
'"14:20","geography";
'day4
'day5
'day6
'day7
'}



Private Sub TxtCode_GotFocus()

End Sub

Private Sub TxtCode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Rt
End Sub

Public Function SARTF()
TxtCvr.Text = TxtCode.Text
TxtCvr.Text = Left(TxtCvr.Text, Len(TxtCvr.Text) - 13)



        Numcolor = &H2469F6
        SknColor = &HFFFFFF
        TxtColor = vbBlack
        SpcColor = &HFF00FF
        CmdColor = &HFF00&
        TxtCvr.Font = "宋体"
        ReflashCmdFormat TxtCvr

        Numcolor = GetSetting("SmileTimetable", "Code", "NumColor")
        SknColor = GetSetting("SmileTimetable", "Code", "BgColor")
        TxtColor = GetSetting("SmileTimetable", "Code", "TxtColor")
        SpcColor = GetSetting("SmileTimetable", "Code", "SpecialColor")
        CmdColor = GetSetting("SmileTimetable", "Code", "CommandColor")

Dim FileN
FileN = App.Path & "\课程表代码记录_" & Format(Now, "yyyymmddhhmmss") & ".rtf"
        TxtCvr.SaveFile FileN

        TxtCvr.Text = TxtCode.Text
        TxtCvr.Text = Left(TxtCvr.Text, Len(TxtCvr.Text) - 13)
        TxtCode.Text = TxtCode.Text & vbCrLf & "保存成功。" & vbCrLf & "由于富文本文档的背景颜色限制，课程表代码记录以bright主题保存。" & vbCrLf & "按任意键来继续。>>"
        Shell "explorer /n,/select," & FileN, vbNormalFocus
        TxtCode.SelStart = Len(TxtCode.Text)
               If ITRef Then ReflashCmdFormat TxtCode Else EnWhite
        TxtCode.SelStart = Len(TxtCode.Text) - 53
        TxtCode.SelLength = 5
        TxtCode.SelColor = CmdColor
                TxtCode.SelStart = Len(TxtCode.Text)
        CvrText = "rtf"
        Covered = True
End Function

Public Function SATxt()
TxtCvr.Text = TxtCode.Text
TxtCvr.Text = Left(TxtCvr.Text, Len(TxtCvr.Text) - 14)


Dim FileNT
FileNT = App.Path & "\课程表代码记录_" & Format(Now, "yyyymmddhhmmss") & ".txt"
        Open FileNT For Output As #77
        Print #77, TxtCode.Text
        Close #77

        TxtCvr.Text = TxtCode.Text
        TxtCvr.Text = Left(TxtCvr.Text, Len(TxtCvr.Text) - 14)
        TxtCode.Text = TxtCode.Text & vbCrLf & "保存成功。" & vbCrLf & "按任意键来继续。>>"
        Shell "explorer /n,/select," & FileNT, vbNormalFocus
        TxtCode.SelStart = Len(TxtCode.Text)
               If ITRef Then ReflashCmdFormat TxtCode Else EnWhite
        TxtCode.SelStart = Len(TxtCode.Text) - 17
        TxtCode.SelLength = 5
        TxtCode.SelColor = CmdColor
                TxtCode.SelStart = Len(TxtCode.Text)
        CvrText = "text"
        Covered = True
End Function

Public Sub Scr()
On Error GoTo agn
TxtCvr.Text = TxtCode.Text
TxtCvr.Text = Left(TxtCvr.Text, Len(TxtCvr.Text) - 12)
TxtCvr.Text = TxtCode.Text
        TxtCvr.Text = Left(TxtCvr.Text, Len(TxtCvr.Text) - 12)
        
            On Error GoTo agn
agn:        Call keybd_event(vbKeySnapshot, 1, 0, 0)
        On Error GoTo agn
        Pc.Picture = Clipboard.GetData(vbCFBitmap)
        Clipboard.Clear
        Dim FileNP
        On Error GoTo agn
        FileNP = App.Path & "\课程表代码记录_" & Format(Now, "yyyymmddhhmmss") & ".bmp"
        On Error GoTo agn
        SavePicture Pc.Picture, FileNP
     On Error GoTo agn
        TxtCode.Text = TxtCode.Text & vbCrLf & "保存成功。" & vbCrLf & "按任意键来继续。>>"
        On Error GoTo agn
                Shell "explorer /n,/select," & FileNP, vbNormalFocus
        TxtCode.SelStart = Len(TxtCode.Text)
               If ITRef Then ReflashCmdFormat TxtCode Else EnWhite
        TxtCode.SelStart = Len(TxtCode.Text) - 17
        TxtCode.SelLength = 5
        TxtCode.SelColor = CmdColor
                TxtCode.SelStart = Len(TxtCode.Text)
        CvrText = "screenshot"
        Covered = True
End Sub

Private Sub UsS_Click()
UsSed = True
If TxtCode.SelText <> "**" And TxtCode.SelText <> "*+" And TxtCode.SelText <> "*-" Then TxtCode.Text = TxtCode.Text & TxtCode.SelText & vbCrLf Else TxtCode.Text = TxtCode.Text & TxtCode.SelText



End Sub
