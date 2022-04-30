Attribute VB_Name = "MdlFunc"
Private Type TGUID
   Data1                            As Long
   Data2                            As Integer
   Data3                            As Integer
   Data4(0 To 7)                    As Byte
End Type
 

Private Declare Function OleLoadPicturePath Lib "oleaut32.dll" (ByVal szURLorPath As Long, ByVal punkCaller As Long, ByVal dwReserved As Long, ByVal clrReserved As OLE_COLOR, ByRef riid As TGUID, ByRef ppvRet As IPicture) As Long
 

Public Function LoadPicture(ByVal strFileName As String) As Picture
   Dim IID  As TGUID
   With IID
      .Data1 = &H7BF80980
      .Data2 = &HBF32
      .Data3 = &H101A
      .Data4(0) = &H8B
      .Data4(1) = &HBB
      .Data4(2) = &H0
      .Data4(3) = &HAA
      .Data4(4) = &H0
      .Data4(5) = &H30
      .Data4(6) = &HC
      .Data4(7) = &HAB
   End With
   
   On Error GoTo LocalErr
   
   OleLoadPicturePath StrPtr(strFileName), 0&, 0&, 0&, IID, LoadPicture
   Exit Function
LocalErr:
   Set LoadPicture = VB.LoadPicture(strFileName)
   Err.Clear
End Function
 



Public Function IEWord()
excel.Workbooks.Open (App.Path & "\word\word.xls")
Dim a As String
excel.ActiveWorkbook.RefreshAll
DWord = excel.Application.Cells(1, 1)
DFrom = excel.Application.Cells(3, 1)
excel.ActiveWorkbook.Saved = True
excel.ActiveWorkbook.Close
IEWord = DWord & Chr(10) & "——" & DFrom
End Function

Public Function Msg(MsgText As String, MsgColor As String, ShowTime As Integer)
FrmMsg.LblText.Caption = MsgText
FrmMsg.BackColor = MsgColor
FrmMsg.Tmr.Interval = ShowTime
FrmMsg.Show
End Function


Public Function EnMiddle(Window As Form)
Window.Move (Screen.Width - Window.Width) / 2, (Screen.Height - Window.Height) / 2
End Function




Public Function NumberTbl(TimeColumn As ListBox, Lessoncolumn As ListBox)

Dim Cnt(0 To 100) As String
Dim Cpr(0 To 100) As String
Dim OutLst(0 To 101) As String
Dim a, B, i, j, UBCNT, Mdl, MdlO, m, w, g
Dim Is0 As Boolean
For g = 0 To TimeColumn.ListCount - 1
If TimeColumn.List(g) = "00:00" Then
Is0 = True
Open App.Path & "\~temp0.tmp" For Output As #3
Write #3, TimeColumn.List(g), Lessoncolumn.List(g)
Close #3
Lessoncolumn.RemoveItem g
End If
Next
Open App.Path & "\~temp.tmp" For Output As #1
Dim t
For t = 0 To TimeColumn.ListCount - 1
Write #1, TimeColumn.List(t), Lessoncolumn.List(t)
Next
Close #1
For a = 0 To TimeColumn.ListCount - 1
Cpr(a) = "00:00"
Cnt(a) = TimeColumn.List(a)
Next

UBCNT = UBound(Cnt)
For i = 0 To UBCNT - 1
Mdl = "00:00"
    For j = 0 To UBound(Cnt)
    If Mdl < Cnt(j) Then
        Mdl = Cnt(j)
        MdlO = j
    End If
    Next j
Cpr(i) = Mdl
Cnt(MdlO) = ""
Next i
w = 0
For m = UBound(Cpr) To 0 Step -1
w = w + 1
OutLst(w) = Cpr(m)
Next
For B = 0 To UBound(OutLst)
TimeColumn.List(B) = OutLst(B)
Next
Dim del0
For del0 = TimeColumn.ListCount - 1 To 0 Step -1

If TimeColumn.List(del0) = "" Or TimeColumn.List(del0) = "00:00" Then TimeColumn.RemoveItem del0
Next
Dim TmTmp(0 To 100)
Open App.Path & "\~temp.tmp" For Input As #2
Dim LEX(0 To 100)
For t = 0 To TimeColumn.ListCount - 1
Input #2, TmTmp(t), LEX(t)
Next
Close #2
Dim h, v
For h = 0 To TimeColumn.ListCount - 1
 For v = 0 To TimeColumn.ListCount - 1
 If TimeColumn.List(h) = TmTmp(v) Then
    Lessoncolumn.List(h) = LEX(v)
End If
Next v
Next h
If Is0 Then
Open App.Path & "\~temp0.tmp" For Input As #4
Dim zero, zerotxt
Input #4, zero, zerotxt
Close #4
TimeColumn.AddItem zero, 0
Lessoncolumn.AddItem zerotxt, 0
End If
End Function


Public Function NumToDay(TheNumber As Integer)
If TheNumber = 1 Then
NumToDay = "星期一"
ElseIf TheNumber = 2 Then
NumToDay = "星期二"
ElseIf TheNumber = 3 Then
NumToDay = "星期三"
ElseIf TheNumber = 4 Then
NumToDay = "星期四"
ElseIf TheNumber = 5 Then
NumToDay = "星期五"
ElseIf TheNumber = 6 Then
NumToDay = "星期六"
ElseIf TheNumber = 7 Then
NumToDay = "星期日"
Else
numday = "ERROR"
End If
End Function

Public Function DayToNum(TheDay As String)
If TheDay = "星期一" Then
DayToNum = 1
ElseIf TheDay = "星期二" Then
DayToNum = 2
ElseIf TheDay = "星期三" Then
DayToNum = 3
ElseIf TheDay = "星期四" Then
DayToNum = 4
ElseIf TheDay = "星期五" Then
DayToNum = 5
ElseIf TheDay = "星期六" Then
DayToNum = 6
ElseIf TheDay = "星期日" Then
DayToNum = 7
Else
DayToNum = 0
End If
End Function



Public Function CodeThemeRead()
SknColor = GetSetting("SmileTimetable", "Code", "BgColor")
Numcolor = GetSetting("SmileTimetable", "Code", "NumColor")
TxtColor = GetSetting("SmileTimetable", "Code", "TxtColor")
SpcColor = GetSetting("SmileTimetable", "Code", "SpecialColor")
CmdColor = GetSetting("SmileTimetable", "Code", "CommandColor")
TxtFont = GetSetting("SmileTimetable", "Code", "TxtFont")
TxtSize = GetSetting("SmileTimetable", "Code", "FontSize")

End Function

Public Function CodeThemeReset()
SaveSetting "SmileTimetable", "Code", "BgColor", &HC0C0C
SaveSetting "SmileTimetable", "Code", "NumColor", &HFFFF&
SaveSetting "SmileTimetable", "Code", "TxtColor", &HFFFFFF
SaveSetting "SmileTimetable", "Code", "SpecialColor", &HFF00FF
SaveSetting "SmileTimetable", "Code", "CommandColor", &HFF00&
SaveSetting "SmileTimetable", "Code", "TxtFont", "宋体"
SaveSetting "SmileTimetable", "Code", "FontSize", 24
End Function

Public Function ReflashCmdFormat()
Dim ppp
ppp = FrmCED.TxtCode.SelStart
On Error GoTo 77
FrmCED.TxtCode.SelStart = 0
FrmCED.TxtCode.SelLength = Len(FrmCED.TxtCode.Text)
FrmCED.TxtCode.SelColor = TxtColor

Dim i
For i = 0 To Len(FrmCED.TxtCode.Text) - 1
FrmCED.TxtCode.SelStart = i
FrmCED.TxtCode.SelLength = 1
On Error Resume Next

On Error GoTo 77
Select Case FrmCED.TxtCode.SelText
Case Chr(34), Chr(44), Chr(58)
FrmCED.TxtCode.SelColor = SpcColor
Case Chr(46), Chr(59), Chr(123), Chr(125), "【", "】", "<", ">", "§"
FrmCED.TxtCode.SelColor = CmdColor
Case 0 To 9
FrmCED.TxtCode.SelColor = Numcolor
End Select

FrmCED.TxtCode.SelLength = 2
Select Case FrmCED.TxtCode.SelText
Case "**", "*+", "*-"
FrmCED.TxtCode.SelColor = CmdColor
End Select

FrmCED.TxtCode.SelLength = 3
Select Case FrmCED.TxtCode.SelText
Case "day"
FrmCED.TxtCode.SelColor = CmdColor
End Select

FrmCED.TxtCode.SelLength = 4
Select Case FrmCED.TxtCode.SelText
Case "help", "quit"
FrmCED.TxtCode.SelColor = CmdColor
End Select

FrmCED.TxtCode.SelLength = 5
Select Case FrmCED.TxtCode.SelText
Case "clear"
FrmCED.TxtCode.SelColor = CmdColor
End Select

FrmCED.TxtCode.SelLength = 6
Select Case FrmCED.TxtCode.SelText
Case "format"
FrmCED.TxtCode.SelColor = CmdColor
End Select
If FstOpen Then
FrmCED.TxtCode.SelLength = 51
Select Case Replace(FrmCED.TxtCode.SelText, " ", "")

Case Replace("//Coded-Editing-Form @SmileTable,S.G.G. [Vision：" & App.Revision & "]", " ", "")
FrmCED.TxtCode.SelColor = CmdColor
Exit Function
End Select
End If
Next
FrmCED.TxtCode.SelLength = 0
FrmCED.TxtCode.SelStart = Len(FrmCED.TxtCode.Text) + 1
FrmCED.TxtCode.SelStart = ppp
Exit Function
77 FrmCED.TxtCode.SelStart = ppp
End Function



Public Function SampleTxt()
'"59chr
SampleTxt = "{" & "day1" & ";" & vbCrLf & "day2" & ";" & vbCrLf & "day3" & ";" & vbCrLf & "day4" & ";" & vbCrLf & "day5" & ";" & vbCrLf & "day6" & ";" & vbCrLf & "day7" & "}"

FrmCED.TxtCode.Text = FrmCED.TxtCode.Text & SampleTxt
ReflashCmdFormat
End Function



Public Function SetTheme(sBgColor As String, sNumcolor As String, sTxtColor As String, sSpecialColor As String, sCommandColor As String, sTxtFont As String)
        SaveSetting "SmileTimetable", "Code", "BgColor", sBgColor
        SaveSetting "SmileTimetable", "Code", "NumColor", sNumcolor
        SaveSetting "SmileTimetable", "Code", "TxtColor", sTxtColor
        SaveSetting "SmileTimetable", "Code", "SpecialColor", sSpecialColor
        SaveSetting "SmileTimetable", "Code", "CommandColor", sCommandColor
        SaveSetting "SmileTimetable", "Code", "TxtFont", sTxtFont
        Numcolor = GetSetting("SmileTimetable", "Code", "NumColor")
        SknColor = GetSetting("SmileTimetable", "Code", "BgColor")
        TxtColor = GetSetting("SmileTimetable", "Code", "TxtColor")
        SpcColor = GetSetting("SmileTimetable", "Code", "SpecialColor")
        CmdColor = GetSetting("SmileTimetable", "Code", "CommandColor")
        FrmCED.BackColor = SknColor
        FrmCED.TxtCode.BackColor = SknColor
        FrmCED.TxtCode.Font = sTxtFont
        ReflashCmdFormat
End Function
