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
If TheNumber = 2 Then
NumToDay = "星期二"
If TheNumber = 3 Then
NumToDay = "星期三"
If TheNumber = 4 Then
NumToDay = "星期四"
If TheNumber = 5 Then
NumToDay = "星期五"
If TheNumber = 6 Then
NumToDay = "星期六"
If TheNumber = 7 Then
NumToDay = "星期日"
End If
End Function

Public Function DayToNum(TheDay As String)
If TheDay = "星期一" Then
DayToNum = 1
If TheDay = "星期二" Then
DayToNum = 2
If TheDay = "星期三" Then
DayToNum = 3
If TheDay = "星期四" Then
DayToNum = 4
If TheDay = "星期五" Then
DayToNum = 5
If TheDay = "星期六" Then
DayToNum = 6
If TheDay = "星期日" Then
DayToNum = 7
End If
End Function
