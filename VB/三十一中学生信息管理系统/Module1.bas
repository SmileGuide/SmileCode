Attribute VB_Name = "Module1"
Public chazhaomoshi As String
Public se As String
Public findout As Integer
Public tim As Integer
Public chazhao As String
Public num As String
Public formn As Integer
Public repaired As Boolean
Public zdd As Boolean
Public stano As Boolean

Sub Main()
If App.PrevInstance Then GoTo 33
App.Title = "三十一中学生信息管理系统"
tipstart = True
welcome.Show
Open "log.log" For Append As #1
Print #1, vbCrLf & vbCrLf & "三十一中学生信息管理系统  " & Trim$(Str$(App.Major)) & "." & Format$(App.Minor, "##00") & "." & Format$(App.Revision, "0000") & " Open " & Now & " Error:" & Err.Number
Close #1
33 End Sub

Function Nosta()
    For ii = 1 To Len(findbox.Text)
    aq = Mid(findbox.Text, ii, 1)
        If aq = "*" Then
            chazhaomoshi = "全部"
            stano = True
        End If
    Next ii
    
End Function

