VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "使用统计"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6780
   Icon            =   "statistics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   6780
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6285
      Top             =   15
   End
   Begin VB.ListBox op 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1200
      ItemData        =   "statistics.frx":0ECA
      Left            =   15
      List            =   "statistics.frx":0EEC
      TabIndex        =   0
      Top             =   0
      Width           =   2340
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1275
      Left            =   2325
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   -15
      Width           =   3105
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "次"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5610
      TabIndex        =   2
      Top             =   180
      Width           =   1380
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con(1000000) As String
Private Sub Form_Load()
Open "log.log" For Input As #1
        Do Until EOF(1)
        nu = nu + 1
        Line Input #1, con(nu)
        DoEvents
        Loop
         Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Open "Log-S.ini" For Output As #1
Print #1, 0
Close #1
End Sub

Private Sub op_Click()
If   = 0 Then
ElseIf opt = 1 Then
ElseIf opt = 2 Then
ElseIf opt = 3 Then
ElseIf opt = 4 Then
ElseIf opt = 5 Then
End If

'————————————————————
op.Text = "..."
le = 0
For al = 1 To UBound(con) - 1
If con(al) = "" Then GoTo ex
For j = 0 To 6
If op.Selected(j) = True Then opt = i
Next
If opt = 0 Then
    For i = 1 To Len(con(al)) - 3
    w = Mid(con(al), i, 4)
    If w = "Open" Then le = le + 1
    Next i
ElseIf opt = 1 Then
    For a = 1 To Len(con(al)) - 2
    w = Mid(con(al), a, 3)
    If w = "Add" Then le = le + 1
    Next a
ElseIf opt = 2 Then
    For b = 1 To Len(con(al)) - 5
    w = Mid(con(al), b, 6)
    If w = "Delete" Then le = le + 1
    Next b
ElseIf opt = 3 Then
    For c = 1 To Len(con(al)) - 3
    w = Mid(con(al), c, 4)
    If w = "Make" Then le = le + 1
    Next c
ElseIf opt = 4 Then
    For d = 1 To Len(con(al)) - 3
    w = Mid(con(al), d, 4)
    If w = "Move" Then le = le + 1
    Next d
ElseIf opt = 5 Then
    For e = 1 To Len(con(al)) - 6
    w = Mid(con(al), e, 7)
    If w = "Changed" Then le = le + 1
    Next e
End If
Next al
ex: Text1.Text = le
End Sub

