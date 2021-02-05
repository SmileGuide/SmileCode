VERSION 5.00
Begin VB.Form letter 
   BackColor       =   &H00C0FFFF&
   Caption         =   "消息"
   ClientHeight    =   6885
   ClientLeft      =   5700
   ClientTop       =   4170
   ClientWidth     =   7110
   FillColor       =   &H00C0FFFF&
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   15.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "letter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   68850
   ScaleMode       =   0  'User
   ScaleWidth      =   7110
   Begin VB.ListBox List1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   6495
      ItemData        =   "letter.frx":1CCA
      Left            =   0
      List            =   "letter.frx":1CCC
      MouseIcon       =   "letter.frx":1CCE
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "（无消息）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
   End
End
Attribute VB_Name = "letter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Open "F:\BBSeed Files\hopestate.bdf" For Input As #1
Input #1, a
Close #1
Open "F:\BBSeed Files\foodstate.bdf" For Input As #3
Input #3, c
Close #3
Open "F:\BBSeed Files\fd.bdf" For Input As #5
Input #5, e
Close #5
If a = 0 Then
ElseIf a = 1 Then
List1.AddItem "小抱抱同意了你的愿望请求"
ElseIf a = 2 Then
List1.AddItem "小抱抱驳回了你的愿望请求"
End If
If c = 0 Then
GoTo 333
ElseIf c = 1 Then
If List1.Columns <> 1 Then
List1.AddItem "小抱抱给你了" & e & "kg养料"
End If
End If
333 End Sub
Private Sub Form_Unload(Cancel As Integer)
main.Show
End Sub

Private Sub List1_Click()
letter.Hide
If List1.Columns = 2 Then
If List1.Selected(0) = True Then
hopecon.Show
List1.RemoveItem 0
Else
Call foodp
List1.RemoveItem 1
End If
Else
If Right(List1.List(0), 2) = "养料" Then
List1.RemoveItem 0
hopecon.Show
Else
List1.RemoveItem 0
food.Show
End If
End If
End Sub
