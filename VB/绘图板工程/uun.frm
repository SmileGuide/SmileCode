VERSION 5.00
Begin VB.Form uun 
   BackColor       =   &H00F8D047&
   ClientHeight    =   2190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6255
   Icon            =   "uun.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   2190
   ScaleWidth      =   6255
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Height          =   975
      Left            =   240
      Picture         =   "uun.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F8D047&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   735
      Left            =   4800
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00F8D047&
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00F8D047&
      Caption         =   "全部关闭"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F8D047&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   735
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   4320
      X2              =   4680
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F8D047&
      Caption         =   "关闭副窗口"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "uun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
closed = True
For i = 1 To 100
On Error GoTo 99
Unload sform(i)
Set sform(i) = Nothing
Next i
99 fordern = 0
fordern = 0
second = False
End Sub

Private Sub Command2_Click()
closed = True
If Val(Text1.Text) <= Val(Text2.Text) Then
For i = Val(Text1.Text) To Val(Text2.Text)
On Error GoTo 99
Unload sform(i)
Next i
99 fordern = 0
second = False
Else
MsgBox "输入值不合法"
closed = False
End If
End Sub

Private Sub Command3_Click()
Unload uun
End Sub
