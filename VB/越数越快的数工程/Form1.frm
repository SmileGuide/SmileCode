VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00F8D047&
   Caption         =   "faster"
   ClientHeight    =   2430
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   19635
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   72
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2430
   ScaleWidth      =   19635
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   1440
   End
   Begin VB.CommandButton Command2 
      Caption         =   "暂停"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   17760
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F8D047&
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Interval = 1
Form1.Refresh
Command1.Visible = False
Command2.Visible = True
Text1.Text = Format(Text1.Text, "0.00")
End Sub

Private Sub Command2_Click()
Timer1.Interval = 0
Form1.Refresh
Command2.Visible = False
Command1.Visible = True
End Sub


Private Sub Text1_Change()
If Val(Text1.Text) > 30000 Then
MsgBox "数太大了"
Text1.Text = ""
End If
End Sub

Private Sub Timer1_Timer()
tt = tt + 1
Text1.Text = Val(Text1.Text) - tt / 10000
Text1.Text = Format(Text1.Text, "0.00")
If Val(Text1.Text) = 0 Then
MsgBox "时间到！"
Beep
Text1.Text = ""
Timer1.Interval = 0
Command1.Visible = True
End If
End Sub
