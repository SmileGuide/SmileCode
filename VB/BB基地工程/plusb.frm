VERSION 5.00
Begin VB.Form plusb 
   BackColor       =   &H00C0FFFF&
   ClientHeight    =   1800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3270
   Icon            =   "plusb.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   3270
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   0
      Top             =   1320
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "不要退出哦~"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "规则：每1分钟操作数大于10，抱抱种子长大1颗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "小蛋蛋加分版"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "plusb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
UserForm1.CheckBox2.Value = False
End Sub


Private Sub Timer1_Timer()
Open "F:\WPE Files\times.bdf" For Input As #1
Input #1, a
Close #1
If a <= 10 Then
MsgBox "一分钟过去了，可你几乎没有操作~"
GoTo 1
Else
c = 1
Open "F:\BBSeed Files\nr.bdf" For Output As #2
Print #2, c
Close #1
End If
1 End Sub
