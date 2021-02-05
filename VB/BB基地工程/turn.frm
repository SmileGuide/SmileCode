VERSION 5.00
Begin VB.Form turn 
   BackColor       =   &H00FFC0FF&
   Caption         =   "跳转"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4530
   Icon            =   "turn.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   4530
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command7 
      BackColor       =   &H0080FF80&
      Caption         =   "去玩红绿灯>"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   4175
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0080FF80&
      Caption         =   "去画画>"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   4175
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080FF80&
      Caption         =   "去破门颓墙敲古寺锈钟>"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   4175
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FF80&
      Caption         =   "去管理养料>"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   4175
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Caption         =   "去浇水>"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   4175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "去组内交流>"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   4175
   End
End
Attribute VB_Name = "turn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell "F:\WPE Files\组内交流.exe", vbMaxmizedFocus
Unload Me
End Sub

Private Sub Command2_Click()
Shell "F:\SE Files\圣哥浏览器.exe", vbMaxmizedFocus
Unload Me
End Sub

Private Sub Command3_Click()
Call foodp
End Sub

Private Sub Command4_Click()
Call foodp
turn.Hide
main.Show
End Sub

Private Sub Command5_Click()
Shell "F:\Clock File\古寺锈钟.exe", vbMaxmizedFocus
Unload Me
End Sub

Private Sub Command6_Click()
Shell "F:\Board Files\drawing-board.exe", vbMaxmizedFocus
Unload Me
End Sub

Private Sub Command7_Click()
Shell "F:\corner File\corner.exe", vbMaxmizedFocus
Unload Me
End Sub
