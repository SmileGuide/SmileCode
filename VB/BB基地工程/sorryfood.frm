VERSION 5.00
Begin VB.Form sorryfood 
   BackColor       =   &H00EBBEEB&
   Caption         =   "很抱歉，您的养料请求已被小抱抱驳回"
   ClientHeight    =   2475
   ClientLeft      =   5835
   ClientTop       =   4575
   ClientWidth     =   5625
   Icon            =   "sorryfood.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   5625
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0097FBCB&
      Caption         =   "重新提交请求"
      Height          =   615
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00EBBEEB&
      Caption         =   "驳回原因"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "sorryfood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
askforfood.Show
End Sub


