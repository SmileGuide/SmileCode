VERSION 5.00
Begin VB.Form sorryhope 
   BackColor       =   &H00EBBEEB&
   Caption         =   "很抱歉，您的愿望已被小抱抱驳回"
   ClientHeight    =   2280
   ClientLeft      =   5835
   ClientTop       =   4305
   ClientWidth     =   4770
   Icon            =   "sorry.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   4770
   Begin VB.CommandButton Command1 
      BackColor       =   &H0097FBCB&
      Caption         =   "重新提交愿望"
      Height          =   615
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
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
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   4215
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "sorryhope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
hope.Show
End Sub
