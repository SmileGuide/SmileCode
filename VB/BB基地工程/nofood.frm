VERSION 5.00
Begin VB.Form nofood 
   BackColor       =   &H00C0C000&
   ClientHeight    =   1530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3930
   Icon            =   "nofood.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1530
   ScaleWidth      =   3930
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "去讨养料"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "小抱抱还没有给你养料呢"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "nofood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
askforfood.Show
End Sub
