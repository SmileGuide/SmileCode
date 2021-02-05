VERSION 5.00
Begin VB.Form makebadprom 
   Caption         =   "承诺书"
   ClientHeight    =   1200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4665
   Icon            =   "makebadprom.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "makebadprom.frx":324A
   ScaleHeight     =   1200
   ScaleWidth      =   4665
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "打印"
      Height          =   975
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Q1 
      Caption         =   "小抱抱承诺为小蛋蛋无偿叠被一次"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "makebadprom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
