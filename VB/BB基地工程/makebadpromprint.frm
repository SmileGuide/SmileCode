VERSION 5.00
Begin VB.Form makebadpromprint 
   BackColor       =   &H00FFFFFF&
   Caption         =   "承诺书"
   ClientHeight    =   2205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4650
   LinkTopic       =   "Form2"
   Picture         =   "makebadpromprint.frx":0000
   ScaleHeight     =   2205
   ScaleWidth      =   4650
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   2160
      Picture         =   "makebadpromprint.frx":324A
      ScaleHeight     =   915
      ScaleWidth      =   2235
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Q1 
      BackColor       =   &H00FFFFFF&
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
Attribute VB_Name = "makebadpromprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
