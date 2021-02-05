VERSION 5.00
Begin VB.Form Startfrom 
   BackColor       =   &H00C0FFFF&
   Caption         =   "向导"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2850
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2910
   ScaleWidth      =   2850
   Begin VB.CommandButton Cm 
      BackColor       =   &H00C0FFC0&
      Caption         =   "浏览器"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   540
      Width           =   2775
   End
   Begin VB.CommandButton Cm 
      BackColor       =   &H00C0FFC0&
      Caption         =   "计算器"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   60
      Width           =   2760
   End
End
Attribute VB_Name = "Startfrom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
