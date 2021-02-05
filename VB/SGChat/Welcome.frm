VERSION 5.00
Begin VB.Form Welcome 
   Caption         =   "欢迎！"
   ClientHeight    =   3880
   ClientLeft      =   44
   ClientTop       =   356
   ClientWidth     =   2672
   LinkTopic       =   "Form1"
   ScaleHeight     =   3880
   ScaleWidth      =   2672
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdAbout 
      Caption         =   "关于..."
      Height          =   484
      Left            =   60
      TabIndex        =   1
      Top             =   3300
      Width           =   1204
   End
   Begin VB.CommandButton CmdEnter 
      Caption         =   "进入→"
      Height          =   484
      Left            =   1380
      TabIndex        =   0
      Top             =   3300
      Width           =   1204
   End
   Begin VB.Image Image1 
      Height          =   3184
      Left            =   60
      Top             =   60
      Width           =   2524
   End
End
Attribute VB_Name = "Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

End Sub
