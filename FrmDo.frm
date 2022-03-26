VERSION 5.00
Begin VB.Form FrmDo 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "管理课程表"
   ClientHeight    =   2502
   ClientLeft      =   1068
   ClientTop       =   -4608
   ClientWidth     =   4278
   BeginProperty Font 
      Name            =   "华文中宋"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmDo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2502
   ScaleWidth      =   4278
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.FileListBox FleTab 
      Height          =   648
      Left            =   3480
      Pattern         =   "*.smtab"
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   726
   End
   Begin VB.ListBox LstTab 
      BackColor       =   &H00C0FFFF&
      Height          =   2064
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   3066
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "删除"
      BeginProperty Font 
         Name            =   "华文仿宋"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   245
      Left            =   1620
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "小工具"
      Top             =   2160
      Width           =   1500
   End
   Begin VB.CommandButton CmdFL 
      BackColor       =   &H00C0FFFF&
      Caption         =   "新建"
      BeginProperty Font 
         Name            =   "华文仿宋"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   245
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "小工具"
      Top             =   2160
      Width           =   1500
   End
End
Attribute VB_Name = "FrmDo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdFL_Click()
StName = InputBox("请输入新建课程表名称：", "新建", "我的新课程表" & Format(Now, "yymmddhhmmss"))
Dim i
For i = 1 To 5
On Error Resume Next
Open App.Path & "\SmTab\" & StName & ".smtab" & i For Output As #1
Write #1, "", ""
Next
Close #1
FleTab.Refresh
Me.Hide
FrmNewEdit.Show
End Sub

Private Sub Form_Load()
On Error Resume Next
MkDir App.Path & "\Table"
FleTab.Path = App.Path & "\Table"
Dim i
For i = 0 To FleTab.ListCount - 1
LstTab.List(i) = FleTab.List(i)
Next i

End Sub
