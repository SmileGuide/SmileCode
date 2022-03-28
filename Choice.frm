VERSION 5.00
Begin VB.Form FrmCho 
   BackColor       =   &H00FEFBBC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择课程表"
   ClientHeight    =   3378
   ClientLeft      =   4818
   ClientTop       =   2814
   ClientWidth     =   3642
   Icon            =   "Choice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3378
   ScaleWidth      =   3642
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.FileListBox FleTab 
      BackColor       =   &H00FEFBBC&
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   1908
      Left            =   120
      Pattern         =   "*.sgtab"
      TabIndex        =   3
      Top             =   480
      Width           =   3364
   End
   Begin VB.CommandButton CmdDo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "管理"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   224
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2460
      Width           =   3364
   End
   Begin VB.CommandButton CmdOK 
      BackColor       =   &H00FEFBBC&
      Caption         =   "即刻开启>"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   545
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   3545
   End
   Begin VB.Label LblKind 
      BackStyle       =   0  'Transparent
      Caption         =   "选择课程表"
      BeginProperty Font 
         Name            =   "华文新魏"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   304
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1384
   End
End
Attribute VB_Name = "FrmCho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FleTab_Click()

End Sub
