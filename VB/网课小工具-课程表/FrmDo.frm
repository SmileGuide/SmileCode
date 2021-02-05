VERSION 5.00
Begin VB.Form FrmDo 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "管理课程表"
   ClientHeight    =   2500
   ClientLeft      =   1068
   ClientTop       =   -4608
   ClientWidth     =   3192
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
   ScaleHeight     =   2500
   ScaleWidth      =   3192
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.FileListBox FleTab 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   13.93
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   2096
      Left            =   60
      Pattern         =   "*.sgtab"
      TabIndex        =   2
      Top             =   60
      Width           =   3064
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
      Caption         =   "编辑/添加"
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

