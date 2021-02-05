VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmWel 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2320
   ClientLeft      =   1204
   ClientTop       =   -2708
   ClientWidth     =   1992
   ControlBox      =   0   'False
   Icon            =   "Welcome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Welcome.frx":1084A
   ScaleHeight     =   2320
   ScaleWidth      =   1992
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CmdAbme 
      BackColor       =   &H00C0FFFF&
      Caption         =   "关于作者"
      Height          =   184
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1984
   End
   Begin MSComctlLib.ProgressBar Pg 
      Height          =   124
      Left            =   0
      TabIndex        =   2
      Top             =   1560
      Width           =   1984
      _ExtentX        =   3500
      _ExtentY        =   219
      _Version        =   393216
      Appearance      =   1
      Max             =   5
   End
   Begin VB.CommandButton CmdEtr 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Start Your Orderly Life →"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   424
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1864
   End
   Begin VB.Label LblPg 
      BackStyle       =   0  'Transparent
      Caption         =   "正在加载窗体…"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   6.43
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   124
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   844
   End
   Begin VB.Label LblWord 
      BackStyle       =   0  'Transparent
      Caption         =   "凡事预则立，不预则废。               ——《礼记·中庸》"
      BeginProperty Font 
         Name            =   "华文行楷"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   1564
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1984
   End
End
Attribute VB_Name = "FrmWel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

