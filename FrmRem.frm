VERSION 5.00
Begin VB.Form FrmRem 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "提醒窗体预览"
   ClientHeight    =   1530
   ClientLeft      =   11970
   ClientTop       =   3318
   ClientWidth     =   4716
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmRem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4716
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdEdit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "了解"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   13.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   424
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "小工具"
      Top             =   1020
      Width           =   4560
   End
   Begin VB.Label LblS 
      BackStyle       =   0  'Transparent
      Caption         =   "XXX距上课还有XXX分，请做好准备"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   244
      Left            =   0
      TabIndex        =   1
      Top             =   780
      Width           =   5224
   End
   Begin VB.Label LblNow 
      BackStyle       =   0  'Transparent
      Caption         =   "12:00"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   42
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   844
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2644
   End
End
Attribute VB_Name = "FrmRem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

