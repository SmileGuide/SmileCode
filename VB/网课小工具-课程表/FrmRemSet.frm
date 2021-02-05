VERSION 5.00
Begin VB.Form FrmRemSet 
   BackColor       =   &H00C0FFFF&
   Caption         =   "提醒窗体个性化设置"
   ClientHeight    =   4784
   ClientLeft      =   16
   ClientTop       =   328
   ClientWidth     =   4872
   Icon            =   "FrmRemSet.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4784
   ScaleWidth      =   4872
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PicCtrlBox 
      Height          =   304
      Left            =   60
      Picture         =   "FrmRemSet.frx":1084A
      ScaleHeight     =   288
      ScaleWidth      =   4744
      TabIndex        =   1
      Top             =   60
      Width           =   4759
   End
   Begin VB.PictureBox PicFrmV 
      BackColor       =   &H00FFFFC0&
      Height          =   1504
      Left            =   60
      ScaleHeight     =   1488
      ScaleWidth      =   4728
      TabIndex        =   0
      Top             =   360
      Width           =   4744
      Begin VB.CommandButton CmdEdit 
         BackColor       =   &H00C0FFC0&
         Caption         =   "朕知道了"
         BeginProperty Font 
            Name            =   "华文中宋"
            Size            =   13.93
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
         TabIndex        =   4
         Top             =   0
         Width           =   2644
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
         TabIndex        =   3
         Top             =   780
         Width           =   5224
      End
   End
   Begin VB.Line LineUd 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   0
      X2              =   4860
      Y1              =   1920
      Y2              =   1920
   End
End
Attribute VB_Name = "FrmRemSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

