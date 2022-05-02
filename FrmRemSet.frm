VERSION 5.00
Begin VB.Form FrmRemSet 
   BackColor       =   &H00C0FFFF&
   Caption         =   "提醒窗体个性化设置"
   ClientHeight    =   2436
   ClientLeft      =   18
   ClientTop       =   330
   ClientWidth     =   5586
   Icon            =   "FrmRemSet.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2436
   ScaleWidth      =   5586
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer T1 
      Interval        =   500
      Left            =   2220
      Top             =   2280
   End
   Begin VB.PictureBox PicFrmV 
      BackColor       =   &H00FFFFC0&
      Height          =   1504
      Left            =   368
      ScaleHeight     =   1482
      ScaleWidth      =   4740
      TabIndex        =   0
      Top             =   480
      Width           =   4764
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
         TabIndex        =   1
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
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   780
         Width           =   5224
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   5132
      X2              =   5132
      Y1              =   480
      Y2              =   2000
   End
   Begin VB.Image ImgControlBox 
      Height          =   360
      Left            =   360
      Picture         =   "FrmRemSet.frx":1084A
      Top             =   120
      Width           =   4764
   End
   Begin VB.Line LineUd 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   368
      X2              =   5132
      Y1              =   2000
      Y2              =   2000
   End
End
Attribute VB_Name = "FrmRemSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

