VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmSet 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "设置"
   ClientHeight    =   7440
   ClientLeft      =   6348
   ClientTop       =   -4170
   ClientWidth     =   7224
   Icon            =   "FrmSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   7224
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame FrmApi 
      BackColor       =   &H00C0FFFF&
      Caption         =   "API图片"
      BeginProperty Font 
         Name            =   "华文仿宋"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3726
      Left            =   60
      TabIndex        =   16
      Top             =   3600
      Width           =   7026
      Begin VB.ComboBox Combo2 
         Height          =   228
         Left            =   1380
         TabIndex        =   17
         Top             =   300
         Width           =   5226
      End
      Begin VB.Image ImgBk 
         BorderStyle     =   1  'Fixed Single
         Height          =   3066
         Left            =   240
         Picture         =   "FrmSet.frx":1084A
         Stretch         =   -1  'True
         Top             =   600
         Width           =   6306
      End
      Begin VB.Label LblAT 
         BackStyle       =   0  'Transparent
         Caption         =   "API接口地址"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   246
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   1206
      End
   End
   Begin MSComDlg.CommonDialog CDFL 
      Left            =   2280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "选择铃声"
      Filter          =   "MEPG-3(*.MP3)|*.mp3|Wave(*.WAV)|*.wav|MIDI(*.MID）|*.mid"
      FontSize        =   12
      Min             =   12
   End
   Begin VB.Frame FrmPers 
      BackColor       =   &H00C0FFFF&
      Caption         =   "主题"
      BeginProperty Font 
         Name            =   "华文仿宋"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3426
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7026
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "CMD自定义主题设置"
         Height          =   2106
         Left            =   3540
         TabIndex        =   6
         Top             =   1140
         Width           =   3066
         Begin VB.Label LblCC 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "指令颜色"
            BeginProperty Font 
               Name            =   "华文宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   282
            Left            =   660
            TabIndex        =   15
            Top             =   1620
            Width           =   1686
         End
         Begin VB.Label LblCMark 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "标点颜色"
            BeginProperty Font 
               Name            =   "华文宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   282
            Left            =   660
            TabIndex        =   14
            Top             =   1260
            Width           =   1686
         End
         Begin VB.Label LblCNum 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "数字颜色"
            BeginProperty Font 
               Name            =   "华文宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   282
            Left            =   660
            TabIndex        =   13
            Top             =   900
            Width           =   1686
         End
         Begin VB.Label LblCM 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "文本颜色"
            BeginProperty Font 
               Name            =   "华文宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   282
            Left            =   660
            TabIndex        =   12
            Top             =   540
            Width           =   1686
         End
         Begin VB.Label LblCG 
            Alignment       =   2  'Center
            BackColor       =   &H000C0C0C&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "（背景颜色）"
            BeginProperty Font 
               Name            =   "华文中宋"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1782
            Left            =   60
            TabIndex        =   11
            Top             =   240
            Width           =   2946
         End
      End
      Begin VB.Frame FrmGui 
         BackColor       =   &H00C0FFFF&
         Caption         =   "GUI自定义主题设置"
         Height          =   2106
         Left            =   120
         TabIndex        =   5
         Top             =   1140
         Width           =   3066
         Begin VB.Label LblGMenu 
            Alignment       =   2  'Center
            BackColor       =   &H000DA3FD&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "补充色调"
            BeginProperty Font 
               Name            =   "华文中宋"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   282
            Left            =   60
            TabIndex        =   10
            Top             =   1620
            Width           =   2886
         End
         Begin VB.Label LblGS 
            Alignment       =   2  'Center
            BackColor       =   &H00FDEEBF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "陪衬色调"
            BeginProperty Font 
               Name            =   "华文中宋"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   282
            Left            =   60
            TabIndex        =   9
            Top             =   1260
            Width           =   2886
         End
         Begin VB.Label LblGD 
            Alignment       =   2  'Center
            BackColor       =   &H0080FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "加强主色调"
            BeginProperty Font 
               Name            =   "华文中宋"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   282
            Left            =   60
            TabIndex        =   8
            Top             =   900
            Width           =   2886
         End
         Begin VB.Label LblGMain 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "主色调"
            BeginProperty Font 
               Name            =   "华文中宋"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   282
            Left            =   60
            TabIndex        =   7
            Top             =   540
            Width           =   2886
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   228
         Left            =   2640
         TabIndex        =   4
         Top             =   720
         Width           =   3966
      End
      Begin VB.ComboBox CboThm 
         Height          =   228
         Left            =   2640
         TabIndex        =   2
         Top             =   300
         Width           =   3966
      End
      Begin VB.Label LblCmd 
         BackStyle       =   0  'Transparent
         Caption         =   "命令行窗口（CMD）主题"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   246
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2406
      End
      Begin VB.Label LblTheme 
         BackStyle       =   0  'Transparent
         Caption         =   "图形用户界面（GUI)主题"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   246
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   2406
      End
   End
   Begin VB.Timer TmrFresh 
      Interval        =   1
      Left            =   720
      Top             =   0
   End
End
Attribute VB_Name = "FrmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LblGM_Click()

End Sub



Private Sub FrmApi_DragDrop(Source As Control, X As Single, Y As Single)

End Sub



Private Sub ImgBk_Click()

End Sub

Private Sub LblCC_Click()

End Sub

Private Sub LblCM_Click()

End Sub

Private Sub LblGS_Click()

End Sub
