VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmSet 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "设置"
   ClientHeight    =   4760
   ClientLeft      =   6348
   ClientTop       =   -4168
   ClientWidth     =   7112
   Icon            =   "FrmSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4760
   ScaleWidth      =   7112
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComDlg.CommonDialog CDFL 
      Left            =   2280
      Top             =   -60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "选择铃声"
      Filter          =   "MEPG-3(*.MP3)|*.mp3|Wave(*.WAV)|*.wav|MIDI(*.MID）|*.mid"
      FontSize        =   12
      Min             =   12
   End
   Begin VB.Timer TmrFresh 
      Interval        =   1
      Left            =   720
      Top             =   0
   End
   Begin VB.Frame FrmRem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "提醒"
      BeginProperty Font 
         Name            =   "华文仿宋"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1024
      Left            =   60
      TabIndex        =   1
      Top             =   3540
      Width           =   6904
      Begin VB.CommandButton CmdView 
         BackColor       =   &H00C0FFC0&
         Caption         =   "提醒窗体个性化设置"
         Height          =   304
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   600
         Width           =   2344
      End
      Begin VB.CheckBox ChkIOR 
         BackColor       =   &H00C0FFFF&
         Caption         =   "开启铃声"
         BeginProperty Font 
            Name            =   "华文仿宋"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   365
         Left            =   180
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Value           =   1  'Checked
         Width           =   1025
      End
      Begin VB.TextBox TxtRF 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "等线"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   1260
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "(请输入铃声位置）"
         Top             =   360
         Width           =   4685
      End
      Begin VB.CommandButton CmdFL 
         BackColor       =   &H00C0FFFF&
         Caption         =   "浏览..."
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
         Left            =   6060
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "小工具"
         Top             =   300
         Width           =   725
      End
   End
   Begin VB.Frame FrmPers 
      BackColor       =   &H00C0FFFF&
      Caption         =   "UI"
      BeginProperty Font 
         Name            =   "华文仿宋"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3484
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6905
      Begin VB.CommandButton CmdN 
         BackColor       =   &H00C0FFFF&
         Caption         =   "恢复默认"
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
         Index           =   7
         Left            =   6060
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "小工具"
         Top             =   2100
         Width           =   725
      End
      Begin VB.CommandButton CmdN 
         BackColor       =   &H00C0FFFF&
         Caption         =   "恢复默认"
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
         Index           =   4
         Left            =   6060
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "小工具"
         Top             =   1740
         Width           =   725
      End
      Begin VB.CommandButton CmdN 
         BackColor       =   &H00C0FFFF&
         Caption         =   "恢复默认"
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
         Index           =   3
         Left            =   6060
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "小工具"
         Top             =   1380
         Width           =   725
      End
      Begin VB.CommandButton CmdN 
         BackColor       =   &H00C0FFFF&
         Caption         =   "恢复默认"
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
         Index           =   2
         Left            =   6060
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "小工具"
         Top             =   1020
         Width           =   725
      End
      Begin VB.CommandButton CmdN 
         BackColor       =   &H00C0FFFF&
         Caption         =   "恢复默认"
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
         Index           =   6
         Left            =   6060
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "小工具"
         Top             =   2820
         Width           =   725
      End
      Begin VB.ComboBox CboF 
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Index           =   1
         ItemData        =   "FrmSet.frx":1084A
         Left            =   1380
         List            =   "FrmSet.frx":1087E
         TabIndex        =   11
         Text            =   "幼圆"
         Top             =   2880
         Width           =   4264
      End
      Begin VB.CommandButton CmdN 
         BackColor       =   &H00C0FFFF&
         Caption         =   "恢复默认"
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
         Index           =   5
         Left            =   6060
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "小工具"
         Top             =   2460
         Width           =   725
      End
      Begin VB.CommandButton CmdN 
         BackColor       =   &H00C0FFFF&
         Caption         =   "恢复默认"
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
         Index           =   1
         Left            =   6060
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "小工具"
         Top             =   660
         Width           =   725
      End
      Begin VB.CommandButton CmdN 
         BackColor       =   &H00C0FFFF&
         Caption         =   "恢复默认"
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
         Index           =   0
         Left            =   6060
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "小工具"
         Top             =   300
         Width           =   725
      End
      Begin VB.ComboBox CboF 
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Index           =   0
         ItemData        =   "FrmSet.frx":10904
         Left            =   1380
         List            =   "FrmSet.frx":10938
         TabIndex        =   5
         Text            =   "幼圆"
         Top             =   2520
         Width           =   4264
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "提醒时间颜色"
         Height          =   184
         Left            =   180
         TabIndex        =   25
         ToolTipText     =   "提醒窗口上现在时间的颜色"
         Top             =   2100
         Width           =   1024
      End
      Begin VB.Shape ShpC 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFF00&
         Height          =   304
         Index           =   5
         Left            =   1380
         Top             =   2040
         Width           =   4624
      End
      Begin VB.Label LblRed 
         BackColor       =   &H00C0C0FF&
         Caption         =   "！ 注意：颜色有重复内容，可能会互相混淆"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   184
         Left            =   1680
         TabIndex        =   20
         ToolTipText     =   "课程表中的紫色总结文字"
         Top             =   3180
         Visible         =   0   'False
         Width           =   3364
      End
      Begin VB.Shape ShpC 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFF00&
         Height          =   304
         Index           =   4
         Left            =   1380
         Top             =   1680
         Width           =   4624
      End
      Begin VB.Label LblNC 
         BackStyle       =   0  'Transparent
         Caption         =   "现在时间颜色"
         Height          =   184
         Left            =   180
         TabIndex        =   19
         ToolTipText     =   "课程表上的实时时间颜色"
         Top             =   1740
         Width           =   1024
      End
      Begin VB.Label LblCpC 
         BackStyle       =   0  'Transparent
         Caption         =   "课程预报颜色色"
         Height          =   184
         Left            =   180
         TabIndex        =   15
         ToolTipText     =   "课程表中的紫色总结文字颜色"
         Top             =   1020
         Width           =   1024
      End
      Begin VB.Label LblTabC 
         BackStyle       =   0  'Transparent
         Caption         =   "课程表颜色"
         Height          =   184
         Left            =   180
         TabIndex        =   14
         ToolTipText     =   "编辑操作窗口的背景色"
         Top             =   1380
         Width           =   1024
      End
      Begin VB.Shape ShpC 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFF00&
         Height          =   304
         Index           =   2
         Left            =   1380
         Top             =   960
         Width           =   4624
      End
      Begin VB.Shape ShpC 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFF00&
         Height          =   304
         Index           =   3
         Left            =   1380
         Top             =   1320
         Width           =   4624
      End
      Begin VB.Label LblOK 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "√"
         Height          =   184
         Index           =   1
         Left            =   5640
         TabIndex        =   12
         Top             =   2880
         Width           =   364
      End
      Begin VB.Label LblTabF 
         BackStyle       =   0  'Transparent
         Caption         =   "课程表字体"
         Height          =   184
         Left            =   180
         TabIndex        =   10
         ToolTipText     =   "课程表中的紫色总结文字"
         Top             =   2880
         Width           =   1024
      End
      Begin VB.Label LblOK 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "√"
         Height          =   184
         Index           =   0
         Left            =   5640
         TabIndex        =   7
         Top             =   2520
         Width           =   364
      End
      Begin VB.Label LblCpF 
         BackStyle       =   0  'Transparent
         Caption         =   "课程预报字体"
         Height          =   184
         Left            =   180
         TabIndex        =   4
         ToolTipText     =   "课程表中的紫色总结文字字体"
         Top             =   2520
         Width           =   1024
      End
      Begin VB.Shape ShpC 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFF00&
         Height          =   304
         Index           =   1
         Left            =   1380
         Top             =   600
         Width           =   4624
      End
      Begin VB.Shape ShpC 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFF00&
         Height          =   304
         Index           =   0
         Left            =   1380
         Top             =   240
         Width           =   4624
      End
      Begin VB.Label LblEdC 
         BackStyle       =   0  'Transparent
         Caption         =   "修改时背景色"
         Height          =   184
         Left            =   180
         TabIndex        =   3
         ToolTipText     =   "编辑操作窗口的背景色"
         Top             =   660
         Width           =   1024
      End
      Begin VB.Label LblRC 
         BackStyle       =   0  'Transparent
         Caption         =   "使用时背景色"
         Height          =   184
         Left            =   180
         TabIndex        =   2
         ToolTipText     =   "非编辑操作窗口的背景色"
         Top             =   300
         Width           =   1024
      End
   End
End
Attribute VB_Name = "FrmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

