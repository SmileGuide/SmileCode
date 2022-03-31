VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{60CC5D62-2D08-11D0-BDBE-00AA00575603}#1.0#0"; "systray.ocx"
Begin VB.Form FrmMn 
   BackColor       =   &H00FEFBBC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "课程表"
   ClientHeight    =   3246
   ClientLeft      =   2274
   ClientTop       =   9294
   ClientWidth     =   5388
   Icon            =   "Mn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3246
   ScaleWidth      =   5388
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MCI.MMControl MMControl1 
      Height          =   546
      Left            =   3000
      TabIndex        =   11
      Top             =   600
      Width           =   1416
      _ExtentX        =   2498
      _ExtentY        =   963
      _Version        =   393216
   End
   Begin SysTrayCtl.cSysTray CTray 
      Left            =   3780
      Top             =   840
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   -1  'True
      TrayIcon        =   "Mn.frx":1084A
      TrayTip         =   "VB 5 - SysTray Control."
   End
   Begin VB.Timer TmrNow 
      Interval        =   500
      Left            =   4500
      Top             =   840
   End
   Begin MCI.MMControl MMCR 
      Height          =   305
      Left            =   3600
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   1085
      _ExtentX        =   1916
      _ExtentY        =   540
      _Version        =   393216
      BorderStyle     =   0
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PlayVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton CmdEdit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "编辑课程表"
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
      Left            =   2580
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "小工具"
      Top             =   300
      Width           =   2580
   End
   Begin VB.CommandButton CmdSet 
      BackColor       =   &H00C0FFFF&
      Caption         =   "更换课程表"
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
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "小工具"
      Top             =   300
      Width           =   2580
   End
   Begin VB.CommandButton CmdTool 
      BackColor       =   &H00C0FFFF&
      Caption         =   "…"
      Height          =   245
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "选项"
      Top             =   300
      Width           =   245
   End
   Begin VB.Frame FrmTab 
      BackColor       =   &H00FEFBBC&
      Height          =   1864
      Left            =   60
      TabIndex        =   1
      Top             =   1260
      Width           =   5224
      Begin VB.ListBox LstO 
         BackColor       =   &H00FEFBBC&
         BeginProperty Font 
            Name            =   "华文中宋"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1464
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   180
         Width           =   544
      End
      Begin VB.ListBox LstTm 
         BackColor       =   &H00FEFBBC&
         BeginProperty Font 
            Name            =   "华文中宋"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1464
         Left            =   900
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   2344
      End
      Begin VB.ListBox LstL 
         BackColor       =   &H00FEFBBC&
         BeginProperty Font 
            Name            =   "华文中宋"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1464
         Left            =   3420
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   1624
      End
   End
   Begin VB.Label LblNow 
      BackStyle       =   0  'Transparent
      Caption         =   "12:00:00"
      BeginProperty Font 
         Name            =   "Orbitron"
         Size            =   26.1
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   606
      Left            =   60
      TabIndex        =   9
      Top             =   480
      Width           =   5226
   End
   Begin VB.Label LblCap 
      BackStyle       =   0  'Transparent
      Caption         =   "我的课程表 星期X"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   305
      Left            =   0
      TabIndex        =   4
      Top             =   60
      Width           =   4805
   End
   Begin VB.Label LblS 
      BackStyle       =   0  'Transparent
      Caption         =   "下节课：--       距上课：--分"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   244
      Left            =   60
      TabIndex        =   0
      Top             =   1020
      Width           =   5224
   End
   Begin VB.Shape ShpCap 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   3
      Height          =   304
      Left            =   0
      Top             =   0
      Width           =   5404
   End
End
Attribute VB_Name = "FrmMn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
LblNow.Caption = Format(Now, "hh:mm:ss")
End Sub

Private Sub TmrNow_Timer()
If LblNow.Caption <> Format(Now, "hh:mm:ss") Then
    LblNow.Caption = Format(Now, "hh:mm:ss")
End If
End Sub
