VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmSet 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "����"
   ClientHeight    =   4764
   ClientLeft      =   6348
   ClientTop       =   -4170
   ClientWidth     =   7224
   Icon            =   "FrmSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4764
   ScaleWidth      =   7224
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComDlg.CommonDialog CDFL 
      Left            =   2280
      Top             =   -60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "ѡ������"
      Filter          =   "MEPG-3(*.MP3)|*.mp3|Wave(*.WAV)|*.wav|MIDI(*.MID��|*.mid"
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "���ķ���"
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
      Begin VB.TextBox TxtRPath 
         BackColor       =   &H00C0FFFF&
         Height          =   250
         Left            =   1200
         TabIndex        =   26
         Top             =   310
         Width           =   4746
      End
      Begin VB.CommandButton CmdView 
         BackColor       =   &H00C0FFC0&
         Caption         =   "���Ѵ�����Ի�����"
         Height          =   304
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   600
         Width           =   2344
      End
      Begin VB.CheckBox ChkIOR 
         BackColor       =   &H00C0FFFF&
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "���ķ���"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   365
         Left            =   180
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Value           =   1  'Checked
         Width           =   1025
      End
      Begin VB.CommandButton CmdFL 
         BackColor       =   &H00DFFFB0&
         Caption         =   "���..."
         BeginProperty Font 
            Name            =   "���ķ���"
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
         ToolTipText     =   "С����"
         Top             =   300
         Width           =   725
      End
   End
   Begin VB.Frame FrmPers 
      BackColor       =   &H00C0FFFF&
      Caption         =   "UI"
      BeginProperty Font 
         Name            =   "���ķ���"
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
      Width           =   7026
      Begin VB.CommandButton CmdN 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�ָ�Ĭ��"
         BeginProperty Font 
            Name            =   "���ķ���"
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
         TabIndex        =   23
         ToolTipText     =   "С����"
         Top             =   2100
         Width           =   906
      End
      Begin VB.CommandButton CmdN 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�ָ�Ĭ��"
         BeginProperty Font 
            Name            =   "���ķ���"
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
         ToolTipText     =   "С����"
         Top             =   1740
         Width           =   906
      End
      Begin VB.CommandButton CmdN 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�ָ�Ĭ��"
         BeginProperty Font 
            Name            =   "���ķ���"
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
         ToolTipText     =   "С����"
         Top             =   1380
         Width           =   906
      End
      Begin VB.CommandButton CmdN 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�ָ�Ĭ��"
         BeginProperty Font 
            Name            =   "���ķ���"
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
         ToolTipText     =   "С����"
         Top             =   1020
         Width           =   906
      End
      Begin VB.CommandButton CmdN 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�ָ�Ĭ��"
         BeginProperty Font 
            Name            =   "���ķ���"
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
         ToolTipText     =   "С����"
         Top             =   2820
         Width           =   906
      End
      Begin VB.ComboBox CboF 
         BeginProperty Font 
            Name            =   "��Բ"
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
         Text            =   "��Բ"
         Top             =   2880
         Width           =   4264
      End
      Begin VB.CommandButton CmdN 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�ָ�Ĭ��"
         BeginProperty Font 
            Name            =   "���ķ���"
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
         ToolTipText     =   "С����"
         Top             =   2460
         Width           =   906
      End
      Begin VB.CommandButton CmdN 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�ָ�Ĭ��"
         BeginProperty Font 
            Name            =   "���ķ���"
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
         ToolTipText     =   "С����"
         Top             =   660
         Width           =   906
      End
      Begin VB.CommandButton CmdN 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�ָ�Ĭ��"
         BeginProperty Font 
            Name            =   "���ķ���"
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
         ToolTipText     =   "С����"
         Top             =   300
         Width           =   906
      End
      Begin VB.ComboBox CboF 
         BeginProperty Font 
            Name            =   "��Բ"
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
         Text            =   "��Բ"
         Top             =   2520
         Width           =   4264
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ����ɫ"
         Height          =   186
         Left            =   180
         TabIndex        =   24
         ToolTipText     =   "���Ѵ���������ʱ�����ɫ"
         Top             =   2100
         Width           =   1146
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
         Caption         =   "!!!ע�⣺��ɫ�ظ��ụ�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   186
         Left            =   2640
         TabIndex        =   20
         ToolTipText     =   "�γ̱��е���ɫ�ܽ�����"
         Top             =   3180
         Visible         =   0   'False
         Width           =   2346
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
         Caption         =   "����ʱ����ɫ"
         Height          =   186
         Left            =   180
         TabIndex        =   19
         ToolTipText     =   "�γ̱��ϵ�ʵʱʱ����ɫ"
         Top             =   1740
         Width           =   1146
      End
      Begin VB.Label LblCpC 
         BackStyle       =   0  'Transparent
         Caption         =   "�γ�Ԥ����ɫɫ"
         Height          =   186
         Left            =   180
         TabIndex        =   15
         ToolTipText     =   "�γ̱��е���ɫ�ܽ�������ɫ"
         Top             =   1020
         Width           =   1206
      End
      Begin VB.Label LblTabC 
         BackStyle       =   0  'Transparent
         Caption         =   "�γ̱���ɫ"
         Height          =   184
         Left            =   180
         TabIndex        =   14
         ToolTipText     =   "�༭�������ڵı���ɫ"
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
         Caption         =   "��"
         Height          =   184
         Index           =   1
         Left            =   5640
         TabIndex        =   12
         Top             =   2880
         Width           =   364
      End
      Begin VB.Label LblTabF 
         BackStyle       =   0  'Transparent
         Caption         =   "�γ̱�����"
         Height          =   186
         Left            =   180
         TabIndex        =   10
         ToolTipText     =   "�γ̱��е���ɫ�ܽ�����"
         Top             =   2880
         Width           =   1206
      End
      Begin VB.Label LblOK 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "��"
         Height          =   184
         Index           =   0
         Left            =   5640
         TabIndex        =   7
         Top             =   2520
         Width           =   364
      End
      Begin VB.Label LblCpF 
         BackStyle       =   0  'Transparent
         Caption         =   "�γ�Ԥ������"
         Height          =   186
         Left            =   180
         TabIndex        =   4
         ToolTipText     =   "�γ̱��е���ɫ�ܽ���������"
         Top             =   2520
         Width           =   1206
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
         Caption         =   "�޸�ʱ����ɫ"
         Height          =   186
         Left            =   180
         TabIndex        =   3
         ToolTipText     =   "�༭�������ڵı���ɫ"
         Top             =   660
         Width           =   1206
      End
      Begin VB.Label LblRC 
         BackStyle       =   0  'Transparent
         Caption         =   "ʹ��ʱ����ɫ"
         Height          =   186
         Left            =   180
         TabIndex        =   2
         ToolTipText     =   "�Ǳ༭�������ڵı���ɫ"
         Top             =   300
         Width           =   1146
      End
   End
End
Attribute VB_Name = "FrmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

