VERSION 5.00
Begin VB.Form FrmEdit 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "     �ҵĿγ̱� - ΢Ц�γ̱� [�༭]"
   ClientHeight    =   3600
   ClientLeft      =   7880
   ClientTop       =   9440
   ClientWidth     =   5528
   Icon            =   "FrmEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5528
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdDay 
      BackColor       =   &H00A5E9FC&
      Caption         =   "����X"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   245
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "С����"
      Top             =   0
      Width           =   540
   End
   Begin VB.TextBox TxtName 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "��������"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   184
      Left            =   1140
      TabIndex        =   19
      Text            =   "�ҵĿγ̱�"
      ToolTipText     =   "�������޸�"
      Top             =   360
      Width           =   4084
   End
   Begin VB.CommandButton CmdFile 
      BackColor       =   &H00A5E9FC&
      Caption         =   "�ļ�"
      BeginProperty Font 
         Name            =   "����"
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
      TabIndex        =   1
      ToolTipText     =   "С����"
      Top             =   0
      Width           =   480
   End
   Begin VB.Frame FrmTab 
      BackColor       =   &H00C0FFFF&
      Height          =   2704
      Left            =   0
      TabIndex        =   0
      Top             =   540
      Width           =   5224
      Begin VB.PictureBox PicO 
         BackColor       =   &H0080FFFF&
         Height          =   2404
         Left            =   120
         ScaleHeight     =   2388
         ScaleWidth      =   648
         TabIndex        =   16
         Top             =   180
         Width           =   664
         Begin VB.ListBox LstO 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "��������"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1808
            ItemData        =   "FrmEdit.frx":1084A
            Left            =   60
            List            =   "FrmEdit.frx":1084C
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   300
            Width           =   544
         End
      End
      Begin VB.PictureBox PicL 
         BackColor       =   &H0080FFFF&
         Height          =   2404
         Left            =   3240
         ScaleHeight     =   2388
         ScaleWidth      =   1788
         TabIndex        =   4
         Top             =   180
         Width           =   1804
         Begin VB.CommandButton CmdLDown 
            BackColor       =   &H00C0FFC0&
            Caption         =   "����"
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
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "С����"
            Top             =   2100
            Width           =   424
         End
         Begin VB.CommandButton CmdLUp 
            BackColor       =   &H00C0FFC0&
            Caption         =   "����"
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
            Left            =   900
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "С����"
            Top             =   2100
            Width           =   424
         End
         Begin VB.CommandButton CmdLAdd 
            BackColor       =   &H00C0FFC0&
            Caption         =   "���"
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
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "С����"
            Top             =   2100
            Width           =   424
         End
         Begin VB.CommandButton CmdLDEL 
            BackColor       =   &H00C0FFC0&
            Caption         =   "ɾ��"
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
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "С����"
            Top             =   2100
            Width           =   424
         End
         Begin VB.ComboBox CboLCha 
            BackColor       =   &H00C0FFFF&
            Height          =   200
            Left            =   60
            TabIndex        =   10
            Top             =   60
            Width           =   1684
         End
         Begin VB.ListBox LstL 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "��������"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1808
            ItemData        =   "FrmEdit.frx":1084E
            Left            =   60
            List            =   "FrmEdit.frx":10850
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   300
            Width           =   1684
         End
      End
      Begin VB.PictureBox PicT 
         BackColor       =   &H0080FFFF&
         Height          =   2404
         Left            =   840
         ScaleHeight     =   2388
         ScaleWidth      =   2328
         TabIndex        =   2
         Top             =   180
         Width           =   2344
         Begin VB.ComboBox CboTMCha 
            BackColor       =   &H00C0FFFF&
            Height          =   200
            Left            =   1260
            TabIndex        =   20
            Top             =   60
            Width           =   1024
         End
         Begin VB.CommandButton CmdTDown 
            BackColor       =   &H00C0FFC0&
            Caption         =   "����"
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
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "С����"
            Top             =   2100
            Width           =   575
         End
         Begin VB.CommandButton CmdTUp 
            BackColor       =   &H00C0FFC0&
            Caption         =   "����"
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
            Left            =   1140
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "С����"
            Top             =   2100
            Width           =   575
         End
         Begin VB.CommandButton CmdTAdd 
            BackColor       =   &H00C0FFC0&
            Caption         =   "���"
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
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "С����"
            Top             =   2100
            Width           =   575
         End
         Begin VB.CommandButton CmdTDel 
            BackColor       =   &H00C0FFC0&
            Caption         =   "ɾ��"
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
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "С����"
            Top             =   2100
            Width           =   575
         End
         Begin VB.ComboBox CboTHCha 
            BackColor       =   &H00C0FFFF&
            Height          =   200
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   1024
         End
         Begin VB.ListBox LstTm 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "��������"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1808
            ItemData        =   "FrmEdit.frx":10852
            Left            =   60
            List            =   "FrmEdit.frx":10854
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   300
            Width           =   2224
         End
         Begin VB.Label LblJ 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "������κ"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   184
            Left            =   1140
            TabIndex        =   21
            Top             =   60
            Width           =   184
         End
      End
   End
   Begin VB.Label LblCpC 
      BackStyle       =   0  'Transparent
      Caption         =   "�γ̱����ƣ�"
      Height          =   184
      Left            =   60
      TabIndex        =   18
      ToolTipText     =   "�γ̱��е���ɫ�ܽ�������ɫ"
      Top             =   360
      Width           =   1084
   End
   Begin VB.Shape ShpMd 
      BackColor       =   &H00A5E9FC&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   244
      Left            =   0
      Top             =   0
      Width           =   5524
   End
   Begin VB.Menu MnFile 
      Caption         =   "�ļ�"
      Visible         =   0   'False
      Begin VB.Menu MnRename 
         Caption         =   "������"
      End
      Begin VB.Menu MnSave 
         Caption         =   "����Ϊ�ҵĿγ̱�"
      End
      Begin VB.Menu MnExIn 
         Caption         =   "�ӵ��ӱ����"
      End
      Begin VB.Menu MnToEx 
         Caption         =   "���Ϊ���ӱ��"
      End
   End
   Begin VB.Menu MnDay 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu MnDI 
         Caption         =   "����1"
         Index           =   1
      End
   End
End
Attribute VB_Name = "FrmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

