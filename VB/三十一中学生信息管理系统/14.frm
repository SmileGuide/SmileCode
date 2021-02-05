VERSION 5.00
Begin VB.Form c14 
   BackColor       =   &H00C0FFFF&
   Caption         =   "31中学生信息管理系统"
   ClientHeight    =   7860
   ClientLeft      =   6420
   ClientTop       =   3884
   ClientWidth     =   13664
   Icon            =   "14.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   13664
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   7695
      Left            =   75
      ScaleHeight     =   7696
      ScaleWidth      =   13200
      TabIndex        =   0
      Top             =   0
      Width           =   13200
      Begin VB.CommandButton toup 
         BackColor       =   &H00C0FFFF&
         Height          =   525
         Left            =   12600
         Picture         =   "14.frx":1B692
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "窗口置顶/取消置顶"
         Top             =   63
         Width           =   555
      End
      Begin VB.TextBox Text 
         BackColor       =   &H0097FBCB&
         DataField       =   "name"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   555
         Index           =   1
         Left            =   2160
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   2  'Automatic
         TabIndex        =   28
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Data Data1 
         BackColor       =   &H00C0FFFF&
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  '缺省游标
         DefaultType     =   2  '使用 ODBC
         Exclusive       =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   1200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   120
         Width           =   10095
      End
      Begin VB.TextBox Text 
         BackColor       =   &H0097FBCB&
         DataField       =   "sex"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   2
         Left            =   2160
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   2  'Automatic
         TabIndex        =   27
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox Text 
         BackColor       =   &H0097FBCB&
         DataField       =   "class"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   3
         Left            =   2160
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   2  'Automatic
         TabIndex        =   26
         Top             =   2280
         Width           =   3135
      End
      Begin VB.TextBox Text 
         BackColor       =   &H0097FBCB&
         DataField       =   "pro"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   4
         Left            =   2160
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   2  'Automatic
         TabIndex        =   25
         Top             =   2880
         Width           =   3135
      End
      Begin VB.TextBox Text 
         BackColor       =   &H0097FBCB&
         DataField       =   "city"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   5
         Left            =   2160
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   2  'Automatic
         TabIndex        =   24
         Top             =   3480
         Width           =   3135
      End
      Begin VB.TextBox Text 
         BackColor       =   &H0097FBCB&
         DataField       =   "street"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   6
         Left            =   2160
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   2  'Automatic
         TabIndex        =   23
         Top             =   4080
         Width           =   3135
      End
      Begin VB.TextBox Text 
         BackColor       =   &H0097FBCB&
         DataField       =   "village"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   7
         Left            =   2160
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   2  'Automatic
         TabIndex        =   22
         Top             =   4680
         Width           =   3135
      End
      Begin VB.TextBox Text 
         BackColor       =   &H0097FBCB&
         DataField       =   "area"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Index           =   8
         Left            =   8160
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   2  'Automatic
         ScrollBars      =   1  'Horizontal
         TabIndex        =   21
         Top             =   945
         Width           =   4965
      End
      Begin VB.TextBox Text 
         BackColor       =   &H0097FBCB&
         DataField       =   "number"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   555
         Index           =   9
         Left            =   8160
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   2  'Automatic
         TabIndex        =   20
         Top             =   1680
         Width           =   4965
      End
      Begin VB.TextBox Text 
         BackColor       =   &H0097FBCB&
         DataField       =   "fname"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   555
         Index           =   10
         Left            =   8160
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   2  'Automatic
         TabIndex        =   19
         Top             =   2280
         Width           =   4965
      End
      Begin VB.TextBox Text 
         BackColor       =   &H0097FBCB&
         DataField       =   "fphone"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   555
         Index           =   11
         Left            =   8160
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   2  'Automatic
         TabIndex        =   18
         Top             =   2880
         Width           =   4965
      End
      Begin VB.TextBox Text 
         BackColor       =   &H0097FBCB&
         DataField       =   "mname"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   555
         Index           =   12
         Left            =   8160
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   2  'Automatic
         TabIndex        =   17
         Top             =   3480
         Width           =   4965
      End
      Begin VB.TextBox Text 
         BackColor       =   &H0097FBCB&
         DataField       =   "mohone"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   555
         Index           =   13
         Left            =   8160
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   2  'Automatic
         TabIndex        =   16
         Top             =   4080
         Width           =   4965
      End
      Begin VB.TextBox Text 
         BackColor       =   &H0097FBCB&
         DataField       =   "other"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   555
         Index           =   14
         Left            =   8160
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   2  'Automatic
         TabIndex        =   15
         Top             =   4680
         Width           =   4965
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "操作"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.36
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2355
         Left            =   240
         TabIndex        =   1
         Top             =   5280
         Width           =   12765
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0FFFF&
            Caption         =   "历史记录"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   720
            Left            =   4575
            TabIndex        =   50
            Top             =   1470
            Width           =   2625
            Begin VB.ComboBox 查找历史 
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   75
               Style           =   2  'Dropdown List
               TabIndex        =   51
               ToolTipText     =   "查找历史"
               Top             =   270
               Width           =   2100
            End
            Begin VB.Image clearreco 
               Appearance      =   0  'Flat
               Height          =   405
               Left            =   2160
               Picture         =   "14.frx":1C254
               Stretch         =   -1  'True
               Top             =   255
               Width           =   450
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "信息复制"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1875
            Left            =   9720
            TabIndex        =   10
            Top             =   360
            Width           =   2760
            Begin VB.TextBox helpcon 
               BackColor       =   &H00FFFFC0&
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   1605
               Left            =   0
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               OLEDragMode     =   1  'Automatic
               ScrollBars      =   2  'Vertical
               TabIndex        =   49
               Text            =   "14.frx":1C696
               Top             =   255
               Visible         =   0   'False
               Width           =   2655
            End
            Begin VB.CommandButton help 
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.36
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   2430
               Picture         =   "14.frx":1C6B5
               Style           =   1  'Graphical
               TabIndex        =   44
               ToolTipText     =   "信息复制帮助"
               Top             =   1560
               Width           =   270
            End
            Begin VB.CommandButton readclipboard 
               BackColor       =   &H00C0FFFF&
               Caption         =   "识别剪贴板"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Left            =   105
               Style           =   1  'Graphical
               TabIndex        =   48
               Top             =   300
               Width           =   345
            End
            Begin VB.CheckBox cap 
               BackColor       =   &H00C0FFFF&
               Caption         =   "标题"
               Height          =   255
               Left            =   1800
               TabIndex        =   47
               Top             =   1425
               Value           =   1  'Checked
               Width           =   855
            End
            Begin VB.CheckBox muti 
               BackColor       =   &H00C0FFFF&
               Caption         =   "多行"
               Height          =   255
               Left            =   1815
               TabIndex        =   46
               Top             =   1170
               Value           =   1  'Checked
               Width           =   855
            End
            Begin VB.ListBox List1 
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H00800080&
               Height          =   1248
               ItemData        =   "14.frx":1C7B7
               Left            =   435
               List            =   "14.frx":1C7E7
               OLEDropMode     =   1  'Manual
               Style           =   1  'Checkbox
               TabIndex        =   14
               ToolTipText     =   "复制内容"
               Top             =   285
               Width           =   1335
            End
            Begin VB.CheckBox Check1 
               BackColor       =   &H00C0FFFF&
               Caption         =   "全选"
               Height          =   180
               Left            =   1815
               TabIndex        =   13
               Top             =   960
               Width           =   855
            End
            Begin VB.CommandButton Command1 
               BackColor       =   &H00C0FFFF&
               Caption         =   "复制"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   15
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1770
               OLEDropMode     =   1  'Manual
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   225
               Width           =   870
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00C0FFFF&
               Caption         =   "反选"
               Height          =   180
               Left            =   1800
               TabIndex        =   11
               Top             =   735
               Width           =   855
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "信息查找"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1875
            Left            =   4440
            TabIndex        =   6
            Top             =   360
            Width           =   5175
            Begin VB.CommandButton help 
               BackColor       =   &H00C0FFC0&
               Caption         =   "？"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.36
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   4620
               Picture         =   "14.frx":1C855
               Style           =   1  'Graphical
               TabIndex        =   45
               ToolTipText     =   "信息复制帮助"
               Top             =   1545
               Width           =   545
            End
            Begin VB.ComboBox Combo3 
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H00FF0000&
               Height          =   300
               ItemData        =   "14.frx":1C957
               Left            =   2895
               List            =   "14.frx":1C96A
               TabIndex        =   43
               Text            =   "含有"
               ToolTipText     =   "查找模式"
               Top             =   1080
               Width           =   1125
            End
            Begin VB.ComboBox Combo2 
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H00FF0000&
               Height          =   300
               ItemData        =   "14.frx":1C997
               Left            =   4005
               List            =   "14.frx":1C9A1
               TabIndex        =   9
               Text            =   "全部"
               ToolTipText     =   "查找范围"
               Top             =   1080
               Width           =   870
            End
            Begin VB.TextBox findbox 
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   18
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00E335DF&
               Height          =   495
               Left            =   180
               OLEDragMode     =   1  'Automatic
               OLEDropMode     =   2  'Automatic
               TabIndex        =   8
               Top             =   450
               Width           =   3495
            End
            Begin VB.CommandButton find 
               BackColor       =   &H00C0FFFF&
               Caption         =   "查找"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   15
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   3810
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   420
               Width           =   1215
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00FFC0C0&
               X1              =   2355
               X2              =   2355
               Y1              =   1050
               Y2              =   930
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00FFC0C0&
               X1              =   75
               X2              =   195
               Y1              =   810
               Y2              =   810
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00FFC0C0&
               X1              =   75
               X2              =   75
               Y1              =   810
               Y2              =   1050
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00FFC0C0&
               X1              =   90
               X2              =   2370
               Y1              =   1050
               Y2              =   1050
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "信息整理"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1875
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Width           =   4095
            Begin VB.CommandButton del 
               BackColor       =   &H00C0FFFF&
               Caption         =   "删除"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   15
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1440
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   660
               Width           =   1215
            End
            Begin VB.CommandButton add 
               BackColor       =   &H00C0FFFF&
               Caption         =   "添加"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   15
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   660
               Width           =   1215
            End
            Begin VB.CommandButton make 
               BackColor       =   &H00C0FFFF&
               Caption         =   "编辑"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   15
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   2760
               Style           =   1  'Graphical
               TabIndex        =   3
               Top             =   660
               Width           =   1215
            End
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   45
         Top             =   15
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Index           =   1
         Left            =   135
         TabIndex        =   42
         Top             =   1110
         Width           =   975
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   41
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "班级"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   40
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "省份"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   39
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "市\县\区"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Index           =   5
         Left            =   120
         TabIndex        =   38
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "街道"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Index           =   6
         Left            =   120
         TabIndex        =   37
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "委\村"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   36
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "小区\屯"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Index           =   8
         Left            =   6240
         TabIndex        =   35
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "具体地址"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Index           =   9
         Left            =   6240
         TabIndex        =   34
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "父亲姓名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Index           =   10
         Left            =   6240
         TabIndex        =   33
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "父亲电话"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Index           =   11
         Left            =   6240
         TabIndex        =   32
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "母亲姓名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Index           =   12
         Left            =   6240
         TabIndex        =   31
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "母亲电话"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Index           =   13
         Left            =   6240
         TabIndex        =   30
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "其他"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.86
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Index           =   14
         Left            =   6240
         TabIndex        =   29
         Top             =   4800
         Width           =   975
      End
   End
End
Attribute VB_Name = "c14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE& = &H1
Private Const SWP_NOMOVE& = &H2

Private Sub add_Click()
Data1.Recordset.AddNew
For i = 1 To 14
Text(i).Locked = False
Next
make.Caption = "查看"
Text(1).SetFocus
Open "log.log" For Append As #1
Print #1, Now & vbCrLf & "Add"
Close #1
End Sub



Private Sub add_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If add.BackColor <> &HFFFF& Then add.BackColor = &HFFFF&
End Sub



Private Sub cap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Command1.BackColor <> &HC0FFFF Then Command1.BackColor = &HC0FFFF
If List1.BackColor <> &HC0FFFF Then List1.BackColor = &HC0FFFF
If help(1).BackColor <> &HC0FFC0 Then help(1).BackColor = &HC0FFC0
If Frame2.Caption <> "信息复制" Then Frame2.Caption = "信息复制"
End Sub
Private Sub Check1_Click()
If Check1.Value = 1 Then
For i = 0 To 12
List1.Selected(i) = True
Next
Else
For i = 0 To 12
List1.Selected(i) = False
Next
End If
End Sub


Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Command1.BackColor <> &HC0FFFF Then Command1.BackColor = &HC0FFFF
If List1.BackColor <> &HC0FFFF Then List1.BackColor = &HC0FFFF
If help(1).BackColor <> &HC0FFC0 Then help(1).BackColor = &HC0FFC0
If Frame2.Caption <> "信息复制" Then Frame2.Caption = "信息复制"
End Sub

Private Sub clearreco_Click()
If MsgBox("是否清空查找历史？", vbExclamation + vbOKCancel, "三十一中学生信息管理系统") = vbOK Then
查找历史.Clear
MsgBox "查找历史已清除"
Open "log.log" For Append As #1
Print #1, Now & vbCrLf & "ClearRecord"
Close #1
End If
End Sub

Private Sub Combo2_Change()
Combo2.Text = chazhao
End Sub

Private Sub Combo2_Click()
chazhaomoshi = Combo2.Text
chazhao = Combo2.Text
End Sub


Private Sub Combo3_Change()
Combo3.BackColor = &HFFFF&
End Sub

Private Sub Command1_Click()
For i = 0 To 12
If List1.Selected(i) = False Then m = m + 1
If List1.Selected(i) = True Then cont = cont & List1.List(i)
Next
If cont = "" Then cont = "None"
Open "log.log" For Append As #1
Print #1, Now & vbCrLf & "Copy(" & (13 - m) & "): " & cont
Close #1
If m = 13 Then a = MsgBox("你还没有选择内容", vbCritical, "三十一中学生信息管理系统"): GoTo 22
Clipboard.Clear
cl = Text(1).Text
For i = 0 To 12
If cap.Value = 1 Then
If muti.Value = 1 Then
If List1.Selected(i) = True Then cl = cl & vbCrLf & List1.List(i) & ":" & Text(i + 2).Text
Else
If List1.Selected(i) = True Then cl = cl & "  " & List1.List(i) & ":" & Text(i + 2).Text
End If
Else
If i = 1 Then cl = ""
If muti.Value = 1 Then
If List1.Selected(i) = True Then cl = cl & vbCrLf & Text(i + 2).Text
Else
If List1.Selected(i) = True Then cl = cl & "  " & Text(i + 2).Text
End If
End If
Next
Clipboard.SetText cl
22  End Sub






Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Command1.BackColor <> &H97FBCB Then Command1.BackColor = &H97FBCB
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
candrag = False
End Sub
Private Sub Command1_OLECompleteDrag(Effect As Long)
Command1.Refresh
List1.Refresh
End Sub



Private Sub cSysTray1_MouseMove(Id As Long)

End Sub

Private Sub Data1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Data1.BackColor = &HFFFFC0
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
Open "log.log" For Append As #1
Print #1, Now & vbCrLf & "Move"
Close #1
End Sub

Private Sub del_Click()
If Data1.Recordset.RecordCount > 0 Then Data1.Recordset.Delete: Data1.Refresh
Open "log.log" For Append As #1
Print #1, Now & vbCrLf & "Delete"
Close #1
End Sub

Private Sub del_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If del.BackColor <> &HFFFF& Then del.BackColor = &HFFFF&
End Sub

Private Sub errcon_Click()
errcon.Visible = False
End Sub

Private Sub find_Click()
If Trim(findbox.Text) = "" Then MsgBox "查找内容不能为空！", , "三十一中学生信息管理系统": GoTo 99
Open "log.log" For Append As #1
Print #1, Now & vbCrLf & "Find: " & findbox.Text
Close #1
If findbox.Text = "*" Then MsgBox "“*”不是一个有效的查找字符串": findbox.Text = "": GoTo 99
For i = 1 To Len(findbox.Text)
If Mid(findbox.Text, i, 1) = "*" And Mid(findbox.Text, i + 1) = "*" Then MsgBox "“" & findbox.Text & "”不是一个有效的查找字符串": findbox.Text = "": GoTo 99
Next
If InStr(1, findbox.Text, "*") >= 1 Then Combo3.Text = "通配符*"
If Combo3.Text = "含有" Then
If chazhaomoshi = "全部" Then
    Data1.Recordset.FindFirst "name Like '*" & Trim(findbox.Text) & "*'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '*" & Trim(findbox.Text) & "*'" Else findout = 1: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '*" & Trim(findbox.Text) & "*'" Else findout = 6: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '*" & Trim(findbox.Text) & "*'" Else findout = 7: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '*" & Trim(findbox.Text) & "*'" Else findout = 8: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '*" & Trim(findbox.Text) & "*'" Else findout = 9: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '*" & Trim(findbox.Text) & "*'" Else findout = 10: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '*" & Trim(findbox.Text) & "*'" Else findout = 11: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '*" & Trim(findbox.Text) & "*'" Else findout = 12: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '*" & Trim(findbox.Text) & "*'" Else findout = 13: GoTo 55
    findout = 14: GoTo 55
ElseIf chazhaomoshi = "下一个" And (Not Data1.Recordset.EOF) Then
    Data1.Recordset.FindNext "name Like '*" & Trim(findbox.Text) & "*'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "street Like '*" & Trim(findbox.Text) & "*'" Else findout = 1: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "village Like '*" & Trim(findbox.Text) & "*'" Else findout = 6: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "area Like '*" & Trim(findbox.Text) & "*'" Else findout = 7: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "number Like '*" & Trim(findbox.Text) & "*'" Else findout = 8: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fname Like '*" & Trim(findbox.Text) & "*'" Else findout = 9: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fphone Like '*" & Trim(findbox.Text) & "*'" Else findout = 10: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mname Like '*" & Trim(findbox.Text) & "*'" Else findout = 11: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mohone Like '*" & Trim(findbox.Text) & "*'" Else findout = 12: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "other Like '*" & Trim(findbox.Text) & "*'" Else findout = 13: GoTo 55
    findout = 14: GoTo 55
ElseIf chazhaomoshi = "下一个" And Data1.Recordset.EOF Then
MsgBox "这已经是最后一条数据了呢", , "三十一中学生信息管理系统"
Else
Data1.Recordset.FindFirst "name Like '*" & Trim(findbox.Text) & "*'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '*" & Trim(findbox.Text) & "*'" Else findout = 1: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '*" & Trim(findbox.Text) & "*'" Else findout = 6: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '*" & Trim(findbox.Text) & "*'" Else findout = 7: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '*" & Trim(findbox.Text) & "*'" Else findout = 8: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '*" & Trim(findbox.Text) & "*'" Else findout = 9: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '*" & Trim(findbox.Text) & "*'" Else findout = 10: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '*" & Trim(findbox.Text) & "*'" Else findout = 11: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '*" & Trim(findbox.Text) & "*'" Else findout = 12: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '*" & Trim(findbox.Text) & "*'" Else findout = 13: GoTo 55
    findout = 14: GoTo 55
End If
55 If Data1.Recordset.NoMatch Then MsgBox "没有找到含有“" & Trim(findbox.Text) & "”的内容", vbQuestion, "三十一中学生信息管理系统":: chazhaomoshi = "全部": chazhao = "全部": Combo2.Text = "全部": GoTo 99
                                                        ElseIf Combo3.Text = "开头为" Then
If chazhaomoshi = "全部" Then
    Data1.Recordset.FindFirst "name Like '" & Trim(findbox.Text) & "*'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '" & Trim(findbox.Text) & "*'" Else findout = 1: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '" & Trim(findbox.Text) & "*'" Else findout = 6: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '" & Trim(findbox.Text) & "*'" Else findout = 7: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '" & Trim(findbox.Text) & "*'" Else findout = 8: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '" & Trim(findbox.Text) & "*'" Else findout = 9: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '" & Trim(findbox.Text) & "*'" Else findout = 10: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '" & Trim(findbox.Text) & "*'" Else findout = 11: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '" & Trim(findbox.Text) & "*'" Else findout = 12: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '" & Trim(findbox.Text) & "*'" Else findout = 13: GoTo 56
    findout = 14: GoTo 56
ElseIf chazhaomoshi = "下一个" And (Not Data1.Recordset.EOF) Then
    Data1.Recordset.FindNext "name Like '" & Trim(findbox.Text) & "*'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "street Like '" & Trim(findbox.Text) & "*'" Else findout = 1: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "village Like '" & Trim(findbox.Text) & "*'" Else findout = 6: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "area Like '" & Trim(findbox.Text) & "*'" Else findout = 7: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "number Like '" & Trim(findbox.Text) & "*'" Else findout = 8: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fname Like '" & Trim(findbox.Text) & "*'" Else findout = 9: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fphone Like '" & Trim(findbox.Text) & "*'" Else findout = 10: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mname Like '" & Trim(findbox.Text) & "*'" Else findout = 11: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mohone Like '" & Trim(findbox.Text) & "*'" Else findout = 12: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "other Like '" & Trim(findbox.Text) & "*'" Else findout = 13: GoTo 56
    findout = 14: GoTo 56
ElseIf chazhaomoshi = "下一个" And Data1.Recordset.EOF Then
MsgBox "这已经是最后一条数据了呢", , "三十一中学生信息管理系统"
Else
Data1.Recordset.FindFirst "name Like '" & Trim(findbox.Text) & "*'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '" & Trim(findbox.Text) & "*'" Else findout = 1: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '" & Trim(findbox.Text) & "*'" Else findout = 6: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '" & Trim(findbox.Text) & "*'" Else findout = 7: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '" & Trim(findbox.Text) & "*'" Else findout = 8: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '" & Trim(findbox.Text) & "*'" Else findout = 9: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '" & Trim(findbox.Text) & "*'" Else findout = 10: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '" & Trim(findbox.Text) & "*'" Else findout = 11: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '" & Trim(findbox.Text) & "*'" Else findout = 12: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '" & Trim(findbox.Text) & "*'" Else findout = 13: GoTo 56
    findout = 14: GoTo 56
End If
56 If Data1.Recordset.NoMatch Then MsgBox "没有找到开头为“" & Trim(findbox.Text) & "”的内容", vbQuestion, "三十一中学生信息管理系统"
                                                            ElseIf Combo3.Text = "结尾为" Then
If chazhaomoshi = "全部" Then
    Data1.Recordset.FindFirst "name Like '*" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '*" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '*" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '*" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '*" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '*" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '*" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '*" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '*" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '*" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 57
    findout = 14: GoTo 57
ElseIf chazhaomoshi = "下一个" And (Not Data1.Recordset.EOF) Then
    Data1.Recordset.FindNext "name Like '*" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "street Like '*" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "village Like '*" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "area Like '*" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "number Like '*" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fname Like '*" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fphone Like '*" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mname Like '*" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mohone Like '*" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "other Like '*" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 57
    findout = 14: GoTo 57
ElseIf chazhaomoshi = "下一个" And Data1.Recordset.EOF Then
MsgBox "这已经是最后一条数据了呢", , "三十一中学生信息管理系统"
Else
Data1.Recordset.FindFirst "name Like '*" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '*" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '*" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '*" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '*" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '*" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '*" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '*" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '*" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '*" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 57
    findout = 14: GoTo 57
End If
57 If Data1.Recordset.NoMatch Then MsgBox "没有找到结尾为“" & Trim(findbox.Text) & "”的内容", vbQuestion, "三十一中学生信息管理系统"
ElseIf Combo3.Text = "严格查找" Then
If chazhaomoshi = "全部" Then
    Data1.Recordset.FindFirst "name = '" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street = '" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village = '" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area = '" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number = '" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname = '" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone = '" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname = '" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone = '" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other = '" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 58
    findout = 14: GoTo 58
ElseIf chazhaomoshi = "下一个" And (Not Data1.Recordset.EOF) Then
    Data1.Recordset.FindNext "name = '" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "street = '" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "village = '" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "area = '" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "number = '" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fname = '" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fphone = '" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mname = '" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mohone = '" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "other = '" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 58
    findout = 14: GoTo 58
ElseIf chazhaomoshi = "下一个" And Data1.Recordset.EOF Then
MsgBox "这已经是最后一条数据了呢", , "三十一中学生信息管理系统"
Else
Data1.Recordset.FindFirst "name = '" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street = '" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village = '" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area = '" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number = '" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname = '" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone = '" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname = '" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone = '" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other = '" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 58
    findout = 14: GoTo 58
End If
58 If Data1.Recordset.NoMatch Then MsgBox "没有找到“" & Trim(findbox.Text) & "”", vbQuestion, "三十一中学生信息管理系统"
ElseIf Combo3.Text = "通配符*" Then
If stano Then GoTo 23
If chazhaomoshi = "全部" Then
23     Data1.Recordset.FindFirst "name Like '" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 67
    findout = 14: GoTo 67
ElseIf chazhaomoshi = "下一个" And (Not Data1.Recordset.EOF) Then
    Data1.Recordset.FindNext "name Like '" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "street Like '" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "village Like '" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "area Like '" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "number Like '" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fname Like '" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fphone Like '" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mname Like '" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mohone Like '" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "other Like '" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 67
    findout = 14: GoTo 67
ElseIf chazhaomoshi = "下一个" And Data1.Recordset.EOF Then
MsgBox "这已经是最后一条数据了呢", , "三十一中学生信息管理系统"
Else
Data1.Recordset.FindFirst "name Like '" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 67
    findout = 14: GoTo 67
    End If
67 If Data1.Recordset.NoMatch Then MsgBox "没有找到“" & Trim(findbox.Text) & "”的内容", vbQuestion, "三十一中学生信息管理系统":: chazhaomoshi = "全部": chazhao = "全部": Combo2.Text = "全部": GoTo 99
End If
On Error GoTo 98
Text(findout).SelStart = InStr(1, Text(findout).Text, Trim(findbox.Text)) - 1
Text(findout).SelLength = Len(Trim(findbox.Text))
Text(findout).SetFocus
Debug.Print InStr(Text(findout).Text, Trim(findbox.Text))
98 If Combo3.Text = "通配符*" Then
Text(findout).SelStart = 0
Text(findout).SelLength = Len(Trim(Text(findout).Text)) + 1
Text(findout).SetFocus
End If
For i = 1 To 查找历史.ListCount
If 查找历史.List(i) = findbox.Text Then repaired = True
Next
If Not repaired Then 查找历史.AddItem findbox.Text
chazhaomoshi = "下一个"
chazhao = "下一个"
Combo2.Text = "下一个"
99 End Sub

Private Sub find_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If find.BackColor <> &HFFFF& Then find.BackColor = &HFFFF&
End Sub

Private Sub findbox_Change()
chazhaomoshi = "全部"
chazhao = "全部"
Combo2.Text = "全部"
repaired = False
End Sub

Private Sub findbox_GotFocus()
findbox.BackColor = &H97FBCB
End Sub



Private Sub findbox_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Trim(findbox.Text) = "" Then MsgBox "查找内容不能为空！", , "三十一中学生信息管理系统": GoTo 99
If findbox.Text = "*" Then MsgBox "“*”不是一个有效的查找字符串": findbox.Text = "": GoTo 99
For i = 1 To Len(findbox.Text)
If Mid(findbox.Text, i, 1) = "*" And Mid(findbox.Text, i + 1) = "*" Then MsgBox "“" & findbox.Text & "”不是一个有效的查找字符串": findbox.Text = "": GoTo 99
Next
If InStr(1, findbox.Text, "*") >= 1 Then Combo3.Text = "通配符*"
If Combo3.Text = "含有" Then
If chazhaomoshi = "全部" Then
    Data1.Recordset.FindFirst "name Like '*" & Trim(findbox.Text) & "*'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '*" & Trim(findbox.Text) & "*'" Else findout = 1: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '*" & Trim(findbox.Text) & "*'" Else findout = 6: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '*" & Trim(findbox.Text) & "*'" Else findout = 7: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '*" & Trim(findbox.Text) & "*'" Else findout = 8: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '*" & Trim(findbox.Text) & "*'" Else findout = 9: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '*" & Trim(findbox.Text) & "*'" Else findout = 10: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '*" & Trim(findbox.Text) & "*'" Else findout = 11: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '*" & Trim(findbox.Text) & "*'" Else findout = 12: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '*" & Trim(findbox.Text) & "*'" Else findout = 13: GoTo 55
    findout = 14: GoTo 55
ElseIf chazhaomoshi = "下一个" And (Not Data1.Recordset.EOF) Then
    Data1.Recordset.FindNext "name Like '*" & Trim(findbox.Text) & "*'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "street Like '*" & Trim(findbox.Text) & "*'" Else findout = 1: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "village Like '*" & Trim(findbox.Text) & "*'" Else findout = 6: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "area Like '*" & Trim(findbox.Text) & "*'" Else findout = 7: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "number Like '*" & Trim(findbox.Text) & "*'" Else findout = 8: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fname Like '*" & Trim(findbox.Text) & "*'" Else findout = 9: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fphone Like '*" & Trim(findbox.Text) & "*'" Else findout = 10: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mname Like '*" & Trim(findbox.Text) & "*'" Else findout = 11: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mohone Like '*" & Trim(findbox.Text) & "*'" Else findout = 12: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "other Like '*" & Trim(findbox.Text) & "*'" Else findout = 13: GoTo 55
    findout = 14: GoTo 55
ElseIf chazhaomoshi = "下一个" And Data1.Recordset.EOF Then
MsgBox "这已经是最后一条数据了呢", , "三十一中学生信息管理系统"
Else
Data1.Recordset.FindFirst "name Like '*" & Trim(findbox.Text) & "*'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '*" & Trim(findbox.Text) & "*'" Else findout = 1: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '*" & Trim(findbox.Text) & "*'" Else findout = 6: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '*" & Trim(findbox.Text) & "*'" Else findout = 7: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '*" & Trim(findbox.Text) & "*'" Else findout = 8: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '*" & Trim(findbox.Text) & "*'" Else findout = 9: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '*" & Trim(findbox.Text) & "*'" Else findout = 10: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '*" & Trim(findbox.Text) & "*'" Else findout = 11: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '*" & Trim(findbox.Text) & "*'" Else findout = 12: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '*" & Trim(findbox.Text) & "*'" Else findout = 13: GoTo 55
    findout = 14: GoTo 55
End If
55 If Data1.Recordset.NoMatch Then MsgBox "没有找到含有“" & Trim(findbox.Text) & "”的内容", vbQuestion, "三十一中学生信息管理系统":: chazhaomoshi = "全部": chazhao = "全部": Combo2.Text = "全部": GoTo 99
                                                        ElseIf Combo3.Text = "开头为" Then
If chazhaomoshi = "全部" Then
    Data1.Recordset.FindFirst "name Like '" & Trim(findbox.Text) & "*'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '" & Trim(findbox.Text) & "*'" Else findout = 1: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '" & Trim(findbox.Text) & "*'" Else findout = 6: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '" & Trim(findbox.Text) & "*'" Else findout = 7: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '" & Trim(findbox.Text) & "*'" Else findout = 8: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '" & Trim(findbox.Text) & "*'" Else findout = 9: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '" & Trim(findbox.Text) & "*'" Else findout = 10: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '" & Trim(findbox.Text) & "*'" Else findout = 11: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '" & Trim(findbox.Text) & "*'" Else findout = 12: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '" & Trim(findbox.Text) & "*'" Else findout = 13: GoTo 56
    findout = 14: GoTo 56
ElseIf chazhaomoshi = "下一个" And (Not Data1.Recordset.EOF) Then
    Data1.Recordset.FindNext "name Like '" & Trim(findbox.Text) & "*'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "street Like '" & Trim(findbox.Text) & "*'" Else findout = 1: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "village Like '" & Trim(findbox.Text) & "*'" Else findout = 6: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "area Like '" & Trim(findbox.Text) & "*'" Else findout = 7: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "number Like '" & Trim(findbox.Text) & "*'" Else findout = 8: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fname Like '" & Trim(findbox.Text) & "*'" Else findout = 9: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fphone Like '" & Trim(findbox.Text) & "*'" Else findout = 10: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mname Like '" & Trim(findbox.Text) & "*'" Else findout = 11: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mohone Like '" & Trim(findbox.Text) & "*'" Else findout = 12: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "other Like '" & Trim(findbox.Text) & "*'" Else findout = 13: GoTo 56
    findout = 14: GoTo 56
ElseIf chazhaomoshi = "下一个" And Data1.Recordset.EOF Then
MsgBox "这已经是最后一条数据了呢", , "三十一中学生信息管理系统"
Else
Data1.Recordset.FindFirst "name Like '" & Trim(findbox.Text) & "*'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '" & Trim(findbox.Text) & "*'" Else findout = 1: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '" & Trim(findbox.Text) & "*'" Else findout = 6: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '" & Trim(findbox.Text) & "*'" Else findout = 7: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '" & Trim(findbox.Text) & "*'" Else findout = 8: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '" & Trim(findbox.Text) & "*'" Else findout = 9: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '" & Trim(findbox.Text) & "*'" Else findout = 10: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '" & Trim(findbox.Text) & "*'" Else findout = 11: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '" & Trim(findbox.Text) & "*'" Else findout = 12: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '" & Trim(findbox.Text) & "*'" Else findout = 13: GoTo 56
    findout = 14: GoTo 56
End If
56 If Data1.Recordset.NoMatch Then MsgBox "没有找到开头为“" & Trim(findbox.Text) & "”的内容", vbQuestion, "三十一中学生信息管理系统"
                                                            ElseIf Combo3.Text = "结尾为" Then
If chazhaomoshi = "全部" Then
    Data1.Recordset.FindFirst "name Like '*" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '*" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '*" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '*" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '*" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '*" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '*" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '*" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '*" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '*" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 57
    findout = 14: GoTo 57
ElseIf chazhaomoshi = "下一个" And (Not Data1.Recordset.EOF) Then
    Data1.Recordset.FindNext "name Like '*" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "street Like '*" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "village Like '*" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "area Like '*" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "number Like '*" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fname Like '*" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fphone Like '*" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mname Like '*" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mohone Like '*" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "other Like '*" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 57
    findout = 14: GoTo 57
ElseIf chazhaomoshi = "下一个" And Data1.Recordset.EOF Then
MsgBox "这已经是最后一条数据了呢", , "三十一中学生信息管理系统"
Else
Data1.Recordset.FindFirst "name Like '*" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '*" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '*" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '*" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '*" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '*" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '*" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '*" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '*" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '*" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 57
    findout = 14: GoTo 57
End If
57 If Data1.Recordset.NoMatch Then MsgBox "没有找到结尾为“" & Trim(findbox.Text) & "”的内容", vbQuestion, "三十一中学生信息管理系统"
ElseIf Combo3.Text = "严格查找" Then
If chazhaomoshi = "全部" Then
    Data1.Recordset.FindFirst "name = '" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street = '" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village = '" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area = '" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number = '" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname = '" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone = '" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname = '" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone = '" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other = '" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 58
    findout = 14: GoTo 58
ElseIf chazhaomoshi = "下一个" And (Not Data1.Recordset.EOF) Then
    Data1.Recordset.FindNext "name = '" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "street = '" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "village = '" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "area = '" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "number = '" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fname = '" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fphone = '" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mname = '" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mohone = '" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "other = '" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 58
    findout = 14: GoTo 58
ElseIf chazhaomoshi = "下一个" And Data1.Recordset.EOF Then
MsgBox "这已经是最后一条数据了呢", , "三十一中学生信息管理系统"
Else
Data1.Recordset.FindFirst "name = '" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street = '" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village = '" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area = '" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number = '" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname = '" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone = '" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname = '" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone = '" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other = '" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 58
    findout = 14: GoTo 58
End If
58 If Data1.Recordset.NoMatch Then MsgBox "没有找到“" & Trim(findbox.Text) & "”", vbQuestion, "三十一中学生信息管理系统"
ElseIf Combo3.Text = "通配符*" Then
If chazhaomoshi = "全部" Then
    Data1.Recordset.FindFirst "name Like '" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 67
    findout = 14: GoTo 67
ElseIf chazhaomoshi = "下一个" And (Not Data1.Recordset.EOF) Then
    Data1.Recordset.FindNext "name Like '" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "street Like '" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "village Like '" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "area Like '" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "number Like '" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fname Like '" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fphone Like '" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mname Like '" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mohone Like '" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "other Like '" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 67
    findout = 14: GoTo 67
ElseIf chazhaomoshi = "下一个" And Data1.Recordset.EOF Then
MsgBox "这已经是最后一条数据了呢", , "三十一中学生信息管理系统"
Else
Data1.Recordset.FindFirst "name Like '" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 67
    findout = 14: GoTo 67
    End If
67 If Data1.Recordset.NoMatch Then MsgBox "没有找到“" & Trim(findbox.Text) & "”的内容", vbQuestion, "三十一中学生信息管理系统":: chazhaomoshi = "全部": chazhao = "全部": Combo2.Text = "全部": GoTo 99
End If

On Error GoTo 98
Text(findout).SelStart = InStr(1, Text(findout).Text, Trim(findbox.Text)) - 1
Text(findout).SelLength = Len(Trim(findbox.Text))
Text(findout).SetFocus
Debug.Print InStr(Text(findout).Text, Trim(findbox.Text))
98 If Combo3.Text = "通配符*" Then
Text(findout).SelStart = 0
Text(findout).SelLength = Len(Trim(Text(findout).Text)) + 1
Text(findout).SetFocus
End If
For i = 1 To 查找历史.ListCount
If 查找历史.List(i) = findbox.Text Then repaired = True
Next
If Not repaired Then 查找历史.AddItem findbox.Text
chazhaomoshi = "下一个"
chazhao = "下一个"
Combo2.Text = "下一个"
99 End If
End Sub

Private Sub findbox_LostFocus()
findbox.BackColor = &HC0FFFF
End Sub

Private Sub findbox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
findbox.BackColor = &HFFFF&
End Sub

Private Sub Form_Load()
zdd = False
Data1.DatabaseName = App.Path & "\三十一中学生信息库.mdb"
Data1.RecordSource = se
If se = "recoall" Then GoTo 46
nam = Right(se, 3)
Data1.Caption = "              庄河31初级中学" & Left(nam, 1) & "年" & Right(nam, 2) & "班学生信息库"
GoTo 88
Label1.Caption = "共有记录" & Data1.Recordset.RecordCount & "条"
GoTo 88
46 Data1.Caption = "               庄河31初级中学学生信息库"
88 End Sub



Private Sub Form_Resize()
On Error GoTo 33
toup.Top = 0
toup.Left = c14.Width - toup.Width - 200
Picture1.Top = c14.Height / 2 - Picture1.Height / 2 - 400
Picture1.Left = c14.Width / 2 - Picture1.Width / 2 - 400
If c14.Width < Picture1.Width Then c14.Width = Picture1.Width
If c14.Height < Picture1.Height Then c14.Height = Picture1.Height
33 End Sub

Private Sub Form_Unload(Cancel As Integer)
welcome.Show
End Sub








Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Command1.BackColor <> &HC0FFFF Then Command1.BackColor = &HC0FFFF
If List1.BackColor <> &HC0FFFF Then List1.BackColor = &HC0FFFF
If help(1).BackColor <> &HC0FFC0 Then help(1).BackColor = &HC0FFC0
If Frame2.Caption <> "信息复制" Then Frame2.Caption = "信息复制"
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If find.BackColor <> &HC0FFFF Then find.BackColor = &HC0FFFF
If findbox.BackColor <> &HC0FFC0 Then findbox.BackColor = &HC0FFFF
If help(0).BackColor <> &HC0FFC0 Then help(0).BackColor = &HC0FFC0
9 End Sub



Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If add.BackColor <> &HC0FFFF Then add.BackColor = &HC0FFFF
If del.BackColor <> &HC0FFFF Then del.BackColor = &HC0FFFF
If make.BackColor <> &HC0FFFF Then make.BackColor = &HC0FFFF
End Sub



Private Sub help_Click(Index As Integer)
If help(0).BackColor <> &HC0FFC0 Then help(0).BackColor = &HC0FFC0
If help(1).BackColor <> &HC0FFC0 Then help(1).BackColor = &HC0FFC0
If Index = 0 Then
    If MsgBox("信息复制栏具有强大的OLE拖放及智能信息处理功能。" & vbCrLf & "    将外部文字（如：某某询问“枫小凝是男的还是女的，他爸爸叫什么？Where does he/she live?”）拖入显示复制内容的列表框，系统将会自动匹配合适的选项。" & vbCrLf & "    (极少数情况下，在输入信息时，若语言较为复杂或特殊（如隐语），智能信息处理功能可能会失误，敬请谅解。)" & vbCrLf & vbCrLf & "即将为您播放帮助视频，是否观看（由于版本不同，帮助视频中的界面可能和本系统有所差异）？", vbInformation + vbYesNoCancel, "三十一中学生信息管理系统") = vbYes Then
On Error GoTo 43
Shell ("explorer " & App.Path & "\“信息复制”框教程.mp4")
Exit Sub
43 MsgBox "帮助文件不存在，请确保帮助文件在程序文件夹内，且名称为““信息复制”框教程.mp4”"
    End If
Else
MsgBox "查找模式列表中的“通配符*”介绍" & vbCrLf & "        此选项可进行更加个性化的查找，“*”代表一个若干长度的字符串(字符串值可以为空)，如“枫*凝”代表开头为“枫”，结尾“凝”的所有记录。" & vbCrLf & "(在其它查找模式下，若查找内容含有“*”，将会自动转换到“通配符*”模式"
End If

End Sub



Private Sub help_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If help(Index).BackColor <> &HC0E0FF Then help(Index).BackColor = &HC0E0FF
End Sub





Private Sub helpcon_DblClick()
helpcon.Visible = ture: Option1.Enabled = True: Check1.Enabled = True: muti.Enabled = True: cap.Enabled = True: List1.Enabled = True
End Sub



Private Sub helpcon_LostFocus()
helpcon.Visible = ture: Option1.Enabled = True: Check1.Enabled = True: muti.Enabled = True: cap.Enabled = True: List1.Enabled = True
End Sub

Private Sub helpcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Frame2.Caption <> "未能识别" Then Frame2.Caption = "未能识别"
End Sub













Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
For b = 0 To 12
List1.Selected(b) = False
Next b
If Effect <> 0 Then Frame2.Caption = "已完成识别"
sl = List1.SelCount
If Data.GetFormat(vbCFText) Then
Gett = Data.GetData(vbCFText)
Gett = Trim(Gett)
Gett = LCase(Gett)
Open "log.log" For Append As #1
Print #1, Now & vbCrLf; "Drag:" & Gett
Close #1
If InStr(1, Gett, "性别") > 0 Or InStr(1, Gett, "男") > 0 Or InStr(1, Gett, "女") > 0 Or InStr(1, Gett, "sex") > 0 Or InStr(1, Gett, "gender") > 0 _
Then List1.Selected(0) = True
If InStr(1, Gett, "班") > 0 Or InStr(1, Gett, "class") > 0 Then List1.Selected(1) = True
If InStr(1, Gett, "省") > 0 Or InStr(1, Gett, "province") > 0 Then List1.Selected(2) = True
If InStr(1, Gett, "市") > 0 Or InStr(1, Gett, "city") > 0 Or InStr(1, Gett, "town") > 0 Or InStr(1, Gett, "城") > 0 Then List1.Selected(3) = True
If InStr(1, Gett, "街道") > 0 Then List1.Selected(4) = True
If InStr(1, Gett, "委") > 0 Or InStr(1, Gett, "村") > 0 Then List1.Selected(5) = True
If InStr(1, Gett, "小区") > 0 Or InStr(1, Gett, "屯") > 0 Then List1.Selected(6) = True
If InStr(1, Gett, "具体地址") > 0 Or InStr(1, Gett, "详细地址") > 0 Then List1.Selected(6) = True
If (InStr(1, Gett, "父亲") > 0 Or InStr(1, Gett, "爸") > 0) And (InStr(1, Gett, "名") > 0 Or InStr(1, Gett, "name") > 0 _
Or InStr(1, Gett, "叫") > 0) Then List1.Selected(8) = True
If (InStr(1, Gett, "父亲") > 0 Or InStr(1, Gett, "爸") > 0) And (InStr(1, Gett, "手机") > 0 Or InStr(1, Gett, "号码") > 0 _
Or InStr(1, Gett, "电话") > 0) Then List1.Selected(9) = True
If (InStr(1, Gett, "母亲") > 0 Or InStr(1, Gett, "妈") > 0) And (InStr(1, Gett, "名") > 0 Or InStr(1, Gett, "name") > 0 _
Or InStr(1, Gett, "叫") > 0) Then List1.Selected(10) = True
If (InStr(1, Gett, "母亲") > 0 Or InStr(1, Gett, "妈") > 0) And (InStr(1, Gett, "手机") > 0 Or InStr(1, Gett, "号码") > 0 _
Or InStr(1, Gett, "电话") > 0) Then List1.Selected(11) = True
If InStr(1, Gett, "地") > 0 Or InStr(1, Gett, "街") > 0 Or InStr(1, Gett, "where") > 0 Or InStr(1, Gett, "住") > 0 Then List1.Selected(2) = True: List1.Selected(3) = True: List1.Selected(4) = True: List1.Selected(5) = True: List1.Selected(6) = True: List1.Selected(7) = True
End If
If InStr(1, Gett, "谁") > 0 Or InStr(1, Gett, "who") > 0 Then
For i = 0 To 12
List1.Selected(i) = True
Next i
End If
If List1.SelCount = sl Then Frame2.Caption = "未能识别"
If Gett = "" Then helpcon.Text = "原因：" & "放下的内容不是文本或为空。（双击隐藏）": helpcon.Visible = True:  helpcon.SetFocus: Option1.Enabled = False: Check1.Enabled = False: muti.Enabled = False: cap.Enabled = False: List1.Enabled = False
For j = 0 To 12
If List1.Selected(j) Then knew = True
Next j
If Not knew And Gett <> "" Then helpcon.Text = "原因：" & "“" & Gett & "”中没有可识别的内容，如果该文本中汉字显示为乱码，说明源程序文字处理被加密或编码不一致，请在源程序中复制文本后点击“识别剪贴板”。（双击隐藏）": helpcon.Visible = True: helpcon.SetFocus: Option1.Enabled = False: Check1.Enabled = False: muti.Enabled = False: cap.Enabled = False: List1.Enabled = False

99 End Sub



Private Sub List1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Frame2.Caption = "放下吧，帮你选择"
End Sub



Private Sub make_Click()
For i = 1 To 14
Text(i).Locked = Not Text(i).Locked
Next
If Text(1).Locked Then
make.Caption = "编辑"
make.SetFocus
Else
make.Caption = "查看"
Text(1).SetFocus
End If
Open "log.log" For Append As #1
Print #1, Now & vbCrLf & "Make"
Close #1
End Sub

Private Sub make_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If make.BackColor <> &HFFFF& Then make.BackColor = &HFFFF&
End Sub

Private Sub muti_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Command1.BackColor <> &HC0FFFF Then Command1.BackColor = &HC0FFFF
If List1.BackColor <> &HC0FFFF Then List1.BackColor = &HC0FFFF
If help(1).BackColor <> &HC0FFC0 Then help(1).BackColor = &HC0FFC0
If Frame2.Caption <> "信息复制" Then Frame2.Caption = "信息复制"
End Sub

Private Sub Newf_Click()
num = se
End Sub

Private Sub Option1_Click()
Option1.Value = Not Option1.Value
For i = 0 To 12
List1.Selected(i) = Not List1.Selected(i)
Next
End Sub









Private Sub Option1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Command1.BackColor <> &HC0FFFF Then Command1.BackColor = &HC0FFFF
If List1.BackColor <> &HC0FFFF Then List1.BackColor = &HC0FFFF
If help(1).BackColor <> &HC0FFC0 Then help(1).BackColor = &HC0FFC0
If Frame2.Caption <> "信息复制" Then Frame2.Caption = "信息复制"
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Text(1).Locked Then
For i = 1 To 14
Text(i).BackColor = &H97FBCB
Next
End If
Data1.BackColor = &HC0FFFF
End Sub



Private Sub readclipboard_Click()
For b = 0 To 12
List1.Selected(b) = False
Next b
Gett = Clipboard.GetText
Frame2.Caption = "已完成识别"
sl = List1.SelCount
Gett = Trim(Gett)
Gett = LCase(Gett)
If InStr(1, Gett, "性别") > 0 Or InStr(1, Gett, "男") > 0 Or InStr(1, Gett, "女") > 0 Or InStr(1, Gett, "sex") > 0 Or InStr(1, Gett, "gender") > 0 _
Then List1.Selected(0) = True
If InStr(1, Gett, "班") > 0 Or InStr(1, Gett, "class") > 0 Then List1.Selected(1) = True
If InStr(1, Gett, "省") > 0 Or InStr(1, Gett, "province") > 0 Then List1.Selected(2) = True
If InStr(1, Gett, "市") > 0 Or InStr(1, Gett, "city") > 0 Or InStr(1, Gett, "town") > 0 Or InStr(1, Gett, "城") > 0 Then List1.Selected(3) = True
If InStr(1, Gett, "街道") > 0 Then List1.Selected(4) = True
If InStr(1, Gett, "委") > 0 Or InStr(1, Gett, "村") > 0 Then List1.Selected(5) = True
If InStr(1, Gett, "小区") > 0 Or InStr(1, Gett, "屯") > 0 Then List1.Selected(6) = True
If InStr(1, Gett, "具体地址") > 0 Or InStr(1, Gett, "详细地址") > 0 Then List1.Selected(6) = True
If (InStr(1, Gett, "父亲") > 0 Or InStr(1, Gett, "爸") > 0) And (InStr(1, Gett, "名") > 0 Or InStr(1, Gett, "name") > 0 _
Or InStr(1, Gett, "叫") > 0) Then List1.Selected(8) = True
If (InStr(1, Gett, "父亲") > 0 Or InStr(1, Gett, "爸") > 0) And (InStr(1, Gett, "手机") > 0 Or InStr(1, Gett, "号码") > 0 _
Or InStr(1, Gett, "电话") > 0) Then List1.Selected(9) = True
If (InStr(1, Gett, "母亲") > 0 Or InStr(1, Gett, "妈") > 0) And (InStr(1, Gett, "名") > 0 Or InStr(1, Gett, "name") > 0 _
Or InStr(1, Gett, "叫") > 0) Then List1.Selected(10) = True
If (InStr(1, Gett, "母亲") > 0 Or InStr(1, Gett, "妈") > 0) And (InStr(1, Gett, "手机") > 0 Or InStr(1, Gett, "号码") > 0 _
Or InStr(1, Gett, "电话") > 0) Then List1.Selected(11) = True
If InStr(1, Gett, "地") > 0 Or InStr(1, Gett, "街") > 0 Or InStr(1, Gett, "where") > 0 Or InStr(1, Gett, "住") > 0 Then List1.Selected(2) = True: List1.Selected(3) = True: List1.Selected(4) = True: List1.Selected(5) = True: List1.Selected(6) = True: List1.Selected(7) = True
If InStr(1, Gett, "谁") > 0 Or InStr(1, Gett, "who") > 0 Then
For i = 0 To 12
List1.Selected(i) = True
Next i
End If
If List1.SelCount = sl Then Frame2.Caption = "未能识别"
For j = 0 To 12
If List1.Selected(j) Then knew = True
Next j
If Not knew Then helpcon.Text = "原因：剪贴板中" & "“" & Gett & "”中没有可识别的内容": helpcon.Visible = True: Option1.Enabled = False: Check1.Enabled = False: muti.Enabled = False: cap.Enabled = False: List1.Enabled = False
99 End Sub


Private Sub Text_Change(Index As Integer)
Open "log.log" For Append As #1
Print #1, Now & vbCrLf & "Changed(" & Text(Index).DataField & "): " & Text(Index).Text
Close #1
End Sub

Private Sub Text_GotFocus(Index As Integer)
For i = 1 To 14
Text(i).BackColor = &H97FBCB
Next


End Sub


Private Sub Text_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Text(1).Locked Then
Text(Index).BackColor = &HFFFF80
End If
End Sub

Private Sub Timer1_Timer()
chazhao = chazhaomoshi
For i = 0 To 查找历史.ListCount - 1
If 查找历史.List(i) = "" Then 查找历史.RemoveItem i
Next
End Sub


Private Sub toup_Click()
zdd = Not zdd
If zdd = True Then
toup.BackColor = &HFFFFC0
SetWindowPos c14.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
Else
toup.BackColor = &HC0FFFF
SetWindowPos c14.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End If
End Sub

Private Sub 查找历史_Click()
chazhaomoshi = "全部"
chazhao = "全部"
Combo2.Text = "全部"
findbox.Text = 查找历史.Text
7 If Trim(findbox.Text) = "" Then MsgBox "查找内容不能为空！", , "三十一中学生信息管理系统": GoTo 99
If findbox.Text = "*" Then MsgBox "“*”不是一个有效的查找字符串": findbox.Text = "": GoTo 99
For i = 1 To Len(findbox.Text)
If Mid(findbox.Text, i, 1) = "*" And Mid(findbox.Text, i + 1) = "*" Then MsgBox "“" & findbox.Text & "”不是一个有效的查找字符串": findbox.Text = "": GoTo 99
Next
If InStr(1, findbox.Text, "*") >= 1 Then Combo3.Text = "通配符*"
If Combo3.Text = "含有" Then
If chazhaomoshi = "全部" Then
    Data1.Recordset.FindFirst "name Like '*" & Trim(findbox.Text) & "*'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '*" & Trim(findbox.Text) & "*'" Else findout = 1: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '*" & Trim(findbox.Text) & "*'" Else findout = 6: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '*" & Trim(findbox.Text) & "*'" Else findout = 7: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '*" & Trim(findbox.Text) & "*'" Else findout = 8: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '*" & Trim(findbox.Text) & "*'" Else findout = 9: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '*" & Trim(findbox.Text) & "*'" Else findout = 10: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '*" & Trim(findbox.Text) & "*'" Else findout = 11: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '*" & Trim(findbox.Text) & "*'" Else findout = 12: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '*" & Trim(findbox.Text) & "*'" Else findout = 13: GoTo 55
    findout = 14: GoTo 55
ElseIf chazhaomoshi = "下一个" And (Not Data1.Recordset.EOF) Then
    Data1.Recordset.FindNext "name Like '*" & Trim(findbox.Text) & "*'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "street Like '*" & Trim(findbox.Text) & "*'" Else findout = 1: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "village Like '*" & Trim(findbox.Text) & "*'" Else findout = 6: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "area Like '*" & Trim(findbox.Text) & "*'" Else findout = 7: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "number Like '*" & Trim(findbox.Text) & "*'" Else findout = 8: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fname Like '*" & Trim(findbox.Text) & "*'" Else findout = 9: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fphone Like '*" & Trim(findbox.Text) & "*'" Else findout = 10: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mname Like '*" & Trim(findbox.Text) & "*'" Else findout = 11: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mohone Like '*" & Trim(findbox.Text) & "*'" Else findout = 12: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "other Like '*" & Trim(findbox.Text) & "*'" Else findout = 13: GoTo 55
    findout = 14: GoTo 55
ElseIf chazhaomoshi = "下一个" And Data1.Recordset.EOF Then
MsgBox "这已经是最后一条数据了呢", , "三十一中学生信息管理系统"
Else
Data1.Recordset.FindFirst "name Like '*" & Trim(findbox.Text) & "*'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '*" & Trim(findbox.Text) & "*'" Else findout = 1: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '*" & Trim(findbox.Text) & "*'" Else findout = 6: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '*" & Trim(findbox.Text) & "*'" Else findout = 7: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '*" & Trim(findbox.Text) & "*'" Else findout = 8: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '*" & Trim(findbox.Text) & "*'" Else findout = 9: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '*" & Trim(findbox.Text) & "*'" Else findout = 10: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '*" & Trim(findbox.Text) & "*'" Else findout = 11: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '*" & Trim(findbox.Text) & "*'" Else findout = 12: GoTo 55
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '*" & Trim(findbox.Text) & "*'" Else findout = 13: GoTo 55
    findout = 14: GoTo 55
End If
55 If Data1.Recordset.NoMatch Then MsgBox "没有找到含有“" & Trim(findbox.Text) & "”的内容", vbQuestion, "三十一中学生信息管理系统":: chazhaomoshi = "全部": chazhao = "全部": Combo2.Text = "全部": GoTo 99
                                                        ElseIf Combo3.Text = "开头为" Then
If chazhaomoshi = "全部" Then
    Data1.Recordset.FindFirst "name Like '" & Trim(findbox.Text) & "*'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '" & Trim(findbox.Text) & "*'" Else findout = 1: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '" & Trim(findbox.Text) & "*'" Else findout = 6: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '" & Trim(findbox.Text) & "*'" Else findout = 7: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '" & Trim(findbox.Text) & "*'" Else findout = 8: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '" & Trim(findbox.Text) & "*'" Else findout = 9: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '" & Trim(findbox.Text) & "*'" Else findout = 10: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '" & Trim(findbox.Text) & "*'" Else findout = 11: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '" & Trim(findbox.Text) & "*'" Else findout = 12: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '" & Trim(findbox.Text) & "*'" Else findout = 13: GoTo 56
    findout = 14: GoTo 56
ElseIf chazhaomoshi = "下一个" And (Not Data1.Recordset.EOF) Then
    Data1.Recordset.FindNext "name Like '" & Trim(findbox.Text) & "*'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "street Like '" & Trim(findbox.Text) & "*'" Else findout = 1: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "village Like '" & Trim(findbox.Text) & "*'" Else findout = 6: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "area Like '" & Trim(findbox.Text) & "*'" Else findout = 7: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "number Like '" & Trim(findbox.Text) & "*'" Else findout = 8: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fname Like '" & Trim(findbox.Text) & "*'" Else findout = 9: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fphone Like '" & Trim(findbox.Text) & "*'" Else findout = 10: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mname Like '" & Trim(findbox.Text) & "*'" Else findout = 11: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mohone Like '" & Trim(findbox.Text) & "*'" Else findout = 12: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "other Like '" & Trim(findbox.Text) & "*'" Else findout = 13: GoTo 56
    findout = 14: GoTo 56
ElseIf chazhaomoshi = "下一个" And Data1.Recordset.EOF Then
MsgBox "这已经是最后一条数据了呢", , "三十一中学生信息管理系统"
Else
Data1.Recordset.FindFirst "name Like '" & Trim(findbox.Text) & "*'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '" & Trim(findbox.Text) & "*'" Else findout = 1: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '" & Trim(findbox.Text) & "*'" Else findout = 6: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '" & Trim(findbox.Text) & "*'" Else findout = 7: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '" & Trim(findbox.Text) & "*'" Else findout = 8: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '" & Trim(findbox.Text) & "*'" Else findout = 9: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '" & Trim(findbox.Text) & "*'" Else findout = 10: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '" & Trim(findbox.Text) & "*'" Else findout = 11: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '" & Trim(findbox.Text) & "*'" Else findout = 12: GoTo 56
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '" & Trim(findbox.Text) & "*'" Else findout = 13: GoTo 56
    findout = 14: GoTo 56
End If
56 If Data1.Recordset.NoMatch Then MsgBox "没有找到开头为“" & Trim(findbox.Text) & "”的内容", vbQuestion, "三十一中学生信息管理系统"
                                                            ElseIf Combo3.Text = "结尾为" Then
If chazhaomoshi = "全部" Then
    Data1.Recordset.FindFirst "name Like '*" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '*" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '*" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '*" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '*" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '*" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '*" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '*" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '*" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '*" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 57
    findout = 14: GoTo 57
ElseIf chazhaomoshi = "下一个" And (Not Data1.Recordset.EOF) Then
    Data1.Recordset.FindNext "name Like '*" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "street Like '*" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "village Like '*" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "area Like '*" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "number Like '*" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fname Like '*" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fphone Like '*" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mname Like '*" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mohone Like '*" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "other Like '*" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 57
    findout = 14: GoTo 57
ElseIf chazhaomoshi = "下一个" And Data1.Recordset.EOF Then
MsgBox "这已经是最后一条数据了呢", , "三十一中学生信息管理系统"
Else
Data1.Recordset.FindFirst "name Like '*" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '*" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '*" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '*" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '*" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '*" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '*" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '*" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '*" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 57
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '*" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 57
    findout = 14: GoTo 57
End If
57 If Data1.Recordset.NoMatch Then MsgBox "没有找到结尾为“" & Trim(findbox.Text) & "”的内容", vbQuestion, "三十一中学生信息管理系统"
ElseIf Combo3.Text = "严格查找" Then
If chazhaomoshi = "全部" Then
    Data1.Recordset.FindFirst "name = '" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street = '" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village = '" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area = '" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number = '" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname = '" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone = '" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname = '" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone = '" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other = '" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 58
    findout = 14: GoTo 58
ElseIf chazhaomoshi = "下一个" And (Not Data1.Recordset.EOF) Then
    Data1.Recordset.FindNext "name = '" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "street = '" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "village = '" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "area = '" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "number = '" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fname = '" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fphone = '" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mname = '" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mohone = '" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "other = '" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 58
    findout = 14: GoTo 58
ElseIf chazhaomoshi = "下一个" And Data1.Recordset.EOF Then
MsgBox "这已经是最后一条数据了呢", , "三十一中学生信息管理系统"
Else
Data1.Recordset.FindFirst "name = '" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street = '" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village = '" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area = '" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number = '" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname = '" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone = '" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname = '" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone = '" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 58
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other = '" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 58
    findout = 14: GoTo 58
End If
58 If Data1.Recordset.NoMatch Then MsgBox "没有找到“" & Trim(findbox.Text) & "”", vbQuestion, "三十一中学生信息管理系统"
ElseIf Combo3.Text = "通配符*" Then
If chazhaomoshi = "全部" Then
    Data1.Recordset.FindFirst "name Like '" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 67
    findout = 14: GoTo 67
ElseIf chazhaomoshi = "下一个" And (Not Data1.Recordset.EOF) Then
    Data1.Recordset.FindNext "name Like '" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "street Like '" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "village Like '" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "area Like '" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "number Like '" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fname Like '" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "fphone Like '" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mname Like '" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "mohone Like '" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindNext "other Like '" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 67
    findout = 14: GoTo 67
ElseIf chazhaomoshi = "下一个" And Data1.Recordset.EOF Then
MsgBox "这已经是最后一条数据了呢", , "三十一中学生信息管理系统"
Else
Data1.Recordset.FindFirst "name Like '" & Trim(findbox.Text) & "'"
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "street Like '" & Trim(findbox.Text) & "'" Else findout = 1: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "village Like '" & Trim(findbox.Text) & "'" Else findout = 6: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "area Like '" & Trim(findbox.Text) & "'" Else findout = 7: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "number Like '" & Trim(findbox.Text) & "'" Else findout = 8: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fname Like '" & Trim(findbox.Text) & "'" Else findout = 9: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "fphone Like '" & Trim(findbox.Text) & "'" Else findout = 10: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mname Like '" & Trim(findbox.Text) & "'" Else findout = 11: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "mohone Like '" & Trim(findbox.Text) & "'" Else findout = 12: GoTo 67
    If Data1.Recordset.NoMatch Then Data1.Recordset.FindFirst "other Like '" & Trim(findbox.Text) & "'" Else findout = 13: GoTo 67
    findout = 14: GoTo 67
    End If
67 If Data1.Recordset.NoMatch Then MsgBox "没有找到“" & Trim(findbox.Text) & "”的内容", vbQuestion, "三十一中学生信息管理系统":: chazhaomoshi = "全部": chazhao = "全部": Combo2.Text = "全部": GoTo 99
End If

On Error GoTo 98
Text(findout).SelStart = InStr(1, Text(findout).Text, Trim(findbox.Text)) - 1
Text(findout).SelLength = Len(Trim(findbox.Text))
Text(findout).SetFocus
Debug.Print InStr(Text(findout).Text, Trim(findbox.Text))
98 If Combo3.Text = "通配符*" Then
Text(findout).SelStart = 0
Text(findout).SelLength = Len(Trim(Text(findout).Text)) + 1
Text(findout).SetFocus
End If
For i = 1 To 查找历史.ListCount
If 查找历史.List(i) = findbox.Text Then repaired = True
Next
If Not repaired Then 查找历史.AddItem findbox.Text
99 End Sub

