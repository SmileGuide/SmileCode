VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00F8D047&
   Caption         =   "评卷板"
   ClientHeight    =   9303
   ClientLeft      =   1813
   ClientTop       =   2863
   ClientWidth     =   15890
   DrawWidth       =   5
   Icon            =   "m.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9303
   ScaleWidth      =   15890
   Begin VB.Frame Frame1 
      BackColor       =   &H00E9C5E9&
      Caption         =   "画笔、橡皮预览"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1575
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1815
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "文字123ABC"
         Height          =   1215
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   1200
         Picture         =   "m.frx":08CA
         Stretch         =   -1  'True
         ToolTipText     =   "预览框介绍"
         Top             =   240
         Width           =   375
      End
      Begin VB.Line Line2 
         BorderWidth     =   5
         X1              =   240
         X2              =   1080
         Y1              =   360
         Y2              =   720
      End
      Begin VB.Line Line3 
         BorderWidth     =   5
         X1              =   600
         X2              =   1080
         Y1              =   1080
         Y2              =   720
      End
      Begin VB.Line Line4 
         BorderWidth     =   5
         X1              =   1320
         X2              =   600
         Y1              =   1320
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderWidth     =   5
         Visible         =   0   'False
         X1              =   480
         X2              =   840
         Y1              =   360
         Y2              =   960
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   975
         Left            =   120
         Shape           =   3  'Circle
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Timer rf 
      Interval        =   1
      Left            =   14880
      Top             =   8400
   End
   Begin VB.Timer ll 
      Left            =   13320
      Top             =   840
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00F8D047&
      Caption         =   "磅"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   9
      ToolTipText     =   "一磅为1/72英寸,等于0.3527 毫米"
      Top             =   2880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog Cf 
      Left            =   13800
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00F8D047&
      Caption         =   "像素"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   8
      ToolTipText     =   "图像元素，分辨率的尺寸单位，1像素为显示屏所能显示的最小单位"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton 画布大小确定 
      BackColor       =   &H00F8D047&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      Picture         =   "m.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "收起"
      Top             =   7200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F8D047&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E335DF&
      Height          =   855
      Left            =   1800
      TabIndex        =   17
      Top             =   4320
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F8D047&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E335DF&
      Height          =   855
      Left            =   1800
      TabIndex        =   16
      Top             =   5760
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   1800
      Max             =   20000
      Min             =   1
      TabIndex        =   15
      Top             =   5160
      Value           =   1
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   375
      Left            =   1800
      Max             =   10000
      Min             =   1
      TabIndex        =   14
      Top             =   6600
      Value           =   1
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "恢复默认"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "恢复画布大小的原始值"
      Top             =   7080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton 笔大小确定 
      BackColor       =   &H00F8D047&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Picture         =   "m.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "收起"
      Top             =   8040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton 单位确定 
      BackColor       =   &H00F8D047&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Picture         =   "m.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "收起"
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00F8D047&
      Caption         =   "缇"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   10
      ToolTipText     =   "计算机的一种测量单位,1像素=15缇"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H00F8D047&
      Caption         =   "英寸"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "1英寸=2,54厘米"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H00F8D047&
      Caption         =   "毫米"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "10毫米相当于1厘米，100毫米相当于1分米,1000毫米相当于1米"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   5655
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "拖拽我来调整画笔和橡皮的大小"
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   939
      _ExtentY        =   10689
      _Version        =   327682
      BorderStyle     =   1
      Orientation     =   1
      LargeChange     =   1
      Min             =   1
      Max             =   30
      SelStart        =   5
      Value           =   5
   End
   Begin MSComDlg.CommonDialog CBC 
      Left            =   0
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   0
      Top             =   4080
   End
   Begin MSComDlg.CommonDialog CommonDialog3 
      Left            =   0
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "绘画作品.bmp"
      Filter          =   "JPEG(*.jpg)|*.jpg|位图(*.bmp)|*.bmp"
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   0
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "JPEG(*.jpg)|*.jpg|位图(*.bmp)|*.bmp"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "绘画作品.bmp"
      Filter          =   "JPEG(*.jpg)|*.jpg|位图(*.bmp)|*.bmp"
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   497
      Left            =   0
      TabIndex        =   0
      Top             =   8806
      Width           =   15890
      _ExtentX        =   28033
      _ExtentY        =   873
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0097FBCB&
      Caption         =   "快捷评卷操作"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   7935
      Left            =   14040
      TabIndex        =   23
      Top             =   0
      Width           =   1815
      Begin VB.CommandButton UNLA 
         BackColor       =   &H0097FBCB&
         Caption         =   "关闭副窗口"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "此功能可以弹出新的编辑窗口，两窗口间数据互相连通"
         Top             =   7200
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H0097FBCB&
         Caption         =   "打开副窗口"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "此功能可以弹出新的编辑窗口，两窗口间数据互相连通"
         Top             =   6600
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0097FBCB&
         Caption         =   "归位"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3480
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0097FBCB&
         Caption         =   "打分"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Height          =   495
         Left            =   120
         Picture         =   "m.frx":2BF2
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "设置"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton JP 
         BackColor       =   &H0097FBCB&
         Caption         =   "减分"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   5520
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0097FBCB&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   4920
         Width           =   1575
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H0097FBCB&
         Caption         =   "红圈"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H0097FBCB&
         Caption         =   "涂鸦红笔"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   1575
      End
      Begin VB.Line Line7 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         X1              =   120
         X2              =   1680
         Y1              =   6360
         Y2              =   6360
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "错题："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Line Line6 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         X1              =   120
         X2              =   1680
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "总分："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Line Line5 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         X1              =   120
         X2              =   1680
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0097FBCB&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   29
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Image Image3 
         Height          =   375
         Left            =   1200
         Picture         =   "m.frx":3140
         Stretch         =   -1  'True
         ToolTipText     =   "评卷框介绍"
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.PictureBox p 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   5
      Height          =   7935
      Left            =   2160
      ScaleHeight     =   7910
      ScaleWidth      =   11865
      TabIndex        =   1
      Top             =   0
      Width           =   11895
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7935
      LargeChange     =   100
      Left            =   1800
      Max             =   7935
      TabIndex        =   22
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008FC4F3&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E335DF&
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   7320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image dww 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   1440
      Picture         =   "m.frx":3A0A
      Stretch         =   -1  'True
      ToolTipText     =   "帮助"
      Top             =   8040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image penw 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   960
      Picture         =   "m.frx":CACC
      Stretch         =   -1  'True
      ToolTipText     =   "帮助"
      Top             =   8040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image backw 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   480
      Picture         =   "m.frx":15B8E
      Stretch         =   -1  'True
      ToolTipText     =   "帮助"
      Top             =   8040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00F8D047&
      Caption         =   "画布长"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   5040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00F8D047&
      Caption         =   "画布宽"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "左右键拖动画布外蓝色区域试试"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   8280
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   0
      Top             =   7320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Menu f 
      Caption         =   "文件"
      Begin VB.Menu cls 
         Caption         =   "新建"
         Shortcut        =   ^N
      End
      Begin VB.Menu uu 
         Caption         =   "-"
      End
      Begin VB.Menu stc 
         Caption         =   "保存到剪贴板"
         Shortcut        =   ^V
      End
      Begin VB.Menu sfc 
         Caption         =   "从剪贴板导入"
         Shortcut        =   ^D
      End
      Begin VB.Menu gg 
         Caption         =   "-"
      End
      Begin VB.Menu Save 
         Caption         =   "保存"
      End
      Begin VB.Menu SV 
         Caption         =   "另存为>"
         Shortcut        =   ^S
      End
      Begin VB.Menu open 
         Caption         =   "打开>"
         Shortcut        =   ^O
      End
      Begin VB.Menu hgh 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "退出"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu bk 
      Caption         =   "画布"
      Begin VB.Menu bna 
         Caption         =   "画布颜色"
         Shortcut        =   {F1}
      End
      Begin VB.Menu bs 
         Caption         =   "画布大小"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu dae 
      Caption         =   "画笔及橡皮"
      Begin VB.Menu menu 
         Caption         =   "画笔及橡皮样式"
         Begin VB.Menu a 
            Caption         =   "直线"
            Shortcut        =   ^L
         End
         Begin VB.Menu b 
            Caption         =   "曲线"
            Shortcut        =   ^C
         End
         Begin VB.Menu sq 
            Caption         =   "矩形"
            Begin VB.Menu sqs 
               Caption         =   "实心"
               Shortcut        =   ^U
            End
            Begin VB.Menu sqe 
               Caption         =   "空心"
               Shortcut        =   ^B
            End
         End
         Begin VB.Menu ci 
            Caption         =   "圆"
            Begin VB.Menu cc1 
               Caption         =   "实心"
               Shortcut        =   ^H
            End
            Begin VB.Menu cc2 
               Caption         =   "空心"
               Shortcut        =   ^R
            End
         End
         Begin VB.Menu wd 
            Caption         =   "文字"
            Shortcut        =   ^W
         End
      End
      Begin VB.Menu color 
         Caption         =   "画笔颜色"
         Shortcut        =   ^P
      End
      Begin VB.Menu c1 
         Caption         =   "-"
      End
      Begin VB.Menu era 
         Caption         =   "橡皮"
         Begin VB.Menu e 
            Caption         =   "橡皮"
            Shortcut        =   ^X
         End
         Begin VB.Menu c 
            Caption         =   "清空"
            Shortcut        =   {DEL}
         End
      End
      Begin VB.Menu c2 
         Caption         =   "-"
      End
      Begin VB.Menu size 
         Caption         =   "画笔及橡皮大小"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu dw 
      Caption         =   "坐标单位"
   End
   Begin VB.Menu mre 
      Caption         =   "更多..."
      Begin VB.Menu dd 
         Caption         =   "勿扰模式"
         Checked         =   -1  'True
      End
      Begin VB.Menu bee 
         Caption         =   "弹出提示音"
         Checked         =   -1  'True
      End
      Begin VB.Menu wh 
         Caption         =   "什么是“勿扰模式”？"
      End
      Begin VB.Menu jh 
         Caption         =   "-"
      End
      Begin VB.Menu hh 
         Caption         =   "快捷键列表"
         Shortcut        =   ^T
      End
      Begin VB.Menu u 
         Caption         =   "-"
      End
      Begin VB.Menu ab 
         Caption         =   "关于..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub a_Click()
ob = "直线"
isp = False
End Sub

Private Sub ab_Click()
ABMe.Show
End Sub

Private Sub b_Click()
ob = "曲线"
End Sub

Private Sub backw_Click()
MsgBox "通过调整画布大小来改变范围"
End Sub

Private Sub bee_Click()
canbeep = Not canbeep
bee.Checked = Not bee.Checked
End Sub

Private Sub bna_Click()
On Error GoTo 23
CBC.ShowColor
pbbkcolor = CBC.color
p.BackColor = CBC.color
23 End Sub

Private Sub bs_Click()
Frame1.Left = 0
Frame1.Top = 0
dww.Visible = False
penw.Visible = False
backw.Visible = True
Label2.Visible = False
Slider1.Visible = False
笔大小确定.Visible = False
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option5.Visible = False
Option6.Visible = False
单位确定.Visible = False
Label4.Visible = True
Label3.Visible = True
Text1.Visible = True
Text2.Visible = True
HScroll1.Visible = True
HScroll2.Visible = True
Command1.Visible = True
画布大小确定.Visible = True
End Sub

Private Sub c_Click()
If canbeep Then Beep
If MsgBox("是否清空画布？", vbOKCancel, "清空") = vbOK Then
p.cls
p.BackColor = CBC.color
End If
End Sub

Private Sub cc1_Click()
ob = "实心圆"
End Sub

Private Sub cc2_Click()
ob = "空心圆"
End Sub

Private Sub che_Click()
pbcolor = vbRed
CommonDialog1.color = vbRed
p.DrawWidth = 10
Frame2.Visible = True
Option4.Value = True
End Sub

Private Sub cls_Click()
If Not saved Then
If canbeep Then Beep
If MsgBox("是否保存图片的更改?", vbOKCancel, "绘图板") = vbOK Then
p.cls
Else
p.cls
GoTo 1
End If
End If
1 End Sub

Private Sub color_Click()
CommonDialog1.ShowColor
pbcolor = CommonDialog1.color
End Sub



Private Sub Command2_Click()
csetting.Show
End Sub

Private Sub Command3_Click()
p.CurrentX = 10
p.CurrentY = 10
p.ForeColor = csetting.cl.color
pbcolor = csetting.cl.color
With p.Font
    .Bold = csetting.cl.FontBold
    .Italic = csetting.cl.FontItalic
    .Name = csetting.cl.FontName
    .size = csetting.cl.FontSize
    .Strikethrough = csetting.cl.FontStrikethru
    .Underline = uned
End With
p.Print Label6.Caption
p.Font.Underline = False
End Sub

Private Sub Command4_Click()
If Not Jianfen Then
FullPoint = fullpointf
Label6.Caption = FullPoint
Else
FullPoint = 0
Label6.Caption = "-0"
End If
End Sub

Private Sub Command6_Click()
forder = forder + 1
fordern = fordern + 1
If forder >= 100 Then
MsgBox "你创建的新窗口太多了，请在关闭窗口界面关闭所有窗口后重新打开副窗口"
GoTo 88
End If
Set sform(forder) = New Form1
second = True
sform(forder).Caption = "评卷板(副窗口" & forder & ")"
sform(forder).Show
distan = distan + 300
sform(forder).Top = sform(forder).Top + distan
sform(forder).Left = sform(forder).Left + distan
UNLA.Enabled = True
88 End Sub

Private Sub dd_Click()
dd.Checked = Not dd.Checked
dised = Not dised
End Sub

Private Sub dw_Click()
Frame1.Left = 0
Frame1.Top = 0
dww.Visible = True
penw.Visible = False
backw.Visible = False
Label4.Visible = False
Label3.Visible = False
Text1.Visible = False
Text2.Visible = False
Command1.Visible = False
画布大小确定.Visible = False
HScroll1.Visible = False
HScroll2.Visible = False
Label2.Visible = False
Slider1.Visible = False
笔大小确定.Visible = False
Option1.Visible = True
Option2.Visible = True
Option3.Visible = True
Option5.Visible = True
Option6.Visible = True
单位确定.Visible = True

End Sub

Private Sub dww_Click()
MsgBox "通过改变选中的选项按钮来调整画布的坐标的单位，对画笔不适用(画笔的单位为像素)"
End Sub

Private Sub e_Click()
CommonDialog1.color = CBC.color
pbcolor = CommonDialog1.color
End Sub


Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_DblClick()
With Frame1
.Top = 0
.Left = 0
End With
With Frame2
.Top = 0
.Left = 14040
End With
End Sub

Private Sub Form_Load()
Jianfen = False
If Not second Then
FullPoint = 100
fullpointf = 100
uned = True
Else
End If
Label6.Caption = FullPoint
resizetime = Now
ob = "曲线"
isp = False
dwh = "缇"
CBC.color = vbWhite
pbcolor = vbWhite
Text1.Text = Form1.p.Width
Text2.Text = Form1.p.Height
WordNum = 1
dised = False
dd.Checked = False
bee.Checked = False
VScroll1.Max = p.Width - 7935
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Frame1.Left = X
Frame1.Top = Y
fcand = True
Else
Frame2.Left = X
Frame2.Top = Y
fcand = True
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If fcand Then
If Button = 1 Then
Frame1.Left = X
Frame1.Top = Y
Timer1.Interval = 0
Label1.Visible = False
Else
Frame2.Left = X
Frame2.Top = Y
Timer1.Interval = 0
Label1.Visible = False
End If
End If

End Sub



Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
fcand = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
If Not closed Then
If Not saved Then
If canbeep Then Beep
If MsgBox("是否保存图片的更改?", vbOKCancel, "绘图板") = vbOK Then
CommonDialog1.FileName = Format(Now, "yyyy年mm月dd日hh时mm分ss秒") & "的绘画作品"
CommonDialog1.ShowSave
SavePicture p.Image, CommonDialog1.FileName
End If
End If
p.cls
Unload Me
End If
fordern = fordern - 1
End Sub

Private Sub fwef_Click()
Dim ddd() As Long
p.ScaleMode = 3
p.AutoRedraw = True
w = p.ScaleWidth
h = p.ScaleHeight
ReDim ddd(w - 1, h - 1)
For i = 0 To w - 1
For j = 0 To h - 1
ddd(i, j) = p.Point(i, j)
Next j
Next i
p.Width = h + 4 * p.BorderStyle
p.Height = w + 4 * p.BorderStyle
For i = 0 To h - 1
For j = 0 To w - 1
p.PSet (i, h - j - 1), ddd(j, i)
Next j
Next i
End Sub

Private Sub hh_Click()
help.Show
End Sub

Private Sub Image2_Click()
MsgBox "    您可以在此处预览画笔和橡皮的效果，拖动画布外的蓝色区域，此框架就会移动到鼠标指针的位置。" & Chr(10) & "    鼠标左键拖动画布外的蓝色区域即可让框架重新定位，双击窗口下方的状态栏即可还原。"
End Sub




Private Sub Image3_Click()
MsgBox "    此模式更适用于老师、学生，含有快捷的红色笔、标准的椭圆圈形笔以及实用的分数统计工具。“打开副窗口”按钮支持弹出若干个空白画布副窗口来评多页的试卷，按下“关闭副窗口”按钮即可选择性地关闭打开的副窗口。" & Chr(10) & "    鼠标右键拖动画布外的蓝色区域即可让框架重新定位，双击窗口下方的状态栏即可还原。"
End Sub

Private Sub JP_Click()
If Not Jianfen Then
FullPoint = FullPoint - Val(Text3.Text)
Label6.Caption = FullPoint
Else
FullPoint = FullPoint + 1
Form1.Label6.Caption = "-" & FullPoint
End If
End Sub





Private Sub ll_Timer()
colour = colour + 1
If Int(colour / 2) = colour / 2 Then Form1.BackColor = first Else Form1.BackColor = RGB(255, colour * 10, 0)
If colour = 24 Then ll.Interval = 0: colour = 1: Frame1.Visible = True
End Sub

Private Sub open_Click()
If Not saved Then
If canbeep Then Beep
If MsgBox("是否保存图片的更改?", vbOKCancel, "绘图板") = vbOK Then
CommonDialog1.FileName = Format(Now, "yyyy年mm月dd日hh时mm分ss秒") & "的绘画作品.bmp"
CommonDialog1.ShowSave
SavePicture p.Image, CommonDialog1.FileName
Else
p.cls
End If
End If
CommonDialog2.ShowOpen
Image1.Picture = LoadPicture(CommonDialog2.FileName)
p.Height = p.Width / Image1.Width * Image1.Height
On Error GoTo 1
p.PaintPicture Image1.Picture, 0, 0, p.Width, p.Height
GoTo 33
1 If Not dised Then MsgBox "你还没有选择文件"
33 End Sub

Private Sub Option4_Click()
CommonDialog1.color = vbRed
pbcolor = vbRed
pbdw = 10
ob = "曲线"
End Sub

Private Sub Option7_Click()
ob = "红圈"
CommonDialog1.color = vbRed
pbcolor = vbRed
End Sub

Private Sub p_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lined = True
x1 = X
y1 = Y
isp = True
canWrite = True
slcirs1 = Y
cancir = True
End Sub

Private Sub p_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.SimpleText = X & "," & Y & "  " & dwh
Slider1.Value = Form1.p.DrawWidth
If ob = "直线" And lined = True Then
StatusBar1.SimpleText = "正在画直线  " & StatusBar1.SimpleText
End If
If ob = "实心矩形" And lined = True Then
StatusBar1.SimpleText = "正在画实心矩形  " & StatusBar1.SimpleText
ElseIf ob = "空心矩形" And lined = True Then
StatusBar1.SimpleText = "正在画空心矩形  " & StatusBar1.SimpleText
End If
If ob = "文字" Then StatusBar1.SimpleText = "正在写文字" & StatusBar1.SimpleText
If ob = "曲线" Then
If isp Then
StatusBar1.SimpleText = "正在画曲线  " & StatusBar1.SimpleText
p.PSet (X, Y)
End If
End If
If ob = "红圈" And cancir Then StatusBar1.SimpleText = "快捷批卷操作：正在画红圈" & StatusBar1.SimpleText
End Sub

Private Sub p_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If cancir And ob = "红圈" Then
p.Circle (x1, y1), 1000, vbRed, , , 0.3
cancir = False
End If
lined = False
If ob = "直线" Then
p.Line (x1, y1)-(X, Y)
    ElseIf ob = "实心矩形" Then
    p.Line (x1, y1)-(X, Y), , BF
        ElseIf ob = "空心矩形" Then
        p.Line (x1, y1)-(X, Y), , B
            ElseIf ob = "实心圆" Then
            mm = p.DrawWidth
            Debug.Print Y - slcirs1
            Debug.Print slcirs1
            If Y - slcirs1 > 0 Then
            pbdw = (Y - slcirs1) / 15
            p.DrawWidth = pbdw
            p.PSet (X, Y)
                Else
                    If Not dised Then MsgBox "请向下拖拽"
                End If
                p.DrawWidth = mm
                    ElseIf ob = "空心圆" Then
                    If Y - slcirs1 > 0 Then
                    p.Circle (X, Y), Y - slcirs1
                     Else
                If Not dised Then MsgBox "请向下拖拽"
                End If
                End If
isp = False
canWrite = False
If ob = "文字" Then
StatusBar1.SimpleText = X & Y & dwh
End If
End Sub



Private Sub penw_Click()
MsgBox "      通过调整刻度尺上的指针位置来改变画笔的大小，作用于：" & Chr(10) & "直线、曲线、空心矩形、空心圆。"
End Sub

Private Sub rf_Timer()
If fordern = 0 Then second = False
If second Then UNLA.Enabled = True Else UNLA.Enabled = False
p.ForeColor = pbcolor
VScroll1.Max = p.Width - 7935
VScroll1.Value = -p.Top
Label2.Caption = Slider1.Value
Shape1.BackColor = pbcolor
Shape1.BorderColor = pbcolor
Line1.BorderColor = pbcolor
Line1.BorderWidth = Slider1.Value
Line2.BorderColor = pbcolor
Line2.BorderWidth = Slider1.Value
Line3.BorderColor = pbcolor
Line3.BorderWidth = Slider1.Value
Line4.BorderColor = pbcolor
Line4.BorderWidth = pbdw
Line1.BorderWidth = pbdw
Line2.BorderWidth = pbdw
Line3.BorderWidth = pbdw
Line4.BorderWidth = pbdw
If ob = "直线" Then
Label5.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
Shape1.Visible = False
Line1.BorderWidth = p.DrawWidth
Line1.Visible = True
ElseIf ob = "曲线" Then
Label5.Visible = False
Shape1.Visible = False
Line1.Visible = False
Line2.Visible = True
Line3.Visible = True
Line4.Visible = True
ElseIf ob = "红圈" Then
Label5.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
Shape1.Shape = 2
Shape1.BackStyle = 0
Shape1.BorderWidth = p.DrawWidth
Line1.Visible = False
Shape1.Visible = True
ElseIf ob = "实心矩形" Then
Label5.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
Shape1.Shape = 1
Shape1.BackStyle = 1
Line1.Visible = False
Shape1.Visible = True
ElseIf ob = "空心矩形" Then
Label5.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
Shape1.Shape = 1
Shape1.BackStyle = 0
Shape1.BorderWidth = p.DrawWidth
Line1.Visible = False
Shape1.Visible = True
ElseIf ob = "文字" Then
Label5.ForeColor = CommonDialog1.color
Label5.Visible = True
Line1.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
Shape1.Visible = False
With Label5.Font
    .Bold = Cf.FontBold
    .Italic = Cf.FontItalic
    .Name = Cf.FontName
    .size = Cf.FontSize
    .Strikethrough = Cf.FontStrikethru
    .Underline = Cf.FontUnderline
End With
If canWrite Then
p.CurrentX = X
p.CurrentY = Y
con = InputBox("请输入插入文字的内容")
p.Print con
If InStr(con, "咕噜") > 0 Then Call love
canWrite = False
ob = "曲线"
isp = False
End If
ElseIf ob = "实心圆" Then
StatusBar1.SimpleText = "正在画实心圆  " & StatusBar1.SimpleText
Label5.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
Shape1.Shape = 3
Shape1.BackStyle = 1
Line1.Visible = False
Shape1.Visible = True
ElseIf ob = "空心圆" Then
StatusBar1.SimpleText = "正在画空心圆  " & StatusBar1.SimpleText
Label5.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
Shape1.Shape = 3
Shape1.BackStyle = 0
Shape1.BorderWidth = p.DrawWidth
Line1.Visible = False
Shape1.Visible = True
End If
End Sub

Private Sub Save_Click()
On Error GoTo 55
SavePicture p.Picture, CommonDialog2.FileName
Exit Sub
55 CommonDialog3.FileName = Format(Now, "yyyy年mm月dd日hh时mm分ss秒") & "的绘画作品.jpg"
CommonDialog3.ShowSave
SavePicture p.Image, CommonDialog3.FileName
saved = True
nnn = nnn + 1
End Sub

Private Sub sfc_Click()
FullPoint = fullpointf
Label6.Caption = FullPoint
If Not saved Then
If canbeep Then Beep
If MsgBox("是否保存图片的更改?", vbOKCancel, "绘图板") = vbOK Then
CommonDialog1.FileName = Format(Now, "yyyy年mm月dd日hh时mm分ss秒") & "的绘画作品.bmp"
CommonDialog1.ShowSave
SavePicture p.Image, CommonDialog1.FileName
Else
p.cls
End If
End If
Image1.Picture = Clipboard.GetData
p.Height = p.Width / Image1.Width * Image1.Height
On Error GoTo 33
p.PaintPicture Image1.Picture, 0, 0, p.Width, p.Height
GoTo 44
33 If Not dised Then MsgBox "你没有复制图片或图片为路径的复制（使用打开即可）"
44 End Sub

Private Sub size_Click()
Frame1.Left = 0
Frame1.Top = 0
dww.Visible = False
penw.Visible = True
backw.Visible = False
Label4.Visible = False
Label3.Visible = False
Text1.Visible = False
Text2.Visible = False
Command1.Visible = False
画布大小确定.Visible = False
HScroll1.Visible = False
HScroll2.Visible = False
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option5.Visible = False
Option6.Visible = False
单位确定.Visible = False
Label2.Visible = True
Slider1.Visible = True
笔大小确定.Visible = True
End Sub
'长方形绘制
'打开图片，保存或另存为
'删除delate
'设置背景色
'修复再次打开画笔粗细还为5


Private Sub sqe_Click()
ob = "空心矩形"
isp = False
End Sub

Private Sub sqs_Click()
ob = "实心矩形"
isp = False
End Sub



Private Sub StatusBar1_DBlClick()
With Frame1
.Top = 0
.Left = 0
End With
With Frame2
.Top = 0
.Left = 14040
End With
End Sub






Private Sub stc_Click()
If Not dised Then MsgBox "保存成功！"
Clipboard.Clear
Clipboard.SetData p.Image
saved = True
nnn = nnn + 1
End Sub

Private Sub sv_Click()
CommonDialog3.FileName = Format(Now, "yyyy年mm月dd日hh时mm分ss秒") & "的绘画作品"
CommonDialog3.ShowSave
On Error GoTo 22
SavePicture p.Image, CommonDialog3.FileName
saved = True
nnn = nnn + 1
GoTo 88
22 MsgBox "抱歉，错误了(^_-)☆"
88 End Sub

Private Sub Text3_Change()
If Val(Text3.Text) > 100 Then Text3.Text = 100
End Sub

Private Sub Text3_GotFocus()
noword = Text3.Text
Text3.Text = ""
Text3.BackColor = &HE9C5E9
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
JP.SetFocus
If Not Jianfen Then
FullPoint = FullPoint - Val(Text3.Text)
Label6.Caption = FullPoint
Else
FullPoint = FullPoint + 1
Form1.Label6.Caption = "-" & FullPoint
End If
End If
End Sub

Private Sub Text3_LostFocus()
Text3.BackColor = &H97FBCB
If Text3.Text = "" Then
Text3.Text = noword
Else
Text3.Text = Val(Text3.Text)
End If
Changed = False
End Sub

Private Sub Timer1_Timer()
Label1.Visible = Not Label1.Visible
End Sub



Private Sub UNLA_Click()
uun.Show
End Sub

Private Sub VScroll1_Change()
p.Top = -VScroll1.Value
End Sub

Private Sub wd_Click()

Cf.ShowFont
With p.Font
    .Bold = Cf.FontBold
    .Italic = Cf.FontItalic
    .Name = Cf.FontName
    .size = Cf.FontSize
    .Strikethrough = Cf.FontStrikethru
    .Underline = Cf.FontUnderline
End With
ob = "文字"
End Sub



Private Sub wh_Click()
MsgBox "使用程序操作不当引发错误时没有提示并没有检测，“勿扰模式”可以让您更好地投身于创作。此模式更适用于能够熟练操作本程序的用户。"
End Sub

Private Sub xxm_Click()
xxm.Checked = Not xxm.Checked
If xxm.Checked = True Then
plusb.Show
Else
Unload plusb
End If
End Sub
Private Sub Slider1_Change()
Label2.Caption = Slider1.Value
pbdw = Slider1.Value
End Sub

Private Sub 笔大小确定_Click()
penw.Visible = False
Label2.Visible = False
Slider1.Visible = False
笔大小确定.Visible = False
End Sub
Private Sub Option1_Click()
Form1.ScaleMode = 1
Form1.p.ScaleMode = 1
dwh = "缇"
End Sub

Private Sub Option2_Click()
Form1.ScaleMode = 2
Form1.p.ScaleMode = 2
dwh = "磅"
End Sub

Private Sub Option3_Click()
Form1.ScaleMode = 3
Form1.p.ScaleMode = 3

dwh = "像素"
End Sub


Private Sub Option5_Click()
Form1.ScaleMode = 5
Form1.p.ScaleMode = 5
dwh = "英寸"
End Sub

Private Sub Option6_Click()
Form1.ScaleMode = 6
Form1.p.ScaleMode = 6
dwh = "毫米"
End Sub


Private Sub Command1_Click()
Form1.p.Width = 11895
Form1.p.Height = 7935
HScroll1.Value = 11895
HScroll2.Value = 7935
End Sub

Private Sub hScroll1_Change()
Text1.Text = HScroll1.Value
End Sub

Private Sub hScroll2_Change()
Text2.Text = HScroll2.Value
End Sub



Private Sub Text1_Change()
HScroll1.Value = Val(Text1.Text)
Form1.p.Width = HScroll1.Value
1 End Sub

Private Sub Text2_Change()
HScroll2.Max = Val(Text2.Text) + 2000
HScroll2.Value = Val(Text2.Text)
Form1.p.Height = HScroll2.Value
1 End Sub

Private Sub 单位确定_Click()
dww.Visible = False
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option5.Visible = False
Option6.Visible = False
单位确定.Visible = False
End Sub

Private Sub 画布大小确定_Click()
backw.Visible = False
Label4.Visible = False
Label3.Visible = False
Text1.Visible = False
Text2.Visible = False
Command1.Visible = False
画布大小确定.Visible = False
HScroll1.Visible = False
HScroll2.Visible = False
End Sub
