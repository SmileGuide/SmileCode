VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form more 
   BackColor       =   &H00FFFF00&
   Caption         =   "更多"
   ClientHeight    =   4215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5745
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   6.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "more.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   5745
   StartUpPosition =   2  '屏幕中心
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BackColor       =   16776960
      ForeColor       =   16576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "愿望"
      TabPicture(0)   =   "more.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Text1"
      Tab(0).Control(3)=   "Command1"
      Tab(0).Control(4)=   "Text2"
      Tab(0).Control(5)=   "Command2"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "养料请求"
      TabPicture(1)   =   "more.frx":0CE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Text3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Text4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command4"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "驳回"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "给予"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2040
         Width           =   4935
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   4935
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "驳回"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -71520
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2040
         Width           =   4935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "实现"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74160
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74520
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   960
         Width           =   4935
      End
      Begin VB.Label Label4 
         Caption         =   "养料请求克数"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "养料解释"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "愿望"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "愿望解释"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   -74880
         TabIndex        =   2
         Top             =   1680
         Width           =   1575
      End
   End
End
Attribute VB_Name = "more"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
a = 1
Open "F:\BBSeed Files\hopestate.bdf" For Output As #1
Print #1, a
Close #1
MsgBox "已实现！"
more.Hide
End Sub

Private Sub Command2_Click()
a = 2
Open "F:\BBSeed Files\hopestate.bdf" For Output As #1
Print #1, a
Close #1
MsgBox "已驳回！"
more.Hide
End Sub

Private Sub Command3_Click()
a = 1
Open "F:\BBSeed Files\foodstate.bdf" For Output As #1
Print #1, a
Close #1
MsgBox "已给予！"
more.Hide
End Sub

Private Sub Command4_Click()
a = 3
Open "F:\BBSeed Files\foodstate.bdf" For Output As #1
Print #1, a
Close #1
MsgBox "已驳回！"
more.Hide
End Sub

Private Sub Form_Load()
Open "F:\BBSeed Files\hopec.bdf" For Input As #1
Input #1, a
Close #1
Open "F:\BBSeed Files\hopereason.bdf" For Input As #2
Input #2, b
Close #2
Open "F:\BBSeed Files\askn.bdf" For Input As #3
Input #3, c
Close #3
Open "F:\BBSeed Files\askc.bdf" For Input As #4
Input #4, d
Close #4
If a = "" Then
Text1.Text = "无"
Else
Text1.Text = a
End If
If b = "" Then
Text2.Text = "无"
Else
Text2.Text = b
End If
If c = "" Then
Text3.Text = "无"
Else
Text3.Text = c
End If
If d = "" Then
Text4.Text = "无"
Else
Text4.Text = d
End If
End Sub
