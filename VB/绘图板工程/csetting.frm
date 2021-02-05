VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form csetting 
   BackColor       =   &H00F8D047&
   Caption         =   "评卷设置"
   ClientHeight    =   3315
   ClientLeft      =   9600
   ClientTop       =   5910
   ClientWidth     =   5010
   Icon            =   "csetting.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3315
   ScaleWidth      =   5010
   Begin VB.CommandButton Command3 
      Height          =   975
      Left            =   3960
      Picture         =   "csetting.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2280
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00F8D047&
      Caption         =   "分数带下划线"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   615
      Left            =   1800
      TabIndex        =   6
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00F8D047&
      Caption         =   "分数字体"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cl 
      Left            =   1800
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   16306247
      FontName        =   "Old English Text MT"
      FontSize        =   72
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00F8D047&
      Caption         =   "分数颜色"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F8D047&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00F8D047&
      Caption         =   "按错题数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00F8D047&
      Caption         =   "减分"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F8D047&
      Caption         =   "满分："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   200
      Width           =   1095
   End
End
Attribute VB_Name = "csetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check1.Value = 1 Then
uned = True
Else
uned = False
End If
End Sub

Private Sub Command1_Click()
cl.ShowColor
End Sub

Private Sub Command2_Click()
cl.ShowFont
End Sub

Private Sub Command3_Click()
Unload csetting
End Sub

Private Sub Form_Load()
If Not Jianfen Then
Option1.Value = True
Text1.Text = FullPoint
Else
Option2.Value = True
Text1.Text = 100
End If
If uned Then
Check1.Value = 1
Else
Check1.Value = 0
End If
End Sub

Private Sub Option1_Click()
Label1.Enabled = True
Text1.Enabled = True
Jianfen = False
Form1.Text3.Visible = True
End Sub

Private Sub Option2_Click()
Label1.Enabled = False
Text1.Enabled = False
Jianfen = True
Form1.Label6.Caption = "-0"
Form1.Text3.Visible = False
FullPoint = 0
fullpointf = 0
End Sub



Private Sub Text1_LostFocus()
FullPoint = Val(Text1.Text)
fullpointf = Val(Text1.Text)
Form1.Label6.Caption = FullPoint
End Sub
