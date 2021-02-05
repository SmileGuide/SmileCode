VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form file 
   Caption         =   "Form1"
   ClientHeight    =   2850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   5430
   StartUpPosition =   3  '窗口缺省
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1680
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   0
      MouseIcon       =   "file.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "上传"
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
      Left            =   840
      TabIndex        =   3
      Top             =   2040
      Width           =   3855
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.DirListBox Dir1 
      Height          =   1560
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      Height          =   1530
      Left            =   3240
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "file"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
File1.FileName = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

