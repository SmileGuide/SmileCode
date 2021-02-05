VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BackColor       =   &H0097FBCB&
   Caption         =   "Form1"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20325
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleWidth      =   20325
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   16200
      TabIndex        =   1
      Top             =   3240
      Width           =   2295
      Begin VB.CommandButton Command1 
         BackColor       =   &H00F8D047&
         Caption         =   "百度地图"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "百度一下"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00F8D047&
         Caption         =   "百度知道"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "CSDN"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2880
         Width           =   1575
      End
   End
   Begin SHDocVwCtl.WebBrowser W 
      Height          =   9615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15855
      ExtentX         =   27966
      ExtentY         =   16960
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
W.Refresh
W.Navigate "map.baidu.com"
End Sub

Private Sub Command2_Click()

W.Refresh
W.Navigate "baidu.com"
End Sub

Private Sub Command3_Click()

W.Refresh
W.Navigate "zhidao.baidu.com"
End Sub

Private Sub Command4_Click()

W.Refresh
W.Navigate "csdn.net"
End Sub

Private Sub Form_Initialize()

W.Width = Form1.Width
W.Height = Form1.Height
End Sub

Private Sub Form_Load()

W.Navigate ""
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CanDrag = True
Frame1.Top = Y
Frame1.Left = X
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

W.Refresh
If CanDrag Then
Frame1.Top = Y
Frame1.Left = X
End If
End Sub



Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CanDrag = False
End Sub

Private Sub W_NewWindow2(ppDisp As Object, Cancel As Boolean)
Dim frm As Form1
Set frm = New Form1
frm.Visible = True
Set ppDisp = W.object
End Sub


