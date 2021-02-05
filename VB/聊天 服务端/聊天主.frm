VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form sever 
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16800
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   16800
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "传文件"
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
      Left            =   0
      TabIndex        =   3
      Top             =   7440
      Width           =   16815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "发送"
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
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   16815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   0
      ScrollBars      =   1  'Horizontal
      TabIndex        =   1
      Top             =   6000
      Width           =   16815
   End
   Begin MSWinsockLib.Winsock W 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   6045
      ItemData        =   "聊天主.frx":0000
      Left            =   0
      List            =   "聊天主.frx":0002
      TabIndex        =   0
      Top             =   0
      Width           =   16815
   End
End
Attribute VB_Name = "sever"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
If W.State = sckConnected Then
W.SendData Text1.Text
Else
MsgBox "对方还没有上线"
End If
End Sub

Private Sub Form_Load()
W.LocalPort = "1000"
W.Bind "1000"
W.Listen
End Sub

Private Sub List1_Click()
For i = 1 To 1000
If filen(i) = "file" Then
If List1.Selected(i + 1) = True Then
GoTo 1
Else
GoTo 20
End If
Next i
1 fn = List1.List(i + 1)
Shell "explorer F:\SChat Files" & fn
20 End Sub

Private Sub W_ConnectionRequest(ByVal requestID As Long)
If W.State <> sckClosed Then
W.Close
List1.AddItem "对方上线了！"
End If
End Sub

Private Sub W_DataArrival(ByVal bytesTotal As Long)
W.GetData getc
List1.AddItem Now & "对方说：" & getc
filen(List1.Index) = "text"
End Sub
