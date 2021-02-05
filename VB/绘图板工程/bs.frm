VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00F8D047&
   Caption         =   "画布大小"
   ClientHeight    =   4680
   ClientLeft      =   6330
   ClientTop       =   4500
   ClientWidth     =   6150
   Icon            =   "bs.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   4680
   ScaleWidth      =   6150
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   2055
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   375
      Left            =   1680
      Max             =   10000
      TabIndex        =   5
      Top             =   2520
      Width           =   3495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   1680
      Max             =   20000
      TabIndex        =   4
      Top             =   1080
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
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   3495
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
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label2 
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
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Form_Load()
Text1.Text = Form1.p.Width
Text2.Text = Form1.p.Height
End Sub

Private Sub Text1_Change()
If Val(Text1.Text) > 20000 Then
MsgBox "内存溢出"
Text2.Text = HScroll2.Value
GoTo 1
End If
HScroll1.Value = Val(Text1.Text)
Form1.p.Width = HScroll1.Value
1 End Sub

Private Sub Text2_Change()
If Val(Text2.Text) > 10000 Then
MsgBox "内存溢出"
Text2.Text = HScroll2.Value
GoTo 1
End If
HScroll2.Value = Val(Text2.Text)
Form1.p.Height = HScroll2.Value
1 End Sub
