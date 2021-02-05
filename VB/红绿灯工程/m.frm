VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Corner"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11415
   ForeColor       =   &H00404040&
   Icon            =   "m.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   11415
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   10560
      Top             =   120
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00808080&
      Caption         =   "ÂÌµÆ"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   465
      Left            =   9360
      TabIndex        =   1
      Top             =   2520
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "ºìµÆ"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   9360
      TabIndex        =   0
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0FFFF&
      BorderColor     =   &H0000FFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   9600
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   10
      Height          =   2175
      Left            =   120
      Top             =   120
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderWidth     =   5
      Height          =   1935
      Left            =   -480
      Shape           =   3  'Circle
      Top             =   240
      Width           =   3495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   10
      X1              =   1200
      X2              =   1200
      Y1              =   2280
      Y2              =   4440
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
For i = 1 To 96
Shape2.Left = Shape2.Left - 100
Next i
Shape2.Left = 9600
End Sub

Private Sub Form_DblClick()
plusb.Show
End Sub

Private Sub Form_Load()
Form1.BackColor = RGB(100, 100, 100)
Option1.BackColor = RGB(100, 100, 100)
Option3.BackColor = RGB(100, 100, 100)
Option3.Value = True
Option1.Value = False
End Sub

Private Sub Option1_Click()
Shape1.BackColor = vbRed
cango = False
End Sub

Private Sub Option3_Click()
Shape1.BackColor = vbGreen
cango = True
End Sub


Private Sub Timer1_Timer()
If cango Then
Shape2.Left = Shape2.Left - 100
If Shape2.Left = -200 Then Shape2.Left = 9600
End If
End Sub
