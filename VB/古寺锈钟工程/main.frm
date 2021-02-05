VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H008FC4F3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÆÆÃÅÍÇÇ½"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5355
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   5355
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H008FC4F3&
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   1560
      Picture         =   "main.frx":1084A
      ScaleHeight     =   1920
      ScaleWidth      =   1920
      TabIndex        =   1
      Top             =   4200
      Width           =   1920
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   1335
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   2880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H008FC4F3&
      Caption         =   "ÔÚ´°ÌåÄÚ»­ÏßÊÔÊÔ"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   855
      Left            =   960
      TabIndex        =   2
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ÇÃÖÓ°ô"
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_DblClick()
plusb.Show
End Sub

Private Sub FORM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
a = True
End Sub

Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If a Then
Label1.Top = Y
Label1.Left = X
Debug.Print X, Y
If X > 1560 And Y > 2760 And X < 3240 And Y < 6960 Then
Shape1.Visible = True
Shape1.Top = 1000
Call r
Shape1.Visible = False
End If
End If
End Sub

Private Sub FORM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
a = False
End Sub


