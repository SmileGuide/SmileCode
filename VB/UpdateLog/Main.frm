VERSION 5.00
Begin VB.Form M 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   10545
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Main.frx":16692
   ScaleHeight     =   5505
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.ListBox List1 
      Height          =   4020
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   9915
   End
End
Attribute VB_Name = "M"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim r As Long
  Dim i
  If Button = 1 Then
    i = ReleaseCapture()
    r = SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
  End If

End Sub
