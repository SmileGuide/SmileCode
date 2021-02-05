VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   11940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17535
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11940
   ScaleWidth      =   17535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Image1.Picture = LoadPicture("C:\Users\user\Desktop\2.jpg")
End Sub

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Unload Me
End Sub


