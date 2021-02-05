VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00F8D047&
   Caption         =   "ª≠± ¥Û–°"
   ClientHeight    =   1485
   ClientLeft      =   6330
   ClientTop       =   4335
   ClientWidth     =   4710
   Icon            =   "size.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1485
   ScaleWidth      =   4710
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Slider1.Value = Form1.p.DrawWidth
Label2.Caption = Slider1.Value & dwh
End Sub

Private Sub Slider1_Change()
Label2.Caption = Slider1.Value & dwh
Form2.p.DrawWidth = Slider1.Value
End Sub

