VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00F8D047&
   Caption         =   "µ¥Î»"
   ClientHeight    =   2925
   ClientLeft      =   6330
   ClientTop       =   4500
   ClientWidth     =   2700
   Icon            =   "dw.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   2925
   ScaleWidth      =   2700
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Option1_Click()
Form1.ScaleMode = 1
Form1.p.ScaleMode = 1
dwh = "ç¾"
End Sub

Private Sub Option2_Click()
Form1.ScaleMode = 2
Form1.p.ScaleMode = 2
dwh = "°õ"
End Sub

Private Sub Option3_Click()
Form1.ScaleMode = 3
Form1.p.ScaleMode = 3
dwh = "ÏñËØ"
End Sub

Private Sub Option4_Click()
Form1.ScaleMode = 4
Form1.p.ScaleMode = 4
dwh = "×Ö·û"
End Sub

Private Sub Option5_Click()
Form1.ScaleMode = 5
Form1.p.ScaleMode = 5
dwh = "Ó¢´ç"
End Sub

Private Sub Option6_Click()
Form1.ScaleMode = 6
Form1.p.ScaleMode = 6
dwh = "ºÁÃ×"
End Sub

