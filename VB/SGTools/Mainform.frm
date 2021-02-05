VERSION 5.00
Begin VB.MDIForm Mainform 
   BackColor       =   &H00C0FFC0&
   Caption         =   "SGToola"
   ClientHeight    =   6765
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   14070
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu F1 
      Caption         =   "打开"
      Begin VB.Menu S1 
         Caption         =   "向导"
      End
      Begin VB.Menu S2 
         Caption         =   "窗口管理器"
      End
   End
End
Attribute VB_Name = "Mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub S1_Click()
Startfrom.Show
End Sub

Private Sub S2_Click()
FMN.Show
End Sub
