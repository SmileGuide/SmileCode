VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form WBoard 
   BackColor       =   &H00FFFFFF&
   Caption         =   "°×°å"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9030
   Icon            =   "WBoard.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5505
   ScaleWidth      =   9030
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin MSComDlg.CommonDialog C 
      Left            =   3090
      Top             =   675
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Color           =   16777215
   End
   Begin VB.Menu m1 
      Caption         =   "m1"
      Visible         =   0   'False
      Begin VB.Menu f1 
         Caption         =   "±³¾°ÑÕÉ«"
      End
   End
End
Attribute VB_Name = "WBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub f1_Click()
On Error GoTo e
C.ShowColor
WBoard.BackColor = C.Color
e: End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu m1
End Sub
