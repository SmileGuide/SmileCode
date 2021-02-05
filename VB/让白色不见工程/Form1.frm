VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   5812
   ClientLeft      =   60
   ClientTop       =   404
   ClientWidth     =   9632
   LinkTopic       =   "Form1"
   ScaleHeight     =   5812
   ScaleWidth      =   9632
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4444
      Left            =   1320
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   4444
      ScaleWidth      =   5464
      TabIndex        =   1
      Top             =   360
      Width           =   5464
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   750
      ScaleHeight     =   1080
      ScaleWidth      =   812
      TabIndex        =   0
      Top             =   585
      Width           =   810
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Sub Form_Load()
Dim rtn As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, vbWhite, 0, LWA_COLORKEY
End Sub
