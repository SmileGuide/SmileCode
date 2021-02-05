VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CLOrder 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "颜色序列"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   3015
   Icon            =   "CLOrder.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OK 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   240
      TabIndex        =   6
      Top             =   6330
      Width           =   2310
   End
   Begin MSComDlg.CommonDialog cdc 
      Left            =   75
      Top             =   1950
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cho 
      Caption         =   "选色"
      Height          =   600
      Left            =   15
      TabIndex        =   5
      Top             =   45
      Width           =   570
   End
   Begin VB.OptionButton Optio 
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   4
      Left            =   630
      TabIndex        =   4
      Top             =   5250
      Width           =   300
   End
   Begin VB.OptionButton Optio 
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   3
      Left            =   660
      TabIndex        =   3
      Top             =   4020
      Width           =   300
   End
   Begin VB.OptionButton Optio 
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   2
      Left            =   660
      TabIndex        =   2
      Top             =   2790
      Width           =   300
   End
   Begin VB.OptionButton Optio 
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   1
      Left            =   675
      TabIndex        =   1
      Top             =   1395
      Width           =   300
   End
   Begin VB.OptionButton Optio 
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   0
      Left            =   660
      TabIndex        =   0
      Top             =   60
      Width           =   300
   End
   Begin VB.Shape Shape 
      FillStyle       =   0  'Solid
      Height          =   1005
      Index           =   4
      Left            =   990
      Shape           =   3  'Circle
      Top             =   5220
      Width           =   990
   End
   Begin VB.Shape Shape 
      FillStyle       =   0  'Solid
      Height          =   1005
      Index           =   3
      Left            =   990
      Shape           =   3  'Circle
      Top             =   4020
      Width           =   990
   End
   Begin VB.Shape Shape 
      FillStyle       =   0  'Solid
      Height          =   1005
      Index           =   2
      Left            =   990
      Shape           =   3  'Circle
      Top             =   2730
      Width           =   990
   End
   Begin VB.Shape Shape 
      FillStyle       =   0  'Solid
      Height          =   1005
      Index           =   1
      Left            =   990
      Shape           =   3  'Circle
      Top             =   1425
      Width           =   990
   End
   Begin VB.Shape Shape 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   1005
      Index           =   0
      Left            =   990
      Shape           =   3  'Circle
      Top             =   225
      Width           =   990
   End
End
Attribute VB_Name = "CLOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
''''''''''''''''''''
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE& = &H1
Private Const SWP_NOMOVE& = &H2

'''''''''''''
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2
Const LWA_COLORKEY = &H1
Private Sub cho_Click()
Dim opche  As Integer
opche = -1
For i = 0 To 4
If Optio(i).Value = True Then opche = i
Next i
If opche = -1 Then
MsgBox "请选择内容", , "选色"
Exit Sub
End If
On Error GoTo 88
cdc.ShowColor
Shape(opche).FillColor = cdc.Color
Shape(opche).BorderColor = cdc.Color
88 End Sub

Private Sub Form_Load()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
Dim rtn As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, vbWhite, 150, LWA_ALPHA
For i = 0 To 4
Shape(i).FillColor = ClSet(i)
Shape(i).BorderColor = ClSet(i)
Next i
End Sub

Private Sub OK_Click()
For i = 0 To 4
ClSet(i) = Shape(i).FillColor
Next i
Mom.slow.Value = 1
Unload CLOrder
End Sub
