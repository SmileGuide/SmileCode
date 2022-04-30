Attribute VB_Name = "MdlDeclare"
Option Explicit
Public DWord As String
Public DFrom As String
Public DMix As String
Public BColorX(1 To 100) As Single
Public BcolorY(1 To 100) As Single
Public FrmWelLoad As Boolean
Public StName As String
Public NDay As Integer
Public Saved As Boolean
Public SelL As Integer
Public Covered As Boolean
Public SknColor
Public CvrText 'which command results the cover
Public Numcolor
Public TxtColor

Public SpcColor
Public CmdColor
Public TxtFont
Public UnDis As Boolean '±‹√‚µ›πÈÀ¿—≠ª∑
Public ITRef As Boolean
Public AlText As String
Public AllTxt As String
Public GruDay As Variant
Public FstOpen As Boolean
Public MxFTS As Variant
Public TxtSize As String







Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST& = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE& = &H1
Public Const SWP_NOMOVE& = &H2


