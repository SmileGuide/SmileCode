VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmWel 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7854
   ClientLeft      =   1206
   ClientTop       =   -2706
   ClientWidth     =   13344
   ControlBox      =   0   'False
   Icon            =   "Welcome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Welcome.frx":F9DA
   Picture         =   "Welcome.frx":10AA4
   ScaleHeight     =   7854
   ScaleWidth      =   13344
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CmdExit 
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   21.9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   12960
      TabIndex        =   4
      Top             =   0
      Width           =   400
   End
   Begin MSComDlg.CommonDialog CmD 
      Left            =   4080
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.bmp"
      DialogTitle     =   "保存图片"
      Filter          =   "高质量（BMP）|*.bmp|高质量（GIF）|*.gif|中等质量（JPG）|*.jpg"
   End
   Begin VB.CommandButton CmdOp 
      Caption         =   "…"
      BeginProperty Font 
         Name            =   "华文彩云"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   498
      Left            =   10860
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5820
      Visible         =   0   'False
      Width           =   498
   End
   Begin VB.Timer T1 
      Interval        =   200
      Left            =   2520
      Top             =   2100
   End
   Begin VB.CommandButton CmdEtr 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Start Your Orderly Life →"
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   15.9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   726
      Left            =   3900
      MouseIcon       =   "Welcome.frx":212EE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Visible         =   0   'False
      Width           =   3546
   End
   Begin MSComctlLib.ProgressBar Pg 
      Height          =   246
      Left            =   960
      TabIndex        =   2
      Top             =   5340
      Width           =   9588
      _ExtentX        =   16912
      _ExtentY        =   434
      _Version        =   393216
      Appearance      =   1
      Max             =   8
      Scrolling       =   1
   End
   Begin VB.Label LblWord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "凡事预则立，不预则废。——《礼记·中庸》"
      BeginProperty Font 
         Name            =   "华文新魏"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7344
   End
   Begin VB.Image ImgBk 
      Height          =   6480
      Left            =   0
      Picture         =   "Welcome.frx":223B8
      Top             =   0
      Width           =   11520
   End
   Begin VB.Menu MnOp 
      Caption         =   "选项"
      Visible         =   0   'False
      Begin VB.Menu MnWordCopy 
         Caption         =   "复制诗词到剪贴板"
      End
      Begin VB.Menu MnSplit 
         Caption         =   "-"
      End
      Begin VB.Menu MnChangePicture 
         Caption         =   "换一张图片"
      End
      Begin VB.Menu MnSaveAs 
         Caption         =   "保存图片"
      End
      Begin VB.Menu MnCopyPicture 
         Caption         =   "复制图片到剪贴板"
      End
   End
End
Attribute VB_Name = "FrmWel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim vpg As Integer
Dim Sen(1 To 21) As String
Dim i As Integer
Dim Frco(1 To 100)
Dim Frcolor
Dim ColorSum
Dim BkLink As String
'Dim BGX, BGY




Private Sub CmdEtr_Click()
FrmMn.Show
Set FrmWel = Nothing
Unload Me
FrmMn.FrmTab.Refresh
FrmMn.LblCap.Refresh
FrmMn.LblNow.Refresh
FrmMn.LblS.Refresh
FrmMn.CmdEdit.Refresh
FrmMn.CmdSet.Refresh
FrmMn.CmdTool.Refresh
FrmMn.LstL.Refresh
FrmMn.LstO.Refresh
FrmMn.LstTm.Refresh
FrmMn.ShpCap.Refresh

FrmMn.Refresh

End Sub

Private Sub CmdEtr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdEtr.Top = CmdEtr.Top + 30

End Sub

Private Sub CmdEtr_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdEtr.Top = CmdEtr.Top - 30
End Sub

Private Sub CmdExit_Click()
End
End Sub

Private Sub CmdOp_Click()
PopupMenu MnOp
End Sub

Private Sub CmdOp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdOp.Top = CmdOp.Top + 30

End Sub

Private Sub CmdOp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdOp.Top = CmdOp.Top - 30
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
End
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
MkDir App.Path & "\SmTab"
FrmMn.Show
FrmMn.Hide
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
BkLink = GetSetting("SmileTimetable", "Picture", "Link")
'https://api.ixiaowai.cn/gqapi/gqapi.php
If BkLink = "" Then
    ImgBk.Picture = LoadPicture("https://api.ixiaowai.cn/gqapi/gqapi.php")
    SaveSetting "SmileTimetable", "Picture", "Link", "https://api.ixiaowai.cn/gqapi/gqapi.php"
Else
ImgBk.Picture = LoadPicture(BkLink)
End If
FrmWel.Picture = ImgBk.Picture
'For BGX = FrmWel.Width - 100 To 1 Step -1
    'For BGY = FrmWel.Height - 100 To 1 Step -1

        'If Me.Point(BGX, BGY) <> 16777215 Then
            'Exit For
        'End If
    'Next BGY
    'If Me.Point(BGX, BGY) <> 16777215 Then
        'Exit For
    'End If
'Next BGX
'Me.Height = BGY
'Me.Width = BGX
Me.Height = ImgBk.Height
Me.Width = ImgBk.Width
LblWord.Caption = IEWord

LblWord.Left = Me.Width / 2 - LblWord.Width / 2
CmdOp.Top = FrmWel.Height - 660
CmdOp.Left = FrmWel.Width - 660
CmdEtr.Left = Me.Width / 2 - CmdEtr.Width / 2
CmdEtr.Top = Me.Height - 1200
Pg.Left = Me.Width / 2 - Pg.Width / 2
Pg.Top = Me.Height - 1140
CmdExit.Left = Me.Width - CmdExit.Width
CmdExit.Top = 0



For i = 1 To 100
Frco(i) = Me.Point(Me.Width / 100 * i, Me.Height / 100 * i)
ColorSum = ColorSum + Frco(i)

Next
Frcolor = ColorSum / 100
Frcolor = Int(Frcolor)
Dim r, g, B
r = Left(Hex(Me.Point(LblWord.Left, LblWord.Top)), 2)
g = Mid(Hex(Me.Point(LblWord.Left, LblWord.Top)), 3, 2)
B = Right(Hex(Me.Point(LblWord.Left, LblWord.Top)), 2)

If r <= Hex(51) And g <= Hex(51) And B <= Hex(51) Then
    LblWord.ForeColor = 16777215 - Frcolor
    CmdEtr.BackColor = Frcolor
    CmdOp.BackColor = Frcolor
Else
    
    LblWord.ForeColor = Frcolor
    CmdEtr.BackColor = 16777215 - Frcolor
    CmdOp.BackColor = 16777215 - Frcolor
    CmdExit.BackColor = 16777215 - Frcolor
End If
EnMiddle Me
FrmWelLoad = True
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ReturnVal As Long
X = ReleaseCapture()
ReturnVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub Form_Paint()
If Me.Point(0, 0) = &HFFFFC0 Then Me.Refresh

End Sub

Private Sub ImgBk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ReturnVal As Long
X = ReleaseCapture()
ReturnVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub MnChangeWord_Click()
LblWord.Caption = IEWord
End Sub

Private Sub LblWord_Click()

End Sub

Private Sub MnChangePicture_Click()
'https://api.ixiaowai.cn/gqapi/gqapi.php
ImgBk.Picture = LoadPicture(BkLink)
FrmWel.Picture = ImgBk.Picture
'For BGX = FrmWel.Width - 100 To 1 Step -1
    'For BGY = FrmWel.Height - 100 To 1 Step -1

        'If Me.Point(BGX, BGY) <> 16777215 Then
            'Exit For
        'End If
    'Next BGY
    'If Me.Point(BGX, BGY) <> 16777215 Then
        'Exit For
    'End If
'Next BGX
'Me.Height = BGY
'Me.Width = BGX
Me.Height = ImgBk.Height
Me.Width = ImgBk.Width
LblWord.Left = Me.Width / 2 - LblWord.Width / 2
CmdOp.Top = FrmWel.Height - 660
CmdOp.Left = FrmWel.Width - 660
CmdEtr.Left = Me.Width / 2 - CmdEtr.Width / 2
CmdEtr.Top = Me.Height - 1200
Pg.Left = Me.Width / 2 - Pg.Width / 2
Pg.Top = Me.Height - 1140

Frcolor = 0
ColorSum = 0
For i = 1 To 100
Frco(i) = Me.Point(Me.Width / 100 * i, Me.Height / 100 * i)
ColorSum = ColorSum + Frco(i)

Next
Frcolor = ColorSum / 100
Frcolor = Int(Frcolor)
Dim r, g, B
r = Left(Hex(Me.Point(LblWord.Left, LblWord.Top)), 2)
g = Mid(Hex(Me.Point(LblWord.Left, LblWord.Top)), 3, 2)
B = Right(Hex(Me.Point(LblWord.Left, LblWord.Top)), 2)

If r <= Hex(51) And g <= Hex(51) And B <= Hex(51) Then
    LblWord.ForeColor = 16777215 - Frcolor
    CmdEtr.BackColor = Frcolor
    CmdOp.BackColor = Frcolor
Else
    LblWord.ForeColor = Frcolor
    On Error GoTo 99
    CmdEtr.BackColor = 16777215 - Frcolor
    CmdOp.BackColor = 16777215 - Frcolor
    CmdExit.BackColor = 16777215 - Frcolor
End If
99 End Sub

Private Sub MnCopyPicture_Click()
Clipboard.Clear
Clipboard.SetData ImgBk.Picture
Msg "已复制", &HDFFFB0, 500
End Sub

Private Sub MnSaveAs_Click()
CmD.InitDir = App.Path
CmD.FileName = Format(Now, "yyyy-mm-dd-hh-mm-ss")
On Error GoTo 99
CmD.ShowSave
SavePicture ImgBk.Picture, CmD.FileName
Msg "已保存至" & CmD.FileName, &HDFFFB0, 1000
99 End Sub

Private Sub MnWordCopy_Click()
Clipboard.Clear
Clipboard.SetText LblWord.Caption
Msg "已复制", &HDFFFB0, 500
End Sub

Private Sub T1_Timer()
    vpg = vpg + 1
    Pg.Value = vpg
    If vpg = 8 Then
        T1.Enabled = False
        Pg.Visible = False
        CmdEtr.Visible = True
        CmdOp.Visible = True
        
    End If
    
End Sub
'''''''''''''''''''''
'https://api.dujin.org/bing/1366.php
' https://api.ixiaowai.cn/gqapi/gqapi.php
'https://api.dujin.org/bing/1920.php
 'https://unsplash.it/1600/900?random
' https://img.xjh.me/random_img.php?type=bg&ctype=nature&return=302
 
 
