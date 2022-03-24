VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmCED 
   BackColor       =   &H000C0C0C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "高级编辑"
   ClientHeight    =   6732
   ClientLeft      =   18
   ClientTop       =   354
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6732
   ScaleWidth      =   11160
   StartUpPosition =   3  '窗口缺省
   Begin RichTextLib.RichTextBox TxtCode 
      Height          =   6726
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11166
      _ExtentX        =   19696
      _ExtentY        =   11864
      _Version        =   393217
      BackColor       =   789516
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"FrmCED.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmCED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub TxtCode_Change()

On Error Resume Next
If Mid(TxtCode.Text, TxtCode.SelStart, 2) = Chr(10) Then GoTo 45
On Error Resume Next
33 TxtCode.SelStart = TxtCode.SelStart - 1
TxtCode.SelLength = 1
TxtCode.SelColor = &HFFFFFF
If Right(TxtCode.Text, 1) = Chr(34) Or Right(TxtCode.Text, 1) = Chr(44) Then
TxtCode.SelLength = 1
TxtCode.SelColor = &HFF00FF

ElseIf Right(TxtCode.Text, 1) = Chr(46) Then
TxtCode.SelLength = 1
TxtCode.SelColor = &HFF00&

ElseIf Right(TxtCode.Text, 1) = Chr(58) Then
TxtCode.SelLength = 1
TxtCode.SelColor = &HFF00FF

ElseIf Right(TxtCode.Text, 1) = Chr(59) Then
TxtCode.SelLength = 1
TxtCode.SelColor = &HFF00&

ElseIf Right(TxtCode.Text, 1) = Chr(123) Then
TxtCode.SelLength = 1
TxtCode.SelColor = &HFF00&

ElseIf Right(TxtCode.Text, 1) = Chr(125) Then
TxtCode.SelLength = 1
TxtCode.SelColor = &HFF00&

ElseIf Asc(Right(TxtCode.Text, 1)) >= 48 And Asc(Right(TxtCode.Text, 1)) <= 57 Then
TxtCode.SelLength = 1
TxtCode.SelColor = &HFFFF&

ElseIf Right(LCase(TxtCode.Text), 3) = "day" Then
TxtCode.SelStart = TxtCode.SelStart - 2
TxtCode.SelLength = 3
TxtCode.SelColor = &HFF00&
TxtCode.SelStart = TxtCode.SelStart + 2

ElseIf Right(LCase(TxtCode.Text), 5) = "clear" Then
TxtCode.SelStart = TxtCode.SelStart - 4
TxtCode.SelLength = 5
TxtCode.SelColor = &HFF00&
TxtCode.SelStart = TxtCode.SelStart + 4

ElseIf Right(LCase(TxtCode.Text), 4) = "quit" Then
TxtCode.SelStart = TxtCode.SelStart - 3
TxtCode.SelLength = 4
TxtCode.SelColor = &HFF00&
TxtCode.SelStart = TxtCode.SelStart + 3

End If
'贴尾归位
TxtCode.SelStart = TxtCode.SelStart + 1
99 TxtCode.SelLength = 0
Exit Sub
45 On Error GoTo 999
If Mid(TxtCode.Text, TxtCode.SelStart - 1, 2) = Chr(46) & vbCrLf Then
        ''''''''''''保存
    ElseIf Mid(TxtCode.Text, TxtCode.SelStart - 5, 6) = "quit" & vbCrLf Then
        Unload Me
    ElseIf Mid(TxtCode.Text, TxtCode.SelStart - 6, 7) = "clear" & vbCrLf Then
        TxtCode.Text = ""
    End If
999 End Sub

'''''''''''''''''''''''
'补充命令的反馈
'help
'format格式套
