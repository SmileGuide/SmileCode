VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmCED 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�߼��༭"
   ClientHeight    =   3732
   ClientLeft      =   18
   ClientTop       =   354
   ClientWidth     =   6294
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3732
   ScaleWidth      =   6294
   StartUpPosition =   3  '����ȱʡ
   Begin RichTextLib.RichTextBox TxtCode 
      Height          =   3726
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6306
      _ExtentX        =   11123
      _ExtentY        =   6572
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"FrmCED.frx":0000
   End
End
Attribute VB_Name = "FrmCED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub TxtCode_Change()


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
'��β��λ
TxtCode.SelStart = TxtCode.SelStart + 1
99 TxtCode.SelLength = 0
Exit Sub
45 On Error Resume Next
If Mid(TxtCode.Text, TxtCode.SelStart - 1, 2) = Chr(46) & vbCrLf Then
        ''''''''''''����
    ElseIf Mid(TxtCode.Text, TxtCode.SelStart - 5, 6) = "quit" & vbCrLf Then
        Unload Me
    ElseIf Mid(TxtCode.Text, TxtCode.SelStart - 6, 7) = "clear" & vbCrLf Then
        TxtCode.Text = ""
    End If
End Sub
