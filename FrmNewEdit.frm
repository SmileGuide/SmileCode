VERSION 5.00
Begin VB.Form FrmNewEdit 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "     �ҵĿγ̱� - ΢Ц�γ̱� [�༭]"
   ClientHeight    =   4182
   ClientLeft      =   7878
   ClientTop       =   9438
   ClientWidth     =   5382
   Icon            =   "FrmNewEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4182
   ScaleWidth      =   5382
   StartUpPosition =   1  '����������
   Begin VB.ComboBox CblDay 
      BackColor       =   &H00FAF9D6&
      Height          =   228
      ItemData        =   "FrmNewEdit.frx":1084A
      Left            =   4380
      List            =   "FrmNewEdit.frx":10863
      TabIndex        =   17
      Text            =   "����һ"
      Top             =   60
      Width           =   846
   End
   Begin VB.TextBox TxtName 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "��������"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   184
      Left            =   1200
      TabIndex        =   10
      Text            =   "�ҵĿγ̱�"
      ToolTipText     =   "�������޸�"
      Top             =   60
      Width           =   3126
   End
   Begin VB.Frame FrmTab 
      BackColor       =   &H00C0FFFF&
      Height          =   3786
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   5224
      Begin VB.CommandButton CmdCode 
         BackColor       =   &H00C0E0FF&
         Caption         =   "�߼��༭..."
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   306
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "С����"
         Top             =   3300
         Width           =   2376
      End
      Begin VB.CommandButton CmdReturn 
         BackColor       =   &H00FED8E7&
         Caption         =   "��ԭ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   306
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "С����"
         Top             =   3300
         Width           =   1176
      End
      Begin VB.CommandButton CmdOK 
         BackColor       =   &H00FDEEBF&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   306
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "С����"
         Top             =   3300
         Width           =   1176
      End
      Begin VB.CommandButton CmdO 
         BackColor       =   &H00C0FFC0&
         Caption         =   "��ʱ������"
         BeginProperty Font 
            Name            =   "���ķ���"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   245
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "С����"
         Top             =   240
         Width           =   1236
      End
      Begin VB.CommandButton CmdTDel 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ɾ��"
         BeginProperty Font 
            Name            =   "���ķ���"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   245
         Left            =   3060
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "С����"
         Top             =   240
         Width           =   756
      End
      Begin VB.CommandButton CmdTAdd 
         BackColor       =   &H00C0FFC0&
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "���ķ���"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   245
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "С����"
         Top             =   240
         Width           =   756
      End
      Begin VB.PictureBox PicO 
         BackColor       =   &H0080FFFF&
         Height          =   2586
         Left            =   120
         ScaleHeight     =   2562
         ScaleWidth      =   642
         TabIndex        =   7
         Top             =   600
         Width           =   664
         Begin VB.ListBox LstO 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "��������"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   2184
            ItemData        =   "FrmNewEdit.frx":1089F
            Left            =   60
            List            =   "FrmNewEdit.frx":108A1
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   300
            Width           =   544
         End
      End
      Begin VB.PictureBox PicL 
         BackColor       =   &H0080FFFF&
         Height          =   2586
         Left            =   3240
         ScaleHeight     =   2562
         ScaleWidth      =   1842
         TabIndex        =   3
         Top             =   600
         Width           =   1866
         Begin VB.ComboBox CboLCha 
            BackColor       =   &H00C0FFFF&
            Height          =   228
            ItemData        =   "FrmNewEdit.frx":108A3
            Left            =   60
            List            =   "FrmNewEdit.frx":108C2
            TabIndex        =   6
            Text            =   "����γ�"
            Top             =   60
            Width           =   1684
         End
         Begin VB.ListBox LstL 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "��������"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   2184
            ItemData        =   "FrmNewEdit.frx":108FC
            Left            =   60
            List            =   "FrmNewEdit.frx":108FE
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   300
            Width           =   1684
         End
      End
      Begin VB.PictureBox PicT 
         BackColor       =   &H0080FFFF&
         Height          =   2586
         Left            =   840
         ScaleHeight     =   2562
         ScaleWidth      =   2322
         TabIndex        =   1
         Top             =   600
         Width           =   2344
         Begin VB.ComboBox CboMC 
            BackColor       =   &H00C0FFFF&
            Height          =   228
            Left            =   1260
            TabIndex        =   11
            Text            =   "�����"
            Top             =   60
            Width           =   1024
         End
         Begin VB.ComboBox CboHC 
            BackColor       =   &H00C0FFFF&
            Height          =   228
            Left            =   60
            TabIndex        =   5
            Text            =   "����ʱ"
            Top             =   60
            Width           =   1024
         End
         Begin VB.ListBox LstTm 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "��������"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   2184
            ItemData        =   "FrmNewEdit.frx":10900
            Left            =   60
            List            =   "FrmNewEdit.frx":10902
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   300
            Width           =   2226
         End
         Begin VB.Label LblJ 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "������κ"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   184
            Left            =   1140
            TabIndex        =   12
            Top             =   60
            Width           =   184
         End
      End
      Begin VB.Label LblSign 
         BackStyle       =   0  'Transparent
         Caption         =   "Smile TimeTable"
         BeginProperty Font 
            Name            =   "Blackadder ITC"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   306
         Left            =   720
         TabIndex        =   16
         ToolTipText     =   "�γ̱��е���ɫ�ܽ�������ɫ"
         Top             =   240
         Width           =   1386
      End
      Begin VB.Image ImgLogo 
         Height          =   384
         Left            =   240
         Picture         =   "FrmNewEdit.frx":10904
         Top             =   180
         Width           =   384
      End
   End
   Begin VB.Label LblCpC 
      BackStyle       =   0  'Transparent
      Caption         =   "�γ̱����ƣ�"
      Height          =   186
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "�γ̱��е���ɫ�ܽ�������ɫ"
      Top             =   60
      Width           =   1086
   End
   Begin VB.Menu MnFile 
      Caption         =   "�ļ�"
      Visible         =   0   'False
      Begin VB.Menu MnRename 
         Caption         =   "������"
      End
      Begin VB.Menu MnSave 
         Caption         =   "����Ϊ�ҵĿγ̱�"
      End
      Begin VB.Menu MnExIn 
         Caption         =   "�ӵ��ӱ����"
      End
      Begin VB.Menu MnToEx 
         Caption         =   "���Ϊ���ӱ��"
      End
   End
   Begin VB.Menu MnDay 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu MnDI 
         Caption         =   "����1"
         Index           =   1
      End
   End
End
Attribute VB_Name = "FrmNewEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TMSel As Integer
Dim LSel As Integer

Private Sub CblDay_Change()
If Not Saved Then
Dim Cls
Cls = MsgBox("�Ƿ񱣴���ģ�", vbOKCancel, "΢Ц�γ̱�")
If Cls = vbOK Then
NumberTbl LstTm, LstL
Open App.Path & "\SmTab\" & StName & ".smtab" & NDay For Output As #1
Dim jm
For jm = 0 To LstTm.ListCount - 1
Write #1, LstTm.List(jm), LstL.List(jm)
Next jm
Close #1
Msg "�ѱ���" & NumToDay(NDay) & "�Ŀγ�", &HFDEEBF, 1000
Save = True
End If

End If
NDay = DayToNum(CblDay.Text)
If NDay = 0 Then Exit Sub
On Error Resume Next
Open App.Path & "\SmTab\" & StName & ".smtab" & NDay For Input As #1
Dim j, CTM, CL, k
LstTm.Clear
LstL.Clear
LstO.Clear
For j = 1 To 100
If EOF(1) = True Then Exit For
Input #1, CTM, CL
LstTm.AddItem CTM
LstL.AddItem CL

Next j
Close #1
For k = 0 To LstTm.ListCount - 1
LstO.AddItem k + 1, k
Next
If LstTm.List(0) = "" Then LstO.Clear
Dim kl, kll
kll = LstO.ListCount - 1
For kl = 0 To kll

If LstTm.List(kl) = "" Then klli = kl: GoTo 99
Next
LstO.Clear
For k = 0 To LstTm.ListCount - 1
LstO.AddItem k + 1, k
Next
Exit Sub
99 kll = LstTm.ListCount - 1
For k = klli To kll
LstTm.RemoveItem (klli)
LstL.RemoveItem (klli)
Next
LstO.Clear
For k = 0 To LstTm.ListCount - 1
LstO.AddItem k + 1, k
Next
End Sub

Private Sub CblDay_Click()
If Not Saved Then
Dim Cls
Cls = MsgBox("�Ƿ񱣴���ģ�", vbOKCancel, "΢Ц�γ̱�")
If Cls = vbOK Then
NumberTbl LstTm, LstL
Open App.Path & "\SmTab\" & StName & ".smtab" & NDay For Output As #1
Dim jm
For jm = 0 To LstTm.ListCount - 1
Write #1, LstTm.List(jm), LstL.List(jm)
Next jm
Close #1
Msg "�ѱ���" & NumToDay(NDay) & "�Ŀγ�", &HFDEEBF, 1000
Save = True
End If

End If
NDay = DayToNum(CblDay.Text)
If NDay = 0 Then Msg "��������ȷ����������ѡ�������б��е�Ԥ��ֵ", &H8080FF, "1500": Exit Sub

Open App.Path & "\SmTab\" & StName & ".smtab" & NDay For Input As #1
Dim j, CTM, CL, k
LstTm.Clear
LstL.Clear
LstO.Clear
For j = 1 To 100
If EOF(1) = True Then Exit For
Input #1, CTM, CL
LstTm.AddItem CTM
LstL.AddItem CL

Next j
Close #1
For k = 0 To LstTm.ListCount - 1
LstO.AddItem k + 1, k
Next
If LstTm.List(0) = "" Then LstO.Clear
Dim kl, kll, klli
kll = LstO.ListCount - 1
For kl = 0 To 100
On Error Resume Next
If LstTm.List(kl) = "" Then klli = kl: GoTo 99
Next
LstO.Clear
For k = 0 To LstTm.ListCount - 1
LstO.AddItem k + 1, k
Next
Exit Sub
99 kll = LstTm.ListCount - 1
For k = klli To kll
LstTm.RemoveItem (klli)
LstL.RemoveItem (klli)
Next
LstO.Clear
For k = 0 To LstTm.ListCount - 1
LstO.AddItem k + 1, k
Next
End Sub

Private Sub CmdO_Click()
Saved = False
NumberTbl LstTm, LstL
End Sub

Private Sub CmdOK_Click()
Saved = True
NumberTbl LstTm, LstL
Open App.Path & "\SmTab\" & StName & ".smtab" & NDay For Output As #1
Dim j
For j = 0 To LstTm.ListCount - 1
Write #1, LstTm.List(j), LstL.List(j)
Next j
Close #1
Msg "�ѱ���" & NumToDay(NDay) & "�Ŀγ�", &HFDEEBF, 1000
End Sub

Private Sub CmdReturn_Click()
Saved = True
Dim p, CTM, CL
LstTm.Clear
LstL.Clear
LstO.Clear
For p = 1 To 100
If EOF(i) Then Exit For
Input #1, CTM, CL
LstTm.AddItem CTM
LstL.AddItem CL
Next
End Sub

Private Sub CmdTAdd_Click()
If Not IsNumeric(CboHC.Text) Then
Msg "����������", &H8080FF, 500
CboHC.Text = ""
CboHC.SetFocus
Exit Sub
End If
If Not IsNumeric(CboMC.Text) Then
Msg "����������", &H8080FF, 500
CboMC.Text = ""
CboMC.SetFocus
Exit Sub
End If
If CboLCha.Text = "����γ�" Or "" Then
Msg "������γ�", &H8080FF, 500
exitsub
End If
Saved = False
CboHC.Text = Format(CboHC.Text, "00")
CboMC.Text = Format(CboMC.Text, "00")
Dim a, B
For a = 0 To LstTm.ListCount - 1
If CboHC.Text & ":" & CboMC.Text = LstTm.List(a) Then
    Msg "ʱ���ظ����������", &H8080FF, 500
    Exit Sub
End If
Next
For B = 0 To LstTm.ListCount - 1
If CboHC.Text & ":" & CboMC.Text = LstTm.List(a) Then
    Msg "ʱ���ظ����������", &H8080FF, 500
    Exit Sub
End If
Next
Dim i
TMSel = -1
For i = 0 To LstTm.ListCount - 1
If LstTm.Selected(i) = True Then TMSel = i
Next
LstTm.AddItem CboHC.Text & ":" & CboMC.Text, TMSel + 1
LstL.AddItem CboLCha.Text, TMSel + 1
Dim k, r, l
l = LstO.ListCount - 1
For r = 0 To l
LstO.RemoveItem 0
Next
For k = 0 To LstTm.ListCount - 1
LstO.AddItem k + 1, k
Next
End Sub




Private Sub CmdTDel_Click()
Saved = False
LstTm.RemoveItem SelL
LstL.RemoveItem SelL
Dim k, r
For r = 0 To LstO.ListCount - 1

LstO.RemoveItem 0
Next
For k = 0 To LstTm.ListCount - 1
LstO.AddItem k + 1, k
Next
End Sub

Private Sub Form_Load()
Saved = True
Dim i, j, k
StName = 1
TxtName.Text = StName
For i = 1 To 24
    CboHC.AddItem Format(i, "00")
Next
For j = 0 To 50 Step 10
i = Format(j, "00")
    CboMC.AddItem Format(j, "00")
Next
'''''''
Dim r, l

'''''''''''
NDay = 1

Open App.Path & "\SmTab\" & StName & ".smtab" & NDay For Input As #1

'''''''
Dim p, CTM, CL
For p = 1 To 100
If EOF(1) = True Then Close #1: Exit For
Input #1, CTM, CL
LstTm.AddItem CTM
LstL.AddItem CL
Next
l = LstO.ListCount - 1
For r = 0 To l
LstO.RemoveItem 0
Next
For k = 0 To LstTm.ListCount - 1
LstO.AddItem k + 1, k
Next
End Sub

Private Sub Form_Paint()
TxtName.SetFocus
End Sub

Private Sub LstL_Click()
Dim i
For i = 0 To LstL.ListCount - 1
If LstL.Selected(i) = True Then
SelL = i
LstTm.Selected(i) = True
LstO.Selected(i) = True
End If
Next


End Sub

Private Sub LstO_Click()
Dim i
For i = 0 To LstTm.ListCount - 1
If LstO.Selected(i) = True Then
SelL = i
LstL.Selected(i) = True
LstTm.Selected(i) = True
End If
Next
End Sub

Private Sub LstTm_Click()
Dim i
For i = 0 To LstTm.ListCount - 1
If LstTm.Selected(i) = True Then
SelL = i
LstL.Selected(i) = True
LstO.Selected(i) = True
End If
Next

End Sub

''''''''''''''
'���ƶ�ʱ��Ĺ���
'ɾ���д�ļ�
'Frmexit


