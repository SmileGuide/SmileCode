VERSION 5.00
Begin VB.Form FrmNewEdit 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "     我的课程表 - 微笑课程表 [编辑]"
   ClientHeight    =   4656
   ClientLeft      =   7878
   ClientTop       =   9438
   ClientWidth     =   5442
   Icon            =   "FrmNewEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4656
   ScaleWidth      =   5442
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox CblDay 
      BackColor       =   &H00DFFFB0&
      Height          =   228
      ItemData        =   "FrmNewEdit.frx":1084A
      Left            =   4380
      List            =   "FrmNewEdit.frx":10863
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "星期一"
      Top             =   60
      Width           =   846
   End
   Begin VB.TextBox TxtName 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "华文中宋"
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
      Text            =   "我的课程表"
      ToolTipText     =   "单击可修改"
      Top             =   60
      Width           =   3126
   End
   Begin VB.Frame FrmTab 
      BackColor       =   &H00C0FFFF&
      Height          =   3726
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   5224
      Begin VB.CommandButton CmdReturn 
         BackColor       =   &H00FDEEBF&
         Caption         =   "还原至保存前"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   306
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "小工具"
         Top             =   3360
         Width           =   1416
      End
      Begin VB.CommandButton CmdOK 
         BackColor       =   &H00FDEEBF&
         Caption         =   "保存该天课程"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   306
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "小工具"
         Top             =   3360
         Width           =   1356
      End
      Begin VB.CommandButton CmdO 
         BackColor       =   &H00C0FFC0&
         Caption         =   "按时间排序"
         BeginProperty Font 
            Name            =   "华文仿宋"
            Size            =   9
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
         ToolTipText     =   "小工具"
         Top             =   240
         Width           =   1236
      End
      Begin VB.CommandButton CmdTDel 
         BackColor       =   &H00C0FFC0&
         Caption         =   "删除"
         BeginProperty Font 
            Name            =   "华文仿宋"
            Size            =   9
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
         ToolTipText     =   "小工具"
         Top             =   240
         Width           =   756
      End
      Begin VB.CommandButton CmdTAdd 
         BackColor       =   &H00C0FFC0&
         Caption         =   "添加"
         BeginProperty Font 
            Name            =   "华文仿宋"
            Size            =   9
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
         ToolTipText     =   "小工具"
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
               Name            =   "华文中宋"
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
         ScaleWidth      =   1782
         TabIndex        =   3
         Top             =   600
         Width           =   1804
         Begin VB.ComboBox CboLCha 
            BackColor       =   &H00C0FFFF&
            Height          =   228
            ItemData        =   "FrmNewEdit.frx":108A3
            Left            =   60
            List            =   "FrmNewEdit.frx":108C2
            TabIndex        =   6
            Text            =   "输入课程"
            Top             =   60
            Width           =   1684
         End
         Begin VB.ListBox LstL 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "华文中宋"
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
            List            =   "FrmNewEdit.frx":1090F
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
            Text            =   "输入分"
            Top             =   60
            Width           =   1024
         End
         Begin VB.ComboBox CboHC 
            BackColor       =   &H00C0FFFF&
            Height          =   228
            Left            =   60
            TabIndex        =   5
            Text            =   "输入时"
            Top             =   60
            Width           =   1024
         End
         Begin VB.ListBox LstTm 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "华文中宋"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   2184
            ItemData        =   "FrmNewEdit.frx":10922
            Left            =   60
            List            =   "FrmNewEdit.frx":10935
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   300
            Width           =   2226
         End
         Begin VB.Label LblJ 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "华文新魏"
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
         Caption         =   "Smile TimeTabe"
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
         ToolTipText     =   "课程表中的紫色总结文字颜色"
         Top             =   240
         Width           =   1386
      End
      Begin VB.Image ImgLogo 
         Height          =   384
         Left            =   240
         Picture         =   "FrmNewEdit.frx":10948
         Top             =   180
         Width           =   384
      End
   End
   Begin VB.Label LblCpC 
      BackStyle       =   0  'Transparent
      Caption         =   "课程表名称："
      Height          =   186
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "课程表中的紫色总结文字颜色"
      Top             =   60
      Width           =   1086
   End
   Begin VB.Menu MnFile 
      Caption         =   "文件"
      Visible         =   0   'False
      Begin VB.Menu MnRename 
         Caption         =   "重命名"
      End
      Begin VB.Menu MnSave 
         Caption         =   "保存为我的课程表"
      End
      Begin VB.Menu MnExIn 
         Caption         =   "从电子表格导入"
      End
      Begin VB.Menu MnToEx 
         Caption         =   "输出为电子表格"
      End
   End
   Begin VB.Menu MnDay 
      Caption         =   "星期"
      Visible         =   0   'False
      Begin VB.Menu MnDI 
         Caption         =   "星期1"
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
NDay = DayToNum(CblDay.Text)
Open App.Path & "\SmTab\" & StName & ".smtab" & NDay For Input As #1
Dim j, CTM, CL, k
For j = 1 To FileLen(App.Path & "\SmTab\" & StName & ".smtab" & NDay)
Input #1, CTM, CL
LstTm.AddItem CTM
LstL.AddItem CL
Next j
Close #1
For k = 0 To LstTm.ListCount - 1
LstO.AddItem k + 1, k
Next
End Sub

Private Sub CmdO_Click()
NumberTbl LstTm, LstL
End Sub

Private Sub CmdOK_Click()
Saved = True
Open App.Path & "\SmTab\" & StName & ".smtab" & NDay For Output As #1
Dim j
For j = 0 To LstTm.ListCount - 1
Write #1, LstTm.List(j), LstL.List(j)
Next j
Close #1
Msg "已保存" & NumToDay(NDay) & "的课程", &HFDEEBF, 1000
End Sub

Private Sub CmdReturn_Click()

Dim i, p, CTM, CL
For i = 0 To LstTm.ListCount
LstTm.RemoveItem i
LstL.RemoveItem i
Next
For p = 1 To FileLen(App.Path & "\SmTab\" & StName & ".smtab" & NDay)
Input #1, CTM, CL
LstTm.AddItem CTM
LstL.AddItem CL
Next
End Sub

Private Sub CmdTAdd_Click()
If Not IsNumeric(CboHC.Text) Then
MsgBox "请输入数字", vbExclamation, "微笑课程表"
CboHC.Text = ""
CboHC.SetFocus
Exit Sub
End If
If Not IsNumeric(CboMC.Text) Then
MsgBox "请输入数字", vbExclamation, "微笑课程表"
CboMC.Text = ""
CboMC.SetFocus
Exit Sub
End If
CboHC.Text = Format(CboHC.Text, "00")
CboMC.Text = Format(CboMC.Text, "00")
Dim a, B
For a = 0 To LstTm.ListCount - 1
If CboHC.Text & ":" & CboMC.Text = LstTm.List(a) Then
    MsgBox "时间重复，不能添加", vbExclamation, "微笑课程表"
    Exit Sub
End If
Next
For B = 0 To LstTm.ListCount - 1
If CboHC.Text & ":" & CboMC.Text = LstTm.List(a) Then
    MsgBox "时间重复，不能添加", vbExclamation, "微笑课程表"
    Exit Sub
End If
Next
Dim i
For i = 0 To LstTm.ListCount - 1
If LstTm.Selected(i) = True Then TMSel = i
Next
On Error GoTo 99
LstTm.AddItem CboHC.Text & ":" & CboMC.Text, TMSel + 1
LstL.AddItem CboLCha.Text, TMSel + 1
Dim k, r
For r = 0 To LstO.ListCount - 1
LstO.RemoveItem r
Next
For k = 0 To LstTm.ListCount - 1
LstO.AddItem k + 1, k
Next
Exit Sub
99 LstTm.AddItem CboHC.Text & ":" & CboMC.Text
LstL.AddItem CboLCha.Text

For r = 0 To LstO.ListCount - 1
LstO.RemoveItem r
Next
For k = 0 To LstTm.ListCount - 1
LstO.AddItem k + 1, k
Next
End Sub




Private Sub CmdTDel_Click()
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
Dim i, j, k
TxtName.Text = StName
For i = 1 To 24
    CboHC.AddItem Format(i, "00")
Next
For j = 0 To 50 Step 10
i = Format(j, "00")
    CboMC.AddItem Format(j, "00")
Next
'''''''
Dim r
For r = 0 To LstO.ListCount - 1
LstO.RemoveItem r
Next
For k = 0 To LstTm.ListCount - 1
LstO.AddItem k + 1, k
Next
'''''''''''
NDay = 1
On Error Resume Next
Open App.Path & "\SmTab\" & StName & ".smtab" & NDay For Input As #1

'''''''
Dim p, CTM, CL
For p = 1 To 100
If EOF(1) = True Then Exit For
Input #1, CTM, CL
LstTm.AddItem CTM
LstL.AddItem CL
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
'完善对时间的过滤
'删除项，写文件
'Frmexit


