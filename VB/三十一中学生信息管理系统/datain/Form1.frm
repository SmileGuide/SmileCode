VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   5730
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   495
      Left            =   1560
      TabIndex        =   15
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox Text 
      DataField       =   "other"
      DataSource      =   "Data1"
      Height          =   1000
      Index           =   14
      Left            =   1080
      TabIndex        =   14
      Text            =   "Text14"
      Top             =   5160
      Width           =   1000
   End
   Begin VB.TextBox Text 
      DataField       =   "mohone"
      DataSource      =   "Data1"
      Height          =   1000
      Index           =   13
      Left            =   1080
      TabIndex        =   13
      Text            =   "Text13"
      Top             =   4080
      Width           =   1000
   End
   Begin VB.TextBox Text 
      DataField       =   "mname"
      DataSource      =   "Data1"
      Height          =   1000
      Index           =   12
      Left            =   1080
      TabIndex        =   12
      Text            =   "Text12"
      Top             =   3360
      Width           =   1000
   End
   Begin VB.TextBox Text 
      DataField       =   "fphone"
      DataSource      =   "Data1"
      Height          =   1000
      Index           =   11
      Left            =   1080
      TabIndex        =   11
      Text            =   "Text11"
      Top             =   2640
      Width           =   1000
   End
   Begin VB.TextBox Text 
      DataField       =   "fname"
      DataSource      =   "Data1"
      Height          =   1000
      Index           =   10
      Left            =   1080
      TabIndex        =   10
      Text            =   "Text10"
      Top             =   1800
      Width           =   1000
   End
   Begin VB.TextBox Text 
      DataField       =   "number"
      DataSource      =   "Data1"
      Height          =   1000
      Index           =   9
      Left            =   1080
      TabIndex        =   9
      Text            =   "Text9"
      Top             =   960
      Width           =   1000
   End
   Begin VB.TextBox Text 
      DataField       =   "area"
      DataSource      =   "Data1"
      Height          =   1000
      Index           =   8
      Left            =   1080
      TabIndex        =   8
      Text            =   "Text8"
      Top             =   120
      Width           =   1000
   End
   Begin VB.TextBox Text 
      DataField       =   "village"
      DataSource      =   "Data1"
      Height          =   1000
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Text            =   "Text7"
      Top             =   5040
      Width           =   1000
   End
   Begin VB.TextBox Text 
      DataField       =   "street"
      DataSource      =   "Data1"
      Height          =   1000
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Text            =   "Text6"
      Top             =   4080
      Width           =   1000
   End
   Begin VB.TextBox Text 
      DataField       =   "city"
      DataSource      =   "Data1"
      Height          =   1000
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Text            =   "Text5"
      Top             =   3240
      Width           =   1000
   End
   Begin VB.TextBox Text 
      DataField       =   "pro"
      DataSource      =   "Data1"
      Height          =   1000
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   2520
      Width           =   1000
   End
   Begin VB.TextBox Text 
      DataField       =   "class"
      DataSource      =   "Data1"
      Height          =   1000
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1800
      Width           =   1000
   End
   Begin VB.TextBox Text 
      DataField       =   "sex"
      DataSource      =   "Data1"
      Height          =   1000
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   840
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "导入"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\@@@@sgbf\三十一中学生信息管理系统\三十一中学生信息库.mdb"
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   2520
      MousePointer    =   1  'Arrow
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "recoall"
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox Text 
      DataField       =   "name"
      DataSource      =   "Data1"
      Height          =   1000
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Excel.Workbooks.Open ("定义路径") '''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For i = 2 To 66
Data1.Recordset.AddNew
For j = 1 To 14
Text(j).Text = Trim(Excel.Application.Cells(i - 1, j))
Next j
Next i
End Sub

Private Sub Command2_Click()
Data1.Recordset.MoveFirst
For i = 1 To 1000
Data1.Recordset.Delete
On Error GoTo 3
Data1.Recordset.MoveNext
Next i
3 Data1.Recordset.MoveFirst

End Sub

Private Sub Picture1_Click()
Data1.UpdateRecord
End Sub

Private Sub Command3_Click()
For o = 1 To Data1.Recordset.RecordCount
If Text(1).Text = "" Then Data1.Recordset.Delete
Data1.Recordset.MoveNext
Next o
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub
