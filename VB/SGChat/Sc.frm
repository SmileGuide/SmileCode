VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Sc 
   Caption         =   "个性"
   ClientHeight    =   1244
   ClientLeft      =   16
   ClientTop       =   328
   ClientWidth     =   3220
   LinkTopic       =   "Form1"
   ScaleHeight     =   1244
   ScaleWidth      =   3220
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdOK 
      Caption         =   "完成"
      Height          =   244
      Left            =   2520
      TabIndex        =   6
      Top             =   960
      Width           =   604
   End
   Begin VB.TextBox Text1 
      Height          =   208
      Left            =   960
      TabIndex        =   5
      Text            =   "改变世界！"
      Top             =   660
      Width           =   2164
   End
   Begin VB.TextBox TxtName 
      Height          =   208
      Left            =   960
      TabIndex        =   3
      Text            =   "用户"
      Top             =   60
      Width           =   2164
   End
   Begin MSComCtl2.DTPicker DTBrth 
      Height          =   184
      Left            =   960
      TabIndex        =   2
      Top             =   360
      Width           =   2164
      _ExtentX        =   3818
      _ExtentY        =   325
      _Version        =   393216
      Format          =   161742849
      CurrentDate     =   2
   End
   Begin VB.Label LblOth 
      Caption         =   "个性签名："
      Height          =   184
      Left            =   60
      TabIndex        =   4
      Top             =   660
      Width           =   844
   End
   Begin VB.Label LblBirth 
      Caption         =   "生日："
      Height          =   184
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   544
   End
   Begin VB.Label LblName 
      Caption         =   "昵称："
      Height          =   184
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   544
   End
End
Attribute VB_Name = "Sc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()

End Sub

Private Sub LblName_Click()

End Sub
