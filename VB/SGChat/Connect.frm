VERSION 5.00
Begin VB.Form Connect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "连接远程主机"
   ClientHeight    =   1916
   ClientLeft      =   12
   ClientTop       =   324
   ClientWidth     =   4328
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1916
   ScaleWidth      =   4328
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdCnt 
      Caption         =   "连接"
      Height          =   304
      Left            =   3000
      TabIndex        =   4
      Top             =   1500
      Width           =   1204
   End
   Begin VB.TextBox TxtIP 
      Height          =   304
      Left            =   60
      TabIndex        =   3
      Top             =   1020
      Width           =   4144
   End
   Begin VB.TextBox TxtName 
      Height          =   304
      Left            =   60
      TabIndex        =   2
      Top             =   300
      Width           =   4144
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助"
      Height          =   304
      Left            =   60
      TabIndex        =   6
      Top             =   1500
      Width           =   1204
   End
   Begin VB.Label LblIP 
      Caption         =   "远程主机IP地址："
      Height          =   304
      Left            =   60
      TabIndex        =   1
      Top             =   780
      Width           =   4324
   End
   Begin VB.Label LblName 
      Caption         =   "远程主机名："
      Height          =   304
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4324
   End
   Begin VB.Label LblStt 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "正在发送请求..."
      Height          =   184
      Left            =   60
      TabIndex        =   5
      Top             =   1620
      Width           =   2704
   End
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

End Sub
