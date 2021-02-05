VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H80000005&
   Caption         =   "聊天"
   ClientHeight    =   4148
   ClientLeft      =   44
   ClientTop       =   628
   ClientWidth     =   7400
   LinkTopic       =   "Form1"
   ScaleHeight     =   4148
   ScaleWidth      =   7400
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu MnuTm 
      Caption         =   "现在时间 00:00"
   End
   Begin VB.Menu MnuTmL 
      Caption         =   "尚能聊天 0分钟"
      Visible         =   0   'False
   End
   Begin VB.Menu MnuW1 
      Caption         =   ""
   End
   Begin VB.Menu Sc 
      Caption         =   "个性"
   End
   Begin VB.Menu MnuBc 
      Caption         =   "背景"
      Begin VB.Menu MnuTrans 
         Caption         =   "修改前方透明度"
      End
   End
   Begin VB.Menu MnuFun 
      Caption         =   "多功能"
      Begin VB.Menu MnuTmr 
         Caption         =   "聊天限时"
      End
      Begin VB.Menu MnuRpl 
         Caption         =   "自动回复"
      End
   End
   Begin VB.Menu MnuFrm 
      Caption         =   "窗口"
      Begin VB.Menu MnuM 
         Caption         =   "多开..."
      End
      Begin VB.Menu MnuW21 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MnuXt 
      Caption         =   "退出"
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MnuTm_Click()

End Sub

Private Sub MnuTmL_Click()

End Sub

Private Sub MnuW1_Click()

End Sub
