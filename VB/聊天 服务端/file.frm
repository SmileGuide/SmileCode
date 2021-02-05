VERSION 5.00
Begin VB.Form file 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   5895
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.DirListBox Dir1 
      Height          =   1560
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      Height          =   1530
      Left            =   3240
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "file"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
File1.FileName = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub
