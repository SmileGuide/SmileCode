VERSION 5.00
Begin VB.Form makebadcmprint 
   Caption         =   "Form1"
   ClientHeight    =   1710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3255
   Icon            =   "makebadcmprint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   3255
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFF00&
      Height          =   1095
      Left            =   0
      Picture         =   "makebadcmprint.frx":324A
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "叠被"
      BeginProperty Font 
         Name            =   "华文彩云"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Kunstler Script"
         Size            =   42
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "makebadcmprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
