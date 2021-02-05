VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Reply 
   Caption         =   "Form1"
   ClientHeight    =   1916
   ClientLeft      =   16
   ClientTop       =   328
   ClientWidth     =   5548
   LinkTopic       =   "Form1"
   ScaleHeight     =   1916
   ScaleWidth      =   5548
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   1444
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   4564
      _ExtentX        =   8051
      _ExtentY        =   2548
      TabWidthStyle   =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Reply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub TabStrip1_Click()
TabStrip1.SelectedItem
End Sub
