VERSION 5.00
Begin VB.Form frmTeamSaankhya 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Developer's Centre"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEnableMenu 
      Caption         =   "Enable All Menus"
      Height          =   420
      Left            =   300
      TabIndex        =   0
      Top             =   405
      Width           =   1620
   End
End
Attribute VB_Name = "frmTeamSaankhya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdEnableMenu_Click()
    For Each mMenu In frmMenu.Controls
        If TypeOf mMenu Is Menu Then
            mMenu.Enabled = True
            mMenu.Visible = True
        End If
    Next
End Sub
