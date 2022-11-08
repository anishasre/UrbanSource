VERSION 5.00
Begin VB.Form frmRptAccountHeads 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Account Heads Viewer"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleMode       =   0  'User
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   4545
      TabIndex        =   4
      Top             =   0
      Width           =   4545
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   2370
      Width           =   1095
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   375
      Left            =   1260
      TabIndex        =   2
      Top             =   2370
      Width           =   1095
   End
   Begin VB.ComboBox cmbAccountGroups 
      Height          =   315
      ItemData        =   "frmRptAccountHeads.frx":0000
      Left            =   2280
      List            =   "frmRptAccountHeads.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1200
      Width           =   1665
   End
   Begin VB.Label Label1 
      Caption         =   "AccountHeadGroups"
      Height          =   285
      Left            =   660
      TabIndex        =   0
      Top             =   1230
      Width           =   1575
   End
End
Attribute VB_Name = "frmRptAccountHeads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Private Sub cmbAccountGroups_Click()
    cmbAccountGroups.ItemData(ListIndex) = cmbAccountGroups.ItemData(ListIndex)
    End Sub
    Private Sub cmdCancel_Click()
        Unload Me
    End Sub

    Private Sub cmdShow_Click()
        Dim frmNewRpt As New frmRptViewer
        Dim arrInput As Variant
        
        Select Case cmbAccountGroups.Text
                Case Is = "Income"
                    arrInput = Array(1)
                Case Is = "Expenditures"
                    arrInput = Array(2)
                Case Is = "Liabilities"
                    arrInput = Array(3)
                Case Is = "Assets"
                    arrInput = Array(4)
                Case Else
                    arrInput = Array(100)
            End Select
        
        frmNewRpt.WindowState = vbMaximized
        
        frmNewRpt.rptFileName = App.Path & "\Reports\rptAccountHeads.rpt"
        frmNewRpt.InputParameters = arrInput
                
        Call frmNewRpt.ShowReport
        frmNewRpt.Show
    End Sub

    Private Sub Form_Load()
        cmbAccountGroups.AddItem "Income"
        cmbAccountGroups.AddItem "Expenditures"
        cmbAccountGroups.AddItem "Liabilities"
        cmbAccountGroups.AddItem "Assets"
    End Sub
    Private Sub Form_Activate()
        Me.Top = 550
        frmRptAccountHeads.Left = (frmMenu.Width - Me.Width) / 2
    End Sub
