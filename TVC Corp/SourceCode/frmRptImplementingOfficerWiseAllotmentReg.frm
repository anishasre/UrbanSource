VERSION 5.00
Begin VB.Form frmRptImplementingOfficerWiseAllotmentReg 
   Caption         =   "Implementing Officer Wise Allotment Register"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5415
   Begin VB.TextBox txtImplementingOfficer 
      Height          =   300
      Left            =   1935
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   555
      Width           =   2640
   End
   Begin VB.CommandButton cmdImlementingOfficer 
      Caption         =   "..."
      Height          =   300
      Left            =   4590
      TabIndex        =   2
      Top             =   555
      Width           =   375
   End
   Begin VB.CommandButton cmdViewImpOffWiseAlltReg 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4245
      TabIndex        =   1
      Top             =   1230
      Width           =   795
   End
   Begin VB.Label lblImpementingOfficer 
      AutoSize        =   -1  'True
      Caption         =   "Impementing Officer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   75
      TabIndex        =   0
      Top             =   585
      Width           =   1830
   End
End
Attribute VB_Name = "frmRptImplementingOfficerWiseAllotmentReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Private Sub cmdImlementingOfficer_Click()
        gbSearchID = -1                                         ''  Setting the Search ID to -1
        frmSearchSubsidiaryAccountHeads.SubLedgerType = 1       ''  1. Implementing Officer
        frmSearchSubsidiaryAccountHeads.Show vbModal
        txtImplementingOfficer.SetFocus
    End Sub

    Private Sub cmdViewImpOffWiseAlltReg_Click()
        Dim objDb As New clsDb
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        Dim mDepartment As Variant
        
        If txtImplementingOfficer.Tag <> "" Then
            arInput = Array(txtImplementingOfficer.Tag)
        Else
            MsgBox "Please select the ImplementingOfficer", vbInformation
            cmdImlementingOfficer.SetFocus
            Exit Sub
        End If
        'arInput = Array(dtpDate.Value, "%", "%", "%", "%")
        frmNewViewer.rptFileName = App.Path & "\Reports\rptImplementingOfficerWiseAllotment.rpt"
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
    End Sub

    Private Sub Form_Activate()
        Me.Width = 5535
        Me.Height = 2475
        Me.Top = 0
        Me.Left = 0
    End Sub
    
    Private Sub txtImplementingOfficer_GotFocus()
         If gbSearchID > 0 Then
            Dim objSubLedger As New clsSubLedger
            objSubLedger.SetSubLedgerDetails (gbSearchID)
            txtImplementingOfficer.Tag = objSubLedger.SubsidiaryAccountHeadID
            txtImplementingOfficer.Text = objSubLedger.NameOfSubLedger
        End If
        gbSearchID = -1
    End Sub
