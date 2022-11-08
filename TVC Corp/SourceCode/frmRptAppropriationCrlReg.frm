VERSION 5.00
Begin VB.Form frmRptAppropriationCrlReg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Appropriation Control Register"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbYear 
      Height          =   315
      Left            =   1815
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1095
      Width           =   2325
   End
   Begin VB.ComboBox cmbScheme 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2595
      Width           =   4245
   End
   Begin VB.ComboBox cmbSource 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1785
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1845
      Width           =   4245
   End
   Begin VB.ComboBox cmbCategory 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   1785
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2205
      Width           =   4245
   End
   Begin VB.CommandButton cmdViewAppCtrlReg 
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
      Left            =   5250
      TabIndex        =   2
      Top             =   3030
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "Financial Year"
      Height          =   210
      Left            =   615
      TabIndex        =   7
      Top             =   1110
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Scheme"
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
      Left            =   1035
      TabIndex        =   6
      Top             =   2640
      Width           =   675
   End
   Begin VB.Label lblCategory 
      AutoSize        =   -1  'True
      Caption         =   "Category"
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
      Left            =   960
      TabIndex        =   1
      Top             =   2205
      Width           =   780
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      Caption         =   "Source"
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
      Left            =   1155
      TabIndex        =   0
      Top             =   1845
      Width           =   585
   End
End
Attribute VB_Name = "frmRptAppropriationCrlReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

    Private Sub cmbScheme_GotFocus()
        Dim msql As String
        If cmbSource.ItemData(cmbSource.ListIndex) = 3 Then
            msql = "SELECT  vchDescription,intID  FROM   faDepSchPro WHERE tnyGroupID IN (1,2) ORDER BY vchDescription asc"
            PopulateList cmbScheme, msql, True, True, True, True
        Else
            msql = "SELECT  vchDescription,intID  FROM   faDepSchPro WHERE tnyGroupID IN (3) ORDER BY vchDescription asc"
            PopulateList cmbScheme, msql, True, True, True, True
        End If
    End Sub

    Private Sub cmbSource_Click()
        If cmbSource.ListIndex > 0 Then
            If cmbSource.ItemData(cmbSource.ListIndex) = 29 Then
                cmbCategory.ListIndex = 2
                cmbCategory.Enabled = False
                cmbScheme.Enabled = False
                cmbCategory.Text = "SCP"
            ElseIf cmbSource.ItemData(cmbSource.ListIndex) = 30 Then
                cmbCategory.ListIndex = 3
                cmbCategory.Enabled = False
                cmbScheme.Enabled = False
                cmbCategory.Text = "TSP"
            ElseIf cmbSource.ItemData(cmbSource.ListIndex) = 3 Then ' B-Fund
                cmbCategory.ListIndex = 0
                cmbCategory.Enabled = False
                cmbScheme.Enabled = True
                cmbScheme.SetFocus
            ElseIf cmbSource.ItemData(cmbSource.ListIndex) = 10 Or _
                    cmbSource.ItemData(cmbSource.ListIndex) = 11 Or _
                    cmbSource.ItemData(cmbSource.ListIndex) = 12 Or _
                    cmbSource.ItemData(cmbSource.ListIndex) = 13 Or _
                    cmbSource.ItemData(cmbSource.ListIndex) = 14 Then

                cmbCategory.Enabled = True
                cmbCategory.Text = "GENERAL"
                cmbScheme.ListIndex = 0
                cmbScheme.Enabled = True
          ElseIf cmbSource.ItemData(cmbSource.ListIndex) = 2 Then ' Centrally Sponsored Scheme Fund
                cmbCategory.ListIndex = 1
                cmbCategory.Enabled = False
                cmbCategory.Text = "GENERAL"
                cmbScheme.Enabled = True
                cmbScheme.SetFocus
            Else
                cmbCategory.ListIndex = 1
                cmbCategory.Enabled = False
                cmbCategory.Text = "GENERAL"
                cmbScheme.ListIndex = 0
                cmbScheme.Enabled = False
            End If
        End If
    End Sub
    
    Private Sub cmdViewAppCtrlReg_Click()
        Dim objDb As New clsDB
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        Dim mDepartment As Variant
        Dim mYearID As Integer
        If cmbYear.ListIndex > -1 Then
            mYearID = cmbYear.ItemData(cmbYear.ListIndex)
        Else
            mYearID = gbFinancialYearID
        End If
        
        If cmbSource.ListIndex < 1 Then
            MsgBox "Please select the Source", vbInformation
            cmbSource.SetFocus
            Exit Sub
            arInput = Array(cmbSource.ItemData(cmbSource.ListIndex), 0, 0, mYearID)
        ElseIf cmbSource.ItemData(cmbSource.ListIndex) = 3 Then
            If cmbScheme.ListIndex < 1 Then
                MsgBox "Please select the Scheme", vbInformation
                cmbScheme.SetFocus
                Exit Sub
            Else
                arInput = Array(cmbSource.ItemData(cmbSource.ListIndex), 0, cmbScheme.ItemData(cmbScheme.ListIndex), mYearID)
            End If
        ElseIf cmbSource.ItemData(cmbSource.ListIndex) = 2 Then
            If cmbScheme.ListIndex < 1 Then
                'MsgBox "Please select the Scheme", vbInformation
                'cmbScheme.SetFocus
                'Exit Sub
                arInput = Array(cmbSource.ItemData(cmbSource.ListIndex), cmbCategory.ItemData(cmbCategory.ListIndex), 0, mYearID)
            
            Else
                arInput = Array(cmbSource.ItemData(cmbSource.ListIndex), cmbCategory.ItemData(cmbCategory.ListIndex), cmbScheme.ItemData(cmbScheme.ListIndex), mYearID)
            End If
        Else
           arInput = Array(cmbSource.ItemData(cmbSource.ListIndex), cmbCategory.ItemData(cmbCategory.ListIndex), 0, mYearID)
        End If
        frmNewViewer.rptFileName = App.Path & "\Reports\rptAppropriationControlRegisterGEN-39.rpt"
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Print
        'frmNewViewer.Show
        'Set frmNewViewer = Nothing
        'Set frmNewRpt = Nothing
        
        
             
    End Sub

    Private Sub Form_Activate()
        Me.Left = 0
        Me.Top = 0
    End Sub

    Private Sub Form_Load()
        Dim msql As String
        
        If gbLBPanchayat = 1 Then
            msql = "Select vchSourceFundName,intSourceFundID From suSourceOfFund Where intSourceFundID In(1,2,3,4,16,17,25,26,27,28,10,11,12,13,14,19,21,29,30,41) Order By vchSourceFundName"
        Else
            msql = "Select vchSourceFundName,intSourceFundID From suSourceOfFund Where intSourceFundID In(1,2,3,4,16,17,19,21,25,26,27,28,29,30,41) Order By vchSourceFundName"
        End If
        PopulateList cmbSource, msql, , True, True, True, enuSourceString.Saankhya
        
        msql = "SELECT vchTransactionCategory,intCategoryID FROM faTransactionCategory"
        PopulateList cmbCategory, msql, True, True, True, True
        
        msql = "SELECT vchDescription ,intID  FROM  faDepSchPro Order By tnyGroupID"
        PopulateList cmbScheme, msql, True, True, True, True
        
        On Error Resume Next
        msql = "Select LTRIM(Str(intFinancialYear)) + '-' + LTRIM(Str(intFinancialYear+1)), intFinancialYearID  From faFinancialYear"
        PopulateList cmbYear, msql, True, False, True, True
        cmbYear.Text = Trim(str(gbFinancialYearID)) + "-" + Trim(str((gbFinancialYearID + 1)))
    End Sub

