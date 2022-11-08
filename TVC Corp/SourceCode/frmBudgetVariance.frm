VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBudgetVariance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Budget Variance & Analysis"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10155
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   10155
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fmeBudgetUtlisation 
      Height          =   3630
      Left            =   0
      TabIndex        =   5
      Top             =   90
      Width           =   6495
      Begin VB.CommandButton cmdFunction 
         Caption         =   "..."
         Height          =   375
         Left            =   2760
         TabIndex        =   29
         Top             =   1448
         Width           =   330
      End
      Begin VB.CommandButton cmdFunctionary 
         Caption         =   "..."
         Height          =   375
         Left            =   6120
         TabIndex        =   28
         Top             =   1448
         Width           =   330
      End
      Begin VB.ComboBox cmbFund 
         Height          =   390
         Left            =   4245
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   945
         Width           =   1725
      End
      Begin VB.TextBox txtFunctionary 
         Enabled         =   0   'False
         Height          =   390
         Left            =   4245
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1440
         Width           =   1725
      End
      Begin VB.ComboBox cmbType 
         Height          =   390
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   945
         Width           =   1725
      End
      Begin VB.TextBox txtFunction 
         Enabled         =   0   'False
         Height          =   390
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1440
         Width           =   1725
      End
      Begin VB.TextBox txtFinancialYear 
         Height          =   390
         Left            =   1395
         TabIndex        =   20
         Top             =   2745
         Width           =   2535
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   375
         Left            =   5040
         TabIndex        =   19
         Top             =   1935
         Width           =   330
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Show"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4140
         TabIndex        =   18
         Top             =   2700
         Width           =   1140
      End
      Begin VB.ComboBox cmbAccountHeads 
         Height          =   390
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1935
         Width           =   3750
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   390
         Left            =   960
         TabIndex        =   24
         Top             =   540
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   688
         _Version        =   393216
         Format          =   15925249
         CurrentDate     =   40044
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   390
         Left            =   4245
         TabIndex        =   27
         Top             =   540
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   688
         _Version        =   393216
         Format          =   15925249
         CurrentDate     =   40044
      End
      Begin VB.Label Label2 
         Caption         =   "Financial Year"
         Height          =   240
         Left            =   135
         TabIndex        =   21
         Top             =   2835
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Account Head"
         Height          =   270
         Index           =   6
         Left            =   45
         TabIndex        =   17
         Top             =   1980
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Functionary"
         Height          =   270
         Index           =   5
         Left            =   3210
         TabIndex        =   11
         Top             =   1440
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Function"
         Height          =   270
         Index           =   4
         Left            =   45
         TabIndex        =   10
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fund"
         Height          =   270
         Index           =   3
         Left            =   3750
         TabIndex        =   9
         Top             =   990
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         Height          =   270
         Index           =   2
         Left            =   45
         TabIndex        =   8
         Top             =   585
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         Height          =   270
         Index           =   1
         Left            =   3615
         TabIndex        =   7
         Top             =   585
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Type"
         Height          =   270
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   990
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3630
      Left            =   6510
      TabIndex        =   0
      Top             =   90
      Width           =   3645
      Begin VB.OptionButton optBudgetUtilisation 
         Caption         =   "Cash && Bank Position"
         Height          =   210
         Index           =   7
         Left            =   90
         TabIndex        =   15
         Top             =   2970
         Width           =   4110
      End
      Begin VB.OptionButton optBudgetUtilisation 
         Caption         =   "Function - Group wise Budget Variance"
         Height          =   435
         Index           =   6
         Left            =   90
         TabIndex        =   14
         Top             =   2340
         Width           =   4110
      End
      Begin VB.OptionButton optBudgetUtilisation 
         Caption         =   "Function - Budget Variance"
         Height          =   210
         Index           =   5
         Left            =   90
         TabIndex        =   13
         Top             =   1980
         Width           =   4110
      End
      Begin VB.OptionButton optBudgetUtilisation 
         Caption         =   "BudgetUtilisation StateMent"
         Height          =   210
         Index           =   4
         Left            =   90
         TabIndex        =   12
         Top             =   1575
         Width           =   4110
      End
      Begin VB.OptionButton optBudgetUtilisation 
         Caption         =   "Quarterly Budget Variance"
         Height          =   210
         Index           =   3
         Left            =   90
         TabIndex        =   4
         Top             =   1215
         Width           =   4110
      End
      Begin VB.OptionButton optBudgetUtilisation 
         Caption         =   "Performance Statement"
         Height          =   210
         Index           =   2
         Left            =   90
         TabIndex        =   3
         Top             =   900
         Width           =   4110
      End
      Begin VB.OptionButton optBudgetUtilisation 
         Caption         =   "Budget Variance"
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   2
         Top             =   585
         Width           =   4110
      End
      Begin VB.OptionButton optBudgetUtilisation 
         Caption         =   "Budget Utilsation"
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Width           =   4110
      End
   End
End
Attribute VB_Name = "frmBudgetVariance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()
    frmSearchAccountHeads.SQLString = "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where intGroupID in(1,2) Order By vchAccountHeadCode"
    frmSearchAccountHeads.Show vbModal
    If gbSearchStr <> "" Then
        cmbAccountHeads.Text = gbSearchStr
        gbSearchStr = ""
    End If
End Sub

Private Sub cmdFunction_Click()           'Done on 25/11/2009
    Dim mTokenV As String
    frmSearchFunction.Show vbModal
    mTokenV = Token(gbSearchStr, " ")
    txtFunction.Text = Trim(gbSearchStr)
    txtFunction.Tag = gbSearchID
    gbSearchStr = " "
    gbSearchID = -1
            
End Sub

Private Sub cmdFunctionary_Click()         'Done on 25/11/2009
    Dim mTokenV As String
    frmSearchFunctionary.Show vbModal
    mTokenV = Token(gbSearchStr, " ")
    txtFunctionary.Text = Trim(gbSearchStr)
    txtFunctionary.Tag = gbSearchID
    gbSearchStr = " "
    gbSearchID = -1
End Sub

Private Sub cmdShow_Click()
    Dim arInput As Variant
    Dim mRptName As String
    Dim mctrls As Integer
    Dim mSelCtrlIndex As Integer
    Dim mFromDate As Date
    Dim mYear As String
    mYear = IIf(Month(dtpTo.Value) > 3, Year(dtpTo.Value), Year(dtpTo.Value) + 1)
    mFromDate = "01/Apr/" + mYear
    For mctrls = 0 To optBudgetUtilisation.Count - 1
        If optBudgetUtilisation(mctrls).Value Then
            mSelCtrlIndex = mctrls
        End If
    Next
    If Len(Trim(txtFinancialYear.Text)) <> 4 Or IsNumeric(Trim(txtFinancialYear.Text)) = False Then
        MsgBox "Please Check the Financial Year Entered"
        txtFinancialYear.SetFocus
        Exit Sub
    End If
    Select Case mSelCtrlIndex
    Case 0:
        If cmbType.ListIndex < 1 Then
            MsgBox "The Type Must be selected"
            cmbType.SetFocus
            Exit Sub
        End If
        arInput = Array(CStr(cmbType.ItemData(cmbType.ListIndex)), "%", "%", Trim(txtFinancialYear.Text), mFromDate, dtpTo.Value)
        mRptName = "rptBudgetAccountHeadwise.rpt"
    Case 1:
        If cmbType.ListIndex < 1 Then
            MsgBox "The Type Must be selected"
            cmbType.SetFocus
            Exit Sub
        End If
        If cmbFund.ListIndex < 1 Then
            MsgBox "Fund Must be Selected"
            cmbFund.SetFocus
            Exit Sub
        End If
        If txtFunction.Tag = "" Then            'Modified 25/11/2009
            MsgBox "Function Must be Selected"
            txtFunction.SetFocus
            Exit Sub
        End If
        If txtFunctionary.Tag = "" Then         'Modified 25/11/2009
            MsgBox "Functionary Must be Selected"
            txtFunctionary.SetFocus
            Exit Sub
        End If
        
        
        'If cmbFunction.ListIndex < 1 Then
           ' MsgBox "Function Must be Selected"
            'cmbFunction.SetFocus
           ' Exit Sub
        'End If
        'If cmbFunctionary.ListIndex < 1 Then
            'MsgBox "Functionary Must be Selected"
           ' cmbFunctionary.SetFocus
            'Exit Sub
        'End If
        
     '*********** Modified on 25/11/2009 ***********
        
        arInput = Array(CStr(cmbType.ItemData(cmbType.ListIndex)), CStr(txtFunction.Tag), CStr(txtFunctionary.Tag), CStr(txtFinancialYear.Text), dtpFrom.Value, dtpTo.Value)
        'arInput = Array(CStr(cmbType.ItemData(cmbType.ListIndex)), CStr(cmbFunction.ItemData(cmbFunction.ListIndex)), CStr(cmbFunctionary.ItemData(cmbFunctionary.ListIndex)), Trim(txtFinancialYear.Text), dtpFrom.Value, dtpTo.Value)
        mRptName = "rptBudgetFunctionFunctionaryWise.rpt"
    Case 2:
        arInput = Array("%", "%", "%", Trim(txtFinancialYear.Text), mFromDate, dtpTo.Value)
        mRptName = "rptPerformanceStatement.rpt"
    Case 3:
        arInput = Array("%", "%", "%", Trim(txtFinancialYear.Text), mFromDate, dtpTo.Value)
        mRptName = "rptBudgetQuaterly.rpt"
    Case 4:
        If cmbType.ListIndex < 1 Then
            MsgBox "The Type Must be selected"
            cmbType.SetFocus
            Exit Sub
        End If
        arInput = Array(CStr(cmbType.ItemData(cmbType.ListIndex)), "%", "%", Trim(txtFinancialYear.Text), mFromDate, dtpTo.Value)
        mRptName = "rptBudgetFunctionWiseMajorHeadwise.rpt"
    Case 5:
        If cmbType.ListIndex < 1 Then
            MsgBox "The Type Must be selected"
            cmbType.SetFocus
            Exit Sub
        End If
        arInput = Array(CStr(cmbType.ItemData(cmbType.ListIndex)), "%", "%", Trim(txtFinancialYear.Text), dtpFrom.Value, dtpTo.Value)
        mRptName = "rptBudgetMajorFunctionWise.rpt"
    Case 6:
        If cmbType.ListIndex < 1 Then
            MsgBox "The Type Must be selected"
            cmbType.SetFocus
            Exit Sub
        End If
        arInput = Array(CStr(cmbType.ItemData(cmbType.ListIndex)), "%", "%", Trim(txtFinancialYear.Text), dtpFrom.Value, dtpTo.Value)
        mRptName = "rptBudgetMajorFunctions.rpt"
    Case 7:
        If cmbAccountHeads.ListIndex < 1 Then
            MsgBox "The Account Head Must be selected"
            cmdBrowse.SetFocus
            Exit Sub
        End If
        arInput = Array(cmbAccountHeads.ItemData(cmbAccountHeads.ListIndex), dtpTo.Value, dtpTo.Value)
        mRptName = "rptCashBankPossition.rpt"
    End Select
    Call ShowReport(arInput, mRptName)
End Sub

Private Sub ShowReport(arInput As Variant, mreportName As String)
    Dim frmNewViewer As New frmRptViewer
    frmNewViewer.rptFileName = App.Path & "\Reports\" & mreportName
    frmNewViewer.WindowState = vbMaximized
    frmNewViewer.WindowState = vbMaximized
    frmNewViewer.InputParameters = arInput
    Call frmNewViewer.ShowReport
    frmNewViewer.Show
End Sub

Private Sub Form_Load()
    Call InitForm
    dtpFrom.Value = gbStartingDate
    dtpTo.Value = gbTransactionDate
    optBudgetUtilisation(0).Value = True
    txtFinancialYear.Text = IIf(Month(Date) > 3, Year(Date), Year(Date) + 1)
End Sub

Private Sub InitForm()
    ''''Tin Type Combo''''''''
    cmbType.Clear
    cmbType.AddItem ""
    cmbType.AddItem ("Income")
    cmbType.ItemData(cmbType.NewIndex) = 1
    cmbType.AddItem ("Expenditure")
    cmbType.ItemData(cmbType.NewIndex) = 2
    cmbType.AddItem ("Liability")
    cmbType.ItemData(cmbType.NewIndex) = 3
    cmbType.AddItem ("Asset")
    cmbType.ItemData(cmbType.NewIndex) = 4
    '''Fund, Account Head Function and Functionary'''''
    'PopulateList cmbFunction, "Select vchFunction,intFunctionID From faFunctions Order By vchFunction", , True, , True
    'PopulateList cmbFunctionary, "Select vchFunctionary,intFunctionaryID From faFunctionaries Order By vchFunctionary", , True, , True
    PopulateList cmbFund, "Select vchFund,intFundID From faFunds Order By vchFund", "General Fund", True, , True
    PopulateList cmbAccountHeads, "Select (vchAccountHeadCode + '  ' + vchAccountHead) as AccHead, intAccountHeadID  From faAccountHeads Where intGroupID in(1,2) Order By vchAccountHeadCode", , True, , True
End Sub

Private Sub MakeDisable()
    cmbAccountHeads.Enabled = False
    txtFunction.Enabled = False
    txtFunctionary.Enabled = False
    cmdFunction.Enabled = False
    cmdFunctionary.Enabled = False
    'cmbFunction.Enabled = False
    'cmbFunctionary.Enabled = False
    cmbType.Enabled = False
    dtpFrom.Enabled = False
    dtpTo.Enabled = False
End Sub

Private Sub optBudgetUtilisation_Click(Index As Integer)
    Call MakeDisable
    Select Case Index
        Case 0, 4:  cmbType.Enabled = True
                    dtpTo.Enabled = True
        Case 1:     cmbType.Enabled = True
                    dtpFrom.Enabled = True
                    txtFunction.Enabled = True
                    txtFunctionary.Enabled = True
                    cmdFunction.Enabled = True
                    cmdFunctionary.Enabled = True
                    'cmbFunction.Enabled = True
                    'cmbFunctionary.Enabled = True
                    dtpTo.Enabled = True
        Case 2, 3:  dtpTo.Enabled = True
        Case 5, 6:  cmbType.Enabled = True
                    dtpFrom.Enabled = True
                    dtpTo.Enabled = True
        Case 7:     cmbAccountHeads.Enabled = True
                    dtpFrom.Enabled = False
                    dtpTo.Enabled = True
    End Select
End Sub
