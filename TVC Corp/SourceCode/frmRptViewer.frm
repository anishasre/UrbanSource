VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRptViewer 
   BackColor       =   &H80000004&
   Caption         =   "~ R e p o r t   V i e w e r ~"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   7470
   WindowState     =   2  'Maximized
   Begin VB.Frame rptFrame 
      BackColor       =   &H80000009&
      Height          =   1575
      Left            =   2100
      TabIndex        =   1
      Top             =   -60
      Width           =   9855
      Begin VB.TextBox txtFund 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6960
         TabIndex        =   17
         Top             =   720
         Width           =   1665
      End
      Begin VB.CheckBox chkMinorAccountHead 
         BackColor       =   &H80000009&
         Caption         =   "Minor Account Head"
         Height          =   255
         Left            =   1425
         TabIndex        =   16
         Top             =   1095
         Width           =   2625
      End
      Begin VB.CommandButton cmdSearchFund 
         Caption         =   "..."
         Height          =   285
         Left            =   8640
         TabIndex        =   15
         Top             =   720
         Width           =   375
      End
      Begin VB.ListBox lstFund 
         BackColor       =   &H80000018&
         Height          =   450
         Left            =   8160
         TabIndex        =   14
         Top             =   180
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.ComboBox cmbAccountGroups 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmRptViewer.frx":0000
         Left            =   6960
         List            =   "frmRptViewer.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "Show Report"
         Height          =   375
         Left            =   6780
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   7980
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtAccountHead 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1410
         TabIndex        =   7
         Top             =   720
         Width           =   3615
      End
      Begin VB.CommandButton cmdSearchAccountHead 
         Caption         =   "..."
         Height          =   285
         Left            =   5070
         TabIndex        =   6
         Top             =   720
         Width           =   375
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   66060291
         CurrentDate     =   39343
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   315
         Left            =   3240
         TabIndex        =   3
         Top             =   240
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   66060291
         CurrentDate     =   39343
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fund Code"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6120
         TabIndex        =   13
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "AccountHeadGroups"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5400
         TabIndex        =   12
         Top             =   270
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Account Head"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   960
         TabIndex        =   5
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3000
         TabIndex        =   4
         Top             =   300
         Width           =   210
      End
   End
   Begin MSComctlLib.TreeView tvwReports 
      Height          =   8640
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   15240
      _Version        =   393217
      HideSelection   =   0   'False
      Style           =   7
      FullRowSelect   =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin CRVIEWER9LibCtl.CRViewer9 crvReport 
      Height          =   9465
      Left            =   2085
      TabIndex        =   18
      Top             =   1530
      Width           =   12840
      lastProp        =   500
      _cx             =   22648
      _cy             =   16695
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "frmRptViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

    Private mvarRptFileName As String
    Private mvarInputParameters As Variant
    Dim mReportFlag As Boolean
    
    Public Function ShowReport() 'rptFileName As String, Optional arrInput As Variant = Null)
            Dim rptFileName As String
            Dim arrInput As Variant
            Set arrInput = Nothing
            Dim Rpt As New CRAXDRT.Report
            Dim mApp As New CRAXDRT.Application
            Dim mLoop As Long
            
            rptFileName = mvarRptFileName
            If IsArray(mvarInputParameters) Then
                arrInput = mvarInputParameters
            End If
            Screen.MousePointer = vbHourglass
            crvReport.DisplayToolbar = True
            'crvReport.Zoom 1
            crvReport.EnableExportButton = True
            Set Rpt = Nothing
            mApp.LogOnServer "ODBC", "dsnFa", "DB_Finance", "FAUser", "FAUser"
            Set Rpt = mApp.OpenReport(mvarRptFileName, 1)
            If IsArray(arrInput) Then
                For mLoop = LBound(arrInput) To UBound(arrInput)
                    Rpt.ParameterFields.Item(mLoop + 1).ClearCurrentValueAndRange
                    Rpt.ParameterFields.Item(mLoop + 1).AddCurrentValue arrInput(mLoop)
                Next mLoop
            End If
            
            crvReport.EnableProgressControl = True
            crvReport.EnableProgressControl = True
            crvReport.DisplayGroupTree = False
            Screen.MousePointer = vbDefault
            crvReport.EnableRefreshButton = False
            crvReport.ReportSource = Rpt
            crvReport.Refresh
            crvReport.Left = 0
            crvReport.Top = 0
            crvReport.Width = Me.Width
            crvReport.Height = Me.Height - 500
            
            mReportFlag = True
            tvwReports.Visible = False
            rptFrame.Visible = False
            crvReport.ViewReport
            
            
    End Function

    Private Sub cmdCancel_Click()
        Unload Me
    End Sub

    Private Sub FormInitialize()
        mReportFlag = False
        txtAccountHead.Text = ""
        txtAccountHead.Tag = ""
        dtpFromDate.Value = gbStartingDate
        dtpToDate.Value = gbEndingDate
        crvReport.Left = 2085
        crvReport.Top = 1530
        tvwReports.Visible = True
        rptFrame.Visible = True
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        If mReportFlag = False Then
            Cancel = True
            Call FormInitialize
        End If
    End Sub

    Private Sub lstFund_dblClick()
        Dim objFund As New clsFund
        objFund.SetFund lstFund.ItemData(lstFund.ListIndex)
        If Not IsNull(objFund.FundID) Then
            txtFund.Tag = objFund.FundID
            txtFund.Text = objFund.FundName
            Set objFund = Nothing
            lstFund.ToolTipText = objFund.FundID
        End If
        lstFund.Visible = False
        txtFund.SetFocus
    End Sub
    
    Private Sub cmdSearchFund_Click()
        Dim mSql As String
        mSql = "Select vchFund, intFundID From faFunds Order By vchFund"
        Call PopulateList(lstFund, mSql, , , , True)
        lstFund.Visible = True
        lstFund.SetFocus
    End Sub

    Private Sub cmdSearchAccountHead_Click()
        'frmSearchAccountHeads.SQLString = ""
        If chkMinorAccountHead.Value = vbChecked Then
            frmSearchAccountHeads.SQLString = "Select (vchMinorAccountHeadCode + '  ' + vchMinorAccountHead) as AccHead, faMinorAccountHeads.intMinorAccountHeadID From faMinorAccountHeads "
        Else
            frmSearchAccountHeads.SQLString = ""
        End If
        frmSearchAccountHeads.Show vbModal
        txtAccountHead.SetFocus
    End Sub

     Private Sub Form_Load()
        Dim newnode As node
        
        'Adding the first Node
        tvwReports.Nodes.Add , , "Reports", "Reports"
       
        Set newnode = tvwReports.Nodes.Add("Reports", tvwChild, "Counter", "Counter Wise")
            Set newnode = tvwReports.Nodes.Add("Counter", tvwChild, "Counter1", "Daily Collection Statements")
            Set newnode = tvwReports.Nodes.Add("Counter", tvwChild, "CashFlow", "Daily Cash Flow")
            
        'Set newnode = tvwReports.Nodes.Add("Counter", tvwChild, "Counter2", "Daily Cancellation Statement")
        Set newnode = tvwReports.Nodes.Add("Reports", tvwChild, "Account", "Statement of Accounts")
            Set newnode = tvwReports.Nodes.Add("Account", tvwChild, "Account1", "Ledger Book")
            Set newnode = tvwReports.Nodes.Add("Account", tvwChild, "Account2", "Cash Book")
            Set newnode = tvwReports.Nodes.Add("Account", tvwChild, "Account3", "Daily Book")
            Set newnode = tvwReports.Nodes.Add("Account", tvwChild, "Account4", "Bank Book")
            Set newnode = tvwReports.Nodes.Add("Account", tvwChild, "Account5", "Journal")
              
        Set newnode = tvwReports.Nodes.Add("Reports", tvwChild, "Register", "Registers")
            Set newnode = tvwReports.Nodes.Add("Register", tvwChild, "Register1", "Cheque Issued Register")
            Set newnode = tvwReports.Nodes.Add("Register", tvwChild, "Register2", "Cheque Received Register")
            'Set newnode = tvwReports.Nodes.Add("Register", tvwChild, "Register3", "Statement of Cheque Received Register")
            'Set newnode = tvwReports.Nodes.Add("Register", tvwChild, "Register4", "Collection Register")
            'Set newnode = tvwReports.Nodes.Add("Register", tvwChild, "Register5", "Memorandum of Collection Register")
            'Set newnode = tvwReports.Nodes.Add("Register", tvwChild, "Register6", "Summary of Daily Collection")
        
        Set newnode = tvwReports.Nodes.Add("Reports", tvwChild, "Masters", "Account Heads")
        
        Set newnode = tvwReports.Nodes.Add("Reports", tvwChild, "FinalAccounts", "Finacial Statements")
            Set newnode = tvwReports.Nodes.Add("FinalAccounts", tvwChild, "TrialBalance", "Trial Balance")
                
            Set newnode = tvwReports.Nodes.Add("FinalAccounts", tvwChild, "IncomeExpenditure", "Income Expenditure Statement")
            Set newnode = tvwReports.Nodes.Add("FinalAccounts", tvwChild, "BalanceSheet", "Balance Sheet")
            Set newnode = tvwReports.Nodes.Add("FinalAccounts", tvwChild, "ReceiptPayment", "Receipt And Payment Statement")
            
        
        'Set Values in Account Group ComboBox
        cmbAccountGroups.AddItem "Income"
        cmbAccountGroups.AddItem "Expenditures"
        cmbAccountGroups.AddItem "Liabilities"
        cmbAccountGroups.AddItem "Assets"
        Call FormInitialize
    End Sub
    Private Sub dtpFromDate_Click()
        dtpFromDate.Value = dtpFromDate.Value
    End Sub
    
    Private Sub dtpToDate_Click()
        dtpToDate.Value = dtpToDate.Value
    End Sub

    Private Sub crvReport_Resize()
        crvReport.Width = ScaleWidth
        crvReport.Height = ScaleHeight
    End Sub

    Public Property Let rptFileName(ByVal vData As String)
        mvarRptFileName = vData
    End Property

    Public Property Let InputParameters(ByVal vData As Variant)
        mvarInputParameters = vData
    End Property

    Private Sub tvwReports_NodeClick(ByVal node As MSComctlLib.node)
        '-----------------------------------'
        '         Added on 29/03/08         '
        '-----------------------------------'
        Dim sSelKey As String
        sSelKey = node.KEY
        Set mvarInputParameters = Nothing
        Select Case node.KEY
            Case "Counter1"     'Daily Collection Statements in counter
                dtpFromDate.Enabled = False
                dtpToDate.Enabled = False
                txtAccountHead.Enabled = False
                cmdSearchAccountHead.Enabled = False
                cmbAccountGroups.Enabled = False
                txtFund.Enabled = False
                cmdSearchFund.Enabled = False
                mvarInputParameters = 0
                mvarRptFileName = App.Path & "\Reports\rptDailyStatementsInCounters.rpt"
            
            Case "Account1"     'Ledger Book Details
                dtpFromDate.Enabled = True
                dtpToDate.Enabled = True
                txtAccountHead.Enabled = True
                cmdSearchAccountHead.Enabled = True
                mvarInputParameters = Array(val(txtAccountHead.Tag), dtpFromDate.Value, dtpToDate.Value)
                mvarRptFileName = App.Path & "\Reports\rptGeneralLedger.rpt"
                                
            Case "Account2"
                dtpFromDate.Enabled = True
                dtpToDate.Enabled = True
                txtAccountHead.Enabled = True
                cmdSearchAccountHead.Enabled = True
                cmbAccountGroups.Enabled = True
                txtFund.Enabled = False
                If chkMinorAccountHead.Value = vbChecked Then
                    mvarInputParameters = Array(val(txtAccountHead.Tag), dtpFromDate.Value, dtpToDate.Value)
                    mvarRptFileName = App.Path & "\Reports\rptCashBookMinorHead.rpt"
                Else
                    mvarInputParameters = Array(val(txtAccountHead.Tag), dtpFromDate.Value, dtpToDate.Value)
                    mvarRptFileName = App.Path & "\Reports\rptCashBook.rpt"
                End If
                      
            Case "Account3"
                dtpFromDate.Enabled = False
                dtpToDate.Enabled = False
                txtAccountHead.Enabled = False
                cmdSearchAccountHead.Enabled = False
                cmbAccountGroups.Enabled = False
                txtFund.Enabled = True
                mvarRptFileName = App.Path & "\Reports\DayBook.rpt"
                
            Case "Account4"
                dtpFromDate.Enabled = True
                dtpToDate.Enabled = True
                txtAccountHead.Enabled = True
                cmdSearchAccountHead.Enabled = True
                cmbAccountGroups.Enabled = True
                cmdSearchFund.Enabled = False
                If chkMinorAccountHead.Value = vbChecked Then
                    mvarInputParameters = Array(val(txtAccountHead.Tag), dtpFromDate.Value, dtpToDate.Value)
                    mvarRptFileName = App.Path & "\Reports\rptBankBookMinorHead.rpt"
                Else
                    mvarInputParameters = Array(val(txtAccountHead.Tag), dtpFromDate.Value, dtpToDate.Value)
                    mvarRptFileName = App.Path & "\Reports\rptBankBook.rpt"
                End If
                
            Case "Account5"
                dtpFromDate.Enabled = True
                dtpToDate.Enabled = True
                txtAccountHead.Enabled = False
                cmdSearchAccountHead.Enabled = False
                cmbAccountGroups.Enabled = False
                cmdSearchFund.Enabled = False
                mvarInputParameters = Array(dtpFromDate.Value, dtpToDate.Value)
                mvarRptFileName = App.Path & "\Reports\rptJournal.rpt"
            
            Case "Account6"
                dtpFromDate.Enabled = True
                dtpToDate.Enabled = True
                txtAccountHead.Enabled = False
                cmdSearchAccountHead.Enabled = False
                cmbAccountGroups.Enabled = False
                cmdSearchFund.Enabled = False
                mvarInputParameters = Array(dtpFromDate.Value, dtpToDate.Value)
                mvarRptFileName = App.Path & "\Reports\rptTrialbalance.rpt"
        
            Case "Register1"    'Cheque Issue Register Details
                dtpFromDate.Enabled = True
                dtpToDate.Enabled = True
                txtAccountHead.Enabled = False
                cmdSearchAccountHead.Enabled = False
                cmbAccountGroups.Enabled = False
                cmdSearchFund.Enabled = True
                txtFund.Enabled = True
                mvarInputParameters = Array(val(txtFund.Tag), dtpFromDate.Value, dtpToDate.Value)
                mvarRptFileName = App.Path & "\Reports\rptChequeIssueRegisterGEN16.rpt"
            
            Case "Register2"    'Cheque Received Register Details
                
                dtpFromDate.Enabled = False
                dtpToDate.Enabled = False
                txtAccountHead.Enabled = False
                cmdSearchAccountHead.Enabled = False
                cmbAccountGroups.Enabled = False
                cmdSearchFund.Enabled = True
                txtFund.Enabled = True
                If Trim(txtFund.Text) = "" Then
                    txtFund.Tag = ""
                    mvarInputParameters = Array(val(txtFund.Tag))
                    mvarRptFileName = App.Path & "\Reports\rptChequeReceivedRegisterGEN9.rpt"
                Else
                    mvarInputParameters = Array(val(txtFund.Tag))
                    mvarRptFileName = App.Path & "\Reports\rptChequeReceivedRegisterGEN9.rpt"
                End If
    
    '            Case "Register3"
    '                dtpFromDate.Enabled = False
    '                dtpToDate.Enabled = False
    '                txtAccountHead.Enabled = False
    '                cmdSearchAccountHead.Enabled = False
    '                cmbAccountGroups.Enabled = False
    '                mvarRptFileName = App.Path & "\Reports\rptStmtOnStatusOfChequesReceivedGEN10.rpt"
    '            Case "Register4"
    '                dtpFromDate.Enabled = False
    '                dtpToDate.Enabled = False
    '                txtAccountHead.Enabled = False
    '                cmdSearchAccountHead.Enabled = False
    '                cmbAccountGroups.Enabled = False
    '                mvarRptFileName = App.Path & "\Reports\rptCollectionRegisterGEN11.rpt"
    '            Case "Register5"
    '                dtpFromDate.Enabled = False
    '                dtpToDate.Enabled = False
    '                txtAccountHead.Enabled = False
    '                cmdSearchAccountHead.Enabled = False
    '                cmbAccountGroups.Enabled = False
    '                mvarRptFileName = App.Path & "\Reports\rptMemorandumofCollectionGEN12.rpt"
    '            Case "Register6"
    '                dtpFromDate.Enabled = False
    '                dtpToDate.Enabled = False
    '                txtAccountHead.Enabled = False
    '                cmdSearchAccountHead.Enabled = False
    '                cmbAccountGroups.Enabled = False
    '                mvarRptFileName = App.Path & "\Reports\rptSummaryofCollectionGEN13.rpt"
            
            
            Case "Masters"
                dtpFromDate.Enabled = False
                dtpToDate.Enabled = False
                cmdSearchAccountHead.Enabled = False
                txtAccountHead.Enabled = False
                cmdSearchFund.Enabled = False
                cmbAccountGroups.Enabled = True
                Select Case cmbAccountGroups.Text
                    Case Is = "Income"
                        mvarInputParameters = Array(1)
                    Case Is = "Expenditures"
                        mvarInputParameters = Array(2)
                    Case Is = "Liabilities"
                        mvarInputParameters = Array(3)
                    Case Is = "Assets"
                        mvarInputParameters = Array(4)
                    Case Else
                        mvarInputParameters = Array(100)
                End Select
    
                mvarRptFileName = App.Path & "\Reports\rptAccountHeads.rpt"
            
         Case "CashFlow"
                dtpFromDate.Enabled = True
                dtpToDate.Enabled = False
                cmdSearchAccountHead.Enabled = False
                txtAccountHead.Enabled = False
                cmdSearchFund.Enabled = False
                cmbAccountGroups.Enabled = False
                mvarInputParameters = Array(dtpFromDate.Value)
                mvarRptFileName = App.Path & "\Reports\DailyCashFlow.rpt"
            
         Case "TrialBalance"
                dtpFromDate.Enabled = True
                dtpToDate.Enabled = True
                cmdSearchAccountHead.Enabled = False
                txtAccountHead.Enabled = False
                cmdSearchFund.Enabled = False
                cmbAccountGroups.Enabled = False
                'if
                mvarInputParameters = Array(dtpFromDate.Value, dtpToDate.Value)
                mvarRptFileName = App.Path & "\Reports\rptTrialbalance.rpt"
         
                
         Case "IncomeExpenditure"
                dtpFromDate.Enabled = True
                dtpToDate.Enabled = True
                cmdSearchAccountHead.Enabled = False
                txtAccountHead.Enabled = False
                cmdSearchFund.Enabled = False
                cmbAccountGroups.Enabled = False
                'if
                mvarInputParameters = Array(dtpFromDate.Value, dtpToDate.Value)
                mvarRptFileName = App.Path & "\Reports\rptIncomeAndExpenditure.rpt"
                                 
         Case "BalanceSheet"
                dtpFromDate.Enabled = True
                dtpToDate.Enabled = False
                cmdSearchAccountHead.Enabled = False
                txtAccountHead.Enabled = False
                cmdSearchFund.Enabled = False
                cmbAccountGroups.Enabled = False
                mvarInputParameters = Array(dtpFromDate.Value)
                mvarRptFileName = App.Path & "\Reports\rptBalanceSheet.rpt"
         Case "ReceiptPayment"
                dtpFromDate.Enabled = True
                dtpToDate.Enabled = True
                cmdSearchAccountHead.Enabled = False
                txtAccountHead.Enabled = False
                cmdSearchFund.Enabled = False
                cmbAccountGroups.Enabled = False
                mvarInputParameters = Array(dtpFromDate.Value, dtpToDate.Value)
                mvarRptFileName = App.Path & "\Reports\rptReceiptsPayments.rpt"
                                    
         End Select
      End Sub
      
      
    Private Sub cmdShow_Click()
        Dim newnode As MSComctlLib.node
        Set newnode = tvwReports.SelectedItem
        If Not tvwReports.SelectedItem Is Nothing Then
            If Not (newnode.KEY = "Reports" Or newnode.KEY = "Counter" _
                Or newnode.KEY = "Account" Or newnode.KEY = "Register" _
                ) Then
                Call tvwReports_NodeClick(newnode)
                ShowReport
            End If
        End If
    End Sub
    
    Private Sub txtAccountHead_GotFocus()
         If gbSearchStr <> "" Then
            txtAccountHead.Text = Token(gbSearchStr, " ")
            txtAccountHead.Text = Trim(gbSearchStr)
            txtAccountHead.Tag = gbSearchID
            gbSearchStr = ""
            gbSearchID = -1
        End If
        txtAccountHead.SelStart = 0
        txtAccountHead.SelLength = Len(txtAccountHead.Text)
    End Sub

