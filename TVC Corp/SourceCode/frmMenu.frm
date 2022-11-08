VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00A65D3E&
   Caption         =   "  S a a n k h y a"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   11880
      TabIndex        =   1
      Top             =   0
      Width           =   11880
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   20160
         Top             =   120
      End
      Begin VB.Label lblDBVersion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DbVersion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   10620
         TabIndex        =   13
         Top             =   405
         Width           =   1005
      End
      Begin VB.Label lblSplash 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   21.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   480
         Left            =   5625
         TabIndex        =   12
         Top             =   180
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Image imgWarning 
         Height          =   240
         Left            =   12240
         Picture         =   "frmMenu.frx":1CCA
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label lblPreYear 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   12525
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   7635
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ver 2.2.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   9630
         TabIndex        =   10
         Top             =   405
         Width           =   720
      End
      Begin VB.Label lblTransactionDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<31-Mar-2008>"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   7665
         TabIndex        =   9
         Top             =   285
         Width           =   1125
      End
      Begin VB.Label lblFinancialYear 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<2007-2008>"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   7665
         TabIndex        =   8
         Top             =   60
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   6195
         TabIndex        =   7
         Top             =   285
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Financial Year :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   6435
         TabIndex        =   6
         Top             =   45
         Width           =   1170
      End
      Begin VB.Label lblCounter 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Counter Name>"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   1020
         TabIndex        =   5
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label lblLoginName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Login Name>"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   1020
         TabIndex        =   4
         Top             =   75
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Counter :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Login :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   450
         TabIndex        =   2
         Top             =   60
         Width           =   510
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   8970
         Picture         =   "frmMenu.frx":281B
         Top             =   0
         Width           =   3000
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   5340
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu Administration 
      Caption         =   "Administration"
      Begin VB.Menu AccountHeads 
         Caption         =   "Opening Balances"
      End
      Begin VB.Menu OpeningCashBook 
         Caption         =   "Opening Cash Book"
      End
      Begin VB.Menu ClosingAccounts 
         Caption         =   "Closing Accounts"
      End
      Begin VB.Menu PublishingUtility 
         Caption         =   "Publishing Utility"
      End
      Begin VB.Menu RequisitionInbox 
         Caption         =   "Requisition Inbox"
      End
      Begin VB.Menu SubLedger 
         Caption         =   "Sub Ledger"
         Begin VB.Menu ImplementingOfficer 
            Caption         =   "Implementing Officer"
         End
         Begin VB.Menu AllotmentClosingBalance 
            Caption         =   "AllotmentClosing Balance"
         End
      End
      Begin VB.Menu SubsidiaryAccount 
         Caption         =   "SubsidiaryAccount"
      End
      Begin VB.Menu FunctionaryHead 
         Caption         =   "Allocate Account heads to Functionaries"
      End
      Begin VB.Menu Bank 
         Caption         =   "Bank Accounts"
      End
      Begin VB.Menu GST 
         Caption         =   "GST"
      End
      Begin VB.Menu AFS 
         Caption         =   "AFS Processing"
      End
      Begin VB.Menu ChequeBook 
         Caption         =   "Cheque Book Details"
      End
      Begin VB.Menu OpeningVoucherEntry 
         Caption         =   "Opening Voucher  Entry"
      End
      Begin VB.Menu Funds 
         Caption         =   "Funds"
      End
      Begin VB.Menu Functions 
         Caption         =   "Functions"
      End
      Begin VB.Menu Functionaries 
         Caption         =   "Functionaries"
      End
      Begin VB.Menu Fields 
         Caption         =   "Fields"
      End
      Begin VB.Menu CollectionRegister 
         Caption         =   "Collection Register"
      End
      Begin VB.Menu StockRegisterReceiptBooks 
         Caption         =   "Stock Register of Receipt Books"
      End
      Begin VB.Menu DefineReportSchedules 
         Caption         =   "Report Schedules "
      End
      Begin VB.Menu BudgetCentres 
         Caption         =   "Budget Centres"
      End
      Begin VB.Menu BudgetRevision 
         Caption         =   "Budget Revision"
      End
      Begin VB.Menu OpeningBalanceSheet 
         Caption         =   "Opening Balance Sheet"
      End
      Begin VB.Menu UpdateOpeningBalance 
         Caption         =   "Update Opening Balance"
      End
      Begin VB.Menu SectionWiseTransactionTypes 
         Caption         =   "Section wise Transaction Types"
      End
      Begin VB.Menu SynchronizeProjectMaster 
         Caption         =   "Synchronize Project Master"
      End
      Begin VB.Menu ConfigurationSettings 
         Caption         =   "Default Configuration"
      End
      Begin VB.Menu SourceOfFundEntry 
         Caption         =   "Source Of Fund Entry"
      End
      Begin VB.Menu Proceedings 
         Caption         =   "Proceedings"
      End
      Begin VB.Menu YearEndProcedure 
         Caption         =   "Year End Procedure"
      End
      Begin VB.Menu CheckOutAccountingAssistant 
         Caption         =   "Check Out For  Accounting Assistant"
      End
      Begin VB.Menu ExtractDataForBudget 
         Caption         =   "Extract Data For Budget"
      End
   End
   Begin VB.Menu Transactions 
      Caption         =   "Transactions"
      Begin VB.Menu CounterReceipt 
         Caption         =   "Receipts"
      End
      Begin VB.Menu Payments 
         Caption         =   "Payments"
      End
      Begin VB.Menu PaymentOrder 
         Caption         =   "Payment Order"
      End
      Begin VB.Menu JournalEntry 
         Caption         =   "Journal Entry"
      End
      Begin VB.Menu ContraEntry 
         Caption         =   "Contra Entry"
      End
      Begin VB.Menu BankReconciliationEntry 
         Caption         =   "Bank Reconciliation Entry"
      End
      Begin VB.Menu BankReconcile 
         Caption         =   "Bank Reconcile"
      End
      Begin VB.Menu ebill 
         Caption         =   "E-bill Voucher genaration"
         Index           =   20
      End
      Begin VB.Menu DecimalEntry 
         Caption         =   "DecimalEntry"
      End
      Begin VB.Menu OpeningAppropriationControlDetails 
         Caption         =   "Opening Appropriation Control Details"
         Begin VB.Menu OpeningLetterOfAuthority 
            Caption         =   "Opening Letter Of Authority"
         End
         Begin VB.Menu OpeningLetterOfAllotment 
            Caption         =   "Opening Letter Of Allotment"
         End
      End
      Begin VB.Menu Allotments 
         Caption         =   "Letter of Authority/Allotments"
         Begin VB.Menu LetterOfAuthority 
            Caption         =   "Letter of Authority"
         End
         Begin VB.Menu LetterOfAllotment 
            Caption         =   "Letter Of Allotment"
         End
         Begin VB.Menu RequisitionForFund 
            Caption         =   "Requisition For Fund By Implementing Officers"
         End
         Begin VB.Menu ListOfCancelledRequisitions 
            Caption         =   "List Of Cancelled Requisitions"
         End
         Begin VB.Menu RequistionRegister 
            Caption         =   "Requistion Register"
         End
         Begin VB.Menu ProjectRegister 
            Caption         =   "Project Register"
         End
         Begin VB.Menu SourceDeductions 
            Caption         =   "Deductions from Letter Of Authority"
         End
         Begin VB.Menu BeneficiaryContribution 
            Caption         =   "Direct Expenditure "
         End
         Begin VB.Menu RemitBackOfUnUtilizedAmount 
            Caption         =   "RemitBack Of UnUtilized Amount"
         End
         Begin VB.Menu LinkRecoveriesToProjectExpenditure 
            Caption         =   "Link Recoveries to Project Expenditure"
         End
         Begin VB.Menu UnAuthorizedDrawal 
            Caption         =   "UnAuthorized Drawal"
         End
         Begin VB.Menu LinkAllotmentsWithReceipts 
            Caption         =   "Link Allotments With Receipts"
         End
      End
      Begin VB.Menu Agreement 
         Caption         =   "Agreement"
      End
      Begin VB.Menu SearchPaymentOrder 
         Caption         =   "Search PaymentOrder"
      End
   End
   Begin VB.Menu Utilities 
      Caption         =   "Utilities"
      Begin VB.Menu PaymentRegister 
         Caption         =   "Payment Register"
      End
      Begin VB.Menu DemandInterface 
         Caption         =   "Demand Interface"
      End
      Begin VB.Menu DemandRegister 
         Caption         =   "Demand Register"
      End
      Begin VB.Menu PayBill 
         Caption         =   "Pay Bill"
      End
      Begin VB.Menu BankReconUtility 
         Caption         =   "BankReconciliationUtility"
      End
      Begin VB.Menu DemandInbox 
         Caption         =   "Demand Inbox"
      End
      Begin VB.Menu ZonalIntegration 
         Caption         =   "Zonal Office(s)"
      End
      Begin VB.Menu RequestForReceiptCancellation 
         Caption         =   "Request For Receipt Cancellation"
      End
      Begin VB.Menu ApprovalOfRecieptCancellation 
         Caption         =   "Approval of Reciept Cancellation"
      End
      Begin VB.Menu CancelPaymentOrder 
         Caption         =   "Cancel Payment Order"
      End
      Begin VB.Menu ReverseEntry 
         Caption         =   "Reverse Entry"
         Begin VB.Menu ReverseEntryList 
            Caption         =   "Reverse Entry List"
         End
         Begin VB.Menu ChequeReturnList 
            Caption         =   "Cheque Return List"
         End
      End
      Begin VB.Menu SearchTransactions 
         Caption         =   "Search Transactions"
      End
      Begin VB.Menu InterruptedReceipt 
         Caption         =   "Interrupted Receipt"
         Begin VB.Menu RequestforInterruptedReceipt 
            Caption         =   "Request For Interrupted Receipt"
         End
         Begin VB.Menu RequestforInterruptedReceiptCancellation 
            Caption         =   "Request For Interrupted Receipt Cancellation"
            Visible         =   0   'False
         End
         Begin VB.Menu ApprovalofInterruptedReceiptCancellation 
            Caption         =   "Approval of Interrupted Receipt Cancellation"
            Visible         =   0   'False
         End
         Begin VB.Menu RequestForInterruptedRecEdit 
            Caption         =   "Request For Interrupted Receipt Edit"
            Visible         =   0   'False
         End
         Begin VB.Menu RequestForInterruptedReceiptDateEdit 
            Caption         =   "Request For Interrupted Receipt Date Edit"
            Visible         =   0   'False
         End
         Begin VB.Menu IssueOfInterruptedReceiptBook 
            Caption         =   "Issue of Interrupted Receipt Book"
         End
         Begin VB.Menu InterruptedRegister 
            Caption         =   "Interrupted Register"
         End
      End
      Begin VB.Menu TreasuryBalanceFinalization 
         Caption         =   "Treasury Balance Finalization"
      End
      Begin VB.Menu AFSFinalization 
         Caption         =   "ACR-Finalization "
         Begin VB.Menu AFSSourceOfFund 
            Caption         =   "ACR-Closing Source Of Fund"
         End
      End
      Begin VB.Menu PortPanchayatData 
         Caption         =   "Port Panchayat Data"
      End
      Begin VB.Menu AccrualDemand 
         Caption         =   "Accrued List of Items"
      End
      Begin VB.Menu InwardChecksAndDds 
         Caption         =   "Inward Checks/DDs From Soochika"
      End
      Begin VB.Menu SearchBuildingTaxRemitance 
         Caption         =   "Search Building Tax Remitance"
      End
      Begin VB.Menu SearchReceipts 
         Caption         =   "Search Receipts Issued"
      End
      Begin VB.Menu SearchAgreements 
         Caption         =   "Search Agreements"
      End
      Begin VB.Menu ChequeRegisterUtility 
         Caption         =   "Cheque Register"
      End
      Begin VB.Menu SendDailyCollectionToHO 
         Caption         =   "Send Daily Collection To H.O."
      End
      Begin VB.Menu SubsidiaryCashBook 
         Caption         =   "Subsidiary Cash Book"
      End
      Begin VB.Menu ListOfPreviousDateReceiptCancellation 
         Caption         =   "List Of PreviousDate's Receipt Cancellation"
      End
      Begin VB.Menu RequestforPendingTransactions 
         Caption         =   "Request for Enable Previous Year's Transaction"
      End
      Begin VB.Menu SynchronizeDetails 
         Caption         =   "Synchronize Details to DB_Sulekha"
      End
      Begin VB.Menu ListOfRegisterOfBills 
         Caption         =   "List Of Register Of Bills"
      End
      Begin VB.Menu ListofWaterBills 
         Caption         =   "List of Water Bills"
      End
      Begin VB.Menu ChangePassword 
         Caption         =   "Change Password"
      End
   End
   Begin VB.Menu Inward 
      Caption         =   "Inward"
      Begin VB.Menu NewInward 
         Caption         =   "New Inward"
      End
      Begin VB.Menu BulkInward 
         Caption         =   "Bulk Inward"
      End
      Begin VB.Menu ManualInward 
         Caption         =   "Manual Inward"
      End
      Begin VB.Menu FrontOfficeDiary 
         Caption         =   "Front Office Diary"
      End
      Begin VB.Menu DespatchDiary 
         Caption         =   "Despatch Diary"
      End
      Begin VB.Menu SearchInward 
         Caption         =   "Search Inward"
      End
      Begin VB.Menu SoochikaMiscellaneousReport 
         Caption         =   "Miscellaneous Reports"
      End
   End
   Begin VB.Menu mnuIntegratedModules 
      Caption         =   "Integrated Modules"
      Begin VB.Menu mnuKMBR 
         Caption         =   "KMBR"
      End
      Begin VB.Menu mnuSevanaPension 
         Caption         =   "Sevana Pension"
      End
      Begin VB.Menu ViewPaymentOrder 
         Caption         =   "View Payment Order"
      End
      Begin VB.Menu AlterTransactionType 
         Caption         =   "Alter Transaction Type"
      End
      Begin VB.Menu IntegratedPayments 
         Caption         =   "Integrated Payments"
      End
      Begin VB.Menu SendCollectionDetailsToMainOffice 
         Caption         =   "Send Collection Details To Main Office"
      End
      Begin VB.Menu ViewVouchers 
         Caption         =   "View Vouchers"
      End
      Begin VB.Menu ListofWaivedFines 
         Caption         =   "List of Waived Fines"
      End
   End
   Begin VB.Menu PDE 
      Caption         =   "PDE"
      Begin VB.Menu ProjectVouchers 
         Caption         =   "Project Vouchers"
      End
      Begin VB.Menu PDEAllotments 
         Caption         =   "Allotments"
      End
   End
   Begin VB.Menu View 
      Caption         =   "View"
      Begin VB.Menu BalanceSheet 
         Caption         =   "Submit AFS to AIMS"
      End
      Begin VB.Menu ChequeRegisterView 
         Caption         =   "Cheque Register"
      End
      Begin VB.Menu viewHeadwiseConsolidation 
         Caption         =   "Headwise Cosolidation"
      End
      Begin VB.Menu ReverseEntryRegister 
         Caption         =   "Reverse Entry Register"
      End
      Begin VB.Menu ListOfSystemGeneratedJournals 
         Caption         =   "List of System Generated Journals"
      End
      Begin VB.Menu ViewRequisitionRegister 
         Caption         =   "Requisition Register"
      End
      Begin VB.Menu ViewPaymentOrderDetails 
         Caption         =   "PaymentOrder Details"
      End
      Begin VB.Menu VoucherExtractStatus 
         Caption         =   "Voucher Extract Status"
      End
      Begin VB.Menu DCBReport 
         Caption         =   "DCB Report"
      End
   End
   Begin VB.Menu Reports 
      Caption         =   "Reports"
      Begin VB.Menu DailyReports 
         Caption         =   "Daily Reports"
         Begin VB.Menu CounterwiseDetails 
            Caption         =   "Counter wise Details"
         End
         Begin VB.Menu Chitta 
            Caption         =   "Chitta"
         End
         Begin VB.Menu mnuCounterDayBook 
            Caption         =   "Counter Day Book"
         End
         Begin VB.Menu CancelledReceipts 
            Caption         =   "Cancelled Receipts"
         End
         Begin VB.Menu HeadwiseConsolidation 
            Caption         =   "Headwise Consolidation"
         End
         Begin VB.Menu DayBookReceipts 
            Caption         =   "Day Book - Receipts"
         End
      End
      Begin VB.Menu AdministratorsReport 
         Caption         =   "Administrators Report"
         Begin VB.Menu Wardwise 
            Caption         =   "Ward Wise"
         End
         Begin VB.Menu AdministratorsChitta 
            Caption         =   "Administrator’s Chitta"
         End
         Begin VB.Menu CounterDayBook 
            Caption         =   "Counter Day Book"
         End
         Begin VB.Menu DepartmentwiseReport 
            Caption         =   "Departmentwise Report"
         End
         Begin VB.Menu ZonalCollection 
            Caption         =   "Zonal Collection"
         End
         Begin VB.Menu CancelledCounterReceipts 
            Caption         =   "Cancelled Counter Receipts"
         End
         Begin VB.Menu TotalReceiptCount 
            Caption         =   "Receipt Count"
         End
         Begin VB.Menu TotalCounterCollection 
            Caption         =   "Total Counter Collection"
         End
         Begin VB.Menu TotalHeadwiseConsolidation 
            Caption         =   "Total Headwise Consolidation"
         End
         Begin VB.Menu DailyCounterConsolidation 
            Caption         =   "Daily Counter Consolidation"
         End
         Begin VB.Menu HeadWiseReport 
            Caption         =   "Head Wise Report"
         End
         Begin VB.Menu DetailedHeadWiseReport 
            Caption         =   "Detailed Head Wise Report "
         End
         Begin VB.Menu PredateReceiptCancelReport 
            Caption         =   "Pre date Receipt Cancel Report"
         End
         Begin VB.Menu CardPaymentReport 
            Caption         =   "Card Payment Report"
         End
      End
      Begin VB.Menu SoochikaInwardReports 
         Caption         =   "Soochika Inward"
         Begin VB.Menu DistributionRegister 
            Caption         =   "Distribution Register"
         End
         Begin VB.Menu SecurityRegister 
            Caption         =   "Security Register"
         End
         Begin VB.Menu ListOfInwards 
            Caption         =   "List Of Inwards"
         End
         Begin VB.Menu Miscellaneousreports 
            Caption         =   "Miscellaneous Reports"
         End
      End
      Begin VB.Menu SourcewiseReports 
         Caption         =   "Source wise Reports"
         Begin VB.Menu sourcewiseReceiptAndPayment 
            Caption         =   "Source-wise Receipts && Payments Statement"
         End
         Begin VB.Menu SectorwiseStatement 
            Caption         =   "Sector-wise Statement of Expenditure from Development Fund"
         End
         Begin VB.Menu CapitalandRevenueExpenditure 
            Caption         =   "Capital and Revenue Expenditure from different"
         End
         Begin VB.Menu ReceiptsunderOwnFundMajorHead 
            Caption         =   "Receipts under Own Fund - Major Head"
         End
         Begin VB.Menu ReceiptsunderOwnFundMinorHead 
            Caption         =   "Receipts under Own Fund - Minor Head"
         End
         Begin VB.Menu ExpenditurefromOwnFundMajorHead 
            Caption         =   "Expenditure from Own Fund - Major Head"
         End
         Begin VB.Menu ExpenditurefromOwnFund 
            Caption         =   "Expenditure from Own Fund - Minor Head"
         End
      End
      Begin VB.Menu rptCashBook 
         Caption         =   "Cash Book"
      End
      Begin VB.Menu rptBankBook 
         Caption         =   "Bank Book"
      End
      Begin VB.Menu rptJournalBook 
         Caption         =   "Journal Book"
      End
      Begin VB.Menu rptLedgerBook 
         Caption         =   "Ledger Book"
      End
      Begin VB.Menu rptTrialBalance 
         Caption         =   "Trial Balance"
      End
      Begin VB.Menu rptBalanceSheet 
         Caption         =   "Balance Sheet"
      End
      Begin VB.Menu rptIncomeAndExpenditure 
         Caption         =   "Income & Expenditure"
      End
      Begin VB.Menu rptReceiptsAndPayments 
         Caption         =   "Receipts & Payments"
      End
      Begin VB.Menu rptBudgetVariance 
         Caption         =   "Budget Variance"
      End
      Begin VB.Menu BudgetAnalysisReport 
         Caption         =   "Budget Analysis Report"
      End
      Begin VB.Menu BankReconciliation 
         Caption         =   "Bank Reconciliation"
      End
      Begin VB.Menu CashFlow 
         Caption         =   "Cash Flow Statement"
      End
      Begin VB.Menu KeyRatio 
         Caption         =   "Key Ratio Analysis"
      End
      Begin VB.Menu ChequeRegister 
         Caption         =   "Cheque Registers"
         Begin VB.Menu ChequeIssue 
            Caption         =   "Cheque Issue"
         End
         Begin VB.Menu ChequeReceived 
            Caption         =   "Cheque Received"
         End
      End
      Begin VB.Menu rptRegisters 
         Caption         =   "Registers"
         Begin VB.Menu rptAppropriationControlRegister 
            Caption         =   "Appropriation Control Register"
         End
         Begin VB.Menu rptAssetReplacementRegister 
            Caption         =   "Asset Replacement Register"
         End
         Begin VB.Menu rptAuthorisationIssuetoSecretary 
            Caption         =   "Authorisation Issue to Secretary"
         End
         Begin VB.Menu rptBillofReceiptsRegister 
            Caption         =   "Bill of Receipts Register"
         End
         Begin VB.Menu rptCollectionregister 
            Caption         =   "Collection Register"
         End
         Begin VB.Menu rptDemandregister 
            Caption         =   "Demand register"
         End
         Begin VB.Menu rptDepositreceivedregister 
            Caption         =   "Deposit Received Register"
         End
         Begin VB.Menu rptDocumentcontrolRegister 
            Caption         =   "Document control Register"
         End
         Begin VB.Menu rptFormGEN40Register 
            Caption         =   "FormGEN-40 Register"
         End
         Begin VB.Menu rptFunctionWiseExpenditure 
            Caption         =   "Function wise Expenditure Subsidiary Ledger "
         End
         Begin VB.Menu rptFunctionwisereceiptsubsidiaryledger 
            Caption         =   "Function wise Receipt Subsidiary Ledger"
         End
         Begin VB.Menu rptFundsReceivedRegister 
            Caption         =   "Funds Received Register"
         End
         Begin VB.Menu rptImmovablePropertyRegister 
            Caption         =   "Immovable Property Register"
         End
         Begin VB.Menu rptImplentingOfficerwiseAllotmentRegister 
            Caption         =   "Implenting Officer wise Allotment Register"
         End
         Begin VB.Menu rptIncomeandExpenditureRegister 
            Caption         =   "Income And Expenditure Register"
         End
         Begin VB.Menu rptLandRegister 
            Caption         =   "Land Register"
         End
         Begin VB.Menu rptLetterofallotment 
            Caption         =   "Letter Of Allotment"
         End
         Begin VB.Menu rptMovablepropertyRegister 
            Caption         =   "Movable property Register"
         End
         Begin VB.Menu rptOfficialReceiptRegister 
            Caption         =   "Official Receipt Register"
         End
         Begin VB.Menu rptPaymentOrderRegister 
            Caption         =   "Payment Order Register"
         End
         Begin VB.Menu rptProjectregister 
            Caption         =   "Project Register"
         End
         Begin VB.Menu rptRegisterofadvances 
            Caption         =   "Register Of Advances"
         End
         Begin VB.Menu rptRegisterofbillsforpayment 
            Caption         =   "Register Of Bills For Payment"
         End
         Begin VB.Menu rptRegisterofPermenantadvance 
            Caption         =   "Register Of Permenant Advance"
         End
         Begin VB.Menu rptRegisterofpubliclightingsystem 
            Caption         =   "Register Of Public Lighting System"
         End
         Begin VB.Menu rptRequesitionforReleaseofFundcodes 
            Caption         =   "Requesition for Release of Fund codes"
         End
         Begin VB.Menu rptStatementofOutstandingLiabilityforexpenses 
            Caption         =   "Statement of Outstanding Liability for Expenses"
         End
         Begin VB.Menu rptStatementonStatusofChequereceived 
            Caption         =   "Statement on Status of Cheque Received"
         End
         Begin VB.Menu rptSubsidiaryRegister 
            Caption         =   "Subsidiary Register"
         End
         Begin VB.Menu rptSummaryofCollectionRegister 
            Caption         =   "Summary of Collection Register"
         End
         Begin VB.Menu rptSummaryStatementodfbills 
            Caption         =   "Summary Statement of Bills"
         End
         Begin VB.Menu rptSummaryStatementofDeposits 
            Caption         =   "Summary Statement of Deposits"
         End
         Begin VB.Menu rptSummaryStatementofRefundandRemission 
            Caption         =   "Summary Statement of Refund and Remission"
         End
         Begin VB.Menu rptSummaryStatementofWriteoffs 
            Caption         =   "Summary Statement of Write-offs"
         End
      End
      Begin VB.Menu ViewSubsidiaryCashBook 
         Caption         =   "View Subsidiary Cash Book"
      End
      Begin VB.Menu TestReport 
         Caption         =   "Test Report"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Exit 
      Caption         =   "Exit"
      Begin VB.Menu TeamSaankhya 
         Caption         =   "Developer Centre"
         Visible         =   0   'False
      End
      Begin VB.Menu Test3 
         Caption         =   "Test 3"
      End
      Begin VB.Menu rptLedgerView 
         Caption         =   "Ledger View"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu YearEndProcess 
         Caption         =   "Year End Process"
      End
      Begin VB.Menu ReportGenerator 
         Caption         =   "Report Generator"
      End
      Begin VB.Menu AccountHeadsNew 
         Caption         =   "Account Heads "
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu SearchBuilding 
         Caption         =   "Search Building"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu TransactionTemp 
         Caption         =   "TransactionTemp"
      End
      Begin VB.Menu DeleteTransactionEntry 
         Caption         =   "Delete Transaction Entry"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Lock 
         Caption         =   "Lock"
         Shortcut        =   ^L
      End
      Begin VB.Menu Logoff 
         Caption         =   "Log Off"
      End
      Begin VB.Menu LogOut 
         Caption         =   "LogOut"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim mTimer                      As Integer ' Added for blinking Previous Year Transaction
    Dim mPreYearMode                As Boolean
    Dim mInterruptedReceiptID       As Boolean 'Added to block receipt counter for IR Register
    
    Private Sub CheckInterruptReceiptRequestStatus()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim mStatus As Variant
        
        On Error GoTo err
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mStatus = ""
        mSql = "Select tnyStatus From faInterruptedRequests"
        'If gbUserTypeID = 3 Then
        mSql = mSql + " Where numUserID =" & gbUserID
        mSql = mSql + " And intCounterID =" & gbCounterID
        mSql = mSql + " And intTypeID = 1"
        'End If
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mStatus = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
        End If
        Rec.Close
        mCnn.Close
        If mStatus <> "" Then
            If mStatus = 1 Or mStatus = 2 Then
                RequestforInterruptedReceipt.Caption = "Cancel request for InterruptedReceipt"
                mInterruptedReceiptID = True
            End If
        Else
                RequestforInterruptedReceipt.Caption = "Request for InterruptedReceipt"
                mInterruptedReceiptID = False
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub
     Private Function CheckSubmissionStatus() As Variant
        Dim mYear As String
       ' Dim mLBCode As Integer
        Dim xmlHttp As Object
        Set xmlHttp = CreateObject("MSXML2.XmlHttp")
        Dim param As String
        Dim mCnn As New ADODB.Connection
        Dim mRec As New ADODB.Recordset
        Dim mRec1 As New ADODB.Recordset
        Dim objdb As New clsDB
        Dim aryIn As Variant
        Dim mCnt    As Integer
        Dim mTotAmt As Double
        Dim params
        Dim mRowCnt As Integer
        Dim Index As Integer
        Dim mResult As String
        Dim mSql As String
        Dim mState As Integer
        Dim mPreYear As Integer
        
        mPreYear = gbFinancialYearID - 1
        
        params = "lbCode=" & CStr(gbLBCODE) + "&year=" + CStr(mPreYear) + "-" + CStr(mPreYear + 1)

        On Error GoTo Message:
            xmlHttp.Open "POST", "http://aims.ksad.kerala.gov.in/esubmission/getAccountStatus.action", False
            xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            xmlHttp.send params
            
            'MsgBox xmlHttp.responseText
            mResult = xmlHttp.responseText
            If mResult = "$NotSubmitted$$" Then
                mState = 0
            ElseIf mResult = "$SubmittedFromSankhya$$" Then
                mState = 1
            ElseIf mResult = "$Submitted$$" Then
                mState = 2
            ElseIf mResult = "$Accepted$$" Then
                mState = 3
            ElseIf Left(mResult, 10) = "$Rejected$" Then
                mState = 4
            End If
''        Else
''            MsgBox ("Connection to the LFA Webservice can not be Established ")
        
        CheckSubmissionStatus = mState
       Exit Function
Message:
        MsgBox ("Connection To The LFA Webservice Can Not be Established Now, Try Again !!!")
        
    End Function
    Private Function pendingTaskDisable()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim objdb   As New clsDB
        Dim mSql    As String
        Dim mDiff   As Integer
        Dim mStat   As Integer
        mStat = 0
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        If gbLBPanchayat = 1 Then
'           If gbLocalBodyID = 1287 Or gbLocalBodyID = 386 Or gbLocalBodyID = 382 gbLocalBodyID = 165 Then
           If gbLocalBodyID = 718 Then ''Vadekkekad   PAllivasal GP 975 Valavannur Gp 92 --Iringalakkuda BP
                mSql = "Select DATEDIFF(day, '1/Apr/2020', getdate()) as Diff From faConfig "
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    mDiff = IIf(IsNull(Rec!diff), "", Rec!diff)
                End If
                
                If mDiff > 321 Then   'jan 30 21 days for selected local bodied
                    mStat = 1
                End If
                Rec.Close
           Else
                mStat = 1
           End If
        Else
            If gbLocalBodyID = 1259 Then  ''Kanur Corp 1268 Wadakkenchery MP For Alappuzha MP 185
                mSql = "Select DATEDIFF(day, '1/Apr/2020', getdate()) as Diff From faConfig "
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                    mDiff = IIf(IsNull(Rec!diff), "", Rec!diff)
                End If
                
                If mDiff > 327 Then   '86 80 65 54 44 Then ' upto Jul 31 110  Jul31 Aug 24
                    mStat = 1
                End If
                Rec.Close
            End If
        End If
        pendingTaskDisable = mStat
    End Function
    
    Private Sub CheckPendingTransactionStatus()
        Dim mCnn    As New ADODB.Connection
        Dim Rec     As New ADODB.Recordset
        Dim mSql    As String
        Dim objdb   As New clsDB
        Dim mStatus As Variant
        Dim mMenu   As Control
        Dim mRequestDate As String
        On Error GoTo err
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        
        mStatus = ""
        mSql = "Select tnyStatus,dtRequestDate From faInterruptedRequests"
        mSql = mSql + " Where  intTypeID = 2 "
        'If gbUserTypeID = 3 Then
        If gbSeatGroupID = gbSeatGroupAccountsClerk Then
            mSql = mSql + " And intCounterID =" & gbCounterID
            mSql = mSql + " And numUserID =" & gbUserID
        End If
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            mStatus = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
            mRequestDate = IIf(IsNull(Rec!dtRequestDate), "", Rec!dtRequestDate)
        End If
        Rec.Close
        mCnn.Close
        If mStatus <> "" Then
            If mStatus = 1 Then
                'If gbUserTypeID = 4 Or gbUserTypeID = 2 Then
                If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                    Timer1.Enabled = True
                    lblPreyear.Caption = "Pevious Year's Transaction Request is Pending for Approval"
                Else
                    lblPreyear.Caption = "Pevious Year's Transaction Request is Pending for Approval"
                    For Each mMenu In frmMenu.Controls
                        If TypeOf mMenu Is Menu Then
                            Debug.Print mMenu.Name
                            If mMenu.Name = "RequestforPendingTransactions" Or mMenu.Name = "Utilities" Then
                                mMenu.Enabled = True
                            Else
                                mMenu.Enabled = False
                            End If
                        End If
                    Next
                    RequestforPendingTransactions.Caption = "Cancel Request for Enable Previous Year's Transaction"
                End If
            ElseIf mStatus = 2 Then
                'If gbUserTypeID = 4 Or gbUserTypeID = 2 Then
                If gbSeatGroupID = gbSeatGroupAccountsOfficer Then
                    Timer1.Enabled = False
                    
                Else
                    mPreYearMode = True
                    gbTransactionDate = mRequestDate
                    If Month(gbTransactionDate) < 4 Then
                        gbFinancialYearID = Year(gbTransactionDate) - 1
                    Else
                        gbFinancialYearID = Year(gbTransactionDate)
                    End If
                    RequestforPendingTransactions.Caption = "Cancel Request for Enable Previous Year's Transaction"
                    CounterReceipt.Enabled = False
                    Timer1.Enabled = True
                    lblPreyear.Caption = "Previous Year Transaction Mode is Enabled"
                End If
            End If
        Else
            Timer1.Enabled = False
            RequestforPendingTransactions.Caption = "Request for Enable Previous Year's Transaction"
        End If
        Exit Sub
err:
        MsgBox err.Description
    End Sub

    Private Sub AccountHeads_Click()
'        frmAccountHeadsNew.Visible = True
'        frmAccountHeadsNew.ZOrder (0)
        frmAccountHeadsOpeningBalance.Visible = True
        frmAccountHeadsOpeningBalance.ZOrder (0)
    End Sub
    Private Sub AccountHeadsNew_Click()
        'frmAccountHeadsNew.Visible = True
        'frmAccountHeadsNew.ZOrder (0)
    End Sub
    Private Sub AccrualDemand_Click()
        frmAccrualDemand.Visible = True
        frmAccrualDemand.ZOrder (0)
    End Sub
    Private Sub AdministratorsChitta_Click()
        frmRptAdmin.Visible = True
        frmRptAdmin.ZOrder (0)
        Call frmRptAdmin.frameVisible
        frmRptAdmin.fraChitta.Visible = True
        frmRptAdmin.Width = 5545
        frmRptAdmin.Height = 1700
    End Sub

   

    Private Sub AFS_Click()
        frmAFSFinalization.Visible = True
        frmAFSFinalization.ZOrder (0)
    End Sub

'    Private Sub AFS_Click()
'        frmAFSClosingBalanceSheet.Visible = True
'        frmAFSClosingBalanceSheet.ZOrder (0)
'    End Sub

    Private Sub AFSSourceOfFund_Click()
        frmAFSClosingSourceOfFund.Visible = True
        frmAFSClosingSourceOfFund.ZOrder (0)
    End Sub

    Private Sub Agreement_Click()
        frmListOfAgreements.Visible = True
        frmListOfAgreements.ZOrder (0)
    End Sub

    Private Sub AllotmentClosingBalance_Click()
        frmAllotmentClosingBalance.Visible = True
        frmAllotmentClosingBalance.ZOrder (0)
    End Sub

    Private Sub AlterTransactionType_Click()
        frmAlterTransactionType.Visible = True
        frmAlterTransactionType.ZOrder (0)
    End Sub
    Private Sub ApprovalofInterruptedReceiptCancellation_Click()
'        frmInterruptedCancellationApproval.Visible = True
'        frmInterruptedCancellationApproval.ZOrder (0)
    End Sub
    Private Sub ApprovalOfRecieptCancellation_Click()
        frmReceiptCancellationList.Visible = True
        frmReceiptCancellationList.ZOrder (0)
    End Sub

    Private Sub BalanceSheet_Click()
            frmViewBalanceSheet.Visible = True
            frmViewBalanceSheet.ZOrder (0)
    End Sub

    Private Sub Bank_Click()
        frmBank.Visible = True
        frmBank.ZOrder (0)
    End Sub
    Private Sub BankReconcile_Click()
        If gbLBPanchayat = 0 Then
            frmBankReconcilationProcess.Visible = True
            frmBankReconcilationProcess.ZOrder (0)
        End If
    End Sub
    Private Sub BankReconciliationEntry_Click()
'        frmBankReconciliation.Visible = True
'        frmBankReconciliation.ZOrder (0)
         If gbLBPanchayat = 0 Then
            frmBankScroll.Visible = True
            frmBankScroll.ZOrder (0)
         End If
  
    End Sub
    Private Sub BankReconUtility_Click()
        frmReconBankList.Visible = True
        frmReconBankList.ZOrder (0)
    End Sub

    Private Sub BeneficiaryContribution_Click()
        frmProjectJournalDetails.LoadMode = 50
        frmProjectJournalDetails.Visible = True
        frmProjectJournalDetails.ZOrder (0)
    End Sub

'    Private Sub BudgetAllocation_Click()
'        frmBudgetAllocation.Visible = True
'        frmBudgetAllocation.ZOrder (0)
'    End Sub
    Private Sub BudgetAnalysisReport_Click()
        frmBudgetVariance.Visible = True
        frmBudgetVariance.ZOrder (0)
    End Sub
    Private Sub BudgetCentres_Click()
        frmBudgetCentre.Visible = True
        frmBudgetCentre.ZOrder (0)
    End Sub
    Private Sub BudgetRevision_Click()
        frmBudgetRevision.Visible = True
        frmBudgetRevision.ZOrder (0)
    End Sub

    Private Sub BulkInward_Click()
'        gbSoochikaVer = gbLinkWithSevana
'        If gbSoochikaVer = 5 Then
'            MsgBox "Bulk inward not permitted !!!", vbInformation, "SOOCHIKA"
'        Else
            frmSoochikaBulkInward.Show
            frmSoochikaBulkInward.ZOrder (0)
'        End If
    End Sub

    Private Sub CancelledReceipts_Click()
        Dim objdb As New clsDB
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        ''frmMenu.Transactions.Enabled = False
        arInput = Array(CStr(gbCounterID), gbTransactionDate, gbTransactionDate, CStr(gbUserID))
        
        frmNewViewer.rptFileName = App.Path & "\Reports\rptCancelledReceipts.rpt"
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
        
    End Sub

    Private Sub CancelPaymentOrder_Click()
        frmPayorderCancellationsList.Visible = True
        frmPayorderCancellationsList.ZOrder (0)
    End Sub
    Private Sub CapitalandRevenueExpenditure_Click()
        frmRptSourceWiseReports.Visible = True
        frmRptSourceWiseReports.cmbReportMenu.ListIndex = 2
        frmRptSourceWiseReports.ZOrder (0)
    End Sub

    Private Sub CardPaymentReport_Click()
        frmRptAdmin.Visible = True
        frmRptAdmin.ZOrder (0)
        
        
        Call frmRptAdmin.frameVisible
        frmRptAdmin.fmeCardPayment.Visible = True
        frmRptAdmin.fmeCardPayment.Left = 0
        frmRptAdmin.fmeCardPayment.Top = 0
        frmRptAdmin.Width = 5445
        frmRptAdmin.Height = 2005
        
      
    End Sub

    Private Sub CashFlow_Click()
        frmRptFilterFields.rptNames = 48
        frmRptFilterFields.Show vbModal
    End Sub
    Private Sub ChangePassword_Click()
        frmChangePassword.Visible = True
        frmChangePassword.ZOrder (0)
    End Sub

    Private Sub CheckOutAccountingAssistant_Click()
        frmAccAssistantCheckout.Visible = True
        frmAccAssistantCheckout.ZOrder (0)
    End Sub

    Private Sub ChequeBook_Click()
        frmChequeBook.Visible = True
        frmChequeBook.ZOrder (0)
    End Sub
    Private Sub ChequeRegisterUtility_Click()
        frmChequeRegister.Visible = True
        frmChequeRegister.ZOrder (0)
    End Sub
    Private Sub ChequeRegisterView_Click()
        frmViewChequeRegister.Visible = True
        frmViewChequeRegister.ZOrder (0)
    End Sub
    Private Sub ChequeReturnList_Click()
        frmListReverseEntryRequests.LoadMode = 2
        frmListReverseEntryRequests.Visible = True
        frmListReverseEntryRequests.ZOrder (0)
    End Sub
    Private Sub Chitta_Click()
        Dim objdb As New clsDB
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        ''frmMenu.Transactions.Enabled = False
        'arInput = Array(gbTransactionDate, gbTransactionDate, CStr(gbCounterID), "%", "%", "%")
        arInput = Array(gbTransactionDate, gbTransactionDate, CStr(gbCounterID), "%", "%", "%", CStr(gbUserID))
        frmNewViewer.rptFileName = App.Path & "\Reports\rptChitta.rpt"
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
    End Sub

    Private Sub ClosingAccounts_Click()
        frmExtractYearWiseList.Visible = True
        frmExtractYearWiseList.ZOrder (0)
    End Sub

    Private Sub CollectionRegister_Click()
        frmCollectionRegister.Visible = True
        frmCollectionRegister.ZOrder (0)
    End Sub
    
    Private Sub ConfigurationSettings_Click()
        frmDefaultSettingsOfReceipts.Visible = True
        'frmDefaultSettingsOfReceipts.ZOrder (0)
    End Sub
    Private Sub ContraEntry_Click()
        frmListOfContraEntries.Visible = True
        frmListOfContraEntries.ZOrder (0)
    End Sub
    
    Private Sub CounterDayBook_Click()
        frmRptAdmin.Visible = True
        frmRptAdmin.ZOrder (0)
        Call frmRptAdmin.frameVisible
        frmRptAdmin.fraDayBook.Visible = True
        frmRptAdmin.Width = 5545
        frmRptAdmin.Height = 1700
    End Sub
   
    Private Sub CounterReceipt_Click()
        CheckInterruptReceiptRequestStatus
            frmReceiptsCounter.mWebExtractMode = False
        If mInterruptedReceiptID = False Then
            frmReceiptsCounter.SoochikaConnected = False
            frmReceiptsCounter.Visible = True
            frmReceiptsCounter.ZOrder (0)
        End If
    End Sub

    Private Sub CounterwiseDetails_Click()
        frmCounterReport.Visible = True
        frmCounterReport.ZOrder (0)
    End Sub
    
    Private Sub DayBookReceipts_Click()
        frmDayBookReceipts.Visible = True
        frmDayBookReceipts.ZOrder (0)
    End Sub

    Private Sub DCBReport_Click()
        frmRefreshDCB.Visible = True
        frmRefreshDCB.ZOrder (0)
    End Sub

    Private Sub DefineReportSchedules_Click()
        frmReportSchedules.Visible = True
        frmReportSchedules.ZOrder (0)
    End Sub
    Private Sub DeleteTransactionEntry_Click()
        frmDeleteJournal.Visible = True
        frmDeleteJournal.ZOrder (0)
    End Sub
    Private Sub DemandInbox_Click()
        frmZonalInbox.Visible = True
        frmZonalInbox.ZOrder (0)
    End Sub
    Private Sub DemandInterface_Click()
        frmListOfDemands.Visible = True
        frmListOfDemands.ZOrder (0)
    End Sub
    Private Sub DemandRegister_Click()
        frmDemandRegister.Visible = True
        frmDemandRegister.ZOrder (0)
    End Sub
    Private Sub DespatchDiary_Click()
        If gbSoochikaVer <> 5 Then
            frmSoochikaDespatchDiary.Show
            frmSoochikaDespatchDiary.ZOrder (0)
        End If
    End Sub
    Private Sub DetailedHeadWiseReport_Click()
        frmRptAdmin.HeadWiseReport = 2
        frmRptAdmin.Visible = True
        frmRptAdmin.ZOrder (0)
        
        Call frmRptAdmin.frameVisible
        frmRptAdmin.fmeheadwiseReport.Visible = True
        frmRptAdmin.fmeheadwiseReport.Left = 0
        frmRptAdmin.fmeheadwiseReport.Top = 0
        frmRptAdmin.Width = 5415
        frmRptAdmin.Height = 2000
    End Sub
    Private Sub DistributionRegister_Click()
        Dim vAryInRpt(1)
        vAryInRpt(0) = CStr(Date)
        frmCRViewer.vShowReport App.Path & "\soochika\Reports", "rptSubUnitRegister.rpt", vAryInRpt
        frmCRViewer.Show 1
    End Sub

    Private Sub ebill_Click(Index As Integer)
        frmWebExtracts.mPreviousYearMode = 0
        frmWebExtracts.Visible = True
        frmWebExtracts.ZOrder (0)
    End Sub

    Private Sub ExpenditurefromOwnFund_Click()
        frmRptSourceWiseReports.Visible = True
        frmRptSourceWiseReports.cmbReportMenu.ListIndex = 6
        frmRptSourceWiseReports.ZOrder (0)
    End Sub
    Private Sub ExpenditurefromOwnFundMajorHead_Click()
        frmRptSourceWiseReports.Visible = True
        frmRptSourceWiseReports.cmbReportMenu.ListIndex = 5
        frmRptSourceWiseReports.ZOrder (0)
    End Sub
    Private Sub ExtractDataForBudget_Click()
        frmExtractDataForBudget.Show vbModal
    End Sub
    Private Sub Fields_Click()
        frmFields.Visible = True
        frmFields.ZOrder (0)
    End Sub
''    Private Sub FinancialYearSettings_Click()
''        frmFinancialYear.Visible = True
''        frmFinancialYear.ZOrder (0)
''    End Sub
    Private Sub FrontOfficeDiary_Click()
'        If gbSoochikaVer <> 5 Then
            frmSoochikaFrontOfficeDiary.Show
            frmSoochikaFrontOfficeDiary.ZOrder (0)
 '       End If
    End Sub
    Private Sub Functionaries_Click()
        frmFunctionary.Visible = True
        frmFunctionary.ZOrder (0)
    End Sub
    Private Sub FunctionaryHead_Click()
        frmFunctionaryHeads.Visible = True
        frmFunctionaryHeads.ZOrder (0)
    End Sub
    Private Sub Functions_Click()
        frmFunctions.Visible = True
        frmFunctions.ZOrder (0)
    End Sub
    Private Sub Funds_Click()
        frmFund.Visible = True
        frmFund.ZOrder (0)
    End Sub
    
    Private Sub GST_Click()
        frmGST.Visible = True
        frmGST.ZOrder (0)
    End Sub

    Private Sub HeadwiseConsolidation_Click()
        Dim objdb As New clsDB
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
        arInput = Array(gbTransactionDate, gbTransactionDate)
        frmNewViewer.rptFileName = App.Path & "\Reports\rptHeadwiseCollection.rpt"
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
    End Sub
    
    Private Sub HeadWiseReport_Click()
        frmRptAdmin.HeadWiseReport = 1
        frmRptAdmin.Visible = True
        frmRptAdmin.ZOrder (0)
        
        Call frmRptAdmin.frameVisible
        frmRptAdmin.fraLedger.Visible = True
        frmRptAdmin.fraLedger.Left = 0
        frmRptAdmin.fraLedger.Top = 0
        frmRptAdmin.Width = 5450
        frmRptAdmin.Height = 2500
    End Sub

Private Sub ImplementingOfficer_Click()
    frmImplementingOfficerList.Visible = True
    frmImplementingOfficerList.ZOrder (0)
End Sub

    Private Sub IntegratedPayments_Click()
        frmIntegratedPayments.Visible = True
        frmIntegratedPayments.ZOrder (0)
    End Sub

    Private Sub InterruptedRegister_Click()
        frmInterruptedReceiptRegister.Visible = True
        frmInterruptedReceiptRegister.ZOrder (0)
    End Sub

    Private Sub InwardChecksAndDds_Click()
        frmInwardValuebles.Visible = True
        frmInwardValuebles.ZOrder (0)
    End Sub
    Private Sub IssueOfInterruptedReceiptBook_Click()
        frmInterruptedReceiptBooks.Visible = True
        frmInterruptedReceiptBooks.ZOrder (0)
    End Sub
    Private Sub JournalEntry_Click()
        frmJournalEntry.PreviousYearMode = 0
        frmJournalEntry.mWebExtractJV = False
        frmJournalEntry.PreviousYearRequestID = -1
        frmJournalEntry.Visible = True
        frmJournalEntry.ZOrder (0)
    End Sub
    Private Sub KeyRatio_Click()
        frmRptFilterFields.rptNames = 49
        frmRptFilterFields.Show vbModal
    End Sub
    Private Sub lblTransactionDate_Click()
        '        Dim objDb   As New clsDB
        '        Dim mCnn    As New ADODB.Connection
        '        Dim Rec     As New ADODB.Recordset
        '        Dim msql    As String
        '        objDb.CreateNewConnection mCnn, enuSourceString.Saankhya
        '        Dim tnyStatus As Integer
        '        msql = "Select tnyStatus,dtRequestDate From faInterruptedRequests"
        '        msql = msql + " Where  intTypeID = 2 And numUserID=" & gbUserID
        '        Rec.Open msql, mCnn
        '        If Not (Rec.EOF And Rec.BOF) Then
        '            tnyStatus = IIf(IsNull(Rec!tnyStatus), "", Rec!tnyStatus)
        '            If tnyStatus = 1 Then
        '                MsgBox "Pending Approval for Pevious Year's Transaction Request", vbInformation
        '                Exit Sub
        '            ElseIf tnyStatus = 2 Then
        '                MsgBox "Pevious Year's Transaction Mode is Active", vbInformation
        '                Exit Sub
        '            End If
        '        Else
        '            If gbCounterSectionID <> gbJSKSectionID And gbSeatGroupID <> gbSeatGroupAuditorsGroup Then  'Added By Sinoj for Restrict date Change (Auditor)
        '                frmChangeDate.Show vbModal
        '            End If
        '        End If
    End Sub
    
    Private Sub LetterOfAllotment_Click()
        frmListOfAllotments.LoadMode = 10
        frmListOfAllotments.AuthorityOrAllotment = "Allotment"
        frmListOfAllotments.Visible = True
        frmListOfAllotments.ZOrder (0)
    End Sub
    
    Private Sub LinkAllotmentsWithReceipts_Click()
        frmLinkAllotmentsWithReceipts.Visible = True
        frmLinkAllotmentsWithReceipts.ZOrder (0)
    End Sub
    
    Private Sub LinkRecoveriesToProjectExpenditure_Click()
        frmLinkRecoveriesToProjectExp.Visible = True
        frmLinkRecoveriesToProjectExp.ZOrder (0)
    End Sub
    
    Private Sub ListOfCancelledRequisitions_Click()
        frmListOfCancelledRequisitions.Visible = True
        frmListOfCancelledRequisitions.ZOrder (0)
    End Sub
    
    Private Sub ListOfInwards_Click()
         Dim vAryInRpt(1)
        vAryInRpt(0) = CStr(Date)
        frmCRViewer.vShowReport App.Path & "\soochika\Reports", "rptDistributionReg.rpt", vAryInRpt
        frmCRViewer.Show 1
    End Sub
    
    Private Sub ListOfPreviousDateReceiptCancellation_Click()
        frmListOfReceiptCancellationRequest.Visible = True
        frmListOfReceiptCancellationRequest.ZOrder (0)
    End Sub
    
    Private Sub ListOfRegisterOfBills_Click()
        frmListOfRegisterOfBills.Visible = True
        frmListOfRegisterOfBills.ZOrder (0)
    End Sub
    
    Private Sub ListOfSystemGeneratedJournals_Click()
        frmViewSystemJournals.Visible = True
        frmViewSystemJournals.ZOrder (0)
    End Sub
    
    Private Sub ListofWaivedFines_Click()
        frmListofWaivedFine.Visible = True
        frmListofWaivedFine.ZOrder (0)
    End Sub
    
    Private Sub LocalBodySettings_Click()
        frmLocalBodySettings.Visible = True
        frmLocalBodySettings.ZOrder (0)
    End Sub
    
    Private Sub ListofWaterBills_Click()
        frmSn_WrBillListOfTransactionDetails.Visible = True
        frmSn_WrBillListOfTransactionDetails.ZOrder (0)
    End Sub
    
    Private Sub Lock_Click()
        frmLock.Visible = True
        frmLock.ZOrder (0)
    End Sub
    
    Private Sub Logoff_Click()
        frmSplash.Show vbModal
    End Sub
    
    Private Sub LogOut_Click()
        Dim objdb As New clsDB
        Dim mCnn As New ADODB.Connection
        Dim mSql As String
        Dim objUser As clsUser
        Dim Rec As New ADODB.Recordset
        If (objdb.SetConnection(mCnn)) Then
            If mPreYearMode Then
                If (MsgBox("Are you Sure to LogOut! " & vbNewLine & "It will Change Your Previous Year's Transaction Mode...", vbExclamation + vbYesNo)) = vbYes Then
                    mSql = "Delete From faInterruptedRequests Where intCounterID=" & gbCounterID & "And numUserID=" & gbUserID
                    objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                    mCnn.Execute "Update faUserMovement Set dtLogoutTime=getdate(),tnyStatus=3 where intID=" & gbSessionID
                    End
                    Exit Sub
                End If
            Else
                objdb.SetConnection mCnn
                mCnn.Execute "Update faUserMovement Set dtLogoutTime=getdate(),tnyStatus=3 where intID=" & gbSessionID
                End
            End If
        End If
    End Sub
    Private Sub ManualInward_Click()
        gbSoochikaVer = gbLinkWithSevana
        If (gbSoochikaVer = 5) Then
            frmUSoochikaManualInward.Show
            frmUSoochikaManualInward.ZOrder (0)
        Else
            frmSoochikaManualInward.Show
            frmSoochikaManualInward.ZOrder (0)
        End If
    End Sub

Private Sub MDIForm_Activate()
        'DATA EXTRACTION
        If gbCounterSectionID <> gbJSKSectionID Then
            If Not IsDate(GetLastExtractedDate) Then
                frmInitialize.Show vbModal
                'MsgBox "LOAD INITIALIZE"
            End If
        End If
End Sub

    Private Sub MDIForm_Load()
        Dim AppPath As String
        Dim mStr As String
        AppPath = Trim(GetSetting("SaankhyaDE", "App", "Path", ""))
        If AppPath <> "" Then
            If App.Path <> AppPath Then
                mStr = mStr + " Another Instance of same application exists in " & vbCrLf
                mStr = mStr + " Path " + AppPath
                mStr = mStr + " This may cause conflict in Versions!"
                MsgBox mStr, vbCritical
                SaveSetting "SaankhyaDE", "App", "Path", CStr(App.Path)
            End If
        Else
            'Note:- Application Path Store's in Registry
            SaveSetting "SaankhyaDE", "App", "Path", CStr(App.Path)
        End If
        
        'SetMenu
        CheckInterruptReceiptRequestStatus
        SaveSetting "Saankhya", "Lock", "UserName", ""
        If Not SetEnvironment Then
            MsgBox "Didn't able to set Enviornment!", vbInformation
            End ' Note this must be Removed Later!! -> By Aiby
                ' Error: Makes visible all menu Items
                ' Menu Should be disabled at the starting
        End If
        If gbLBType = 3 Or gbLBType = 4 Then
            lblVersion.Caption = "Ver:" & gbVerID & "." & gbVerSubID
            lblDBVersion.Caption = "DBVer:" & gbDBVerID & "." & gbDBSubVerID
        Else
            lblVersion.Caption = "Ver:" & gbPVerID & "." & gbPVerSubID
            lblDBVersion.Caption = "DB Ver:" & gbPDBVerID & "." & gbPDBSubVerID
        End If
        
        If gbSeatGroupID = gbSeatByDeveloper Then
        Else
            MenuDetails
        End If
        
        CheckPendingTransactionStatus
     
        Me.Caption = Me.Caption & " [ " & gbLBTitle & " ]"
        gbCounterStatusFlag = False
        lblLoginName.Caption = gbUserName
        lblCounter.Caption = gbCounterName & "( " & gbCounterNo & " )"
        lblFinancialYear.Caption = CStr(gbFinancialYearID) & "-" & CStr(gbFinancialYearID + 1)
        lblTransactionDate.Caption = CStr(DdMmmYy(gbTransactionDate))
        

        '&H8000000C&
      
      
    End Sub
    
    Private Function GetLastExtractedDate() As Variant        'FUNCTION TO GET dtLastExtractedDate
    
        Dim mCnn                As New ADODB.Connection
        Dim objdb               As New clsDB
        Dim Rec                 As New ADODB.Recordset
        Dim mSql                As String
        Dim dtLastExtractedDate     As Variant
        dtLastExtractedDate = Null
        objdb.SetConnection mCnn
        
        mSql = " SELECT dtLastExtractedDate   FROM faConfig"
        Rec.Open mSql, mCnn
        If Not (Rec.EOF And Rec.BOF) Then
            If Rec!dtLastExtractedDate <> "" Then
                dtLastExtractedDate = DdMmYy(Rec!dtLastExtractedDate)
            End If
            Rec.Close
        End If
        GetLastExtractedDate = dtLastExtractedDate
    End Function
    
    Public Sub MenuDetails()
        Dim objdb   As New clsDB
        Dim mCon    As New ADODB.Connection
        Dim mRec    As New ADODB.Recordset
        Dim mSql    As String
        Dim mMenu   As Control
        Dim mMac    As String
        Dim mIsIKM  As Boolean
        
        'mSQL = "Select vchMenu from faMenumanagements "
        'mSQL = mSQL + " Inner Join faMenuMasters ON faMenuMasters.intMenuID=faMenumanagements.intMenuID "
        'mSQL = mSQL + " where intSeatGroupID=(Select intseatGroupId from faSeatsGroup "
        'mSQL = mSQL + " INNER Join DB_Masters..GL_Seats on faSeatsGroup.intSeatGroupId=DB_Masters..GL_Seats.intGroupId  "
        'mSQL = mSQL + " Where DB_Masters..GL_Seats.numSeatID=" & gbSeatID & ") order by intParentid"

        '------------------------------------------------------------------------------------------------'
        ' Modified by Aiby on 18-Jan-2008
        '------------------------------------------------------------------------------------------------'
        mSql = "Select vchMenu from faMenumanagements "
        mSql = mSql + " Inner Join faMenuMasters ON faMenuMasters.intMenuID=faMenumanagements.intMenuID "
        mSql = mSql + " where intSeatGroupID = " & gbSeatGroupID
        mSql = mSql + " Order By intParentid"

        
        mMac = GetMacAddress
        mIsIKM = IsIKMLAB(mMac)
        Set mRec = objdb.ExecuteSP(mSql, , , , mCon, adCmdText)
        'Note:- Changed By Aiby On 8-Feb-2010
        On Error Resume Next
        For Each mMenu In frmMenu.Controls
            If TypeOf mMenu Is Menu Then
                'Debug.Print mMenu.Name
                mMenu.Enabled = False
                mMenu.Visible = False
                
            End If
        Next
        
        
        While Not (mRec.EOF Or mRec.BOF)
        For Each mMenu In frmMenu.Controls
            If TypeOf mMenu Is Menu Then
                   If mMenu.Name = mRec!vchMenu Then
                        mMenu.Enabled = True
                        mMenu.Visible = True
                        Exit For
                   End If
                   If mIsIKM Then
                        If mMenu.Name = "TeamSaankhya" Then
                             mMenu.Enabled = True
                             mMenu.Visible = True
                        End If
                   End If
            End If
        Next
        mRec.MoveNext
        Wend

        If gbSeatGroupID = gbSeatGroupAccountsOfficer Or gbSeatGroupID = gbSeatGroupCashSuperintended Then 'gbUserTypeID = 2 Or gbUserTypeID = 4 Then
            RequestforPendingTransactions.Caption = "Approval of Pending Transaction Request"
            frmInterruptReceiptStatus.Timer1.Enabled = True
            frmInterruptReceiptStatus.Visible = True
            frmInterruptReceiptStatus.Left = 0
            frmInterruptReceiptStatus.Top = 4920
            RequestforInterruptedReceipt.Caption = "Approval of InterruptedReceipt Request"
'            frmInterruptReceiptStatus.Timer1.Enabled = True
            'frmInterruptReceiptStatus.Picture2.Visible = True
        End If
       
        On Error GoTo 0
        
    End Sub

    Private Sub MDIForm_Resize()
        On Error Resume Next
        frmInterruptReceiptStatus.Top = Me.Height - 2100
        Me.Height = 8700
        Me.Width = 12120
        
    End Sub
    Private Sub MDIForm_Unload(Cancel As Integer)
'        Dim objUser As New clsUser
'        objUser.LogOut
        Dim mCnn    As New ADODB.Connection
        Dim objdb   As New clsDB
        Dim mSql    As String
        If Cancel = 1 Then
            Exit Sub
        Else
            Cancel = 1
            LogOut_Click
        End If
    End Sub

    
    Private Sub Miscellaneousreports_Click()
         frmSoochikaMiscellaneous.Show
    End Sub

    Private Sub mnuAllotments_Click()
        
    End Sub

    Private Sub mnuCounterDayBook_Click()
        
        Dim frmNewRpt As New frmRptViewer
        Dim arInput As Variant
        Dim frmNewViewer As New frmRptViewer
          
        arInput = Array(gbTransactionDate, gbTransactionDate, "%", "%", "%", "%", CStr(gbUserID))
        frmNewViewer.rptFileName = App.Path & "\Reports\rptCounterDayBook.rpt"
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.WindowState = vbMaximized
        frmNewViewer.InputParameters = arInput
        Call frmNewViewer.ShowReport
        frmNewViewer.Show
    End Sub

    Private Sub mnuKMBR_Click()
        frmKMBR.Visible = True
        frmKMBR.ZOrder (0)
    End Sub

    Private Sub mnuSearchSubsidiaryAccountHeads_Click()
        frmSearchSubsidiaryAccountHeads.Show vbModal, Me
    End Sub



    Private Sub mnuSevanaPension_Click()
        'frmSevanaPension.Visible = True
    End Sub

    Private Sub NewInward_Click()
        gbSoochikaVer = gbLinkWithSevana
        If (gbSoochikaVer = 5) Then
            frmUSoochikaInward.Visible = True
            frmUSoochikaInward.ZOrder (0)
        Else
            frmSoochikaInward.Visible = True
        End If
    End Sub

    Private Sub OpeningAppropriationControlRegister_Click()
        'frmOpeningAppCntrRegisterDetails.Visible = True
        'frmOpeningAppCntrRegisterDetails.ZOrder (0)
'''        frmListOfAllotments.LoadMode = 50
'''        frmListOfAllotments.AuthorityOrAllotment = "Opening Appropriation Control Register details"
'''        frmListOfAllotments.Visible = True
'''        frmListOfAllotments.ZOrder (0)
    End Sub

'    Private Sub OpeningBalance_Click()
'        frmAccountHeadsOpeningBalance.Visible = True
'        frmAccountHeadsOpeningBalance.ZOrder (0)
'    End Sub

    Private Sub OpeningBalanceSheet_Click()
        frmOpeningBalace.Visible = True
        frmOpeningBalace.ZOrder (0)
    End Sub
    
    Private Sub OpeningCashBook_Click()
        'If gbFinancialYearID = 2011 Then
        'If OBRPForCuurentVersion Then
            frmOpeningWizard.Visible = True
            frmOpeningWizard.ZOrder (0)
        'End If
        'End If
    End Sub
    Private Function OBRPForCuurentVersion() As Boolean
        Dim mSql        As String
        Dim mCnn        As New ADODB.Connection
        Dim Rec         As New ADODB.Recordset
        Dim mRec         As New ADODB.Recordset
        Dim objdb       As New clsDB
        objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
        mSql = "Select * From faFinancialYear where tinCurrentFinancialYearFlag=1"
        Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
        If Not (Rec.EOF And Rec.BOF) Then
            If Rec!intFinancialYearID = gbFinancialYearID Then
                mSql = "Select * From faVouchers Where intTransactionTypeID in (3007) And intFinancialYearID=" & Rec!intFinancialYearID - 1
                Set Rec = objdb.ExecuteSP(mSql, , , , mCnn, adCmdText)
                If (Rec.EOF And Rec.BOF) Then
                    OBRPForCuurentVersion = True
                End If
            End If
        End If
    End Function

    Private Sub OpeningLetterOfAllotment_Click()
        frmListOfAllotments.LoadMode = 50
        frmListOfAllotments.AuthorityOrAllotment = "OpeningAllotment"
        frmListOfAllotments.Visible = True
        frmListOfAllotments.ZOrder (0)
    End Sub

    Private Sub OpeningLetterOfAuthority_Click()
        frmListOfAllotments.LoadMode = 50
        frmListOfAllotments.AuthorityOrAllotment = "OpeningAuthority"
        frmListOfAllotments.Visible = True
        frmListOfAllotments.ZOrder (0)
    End Sub

    Private Sub OpeningVoucherEntry_Click()
        frmBankScrollOpening.Visible = True
        frmBankScrollOpening.ZOrder (0)
    End Sub

    Private Sub PayBill_Click()
        frmPayBill.Show vbModal
    End Sub

    Private Sub PaymentOrder_Click()
        ''If gbUserTypeID = 1 Or gbUserTypeID = 2 Then    '   Approving   '
        ''    frmPaymentOrder.LoadMode = 2
        ''Else                                            '   Normal      '
        ''    frmPaymentOrder.LoadMode = 1
        ''End If
        ''frmPaymentOrder.LoadMode = 1
        ''frmPaymentOrder.Visible = True
        ''frmPaymentOrder.ZOrder (0)
        
        frmViewPaymentOrder.Visible = True
        frmViewPaymentOrder.ZOrder (0)
    End Sub

    Private Sub PaymentRegister_Click()
        frmViewPayments.Visible = True
        frmViewPayments.ZOrder (0)
    End Sub

    Private Sub Payments_Click()
        ''frmPayments.Visible = True
        ''frmPayments.ZOrder (0)
'        frmIntegratedPayments.Visible = True
'        frmIntegratedPayments.ZOrder (0)

        frmListOfPayments.Visible = True
        frmListOfPayments.ZOrder (0)
    End Sub
  
   
    Private Sub PDEAllotments_Click()
        frmPDEAllotments.Visible = True
        frmPDEAllotments.ZOrder (0)
    End Sub

    Private Sub PortPanchayatData_Click()
        frmPortTransactionsToNewLBs.Show
        frmPortTransactionsToNewLBs.ZOrder (0)
    End Sub

    Private Sub PredateReceiptCancelReport_Click()
        frmRptAdmin.Visible = True
        frmRptAdmin.ZOrder (0)
        
        Call frmRptAdmin.frameVisible
        frmRptAdmin.fmePreviousDatesReceiptCancellation.Visible = True
        frmRptAdmin.Width = 5550
        frmRptAdmin.Height = 1605
    End Sub

    Private Sub Proceedings_Click()
        frmProceedings.chkEdit.Value = 1
        frmProceedings.Show vbModal
    End Sub
    Private Sub Proftax_Click()
        frmBankScrollOpening.Show
        frmBankScrollOpening.ZOrder (0)
    End Sub

    Private Sub ProjectRegister_Click()
        frmProjectDetails.Visible = True
        frmProjectDetails.ZOrder (0)
    End Sub

    Private Sub ProjectVouchers_Click()
        frmListOfProjectVouchers.Show
        frmListOfProjectVouchers.ZOrder (0)
    End Sub

    Private Sub PropertyTax_Click()
       frmPropertyTaxCalculator.Show vbModal
       frmPropertyTaxCalculator.ZOrder (0)
    End Sub
    
    Private Sub LetterOfAuthority_Click()
        frmListOfAllotments.LoadMode = 10
        frmListOfAllotments.AuthorityOrAllotment = "Authority"
        frmListOfAllotments.Visible = True
        frmListOfAllotments.ZOrder (0)
    End Sub

    Private Sub PublishingUtility_Click()
         frmPublishingUtility.Visible = True
         frmPublishingUtility.ZOrder (0)
    End Sub

    Private Sub ReceiptsunderOwnFundMajorHead_Click()
        frmRptSourceWiseReports.Visible = True
        frmRptSourceWiseReports.cmbReportMenu.ListIndex = 3
        frmRptSourceWiseReports.ZOrder (0)
    End Sub

    Private Sub ReceiptsunderOwnFundMinorHead_Click()
        frmRptSourceWiseReports.Visible = True
        frmRptSourceWiseReports.cmbReportMenu.ListIndex = 4
        frmRptSourceWiseReports.ZOrder (0)
    End Sub
    Private Sub RentOnLandAndBuildings_Click()
        frmRentOnLandBuildings.Visible = True
        frmRentOnLandBuildings.ZOrder (0)
    End Sub

    Private Sub RemitBackOfUnUtilizedAmount_Click()
        frmRemitBackofUnUtilizeddrawnAmounts.Visible = True
        frmRemitBackofUnUtilizeddrawnAmounts.ZOrder (0)
    End Sub

    Private Sub ReportGenerator_Click()
        'frmUCheckServer.Visible = True
        'frmUCheckServer.ZOrder (0)
        
        frmReportGenerator.Visible = True
        frmReportGenerator.ZOrder (0)
        
    End Sub
    Private Sub RequestForInterruptedRecEdit_Click()
'        frmListOfInterruptedEditRequests.Visible = True
'        frmListOfInterruptedEditRequests.ZOrder (0)
    End Sub
    Private Sub RequestForInterruptedReceipt_Click()
        frmInterruptedReceiptRequest.Show vbModal
    End Sub
    Private Sub RequestforInterruptedReceiptCancellation_Click()
'        frmInterruptedCancellationRequest.Visible = True
'        frmInterruptedCancellationRequest.ZOrder (0)
    End Sub

    Private Sub RequestForInterruptedReceiptDateEdit_Click()
'        frmInterruptedDateEditRequest.Visible = True
'        frmInterruptedDateEditRequest.ZOrder (0)
    End Sub

    Private Sub RequestforPendingTransactions_Click()
    
        
        If pendingTaskDisable = 1 Then
            MsgBox "Pending Task process is Disabled", vbApplicationModal
            Exit Sub
        Else
            If CheckSubmissionStatus = "$Accepted$$" Or CheckSubmissionStatus = "$Submitted$$" Then
                MsgBox ("AFS Submitted to LFA. No Further Modification is Possible"), vbApplicationModal
                Exit Sub
            ElseIf CheckSubmissionStatus = 2 Or CheckSubmissionStatus = 3 Then
            
                MsgBox ("AFS Submitted to LFA. No Further Modification is Possible"), vbApplicationModal
                Exit Sub
            ElseIf CheckSubmissionStatus = 4 Then
                If pendingTaskDisable = 1 Then
                   MsgBox ("AFS Submitted to LFA. No Further Modification is Possible"), vbApplicationModal
                    Exit Sub
                Else
                  If gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupID = gbSeatGroupAccountsOfficer Or gbSeatGroupID = gbSeatGroupAccountsSuperintended Then
                    frmPendingTasks.Visible = True
                    frmPendingTasks.ZOrder (0)
                  End If
                End If
            Else
            'frmPendingTransactionRequest.Show vbModal
                If gbSeatGroupID = gbSeatGroupAccountsClerk Or gbSeatGroupID = gbSeatGroupAccountsOfficer Or gbSeatGroupID = gbSeatGroupAccountsSuperintended Then
                    frmPendingTasks.Visible = True
                    frmPendingTasks.ZOrder (0)
                End If
            End If
        End If
    End Sub
    Private Sub RequestForReceiptCancellation_Click()
        frmCancelReceipt.LoadMode = 2
        frmCancelReceipt.Show vbModal, Me
    End Sub

    Private Sub RequisitionForFund_Click()
        frmListOfRequisitions.Visible = True
        frmListOfRequisitions.ZOrder (0)
    End Sub



    Private Sub RequisitionInbox_Click()
        frmRequisitionInbox.Visible = True
        frmRequisitionInbox.ZOrder (0)
    End Sub

    Private Sub RequistionRegister_Click()
        frmRequisitionRegister.Visible = True
        frmRequisitionRegister.ZOrder (0)
    End Sub

    Private Sub ReverseEntryList_Click()
        frmListReverseEntryRequests.LoadMode = 1
        frmListReverseEntryRequests.Visible = True
        frmListReverseEntryRequests.ZOrder (0)
    End Sub
    
    Private Sub ReverseEntryRegister_Click()
        frmViewReverseEntryDetails.Visible = True
        frmViewReverseEntryDetails.ZOrder (0)
    End Sub

    Private Sub rptAppropriationControlRegister_Click()
        frmRptAppropriationCrlReg.Visible = True
        frmRptAppropriationCrlReg.ZOrder (0)
    End Sub

    Private Sub rptAssetReplacementRegister_Click()
        frmRptFilterFields.rptNames = 10
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptAuthorisationIssuetoSecretary_Click()
        frmRptFilterFields.rptNames = 11
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptBalanceSheet_Click()
'        frmRptFilterFields.rptNames = 2
'        frmRptFilterFields.Show vbModal
        frmRptFinancialFilterFields.Visible = True
        frmRptFinancialFilterFields.ZOrder (0)
        Call frmRptFinancialFilterFields.cmdBalanceSheet_Click
    End Sub
    
    Private Sub rptBankBook_Click()
'        frmRptFilterFields.rptNames = 6
'        frmRptFilterFields.Show vbModal
        frmRptFinancialFilterFields.Visible = True
        frmRptFinancialFilterFields.ZOrder (0)
        Call frmRptFinancialFilterFields.cmdBankBook_Click
    End Sub

    Private Sub rptBillofReceiptsRegister_Click()
        frmRptFilterFields.rptNames = 12
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptBudgetVariance_Click()
        frmRptFilterFields.rptNames = 44
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptCashBook_Click()
'        frmRptFilterFields.rptNames = 5
'        frmRptFilterFields.Show vbModal
        frmRptFinancialFilterFields.Visible = True
        frmRptFinancialFilterFields.ZOrder (0)
        Call frmRptFinancialFilterFields.cmdCashBook_Click
    End Sub
    
    Private Sub rptCollectionregister_Click()
        frmRptFilterFields.rptNames = 13
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptDemandregister_Click()
        frmRptFilterFields.rptNames = 14
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptDepositreceivedregister_Click()
        frmRptFilterFields.rptNames = 15
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptDocumentcontrolRegister_Click()
        frmRptFilterFields.rptNames = 16
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptFormGEN40Register_Click()
        frmRptFilterFields.rptNames = 17
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptFunctionWiseExpenditure_Click()
        frmRptFilterFields.rptNames = 18
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptFunctionwisereceiptsubsidiaryledger_Click()
        frmRptFilterFields.rptNames = 19
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptFundsReceivedRegister_Click()
        frmRptFilterFields.rptNames = 20
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptImmovablePropertyRegister_Click()
        frmRptFilterFields.rptNames = 21
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptImplentingOfficerwiseAllotmentRegister_Click()
        frmRptImplementingOfficerWiseAllotmentReg.Visible = True
        frmRptImplementingOfficerWiseAllotmentReg.ZOrder (0)
    End Sub

    Private Sub rptIncomeAndExpenditure_Click()
        frmRptFinancialFilterFields.Visible = True
        frmRptFinancialFilterFields.ZOrder (0)
        Call frmRptFinancialFilterFields.cmdIncomeExpenditure_Click
    End Sub
    
    Private Sub rptIncomeandExpenditureRegister_Click()
        frmRptFilterFields.rptNames = 23
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptJournalBook_Click()
'        frmRptFilterFields.rptNames = 7
'        frmRptFilterFields.Show vbModal
        frmRptFinancialFilterFields.Visible = True
        frmRptFinancialFilterFields.ZOrder (0)
        Call frmRptFinancialFilterFields.cmdJounalBook_Click
    End Sub

    Private Sub rptLandRegister_Click()
        frmRptFilterFields.rptNames = 24
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptLedgerBook_Click()
'        frmRptFilterFields.rptNames = 8
'        frmRptFilterFields.Show vbModal
        frmRptFinancialFilterFields.Visible = True
        frmRptFinancialFilterFields.ZOrder (0)
        Call frmRptFinancialFilterFields.cmdLedgerBook_Click
    End Sub
        
    Private Sub rptLedgerView_Click()
        frmRPTLedgerView.Visible = True
        frmRPTLedgerView.Left = 0
        frmRPTLedgerView.Top = 0
        frmRPTLedgerView.ZOrder (0)
    End Sub

    Private Sub rptLetterofallotment_Click()
        frmRptFilterFields.rptNames = 25
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptMemorandumofcollectionRegister_Click()
        frmRptFilterFields.rptNames = 26
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptMovablepropertyRegister_Click()
        frmRptFilterFields.rptNames = 27
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptOfficialReceiptRegister_Click()
        frmRptFilterFields.rptNames = 28
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptPaymentOrderRegister_Click()
        frmRptFilterFields.rptNames = 29
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptProjectregister_Click()
        frmRptFilterFields.rptNames = 30
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptReceiptsAndPayments_Click()
'        frmRptFilterFields.rptNames = 4
'        frmRptFilterFields.Show vbModal
        frmRptFinancialFilterFields.Visible = True
        frmRptFinancialFilterFields.ZOrder (0)
        Call frmRptFinancialFilterFields.cmdReceiptPayment_Click
    End Sub


    Private Sub rptRegisterofadvances_Click()
        frmRptFilterFields.rptNames = 31
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptRegisterofbillsforpayment_Click()
        frmRptFilterFields.rptNames = 32
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptRegisterofPermenantadvance_Click()
        frmRptFilterFields.rptNames = 33
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptRegisterofpubliclightingsystem_Click()
        frmRptFilterFields.rptNames = 34
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptRequesitionforReleaseofFundcodes_Click()
        frmRptFilterFields.rptNames = 35
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptStatementofOutstandingLiabilityforexpenses_Click()
        frmRptFilterFields.rptNames = 36
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptStatementonStatusofChequereceived_Click()
        frmRptFilterFields.rptNames = 37
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptSubsidiaryRegister_Click()
        frmRptFilterFields.rptNames = 38
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptSummaryofCollectionRegister_Click()
        frmRptFilterFields.rptNames = 39
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptSummaryStatementodfbills_Click()
        frmRptFilterFields.rptNames = 40
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptSummaryStatementofDeposits_Click()
        frmRptFilterFields.rptNames = 41
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptSummaryStatementofRefundandRemission_Click()
        frmRptFilterFields.rptNames = 42
        frmRptFilterFields.Show vbModal
    End Sub
    
    Private Sub ChequeIssue_Click()
        frmRptFilterFields.rptNames = 45
        
        frmRptFilterFields.Show vbModal
    End Sub
    Private Sub ChequeReceived_Click()
        frmRptFilterFields.rptNames = 46
        frmRptFilterFields.Show vbModal
    End Sub
    Private Sub BankReconciliation_Click()
        frmRptFilterFields.rptNames = 47
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptSummaryStatementofWriteoffs_Click()
        frmRptFilterFields.rptNames = 43
        frmRptFilterFields.Show vbModal
    End Sub

    Private Sub rptTrialBalance_Click()
'        frmRptFilterFields.rptNames = 1
'        frmRptFilterFields.Show vbModal
        frmRptFinancialFilterFields.Visible = True
        frmRptFinancialFilterFields.ZOrder (0)
        Call frmRptFinancialFilterFields.cmdTrialBalance_Click
    End Sub

    Private Sub SearchAgreements_Click()
        frmSearchAgreements.Visible = True
        frmSearchAgreements.ZOrder (0)
    End Sub

    Private Sub SearchBuilding_Click()
        frmSearchBuildingDetails.Visible = True
        frmSearchBuildingDetails.ZOrder (0)
    End Sub
    
    Private Sub SearchBuildingTaxRemitance_Click()
        frmSearchPropertyTaxFromReceipts.Visible = True
        frmSearchPropertyTaxFromReceipts.ZOrder (0)
    End Sub

    Private Sub SearchInward_Click()
        'If (gbSoochikaVer <> 5) Then
            'frmSoochikaSearch.Show
            'frmSoochikaSearch.ZOrder (0)
        'End If
        'chnaged by soumya v S in Nov13
         Dim mCnn As New ADODB.Connection
         Dim Rec As New ADODB.Recordset
         Dim arrIn As Variant
         Dim arrOut As Variant
         Dim objdb As New clsDB
            
         If (objdb.CreateNewConnection(mCnn, enuSourceString.SoochikaUnicode) = False) Then
          frmSoochikaSearch.Show
          frmSoochikaSearch.ZOrder (0)
          Else
         frmUsoochikaSearch.Show
        End If
  
  
    End Sub

    Private Sub SearchPaymentOrder_Click()
        frmSearchPaymentOrder.Visible = True
        frmSearchPaymentOrder.ZOrder (0)
    End Sub

    Private Sub SearchReceipts_Click()
        frmSearchReceipts.Visible = True
        frmSearchReceipts.ZOrder (0)
    End Sub

    Private Sub SearchTransactions_Click()
        frmSearchVouchers.chkInterrupted.Visible = True
        frmSearchVouchers.Show vbModal
    End Sub


    
    Private Sub SectionWiseTransactionTypes_Click()
        frmSectionWiseTransactionTypes.Visible = True
        frmSectionWiseTransactionTypes.ZOrder (0)
    End Sub
    Private Sub SectorwiseStatement_Click()
        frmRptSourceWiseReports.Visible = True
        frmRptSourceWiseReports.cmbReportMenu.ListIndex = 1
        frmRptSourceWiseReports.ZOrder (0)
    End Sub
    
    Private Sub SecurityRegister_Click()
        Dim vAryInRpt(1)
        vAryInRpt(0) = CStr(Date)
        frmCRViewer.vShowReport App.Path & "\soochika\Reports", "rptCDR.rpt", vAryInRpt
        frmCRViewer.Show 1
    End Sub
    
    Private Sub SendCollectionDetailsToMainOffice_Click()
        frmDemandInterface.Visible = True
        frmDemandInterface.ZOrder (0)
    End Sub
    
    Private Sub SendDailyCollectionToHO_Click()
        frmZonalDaily.Visible = True
        frmZonalDaily.ZOrder (0)
    End Sub

    Private Sub SoochikaMiscellaneousReport_Click()
        frmSoochikaMiscellaneous.Visible = True  'Changed On 25/10/2011 By Poornima
        frmSoochikaMiscellaneous.ZOrder (0)
    End Sub

    Private Sub SourceDeductions_Click()
        frmProjectJournalDetails.LoadMode = 10
        frmProjectJournalDetails.Visible = True
        frmProjectJournalDetails.ZOrder (0)
    End Sub

    Private Sub SourceOfFundEntry_Click()
        frmSourceWiseOpening.Visible = True
        frmSourceWiseOpening.ZOrder (0)
    End Sub

    Private Sub sourcewiseReceiptAndPayment_Click()
        frmRptSourceWiseReports.Visible = True
        frmRptSourceWiseReports.cmbReportMenu.ListIndex = 0
        frmRptSourceWiseReports.ZOrder (0)
    End Sub

    Private Sub StockRegisterReceiptBooks_Click()
        frmStockRegisterOfReceiptBooks.Visible = True
        frmStockRegisterOfReceiptBooks.ZOrder (0)
    End Sub
    Private Sub SubsidiaryAccount_Click()
        frmCreateSubsidiaryAccountHeads.Visible = True
        frmCreateSubsidiaryAccountHeads.ZOrder (0)
    End Sub

    Private Sub SubsidiaryCashBook_Click()
        frmListOfSubsidiaryCashTransfers.Visible = True
        frmListOfSubsidiaryCashTransfers.ZOrder (0)
    End Sub

    Private Sub Sulekha_Click()
        frmSulekhaIntegration.Visible = True
        frmSulekhaIntegration.ZOrder (0)
    End Sub

    Private Sub SynchronizeDetails_Click()
    
    MsgBox "THIS FUNCTIONALITY IS NO MORE AVAILABLE", vbInformation
'        frmPortSulekha.Visible = True
'        frmPortSulekha.ZOrder (0)
    End Sub

    Private Sub SynchronizeProjectMaster_Click()
'        frmSynchronizeProjectMaster.Show vbModal
    End Sub
    
Private Sub TeamSaankhya_Click()
    frmTeamSaankhya.Visible = True
    frmTeamSaankhya.ZOrder (0)
End Sub

    Private Sub Test_Click()
        frmTest2.Visible = True
        frmTest2.ZOrder (0)

    End Sub
    
    Private Sub Test3_Click()
'        frmTest2.Visible = True
'        frmTest2.ZOrder (0)
'        frmPortTransactionsToNewLBs.Visible = True
'        frmPortTransactionsToNewLBs.ZOrder (0)
        frmDatabaseConnection.Visible = True
        frmDatabaseConnection.ZOrder (0)
    End Sub

    Private Sub TestReport_Click()
        frmRptViewer.Visible = True
        frmRptViewer.ZOrder (0)
    End Sub
    
    Private Sub Timer1_Timer()
        If mTimer = 0 Then
            lblPreyear.Visible = True
            imgWarning.Visible = True
            mTimer = 1
            If gbTransactionDate <> Date Then
                lblSplash.Visible = True
            End If
            
            Exit Sub
        End If
        If mTimer = 1 Then
            lblPreyear.Visible = False
            imgWarning.Visible = False
            mTimer = 0
            If gbTransactionDate <> Date Then
                lblSplash.Visible = False
            End If
            
            Exit Sub
        End If
    End Sub

    Private Sub TotalReceiptCount_Click()
        frmRptAdmin.Visible = True
        frmRptAdmin.ZOrder (0)

        Call frmRptAdmin.frameVisible
        frmRptAdmin.fmeReceiptCount.Visible = True
        frmRptAdmin.fmeReceiptCount.Left = 0
        frmRptAdmin.fmeReceiptCount.Top = 0
        frmRptAdmin.Width = 5445
        frmRptAdmin.Height = 2005
    End Sub

    Private Sub TransactionTemp_Click()
       frmReconBankList.Visible = True
       frmReconBankList.ZOrder (0)
    End Sub
    
'  Commented on 22/10/2019
    
'    Private Sub TransactionType_Click()
'        frmTransactionTypes.Visible = True
'        frmTransactionTypes.ZOrder (0)
'    End Sub

    Private Sub TransactionTypeMapping_Click()
        frmTransactionsMapping.Visible = True
        frmTransactionsMapping.ZOrder (0)
    End Sub

    Private Sub TreasuryBalanceFinalization_Click()
        frmCBSourceofFundTreasury.Visible = True
        frmCBSourceofFundTreasury.ZOrder (0)
    End Sub

    Private Sub UnAuthorizedDrawal_Click()
        frmListOfRequisitions.LoadMode = 10
        frmListOfRequisitions.Visible = True
        frmListOfRequisitions.ZOrder (0)
    End Sub

    Private Sub UpdateOpeningBalance_Click()
        frmUpdateAccountHeadBalance.Visible = True
        frmUpdateAccountHeadBalance.ZOrder (0)
    End Sub
   


    Private Sub viewHeadwiseConsolidation_Click()
       frmViewHeadWiseConsolidation.Visible = True
       frmViewHeadWiseConsolidation.ZOrder (0)
    End Sub

    Private Sub ViewPaymentOrder_Click()
'        frmViewPaymentOrder.Top = 100
'        frmViewPaymentOrder.Left = (Me.Width - frmViewPaymentOrder.Width) / 2
'        frmViewPaymentOrder.Visible vbModal, frmMenu
        frmViewPaymentOrder.Visible = True
        frmViewPaymentOrder.ZOrder (0)
    End Sub
    
Private Sub ViewPaymentOrderDetails_Click()
        frmViewPaymentOrder.ViewMode = 50
        frmViewPaymentOrder.Visible = True
        frmViewPaymentOrder.ZOrder (0)
End Sub



    Private Sub ViewRequisitionRegister_Click()
        frmViewRequisitionRegister.Visible = True
        frmViewRequisitionRegister.ZOrder (0)
    End Sub

    Private Sub ViewSubsidiaryCashBook_Click()
        frmViewSubsidiaryCashBook.Visible = True
        frmViewSubsidiaryCashBook.ZOrder (0)
    End Sub
    
    Private Sub ViewVouchers_Click()
        frmViewVoucher.Show vbModal
    End Sub

    Private Sub VoucherExtractStatus_Click()
        frmViewVoucherExtractStatus.Visible = True
        frmViewVoucherExtractStatus.ZOrder (0)
    End Sub

'    Private Sub VoucherUtility_Click()
'        Dim mPassword As String
'
'        mPassword = InputBox("Enter the Utility Password", "Utility PassWord")
'        If mPassword = "DingDong" Then
'            frmVoucherUtility.Visible = True
'            frmVoucherUtility.ZOrder (0)
'        Else
'            MsgBox "Wrong PassWord"
'            Exit Sub
'        End If
'    End Sub

    Private Sub Wardwise_Click()
        frmRptAdmin.Visible = True
        frmRptAdmin.ZOrder (0)
                
        Call frmRptAdmin.frameVisible
        frmRptAdmin.fmeWardWiseTransactionTypes.Visible = True
        frmRptAdmin.Width = 5535
        frmRptAdmin.Height = 2475
    End Sub
    Private Sub DepartmentwiseReport_Click()
        frmRptAdmin.Visible = True
        frmRptAdmin.ZOrder (0)
                
        Call frmRptAdmin.frameVisible
        frmRptAdmin.fmeDept.Visible = True
        frmRptAdmin.Width = 5520
        frmRptAdmin.Height = 2040
    End Sub
    Private Sub CancelledCounterReceipts_Click()
        frmRptAdmin.Visible = True
        frmRptAdmin.ZOrder (0)
        
        Call frmRptAdmin.frameVisible
        frmRptAdmin.fmeCancelled.Visible = True
        frmRptAdmin.Width = 5565
        frmRptAdmin.Height = 1500
    End Sub
    Private Sub TotalCounterCollection_Click()
        frmRptAdmin.Visible = True
        frmRptAdmin.ZOrder (0)
        
        Call frmRptAdmin.frameVisible
        frmRptAdmin.fmeTotalCounterCollection.Visible = True
        frmRptAdmin.Width = 5550
        frmRptAdmin.Height = 1605
    End Sub
    Private Sub TotalHeadwiseConsolidation_Click()
        frmRptAdmin.Visible = True
        frmRptAdmin.ZOrder (0)
        
        Call frmRptAdmin.frameVisible
        frmRptAdmin.fmeTotalHeadWiseCollection.Visible = True
        frmRptAdmin.Width = 5550
        frmRptAdmin.Height = 1635
    End Sub
    Private Sub DailyCounterConsolidation_Click()
        frmRptAdmin.Visible = True
        frmRptAdmin.ZOrder (0)
        
        Call frmRptAdmin.frameVisible
        frmRptAdmin.fmeDailyCounterConsolidation.Visible = True
        frmRptAdmin.Width = 5550
        frmRptAdmin.Height = 1620
    End Sub
    Private Sub YearEndProcedure_Click()
        frmYearEndProcess.Visible = True
        frmYearEndProcess.ZOrder (0)
    End Sub

    Private Sub YearEndProcess_Click()
        frmYearEndProcess.Visible = True
        frmYearEndProcess.ZOrder (0)
    End Sub
    
    Public Property Let InterruptedReceiptModeIR(ByVal mData As Boolean)
        mInterruptedReceiptID = mData
    End Property
    
    Private Sub ZonalCollection_Click()
    
    '''Old Report
        frmRptAdmin.Visible = True
        frmRptAdmin.ZOrder (0)
        Call frmRptAdmin.frameVisible
        frmRptAdmin.fmeZonalCollection.Visible = True
'        frmRptAdmin.Width = 5545
'        frmRptAdmin.Height = 9330

        frmRptAdmin.Width = 5545
        frmRptAdmin.Height = 1620
    End Sub

    Private Sub ZonalIntegration_Click()
        frmZonalMain.Visible = True
        frmZonalMain.ZOrder (0)
    End Sub
