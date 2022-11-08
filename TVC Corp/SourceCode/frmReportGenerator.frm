VERSION 5.00
Begin VB.Form frmReportGenerator 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "   ~~  Reports  ~~"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCheckReconciliation 
      Caption         =   "Check Reconciliation"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1755
      TabIndex        =   19
      Top             =   8340
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ledger Book"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1755
      TabIndex        =   18
      Top             =   7785
      Width           =   2295
   End
   Begin VB.CommandButton cmdBankReconciliation 
      Caption         =   "Bank Reconciliation"
      Height          =   420
      Left            =   1755
      TabIndex        =   17
      Top             =   7290
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Income && Expenditure"
      Default         =   -1  'True
      Height          =   390
      Left            =   1770
      TabIndex        =   16
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton cmdChequeRegister 
      Caption         =   "Cheque Register"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1785
      TabIndex        =   15
      Top             =   6360
      Width           =   2310
   End
   Begin VB.CommandButton cmdPropertyTax 
      Caption         =   "Property Tax Register"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1785
      TabIndex        =   14
      Top             =   5880
      Width           =   2310
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Deposit Register"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4275
      TabIndex        =   13
      Top             =   4905
      Width           =   1260
   End
   Begin VB.CommandButton cmdMonthlyConslidation 
      Caption         =   "Monthly Consolidation"
      Height          =   435
      Left            =   1785
      TabIndex        =   12
      Top             =   5385
      Width           =   2205
   End
   Begin VB.CommandButton cmdLedgerBook 
      Caption         =   "Ledger Book"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2820
      TabIndex        =   11
      Top             =   4905
      Width           =   1260
   End
   Begin VB.TextBox txtHeadCode 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Top             =   4995
      Width           =   1320
   End
   Begin VB.CommandButton cmdTrialBalanceConsolidation 
      Caption         =   "Trial Balance (Consolidation)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1830
      TabIndex        =   8
      Top             =   3000
      Width           =   2310
   End
   Begin VB.CommandButton cmdBalanceSheet 
      Caption         =   "Balance Sheet"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1830
      TabIndex        =   7
      Top             =   4095
      Width           =   2310
   End
   Begin VB.CommandButton cmdIncomeAndExpediture 
      Caption         =   "Income and Expenditure"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1830
      TabIndex        =   6
      Top             =   3540
      Width           =   2310
   End
   Begin VB.TextBox txtDate 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2070
      TabIndex        =   5
      Top             =   195
      Width           =   1800
   End
   Begin VB.CommandButton cmdOpeningBalance 
      Caption         =   "Opening Balance"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1830
      TabIndex        =   3
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdTrialBalance 
      Caption         =   "Trial Balance"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1830
      TabIndex        =   2
      Top             =   2460
      Width           =   2310
   End
   Begin VB.CommandButton cmdRPConsolidation 
      Caption         =   "R && P (Major Head )"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1830
      TabIndex        =   1
      Top             =   1920
      Width           =   2310
   End
   Begin VB.CommandButton cmdReceiptsAndPayments 
      Caption         =   "Receipts && Payments"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1830
      TabIndex        =   0
      Top             =   1320
      Width           =   2310
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Head Code"
      Height          =   195
      Left            =   540
      TabIndex        =   9
      Top             =   5040
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      Height          =   195
      Left            =   1515
      TabIndex        =   4
      Top             =   240
      Width           =   345
   End
End
Attribute VB_Name = "frmReportGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBalanceSheet_Click()
    
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQL As String
    Dim RecBalance As New ADODB.Recordset
    Dim varInPut As Variant
    Dim mFlag As Boolean
    
    Dim mTotalI As Currency
    Dim mTotalE As Currency
    
    FileInitialize
    
    objDB.SetConnection mCnn
    mSQL = "Select * From faAccountHeads Where tinType IN (3, 4)  Order By tinType, faAccountHeads.vchAccountHeadCode"
    Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
    
    Print #gbFileNO, "  Balance Sheet "
    Print #gbFileNO, "======================================"
    While Not Rec.EOF
        If Not mFlag And Rec!tinType = 4 Then
                mFlag = True
                Print #gbFileNO, Tab(45); "=============================================="
                Print #gbFileNO, Tab(11); PadL("A:  LIABILITIES", 30); Tab(45); PadL(Format(mTotalI, "0.00"), 14)
                Print #gbFileNO, Tab(45); "=============================================="
        End If
        varInPut = Array(Rec!intAccountHeadID)
        Set RecBalance = objDB.ExecuteSP("spGetClosingBalance", varInPut, , , mCnn, adCmdStoredProc)
        If Not (RecBalance.EOF And RecBalance.BOF) Then
            If RecBalance!Balance <> 0 Then
                Print #gbFileNO, Rec!vchAccountHeadCode; "  "; PadL(Rec!vchAccountHead, 30);
                If Rec!tinType = 3 Then
                    Print #gbFileNO, Tab(45); PadL(Format(RecBalance!Balance * -1, "0.00"), 14);
                    mTotalI = mTotalI + Format(RecBalance!Balance * -1, "0.00")
                Else
                    Print #gbFileNO, Tab(45); PadL(Format(RecBalance!Balance, "0.00"), 14);
                    mTotalE = mTotalE + Format(RecBalance!Balance, "0.00")
                End If
                Print #gbFileNO,
            End If
        End If
        RecBalance.Close
        Rec.MoveNext
    Wend
    Rec.Close
    
    Print #gbFileNO, Tab(45); "=============================================="
    Print #gbFileNO, Tab(11); PadL("B:   ASSETS", 30); Tab(45); PadL(Format(mTotalE, "0.00"), 14)
    Print #gbFileNO, Tab(45); "=============================================="
    
    Close #gbFileNO
    ShellPad
    
End Sub

Private Sub cmdBankReconciliation_Click()
    Dim objDB As New clsDB
    Dim Rec As New ADODB.Recordset
    Dim mCnn As New ADODB.Connection
    Dim mSQL As String
    Dim mSQL1 As String
    Dim mPayment As Double
    Dim mReceipt As Double
    Dim mCredit As Double
    Dim mDebit As Double
    
    objDB.SetConnection mCnn
    
    
    mSQL = "Select dtDate, intVoucherID, vchInstrumentNo, fltAmount, tnyVoucherTypeID From faVouchers Where tnyReconciled is Null AND intKeyID1 = 1506 AND dtDate BETWEEN '1-Apr-2008' AND '30-Apr-2008' "
    FileInitialize
    Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic
    While Not Rec.EOF
        Print #gbFileNO, Rec!intVoucherID, Rec!dtDate, Rec!fltAmount
        If (Rec!tnyVoucherTypeID = 20) Then
            mPayment = mPayment + Rec!fltAmount
        ElseIf (Rec!tnyVoucherTypeID = 10) Then
            mReceipt = mReceipt + Rec!fltAmount
        End If
        Rec.MoveNext
    Wend
    
    Print #gbFileNO, "____________Total______________"
    Print #gbFileNO, mPayment, mReceipt
    Rec.Close


    mSQL1 = "Select intReconciliationid,vchChequeNo,dtBankEntryDate,fltDrAmount,fltCrAmount From faBankReconciliationEntries WHERE tnyReconciled is Null AND dtBANKENTRyDate BETWEEN '1-Apr-2008' AND '30-Sep-2008'"
    Rec.Open mSQL1, mCnn, adOpenDynamic, adLockOptimistic
    While Not Rec.EOF
        Print #gbFileNO, Rec!intReconciliationID, Rec!vchChequeNo, Rec!dtBankEntryDate, Rec!fltDrAmount, Rec!fltCrAmount
        mCredit = mCredit + IIf(IsNull(Rec!fltCrAmount), 0, Rec!fltCrAmount) '''''IIf(IsNull(Rec!fltCrAmount), 0)
        mDebit = mDebit + IIf(IsNull(Rec!fltDrAmount), 0, Rec!fltDrAmount)
        Rec.MoveNext
    Wend
    Print #gbFileNO, "_________BANK TOTAL_________"
    Print #gbFileNO, mCredit, mDebit
    Close #gbFileNO
    ShellPad
End Sub

Private Sub cmdCheckReconciliation_Click()
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    
    Dim Rec As New ADODB.Recordset
    Dim RecSum As New ADODB.Recordset
    
    Dim mSQL As String
    Dim mTokenID As Long
    Dim mTotalAmt As Double
    Dim mBankAmount As Double
    Dim mCount As Long
    
    objDB.SetConnection mCnn
    mSQL = "Select * From faBankReconciliationEntries Where intBankAccountHeadID = 1506 And tnyReconciled > 0" ' AND intReconciliationID Between 300 and 400"
    Rec.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic
    FileInitialize
    While Not Rec.EOF
        mCount = mCount + 1
        'If mCount = 100 Then GoTo SkipLoop
        mTotalAmt = 0
        mSQL = "Select Sum(fltAmount) fltAmount From faOpeningVouchers Where intAccountHeadID = 1506 and tnyReconciled > 0 AND numTockenID = " & Rec!intReconciliationID
        RecSum.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic
        If IsNumeric(RecSum!fltAmount) Then
            mTotalAmt = RecSum!fltAmount
        End If
        RecSum.Close
        
        mSQL = "Select Sum(faTransactionChild.fltAmount) fltAmount From faVouchers Inner Join "
        mSQL = mSQL + " faTransactions On faTransactions.intVoucherID = faVouchers.intVoucherID Inner Join"
        mSQL = mSQL + " faTransactionChild On faTransactionChild.intTransactionID = faTransactions.intTransactionID"
        mSQL = mSQL + " Where intAccountHeadID = 1506 And tnyReconciled > 0 And numTockenID = " & Rec!intReconciliationID
        RecSum.Open mSQL, mCnn, adOpenDynamic, adLockOptimistic
        If IsNumeric(RecSum!fltAmount) Then
            mTotalAmt = mTotalAmt + RecSum!fltAmount
        End If
        RecSum.Close
            
        If IsNumeric(Rec!fltDrAmount) Then
            mBankAmount = Rec!fltDrAmount
        Else
            mBankAmount = Rec!fltCrAmount
        End If
        If mBankAmount = mTotalAmt Then
            'Print #gbFileNO, Rec!intReconciliationID, , mBankAmount, mTotalAmt
        Else
            Print #gbFileNO, "Matching  " & Rec!intReconciliationID, , mBankAmount, mTotalAmt
        End If
        Rec.MoveNext
    Wend
SkipLoop:
    Close #gbFileNO
    ShellPad
    
End Sub

Private Sub cmdChequeRegister_Click()
    Dim objAc As New clsAccounts
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQL As String
    Dim varInPut As Variant
    Dim mFlag As Boolean
    
    Dim mTotalDr As Currency
    Dim mTotalCr As Currency
    
    Dim mDailyDrTotal As Currency
    Dim mDailyCrTotal As Currency
    
    
    Dim mDt As Date
    Dim mPageNo As Integer
    Dim mLineNo As Integer
    Dim mID As Long
    
        FileInitialize
        
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO, Tab(10); objAc.AccountCode, objAc.AccountHead
        
        Print #gbFileNO, "--------------------------------------------------------------------------------------------------------------------------------------------"
        Print #gbFileNO,
        Print #gbFileNO, "---------------------------------------------------------------------------------------------------------------------------------------------"
  
        
        mLineNo = 7
        objDB.SetConnection mCnn
        Set Rec = objDB.ExecuteSP("spRptLedgerForDOS2", , , , mCnn, adCmdStoredProc)
        If Not (Rec.BOF And Rec.EOF) Then
            Dim mVoucherNo As String
            Dim mInstrumentNo As String
            Dim mInstDate As String
            Dim mNameOfBank As String
            Dim mPlaceOfBank As String
            Dim mName As String
            Dim mHName As String
            Dim mID2 As Double
            mID = -1
            While Not Rec.EOF
                mDt = Rec!dtDate
                
                If Not IsNull(Rec!intVoucherID) Then
                    mID2 = Rec!intVoucherID
                Else
                    mID2 = 0
                End If
                If mID <> mID2 Then
                        mID = mID2
                        If Not IsNull(Rec.Fields(15)) Then
                            mVoucherNo = Rec.Fields(15)
                        Else
                            mVoucherNo = "-"
                        End If
                        
                        If Not IsNull(Rec!vchInstrumentNo) Then
                            mInstrumentNo = Rec!vchInstrumentNo
                        Else
                            mVoucherNo = "-"
                        End If
                        
                        If IsDate(Rec!dtInstrumentDate) Then
                            mInstDate = DdMmmYy(Rec!dtInstrumentDate)
                        Else
                            mInstDate = "-"
                        End If
                        
                        If Not IsNull(Rec!vchBank) Then
                            mNameOfBank = Rec!vchBank
                        Else
                            mNameOfBank = ""
                        End If
                        
                        If Not IsNull(Rec!vchBankPlace) Then
                            mPlaceOfBank = Rec!vchBankPlace
                        Else
                            mPlaceOfBank = "-"
                        End If
                        
                        If Not IsNull(Rec!vchName) Then
                            mName = Rec!vchName
                        Else
                            mName = "-"
                        End If
                        
                        If Not IsNull(Rec!vchHouseName) Then
                            mHName = Rec!vchHouseName & ", "
                        Else
                            mHName = "-"
                        End If
                        
                        Print #gbFileNO,
                        Print #gbFileNO,
                        Print #gbFileNO, Column(7, DdMmmYy(mDt), 11, mVoucherNo, 11, (mInstrumentNo + "/" + mInstDate), 18, mNameOfBank, 20, mPlaceOfBank, 20)
                        'Print #gbFileNO,
                        'Print #gbFileNO, DdMmmYy(mDt); " "; PadL(mVoucherNo, 11); " ";
                        'Print #gbFileNO, PadR(mInstrumentNo + " /" + mInstDate, 15); "      Amount : "; Rec!fltAmount
                        Print #gbFileNO, "               "; mName; mHName
                        Print #gbFileNO, "               "; "("; Rec!vchDescription; ")"
                End If
                
                Print #gbFileNO, "               "; Rec!vchAccountHeadCode; "  "; PadR(Rec!vchAccountHead, 30); "  "; PadL(Format(Rec!fltAmt, "0.00"), 10)
                Rec.MoveNext
            Wend
        End If
        Close #gbFileNO
        ShellPad
End Sub

Private Sub cmdIncomeAndExpediture_Click()
    
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQL As String
    Dim RecBalance As New ADODB.Recordset
    Dim varInPut As Variant
    Dim mFlag As Boolean
    
    Dim mTotalI As Currency
    Dim mTotalE As Currency
    
    FileInitialize
    
    objDB.SetConnection mCnn
    mSQL = "Select * From faAccountHeads Where tinType IN (1, 2)  Order By tinType, faAccountHeads.vchAccountHeadCode"
    Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
    
    Print #gbFileNO, "   Income & Expenditure Statements    "
    Print #gbFileNO, "======================================"
    While Not Rec.EOF
        
        
        If Not mFlag And Rec!tinType = 2 Then
                mFlag = True
                Print #gbFileNO, Tab(45); "=============================================="
                Print #gbFileNO, Tab(11); PadL("A:  TOTAL INCOME", 30); Tab(45); PadL(Format(mTotalI, "0.00"), 14)
                Print #gbFileNO, Tab(45); "=============================================="
        End If
        
        
        If Left(Rec!vchAccountHeadCode, 3) = "280" Then
            GoTo lblPriorPeriod:
        End If
        varInPut = Array(Rec!intAccountHeadID)
        Set RecBalance = objDB.ExecuteSP("spGetClosingBalance", varInPut, , , mCnn, adCmdStoredProc)
        If Not (RecBalance.EOF And RecBalance.BOF) Then
            If RecBalance!Balance <> 0 Then
                Print #gbFileNO, Rec!vchAccountHeadCode; "  "; PadL(Rec!vchAccountHead, 30);
                Print #gbFileNO, Tab(45); PadL(Format(Abs(RecBalance!Balance), "0.00"), 14);
                If Rec!tinType = 1 Then
                    mTotalI = mTotalI + Format(Abs(RecBalance!Balance), "0.00")
                Else
                    mTotalE = mTotalE + Format(RecBalance!Balance, "0.00")
                End If
                Print #gbFileNO,
            End If
        End If
        RecBalance.Close
        
        
            
        Rec.MoveNext
    Wend
    
    
    
lblPriorPeriod:

    Print #gbFileNO, Tab(45); "=============================================="
    Print #gbFileNO, Tab(11); PadL("B:   TOTAL EXPENDITURE", 30); Tab(45); PadL(Format(mTotalE, "0.00"), 14)
    Print #gbFileNO, Tab(45); "=============================================="
    
    Print #gbFileNO, "A-B : Gross Surplus/(Deficit) of"
    Print #gbFileNO, "      Income over Expediture Before"
    Print #gbFileNO, "      Prior period items";
    Print #gbFileNO, Tab(45); PadL(Format(mTotalI - mTotalE, "0.00"), 14)
    Print #gbFileNO,
    
    Print #gbFileNO, "Add: Prior PeriodItems (Net)"
    
    While Not Rec.EOF
        If Not mFlag And Rec!tinType = 2 Then
                mFlag = True
                Print #gbFileNO, Tab(45); "=============================================="
                Print #gbFileNO, Tab(45); PadL(Format(mTotalI, "0.00"), 14)
                Print #gbFileNO, Tab(45); "=============================================="
        End If
        
        varInPut = Array(Rec!intAccountHeadID)
        Set RecBalance = objDB.ExecuteSP("spGetClosingBalance", varInPut, , , mCnn, adCmdStoredProc)
        If Not (RecBalance.EOF And RecBalance.BOF) Then
            If RecBalance!Balance <> 0 Then
                Print #gbFileNO, Rec!vchAccountHeadCode; "  "; PadL(Rec!vchAccountHead, 30);
                Print #gbFileNO, Tab(45); PadL(Format(Abs(RecBalance!Balance), "0.00"), 14);
                mTotalE = mTotalE - Format(RecBalance!Balance, "0.00")
                
                Print #gbFileNO,
            End If
        End If
        RecBalance.Close
        Rec.MoveNext
    Wend
    
    Rec.Close
    
    
    Print #gbFileNO, Tab(45); "=============================================="
    Print #gbFileNO, Tab(45); PadL(Format(mTotalI - mTotalE, "0.00"), 14)
    Print #gbFileNO, Tab(45); "=============================================="
    Print #gbFileNO,
    Print #gbFileNO, "Net balance being surplus/deficit carried"
    Print #gbFileNO, "over to Municipal Fund"


    Close #gbFileNO
    ShellPad
    
    
End Sub

Private Sub cmdLedgerBook_Click()
    Dim objAc As New clsAccounts
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQL As String
    Dim varInPut As Variant
    Dim mFlag As Boolean
    
    Dim mTotalDr As Currency
    Dim mTotalCr As Currency
    
    Dim mDailyDrTotal As Currency
    Dim mDailyCrTotal As Currency
    
    
    Dim mDt As Date
    Dim mPageNo As Integer
    Dim mLineNo As Integer
    Dim mLoop As Integer
    
    objAc.SetAccountCode (Trim(txtHeadCode))
    If objAc.AccountHeadID > 0 Then
        If IsDate(txtDate) Then
            varInPut = Array(objAc.AccountHeadID, "1-Apr-2008", txtDate)
        Else
            varInPut = Array(objAc.AccountHeadID)
        End If
        FileInitialize
        
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO, Tab(10); objAc.AccountCode, objAc.AccountHead
        Print #gbFileNO, "---------------------------------------------------------------------------------------------------------------------------------"
        Print #gbFileNO, "Vou.ID Trn. Date   Head Code  Account Head                      Debit       Credit      Balance Type Tr.ID   Cheque No.\Date"
        Print #gbFileNO, "---------------------------------------------------------------------------------------------------------------------------------"
        mLineNo = 7
        objDB.SetConnection mCnn
        Set Rec = objDB.ExecuteSP("spRptLedgerForDOS", varInPut, , , mCnn, adCmdStoredProc)
        If Not (Rec.BOF And Rec.EOF) Then
            mDt = Rec!dtTransactionDate
            While Not Rec.EOF
                
                If mLineNo >= 67 Then
                    mPageNo = mPageNo + 1
                    mLineNo = 0
                    Print #gbFileNO,
                    Print #gbFileNO, Tab(100); "Page No:"; mPageNo
                    Print #gbFileNO,
                    Print #gbFileNO,
                    Print #gbFileNO, ' Page Ending on 72 Line
                    
                    
                    Print #gbFileNO,
                    Print #gbFileNO,
                    Print #gbFileNO,
                    Print #gbFileNO, Tab(10); objAc.AccountCode, objAc.AccountHead
                    Print #gbFileNO, "---------------------------------------------------------------------------------------------------------------------------------"
                    Print #gbFileNO, "Vou.ID Trn. Date   Head Code  Account Head                      Debit       Credit      Balance Type Tr.ID   Cheque No.\Date"
                    Print #gbFileNO, "---------------------------------------------------------------------------------------------------------------------------------"
                    mLineNo = 7
                End If
                
                Print #gbFileNO, PadL(str(Rec!intVoucherID), 7); " ";
                Print #gbFileNO, DdMmmYy(Rec!dtTransactionDate); " ";
                Print #gbFileNO, Rec!vchAccountHeadCode; "  ";
                Print #gbFileNO, PadR(Rec!vchAccountHead, 25); "|";
                'If Rec!tinDebitOrCreditFlag = 1 Then
                If Rec!tinDrOrCr = 1 Then
                    mTotalCr = mTotalCr + Format(Rec!fltAmount, "0.00")
                    mDailyCrTotal = mDailyCrTotal + Format(Rec!fltAmount, "0.00")
                    Print #gbFileNO, Tab(71); PadL(Format(Rec!fltAmount, "0.00"), 12);
                Else
                    mTotalDr = mTotalDr + Format(Rec!fltAmount, "0.00")
                    mDailyDrTotal = mDailyDrTotal + Format(Rec!fltAmount, "0.00")
                    Print #gbFileNO, Tab(58); PadL(Format(Rec!fltAmount, "0.00"), 12);
                End If
                
                Print #gbFileNO, Tab(84); "|"; PadL(Format(mTotalDr - mTotalCr, "0.00"), 11); "   ";
                Print #gbFileNO, Rec!vchGroup; "  "; PadL(Trim(str(Rec!intTransactionID)), 4); "   "; IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo);
                If IsDate(Rec!dtInstrumentDate) Then
                    Print #gbFileNO, "\" + DdMmmYy(Rec!dtInstrumentDate)
                Else
                    Print #gbFileNO,
                End If
                mLineNo = mLineNo + 1
                
                
                Rec.MoveNext
                If Not Rec.EOF Then
                    If mDt <> Rec!dtTransactionDate Then
                        mDt = Rec!dtTransactionDate
                        If (mLineNo + 4) >= 67 Then
                            For mLoop = mLineNo To 67
                                Print #gbFileNO,
                            Next
                            mPageNo = mPageNo + 1
                            mLineNo = 0
                            Print #gbFileNO,
                            Print #gbFileNO, Tab(100); "Page No:"; mPageNo
                            Print #gbFileNO,
                            Print #gbFileNO,
                            Print #gbFileNO, ' Page Ending on 72 Line
                            
                            
                            Print #gbFileNO,
                            Print #gbFileNO,
                            Print #gbFileNO,
                            Print #gbFileNO, Tab(10); objAc.AccountCode, objAc.AccountHead
                            Print #gbFileNO, "---------------------------------------------------------------------------------------------------------------------------------"
                            Print #gbFileNO, "Vou.ID Trn. Date   Head Code  Account Head                      Debit       Credit      Balance Type Tr.ID   Cheque No.\Date"
                            Print #gbFileNO, "---------------------------------------------------------------------------------------------------------------------------------"
                            mLineNo = 7
                        End If
                        
                        
lblSubTotal:
                        Print #gbFileNO, Tab(56); "=============================="
                        Print #gbFileNO, Tab(56); PadL(Format(mDailyDrTotal, "0.00"), 14);
                        Print #gbFileNO, Tab(71); PadL(Format(mDailyCrTotal, "0.00"), 14);
                        Print #gbFileNO, Tab(56); "=============================="
                        If (mTotalDr - mTotalCr) > -1 Then
                            Print #gbFileNO, Tab(45); "Balance  : "; Tab(56); PadL(Format(mTotalDr - mTotalCr, "0.00"), 14)
                        Else
                            Print #gbFileNO, Tab(45); "Balance  : "; Tab(71); PadL(Format(mTotalCr - mTotalDr, "0.00"), 14)
                        End If
                        mDailyDrTotal = 0
                        mDailyCrTotal = 0
                         mLineNo = mLineNo + 4
                    End If
                Else
                    GoTo lblSubTotal:
                End If
            Wend
        End If
        Close #gbFileNO
        ShellPad
    End If
    
End Sub

Private Sub cmdMonthlyConslidation_Click()
 
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQL As String
    Dim RecBalance As New ADODB.Recordset
    Dim varInPut As Variant
    Dim mFlag As Boolean
    
    Dim mTotalDr As Currency
    Dim mTotalCr As Currency
    
    mSQL = "Select faVoucherChild.intAccountHeadID , faAccountHeads.vchAccountHeadCode , faAccountHeads.vchAccountHead, Sum(faVoucherChild.fltAmount) Amount"
    mSQL = mSQL + " from faVoucherChild Inner Join"
    mSQL = mSQL + "     faVouchers On faVouchers.intVoucherID = faVoucherChild.intVoucherID Inner Join"
    mSQL = mSQL + "     faAccountHeads on faAccountHeads.intAccountHeadID = faVoucherChild.intAccountHeadID"
    mSQL = mSQL + " Where dtDate Between '01-Dec-2008' ANd '31-Dec-2008' And tnyCancelFlag <> 1  Group By faVoucherChild.intAccountHeadID , faAccountHeads.vchAccountHeadCode , faAccountHeads.vchAccountHead"
    mSQL = mSQL + " Order By vchAccountHeadCode"
    
    FileInitialize
    
    objDB.SetConnection mCnn
    Rec.Open mSQL, mCnn
    Print #gbFileNO, Heading("Kozhikode Corporation ", , True, True)
    Print #gbFileNO,
    Print #gbFileNO, Heading("Monthly Head wise Consolidation", , True)
    Print #gbFileNO, Heading("===================================")
    Print #gbFileNO,
Print #gbFileNO, "---------------------------------------------------------------------------"
Print #gbFileNO, "Head Code  Account Head                                              Amount"
Print #gbFileNO, "---------------------------------------------------------------------------"
    If Not (Rec.BOF And Rec.EOF) Then
        While Not Rec.EOF
            Print #gbFileNO, Rec!vchAccountHeadCode; "  "; PadR(Rec!vchAccountHead, 50); PadL(Format(Rec!Amount, "0.00"), 14)
            mTotalDr = mTotalDr + Format(Rec!Amount, "0.00")
            Rec.MoveNext
        Wend
        Print #gbFileNO, "                                                     ------------------------"
        Print #gbFileNO, "                                                            " & PadL(Format(mTotalDr, "0.00"), 15)
        Print #gbFileNO, "                                                     ========================"
    End If
    Close #gbFileNO
    ShellPad
    Shell "Print " & gbFileName
End Sub

Private Sub cmdOpeningBalance_Click()
    
    
        
    
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQL As String
    Dim RecBalance As New ADODB.Recordset
    Dim varInPut As Variant
    Dim mFlag As Boolean
    
    Dim mTotalDr As Currency
    Dim mTotalCr As Currency
    
    FileInitialize
    
    objDB.SetConnection mCnn
    
    
    
    'mSQL = "Select * From faAccountHeads Where NOT tinDebitOrCredit Is Null Order By tinDebitOrCredit, vchAccountHeadCode"
    
    mSQL = "Select * From faTransactionChild  Inner Join"
    mSQL = mSQL + "    faAccountHeads On faAccountHeads.intAccountHeadID = faTransactionChild.intAccountHeadID"
    mSQL = mSQL + " Where faTransactionChild.intTransactionID = 0 Order By faTransactionChild.tinDebitOrCreditFlag, faAccountHeads.vchAccountHeadCode"




    
    Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
    
    Print #gbFileNO, "   OPENING BALANCE   "
    Print #gbFileNO, "====================="
    While Not Rec.EOF
        If Rec!tinDebitOrCredit Then
            Print #gbFileNO, Rec!vchAccountHeadCode; " "; PadR(Rec!vchAccountHead, 40);
            Print #gbFileNO, Tab(60); PadL(Format(Rec!fltOpeningBalance, "0.00"), 12)
            mTotalDr = mTotalDr + Format(Rec!fltOpeningBalance, "0.00")
        Else
            Print #gbFileNO, Rec!vchAccountHeadCode; " "; PadR(Rec!vchAccountHead, 40);
            Print #gbFileNO, Tab(75); PadL(Format(Rec!fltOpeningBalance, "0.00"), 12)
            mTotalCr = mTotalCr + Format(Rec!fltOpeningBalance, "0.00")
        End If
        
        
        Rec.MoveNext
    Wend
    Rec.Close
    Print #gbFileNO, Tab(56); "----------------------------------------------------------"
    Print #gbFileNO, Tab(58); PadL(Format(mTotalDr, "0.00"), 14); Tab(75); PadL(Format(mTotalCr, "0.00"), 14)
    Print #gbFileNO, Tab(56); "----------------------------------------------------------"
    Close #gbFileNO
    ShellPad
    
    
End Sub

Private Sub cmdPropertyTax_Click()

Dim objAc As New clsAccounts
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQL As String
    Dim varInPut As Variant
    Dim mFlag As Boolean
    
    Dim mTotalDr As Currency
    Dim mTotalCr As Currency
    
    Dim mDailyDrTotal As Currency
    Dim mDailyCrTotal As Currency
    
    
    Dim mDt As Date
    Dim mPageNo As Integer
    Dim mLineNo As Integer
    Dim mID As Long
    
        FileInitialize
        
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO, Tab(10); objAc.AccountCode, objAc.AccountHead
        
        Print #gbFileNO, "--------------------------------------------------------------------------------------------------------------------------------------------"
        Print #gbFileNO, "Vou.ID       Receipt No.                     Amount   W\DoorNo  Name                                  Narration"
        Print #gbFileNO, "---------------------------------------------------------------------------------------------------------------------------------------------"
  
        
        mLineNo = 7
        objDB.SetConnection mCnn
        Set Rec = objDB.ExecuteSP("spRptLedgerForDOSPropertyTax", , , , mCnn, adCmdStoredProc)
        If Not (Rec.BOF And Rec.EOF) Then
            'mDt = Rec!dtTransactionDate
            While Not Rec.EOF
                If mID <> Rec!intVoucherID Then
                    Print #gbFileNO,
                    Print #gbFileNO, DdMmmYy(Rec!dtDate); " ";
                    Print #gbFileNO, Rec!intVoucherNo;
                    mSQL = ""
                    If Not IsNull(Rec!intWardNo) Then
                        mSQL = str(Rec!intWardNo) + "/"
                    End If
                    If Not IsNull(Rec!intDoorNo) Then
                        mSQL = mSQL + str(Rec!intDoorNo)
                    End If
                    If Not IsNull(Rec!vchDoorNo2) Then
                        mSQL = mSQL + "-" + Rec!vchDoorNo2
                    End If
                    
                    mSQL = ""
                    Print #gbFileNO, mSQL; "   ";
                    If Not IsNull(Rec!vchName) Then
                        mSQL = Rec!vchName
                    End If
                    Print #gbFileNO, PadR(mSQL, 50)
                    Print #gbFileNO,
                    
                    mID = Rec!intVoucherID
                End If
                Print #gbFileNO, Tab(5); PadR(Rec!vchAccountHead, 45); "  "; Rec!tnyPeriodID; "  "; Rec!intYearID;
                Print #gbFileNO, PadL(Format(Rec!fltAmt, "0.00"), 13)
                
                mLine = mLine + 1
                Rec.MoveNext
            Wend
        End If
        Close #gbFileNO
        ShellPad
    
    



End Sub

Private Sub cmdReceiptsAndPayments_Click()

    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQL As String
    Dim RecBalance As New ADODB.Recordset
    Dim varInPut As Variant
    Dim mFlag As Boolean

    Dim mTotalDr As Currency
    Dim mTotalCr As Currency

'    FileInitialize
'    objDB.SetConnection mCnn
'
'    mSQL = "Select * From faAccountHeads where intMajorAccountHeadID = 40 AND tinHiddenFlag = 0 Order By vchAccountHeadCode"
'    Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
'
'    Print #gbFileNO, "   OPENING BALANCE"
'    Print #gbFileNO, "==================="
'    While Not Rec.EOF
'        If Rec!fltOpeningBalance <> 0 Then
'            Print #gbFileNO, Rec!vchAccountHeadCode; "  "; PadL(Rec!vchAccountHead, 30);
'            Print #gbFileNO, Tab(90); PadL(Format(Rec!fltOpeningBalance, "00"), 12)
'            mTotalDr = mTotalDr + Format(Rec!fltOpeningBalance, "00")
'
'        End If
'        Rec.MoveNext
'    Wend
'    Rec.Close
'
'
'    mSQL = " Select Distinct faAccountHeads.intAccountHeadID,faAccountHeads.intMajorAccountHeadID,"
'    mSQL = mSQL + " faAccountHeads.intMinorAccountHeadID , faAccountHeads.vchAccountHeadCode ,"
'    mSQL = mSQL + " faAccountHeads.vchAccountHead, vchMinorAccountHeadCode, vchMinorAccountHead,"
'    mSQL = mSQL + " vchMajorAccountHeadCode , vchMajorAccountHead, intOperating "
'    mSQL = mSQL + " From faTransactionChild Inner Join"
'    mSQL = mSQL + " faTransactions ON faTransactions.intTransactionID = faTransactionChild.intTransactionID Inner Join"
'    mSQL = mSQL + " faAccountHeads ON faAccountHeads.intAccountHeadID = faTransactionChild.intAccountHeadID Inner Join"
'    mSQL = mSQL + " faMinorAccountHeads On faMinorAccountHeads.intMinorAccountHeadID = faAccountHeads.intMinorAccountHeadID Inner Join"
'    mSQL = mSQL + " faMajorAccountHeads On faMajorAccountHeads.intMajorAccountHeadID = faAccountHeads.intMajorAccountHeadID"
'    mSQL = mSQL + " Where faTransactions.intGroupID In (10,20) AND faAccountHeads.intMajorAccountHeadID <>40"
'    mSQL = mSQL + " Order By intOperating, faAccountHeads.intMajorAccountHeadID, faAccountHeads.intMinorAccountHeadID,"
'    mSQL = mSQL + " faAccountHeads.vchAccountHeadCode "
'
'    Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
'
'    Print #gbFileNO,
'    Print #gbFileNO,
'    Print #gbFileNO,
'    Print #gbFileNO, "   RECEIPTS"
'    Print #gbFileNO, "======================================================"
'    Print #gbFileNO,
'    Print #gbFileNO, "  Operating Payments"
'    Print #gbFileNO, "------------------------------------------------------"
'
'    While Not Rec.EOF
'        varInput = Array(Rec!intAccountHeadID)
'        Set RecBalance = objDB.ExecuteSP("spGetClosingBalanceWithOutOpening", varInput, , , mCnn, adCmdStoredProc)
'        If Not (RecBalance.EOF And RecBalance.BOF) Then
'            If RecBalance!Cr <> 0 Then
'                If Not mFlag Then
'                    If Rec!intOperating = 1 Then
'                        Print #gbFileNO, "------------------------------------------------------"
'                        Print #gbFileNO, "  Non-Operating Receipts"
'                        Print #gbFileNO, "------------------------------------------------------"
'                        mFlag = True
'                    End If
'                End If
'                Print #gbFileNO, Rec!vchAccountHeadCode; "  "; Rec!vchAccountHead;
'                Print #gbFileNO, Tab(90); PadL(Format(RecBalance!Cr, "00"), 12)
'                mTotalDr = mTotalDr + Format(RecBalance!Cr, "00")
'            End If
'        End If
'        RecBalance.Close
'        Rec.MoveNext
'    Wend
'    Rec.MoveFirst
'    mFlag = False
'
'    Print #gbFileNO,
'    Print #gbFileNO,
'    Print #gbFileNO,
'    Print #gbFileNO, "   PAYMENTS"
'    Print #gbFileNO, "======================================================"
'    Print #gbFileNO,
'    Print #gbFileNO, "  Operating Payments"
'    Print #gbFileNO, "------------------------------------------------------"
'
'    While Not Rec.EOF
'        varInput = Array(Rec!intAccountHeadID)
'        Set RecBalance = objDB.ExecuteSP("spGetClosingBalanceWithOutOpening", varInput, , , mCnn, adCmdStoredProc)
'        If Not (RecBalance.EOF And RecBalance.BOF) Then
'            If RecBalance!Dr <> 0 Then
'                If Not mFlag Then
'                    If Rec!intOperating = 1 Then
'                        Print #gbFileNO, "------------------------------------------------------"
'                        Print #gbFileNO, "  Non-Operating Payments"
'                        Print #gbFileNO, "------------------------------------------------------"
'                        mFlag = True
'                    End If
'                End If
'                Print #gbFileNO, Rec!vchAccountHeadCode; "  "; Rec!vchAccountHead;
'                Print #gbFileNO, Tab(118); PadL(Format(RecBalance!Dr, "00"), 12)
'                mTotalCr = mTotalCr + Format(RecBalance!Dr, "00")
'            End If
'        End If
'        RecBalance.Close
'        Rec.MoveNext
'    Wend
'    Rec.Close
'
'    mSQL = "Select * From faAccountHeads where intMajorAccountHeadID = 40 AND tinHiddenFlag = 0 Order By vchAccountHeadCode"
'    Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
'
'    Print #gbFileNO,
'    Print #gbFileNO,
'    Print #gbFileNO,
'    Print #gbFileNO, "   CLOSING BALANCE "
'    Print #gbFileNO, "==================="
'    While Not Rec.EOF
'        varInput = Array(Rec!intAccountHeadID)
'        Set RecBalance = objDB.ExecuteSP("spGetClosingBalance", varInput, , , mCnn, adCmdStoredProc)
'        If Not (RecBalance.EOF And RecBalance.BOF) Then
'            If RecBalance!Balance > 0 Then
'            Print #gbFileNO, Rec!vchAccountHeadCode; "  "; PadL(Rec!vchAccountHead, 30);
'            Print #gbFileNO, Tab(118); PadL(Format(RecBalance!Balance, "00"), 12)
'            'mTotalDr = mTotalDr + Format(RecBalance!Balance, "00")
'            ElseIf RecBalance!Balance < 0 Then
'            Print #gbFileNO, Rec!vchAccountHeadCode; "  "; PadL(Rec!vchAccountHead, 30);
'            Print #gbFileNO, Tab(90); PadL(Format(RecBalance!Balance, "00"), 12)
'            'mTotalCr = mTotalCr + Format(RecBalance!Balance, "00")
'            End If
'            mTotalCr = mTotalCr + Format(RecBalance!Balance, "00")
'        End If
'        RecBalance.Close
'
'        Rec.MoveNext
'    Wend
'    Rec.Close
'
'
'    Print #gbFileNO, Tab(72); "============================================================"
'    Print #gbFileNO, Tab(90); PadL(Format(mTotalDr, "00"), 12); Tab(118); PadL(Format(mTotalCr, "00"), 12);
'    Print #gbFileNO, Tab(72); "============================================================"
'    Close #gbFileNO
'    ShellPad
'
'
'



    FileInitialize

    objDB.SetConnection mCnn

    mSQL = "Select * From faAccountHeads where intMajorAccountHeadID = 40 AND tinHiddenFlag = 0 Order By vchAccountHeadCode"
    Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic

    Print #gbFileNO, "   OPENING BALANCE"
    Print #gbFileNO, "==================="
    While Not Rec.EOF
        If Rec!fltOpeningBalance <> 0 Then
            Print #gbFileNO, Rec!vchAccountHeadCode; "  "; PadL(Rec!vchAccountHead, 30);
            Print #gbFileNO, Tab(50); PadL(Format(Rec!fltOpeningBalance, "00"), 12)
            mTotalDr = mTotalDr + Format(Rec!fltOpeningBalance, "00")
        End If
        Rec.MoveNext
    Wend
    Rec.Close


    mSQL = " Select Distinct faAccountHeads.intAccountHeadID,faAccountHeads.intMajorAccountHeadID,"
    mSQL = mSQL + " faAccountHeads.intMinorAccountHeadID , faAccountHeads.vchAccountHeadCode ,"
    mSQL = mSQL + " faAccountHeads.vchAccountHead, vchMinorAccountHeadCode, vchMinorAccountHead,"
    mSQL = mSQL + " vchMajorAccountHeadCode , vchMajorAccountHead, intOperating "
    mSQL = mSQL + " From faTransactionChild Inner Join"
    mSQL = mSQL + " faTransactions ON faTransactions.intTransactionID = faTransactionChild.intTransactionID Inner Join"
    mSQL = mSQL + " faAccountHeads ON faAccountHeads.intAccountHeadID = faTransactionChild.intAccountHeadID Inner Join"
    mSQL = mSQL + " faMinorAccountHeads On faMinorAccountHeads.intMinorAccountHeadID = faAccountHeads.intMinorAccountHeadID Inner Join"
    mSQL = mSQL + " faMajorAccountHeads On faMajorAccountHeads.intMajorAccountHeadID = faAccountHeads.intMajorAccountHeadID"
    mSQL = mSQL + " Where faTransactions.intGroupID In (10,20) AND faAccountHeads.intMajorAccountHeadID <>40"
    mSQL = mSQL + " Order By intOperating, faAccountHeads.intMajorAccountHeadID, faAccountHeads.intMinorAccountHeadID,"
    mSQL = mSQL + " faAccountHeads.vchAccountHeadCode "

    Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic

    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO, "   RECEIPTS"
    Print #gbFileNO, "======================================================"

    Print #gbFileNO,
    Print #gbFileNO, "  Operating Payments"
    Print #gbFileNO, "------------------------------------------------------"

    While Not Rec.EOF
        varInPut = Array(Rec!intAccountHeadID)
        Set RecBalance = objDB.ExecuteSP("spGetClosingBalanceWithOutOpening", varInPut, , , mCnn, adCmdStoredProc)
        If Not (RecBalance.EOF And RecBalance.BOF) Then
            If RecBalance!CR <> 0 Then
                If Not mFlag Then
                    If Rec!intOperating = 1 Then
                        Print #gbFileNO, "------------------------------------------------------"
                        Print #gbFileNO, "  Non-Operating Receipts"
                        Print #gbFileNO, "------------------------------------------------------"
                        mFlag = True
                    End If
                End If
                Print #gbFileNO, Rec!vchAccountHeadCode; "  "; PadL(Rec!vchAccountHead, 30);
                Print #gbFileNO, Tab(50); PadL(Format(RecBalance!CR, "00"), 12); Tab(88); Rec!vchAccountHead
                mTotalDr = mTotalDr + Format(RecBalance!CR, "00")
            End If
        End If
        RecBalance.Close
        Rec.MoveNext
    Wend
    Rec.MoveFirst
    mFlag = False

    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO, "   PAYMENTS"
    Print #gbFileNO, "======================================================"

    Print #gbFileNO,
    Print #gbFileNO, "  Operating Payments"
    Print #gbFileNO, "------------------------------------------------------"

    While Not Rec.EOF
        varInPut = Array(Rec!intAccountHeadID)
        Set RecBalance = objDB.ExecuteSP("spGetClosingBalanceWithOutOpening", varInPut, , , mCnn, adCmdStoredProc)
        If Not (RecBalance.EOF And RecBalance.BOF) Then
            If RecBalance!Dr <> 0 Then
                If Not mFlag Then
                    If Rec!intOperating = 1 Then
                        Print #gbFileNO, "------------------------------------------------------"
                        Print #gbFileNO, "  Non-Operating Payments"
                        Print #gbFileNO, "------------------------------------------------------"
                        mFlag = True
                    End If
                End If
                Print #gbFileNO, Rec!vchAccountHeadCode; "  "; PadL(Rec!vchAccountHead, 30);
                Print #gbFileNO, Tab(74); PadL(Format(RecBalance!Dr, "00"), 12); Tab(88); Rec!vchAccountHead
                mTotalCr = mTotalCr + Format(RecBalance!Dr, "00")
            End If
        End If
        RecBalance.Close
        Rec.MoveNext
    Wend
    Rec.Close

    mSQL = "Select * From faAccountHeads where intMajorAccountHeadID = 40 AND tinHiddenFlag = 0 Order By vchAccountHeadCode"
    Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic

    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO, "   CLOSING BALANCE "
    Print #gbFileNO, "==================="
    While Not Rec.EOF
        varInPut = Array(Rec!intAccountHeadID)
        Set RecBalance = objDB.ExecuteSP("spGetClosingBalance", varInPut, , , mCnn, adCmdStoredProc)
        If Not (RecBalance.EOF And RecBalance.BOF) Then
            If RecBalance!Balance > 0 Then
            Print #gbFileNO, Rec!vchAccountHeadCode; "  "; PadL(Rec!vchAccountHead, 30);
            Print #gbFileNO, Tab(74); PadL(Format(RecBalance!Balance, "00"), 12)
            'mTotalDr = mTotalDr + Format(RecBalance!Balance, "00")
            ElseIf RecBalance!Balance < 0 Then
            Print #gbFileNO, Rec!vchAccountHeadCode; "  "; PadL(Rec!vchAccountHead, 30);
            Print #gbFileNO, Tab(50); PadL(Format(RecBalance!Balance, "00"), 12)
            'mTotalCr = mTotalCr + Format(RecBalance!Balance, "00")
            End If
            mTotalCr = mTotalCr + Format(RecBalance!Balance, "00")
        End If
        RecBalance.Close

        Rec.MoveNext
    Wend
    Rec.Close


    Print #gbFileNO, Tab(50); "============================================================"
    Print #gbFileNO, Tab(50); PadL(Format(mTotalDr, "00"), 12); Tab(74); PadL(Format(mTotalCr, "00"), 12);
    Print #gbFileNO, Tab(50); "============================================================"
    Close #gbFileNO
    ShellPad



    
End Sub

Private Sub cmdRegister_Click()
    Dim objAc As New clsAccounts
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQL As String
    Dim varInPut As Variant
    Dim mFlag As Boolean
    
    Dim mTotalDr As Currency
    Dim mTotalCr As Currency
    
    Dim mDailyDrTotal As Currency
    Dim mDailyCrTotal As Currency
    
    
    Dim mDt As Date
    Dim mPageNo As Integer
    Dim mLineNo As Integer
    
    objAc.SetAccountCode (Trim(txtHeadCode))
    If objAc.AccountHeadID > 0 Then
        varInPut = Array(objAc.AccountHeadID)
        FileInitialize
        
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO, Tab(10); objAc.AccountCode, objAc.AccountHead
        
        Print #gbFileNO, "---------------------------------------------------------------------------------------------------------------------------------------------"
        Print #gbFileNO, "Vou.ID       Receipt No.                     Amount   W\DoorNo  Name                                  Narration"
        Print #gbFileNO, "---------------------------------------------------------------------------------------------------------------------------------------------"
  
        
        mLineNo = 7
        objDB.SetConnection mCnn
        Set Rec = objDB.ExecuteSP("spRptLedgerForDOS", varInPut, , , mCnn, adCmdStoredProc)
        If Not (Rec.BOF And Rec.EOF) Then
            mDt = Rec!dtTransactionDate
            While Not Rec.EOF
                 If Rec!vchGroup <> "R" Then GoTo lblSubTotal:
                
                'Print #gbFileNO, PadL(str(Rec!intVoucherID), 7); " ";
                Print #gbFileNO, DdMmmYy(Rec!dtTransactionDate); " ";
                Print #gbFileNO, Rec!intVoucherNo; "  ";
                'Print #gbFileNO, PadR(Rec!vchAccountHead, 25); "|";
                'If Rec!tinDebitOrCreditFlag = 1 Then
                If Rec!tinDrOrCr = 1 Then
                    mTotalCr = mTotalCr + Format(Rec!fltAmount, "0.00")
                    mDailyCrTotal = mDailyCrTotal + Format(Rec!fltAmount, "0.00")
                    Print #gbFileNO, PadR(" ", 13); PadL(Format(Rec!fltAmount, "0.00"), 12);
                Else
                    mTotalDr = mTotalDr + Format(Rec!fltAmount, "0.00")
                    mDailyDrTotal = mDailyDrTotal + Format(Rec!fltAmount, "0.00")
                    Print #gbFileNO, PadL(Format(Rec!fltAmount, "0.00"), 12); PadR(" ", 13);
                End If
                
                'Print #gbFileNO, Tab(84); "|"; PadL(Format(mTotalDr - mTotalCr, "0.00"), 11); "   ";
                Print #gbFileNO, "  "; Rec!vchGroup; " "; 'IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo);
'                If IsDate(Rec!dtInstrumentDate) Then
'                    Print #gbFileNO, PadR("\" + DdMmmYy(Rec!dtInstrumentDate), 20);
'                Else
'                    Print #gbFileNO, Space(20);
'                End If
                mSQL = IIf(IsNull(Rec!intWardNo), " ", Rec!intWardNo)
                mSQL = mSQL & "\" & IIf(IsNull(Rec!intDoorNo), "", Rec!intDoorNo)
                Print #gbFileNO, PadR(mSQL, 8); "  ";
                Print #gbFileNO, PadR(IIf(IsNull(Rec!vchName), "", Rec!vchName), 30); " ";
                Print #gbFileNO, PadR(IIf(IsNull(Rec!vchDescription), "", Rec!vchDescription), 45)
lblSubTotal:
                mLine = mLine + 1
                Rec.MoveNext

            Wend
        End If
        Close #gbFileNO
        ShellPad
    End If
    

End Sub

Private Sub cmdRPConsolidation_Click()

        
    
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQL As String
    Dim RecBalance As New ADODB.Recordset
    Dim varInPut As Variant
    Dim mFlag As Boolean
    
    Dim mTotalDr As Currency
    Dim mTotalCr As Currency
    
    FileInitialize
    
    objDB.SetConnection mCnn
    
    'mSQL = "Select * From faMinorAccountHeads where intMajorAccountHeadID = 40 Order By vchMinorAccountHeadCode"
    
    mSQL = "Select faAccountHeads.intGroupID, Sum(fltOpeningBalance) fltOpeningBalance From faAccountHeads Inner Join"
    mSQL = mSQL + " faMinorAccountHeads On faMinorAccountHeads.intMinorAccountHeadID = faAccountHeads.intMinorAccountHeadID"
    mSQL = mSQL + " Where faAccountHeads.intMajorAccountHeadID = 40"
    mSQL = mSQL + " Group By faAccountHeads.intGroupID"

    
    Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
    
    Print #gbFileNO, "   OPENING BALANCE   "
    Print #gbFileNO, "====================="
    While Not Rec.EOF
        If Rec!fltOpeningBalance <> 0 Then
            If Rec!intGroupID = 1 Then
                Print #gbFileNO, "  Cash";
            Else
                Print #gbFileNO, "  Bank";
            End If
            Print #gbFileNO, Tab(50); PadL(Format(Rec!fltOpeningBalance, "00"), 12)
            mTotalDr = mTotalDr + Format(Rec!fltOpeningBalance, "00")
        End If
        Rec.MoveNext
    Wend
    Rec.Close
    
    mSQL = "        Select Distinct faAccountHeads.intMajorAccountHeadID, vchMajorAccountHeadCode , "
    mSQL = mSQL + " vchMajorAccountHead,intOperating "
    mSQL = mSQL + " From faTransactionChild Inner Join"
    mSQL = mSQL + " faTransactions ON faTransactions.intTransactionID = faTransactionChild.intTransactionID Inner Join"
    mSQL = mSQL + " faAccountHeads ON faAccountHeads.intAccountHeadID = faTransactionChild.intAccountHeadID Inner Join"
    mSQL = mSQL + " faMajorAccountHeads On faMajorAccountHeads.intMajorAccountHeadID = faAccountHeads.intMajorAccountHeadID"
    mSQL = mSQL + " Where faTransactions.intGroupID In (10,20) AND faAccountHeads.intMajorAccountHeadID <> 40"
    mSQL = mSQL + " Order By intOperating, faAccountHeads.intMajorAccountHeadID"
    
    Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO, "   RECEIPTS"
    Print #gbFileNO, '"======================================================"
    Print #gbFileNO, "  Operating Receipts"
    Print #gbFileNO, '"------------------------------------------------------"
    
    While Not Rec.EOF
        varInPut = Array(Rec!intMajorAccountHeadID)
        Set RecBalance = objDB.ExecuteSP("spGetClosingBalanceMajorHeadWithOutOpening", varInPut, , , mCnn, adCmdStoredProc)
        If Not (RecBalance.EOF And RecBalance.BOF) Then
            If RecBalance!CR <> 0 Then
                If Not mFlag Then
                    If Rec!intOperating = 1 Then
                        Print #gbFileNO, '"------------------------------------------------------"
                        Print #gbFileNO, "  Non-Operating Receipts"
                        Print #gbFileNO, '"------------------------------------------------------"
                        mFlag = True
                    End If
                End If
                    
                'Print #gbFileNO, Rec!vchMajorAccountHeadCode; "  "; PadL(Rec!vchMajorAccountHead, 30);
                Print #gbFileNO, Rec!vchMajorAccountHeadCode; "  "; Rec!vchMajorAccountHead;
                Print #gbFileNO, Tab(50); PadL(Format(RecBalance!CR, "00"), 12)
                mTotalDr = mTotalDr + Format(RecBalance!CR, "00")
            End If
        End If
        RecBalance.Close
        Rec.MoveNext
    Wend
    Rec.MoveFirst
    mFlag = False
    
    
    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO, "   PAYMENTS"
    Print #gbFileNO, '"======================================================"
    Print #gbFileNO, "     Operating Payments"
    Print #gbFileNO, '"------------------------------------------------------"
                    
    While Not Rec.EOF
        varInPut = Array(Rec!intMajorAccountHeadID)
        Set RecBalance = objDB.ExecuteSP("spGetClosingBalanceMajorHeadWithOutOpening", varInPut, , , mCnn, adCmdStoredProc)
        If Not (RecBalance.EOF And RecBalance.BOF) Then
            If RecBalance!Dr <> 0 Then
                If Not mFlag Then
                    If Rec!intOperating = 1 Then
                    
                        Print #gbFileNO, '"------------------------------------------------------"
                        Print #gbFileNO, "     Non-Operating Payments"
                        Print #gbFileNO, ' "------------------------------------------------------"
                        mFlag = True
                    End If
                End If
                    
                'Print #gbFileNO, Rec!vchMajorAccountHeadCode; "  "; PadL(Rec!vchMajorAccountHead, 30);
                Print #gbFileNO, Rec!vchMajorAccountHeadCode; "  "; Rec!vchMajorAccountHead;
                Print #gbFileNO, Tab(74); PadL(Format(RecBalance!Dr, "00"), 12)
                mTotalCr = mTotalCr + Format(RecBalance!Dr, "00")
            End If
        End If
        RecBalance.Close
        Rec.MoveNext
    Wend
    Rec.Close
    
    Dim mClosingCash As Double
    Dim mClosingBank As Double
    
    mSQL = "Select * From faAccountHeads where intMajorAccountHeadID = 40 AND tinHiddenFlag = 0 Order By intGroupID,vchAccountHeadCode"
    Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
    
    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO,
    Print #gbFileNO, "   CLOSING BALANCE "
    Print #gbFileNO, '"==================="
    While Not Rec.EOF
        varInPut = Array(Rec!intAccountHeadID)
        Set RecBalance = objDB.ExecuteSP("spGetClosingBalance", varInPut, , , mCnn, adCmdStoredProc)
        If Not (RecBalance.EOF And RecBalance.BOF) Then
            If Rec!intGroupID = 1 Then
                mClosingCash = mClosingCash + RecBalance!NetBalance
            ElseIf Rec!intGroupID = 2 Then
                mClosingBank = mClosingBank + RecBalance!NetBalance
            End If
            mTotalCr = mTotalCr + Format(RecBalance!NetBalance, "00")
        End If
        RecBalance.Close
        Rec.MoveNext
    Wend
    Rec.Close
    
    Print #gbFileNO, " Cash";
    Print #gbFileNO, Tab(74); PadL(Format(mClosingCash, "00"), 12)
    
    Print #gbFileNO, " Bank";
    Print #gbFileNO, Tab(74); PadL(Format(mClosingBank, "00"), 12)

    
    Print #gbFileNO, Tab(50); "============================================================"
    Print #gbFileNO, Tab(50); PadL(Format(mTotalDr, "00"), 12); Tab(74); PadL(Format(mTotalCr, "00"), 12);
    Print #gbFileNO, Tab(50); "============================================================"
    Close #gbFileNO
    ShellPad



End Sub

Private Sub cmdTrialBalance_Click()
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQL As String
    Dim RecBalance As New ADODB.Recordset
    Dim varInPut As Variant
    Dim mFlag As Boolean
    Dim mOpeningDr As Double
    Dim mOpeningCr As Double
    Dim mOBDr As Double
    Dim mOBCr As Double
    
    Dim mCurrentDr As Double
    Dim mCurrentCr As Double
    
    Dim mTotalDr As Currency
    Dim mTotalCr As Currency
    
    FileInitialize
    
    objDB.SetConnection mCnn
    
    
    mSQL = "Select * From faAccountHeads Order By vchAccountHeadCode"
    Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
    
    Print #gbFileNO, "             TRIAL BALANCE"
    Print #gbFileNO, "========================================================================================================================================="
    Print #gbFileNO, "  Code                           Head     Opening(Dr)    Opening(Cr)    Current(Dr)    Current(Cr)        Closing(Dr)     Closing(Cr)"
    Print #gbFileNO, "==========================================================================================================================================="
    While Not Rec.EOF
        mOBDr = 0
        mOBCr = 0
        
        varInPut = Array(Rec!intAccountHeadID)
        Set RecBalance = objDB.ExecuteSP("spGetClosingBalance", varInPut, , , mCnn, adCmdStoredProc)
        If Not (RecBalance.EOF And RecBalance.BOF) Then
            If RecBalance!Balance > 0 Then
                Print #gbFileNO, Rec!vchAccountHeadCode; "  "; PadL(Rec!vchAccountHead, 30);
                If Not IsNull(RecBalance!obflag) Then
                If Rec!tinDebitOrCredit Then
                    Print #gbFileNO, Tab(45); PadL(Format(Rec!fltOpeningBalance, "0.00"), 13);
                    mOpeningDr = mOpeningDr + Format(Rec!fltOpeningBalance, "0.00")
                    mOBDr = Format(Rec!fltOpeningBalance, "0.00")
                Else
                    Print #gbFileNO, Tab(60); PadL(Format(Rec!fltOpeningBalance, "0.00"), 13);
                    mOpeningCr = mOpeningCr + Format(Rec!fltOpeningBalance, "0.00")
                    mOBCr = Format(Rec!fltOpeningBalance, "0.00")
                End If
                End If
                Print #gbFileNO, Tab(74); PadL(Format(RecBalance!Dr - mOBDr, "0.00"), 13);
                Print #gbFileNO, Tab(89); PadL(Format(RecBalance!CR - mOBCr, "0.00"), 13);
                
                mCurrentDr = mCurrentDr + Format(RecBalance!Dr - mOBDr, "0.00")
                mCurrentCr = mCurrentCr + Format(RecBalance!CR - mOBCr, "0.00")
                
                Print #gbFileNO, Tab(110); PadL(Format(RecBalance!Balance, "0.00"), 13);
                mTotalDr = mTotalDr + Format(RecBalance!Balance, "0.00")
                Print #gbFileNO, Tab(140); Rec!vchAccountHead
            Else
                Print #gbFileNO, Rec!vchAccountHeadCode; "  "; PadL(Rec!vchAccountHead, 30);
                If Not IsNull(RecBalance!obflag) Then
                If Rec!tinDebitOrCredit Then
                    Print #gbFileNO, Tab(45); PadL(Format(Rec!fltOpeningBalance, "0.00"), 13);
                    mOpeningDr = mOpeningDr + Format(Rec!fltOpeningBalance, "0.00")
                    mOBDr = Format(Rec!fltOpeningBalance, "0.00")
                Else
                    Print #gbFileNO, Tab(60); PadL(Format(Rec!fltOpeningBalance, "0.00"), 13);
                    mOpeningCr = mOpeningCr + Format(Rec!fltOpeningBalance, "0.00")
                    mOBCr = Format(Rec!fltOpeningBalance, "0.00")
                End If
                End If
                Print #gbFileNO, Tab(74); PadL(Format(RecBalance!Dr - mOBDr, "0.00"), 13);
                Print #gbFileNO, Tab(89); PadL(Format(RecBalance!CR - mOBCr, "0.00"), 13);
                mCurrentDr = mCurrentDr + Format(RecBalance!Dr - mOBDr, "0.00")
                mCurrentCr = mCurrentCr + Format(RecBalance!CR - mOBCr, "0.00")
                
                Print #gbFileNO, Tab(125); PadL(Format(Abs(RecBalance!Balance), "0.00"), 13);
                mTotalCr = mTotalCr + Format(Abs(RecBalance!Balance), "0.00")
                
                Print #gbFileNO, Tab(140); Rec!vchAccountHead
                
            End If
        End If
        RecBalance.Close
        
        Rec.MoveNext
    Wend
    Rec.Close
    Print #gbFileNO, "============================================================================================================================================="
    Print #gbFileNO, Tab(45); PadL(Format(mOpeningDr, "0.00"), 13); Tab(60); PadL(Format(mOpeningCr, "0.00"), 13);
    Print #gbFileNO, Tab(74); PadL(Format(mCurrentDr, "0.00"), 13); Tab(89); PadL(Format(mCurrentCr, "0.00"), 13);
    Print #gbFileNO, Tab(112); PadL(Format(mTotalDr, "0.00"), 13); Tab(127); PadL(Format(mTotalCr, "0.00"), 13)
    Print #gbFileNO, "============================================================================================================================================="
    Close #gbFileNO
    ShellPad
End Sub

Private Sub Command1_Click()

    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQL As String
    Dim RecBalance As New ADODB.Recordset
    Dim varInPut As Variant
    Dim mCurrentBalance As Variant
    Dim mTotalInc As Variant
    Dim mTotalExp As Variant
    Dim mTotalPrior As Variant
    objDB.SetConnection mCnn
    
    objDB.ExecuteSP "Delete From faIEStatment", , , , mCnn, adCmdText
    mSQL = "Select * From faAccountHeads Where tinType IN (1, 2)  Order By tinType, faAccountHeads.vchAccountHeadCode"
    Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
    While Not Rec.EOF
        
        '@intLocalBodyID    [int],
        '@intAccountHeadID  [int],
        '@vchAccountHeadCode    [varchar](15),
        '@vchAccountHead    [varchar](250),
        '@fltDrAmount       [numeric](18,2),
        '@fltCrAmount       [numeric](18,2)
        
        varInPut = Array(Rec!intAccountHeadID, Null, "30-Sep-2008")
        Set RecBalance = objDB.ExecuteSP("spGetClosingBalance", varInPut, , , mCnn, adCmdStoredProc)
        mCurrentBalance = 0
        If Not (RecBalance.EOF And RecBalance.BOF) Then
            mCurrentBalance = RecBalance!Balance
        End If
        RecBalance.Close
        
        If Left(Rec!vchAccountHeadCode, 3) = "280" Then
            intReportGroupID = 4
            mTotalPrior = mTotalPrior + mCurrentBalance
        Else
            intReportGroupID = Rec!tinType
            If Rec!tinType = 1 Then
                mCurrentBalance = mCurrentBalance * -1
                mTotalInc = mTotalInc + mCurrentBalance
            ElseIf Rec!tinType = 2 Then
                mTotalExp = mTotalExp + mCurrentBalance
            End If
        End If
        
        varInPut = Array(gbLocalBodyID, Rec!intAccountHeadID, Rec!vchAccountHeadCode, Rec!vchAccountHead, mCurrentBalance, Null, Rec!tinType, intReportGroupID)
        objDB.ExecuteSP "spSaveIE", varInPut, , , mCnn, adCmdStoredProc
        
        Rec.MoveNext
    Wend
    
    varInPut = Array(gbLocalBodyID, Null, Null, "Gross", mTotalInc - mTotalExp, Null, Null, 3)
    objDB.ExecuteSP "spSaveIE", varInPut, , , mCnn, adCmdStoredProc
    
    varInPut = Array(gbLocalBodyID, Null, Null, "Net", mTotalInc - mTotalExp - mTotalPrior, Null, Null, 5)
    objDB.ExecuteSP "spSaveIE", varInPut, , , mCnn, adCmdStoredProc
    
    '============================================================================================='
    '        If Left(Rec!vchAccountHeadCode, 3) = "280" Then
    '            GoTo lblPriorPeriod:
    '        End If
    '============================================================================================='
    
    
    
    
End Sub

Private Sub cmdTrialBalanceConsolidation_Click()

    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQL As String
    Dim RecBalance As New ADODB.Recordset
    Dim varInPut As Variant
    Dim mFlag As Boolean
    Dim mOpeningDr As Double
    Dim mOpeningCr As Double
    Dim mOBDr As Double
    Dim mOBCr As Double
    
    Dim mCurrentDr As Double
    Dim mCurrentCr As Double
    
    Dim mTotalDr As Currency
    Dim mTotalCr As Currency
    
    FileInitialize
    
    objDB.SetConnection mCnn
    
    
    mSQL = "Select * From faMajorAccountHeads Order By faMajorAccountHeads.vchMajorAccountHeadCode"
    Rec.Open mSQL, mCnn, adOpenKeyset, adLockOptimistic
    
    Print #gbFileNO, "          TRIAL BALANCE - (Major Head Wise)   "
    Print #gbFileNO, "====================================================="
    
    Set Rec = objDB.ExecuteSP("spGetClosingBalanceMajorHeadWise", , , , mCnn, adCmdStoredProc)
    While Not Rec.EOF
        Print #gbFileNO, Rec!vchMajorAccountHeadCode; "  "; PadL(Format(Rec!Dr, "0.00"), 14); PadL(Format(Rec!CR, "0.00"), 14); PadL(Format(Rec!Balance, "0.00"), 14)
        mTotalDr = mTotalDr + Format(Rec!Dr, "0.00")
        mTotalCr = mTotalCr + Format(Rec!CR, "0.00")
        Rec.MoveNext
    Wend
    Print #gbFileNO,
    Print #gbFileNO, Tab(11); PadL(Format(mTotalDr, "0.00"), 14); PadL(Format(mTotalCr, "0.00"), 14)
    Close #gbFileNO
    ShellPad

End Sub

Private Sub Command2_Click()
    Dim objAc As New clsAccounts
    Dim objDB As New clsDB
    Dim mCnn As New ADODB.Connection
    Dim Rec As New ADODB.Recordset
    Dim mSQL As String
    Dim varInPut As Variant
    Dim mFlag As Boolean
    
    Dim mTotalDr As Currency
    Dim mTotalCr As Currency
    
    Dim mDailyDrTotal As Currency
    Dim mDailyCrTotal As Currency
    
    
    Dim mDt As Date
    Dim mPageNo As Integer
    Dim mLineNo As Integer
    Dim mLoop As Integer
    
    objAc.SetAccountCode (Trim(txtHeadCode))
    If objAc.AccountHeadID > 0 Then
        varInPut = Array(objAc.AccountHeadID)
        FileInitialize
        
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO,
        Print #gbFileNO, Tab(10); objAc.AccountCode, objAc.AccountHead
        Print #gbFileNO, "---------------------------------------------------------------------------------------------------------------------------------"
        Print #gbFileNO, "Vou.ID Trn. Date   Head Code  Account Head                      Debit       Credit      Balance Type Tr.ID   Cheque No.\Date"
        Print #gbFileNO, "---------------------------------------------------------------------------------------------------------------------------------"
        mLineNo = 7
        objDB.SetConnection mCnn
        Set Rec = objDB.ExecuteSP("spRptLedgerForDOS", varInPut, , , mCnn, adCmdStoredProc)
        If Not (Rec.BOF And Rec.EOF) Then
            mDt = DateSerial(Year(Rec!dtTransactionDate), Month(Rec!dtTransactionDate), 1)
            While Not Rec.EOF
                
                If mLineNo >= 67 Then
                    mPageNo = mPageNo + 1
                    mLineNo = 0
                    Print #gbFileNO,
                    Print #gbFileNO, Tab(100); "Page No:"; mPageNo
                    Print #gbFileNO,
                    Print #gbFileNO,
                    Print #gbFileNO, ' Page Ending on 72 Line
                    
                    
                    Print #gbFileNO,
                    Print #gbFileNO,
                    Print #gbFileNO,
                    Print #gbFileNO, Tab(10); objAc.AccountCode, objAc.AccountHead
                    Print #gbFileNO, "---------------------------------------------------------------------------------------------------------------------------------"
                    Print #gbFileNO, "Vou.ID Trn. Date   Head Code  Account Head                      Debit       Credit      Balance Type Tr.ID   Cheque No.\Date"
                    Print #gbFileNO, "---------------------------------------------------------------------------------------------------------------------------------"
                    mLineNo = 7
                End If
                
                Print #gbFileNO, PadL(str(Rec!intVoucherID), 7); " ";
                Print #gbFileNO, DdMmmYy(Rec!dtTransactionDate); " ";
                Print #gbFileNO, Rec!vchAccountHeadCode; "  ";
                Print #gbFileNO, PadR(Rec!vchAccountHead, 25); "|";
                'If Rec!tinDebitOrCreditFlag = 1 Then
                If Rec!tinDrOrCr = 1 Then
                    mTotalCr = mTotalCr + Format(Rec!fltAmount, "0.00")
                    mDailyCrTotal = mDailyCrTotal + Format(Rec!fltAmount, "0.00")
                    Print #gbFileNO, Tab(71); PadL(Format(Rec!fltAmount, "0.00"), 12);
                Else
                    mTotalDr = mTotalDr + Format(Rec!fltAmount, "0.00")
                    mDailyDrTotal = mDailyDrTotal + Format(Rec!fltAmount, "0.00")
                    Print #gbFileNO, Tab(58); PadL(Format(Rec!fltAmount, "0.00"), 12);
                End If
                
                Print #gbFileNO, Tab(84); "|"; PadL(Format(mTotalDr - mTotalCr, "0.00"), 11); "   ";
                Print #gbFileNO, Rec!vchGroup; "  "; PadL(Trim(str(Rec!intTransactionID)), 4); "   "; IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo);
                If IsDate(Rec!dtInstrumentDate) Then
                    Print #gbFileNO, "\" + DdMmmYy(Rec!dtInstrumentDate)
                Else
                    Print #gbFileNO,
                End If
                mLineNo = mLineNo + 1
                
                
                Rec.MoveNext
                If Not Rec.EOF Then
                    If Month(mDt) <> Month(Rec!dtTransactionDate) Then
                        mDt = DateSerial(Year(Rec!dtTransactionDate), Month(Rec!dtTransactionDate), 1)
                        If (mLineNo + 4) >= 67 Then
                            For mLoop = mLineNo To 67
                                Print #gbFileNO,
                            Next
                            mPageNo = mPageNo + 1
                            mLineNo = 0
                            Print #gbFileNO,
                            Print #gbFileNO, Tab(100); "Page No:"; mPageNo
                            Print #gbFileNO,
                            Print #gbFileNO,
                            Print #gbFileNO, ' Page Ending on 72 Line
                            
                            
                            Print #gbFileNO,
                            Print #gbFileNO,
                            Print #gbFileNO,
                            Print #gbFileNO, Tab(10); objAc.AccountCode, objAc.AccountHead
                            Print #gbFileNO, "---------------------------------------------------------------------------------------------------------------------------------"
                            Print #gbFileNO, "Vou.ID Trn. Date   Head Code  Account Head                      Debit       Credit      Balance Type Tr.ID   Cheque No.\Date"
                            Print #gbFileNO, "---------------------------------------------------------------------------------------------------------------------------------"
                            mLineNo = 7
                        End If
                        
                        
lblSubTotal:
                        Print #gbFileNO, Tab(56); "=============================="
                        Print #gbFileNO, Tab(56); PadL(Format(mDailyDrTotal, "0.00"), 14);
                        Print #gbFileNO, Tab(71); PadL(Format(mDailyCrTotal, "0.00"), 14);
                        Print #gbFileNO, Tab(56); "=============================="
                        If (mTotalDr - mTotalCr) > -1 Then
                            Print #gbFileNO, Tab(45); "Balance  : "; Tab(56); PadL(Format(mTotalDr - mTotalCr, "0.00"), 14)
                        Else
                            Print #gbFileNO, Tab(45); "Balance  : "; Tab(71); PadL(Format(mTotalCr - mTotalDr, "0.00"), 14)
                        End If
                        mDailyDrTotal = 0
                        mDailyCrTotal = 0
                         mLineNo = mLineNo + 4
                    End If
                Else
                    GoTo lblSubTotal:
                End If
            Wend
        End If
        Close #gbFileNO
        ShellPad
    End If
End Sub

Private Sub Form_Activate()
    Me.Left = (frmMenu.Width - Me.Width) / 2
End Sub

Private Sub txtDate_LostFocus()
    txtDate.Text = DdMmmYy(txtDate)
End Sub
