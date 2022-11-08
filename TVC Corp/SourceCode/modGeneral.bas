Attribute VB_Name = "modGeneral"
Option Explicit

Public Sub GeneratePayBillJournals(mPaymentOrderNo As Double, Optional mPreviousYearMode As Integer = 0)
        Dim mVoucher            As uVoucher
        Dim mVouChildTbl        As uVChild
        
        Dim mTranTable          As uTr
        Dim mTranChildTbl       As uTrChild
        
        Dim arrInput            As Variant
        Dim arrOutPut           As Variant
        Dim mintVoucherID       As Variant
        Dim mintTransactionID   As Variant
        Dim mCommonDescription  As String
        
        Dim objDb               As New clsDB
        Dim Rec                 As New ADODB.Recordset
        Dim RecChild            As New ADODB.Recordset
        Dim mCn                As New ADODB.Connection
        Dim mCnRead             As New ADODB.Connection
        Dim mSql                As String
        
        Dim mNetSalaryAmt       As Double
        Dim mGrossSalaryAmt     As Double
        Dim mPensionAmt         As Double
        Dim mCpAmt              As Double
        Dim mSLNo               As Integer
        Dim mSalaryHeadCode     As String
        Dim mSalaryHeadID       As Integer
        Dim objAc               As New clsAccounts
        
        Dim mPensionAmount      As Double
        Dim mPensionCategory    As Integer
        Dim mCPAmount      As Double
        Dim mCPCategory    As Integer
    
        '----------------------------------------------------------------------------- '
        ' Opening PaymentOrder Table And Child Tables
        '----------------------------------------------------------------------------- '
        objDb.SetConnection mCnRead
        mSql = "Select * From faPayOrder Where vchPayOrderNo = " & mPaymentOrderNo
        Rec.Open mSql, mCnRead, adOpenDynamic, adLockOptimistic, adCmdText
        If Rec.BOF And Rec.EOF Then
            MsgBox "No Pay Order Found For Generate Pay Bill Journals", vbInformation
            Exit Sub
        Else
            mSql = "Select * From faPayOrderChild Where intPayOrderID = " & Rec!intPayOrderID
            RecChild.Open mSql, mCnRead, adOpenDynamic, adLockOptimistic, adCmdText
            If RecChild.BOF And RecChild.EOF Then
                MsgBox "Payment Order Details not found for this Pay Order", vbInformation
                Exit Sub
            End If
        End If
        
        '----------------------------------------------------------------- '
        ' First Journal Voucher                                            '
        ' Salary A/c Dr to                                                 '
        '   Gross Salary Payabl                                            '
        '----------------------------------------------------------------- '
        ' Debit                                                            '
        '                                                                  '
        ' 210100101        Salaries -Secretary                             '
        ' 210100102        Salaries - Municipal Engineer                   '
        ' 210100103        Salaries - Health Officer                       '
        ' 210100104        Salaries - Permanent Staff                      '
        ' 210100105        Salaries - Temporary Staff                      '
        ' 210100106        Salaries - Contingent Staff                     '
        '                                                                  '
        ' Credit                                                           '
        ' 350110100        Gross Salary Payable                            '
        '                                                                  '
        '----------------------------------------------------------------- '
        
         '''''''''----------------------------------
        
        ''----Added On 4-12-12 By Anisha
        RecChild.MoveFirst
        While Not RecChild.EOF
            If RecChild!tnyCategoryFlag = 5 Then
                mPensionAmount = RecChild!numAmount
                mPensionCategory = 5
                'GoTo Pension:
            End If
            If RecChild!tnyCategoryFlag = 6 Then
                mCPAmount = RecChild!numAmount
                mCPCategory = 6
                'GoTo CP:
            End If
            RecChild.MoveNext
        Wend
        '''''''''----------------------------------
        
        RecChild.MoveFirst
        While Not RecChild.EOF
            If RecChild!tnyCategoryFlag = 1 Then
                mSalaryHeadID = Rec!intCashOrBankHeadID
                GoTo GrossSalary
            End If
            RecChild.MoveNext
        Wend
        GoTo ErrNoGr:
        
GrossSalary:
        With mVoucher
            .intVoucherID_1 = -1
            .intLocalBodyID_2 = gbLocalBodyID
            .intTransactionID_3 = Null
            .intTransactionTypeID_4 = Rec!intTransactionTypeID
            .tnyVoucherTypeID_5 = 40
            .intVoucherNo_6 = Null
            .intBookNo_7 = Null
            .dtDate_8 = Rec!dtDueDate
            .fltAmount_9 = RecChild!numAmount
             mGrossSalaryAmt = RecChild!numAmount
            .intInstrumentTypeID_10 = Null
            .vchInstrumentNo_11 = Null
            .dtInstrumentDate_12 = Null
            .vchDescription_13 = Rec!vchDescription
            .numZoneID_14 = Null
            .numWardID_15 = Null
            .intDoorNoP1_16 = Null
            .vchDoorNoP2_17 = Null
            .vchDoorNoP3_18 = Null
            .intUserID_19 = gbUserID
            .intCounterID_20 = gbCounterID
            .numSubLedgerID_21 = Null
            .intKeyID1_22 = mSalaryHeadID  'gbAcHeadIDNetSalaryPayable  'Debit to Net Salary Payable
            .intKeyID2_23 = mPaymentOrderNo
            .intExternalApplicationID_24 = 115
            .intExternalModuleID_25 = 61 'PaymentOrder-SthapanaInterface Module
            
            '''To Get FinancialYearID For previous Year Transactions
            'If .dtDate_8 < DateAdd("yyyy", -1, gbStartingDate) Or .dtDate_8 > DateAdd("yyyy", -1, gbEndingDate) Then
            If mPreviousYearMode Then
                .intFinancialYearID_26 = gbFinancialYearID - 1
            Else
                .intFinancialYearID_26 = gbFinancialYearID
            End If
            .tnyShiftID_27 = Null
            .tnyPrintFlag_28 = Null
            .tnyCancelFlag_29 = Null
            .vchBank_33 = Null
            .vchBankPlace_34 = Null
            .intFundID_35 = Null
            .numSeatID = gbSeatID
            .intSessionID = gbSessionID
            .vchRefNo = Null
            .fltRoundOff = Null
            .fltAdvAmtAdj = Null
            .numInwardNo = Null
            .tnyStatus_32 = 0
            .numLocationID = Null
            
            arrInput = Array(.intVoucherID_1, .intLocalBodyID_2, .intTransactionID_3, .intTransactionTypeID_4, _
            .tnyVoucherTypeID_5, .intVoucherNo_6, .intBookNo_7, .dtDate_8, _
            .fltAmount_9, .intInstrumentTypeID_10, .vchInstrumentNo_11, .dtInstrumentDate_12, _
            .vchDescription_13, .numZoneID_14, .numWardID_15, .intDoorNoP1_16, _
            .vchDoorNoP2_17, .vchDoorNoP3_18, .intUserID_19, .intCounterID_20, _
            .numSubLedgerID_21, .intKeyID1_22, .intKeyID2_23, .intExternalApplicationID_24, _
            .intExternalModuleID_25, .intFinancialYearID_26, .tnyShiftID_27, _
            .tnyPrintFlag_28, .tnyCancelFlag_29, .vchBank_33, .vchBankPlace_34, _
            .intFundID_35, .numSeatID, .intSessionID, .vchRefNo, _
            .fltRoundOff, .fltAdvAmtAdj, .numInwardNo, .tnyStatus_32, _
            .numLocationID)
        End With
        
        objAc.SetAccountID mSalaryHeadID
        If objAc.AccountHeadID > 0 Then
            mSalaryHeadCode = objAc.AccountCode   ' Debit Head
        Else
            GoTo ErrNoPenHeadNotFound:
        End If
        
        
        objDb.CreateNewConnection mCn, enuSourceString.Saankhya
        'mCn.BeginTrans
        'On Error GoTo ErrRollBack:
        objDb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCn
        If IsNumeric(arrOutPut(0, 0)) Then
            mintVoucherID = arrOutPut(0, 0)
        Else
            MsgBox "Error : Voucher Table didnt able to save!", vbInformation
            GoTo ErrRollBack:
        End If
        
        'Note:- Gross Salary AccountHead to the Voucher Child
        With mVouChildTbl
            .intVoucherID_1 = mintVoucherID
            .intLocalBodyID_2 = gbLocalBodyID
            .intSlNo_3 = 1
            .intAccountHeadID_4 = gbAcHeadIDGrossSalaryPayable
            .tnyDebitOrCredit_5 = 0
            If IsDate(Rec!dtKeyDate) Then
                .intYearID_6 = Year(Rec!dtKeyDate)
                .tnyPeriodID_7 = Month(Rec!dtKeyDate)
            Else
                .intYearID_6 = Null
                .tnyPeriodID_7 = Null
            End If
            .tnyArrearFlag_8 = Null
            .numDemandID_9 = Rec!intKeyID
            .fltAmount_10 = RecChild!numAmount
        
            arrInput = Array(.intVoucherID_1, _
            .intLocalBodyID_2, _
            .intSlNo_3, _
            .intAccountHeadID_4, _
            .tnyDebitOrCredit_5, _
            .intYearID_6, _
            .tnyPeriodID_7, _
            .tnyArrearFlag_8, _
            .numDemandID_9, _
            .fltAmount_10)
            objDb.ExecuteSP "spSaveVoucherChild", arrInput, , , mCn
        End With
        
        
        With mTranTable
            .intTransactionID = -1
            .intLocalBodyID = gbLocalBodyID
            .dtTransactionDate = Rec!dtDueDate
            'If .dtTransactionDate < DateAdd("yyyy", -1, gbStartingDate) Or .dtTransactionDate > DateAdd("yyyy", -1, gbEndingDate) Then
            If mPreviousYearMode Then
                .intFinancialYearID = gbFinancialYearID - 1
            Else
                .intFinancialYearID = gbFinancialYearID
            End If
            .intExternalApplicationID = Null
            .intExternalApplicationModuleID = Null
            .intFunctionID = IIf(Rec!intFunctionID = 0, Null, Rec!intFunctionID)
            .intFunctionaryID = IIf(Rec!intFunctionaryID = 0, Null, Rec!intFunctionaryID)
            .intFieldID = Null
            .intFundID = gbFundID
            .intBudgetCentreID = Null
            .vchNarration = Rec!vchDescription
            .intTransactionTypeID = Rec!intTransactionTypeID
            .intProcessID = Null
            .vchGroup = "JV"
            .intGroupID = 40
            .intKeyID = Null
            .numSubLedgerID = Null
            .numUserID = gbUserID
            .intVoucherID = mintVoucherID
            
            arrInput = Array(.intTransactionID, _
            .intLocalBodyID, _
            .intFinancialYearID, _
            .dtTransactionDate, _
            .intExternalApplicationID, _
            .intExternalApplicationModuleID, _
            .intFunctionID, _
            .intFunctionaryID, _
            .intFieldID, _
            .intFundID, _
            .intBudgetCentreID, _
            .vchNarration, _
            .intTransactionTypeID, _
            .intProcessID, _
            .vchGroup, _
            .intGroupID, _
            .intKeyID, _
            .numSubLedgerID, _
            .numUserID, _
            .intVoucherID)
        
        End With
        
        objDb.ExecuteSP "spSaveTransactions", arrInput, arrOutPut, , mCn
        If IsNumeric(arrOutPut(0, 0)) Then
            mintTransactionID = arrOutPut(0, 0)
        End If
        
        With mTranChildTbl
            .intTransactionID = mintTransactionID
            .intSerialNo = 1
            .intAccountHeadID = mSalaryHeadID 'Rec!intCashOrBankHeadID
            .fltAmount = RecChild!numAmount
            .tinDebitOrCreditFlag = 1
            .intByAccountHeadID = Null
            .vchNarration = RecChild!vchDescription
            .intFundID = gbFundID
            
            arrInput = Array(.intTransactionID, _
            .intSerialNo, _
            .intAccountHeadID, _
            .fltAmount, _
            .tinDebitOrCreditFlag, _
            .intByAccountHeadID, _
            .vchNarration, _
            .intFundID)
            
            objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCn
            
            
            .intSerialNo = 2
            .intAccountHeadID = gbAcHeadIDGrossSalaryPayable
            .tinDebitOrCreditFlag = 0
            .intByAccountHeadID = Rec!intCashOrBankHeadID
            
            arrInput = Array(.intTransactionID, _
            .intSerialNo, _
            .intAccountHeadID, _
            .fltAmount, _
            .tinDebitOrCreditFlag, _
            .intByAccountHeadID, _
            .vchNarration, _
            .intFundID)
            
            objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCn
        End With
        
        
TryDeductions:

        '----------------------------------------------------------------- '
        ' Second Journal Voucher                                           '
        ' Gross Salary A/c Dr to                                           '
        '   Deductions          Cr                                         '
        '   Net Salary Payable  Cr                                         '
        '----------------------------------------------------------------- '
        ' Debit                                                            '
        ' 350110100        Gross Salary Payable                            '
        '                                                                  '
        ' Credit                                                           '
        '     (---      Deduction Heads  --- )                             '
        '     350110200    Net Salary Payable                              '
        '----------------------------------------------------------------- '
        RecChild.MoveFirst
        While Not RecChild.EOF
            If RecChild!tnyCategoryFlag = 3 Then
                mNetSalaryAmt = RecChild!numAmount
                GoTo Deductions
            End If
            RecChild.MoveNext
        Wend
        GoTo TryPension:
Deductions:
        
        With mVoucher
            .intVoucherID_1 = -1
            .intLocalBodyID_2 = gbLocalBodyID
            .intTransactionID_3 = Null
            .intTransactionTypeID_4 = Rec!intTransactionTypeID
            .tnyVoucherTypeID_5 = 40
            .intVoucherNo_6 = Null
            .intBookNo_7 = Null
            .dtDate_8 = Rec!dtDueDate
            .fltAmount_9 = mGrossSalaryAmt
            .intInstrumentTypeID_10 = Null
            .vchInstrumentNo_11 = Null
            .dtInstrumentDate_12 = Null
            .vchDescription_13 = Rec!vchDescription
            .numZoneID_14 = Null
            .numWardID_15 = Null
            .intDoorNoP1_16 = Null
            .vchDoorNoP2_17 = Null
            .vchDoorNoP3_18 = Null
            .intUserID_19 = gbUserID
            .intCounterID_20 = gbCounterID
            .numSubLedgerID_21 = Null
            .intKeyID1_22 = gbAcHeadIDGrossSalaryPayable  'Credit to GrossSalary Payable
            .intKeyID2_23 = mPaymentOrderNo
            .intExternalApplicationID_24 = 115
            .intExternalModuleID_25 = 61 'PaymentOrder-SthapanaInterface Module
            '''To Get FinancialYearID For previous Year Transactions
            'If .dtDate_8 < DateAdd("yyyy", -1, gbStartingDate) Or .dtDate_8 > DateAdd("yyyy", -1, gbEndingDate) Then
            If mPreviousYearMode Then
                .intFinancialYearID_26 = gbFinancialYearID - 1
            Else
                .intFinancialYearID_26 = gbFinancialYearID
            End If
            
            '.intFinancialYearID_26 = gbFinancialYearID
            .tnyShiftID_27 = Null
            .tnyPrintFlag_28 = Null
            .tnyCancelFlag_29 = Null
            .vchBank_33 = Null
            .vchBankPlace_34 = Null
            .intFundID_35 = Null
            .numSeatID = gbSeatID
            .intSessionID = gbSessionID
            .vchRefNo = Null
            .fltRoundOff = Null
            .fltAdvAmtAdj = Null
            .numInwardNo = Null
            .tnyStatus_32 = 0
            .numLocationID = Null
            
            arrInput = Array(.intVoucherID_1, .intLocalBodyID_2, .intTransactionID_3, .intTransactionTypeID_4, _
            .tnyVoucherTypeID_5, .intVoucherNo_6, .intBookNo_7, .dtDate_8, _
            .fltAmount_9, .intInstrumentTypeID_10, .vchInstrumentNo_11, .dtInstrumentDate_12, _
            .vchDescription_13, .numZoneID_14, .numWardID_15, .intDoorNoP1_16, _
            .vchDoorNoP2_17, .vchDoorNoP3_18, .intUserID_19, .intCounterID_20, _
            .numSubLedgerID_21, .intKeyID1_22, .intKeyID2_23, .intExternalApplicationID_24, _
            .intExternalModuleID_25, .intFinancialYearID_26, .tnyShiftID_27, _
            .tnyPrintFlag_28, .tnyCancelFlag_29, .vchBank_33, .vchBankPlace_34, _
            .intFundID_35, .numSeatID, .intSessionID, .vchRefNo, _
            .fltRoundOff, .fltAdvAmtAdj, .numInwardNo, .tnyStatus_32, _
            .numLocationID)
        End With
        objDb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCn
        If IsNumeric(arrOutPut(0, 0)) Then
            mintVoucherID = arrOutPut(0, 0)
        End If
        
        
        With mTranTable
            .intTransactionID = -1
            .intLocalBodyID = gbLocalBodyID
            '.intFinancialYearID = gbFinancialYearID
            .dtTransactionDate = Rec!dtDueDate
            '''To Get FinancialYearID For previous Year Transactions
            'If .dtTransactionDate < DateAdd("yyyy", -1, gbStartingDate) Or .dtTransactionDate > DateAdd("yyyy", -1, gbEndingDate) Then
            If mPreviousYearMode Then
                .intFinancialYearID = gbFinancialYearID - 1
            Else
                .intFinancialYearID = gbFinancialYearID
            End If
            .intExternalApplicationID = Null
            .intExternalApplicationModuleID = Null
            .intFunctionID = IIf(Rec!intFunctionID = 0, Null, Rec!intFunctionID)
            .intFunctionaryID = IIf(Rec!intFunctionaryID = 0, Null, Rec!intFunctionaryID)
            .intFieldID = Null
            .intFundID = gbFundID
            .intBudgetCentreID = Null
            .vchNarration = Rec!vchDescription
            .intTransactionTypeID = Rec!intTransactionTypeID
            .intProcessID = Null
            .vchGroup = "JV"
            .intGroupID = 40
            .intKeyID = Null
            .numSubLedgerID = Null
            .numUserID = gbUserID
            .intVoucherID = mintVoucherID
            
            arrInput = Array(.intTransactionID, _
            .intLocalBodyID, _
            .intFinancialYearID, _
            .dtTransactionDate, _
            .intExternalApplicationID, _
            .intExternalApplicationModuleID, _
            .intFunctionID, _
            .intFunctionaryID, _
            .intFieldID, _
            .intFundID, _
            .intBudgetCentreID, _
            .vchNarration, _
            .intTransactionTypeID, _
            .intProcessID, _
            .vchGroup, _
            .intGroupID, _
            .intKeyID, _
            .numSubLedgerID, _
            .numUserID, _
            .intVoucherID)
        
        End With
        
        objDb.ExecuteSP "spSaveTransactions", arrInput, arrOutPut, , mCn
        If IsNumeric(arrOutPut(0, 0)) Then
            mintTransactionID = arrOutPut(0, 0)
        End If
        
        'Note:-Gross Salary Payable A/c Debtor
        With mTranChildTbl
            .intTransactionID = mintTransactionID
            .intSerialNo = 1
            .intAccountHeadID = gbAcHeadIDGrossSalaryPayable
            .fltAmount = mGrossSalaryAmt
            .tinDebitOrCreditFlag = 1
            .intByAccountHeadID = Null
            .vchNarration = RecChild!vchDescription
            .intFundID = gbFundID
            
            arrInput = Array(.intTransactionID, _
            .intSerialNo, _
            .intAccountHeadID, _
            .fltAmount, _
            .tinDebitOrCreditFlag, _
            .intByAccountHeadID, _
            .vchNarration, _
            .intFundID)
            
            objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCn
        End With
        
        mSLNo = 0
        RecChild.MoveFirst
        While Not RecChild.EOF
            If RecChild!tnyCategoryFlag = 2 Or RecChild!tnyCategoryFlag = 3 Then
                        mSLNo = mSLNo + 1
                        'Note:- Deduction Heads and Net Salary to Voucher Child
                        With mVouChildTbl
                            .intVoucherID_1 = mintVoucherID
                            .intLocalBodyID_2 = gbLocalBodyID
                            .intSlNo_3 = mSLNo
                            .intAccountHeadID_4 = RecChild!intAccountHeadID
                            .tnyDebitOrCredit_5 = 0
                            If IsDate(Rec!dtKeyDate) Then
                                .intYearID_6 = Year(Rec!dtKeyDate)
                                .tnyPeriodID_7 = Month(Rec!dtKeyDate)
                            Else
                                .intYearID_6 = Null
                                .tnyPeriodID_7 = Null
                            End If
                            .tnyArrearFlag_8 = Null
                            .numDemandID_9 = Rec!intKeyID
                            .fltAmount_10 = RecChild!numAmount
                        
                            arrInput = Array(.intVoucherID_1, _
                            .intLocalBodyID_2, _
                            .intSlNo_3, _
                            .intAccountHeadID_4, _
                            .tnyDebitOrCredit_5, _
                            .intYearID_6, _
                            .tnyPeriodID_7, _
                            .tnyArrearFlag_8, _
                            .numDemandID_9, _
                            .fltAmount_10)
                            objDb.ExecuteSP "spSaveVoucherChild", arrInput, , , mCn
                        End With
                                
                        With mTranChildTbl
                            .intTransactionID = mintTransactionID
                            .intSerialNo = mSLNo + 1
                            .intAccountHeadID = RecChild!intAccountHeadID
                            .fltAmount = RecChild!numAmount
                            .tinDebitOrCreditFlag = 0
                            .intByAccountHeadID = gbAcHeadIDGrossSalaryPayable
                            .vchNarration = RecChild!vchDescription
                            .intFundID = gbFundID
                            
                            arrInput = Array(.intTransactionID, _
                            .intSerialNo, _
                            .intAccountHeadID, _
                            .fltAmount, _
                            .tinDebitOrCreditFlag, _
                            .intByAccountHeadID, _
                            .vchNarration, _
                            .intFundID)
                            
                            objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCn
                        End With
            End If
            RecChild.MoveNext
        Wend
        
        
        RecChild.MoveFirst
        While Not RecChild.EOF
            If RecChild!tnyCategoryFlag = 3 Then
                GoTo TryPension:
            End If
            RecChild.MoveNext
        Wend
        
       
        
        GoTo CleanUp:
TryPension:
        '----------------------------------------------------------------- '
        ' Third Journal Voucher                                           '
        ' Pention Contribution  A/c Dr to                                           '
        '   Pention Payable         Cr
        '----------------------------------------------------------------- '
        ' Permenent Staff   Dr - 210300104    Cr - 350110600
        ' Secretary         Dr - 210300101    Cr - 350110700
        ' Contingent        Dr - 210300201    Cr - 311700100
        '----------------------------------------------------------------- '
        ' 210100101        Salaries -Secretary                             '
        ' 210100102        Salaries - Municipal Engineer                   '
        ' 210100103        Salaries - Health Officer                       '
        ' 210100104        Salaries - Permanent Staff                      '
        ' 210100105        Salaries - Temporary Staff                      '
        ' 210100106        Salaries - Contingent Staff
        
       
        
        Dim mPenContributionHeadCode As String
        Dim mPenPayableHeadCode      As String
        '''''''''----------------------------------
        ''----Added On 4-12-12 By Anisha

        If mPensionCategory = 5 Then
            mPensionAmt = mPensionAmount
        Else
            If mCPCategory = 6 Then GoTo EndOfPension:
            mPensionAmt = Format(mGrossSalaryAmt * 15 / 100, "0#")
        End If
        ''-------------------------------------
        
        objAc.SetAccountID mSalaryHeadID
        If objAc.AccountHeadID > 0 Then
            mSalaryHeadCode = objAc.AccountCode   ' Debit Head
        Else
            GoTo ErrNoPenHeadNotFound:
        End If
        
        If mSalaryHeadCode = "210100101" Then ' Secretary
            mPenContributionHeadCode = "210300101"
            mPenPayableHeadCode = "350110700"
        ElseIf mSalaryHeadCode = "210100106" Then ' Contingent Staff
            mPenContributionHeadCode = "210300201"
            mPenPayableHeadCode = "311700100"
'''''        Else    ' Permenent Staff
'''''            mPenContributionHeadCode = "210300104"
'''''            mPenPayableHeadCode = "350110600"
'''''        End If
        ElseIf mSalaryHeadCode = "210100104" Or mSalaryHeadCode = "210100103" Or mSalaryHeadCode = "210100102" Then   ' Permenent Staff /Health Officer/Muncipal Engineer
            mPenContributionHeadCode = "210300104"
            mPenPayableHeadCode = "350110600"
        ''--- Added On 4.8.11
        ElseIf mSalaryHeadCode = "210100105" Then
            GoTo EndOfPension
        End If
        
        With mVoucher
            .intVoucherID_1 = -1
            .intLocalBodyID_2 = gbLocalBodyID
            .intTransactionID_3 = Null
            .intTransactionTypeID_4 = Rec!intTransactionTypeID
            .tnyVoucherTypeID_5 = 40
            .intVoucherNo_6 = Null
            .intBookNo_7 = Null
            .dtDate_8 = Rec!dtDueDate
            .fltAmount_9 = mPensionAmt
             mGrossSalaryAmt = RecChild!numAmount
            .intInstrumentTypeID_10 = Null
            .vchInstrumentNo_11 = Null
            .dtInstrumentDate_12 = Null
            
            .vchDescription_13 = "Being the Pension Contribution for the month of " & MonthName((Month(Rec!dtKeyDate))) & "," & Year(Rec!dtKeyDate)
            .numZoneID_14 = Null
            .numWardID_15 = Null
            .intDoorNoP1_16 = Null
            .vchDoorNoP2_17 = Null
            .vchDoorNoP3_18 = Null
            .intUserID_19 = gbUserID
            .intCounterID_20 = gbCounterID
            .numSubLedgerID_21 = Null
            objAc.SetAccountCode mPenContributionHeadCode
            If objAc.AccountHeadID > 0 Then
                .intKeyID1_22 = objAc.AccountHeadID  ' Debit Head
            Else
                GoTo ErrNoPenHeadNotFound:
            End If
            .intKeyID2_23 = mPaymentOrderNo
            .intExternalApplicationID_24 = 115
            .intExternalModuleID_25 = 61 'PaymentOrder-SthapanaInterface Module
            '.intFinancialYearID_26 = gbFinancialYearID
            'If .dtDate_8 < DateAdd("yyyy", -1, gbStartingDate) Or .dtDate_8 > DateAdd("yyyy", -1, gbEndingDate) Then
            If mPreviousYearMode Then
                .intFinancialYearID_26 = gbFinancialYearID - 1
            Else
                .intFinancialYearID_26 = gbFinancialYearID
            End If
            .tnyShiftID_27 = Null
            .tnyPrintFlag_28 = Null
            .tnyCancelFlag_29 = Null
            .vchBank_33 = Null
            .vchBankPlace_34 = Null
            .intFundID_35 = Null
            .numSeatID = gbSeatID
            .intSessionID = gbSessionID
            .vchRefNo = Null
            .fltRoundOff = Null
            .fltAdvAmtAdj = Null
            .numInwardNo = Null
            .tnyStatus_32 = 0
            .numLocationID = Null
            
            arrInput = Array(.intVoucherID_1, .intLocalBodyID_2, .intTransactionID_3, .intTransactionTypeID_4, _
            .tnyVoucherTypeID_5, .intVoucherNo_6, .intBookNo_7, .dtDate_8, _
            .fltAmount_9, .intInstrumentTypeID_10, .vchInstrumentNo_11, .dtInstrumentDate_12, _
            .vchDescription_13, .numZoneID_14, .numWardID_15, .intDoorNoP1_16, _
            .vchDoorNoP2_17, .vchDoorNoP3_18, .intUserID_19, .intCounterID_20, _
            .numSubLedgerID_21, .intKeyID1_22, .intKeyID2_23, .intExternalApplicationID_24, _
            .intExternalModuleID_25, .intFinancialYearID_26, .tnyShiftID_27, _
            .tnyPrintFlag_28, .tnyCancelFlag_29, .vchBank_33, .vchBankPlace_34, _
            .intFundID_35, .numSeatID, .intSessionID, .vchRefNo, _
            .fltRoundOff, .fltAdvAmtAdj, .numInwardNo, .tnyStatus_32, _
            .numLocationID)
        End With
        objDb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCn
        If IsNumeric(arrOutPut(0, 0)) Then
            mintVoucherID = arrOutPut(0, 0)
        Else
            MsgBox "Error : Voucher Table didnt able to save!", vbInformation
            Exit Sub
        End If
        
        'Note:- Pension Payable AccountHead to the Voucher Child
        With mVouChildTbl
            .intVoucherID_1 = mintVoucherID
            .intLocalBodyID_2 = gbLocalBodyID
            .intSlNo_3 = 1
            objAc.SetAccountCode mPenPayableHeadCode
            If objAc.AccountHeadID > 0 Then
                .intAccountHeadID_4 = objAc.AccountHeadID  ' Credit Head
            Else
                GoTo ErrNoPenHeadNotFound:
            End If
            .tnyDebitOrCredit_5 = 0
            .intYearID_6 = Year(Rec!dtKeyDate)
            .tnyPeriodID_7 = Month(Rec!dtKeyDate)
            .tnyArrearFlag_8 = Null
            .numDemandID_9 = Rec!intKeyID
            .fltAmount_10 = mPensionAmt
        
            arrInput = Array(.intVoucherID_1, _
            .intLocalBodyID_2, _
            .intSlNo_3, _
            .intAccountHeadID_4, _
            .tnyDebitOrCredit_5, _
            .intYearID_6, _
            .tnyPeriodID_7, _
            .tnyArrearFlag_8, _
            .numDemandID_9, _
            .fltAmount_10)
            objDb.ExecuteSP "spSaveVoucherChild", arrInput, , , mCn
        End With
        
        
        With mTranTable
            .intTransactionID = -1
            .intLocalBodyID = gbLocalBodyID
            '.intFinancialYearID = gbFinancialYearID
            .dtTransactionDate = Rec!dtDueDate
            'If .dtTransactionDate < DateAdd("yyyy", -1, gbStartingDate) Or .dtTransactionDate > DateAdd("yyyy", -1, gbEndingDate) Then
            If mPreviousYearMode Then
                .intFinancialYearID = gbFinancialYearID - 1
            Else
                .intFinancialYearID = gbFinancialYearID
            End If
            .intExternalApplicationID = Null
            .intExternalApplicationModuleID = Null
            .intFunctionID = IIf(Rec!intFunctionID = 0, Null, Rec!intFunctionID)
            .intFunctionaryID = IIf(Rec!intFunctionaryID = 0, Null, Rec!intFunctionaryID)
            .intFieldID = Null
            .intFundID = gbFundID
            .intBudgetCentreID = Null
            .vchNarration = "Being the Pension Contribution for the month of " & MonthName(Month(Rec!dtKeyDate), True) & "," & Year(Rec!dtKeyDate)
            .intTransactionTypeID = Rec!intTransactionTypeID
            .intProcessID = Null
            .vchGroup = "JV"
            .intGroupID = 40
            .intKeyID = Null
            .numSubLedgerID = Null
            .numUserID = gbUserID
            .intVoucherID = mintVoucherID
            
            arrInput = Array(.intTransactionID, _
            .intLocalBodyID, _
            .intFinancialYearID, _
            .dtTransactionDate, _
            .intExternalApplicationID, _
            .intExternalApplicationModuleID, _
            .intFunctionID, _
            .intFunctionaryID, _
            .intFieldID, _
            .intFundID, _
            .intBudgetCentreID, _
            .vchNarration, _
            .intTransactionTypeID, _
            .intProcessID, _
            .vchGroup, _
            .intGroupID, _
            .intKeyID, _
            .numSubLedgerID, _
            .numUserID, _
            .intVoucherID)
        
        End With
        
        objDb.ExecuteSP "spSaveTransactions", arrInput, arrOutPut, , mCn
        If IsNumeric(arrOutPut(0, 0)) Then
            mintTransactionID = arrOutPut(0, 0)
        End If
        
        With mTranChildTbl
            .intTransactionID = mintTransactionID
            .intSerialNo = 1
            
            objAc.SetAccountCode mPenContributionHeadCode
            If objAc.AccountHeadID > 0 Then
                .intAccountHeadID = objAc.AccountHeadID  ' Debit Head
            Else
                GoTo ErrNoPenHeadNotFound:
            End If
            .fltAmount = mPensionAmt
            .tinDebitOrCreditFlag = 1
            .intByAccountHeadID = Null
            .vchNarration = RecChild!vchDescription
            .intFundID = gbFundID
            
            arrInput = Array(.intTransactionID, _
            .intSerialNo, _
            .intAccountHeadID, _
            .fltAmount, _
            .tinDebitOrCreditFlag, _
            .intByAccountHeadID, _
            .vchNarration, _
            .intFundID)
            
            objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCn
            
            .intSerialNo = 2
            objAc.SetAccountCode mPenPayableHeadCode
            If objAc.AccountHeadID > 0 Then
                .intAccountHeadID = objAc.AccountHeadID  ' Credit Head
            Else
                GoTo ErrNoPenHeadNotFound:
            End If
            .tinDebitOrCreditFlag = 0
            
            objAc.SetAccountCode mPenContributionHeadCode
            If objAc.AccountHeadID > 0 Then
                .intByAccountHeadID = objAc.AccountHeadID   ' Debit Head
            Else
                GoTo ErrNoPenHeadNotFound:
            End If
            
            arrInput = Array(.intTransactionID, _
            .intSerialNo, _
            .intAccountHeadID, _
            .fltAmount, _
            .tinDebitOrCreditFlag, _
            .intByAccountHeadID, _
            .vchNarration, _
            .intFundID)
            
            objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCn
        End With
EndOfPension:

TryCP:
        '----------------------------------------------------------------- '
        ' Fourth Journal Voucher                                           '
        ' Contributory Pension Fund A/c Dr to                                           '
        '   Pension Payable         Cr
        '----------------------------------------------------------------- '
        ' Permenent Staff   Dr - 210300500    Cr - 350110601
        ' Secretary         Dr - 210300500    Cr - 350110601
        ' Contingent        Dr - 210300500    Cr - 350110601
        '----------------------------------------------------------------- '
        ' 210100101        Salaries -Secretary                             '
        ' 210100102        Salaries - Municipal Engineer                   '
        ' 210100103        Salaries - Health Officer                       '
        ' 210100104        Salaries - Permanent Staff                      '
        ' 210100105        Salaries - Temporary Staff                      '
        ' 210100106        Salaries - Contingent Staff
        
       
        
        Dim mCPContributionHeadCode As String
        Dim mCPPayableHeadCode      As String
        '''''''''----------------------------------
        ''----Added On 27-11-17 By Anisha
         
        
        If mCPCategory = 6 Then
            mCpAmt = mCPAmount
        Else
            mCPAmount = 0
            GoTo EndOfCP:
        End If
        ''-------------------------------------
        
        objAc.SetAccountID mSalaryHeadID
        If objAc.AccountHeadID > 0 Then
            mSalaryHeadCode = objAc.AccountCode   ' Debit Head
        Else
            GoTo ErrNoPenHeadNotFound:
        End If
        
        If mSalaryHeadCode = "210100101" Then ' Secretary
            mCPContributionHeadCode = "210300500"
            mCPPayableHeadCode = "350110601"
        ElseIf mSalaryHeadCode = "210100106" Then ' Contingent Staff
            mCPContributionHeadCode = "210300500"
            mCPPayableHeadCode = "350110601"
        ElseIf mSalaryHeadCode = "210100104" Or mSalaryHeadCode = "210100103" Or mSalaryHeadCode = "210100102" Then   ' Permenent Staff /Health Officer/Muncipal Engineer
            mCPContributionHeadCode = "210300500"
            mCPPayableHeadCode = "350110601"
        ''--- Added On 4.8.11
        ElseIf mSalaryHeadCode = "210100105" Then
            GoTo EndOfCP
        End If
        
        With mVoucher
            .intVoucherID_1 = -1
            .intLocalBodyID_2 = gbLocalBodyID
            .intTransactionID_3 = Null
            .intTransactionTypeID_4 = Rec!intTransactionTypeID
            .tnyVoucherTypeID_5 = 40
            .intVoucherNo_6 = Null
            .intBookNo_7 = Null
            .dtDate_8 = Rec!dtDueDate
            .fltAmount_9 = mCpAmt
             mGrossSalaryAmt = RecChild!numAmount
            .intInstrumentTypeID_10 = Null
            .vchInstrumentNo_11 = Null
            .dtInstrumentDate_12 = Null
            
            .vchDescription_13 = "Being the Contributory Pension for the month of " & MonthName((Month(Rec!dtKeyDate))) & "," & Year(Rec!dtKeyDate)
            .numZoneID_14 = Null
            .numWardID_15 = Null
            .intDoorNoP1_16 = Null
            .vchDoorNoP2_17 = Null
            .vchDoorNoP3_18 = Null
            .intUserID_19 = gbUserID
            .intCounterID_20 = gbCounterID
            .numSubLedgerID_21 = Null
            objAc.SetAccountCode mCPContributionHeadCode
            If objAc.AccountHeadID > 0 Then
                .intKeyID1_22 = objAc.AccountHeadID  ' Debit Head
            Else
                GoTo ErrNoPenHeadNotFound:
            End If
            .intKeyID2_23 = mPaymentOrderNo
            .intExternalApplicationID_24 = 115
            .intExternalModuleID_25 = 61 'PaymentOrder-SthapanaInterface Module
            '.intFinancialYearID_26 = gbFinancialYearID
            'If .dtDate_8 < DateAdd("yyyy", -1, gbStartingDate) Or .dtDate_8 > DateAdd("yyyy", -1, gbEndingDate) Then
            If mPreviousYearMode Then
                .intFinancialYearID_26 = gbFinancialYearID - 1
            Else
                .intFinancialYearID_26 = gbFinancialYearID
            End If
            .tnyShiftID_27 = Null
            .tnyPrintFlag_28 = Null
            .tnyCancelFlag_29 = Null
            .vchBank_33 = Null
            .vchBankPlace_34 = Null
            .intFundID_35 = Null
            .numSeatID = gbSeatID
            .intSessionID = gbSessionID
            .vchRefNo = Null
            .fltRoundOff = Null
            .fltAdvAmtAdj = Null
            .numInwardNo = Null
            .tnyStatus_32 = 0
            .numLocationID = Null
            
            arrInput = Array(.intVoucherID_1, .intLocalBodyID_2, .intTransactionID_3, .intTransactionTypeID_4, _
            .tnyVoucherTypeID_5, .intVoucherNo_6, .intBookNo_7, .dtDate_8, _
            .fltAmount_9, .intInstrumentTypeID_10, .vchInstrumentNo_11, .dtInstrumentDate_12, _
            .vchDescription_13, .numZoneID_14, .numWardID_15, .intDoorNoP1_16, _
            .vchDoorNoP2_17, .vchDoorNoP3_18, .intUserID_19, .intCounterID_20, _
            .numSubLedgerID_21, .intKeyID1_22, .intKeyID2_23, .intExternalApplicationID_24, _
            .intExternalModuleID_25, .intFinancialYearID_26, .tnyShiftID_27, _
            .tnyPrintFlag_28, .tnyCancelFlag_29, .vchBank_33, .vchBankPlace_34, _
            .intFundID_35, .numSeatID, .intSessionID, .vchRefNo, _
            .fltRoundOff, .fltAdvAmtAdj, .numInwardNo, .tnyStatus_32, _
            .numLocationID)
        End With
        objDb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCn
        If IsNumeric(arrOutPut(0, 0)) Then
            mintVoucherID = arrOutPut(0, 0)
        Else
            MsgBox "Error : Voucher Table didnt able to save!", vbInformation
            Exit Sub
        End If
        
        'Note:- Pension Payable AccountHead to the Voucher Child
        With mVouChildTbl
            .intVoucherID_1 = mintVoucherID
            .intLocalBodyID_2 = gbLocalBodyID
            .intSlNo_3 = 1
            objAc.SetAccountCode mCPPayableHeadCode
            If objAc.AccountHeadID > 0 Then
                .intAccountHeadID_4 = objAc.AccountHeadID  ' Credit Head
            Else
                GoTo ErrNoPenHeadNotFound:
            End If
            .tnyDebitOrCredit_5 = 0
            .intYearID_6 = Year(Rec!dtKeyDate)
            .tnyPeriodID_7 = Month(Rec!dtKeyDate)
            .tnyArrearFlag_8 = Null
            .numDemandID_9 = Rec!intKeyID
            .fltAmount_10 = mCpAmt
        
            arrInput = Array(.intVoucherID_1, _
            .intLocalBodyID_2, _
            .intSlNo_3, _
            .intAccountHeadID_4, _
            .tnyDebitOrCredit_5, _
            .intYearID_6, _
            .tnyPeriodID_7, _
            .tnyArrearFlag_8, _
            .numDemandID_9, _
            .fltAmount_10)
            objDb.ExecuteSP "spSaveVoucherChild", arrInput, , , mCn
        End With
        
        
        With mTranTable
            .intTransactionID = -1
            .intLocalBodyID = gbLocalBodyID
            '.intFinancialYearID = gbFinancialYearID
            .dtTransactionDate = Rec!dtDueDate
            'If .dtTransactionDate < DateAdd("yyyy", -1, gbStartingDate) Or .dtTransactionDate > DateAdd("yyyy", -1, gbEndingDate) Then
            If mPreviousYearMode Then
                .intFinancialYearID = gbFinancialYearID - 1
            Else
                .intFinancialYearID = gbFinancialYearID
            End If
            .intExternalApplicationID = Null
            .intExternalApplicationModuleID = Null
            .intFunctionID = IIf(Rec!intFunctionID = 0, Null, Rec!intFunctionID)
            .intFunctionaryID = IIf(Rec!intFunctionaryID = 0, Null, Rec!intFunctionaryID)
            .intFieldID = Null
            .intFundID = gbFundID
            .intBudgetCentreID = Null
            .vchNarration = "Being the Contrubutory Pension for the month of " & MonthName(Month(Rec!dtKeyDate), True) & "," & Year(Rec!dtKeyDate)
            .intTransactionTypeID = Rec!intTransactionTypeID
            .intProcessID = Null
            .vchGroup = "JV"
            .intGroupID = 40
            .intKeyID = Null
            .numSubLedgerID = Null
            .numUserID = gbUserID
            .intVoucherID = mintVoucherID
            
            arrInput = Array(.intTransactionID, _
            .intLocalBodyID, _
            .intFinancialYearID, _
            .dtTransactionDate, _
            .intExternalApplicationID, _
            .intExternalApplicationModuleID, _
            .intFunctionID, _
            .intFunctionaryID, _
            .intFieldID, _
            .intFundID, _
            .intBudgetCentreID, _
            .vchNarration, _
            .intTransactionTypeID, _
            .intProcessID, _
            .vchGroup, _
            .intGroupID, _
            .intKeyID, _
            .numSubLedgerID, _
            .numUserID, _
            .intVoucherID)
        
        End With
        
        objDb.ExecuteSP "spSaveTransactions", arrInput, arrOutPut, , mCn
        If IsNumeric(arrOutPut(0, 0)) Then
            mintTransactionID = arrOutPut(0, 0)
        End If
        
        With mTranChildTbl
            .intTransactionID = mintTransactionID
            .intSerialNo = 1
            
            objAc.SetAccountCode mCPContributionHeadCode
            If objAc.AccountHeadID > 0 Then
                .intAccountHeadID = objAc.AccountHeadID  ' Debit Head
            Else
                GoTo ErrNoPenHeadNotFound:
            End If
            .fltAmount = mCpAmt
            .tinDebitOrCreditFlag = 1
            .intByAccountHeadID = Null
            .vchNarration = RecChild!vchDescription
            .intFundID = gbFundID
            
            arrInput = Array(.intTransactionID, _
            .intSerialNo, _
            .intAccountHeadID, _
            .fltAmount, _
            .tinDebitOrCreditFlag, _
            .intByAccountHeadID, _
            .vchNarration, _
            .intFundID)
            
            objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCn
            
            .intSerialNo = 2
            objAc.SetAccountCode mCPPayableHeadCode
            If objAc.AccountHeadID > 0 Then
                .intAccountHeadID = objAc.AccountHeadID  ' Credit Head
            Else
                GoTo ErrNoPenHeadNotFound:
            End If
            .tinDebitOrCreditFlag = 0
            
            objAc.SetAccountCode mCPContributionHeadCode
            If objAc.AccountHeadID > 0 Then
                .intByAccountHeadID = objAc.AccountHeadID   ' Debit Head
            Else
                GoTo ErrNoPenHeadNotFound:
            End If
            
            arrInput = Array(.intTransactionID, _
            .intSerialNo, _
            .intAccountHeadID, _
            .fltAmount, _
            .tinDebitOrCreditFlag, _
            .intByAccountHeadID, _
            .vchNarration, _
            .intFundID)
            
            objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCn
        End With
EndOfCP:

        'Note:- Changing the Status of PayOrder as Approved!
        mSql = "Update faPayOrder Set tnyStatus = 1 Where vchPayOrderNo = " & mPaymentOrderNo
        objDb.ExecuteSP mSql, , , , mCn, adCmdText
        'mCn.CommitTrans
CleanUp:
        Exit Sub
        
ErrNoPenHeadNotFound:
        MsgBox "Pension Head not found", vbInformation
        GoTo ErrRollBack:
ErrNoGr:
    MsgBox "No Gross Amount Found to Proceed!", vbInformation
    GoTo ErrRollBack:

ErrRollBack:
    'mCn.RollbackTrans
    Set mCn = Nothing
End Sub
    Public Function xmlTORecordset(sXml) As ADODB.Recordset
        ''To Read RecordSet From Xml
        Dim oStream         As ADODB.Stream
        Dim oRecordset      As ADODB.Recordset
        Set oStream = New ADODB.Stream
        oStream.Open
        oStream.WriteText sXml
        oStream.Position = 0
        Set oRecordset = New ADODB.Recordset
        oRecordset.Open oStream
        oStream.Close
        Set oStream = Nothing
        Set xmlTORecordset = oRecordset
        Set oRecordset = Nothing
    End Function

    Public Function funCancelPropertyTax(mRecieptNo As Variant, mVoucherID As Long)
        On Error GoTo Err:
            Dim mcnn As New ADODB.Connection
            'Dim mCnnSanchaya As New ADODB.Connection
            Dim Rec As New Recordset
            Dim mSql As String
            Dim objDb As New clsDB
            Dim arrIn As Variant
            Dim mQry As String
            
            
            Dim blnConfig As Boolean
            Dim blnOtherZoneOfficeFlag As Boolean
            
            If objDb.SetConnection(mcnn) Then
                mQry = "Select tnyLinkWithPropertyTax from faConfig"
                Rec.Open mQry, mcnn
                If IsNull(Rec!tnyLinkWithPropertyTax) Then
                    blnConfig = False
                ElseIf val(Rec!tnyLinkWithPropertyTax) = 1 Then
                    blnConfig = True
                Else
                    blnConfig = False
                End If
                If Rec.State = 1 Then Rec.Close
                
                mSql = "Select numZoneID as ZoneID from faVouchers Where intVoucherNo = " & Trim(mRecieptNo)
                Rec.Open mSql, mcnn
                If Not (Rec.EOF Or Rec.BOF) Then
                    If Rec!ZoneID <> gbLocationID Then
                        blnOtherZoneOfficeFlag = True
                    Else
                        blnOtherZoneOfficeFlag = False
                    End If
                End If
            Else
                MsgBox "Connection To Finance Does not Exist, Please Contact your System Administrtor", vbInformation
            End If
            
            If blnConfig = True Then
                Set mcnn = Nothing
                If objDb.CreateNewConnection(mcnn, enuSourceString.SanchayaLite) Then
                    If blnOtherZoneOfficeFlag = False Then
                        arrIn = Array(Trim(mRecieptNo))
                        objDb.ExecuteSP "spReverseDemandFromSaankhya", arrIn, , , mcnn
                    Else
                        '---------------------------------------------------------------'
                        ' Other Zone Office Collection Modified on 13-aug-2009 By cijith'
                        '---------------------------------------------------------------'
                        arrIn = Array(gbLocationID, mVoucherID)
                        objDb.ExecuteSP "HOSaanOtherCollectionCancel", arrIn, , , mcnn
                        '----------------------------------------------------------'
                    End If
                Else
                    MsgBox "Connection To Sanchaya Does not Exist, Please Contact your System Administrtor", vbInformation
                End If
            End If
        Exit Function
Err:
        MsgBox (Error$)
    End Function
Public Sub GeneratePayBillJournalsForPanchayat(mPaymentOrderNo As Double, Optional mPreviousYearMode As Integer = 0)   'ADDED BY MINU FOR PANCHAYAT ON 11-01-2011
        Dim mVoucher            As uVoucher
        Dim mVouChildTbl        As uVChild
        
        Dim mTranTable          As uTr
        Dim mTranChildTbl       As uTrChild
        
        Dim arrInput            As Variant
        Dim arrOutPut           As Variant
        Dim mintVoucherID       As Variant
        Dim mintTransactionID   As Variant
        Dim mCommonDescription  As String
        
        Dim objDb               As New clsDB
        Dim Rec                 As New ADODB.Recordset
        Dim RecChild            As New ADODB.Recordset
        Dim mCn                As New ADODB.Connection
        Dim mCnRead             As New ADODB.Connection
        Dim mSql                As String
        
        Dim mNetSalaryAmt       As Double
        Dim mGrossSalaryAmt     As Double
        Dim mPensionAmt         As Double
        Dim mSLNo               As Integer
        Dim mSalaryHeadCode     As String
        Dim mSalaryHeadID       As Integer
        Dim objAc               As New clsAccounts
        
        Dim mPensionAmount      As Double
        Dim mPensionCategory    As Integer
        '----------------------------------------------------------------------------- '
        ' Opening PaymentOrder Table And Child Tables
        '----------------------------------------------------------------------------- '
        objDb.SetConnection mCnRead
        mSql = "Select * From faPayOrder Where vchPayOrderNo = " & mPaymentOrderNo
        Rec.Open mSql, mCnRead, adOpenDynamic, adLockOptimistic, adCmdText
        If Rec.BOF And Rec.EOF Then
            MsgBox "No Pay Order Found For Generate Pay Bill Journals", vbInformation
            Exit Sub
        Else
            mSql = "Select * From faPayOrderChild Where intPayOrderID = " & Rec!intPayOrderID
            RecChild.Open mSql, mCnRead, adOpenDynamic, adLockOptimistic, adCmdText
            If RecChild.BOF And RecChild.EOF Then
                MsgBox "Payment Order Details not found for this Pay Order", vbInformation
                Exit Sub
            End If
        End If
        
        '----------------------------------------------------------------- '
        ' First Journal Voucher                                            '
        ' Salary A/c Dr to                                                 '
        '   Gross Salary Payabl                                            '
        '----------------------------------------------------------------- '
        ' Debit                                                            '
        '                                                                  '
        ' 210100101        Salaries -Secretary                             '
        ' 210100102        Salaries - Permanent Staff                      '
        ' 210100103        Salaries - Temporary Staff-Recruited through
        '                            Employment Exchange                   '
        ' 210100104        Salaries - Full Time Contingent Staff                      '
        ' 210100105        Salaries - Part Time Contingent Staff                      '
        ' 210100106        Salaries - Contract Staff                     '
        ' 210100107        Salaries - Honorarium Staff                                                                 '
        '
        ' Credit                                                           '
        ' 350110101        Gross Salary Payable                            '
        '                                                                  '
        '----------------------------------------------------------------- '
        RecChild.MoveFirst
        While Not RecChild.EOF
            If RecChild!tnyCategoryFlag = 5 Then
                mPensionAmount = RecChild!numAmount
                mPensionCategory = 5
            End If
            RecChild.MoveNext
        Wend
        
        RecChild.MoveFirst
        While Not RecChild.EOF
            If RecChild!tnyCategoryFlag = 1 Then
                mSalaryHeadID = Rec!intCashOrBankHeadID
                GoTo GrossSalary
            End If
            RecChild.MoveNext
        Wend
        GoTo ErrNoGr:
        
GrossSalary:
        With mVoucher
            .intVoucherID_1 = -1
            .intLocalBodyID_2 = gbLocalBodyID
            .intTransactionID_3 = Null
            .intTransactionTypeID_4 = Rec!intTransactionTypeID
            .tnyVoucherTypeID_5 = 40
            .intVoucherNo_6 = Null
            .intBookNo_7 = Null
            .dtDate_8 = Rec!dtDueDate
            .fltAmount_9 = RecChild!numAmount
             mGrossSalaryAmt = RecChild!numAmount
            .intInstrumentTypeID_10 = Null
            .vchInstrumentNo_11 = Null
            .dtInstrumentDate_12 = Null
            .vchDescription_13 = Rec!vchDescription
            .numZoneID_14 = Null
            .numWardID_15 = Null
            .intDoorNoP1_16 = Null
            .vchDoorNoP2_17 = Null
            .vchDoorNoP3_18 = Null
            .intUserID_19 = gbUserID
            .intCounterID_20 = gbCounterID
            .numSubLedgerID_21 = Null
            .intKeyID1_22 = mSalaryHeadID  'gbAcHeadIDNetSalaryPayable  'Debit to Net Salary Payable
            .intKeyID2_23 = mPaymentOrderNo
            .intExternalApplicationID_24 = 115
            .intExternalModuleID_25 = 61 'PaymentOrder-SthapanaInterface Module
            If mPreviousYearMode Then
                .intFinancialYearID_26 = gbFinancialYearID - 1
            Else
                .intFinancialYearID_26 = gbFinancialYearID
            End If
            
            '.intFinancialYearID_26 = gbFinancialYearID
            .tnyShiftID_27 = Null
            .tnyPrintFlag_28 = Null
            .tnyCancelFlag_29 = Null
            .vchBank_33 = Null
            .vchBankPlace_34 = Null
            .intFundID_35 = Null
            .numSeatID = gbSeatID
            .intSessionID = gbSessionID
            .vchRefNo = Null
            .fltRoundOff = Null
            .fltAdvAmtAdj = Null
            .numInwardNo = Null
            .tnyStatus_32 = 0
            .numLocationID = Null
            
            arrInput = Array(.intVoucherID_1, .intLocalBodyID_2, .intTransactionID_3, .intTransactionTypeID_4, _
            .tnyVoucherTypeID_5, .intVoucherNo_6, .intBookNo_7, .dtDate_8, _
            .fltAmount_9, .intInstrumentTypeID_10, .vchInstrumentNo_11, .dtInstrumentDate_12, _
            .vchDescription_13, .numZoneID_14, .numWardID_15, .intDoorNoP1_16, _
            .vchDoorNoP2_17, .vchDoorNoP3_18, .intUserID_19, .intCounterID_20, _
            .numSubLedgerID_21, .intKeyID1_22, .intKeyID2_23, .intExternalApplicationID_24, _
            .intExternalModuleID_25, .intFinancialYearID_26, .tnyShiftID_27, _
            .tnyPrintFlag_28, .tnyCancelFlag_29, .vchBank_33, .vchBankPlace_34, _
            .intFundID_35, .numSeatID, .intSessionID, .vchRefNo, _
            .fltRoundOff, .fltAdvAmtAdj, .numInwardNo, .tnyStatus_32, _
            .numLocationID)
        End With
        
        objAc.SetAccountID mSalaryHeadID
        If objAc.AccountHeadID > 0 Then
            mSalaryHeadCode = objAc.AccountCode   ' Debit Head
        Else
            GoTo ErrNoPenHeadNotFound:
        End If
        
        
        objDb.CreateNewConnection mCn, enuSourceString.Saankhya
        'mCn.BeginTrans
        'On Error GoTo ErrRollBack:
        objDb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCn
        If IsNumeric(arrOutPut(0, 0)) Then
            mintVoucherID = arrOutPut(0, 0)
        Else
            MsgBox "Error : Voucher Table didnt able to save!", vbInformation
            GoTo ErrRollBack:
        End If
        
        'Note:- Gross Salary AccountHead to the Voucher Child
        With mVouChildTbl
            .intVoucherID_1 = mintVoucherID
            .intLocalBodyID_2 = gbLocalBodyID
            .intSlNo_3 = 1
            .intAccountHeadID_4 = gbAcHeadIDGrossSalaryPayable
            .tnyDebitOrCredit_5 = 0
            If IsDate(Rec!dtKeyDate) Then
                .intYearID_6 = Year(Rec!dtKeyDate)
                .tnyPeriodID_7 = Month(Rec!dtKeyDate)
            Else
                .intYearID_6 = Null
                .tnyPeriodID_7 = Null
            End If
            .tnyArrearFlag_8 = Null
            .numDemandID_9 = Rec!intKeyID
            .fltAmount_10 = RecChild!numAmount
        
            arrInput = Array(.intVoucherID_1, _
            .intLocalBodyID_2, _
            .intSlNo_3, _
            .intAccountHeadID_4, _
            .tnyDebitOrCredit_5, _
            .intYearID_6, _
            .tnyPeriodID_7, _
            .tnyArrearFlag_8, _
            .numDemandID_9, _
            .fltAmount_10)
            objDb.ExecuteSP "spSaveVoucherChild", arrInput, , , mCn
        End With
        
        
        With mTranTable
            .intTransactionID = -1
            .intLocalBodyID = gbLocalBodyID
            '.intFinancialYearID = gbFinancialYearID
            .dtTransactionDate = Rec!dtDueDate
            'If .dtTransactionDate < DateAdd("yyyy", -1, gbStartingDate) Or .dtTransactionDate > DateAdd("yyyy", -1, gbEndingDate) Then
            If mPreviousYearMode Then
                .intFinancialYearID = gbFinancialYearID - 1
            Else
                .intFinancialYearID = gbFinancialYearID
            End If
            .intExternalApplicationID = Null
            .intExternalApplicationModuleID = Null
            .intFunctionID = IIf(Rec!intFunctionID = 0, Null, Rec!intFunctionID)
            .intFunctionaryID = IIf(Rec!intFunctionaryID = 0, Null, Rec!intFunctionaryID)
            .intFieldID = Null
            .intFundID = gbFundID
            .intBudgetCentreID = Null
            .vchNarration = Rec!vchDescription
            .intTransactionTypeID = Rec!intTransactionTypeID
            .intProcessID = Null
            .vchGroup = "JV"
            .intGroupID = 40
            .intKeyID = Null
            .numSubLedgerID = Null
            .numUserID = gbUserID
            .intVoucherID = mintVoucherID
            
            arrInput = Array(.intTransactionID, _
            .intLocalBodyID, _
            .intFinancialYearID, _
            .dtTransactionDate, _
            .intExternalApplicationID, _
            .intExternalApplicationModuleID, _
            .intFunctionID, _
            .intFunctionaryID, _
            .intFieldID, _
            .intFundID, _
            .intBudgetCentreID, _
            .vchNarration, _
            .intTransactionTypeID, _
            .intProcessID, _
            .vchGroup, _
            .intGroupID, _
            .intKeyID, _
            .numSubLedgerID, _
            .numUserID, _
            .intVoucherID)
        
        End With
        
        objDb.ExecuteSP "spSaveTransactions", arrInput, arrOutPut, , mCn
        If IsNumeric(arrOutPut(0, 0)) Then
            mintTransactionID = arrOutPut(0, 0)
        End If
        
        With mTranChildTbl
            .intTransactionID = mintTransactionID
            .intSerialNo = 1
            .intAccountHeadID = mSalaryHeadID 'Rec!intCashOrBankHeadID
            .fltAmount = RecChild!numAmount
            .tinDebitOrCreditFlag = 1
            .intByAccountHeadID = Null
            .vchNarration = RecChild!vchDescription
            .intFundID = gbFundID
            
            arrInput = Array(.intTransactionID, _
            .intSerialNo, _
            .intAccountHeadID, _
            .fltAmount, _
            .tinDebitOrCreditFlag, _
            .intByAccountHeadID, _
            .vchNarration, _
            .intFundID)
            
            objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCn
            
            
            .intSerialNo = 2
            .intAccountHeadID = gbAcHeadIDGrossSalaryPayable
            .tinDebitOrCreditFlag = 0
            .intByAccountHeadID = Rec!intCashOrBankHeadID
            
            arrInput = Array(.intTransactionID, _
            .intSerialNo, _
            .intAccountHeadID, _
            .fltAmount, _
            .tinDebitOrCreditFlag, _
            .intByAccountHeadID, _
            .vchNarration, _
            .intFundID)
            
            objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCn
        End With
        
        
TryDeductions:

        '----------------------------------------------------------------- '
        ' Second Journal Voucher                                           '
        ' Gross Salary A/c Dr to                                           '
        '   Deductions          Cr                                         '
        '   Net Salary Payable  Cr                                         '
        '----------------------------------------------------------------- '
        ' Debit                                                            '
        ' 350110101        Gross Salary Payable                            '
        '                                                                  '
        ' Credit                                                           '
        '     (---      Deduction Heads  --- )                             '
        '     350110102    Net Salary Payable                              '
        '----------------------------------------------------------------- '
        RecChild.MoveFirst
        While Not RecChild.EOF
            If RecChild!tnyCategoryFlag = 3 Then
                mNetSalaryAmt = RecChild!numAmount
                GoTo Deductions
            End If
            RecChild.MoveNext
        Wend
        GoTo TryPension:
Deductions:
        
        With mVoucher
            .intVoucherID_1 = -1
            .intLocalBodyID_2 = gbLocalBodyID
            .intTransactionID_3 = Null
            .intTransactionTypeID_4 = Rec!intTransactionTypeID
            .tnyVoucherTypeID_5 = 40
            .intVoucherNo_6 = Null
            .intBookNo_7 = Null
            .dtDate_8 = Rec!dtDueDate
            .fltAmount_9 = mGrossSalaryAmt
            .intInstrumentTypeID_10 = Null
            .vchInstrumentNo_11 = Null
            .dtInstrumentDate_12 = Null
            .vchDescription_13 = Rec!vchDescription
            .numZoneID_14 = Null
            .numWardID_15 = Null
            .intDoorNoP1_16 = Null
            .vchDoorNoP2_17 = Null
            .vchDoorNoP3_18 = Null
            .intUserID_19 = gbUserID
            .intCounterID_20 = gbCounterID
            .numSubLedgerID_21 = Null
            .intKeyID1_22 = gbAcHeadIDGrossSalaryPayable  'Credit to GrossSalary Payable
            .intKeyID2_23 = mPaymentOrderNo
            .intExternalApplicationID_24 = 115
            .intExternalModuleID_25 = 61 'PaymentOrder-SthapanaInterface Module
            '.intFinancialYearID_26 = gbFinancialYearID
            'If .dtDate_8 < DateAdd("yyyy", -1, gbStartingDate) Or .dtDate_8 > DateAdd("yyyy", -1, gbEndingDate) Then
            If mPreviousYearMode Then
                .intFinancialYearID_26 = gbFinancialYearID - 1
            Else
                .intFinancialYearID_26 = gbFinancialYearID
            End If
            .tnyShiftID_27 = Null
            .tnyPrintFlag_28 = Null
            .tnyCancelFlag_29 = Null
            .vchBank_33 = Null
            .vchBankPlace_34 = Null
            .intFundID_35 = Null
            .numSeatID = gbSeatID
            .intSessionID = gbSessionID
            .vchRefNo = Null
            .fltRoundOff = Null
            .fltAdvAmtAdj = Null
            .numInwardNo = Null
            .tnyStatus_32 = 0
            .numLocationID = Null
            
            arrInput = Array(.intVoucherID_1, .intLocalBodyID_2, .intTransactionID_3, .intTransactionTypeID_4, _
            .tnyVoucherTypeID_5, .intVoucherNo_6, .intBookNo_7, .dtDate_8, _
            .fltAmount_9, .intInstrumentTypeID_10, .vchInstrumentNo_11, .dtInstrumentDate_12, _
            .vchDescription_13, .numZoneID_14, .numWardID_15, .intDoorNoP1_16, _
            .vchDoorNoP2_17, .vchDoorNoP3_18, .intUserID_19, .intCounterID_20, _
            .numSubLedgerID_21, .intKeyID1_22, .intKeyID2_23, .intExternalApplicationID_24, _
            .intExternalModuleID_25, .intFinancialYearID_26, .tnyShiftID_27, _
            .tnyPrintFlag_28, .tnyCancelFlag_29, .vchBank_33, .vchBankPlace_34, _
            .intFundID_35, .numSeatID, .intSessionID, .vchRefNo, _
            .fltRoundOff, .fltAdvAmtAdj, .numInwardNo, .tnyStatus_32, _
            .numLocationID)
        End With
        objDb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCn
        If IsNumeric(arrOutPut(0, 0)) Then
            mintVoucherID = arrOutPut(0, 0)
        End If
        
        
        With mTranTable
            .intTransactionID = -1
            .intLocalBodyID = gbLocalBodyID
            '.intFinancialYearID = gbFinancialYearID
            .dtTransactionDate = Rec!dtDueDate
            'If .dtTransactionDate < DateAdd("yyyy", -1, gbStartingDate) Or .dtTransactionDate > DateAdd("yyyy", -1, gbEndingDate) Then
            If mPreviousYearMode Then
                .intFinancialYearID = gbFinancialYearID - 1
            Else
                .intFinancialYearID = gbFinancialYearID
            End If
            .intExternalApplicationID = Null
            .intExternalApplicationModuleID = Null
            .intFunctionID = IIf(Rec!intFunctionID = 0, Null, Rec!intFunctionID)
            .intFunctionaryID = IIf(Rec!intFunctionaryID = 0, Null, Rec!intFunctionaryID)
            .intFieldID = Null
            .intFundID = gbFundID
            .intBudgetCentreID = Null
            .vchNarration = Rec!vchDescription
            .intTransactionTypeID = Rec!intTransactionTypeID
            .intProcessID = Null
            .vchGroup = "JV"
            .intGroupID = 40
            .intKeyID = Null
            .numSubLedgerID = Null
            .numUserID = gbUserID
            .intVoucherID = mintVoucherID
            
            arrInput = Array(.intTransactionID, _
            .intLocalBodyID, _
            .intFinancialYearID, _
            .dtTransactionDate, _
            .intExternalApplicationID, _
            .intExternalApplicationModuleID, _
            .intFunctionID, _
            .intFunctionaryID, _
            .intFieldID, _
            .intFundID, _
            .intBudgetCentreID, _
            .vchNarration, _
            .intTransactionTypeID, _
            .intProcessID, _
            .vchGroup, _
            .intGroupID, _
            .intKeyID, _
            .numSubLedgerID, _
            .numUserID, _
            .intVoucherID)
        
        End With
        
        objDb.ExecuteSP "spSaveTransactions", arrInput, arrOutPut, , mCn
        If IsNumeric(arrOutPut(0, 0)) Then
            mintTransactionID = arrOutPut(0, 0)
        End If
        
        'Note:-Gross Salary Payable A/c Debtor
        With mTranChildTbl
            .intTransactionID = mintTransactionID
            .intSerialNo = 1
            .intAccountHeadID = gbAcHeadIDGrossSalaryPayable
            .fltAmount = mGrossSalaryAmt
            .tinDebitOrCreditFlag = 1
            .intByAccountHeadID = Null
            .vchNarration = RecChild!vchDescription
            .intFundID = gbFundID
            
            arrInput = Array(.intTransactionID, _
            .intSerialNo, _
            .intAccountHeadID, _
            .fltAmount, _
            .tinDebitOrCreditFlag, _
            .intByAccountHeadID, _
            .vchNarration, _
            .intFundID)
            
            objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCn
        End With
        
        
        mSLNo = 0
        RecChild.MoveFirst
        While Not RecChild.EOF
            If RecChild!tnyCategoryFlag = 2 Or RecChild!tnyCategoryFlag = 3 Then
                        mSLNo = mSLNo + 1
                        'Note:- Deduction Heads and Net Salary to Voucher Child
                        With mVouChildTbl
                            .intVoucherID_1 = mintVoucherID
                            .intLocalBodyID_2 = gbLocalBodyID
                            .intSlNo_3 = mSLNo
                            .intAccountHeadID_4 = RecChild!intAccountHeadID
                            .tnyDebitOrCredit_5 = 0
                            If IsDate(Rec!dtKeyDate) Then
                                .intYearID_6 = Year(Rec!dtKeyDate)
                                .tnyPeriodID_7 = Month(Rec!dtKeyDate)
                            Else
                                .intYearID_6 = Null
                                .tnyPeriodID_7 = Null
                            End If
                            .tnyArrearFlag_8 = Null
                            .numDemandID_9 = Rec!intKeyID
                            .fltAmount_10 = RecChild!numAmount
                        
                            arrInput = Array(.intVoucherID_1, _
                            .intLocalBodyID_2, _
                            .intSlNo_3, _
                            .intAccountHeadID_4, _
                            .tnyDebitOrCredit_5, _
                            .intYearID_6, _
                            .tnyPeriodID_7, _
                            .tnyArrearFlag_8, _
                            .numDemandID_9, _
                            .fltAmount_10)
                            objDb.ExecuteSP "spSaveVoucherChild", arrInput, , , mCn
                        End With
                                
                        With mTranChildTbl
                            .intTransactionID = mintTransactionID
                            .intSerialNo = mSLNo + 1
                            .intAccountHeadID = RecChild!intAccountHeadID
                            .fltAmount = RecChild!numAmount
                            .tinDebitOrCreditFlag = 0
                            .intByAccountHeadID = gbAcHeadIDGrossSalaryPayable
                            .vchNarration = RecChild!vchDescription
                            .intFundID = gbFundID
                            
                            arrInput = Array(.intTransactionID, _
                            .intSerialNo, _
                            .intAccountHeadID, _
                            .fltAmount, _
                            .tinDebitOrCreditFlag, _
                            .intByAccountHeadID, _
                            .vchNarration, _
                            .intFundID)
                            
                            objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCn
                        End With
            End If
            RecChild.MoveNext
        Wend
        
        
        RecChild.MoveFirst
        While Not RecChild.EOF
            If RecChild!tnyCategoryFlag = 3 Then
                GoTo TryPension:
            End If
            RecChild.MoveNext
        Wend
        GoTo CleanUp:
TryPension:
        '----------------------------------------------------------------- '
        ' Third Journal Voucher                                           '
        ' Pention Contribution  A/c Dr to                                           '
        '   Pention Payable         Cr
        '----------------------------------------------------------------- '
        ' Permenent Staff       Dr - 210100102    Cr - 210300102
        ' Secretary             Dr - 210100101    Cr - 210300101
        ' Contingent-Full Time  Dr - 210100104    Cr - 210300103
        ' Contingent-Part Time  Dr - 210100105    Cr - 210300104
        '----------------------------------------------------------------- '
        ' 210100101        Salaries -Secretary                             '
        ' 210100102        Salaries - Permanent Staff                      '
        ' 210100103        Salaries - Temporary Staff-Recruited through
        '                            Employment Exchange                   '
        ' 210100104        Salaries - Full Time Contingent Staff           '
        ' 210100105        Salaries - Part Time Contingent Staff           '
        ' 210100106        Salaries - Contract Staff                       '
        ' 210100107        Salaries - Honorarium Staff                     '
        
      
        
        
        Dim mPenContributionHeadCode As String
        Dim mPenPayableHeadCode      As String
        
        
        If mPensionCategory = 5 Then
            mPensionAmt = mPensionAmount
        Else
            mPensionAmt = Format(mGrossSalaryAmt * 15 / 100, "0#")
        End If
        
       ' mPensionAmt = Format(mGrossSalaryAmt * 15 / 100, "0#")
        
        objAc.SetAccountID mSalaryHeadID
        If objAc.AccountHeadID > 0 Then
            mSalaryHeadCode = objAc.AccountCode   ' Debit Head
        Else
            GoTo ErrNoPenHeadNotFound:
        End If
        
        If mSalaryHeadCode = "210100101" Then ' Secretary
            mPenContributionHeadCode = "210300101"
            mPenPayableHeadCode = "350110104"
        ElseIf mSalaryHeadCode = "210100104" Then ' Contingent Staff-Full time
            mPenContributionHeadCode = "210300103"
            mPenPayableHeadCode = "350110104"
        ElseIf mSalaryHeadCode = "210100105" Then ' Contingent Staff-Part time
            mPenContributionHeadCode = "210300104"
            mPenPayableHeadCode = "350110104"
        ElseIf mSalaryHeadCode = "210100102" Then    ' Permenent Staff
            mPenContributionHeadCode = "210300102"
            mPenPayableHeadCode = "350110104"
        End If
        
        With mVoucher
            .intVoucherID_1 = -1
            .intLocalBodyID_2 = gbLocalBodyID
            .intTransactionID_3 = Null
            .intTransactionTypeID_4 = Rec!intTransactionTypeID
            .tnyVoucherTypeID_5 = 40
            .intVoucherNo_6 = Null
            .intBookNo_7 = Null
            .dtDate_8 = Rec!dtDueDate
            .fltAmount_9 = mPensionAmt
             mGrossSalaryAmt = RecChild!numAmount
            .intInstrumentTypeID_10 = Null
            .vchInstrumentNo_11 = Null
            .dtInstrumentDate_12 = Null
           
            '.vchDescription_13 = "Being the Pension Contribution for the month of " & Format(Month(Rec!dtKeyDate), "mmm") & "," & Year(Rec!dtKeyDate)
            .vchDescription_13 = "Being the Pension Contribution for the month of " & MonthName(Month(Rec!dtKeyDate), True) & "," & Year(Rec!dtKeyDate)
            .numZoneID_14 = Null
            .numWardID_15 = Null
            .intDoorNoP1_16 = Null
            .vchDoorNoP2_17 = Null
            .vchDoorNoP3_18 = Null
            .intUserID_19 = gbUserID
            .intCounterID_20 = gbCounterID
            .numSubLedgerID_21 = Null
            objAc.SetAccountCode mPenContributionHeadCode
            If objAc.AccountHeadID > 0 Then
                .intKeyID1_22 = objAc.AccountHeadID  ' Debit Head
            Else
                GoTo ErrNoPenHeadNotFound:
            End If
            .intKeyID2_23 = mPaymentOrderNo
            .intExternalApplicationID_24 = 115
            .intExternalModuleID_25 = 61 'PaymentOrder-SthapanaInterface Module
            '.intFinancialYearID_26 = gbFinancialYearID
            'If .dtDate_8 < DateAdd("yyyy", -1, gbStartingDate) Or .dtDate_8 > DateAdd("yyyy", -1, gbEndingDate) Then
            If mPreviousYearMode Then
                .intFinancialYearID_26 = gbFinancialYearID - 1
            Else
                .intFinancialYearID_26 = gbFinancialYearID
            End If
            .tnyShiftID_27 = Null
            .tnyPrintFlag_28 = Null
            .tnyCancelFlag_29 = Null
            .vchBank_33 = Null
            .vchBankPlace_34 = Null
            .intFundID_35 = Null
            .numSeatID = gbSeatID
            .intSessionID = gbSessionID
            .vchRefNo = Null
            .fltRoundOff = Null
            .fltAdvAmtAdj = Null
            .numInwardNo = Null
            .tnyStatus_32 = 0
            .numLocationID = Null
            
            arrInput = Array(.intVoucherID_1, .intLocalBodyID_2, .intTransactionID_3, .intTransactionTypeID_4, _
            .tnyVoucherTypeID_5, .intVoucherNo_6, .intBookNo_7, .dtDate_8, _
            .fltAmount_9, .intInstrumentTypeID_10, .vchInstrumentNo_11, .dtInstrumentDate_12, _
            .vchDescription_13, .numZoneID_14, .numWardID_15, .intDoorNoP1_16, _
            .vchDoorNoP2_17, .vchDoorNoP3_18, .intUserID_19, .intCounterID_20, _
            .numSubLedgerID_21, .intKeyID1_22, .intKeyID2_23, .intExternalApplicationID_24, _
            .intExternalModuleID_25, .intFinancialYearID_26, .tnyShiftID_27, _
            .tnyPrintFlag_28, .tnyCancelFlag_29, .vchBank_33, .vchBankPlace_34, _
            .intFundID_35, .numSeatID, .intSessionID, .vchRefNo, _
            .fltRoundOff, .fltAdvAmtAdj, .numInwardNo, .tnyStatus_32, _
            .numLocationID)
        End With
        objDb.ExecuteSP "spSaveVoucher", arrInput, arrOutPut, , mCn
        If IsNumeric(arrOutPut(0, 0)) Then
            mintVoucherID = arrOutPut(0, 0)
        Else
            MsgBox "Error : Voucher Table didnt able to save!", vbInformation
            Exit Sub
        End If
        
        'Note:- Pension Payable AccountHead to the Voucher Child
        With mVouChildTbl
            .intVoucherID_1 = mintVoucherID
            .intLocalBodyID_2 = gbLocalBodyID
            .intSlNo_3 = 1
            objAc.SetAccountCode mPenPayableHeadCode
            If objAc.AccountHeadID > 0 Then
                .intAccountHeadID_4 = objAc.AccountHeadID  ' Credit Head
            Else
                GoTo ErrNoPenHeadNotFound:
            End If
            .tnyDebitOrCredit_5 = 0
            .intYearID_6 = Year(Rec!dtKeyDate)
            .tnyPeriodID_7 = Month(Rec!dtKeyDate)
            .tnyArrearFlag_8 = Null
            .numDemandID_9 = Rec!intKeyID
            .fltAmount_10 = mPensionAmt
        
            arrInput = Array(.intVoucherID_1, _
            .intLocalBodyID_2, _
            .intSlNo_3, _
            .intAccountHeadID_4, _
            .tnyDebitOrCredit_5, _
            .intYearID_6, _
            .tnyPeriodID_7, _
            .tnyArrearFlag_8, _
            .numDemandID_9, _
            .fltAmount_10)
            objDb.ExecuteSP "spSaveVoucherChild", arrInput, , , mCn
        End With
        
        
        With mTranTable
            .intTransactionID = -1
            .intLocalBodyID = gbLocalBodyID
            '.intFinancialYearID = gbFinancialYearID
            .dtTransactionDate = Rec!dtDueDate
            'If .dtTransactionDate < DateAdd("yyyy", -1, gbStartingDate) Or .dtTransactionDate > DateAdd("yyyy", -1, gbEndingDate) Then
            If mPreviousYearMode Then
                .intFinancialYearID = gbFinancialYearID - 1
            Else
                .intFinancialYearID = gbFinancialYearID
            End If
            .intExternalApplicationID = Null
            .intExternalApplicationModuleID = Null
            .intFunctionID = IIf(Rec!intFunctionID = 0, Null, Rec!intFunctionID)
            .intFunctionaryID = IIf(Rec!intFunctionaryID = 0, Null, Rec!intFunctionaryID)
            .intFieldID = Null
            .intFundID = gbFundID
            .intBudgetCentreID = Null
            '.vchNarration = "Being the Pension Contribution for the month of " & Format(Month(Rec!dtKeyDate), "mmmm") & "," & Year(Rec!dtKeyDate)
            .vchNarration = "Being the Pension Contribution for the month of " & MonthName(Month(Rec!dtKeyDate), True) & "," & Year(Rec!dtKeyDate)
            .intTransactionTypeID = Rec!intTransactionTypeID
            .intProcessID = Null
            .vchGroup = "JV"
            .intGroupID = 40
            .intKeyID = Null
            .numSubLedgerID = Null
            .numUserID = gbUserID
            .intVoucherID = mintVoucherID
            
            arrInput = Array(.intTransactionID, _
            .intLocalBodyID, _
            .intFinancialYearID, _
            .dtTransactionDate, _
            .intExternalApplicationID, _
            .intExternalApplicationModuleID, _
            .intFunctionID, _
            .intFunctionaryID, _
            .intFieldID, _
            .intFundID, _
            .intBudgetCentreID, _
            .vchNarration, _
            .intTransactionTypeID, _
            .intProcessID, _
            .vchGroup, _
            .intGroupID, _
            .intKeyID, _
            .numSubLedgerID, _
            .numUserID, _
            .intVoucherID)
        
        End With
        
        objDb.ExecuteSP "spSaveTransactions", arrInput, arrOutPut, , mCn
        If IsNumeric(arrOutPut(0, 0)) Then
            mintTransactionID = arrOutPut(0, 0)
        End If
        
        With mTranChildTbl
            .intTransactionID = mintTransactionID
            .intSerialNo = 1
            
            objAc.SetAccountCode mPenContributionHeadCode
            If objAc.AccountHeadID > 0 Then
                .intAccountHeadID = objAc.AccountHeadID  ' Debit Head
            Else
                GoTo ErrNoPenHeadNotFound:
            End If
            .fltAmount = mPensionAmt
            .tinDebitOrCreditFlag = 1
            .intByAccountHeadID = Null
            .vchNarration = RecChild!vchDescription
            .intFundID = gbFundID
            
            arrInput = Array(.intTransactionID, _
            .intSerialNo, _
            .intAccountHeadID, _
            .fltAmount, _
            .tinDebitOrCreditFlag, _
            .intByAccountHeadID, _
            .vchNarration, _
            .intFundID)
            
            objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCn
            
            .intSerialNo = 2
            objAc.SetAccountCode mPenPayableHeadCode
            If objAc.AccountHeadID > 0 Then
                .intAccountHeadID = objAc.AccountHeadID  ' Credit Head
            Else
                GoTo ErrNoPenHeadNotFound:
            End If
            .tinDebitOrCreditFlag = 0
            
            objAc.SetAccountCode mPenContributionHeadCode
            If objAc.AccountHeadID > 0 Then
                .intByAccountHeadID = objAc.AccountHeadID   ' Debit Head
            Else
                GoTo ErrNoPenHeadNotFound:
            End If
            
            
            arrInput = Array(.intTransactionID, _
            .intSerialNo, _
            .intAccountHeadID, _
            .fltAmount, _
            .tinDebitOrCreditFlag, _
            .intByAccountHeadID, _
            .vchNarration, _
            .intFundID)
            
            objDb.ExecuteSP "spSaveTransactionChild", arrInput, , , mCn
        End With
        'Note:- Changing the Status of PayOrder as Approved!
        mSql = "Update faPayOrder Set tnyStatus = 1 Where vchPayOrderNo = " & mPaymentOrderNo
        objDb.ExecuteSP mSql, , , , mCn, adCmdText
        'mCn.CommitTrans
CleanUp:
        Exit Sub
        
ErrNoPenHeadNotFound:
        MsgBox "Pension Head not found", vbInformation
        GoTo ErrRollBack:
ErrNoGr:
    MsgBox "No Gross Amount Found to Proceed!", vbInformation
    GoTo ErrRollBack:

ErrRollBack:
    'mCn.RollbackTrans
    Set mCn = Nothing
End Sub




