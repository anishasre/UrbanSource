VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReverseProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Option Explicit
    ' Variable Declarations '
    Private mvarRequestID                           As Integer
    Private mvarRequestDate                         As Date
    Private mvarCategoryID                          As Integer
    Private mvarVoucherTypeID                       As Integer
    Private mvarReasonID                            As Integer
    Private mVarReason                              As String
    Private mvarRemarks                             As String
    Private mvarRequestedUserID                     As Double
    Private mvarRequestedSeatID                     As Double
    Private mvarAuthorisedByAO                      As Double
    Private mvarAuthorisationDateAO                 As Date
    Private mvarForwardedSeatID                     As Double
    Private mvarFinancialYearID                     As Integer
    Private mvarStatus                              As Integer
    Private mvarAuthorisedBySec                     As Double
    Private mvarAuthorisationDateSec                As Date
    Private mvarDemandNo                            As Double
    Private mvarDemandID                            As Double
    Private mvarPaymentVoucherID                    As Integer
    Private mvarPaid                                As Integer
    Private mvarRecoveryRemitted                    As Integer
    Private mvarChequeCancelled                     As Integer
    Private mvarInstrumentNo                        As String
    Private mvarInstrumentDate                      As Date
    
    Private intVoucherID                            As Variant
    Private intVoucherNo                            As Variant
    Private intTransactionTypeID                    As Variant
    Private intInstrumentTypeID                     As Variant
    Private mMode                                   As Boolean 'false -Double transaction true-single Transaction
    Private mTransactionDate                        As Variant
    
    Public Property Let VoucherID(ByVal mData As Variant)
        intVoucherID = mData
    End Property
    Public Property Get VoucherID() As Variant
        VoucherID = intVoucherID
    End Property
    Public Property Let VoucherNo(ByVal mData As Variant)
        intVoucherNo = mData
    End Property
    Public Property Get VoucherNo() As Variant
        VoucherNo = intVoucherNo
    End Property
    Public Property Let TransactionTypeID(ByVal mData As Variant)
        intTransactionTypeID = mData
    End Property
    Public Property Get TransactionTypeID() As Variant
        TransactionTypeID = intTransactionTypeID
    End Property

    Public Function ReverseProcess(mRequestID As Long, mReasonID As Integer) As Variant
    'mRequestID is request id of the thansactions to be reversed
    'mReasonID :what the reason for reversing
    '*****************
    'This Function has a Return vlaue ie Voucherid of Reversed Transactions
    '*****************
         On Error GoTo err:
            Dim objdb       As New clsDB
            Dim mSql        As String
            Dim Rec         As New ADODB.Recordset
            Dim mCnn        As New ADODB.Connection
            Dim fltAmount   As Variant
            Dim dtDate      As Date
            Dim vchUser     As Variant
            Dim mVoucherType As Integer
            Dim intReasonID     As Integer
            Dim aryOut      As Variant
            
            
            If objdb.SetConnection(mCnn) Then
                mSql = "Select * From faVouchers Inner Join faReverseEntryChild On faVouchers.intVoucherID=faReverseEntryChild.intVoucherID Where intRequestID=" & mRequestID
                Rec.Open mSql, mCnn
                aryOut = ""
                While Not (Rec.EOF Or Rec.BOF)
                        intVoucherID = IIf(IsNull(Rec!intVoucherID), "", Rec!intVoucherID)
                        intVoucherNo = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                        intTransactionTypeID = IIf(IsNull(Rec!intTransactionTypeID), "", Rec!intTransactionTypeID)
                        intInstrumentTypeID = IIf(IsNull(Rec!intInstrumentTypeID), "", Rec!intInstrumentTypeID)
                        mVoucherType = IIf(IsNull(Rec!tnyVoucherTypeID), "", Rec!tnyVoucherTypeID)
                        If mVoucherType = 10 Then
                            Select Case intInstrumentTypeID
'                                Case gbInstrumentCash           'Cash
'                                    '----------------------------------------------------'
'                                    If mReasonID = 3 Then       'Amount
'                                        aryOut = Null
'                                        ' DoubleTransaction
'                                    ElseIf mReasonID = 4 Then       'Particulars
'                                        aryOut = Null
'                                        ' DoubleTransaction
'                                    ElseIf mReasonID = 5 Then       'Account Head
'                                        aryOut = Null
'                                    Else
'                                        aryOut = Null
'                                    End If
                                Case gbInstrumentCheque         'Cheque
                                    If mReasonID = 500 Then       'Cheque Dishonour
                                        mMode = True
                                        aryOut = CStr(aryOut) + CStr(SingleTransaction) + "-"
                                    ElseIf mReasonID = 2 Then   'Cheque Duplication
                                        mMode = True
                                        aryOut = CStr(aryOut) + CStr(SingleTransaction) + "-"
                                    ElseIf mReasonID = 3 Then   'Amount
                                         aryOut = Null
                                    ElseIf mReasonID = 4 Then   'Particulars
                                        aryOut = Null
                                        ' DoubleTransaction
                                    ElseIf mReasonID = 5 Then   'Account Head
                                        aryOut = Null
                                    End If
                                Case Else
                                        aryOut = Null
                            End Select
                        ElseIf mVoucherType = 30 Then
                            mMode = True
                            aryOut = CStr(aryOut) + CStr(SingleTransaction) + "-"
                        ElseIf mVoucherType = 40 Then
                            mMode = True
                            aryOut = CStr(aryOut) + CStr(SingleTransaction) + "-"
                        Else
                            aryOut = Null
                        End If
                    Rec.MoveNext
                Wend
            End If
            
            If Not IsNull(aryOut) Then
               ' mSQL = "Update faReverseEntry Set tnyStatus=2,numApprovedUserID=" & gbUserID & " ,dtApprovedDate=" & gbTransactionDate & " Where intRequestID=" & mRequestID
                mSql = "Update faReverseEntry Set tnyStatus=2,numAuthorisedByAO=" & gbUserID & " ,dtAuthorisationDateAO=" & gbTransactionDate & " Where intRequestID=" & mRequestID
                objdb.ExecuteSP mSql, , , , mCnn, adCmdText
                ReverseProcess = Split(aryOut, "-")
            Else
                ReverseProcess = Null
            End If
            
        Exit Function
err:
        MsgBox (Error$)
    End Function
    
    Private Function SingleTransaction() As Variant
            Dim objdb       As New clsDB
            Dim Rec         As New ADODB.Recordset
            Dim mCnn        As New ADODB.Connection
            Dim arrIn       As Variant
            Dim arrOut      As Variant
            Dim mSql        As String
            Dim mReverseID   As Variant
            Dim mReverseNo   As Variant
            '--------This will Reverse Requested Transaction.Results only Single Transaction
            'ie if Requested Voucher is Receipt Then Result is Payment voucher and vise versa
            'if Requested is Conta then Contra ,if it JV then JV
            
            arrIn = Array(intVoucherID, gbTransactionDate)
            objdb.CreateNewConnection mCnn, enuSourceString.Saankhya
            objdb.ExecuteSP "spSaveReverseVouchers", arrIn, arrOut, , mCnn, adCmdStoredProc
            If IsArray(arrOut) Then
                mReverseID = arrOut(0, 0)
                mReverseNo = arrOut(1, 0)
            End If
            '----If mode is true Demand for proprerty tax will be open-------------------
            Rec.Open "Select * From favouchers where intVoucherId=" & intVoucherID, mCnn
            If Rec!intTransactionTypeID = gbTransactionTypePTax And mMode = True Then
                Call PropertyTaxReverse
            End If
            '----------------------------------------------------------------------------
            SingleTransaction = mReverseID
            Exit Function
        End Function
        
        Public Function ReverseTransaction(intVrID, mCn As ADODB.Connection) As Variant
            Dim objdb           As New clsDB
            Dim Rec             As New ADODB.Recordset
            Dim arrIn           As Variant
            Dim arrOut          As Variant
            Dim mSql            As String
            Dim mReverseID      As Variant
            Dim mReverseNo      As Variant
            Dim mTrDate         As Date
            Dim mTrType         As Integer
            '--------This will Reverse Requested Transaction.Results only Single Transaction
            'ie if Requested Voucher is Receipt Then Result is Payment voucher and vise versa
            'if Requested is Conta then Contra ,if it JV then JV
            
            If IsDate(mTransactionDate) Then
                mTrDate = mTransactionDate
            Else
                mTrDate = gbTransactionDate
            End If
            
            arrIn = Array(intVrID, mTrDate)
            objdb.CreateNewConnection mCn, enuSourceString.Saankhya
            objdb.ExecuteSP "spSaveReverseVouchers", arrIn, arrOut, , mCn, adCmdStoredProc
            If IsArray(arrOut) Then
                mReverseID = arrOut(0, 0)
                mReverseNo = arrOut(1, 0)
            Else
                'MsgBox "Unexpected Error Occured in Stored Procedure"
                Exit Function
            End If
            
            '--- Updation in voucher Table
            mCn.Execute "Update faVouchers set tnysync=Null,intExternalModuleID=55,tnyReversed=1,dtValueDate='" & Format(mTrDate, "dd-MMM-yyyy") & "' Where intVoucherID=" & mReverseID
            mCn.Execute "Update faVouchers set tnysync=Null,tnyReversed=1 ,dtValueDate='" & Format(mTrDate, "dd-MMM-yyyy") & "' Where intVoucherID=" & intVrID
            
            '--- Updation in faTransaction Table
            mCn.Execute "Update faTransactions set tnysync=Null,intExternalApplicationModuleID=55,tnyReversed=1,dtValueDate='" & Format(mTrDate, "dd-MMM-yyyy") & "' Where intVoucherID=" & mReverseID
            mCn.Execute "Update faTransactions set tnysync=Null,tnyReversed=1,dtValueDate='" & Format(mTrDate, "dd-MMM-yyyy") & "' Where intVoucherID=" & intVrID
            
            '---- Demand for proprerty tax will be set as open-------------------
'            Rec.Open "Select * From favouchers where intVoucherId=" & intVrID, mCn
'            If Not (Rec.EOF And Rec.BOF) Then
'                mTrType = IIf(IsNull(Rec!intTransactionTypeID), 0, Rec!intTransactionTypeID)
'            End If
           ' mCn.Close
            mTrType = TransactionTypeID
            If mTrType = gbTransactionTypePTax Then '--And gbLinkWithPropertyTax Then
                If gbLinkWithPropertyTax Then
                    Call PropertyTaxReverse
                ElseIf gbFetchDemandFromWeb Then  '' added on 2019 Dec
                    If mTrType = 9 Or mTrType = 10 Or mTrType = 11 Then
                        '' skip Sanchaya web service
                    Else
                        If gbLocalBodyID <> 169 Then
                            PTaxWebDemand (val(intVrID))
                        End If
                    End If
                End If
            ElseIf (mTrType = gbTransactionTypeRentOnBuilding Or mTrType = gbTransactionTypeRentOnLand) _
            And gbLinkWithRentOnLand Then
                Call RLBReverse
            End If
           '----------------------------------------------------------------------------
            ReverseTransaction = mReverseID
        End Function
        
             Private Function PTaxWebDemand(mVoucherID As Long)
                    Dim Rec             As New ADODB.Recordset
                    Dim mCollPost       As String
                    Dim mColZoneID      As String
                    Dim mBuildingIdWeb  As String
                    Dim mColAmt            As String
                    Dim mColDate        As String
                    Dim mColReceiptNo   As String
                    Dim mColBookNo      As String
                    Dim mColPeriodId     As String
                    Dim mColYearID       As String
                    Dim mHash           As String
                    Dim mCollOut        As String
'                    Dim node            As IXMLDOMNode
'                    Dim DataNodes       As IXMLDOMNodeList
                    Dim mUrl            As String
                    Dim objSOAP         As Variant
                    Dim mLen            As Integer
                    Dim mColAccID       As String
                    Dim mColKeyID       As String
                    Dim objdb As New clsDB

                mUrl = gbDefaultUrlSanchayaPost
                Set objSOAP = CreateObject("MSSOAP.SOAPClient30")
                objSOAP.MSSoapInit mUrl + "?WSDL"
           
    
                    Set Rec = GetRecordSet("spGetVoucherDetails " & mVoucherID & ", " & gbLocalBodyID, adOpenKeyset, adLockOptimistic)
                    If Not (Rec.EOF And Rec.BOF) Then
                        While Not Rec.EOF
                            
                            mColAmt = IIf(IsNull(Rec!fltAmount), "", Rec!fltAmount)
                            mColDate = IIf(IsNull(Rec!dtDate), "", Rec!dtDate)
                            mColReceiptNo = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                            mColBookNo = IIf(IsNull(Rec!intBookNo), "", Rec!intBookNo)
                            mColPeriodId = IIf(IsNull(Rec!tnyPeriodID), "", Rec!tnyPeriodID)
                            mColYearID = IIf(IsNull(Rec!intYearID), "", Rec!intYearID)
                            mBuildingIdWeb = IIf(IsNull(Rec!numSubLedgerID), "", Rec!numSubLedgerID)
                            mColZoneID = IIf(IsNull(Rec!numZoneID), "", Rec!numZoneID)
                            mColAccID = IIf(IsNull(Rec!intAccountHeadID), "", Rec!intAccountHeadID)
                            mColKeyID = IIf(IsNull(Rec!numDemandID), "", Rec!numDemandID)
                            If mColAccID <> gbAcHeadIDPenalInterest Then
                                mCollPost = mCollPost + CStr(gbLBID) + "#" + CStr(mColZoneID) + "#" + CStr(mBuildingIdWeb) + "#"
                                mCollPost = mCollPost + CStr(mColYearID) + "#" + CStr(mColPeriodId) + "#" + CStr(mVoucherID) + "#"
                                mCollPost = mCollPost + CStr(mColBookNo) + "#" + CStr(mColReceiptNo) + "#" + CStr(mColDate) + "#"
                                mCollPost = mCollPost + CStr(gbFinancialYearID) + "#" + CStr(mColAmt) + "#" + CStr(gbLBName) + "#"
                                mCollPost = mCollPost + CStr(mColAccID) + "#" + CStr(mColKeyID)
                            End If
                            Rec.MoveNext
                            mCollPost = mCollPost + "~"
                        Wend
                        mLen = Len(mCollPost) - 1
                        mCollPost = Left$(mCollPost, mLen - 1)
                        mHash = CStr(mVoucherID) + CStr(mBuildingIdWeb) + "ikm#9567" + CStr(mColDate) + "*ikm#9567"
                        mCollOut = objSOAP.Saankhyaa_CollectionPostingCancel(mCollPost, mHash)
                    End If
  
    End Function
'        Private Sub DoubleTransaction()
'            Dim objDb       As New clsDB
'            Dim Rec         As New ADODB.Recordset
'            Dim mCnn        As New ADODB.Connection
'            Dim ArrIn       As Variant
'            Dim arrOut      As Variant
'            Dim mSql        As String
'            ''-------To Reverse a Voucher  Some Voucher (Depends on Reason)need Two Transactions
'            ''-------ie Requested Voucher is Receipt then its result will be A Payment & Receipt Voucher
'            mMode = False
'            'DoubleTransaction = SingleTransaction
'            'Pendiinding Transaction
'            'capture inputs to save data in temp table and to print details
'            '--
'            '--
'
'        End Sub



        Private Sub PropertyTaxReverse()
            On Error GoTo Err1:
                Dim objdb As New clsDB
                Dim mCnn As New ADODB.Connection
                Dim Rec As New Recordset
                Dim mSql As String
                Dim arrIn As Variant
                Dim mQry As String
                Dim mConfig As Boolean
                Dim mOtherZoneFlag As Boolean
    
                If objdb.SetConnection(mCnn) Then
                    mQry = "Select tnyLinkWithPropertyTax from faConfig"
                    Rec.Open mQry, mCnn
                    If IsNull(Rec!tnyLinkWithPropertyTax) Then
                        mConfig = False
                    ElseIf val(Rec!tnyLinkWithPropertyTax) = 1 Then
                        mConfig = True
                    Else
                        mConfig = False
                    End If
                    If Rec.State = 1 Then Rec.Close
                    mSql = "Select numZoneID,intVoucherNo from faVouchers Where intVoucherID = " & intVoucherID
                    Rec.Open mSql, mCnn
                    If Not (Rec.EOF Or Rec.BOF) Then
                        arrIn = Array(Rec!intVoucherNo)
                        If Rec!numZoneID <> gbLocationID Then
                            mOtherZoneFlag = True
                        Else
                            mOtherZoneFlag = False
                        End If
                    End If
                Else
                    MsgBox "Connection Does not Exist", vbInformation
                End If
    
                If mConfig = True Then
                    Set mCnn = Nothing
                    If objdb.CreateNewConnection(mCnn, enuSourceString.SanchayaLite) Then
                        If mOtherZoneFlag = False Then
                            objdb.ExecuteSP "spReverseDemandFromSaankhya", arrIn, , , mCnn
                        Else
                            '---------------------------------------------------------------'
                            ' Other Zone Office Collection
                            '---------------------------------------------------------------'
                            objdb.ExecuteSP "HOSaanOtherCollectionCancel", arrIn, , , mCnn
                            '----------------------------------------------------------'
                        End If
                    Else
                        MsgBox "Connection Does not Exist", vbInformation
                    End If
                End If
            Exit Sub
Err1:
            MsgBox (Error$)
        End Sub
    Private Sub RLBReverse()
        'Cancel receipt for Rent On Land Integration
        On Error GoTo err:
            Dim mCnn As New ADODB.Connection
            Dim Rec     As New ADODB.Recordset
            Dim aryIn As Variant
            Dim objdb As New clsDB
            Dim mSql    As String
            Dim mVoucherNo  As String
            If objdb.CreateNewConnection(mCnn, enuSourceString.Sanchaya) Then
                mSql = "Select intVoucherNo From faVouchers Where intVoucherID=" & intVoucherID
                Rec.Open mSql, mCnn
                If Not (Rec.EOF And Rec.BOF) Then
                  mVoucherNo = IIf(IsNull(Rec!intVoucherNo), "", Rec!intVoucherNo)
                Else
                
                End If
                aryIn = Array(mVoucherNo, intVoucherID, 2)
                objdb.ExecuteSP "spSanSnRentDemandReverce", aryIn, , , mCnn, adCmdStoredProc
            Else
                MsgBox "Connection To Sanchaya does not exist, Please contact your System Administrator", vbInformation
            End If
        Exit Sub
err:
        MsgBox (Error$)
    End Sub
    Private Sub Class_Initialize()
        intVoucherID = Null
        intVoucherNo = Null
        intTransactionTypeID = Null
        intInstrumentTypeID = Null
    End Sub
    
'------------------------------------------

'------------------------------------------

' Property Definitions '
    Public Property Let RequestID(ByVal vData As Integer)
        mvarRequestID = vData
    End Property
    Public Property Get RequestID() As Integer
        RequestID = mvarRequestID
    End Property
    Public Property Let RequestDate(ByVal vData As Date)
        mvarRequestDate = vData
    End Property
    Public Property Get equestDate() As Date
        RequestDate = mvarRequestDate
    End Property
    Public Property Let CategoryID(ByVal vData As Integer)
        mvarCategoryID = vData
    End Property
    Public Property Get CategoryID() As Integer
        CategoryID = mvarCategoryID
    End Property
    Public Property Let VoucherTypeID(ByVal vData As Integer)
        mvarVoucherTypeID = vData
    End Property
    Public Property Get VoucherTypeID() As Integer
        VoucherTypeID = mvarVoucherTypeID
    End Property
    Public Property Let ReasonID(ByVal vData As Integer)
        mvarReasonID = vData
    End Property
    Public Property Get ReasonID() As Integer
        ReasonID = mvarReasonID
    End Property
    Public Property Let Reason(ByVal vData As Integer)
        mVarReason = vData
    End Property
    Public Property Get Reason() As Integer
        Reason = mVarReason
    End Property
    Public Property Let Remarks(ByVal vData As String)
        mvarRemarks = vData
    End Property
    Public Property Get Remarks() As String
        Remarks = mvarRemarks
    End Property
    Public Property Let RequestedUserID(ByVal vData As Double)
        mvarRequestedUserID = vData
    End Property
    Public Property Get RequestedUserID() As Double
        RequestedUserID = mvarRequestedUserID
    End Property
    Public Property Let RequestedSeatID(ByVal vData As Double)
        mvarRequestedSeatID = vData
    End Property
    Public Property Get RequestedSeatID() As Double
        RequestedSeatID = mvarRequestedSeatID
    End Property
    Public Property Let AuthorisedByAO(ByVal vData As Double)
        mvarAuthorisedByAO = vData
    End Property
    Public Property Get AuthorisedByAO() As Double
        AuthorisedByAO = mvarAuthorisedByAO
    End Property
    Public Property Let AuthorisationDateAO(ByVal vData As Date)
        mvarAuthorisationDateAO = vData
    End Property
    Public Property Get AuthorisationDateAO() As Date
        AuthorisationDateAO = mvarAuthorisationDateAO
    End Property
    Public Property Let ForwardedSeatID(ByVal vData As Double)
        mvarForwardedSeatID = vData
    End Property
    Public Property Get ForwardedSeatID() As Double
        ForwardedSeatID = mvarForwardedSeatID
    End Property
    Public Property Let FinancialYearID(ByVal vData As Integer)
        mvarFinancialYearID = vData
    End Property
    Public Property Get FinancialYearID() As Integer
        FinancialYearID = mvarFinancialYearID
    End Property
    Public Property Let Status(ByVal vData As Integer)
        mvarStatus = vData
    End Property
    Public Property Get Status() As Integer
        Status = mvarStatus
    End Property
    Public Property Let AuthorisedBySec(ByVal vData As Double)
        mvarAuthorisedBySec = vData
    End Property
    Public Property Get AuthorisedBySec() As Double
        AuthorisedBySec = mvarAuthorisedBySec
    End Property
    Public Property Let AuthorisationDateSec(ByVal vData As Date)
        mvarAuthorisationDateSec = vData
    End Property
    Public Property Get AuthorisationDateSec() As Date
        AuthorisationDateSec = mvarAuthorisationDateSec
    End Property
    Public Property Let DemandNo(ByVal vData As Double)
        mvarDemandNo = vData
    End Property
    Public Property Get DemandNo() As Double
        DemandNo = mvarDemandNo
    End Property
    Public Property Let DemandID(ByVal vData As Double)
        mvarDemandID = vData
    End Property
    Public Property Get DemandID() As Double
        DemandID = mvarDemandID
    End Property
    Public Property Let PaymentVoucherID(ByVal vData As Integer)
        mvarPaymentVoucherID = vData
    End Property
    Public Property Get PaymentVoucherID() As Integer
        PaymentVoucherID = mvarPaymentVoucherID
    End Property
    Public Property Let Paid(ByVal vData As Integer)
        mvarPaid = vData
    End Property
    Public Property Get Paid() As Integer
        Paid = mvarPaid
    End Property
    Public Property Let RecoveryRemitted(ByVal vData As Integer)
        mvarRecoveryRemitted = vData
    End Property
    Public Property Get RecoveryRemitted() As Integer
        RecoveryRemitted = mvarRecoveryRemitted
    End Property
    Public Property Let ChequeCancelled(ByVal vData As Integer)
        mvarChequeCancelled = vData
    End Property
    Public Property Get ChequeCancelled() As Integer
        ChequeCancelled = mvarChequeCancelled
    End Property
    Public Property Let InstrumentNo(ByVal vData As String)
        mvarInstrumentNo = vData
    End Property
    Public Property Get InstrumentNo() As String
        InstrumentNo = mvarInstrumentNo
    End Property
    Public Property Let InstrumentDate(ByVal vData As Date)
        mvarInstrumentDate = vData
    End Property
    Public Property Get InstrumentDate() As Date
        InstrumentDate = mvarInstrumentDate
    End Property
    
    Public Property Let TransactionDate(ByVal vData As Date)
        mTransactionDate = vData
    End Property
    Public Property Get TransactionDate() As Date
        TransactionDate = mTransactionDate
    End Property
    
    
       ' Sub or Function Definitions   '
    Public Sub SetReverseEntry()
           Dim objdb As New clsDB
           Dim mCnn As New ADODB.Connection
           Dim Rec As New ADODB.Recordset
           Dim mSql As String
           'Creating Connction              '
           If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
               MsgBox "Connection to Saankhya Not Present", vbCritical
               Exit Sub
           End If
            mSql = "Select * From faReverseEntry "
            mSql = " INNER JOIN faReverseEntryChild ON faReverseEntryChild.intRequestID=faReverseEntry.intRequestID"
            mSql = " INNER JOIN faReasons ON faReasons.intReasonID=faReverseEntry.intReasonID"
            mSql = " Where intRequstID = '" & mvarRequestID & "'"
            Rec.Open mSql, mCnn
            Call ClearVariables
            If Not (Rec.EOF And Rec.BOF) Then
                mvarRequestID = IIf(IsNull(Rec!intRequestID), -1, Rec!intRequestID)
                mvarRequestDate = IIf(IsNull(Rec!dtRequestDate), "", Rec!dtRequestDate)
                mvarCategoryID = IIf(IsNull(Rec!intCategoryID), -1, Rec!intCategoryID)
                mvarVoucherTypeID = IIf(IsNull(Rec!tnyVoucherTypeID), -1, Rec!tnyVoucherTypeID)
                mvarReasonID = IIf(IsNull(Rec!intReasonID), -1, Rec!intReasonID)
                mVarReason = IIf(IsNull(Rec!vchReason), -1, Rec!vchReason)
                mvarRemarks = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
                mvarRequestedUserID = IIf(IsNull(Rec!numRequestedUserID), -1, Rec!numRequestedUserID)
                mvarRequestedSeatID = IIf(IsNull(Rec!numRequestedSeatID), -1, Rec!numRequestedSeatID)
                mvarAuthorisedByAO = IIf(IsNull(Rec!numAuthorisedByAO), -1, Rec!numAuthorisedByAO)
                mvarAuthorisationDateAO = IIf(IsNull(Rec!dtAuthorisationDateAO), "", Rec!dtAuthorisationDateAO)
                mvarForwardedSeatID = IIf(IsNull(Rec!numForwardedSeatID), -1, Rec!numForwardedSeatID)
                mvarFinancialYearID = IIf(IsNull(Rec!intFinancialYearID), -1, Rec!intFinancialYearID)
                mvarStatus = IIf(IsNull(Rec!tnyStatus), -1, Rec!tnyStatus)
                mvarAuthorisedBySec = IIf(IsNull(Rec!numAuthorisedBySec), -1, Rec!numAuthorisedBySec)
                mvarAuthorisationDateSec = IIf(IsNull(Rec!dtAuthorisationDateSec), "", Rec!dtAuthorisationDateSec)
                mvarDemandNo = IIf(IsNull(Rec!numDemandNo), -1, Rec!numDemandNo)
                mvarDemandID = IIf(IsNull(Rec!numDemandID), -1, Rec!numDemandID)
                mvarPaymentVoucherID = IIf(IsNull(Rec!intPaymentVoucherID), -1, Rec!intPaymentVoucherID)
                mvarPaid = IIf(IsNull(Rec!tnyPaid), -1, Rec!tnyPaid)
                mvarRecoveryRemitted = IIf(IsNull(Rec!tnyRecoveryRemitted), -1, Rec!tnyRecoveryRemitted)
                mvarChequeCancelled = IIf(IsNull(Rec!tnyChequeCancelled), -1, Rec!tnyChequeCancelled)
                mvarInstrumentNo = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                mvarInstrumentDate = IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
           End If
           Rec.Close
           mCnn.Close
    End Sub
    
    Public Sub SetReverseEntryByDemand(mValDemand As String) '  Payorder Number Or PayorderID
           Dim objdb As New clsDB
           Dim mCnn As New ADODB.Connection
           Dim Rec As New ADODB.Recordset
           Dim mSql As String
           'Creating Connction              '
           If objdb.CreateNewConnection(mCnn, enuSourceString.Saankhya) = False Then
               MsgBox "Connection to Saankhya Not Present", vbCritical
               Exit Sub
           End If
            mSql = "Select * From faReverseEntry "
            mSql = " INNER JOIN faReasons ON faReasons.intReasonID=faReverseEntry.intReasonID"
            mSql = " Where " & mValDemand & " in(numDemandID,numDemandNo)"
            Rec.Open mSql, mCnn
            Call ClearVariables
            If Not (Rec.EOF And Rec.BOF) Then
                mvarRequestID = IIf(IsNull(Rec!intRequestID), -1, Rec!intRequestID)
                mvarRequestDate = IIf(IsNull(Rec!dtRequestDate), "", Rec!dtRequestDate)
                mvarCategoryID = IIf(IsNull(Rec!intCategoryID), -1, Rec!intCategoryID)
                mvarVoucherTypeID = IIf(IsNull(Rec!tnyVoucherTypeID), -1, Rec!tnyVoucherTypeID)
                mvarReasonID = IIf(IsNull(Rec!intReasonID), -1, Rec!intReasonID)
                mVarReason = IIf(IsNull(Rec!vchReason), -1, Rec!vchReason)
                mvarRemarks = IIf(IsNull(Rec!vchRemarks), "", Rec!vchRemarks)
                mvarRequestedUserID = IIf(IsNull(Rec!numRequestedUserID), -1, Rec!numRequestedUserID)
                mvarRequestedSeatID = IIf(IsNull(Rec!numRequestedSeatID), -1, Rec!numRequestedSeatID)
                mvarAuthorisedByAO = IIf(IsNull(Rec!numAuthorisedByAO), -1, Rec!numAuthorisedByAO)
                mvarAuthorisationDateAO = IIf(IsNull(Rec!dtAuthorisationDateAO), "", Rec!dtAuthorisationDateAO)
                mvarForwardedSeatID = IIf(IsNull(Rec!numForwardedSeatID), -1, Rec!numForwardedSeatID)
                mvarFinancialYearID = IIf(IsNull(Rec!intFinancialYearID), -1, Rec!intFinancialYearID)
                mvarStatus = IIf(IsNull(Rec!tnyStatus), -1, Rec!tnyStatus)
                mvarAuthorisedBySec = IIf(IsNull(Rec!numAuthorisedBySec), -1, Rec!numAuthorisedBySec)
                mvarAuthorisationDateSec = IIf(IsNull(Rec!dtAuthorisationDateSec), "", Rec!dtAuthorisationDateSec)
                mvarDemandNo = IIf(IsNull(Rec!numDemandNo), -1, Rec!numDemandNo)
                mvarDemandID = IIf(IsNull(Rec!numDemandID), -1, Rec!numDemandID)
                mvarPaymentVoucherID = IIf(IsNull(Rec!intPaymentVoucherID), -1, Rec!intPaymentVoucherID)
                mvarPaid = IIf(IsNull(Rec!tnyPaid), -1, Rec!tnyPaid)
                mvarRecoveryRemitted = IIf(IsNull(Rec!tnyRecoveryRemitted), -1, Rec!tnyRecoveryRemitted)
                mvarChequeCancelled = IIf(IsNull(Rec!tnyChequeCancelled), -1, Rec!tnyChequeCancelled)
                mvarInstrumentNo = IIf(IsNull(Rec!vchInstrumentNo), "", Rec!vchInstrumentNo)
                mvarInstrumentDate = IIf(IsNull(Rec!dtInstrumentDate), "", Rec!dtInstrumentDate)
           End If
           Rec.Close
           mCnn.Close
    End Sub
   ' Clear Variables   '
    Private Sub ClearVariables()
       mvarRequestID = -1
       mvarRequestDate = ""
       mvarCategoryID = -1
       mvarVoucherTypeID = -1
       mvarReasonID = -1
       mvarRemarks = ""
       mvarRequestedUserID = -1
       mvarRequestedSeatID = -1
       mvarAuthorisedByAO = -1
       mvarAuthorisationDateAO = ""
       mvarForwardedSeatID = -1
       mvarFinancialYearID = -1
       mvarStatus = -1
       mvarAuthorisedBySec = -1
       mvarAuthorisationDateSec = ""
       mvarDemandNo = -1
       mvarDemandID = -1
       mvarPaymentVoucherID = -1
       mvarPaid = -1
       mvarRecoveryRemitted = -1
       mvarChequeCancelled = -1
       mvarInstrumentNo = ""
       mvarInstrumentDate = ""
       mTransactionDate = Null
    End Sub

    
